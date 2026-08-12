using System;
using System.IO;
using System.Reflection;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 系统公共工具类，提供全局通用的路径获取、目录检索等工具方法
    /// </summary>
    public static class Tool
    {
        /// <summary>
        /// 安全获取当前插件 DLL / XLL 文件所在的实际物理目录路径 (支持 publish 及 bin 输出目录)
        /// </summary>
        /// <returns>插件物理目录绝对路径，若无法获取则返回 BaseDirectory 兜底</returns>
        public static string GetAppDirectory()
        {
            string currentDir = "";

            // 1. 尝试从当前运行程序集的 Location 获取物理路径 (需安全校验防范内存加载引发的空路径异常)
            try
            {
                // 获取程序集 Location 绝对路径
                string asmLocation = Assembly.GetExecutingAssembly().Location;

                // 判断 Location 字符串有效性
                if (!string.IsNullOrWhiteSpace(asmLocation))
                {
                    // 提取所在的文件夹路径
                    currentDir = Path.GetDirectoryName(asmLocation) ?? "";
                }
            }
            catch { }

            // 2. 若 Location 为空 (例如打包内存加载情况)，尝试获取 Excel-DNA 的 XLL 文件物理路径
            if (string.IsNullOrWhiteSpace(currentDir))
            {
                try
                {
                    // 获取 XLL 文件的绝对物理路径
                    string xllPath = ExcelDna.Integration.ExcelDnaUtil.XllPath;

                    // 判断 XLL 路径有效性
                    if (!string.IsNullOrWhiteSpace(xllPath))
                    {
                        // 提取 XLL 文件所在的文件夹路径
                        currentDir = Path.GetDirectoryName(xllPath) ?? "";
                    }
                }
                catch { }
            }

            // 3. 若仍为空，再次兜底获取 AppDomain.CurrentDomain.BaseDirectory
            if (string.IsNullOrWhiteSpace(currentDir))
            {
                try
                {
                    // 获取当前应用域根目录
                    currentDir = AppDomain.CurrentDomain.BaseDirectory ?? "";
                }
                catch { }
            }

            // 返回最终确定的物理目录路径
            return currentDir;
        }

        /// <summary>
        /// 获取当前 Windows 系统 AppData 目录下本插件专属的数据与配置存储目录路径
        /// </summary>
        /// <returns>%AppData%\ExcelAddInDemo 专属目录全路径</returns>
        public static string GetAppDataDirectory()
        {
            // 拼接 AppData 专用数据与配置文件保存目录
            string appDataDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ExcelAddInDemo"
            );

            // 检查文件夹是否存在，不存在则自动创建
            if (!Directory.Exists(appDataDir))
            {
                Directory.CreateDirectory(appDataDir);
            }

            // 返回 AppData 目录全路径
            return appDataDir;
        }

        /// <summary>
        /// 清洗指定单元格区域内公式中包含的模板外部文件绝对路径引用 (如 [CabinetTemplate.xlsx])
        /// 显式跳过 A 列 (锚点列)，在 100% 擦除公式物理路径的同时，绝对保护名称管理器与超链接
        /// </summary>
        /// <param name="targetRange">需要执行公式清洗的 Excel 单元格 Range 区域</param>
        public static void CleanRangeFormulas(Microsoft.Office.Interop.Excel.Range targetRange)
        {
            try
            {
                // 若目标区域对象为空则直接退出
                if (targetRange == null) return;

                // 尝试提取区域内所有包含公式的单元格集合 (提升遍历效率)
                Microsoft.Office.Interop.Excel.Range? formulaCells = null;
                try
                {
                    // 获取包含公式的单元格区域
                    formulaCells = targetRange.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeFormulas);
                }
                catch { }

                // 若 SpecialCells 未提取到或抛出异常，兜底直接遍历 targetRange 本身
                if (formulaCells == null)
                {
                    formulaCells = targetRange;
                }

                // 遍历包含公式的每一个单元格
                foreach (Microsoft.Office.Interop.Excel.Range cell in formulaCells)
                {
                    try
                    {
                        // 重点：显式跳过 A 列 (第 1 列)，绝对不触摸 A 列，100% 保护 A 列上绑定的定义名称与超链接
                        if (cell.Column == 1) continue;

                        // 读取单元格公式文本
                        string formula = Convert.ToString(cell.Formula) ?? "";

                        // 只有当公式中明确包含 .xlsx 外部文件引用时才进行精准替换
                        if (!string.IsNullOrEmpty(formula) && formula.IndexOf(".xlsx", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // 1. 正则匹配 '物理路径\[文件名.xlsx]工作表名'! 结构，擦除物理路径与文件名
                            string cleanedFormula = System.Text.RegularExpressions.Regex.Replace(
                                formula,
                                @"\'[^\'\n\r]*?\[[^\]\n\r]+\.[xX][lL][sS][a-zA-Z0-9]*\]([^\'\n\r]*)\'!",
                                m =>
                                {
                                    // 提取捕获到的纯工作表名 (如 "项目信息")
                                    string sheetName = m.Groups[1].Value.Trim();
                                    // 若表名有效保留 "工作表名!"，为空则返回空字符串
                                    return string.IsNullOrEmpty(sheetName) ? "" : $"{sheetName}!";
                                }
                            );

                            // 2. 清理残留的不带单引号的中括号文件名 (如 [CabinetTemplate.xlsx])
                            cleanedFormula = System.Text.RegularExpressions.Regex.Replace(
                                cleanedFormula,
                                @"\[[^\]\n\r]+\.[xX][lL][sS][a-zA-Z0-9]*\]",
                                ""
                            );

                            // 若清洗后的公式发生改变，强制写回单元格
                            if (!string.Equals(cleanedFormula, formula, StringComparison.Ordinal))
                            {
                                cell.Formula = cleanedFormula;
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}
