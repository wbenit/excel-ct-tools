using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ExcelAddInDemo.Controllers;

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
        /// 将计费公式配置项集合转换为可直接批量写入 Excel 的二维数据矩阵 (规则 6 & 规则 7)
        /// </summary>
        /// <param name="items">公式配置项列表</param>
        /// <param name="cabDetRow">箱柜信息行物理行号 (Cab_Det)</param>
        /// <param name="subsumRow">计费小计起始物理行号 (Cab_Subsum)</param>
        /// <param name="compStartRow">元器件起始物理行号</param>
        /// <param name="compEndRow">元器件终止物理行号</param>
        /// <param name="totalCols">输出矩阵总列数 (默认 17 列，对应 A 列至 Q 列)</param>
        /// <returns>构建完成的二维数据与公式矩阵</returns>
        public static object[,] BuildFeeMatrix(
            List<FormulaItemModel> items,
            int cabDetRow,
            int subsumRow,
            int compStartRow,
            int compEndRow,
            int totalCols = 17)
        {
            // 若入参为空，返回空矩阵
            if (items == null || items.Count == 0) return new object[0, 0];

            int n = items.Count;
            // 构建 N 行 totalCols 列的二维数据矩阵
            object[,] feeMatrix = new object[n, totalCols];

            // 遍历每个配置项逐行转换
            for (int i = 0; i < n; i++)
            {
                var item = items[i];

                // A 列 (索引 0): 序号处理
                if (item.No == "总计" || i == n - 1)
                {
                    // 总计行标记为“总计”
                    feeMatrix[i, 0] = "总计";
                }
                else if (item.No == "[序号]" || string.IsNullOrWhiteSpace(item.No))
                {
                    // 动态序号公式: =ROW()-ROW(A${cabDetRow+1})
                    feeMatrix[i, 0] = $"=ROW()-ROW(A${cabDetRow + 1})";
                }
                else if (item.No.StartsWith("="))
                {
                    // 自定义公式
                    feeMatrix[i, 0] = item.No;
                }
                else
                {
                    // 普通文本序号
                    feeMatrix[i, 0] = item.No;
                }

                // B 列 (索引 1): 元件/费用名称
                feeMatrix[i, 1] = item.Name ?? string.Empty;

                // C 列 (索引 2): 规格型号
                feeMatrix[i, 2] = item.Model ?? string.Empty;

                // D 列 (索引 3): 生产厂家
                feeMatrix[i, 3] = item.Manufacturer ?? string.Empty;

                // E 列 (索引 4): 单位
                feeMatrix[i, 4] = item.Unit ?? string.Empty;

                // F 列 (索引 5): 数量 (支持公式行号平移)
                if (!string.IsNullOrEmpty(item.Quantity))
                {
                    if (item.Quantity.StartsWith("="))
                        feeMatrix[i, 5] = TransformFormulaRowOffset(item.Quantity, subsumRow);
                    else
                        feeMatrix[i, 5] = item.Quantity;
                }
                else
                {
                    feeMatrix[i, 5] = string.Empty;
                }

                // G 列 (索引 6): 单价 (支持公式行号平移)
                if (!string.IsNullOrEmpty(item.Price))
                {
                    if (item.Price.StartsWith("="))
                        feeMatrix[i, 6] = TransformFormulaRowOffset(item.Price, subsumRow);
                    else
                        feeMatrix[i, 6] = item.Price;
                }
                else
                {
                    feeMatrix[i, 6] = string.Empty;
                }

                // H 列 (索引 7): 销售总价公式转换
                if (!string.IsNullOrEmpty(item.TotalPriceFormula))
                {
                    if (item.TotalPriceFormula == "[总价小计]")
                    {
                        // 动态汇总元器件区域
                        feeMatrix[i, 7] = $"=SUM(H{compStartRow}:H{compEndRow})";
                    }
                    else if (item.TotalPriceFormula.StartsWith("="))
                    {
                        // 相对公式平移
                        feeMatrix[i, 7] = TransformFormulaRowOffset(item.TotalPriceFormula, subsumRow);
                    }
                    else
                    {
                        feeMatrix[i, 7] = item.TotalPriceFormula;
                    }
                }
                else
                {
                    feeMatrix[i, 7] = string.Empty;
                }

                // I 列 (索引 8): 备注 (保留空字符串)
                feeMatrix[i, 8] = string.Empty;

                // J 列 (索引 9): 成本单价
                feeMatrix[i, 9] = item.CostPrice ?? string.Empty;

                // K 列 (索引 10): 成本总价公式转换
                if (!string.IsNullOrEmpty(item.CostTotalPriceFormula))
                {
                    if (item.CostTotalPriceFormula == "[成本总价小计]")
                    {
                        // 动态汇总元器件成本区域
                        feeMatrix[i, 10] = $"=SUM(K{compStartRow}:K{compEndRow})";
                    }
                    else if (item.CostTotalPriceFormula.StartsWith("="))
                    {
                        // 相对成本公式平移
                        feeMatrix[i, 10] = TransformFormulaRowOffset(item.CostTotalPriceFormula, subsumRow);
                    }
                    else
                    {
                        feeMatrix[i, 10] = item.CostTotalPriceFormula;
                    }
                }
                else
                {
                    feeMatrix[i, 10] = string.Empty;
                }

                // 若有超过 16 列的输出，Q 列 (索引 16): 类别
                if (totalCols > 16)
                {
                    feeMatrix[i, 16] = item.Category ?? string.Empty;
                }
            }

            // 返回构建完成的二维矩阵
            return feeMatrix;
        }

        /// <summary>
        /// 动态平移公式表达式中的相对行号，将其映射到箱柜物理小计行 (如将 H2 转换为 H{subtotalRow})
        /// </summary>
        /// <param name="formula">待平移的公式字符串</param>
        /// <param name="subtotalRow">箱柜小计行实际物理行号</param>
        /// <returns>平移修正后的公式字符串</returns>
        public static string TransformFormulaRowOffset(string formula, int subtotalRow)
        {
            // 校验公式格式是否以等号开头
            if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith("=")) return formula;

            // 正则匹配公式中的单元格引用与行号 (如 H2, H3, H4, H5, H6, K2, K5, F7, G7 等)
            return System.Text.RegularExpressions.Regex.Replace(formula, @"([A-Z]+)(\d+)", match =>
            {
                string col = match.Groups[1].Value;
                if (int.TryParse(match.Groups[2].Value, out int rowNum))
                {
                    // 若模板中行号在 2~10 之间，平移偏移量 (rowNum - 2)
                    if (rowNum >= 1 && rowNum <= 10)
                    {
                        // 计算实际物理行号
                        int realRow = subtotalRow + (rowNum - 1);
                        return $"{col}{realRow}";
                    }
                }
                return match.Value;
            });
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
