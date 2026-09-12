using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 系统通用工具类，提供全局通用的路径获取、目录检索等工具方法
    /// </summary>
    public static class Tool
    {
        // 静态构造函数：注册全局程序集动态解析，确保在 CAD (acad.exe) 等外部宿主调用时能自动定位加载依赖 DLL
        static Tool()
        {
            try
            {
                // 挂载全局 AssemblyResolve 事件监听
                AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            }
            catch { }
        }

        /// <summary>
        /// 全局程序集动态解析处理函数，优先从插件物理目录及应用基目录加载
        /// </summary>
        private static Assembly? CurrentDomain_AssemblyResolve(object? sender, ResolveEventArgs args)
        {
            try
            {
                // 提取缺失程序集的 DLL 文件名
                string dllName = new AssemblyName(args.Name).Name + ".dll";
                // 获取插件实际安装/运行物理目录
                string appDir = GetAppDirectory();
                if (!string.IsNullOrWhiteSpace(appDir))
                {
                    // 拼接插件目录下的目标路径
                    string target = Path.Combine(appDir, dllName);
                    if (File.Exists(target))
                    {
                        // 从插件目录动态加载
                        return Assembly.LoadFrom(target);
                    }
                }

                // 回退尝试从当前应用域基目录查找
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                if (!string.IsNullOrWhiteSpace(baseDir))
                {
                    // 拼接基目录下的目标路径
                    string target = Path.Combine(baseDir, dllName);
                    if (File.Exists(target))
                    {
                        // 从基目录动态加载
                        return Assembly.LoadFrom(target);
                    }
                }
            }
            catch { }
            return null;
        }

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

            // 2. 若 Location 为空 (例如打包内存加载情况)，尝试安全获取 Excel-DNA 的 XLL 文件物理路径
            if (string.IsNullOrWhiteSpace(currentDir))
            {
                try
                {
                    // 安全获取 XLL 文件的绝对物理路径 (不触发 JIT 强制加载 ExcelDna.Integration)
                    string? xllPath = ExcelDnaSafeAccessor.GetXllPath();

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
        /// 获取当前插件运行目录下的 data 专属数据与配置存储目录路径
        /// </summary>
        /// <returns>插件运行目录/data 专属目录全路径</returns>
        public static string GetAppDataDirectory()
        {
            // 获取插件当前运行的根物理目录
            string appDir = GetAppDirectory();

            // 拼接插件目录下的 data 专用数据与配置文件保存目录
            string appDataDir = Path.Combine(appDir, "data");

            // 检查文件夹是否存在，不存在则自动创建
            if (!Directory.Exists(appDataDir))
            {
                // 创建 data 文件夹
                Directory.CreateDirectory(appDataDir);
            }

            // 返回 data 目录全路径
            return appDataDir;
        }

        /// <summary>
        /// 提取定义名称中的纯标识文本 (清理可能存在的工作表前缀、单引号、等号与空格)
        /// </summary>
        /// <param name="rawName">原始定义名称字符串</param>
        /// <returns>清洗后的纯定义名称标识</returns>
        public static string ExtractCleanNameStr(string rawName)
        {
            // 校验入参有效性
            if (string.IsNullOrWhiteSpace(rawName)) return string.Empty;
            string clean = rawName;
            // 剔除可能包含的工作表前缀 (例如 '分类1'!Cab_Sum_1 -> Cab_Sum_1)
            if (clean.Contains("!"))
            {
                // 提取感叹号之后的纯名称标识
                clean = clean.Substring(clean.IndexOf("!") + 1);
            }
            // 修剪首尾可能存在的单引号、等号、空格与双引号
            return clean.Trim('\'', '=', ' ', '"');
        }

        /// <summary>
        /// 从定义名称全名中安全解析提取箱柜序号数字 (支持 Cab_Sum_ / Cab_Det_ / Cab_Subsum_ / Cab_Tolsum_)
        /// </summary>
        /// <param name="fullName">定义名称全称</param>
        /// <param name="sumPrefix">汇总行前缀 (可选)</param>
        /// <param name="detPrefix">信息行前缀 (可选)</param>
        /// <param name="subsumPrefix">小计行前缀 (可选)</param>
        /// <param name="tolsumPrefix">总计行前缀 (可选)</param>
        /// <returns>提取出的箱柜数字序号，失败返回 0</returns>
        public static int ExtractIndexFromName(string fullName, string? sumPrefix = null, string? detPrefix = null, string? subsumPrefix = null, string? tolsumPrefix = null)
        {
            // 校验输入字符串是否为空
            if (string.IsNullOrWhiteSpace(fullName)) return 0;

            // 提取或回退默认前缀 (复用 CabinetPrefixConfig 值类型)
            var currentPrefixes = CabinetPrefixConfig.Current;
            sumPrefix = sumPrefix ?? currentPrefixes.SumPrefix;
            detPrefix = detPrefix ?? currentPrefixes.DetPrefix;
            subsumPrefix = subsumPrefix ?? currentPrefixes.SubsumPrefix;
            tolsumPrefix = tolsumPrefix ?? currentPrefixes.TolsumPrefix;

            // 清理可能存在的工作表前缀与单引号/等号
            string cleanName = ExtractCleanNameStr(fullName);

            // 遍历 4 个前缀进行匹配提取序号
            string[] prefixes = new[] { sumPrefix, detPrefix, subsumPrefix, tolsumPrefix };
            foreach (var prefix in prefixes)
            {
                // 若匹配以前缀开头
                if (!string.IsNullOrEmpty(prefix) && cleanName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    // 截取前缀后的数字文本
                    string numStr = cleanName.Substring(prefix.Length);
                    // 解析数字
                    if (int.TryParse(numStr, out int k)) return k;
                }
            }

            // 未匹配到返回 0
            return 0;
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
                    // 若模板中行号在 1~10 之间，平移偏移量 (rowNum - 1)
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
                        // 动态汇总元器件销售总价区域 (自适应插入行并保留两位小数)
                        feeMatrix[i, 7] = $"=ROUND(SUM(H{compStartRow - 1}:INDEX(H:H,ROW()-1)),2)";
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
                        // 动态汇总元器件成本总价区域 (自适应插入行并保留两位小数)
                        feeMatrix[i, 10] = $"=ROUND(SUM(K{compStartRow - 1}:INDEX(K:K,ROW()-1)),2)";
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
        /// 构建元器件区域的二维公式与数据矩阵 (17 列，覆盖 A 列至 Q 列)
        /// 遵循规则 6 与规则 7，并根据要求动态填充 F、G、H、J、K、L、N、Q 列的自适应空行判断公式
        /// </summary>
        /// <param name="compStartRow">元器件起始物理行号</param>
        /// <param name="compEndRow">元器件终止物理行号</param>
        /// <param name="cabDetRow">箱柜明细信息行物理行号 (用于 A 列序号偏移)</param>
        /// <param name="totalCols">矩阵总列数 (默认 17 列，对应 A 列至 Q 列)</param>
        /// <param name="components">可选的已有元器件实体数据列表</param>
        /// <returns>构建完成的元器件二维数据与公式矩阵</returns>
        public static object[,] BuildComponentRowsMatrix(
            int compStartRow,
            int compEndRow,
            int cabDetRow,
            int totalCols = 17,
            List<Models.ComponentItem>? components = null)
        {
            // 校验行号区间有效性
            if (compEndRow < compStartRow || compStartRow <= 0) return new object[0, 0];

            int rowCount = compEndRow - compStartRow + 1;
            // 确保总列数至少为 17 列
            int cols = Math.Max(totalCols, 17);
            object[,] matrix = new object[rowCount, cols];

            // 遍历元器件区域每一行填充数据与自适应公式
            for (int r = 0; r < rowCount; r++)
            {
                // 当前单元格的绝对物理行号
                int currPhysicalRow = compStartRow + r;

                // A 列 (索引 0): 动态序号公式
                matrix[r, 0] = $"=ROW()-ROW(A${cabDetRow + 1})";

                // B 列 (索引 1): 元件名称
                string name = string.Empty;
                // C 列 (索引 2): 规格型号
                string spec = string.Empty;
                // D 列 (索引 3): 生产厂家
                string mfr = string.Empty;
                // E 列 (索引 4): 计量单位
                string unit = string.Empty;
                // M 列 (索引 12): 面价/基准价
                object mVal = string.Empty;

                // 若传入了元器件实体且当前行索引在范围内
                if (components != null && r < components.Count)
                {
                    var comp = components[r];
                    name = comp.Name ?? string.Empty;
                    spec = comp.Specification ?? string.Empty;
                    mfr = comp.Manufacturer ?? string.Empty;
                    unit = comp.Unit ?? string.Empty;
                    if (comp.UnitPrice > 0) mVal = comp.UnitPrice;
                }

                matrix[r, 1] = name;
                matrix[r, 2] = spec;
                matrix[r, 3] = mfr;
                matrix[r, 4] = unit;

                // F 列 (索引 5): 数量 (若有实体按真实数量赋值，否则采用空行自适应公式)
                if (components != null && r < components.Count)
                {
                    decimal qty = components[r].Quantity > 0 ? components[r].Quantity : 1;
                    matrix[r, 5] = qty;
                }
                else
                {
                    matrix[r, 5] = $"=IF(AND(B{currPhysicalRow}=\"\",C{currPhysicalRow}=\"\"),\"\",1)";
                }

                // G 列 (索引 6): 销售单价 =IF(AND(B{row}="",C{row}=""),"",ROUND(M{row}*L{row}*N{row},2))
                matrix[r, 6] = $"=IF(AND(B{currPhysicalRow}=\"\",C{currPhysicalRow}=\"\"),\"\",ROUND(M{currPhysicalRow}*L{currPhysicalRow}*N{currPhysicalRow},2))";

                // H 列 (索引 7): 销售总价 =IF(AND(B{row}="",C{row}=""),"",ROUND(F{row}*G{row},2))
                matrix[r, 7] = $"=IF(AND(B{currPhysicalRow}=\"\",C{currPhysicalRow}=\"\"),\"\",ROUND(F{currPhysicalRow}*G{currPhysicalRow},2))";

                // I 列 (索引 8): 备注 (保留空字符串)
                matrix[r, 8] = string.Empty;

                // J 列 (索引 9): 成本单价 =IF(AND(B{row}="",C{row}=""),"",ROUND(M{row}*N{row},2))
                if (components != null && r < components.Count && components[r].CostUnitPrice > 0)
                {
                    matrix[r, 9] = components[r].CostUnitPrice;
                }
                else
                {
                    matrix[r, 9] = $"=IF(AND(B{currPhysicalRow}=\"\",C{currPhysicalRow}=\"\"),\"\",ROUND(M{currPhysicalRow}*N{currPhysicalRow},2))";
                }

                // K 列 (索引 10): 成本总价 =IF(AND(B{row}="",C{row}=""),"",ROUND(J{row}*F{row},2))
                matrix[r, 10] = $"=IF(AND(B{currPhysicalRow}=\"\",C{currPhysicalRow}=\"\"),\"\",ROUND(J{currPhysicalRow}*F{currPhysicalRow},2))";

                // L 列 (索引 11): 加价系数 =IF(AND(B{row}="",C{row}=""),"",1)
                matrix[r, 11] = $"=IF(AND(B{currPhysicalRow}=\"\",C{currPhysicalRow}=\"\"),\"\",1)";

                // M 列 (索引 12): 面价/基准单价
                matrix[r, 12] = mVal;

                // N 列 (索引 13): 折扣/采购系数 =IF(AND(B{row}="",C{row}=""),"",1)
                matrix[r, 13] = $"=IF(AND(B{currPhysicalRow}=\"\",C{currPhysicalRow}=\"\"),\"\",1)";

                // O 列 (索引 14): 预留
                matrix[r, 14] = string.Empty;

                // P 列 (索引 15): 预留
                matrix[r, 15] = string.Empty;

                // Q 列 (索引 16): 类别 =IF(AND(B{row}="",C{row}=""),"","元件")
                matrix[r, 16] = $"=IF(AND(B{currPhysicalRow}=\"\",C{currPhysicalRow}=\"\"),\"\",\"元件\")";
                if (totalCols >= 20)
                {
                    // 覆盖写入 A 列至 Q 列完整元器件二维数组
                    matrix[r, 20] = spec;
                }
            }

            // 返回构建完成的元器件二维矩阵
            return matrix;
        }

        /// <summary>
        /// 清洗指定单元格区域内公式中包含的模板外部文件绝对路径引用 (如 [CabinetTemplate.xlsx])
        /// 显式跳过 A 列 (锚点列)，在 100% 擦除公式物理路径的同时，绝对保护名称管理器与超链接
        /// </summary>
        /// <param name="targetRange">需要执行公式清洗的 Excel 单元格 Range 区域</param>
        public static void CleanRangeFormulas(dynamic targetRange)
        {
            try
            {
                // 若目标区域对象为空则直接退出
                if (targetRange == null) return;

                // 尝试提取区域内所有包含公式的单元格集合 (提升遍历效率)
                dynamic? formulaCells = null;
                try
                {
                    // 获取包含公式的单元格区域 (-4123 对应 xlCellTypeFormulas)
                    formulaCells = targetRange.SpecialCells(-4123);
                }
                catch { }

                // 若 SpecialCells 未提取到或抛出异常，兜底直接遍历 targetRange 本身
                if (formulaCells == null)
                {
                    formulaCells = targetRange;
                }

                // 遍历包含公式的每一个单元格
                foreach (dynamic cell in formulaCells)
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

        /// <summary>
        /// 对比区域现有行数与目标行数，在指定起始物理行位置自动完成插行或删行对齐
        /// 遵循规则 6 紧凑无空行原则
        /// </summary>
        /// <param name="sheet">目标工作表 COM 引用</param>
        /// <param name="startRow">插删行的基准起始物理行号</param>
        /// <param name="currentCount">当前现有行数</param>
        /// <param name="targetCount">目标所需行数</param>
        /// <returns>行数变化差值 (targetCount - currentCount，正数表示插入行数，负数表示删除行数)</returns>
        public static int AlignRowRangeCount(dynamic sheet, int startRow, int currentCount, int targetCount)
        {
            // 校验工作表与行号参数有效性
            if (sheet == null || startRow <= 0) return 0;

            // 计算行数差值
            int diff = targetCount - currentCount;

            // 1. 若目标行数多于现有行数：在起始行处向下插入差值行
            if (diff > 0)
            {
                // 获取待插入行的 Range 区域
                dynamic insertRange = sheet.Rows[$"{startRow}:{startRow + diff - 1}"];
                // 执行向下位移插入新行
                insertRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
            }
            // 2. 若目标行数少于现有行数：从起始行起向上删除多余行
            else if (diff < 0)
            {
                // 计算需要删除的行数
                int deleteCount = -diff;
                // 获取待删除行的 Range 区域
                dynamic deleteRange = sheet.Rows[$"{startRow}:{startRow + deleteCount - 1}"];
                // 执行向上位移删除行
                deleteRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
            }

            // 返回行数变化差值
            return diff;
        }

        /// <summary>
        /// 扫描定义名称集合，构建 箱柜序号 → CabinetAnchorModel 锚点列表
        /// 自动过滤非当前工作表的跨表引用，仅保留属于 currentSheetName 的锚点
        /// </summary>
        /// <param name="allNames">已收集的工作簿/工作表定义名称列表（dynamic COM 对象）</param>
        /// <param name="currentSheetName">当前活动工作表名称，用于过滤跨表引用</param>
        /// <param name="sumPrefix">汇总行定义名称前缀</param>
        /// <param name="detPrefix">箱柜信息行定义名称前缀</param>
        /// <param name="subsumPrefix">小计行定义名称前缀</param>
        /// <param name="tolsumPrefix">总计行定义名称前缀</param>
        /// <returns>按汇总行物理行号升序排列的有效箱柜锚点强类型列表</returns>
        public static List<KeyValuePair<int, Models.CabinetAnchorModel>> BuildCabinetMap(
            IEnumerable<dynamic> allNames,
            string currentSheetName,
            string sumPrefix, string detPrefix,
            string subsumPrefix, string tolsumPrefix)
        {
            // 构建中间字典，Key 为箱柜序号，Value 为强类型锚点模型
            var cabinetDict = new Dictionary<int, Models.CabinetAnchorModel>();

            // 遍历所有定义名称，逐个解析并填充锚点字典
            foreach (dynamic name in allNames)
            {
                try
                {
                    // 清洗提取定义名称字符串
                    string clean = ExtractCleanNameStr(Convert.ToString(name.Name) ?? "");

                    // 提取箱柜数字序号，无法匹配则跳过
                    int k = ExtractIndexFromName(clean, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                    if (k <= 0) continue;

                    // 安全读取定义名称所指向的单元格 Range 引用
                    dynamic? refRange = null;
                    try { refRange = name.RefersToRange; } catch { }
                    if (refRange == null) continue;

                    // 校验该定义名称是否属于当前活动工作表，避免跨表误取
                    string refSheetName = "";
                    try { refSheetName = refRange.Worksheet.Name; } catch { }
                    if (!string.IsNullOrEmpty(refSheetName) &&
                        !string.Equals(refSheetName, currentSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        // 不属于当前工作表，跳过
                        continue;
                    }

                    // 初始化字典中该序号的锚点模型
                    if (!cabinetDict.ContainsKey(k)) cabinetDict[k] = new Models.CabinetAnchorModel();

                    // 匹配 Det 锚点（箱柜信息行）
                    if (clean.StartsWith(detPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        cabinetDict[k].Det = refRange;
                    }
                    // 匹配 Sum 锚点（汇总行）
                    else if (clean.StartsWith(sumPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        cabinetDict[k].Sum = refRange;
                    }
                    // 匹配 Subsum 锚点（小计行）
                    else if (clean.StartsWith(subsumPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        cabinetDict[k].Subsum = refRange;
                    }
                    // 匹配 Tolsum 锚点（总计行）
                    else if (clean.StartsWith(tolsumPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        cabinetDict[k].Tolsum = refRange;
                    }
                }
                catch { }
            }

            // 兼容普通有明细箱柜 (同时具备 Sum 和 Det) 与纯汇总无明细箱柜 (仅具备 Sum 锚点)
            return cabinetDict
                .Where(x => x.Value.Sum != null || x.Value.Det != null)
                .OrderBy(x => x.Value.Sum != null ? (int)x.Value.Sum.Row : (int)x.Value.Det.Row)
                .ToList();
        }

        /// <summary>
        /// 安全获取当前 Excel 活动运行环境上下文 (Application, ActiveWorkbook, ActiveSheet)
        /// 支持外部显式传入指定对象以进行覆盖
        /// </summary>
        /// <param name="explicitApp">外部显式传入的 Application 实例 (可选)</param>
        /// <param name="explicitSheet">外部显式传入的 Worksheet 实例 (可选)</param>
        /// <returns>包含 app, wb, sheet 的强类型上下文实体，任一关键对象获取失败时返回 null</returns>
        public static Models.ExcelContext? GetActiveExcelContext(
            dynamic? explicitApp = null,
            dynamic? explicitSheet = null)
        {
            // 获取当前运行的 Excel Application COM 接口实例 (优先使用传入实例，回退 ExcelDnaSafeAccessor)
            dynamic? app = explicitApp ?? ExcelDnaSafeAccessor.GetApplication();
            if (app == null) return null;

            // 获取当前激活的工作簿对象
            dynamic? wb = app.ActiveWorkbook;
            if (wb == null) return null;

            // 获取目标工作表对象 (优先使用传入工作表，回退活动工作表)
            dynamic? sheet = explicitSheet ?? wb.ActiveSheet;
            if (sheet == null) return null;

            // 返回包装完成的强类型上下文实体对象
            return new Models.ExcelContext(app, wb, sheet);
        }

        /// <summary>
        /// 安全收集指定工作簿与工作表中的所有定义名称 dynamic COM 对象 (双作用域完整扫描)
        /// 若初次扫描未获取到任何定义名称且指定了有效工作表，将自动触发智能识别并重新构建定义名称
        /// </summary>
        /// <param name="wb">目标工作簿 COM 对象 (可选)</param>
        /// <param name="sheet">目标工作表 COM 对象 (可选)</param>
        /// <param name="autoRebuildIfEmpty">若定义名称为空是否自动触发重新构建 (默认 true)</param>
        /// <returns>合并后的定义名称列表</returns>
        public static List<dynamic> CollectAllDefinedNames(
            dynamic? wb,
            dynamic? sheet,
            bool autoRebuildIfEmpty = true)
        {
            // 创建定义名称列表容器
            var allNames = new List<dynamic>();

            // 内部辅助局部方法：执行双作用域扫描
            void ScanNames()
            {
                // 1. 尝试遍历收集工作簿级定义名称
                if (wb != null)
                {
                    try
                    {
                        // 遍历工作簿名称集合
                        if (wb.Names != null)
                        {
                            foreach (dynamic n in wb.Names) allNames.Add(n);
                        }
                    }
                    catch { }
                }

                // 2. 尝试遍历收集工作表级定义名称
                if (sheet != null)
                {
                    try
                    {
                        // 遍历工作表名称集合
                        if (sheet.Names != null)
                        {
                            foreach (dynamic n in sheet.Names) allNames.Add(n);
                        }
                    }
                    catch { }
                }
            }

            // 执行初次名称扫描
            ScanNames();

            // 若扫描结果为空且允许自动重建且工作表对象有效
            if ((allNames == null || allNames.Count == 0) && autoRebuildIfEmpty && sheet != null)
            {
                // 自动触发单表智能识别与定义名称补齐重建
                FixAndFillCabinetNamesForSheet(sheet);

                // 清空后重新执行名称扫描
                allNames?.Clear();
                ScanNames();
            }

            // 返回最终合并收集的所有定义名称集合 (保障非空)
            return allNames ?? new List<dynamic>();
        }

        /// <summary>
        /// 快捷公共方法：获取指定工作表中按汇总行物理行号升序排列的有效箱柜映射列表
        /// 内部自动读取系统配置前缀、自动聚合双作用域定义名称并完成锚点结构化构建
        /// 若定义名称缺失或为空，底层 CollectAllDefinedNames 自动触发智能重建补齐
        /// </summary>
        /// <param name="sheet">目标工作表 COM 对象</param>
        /// <param name="wb">所属工作簿 COM 对象 (可选，若为空自动通过 sheet.Parent 向上获取)</param>
        /// <returns>按汇总行物理行号升序排列的有效箱柜映射列表</returns>
        public static List<KeyValuePair<int, Models.CabinetAnchorModel>> GetSheetValidCabinets(
            object sheet,
            object? wb = null)
        {
            // 校验工作表对象有效性
            if (sheet == null) return new List<KeyValuePair<int, Models.CabinetAnchorModel>>();

            // 转换为 dynamic 方便调用 COM 属性
            dynamic dSheet = sheet;
            dynamic? dWb = wb;

            // 若工作簿为空尝试向上从 sheet.Parent 获取
            if (dWb == null)
            {
                // 向上追溯所属工作簿
                try { dWb = dSheet.Parent; } catch { }
            }

            // 读取全局配置中的 4 个定义名称前缀 (零堆分配)
            var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

            // 提取当前工作表纯文本名称
            string sheetName = Convert.ToString(dSheet.Name) ?? "";

            // 收集双作用域所有定义名称 (若为空底层自动触发智能重建)
            var allNames = CollectAllDefinedNames(dWb, dSheet, autoRebuildIfEmpty: true);

            // 调用 BuildCabinetMap 构建有效箱柜列表
            var validCabinets = BuildCabinetMap(allNames, sheetName, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);

            // 若构建结果仍为空，进行二次容错重建与重新构建
            if (validCabinets == null || validCabinets.Count == 0)
            {
                // 强制触发定义名称补齐重建
                FixAndFillCabinetNamesForSheet(dSheet);
                // 再次收集定义名称
                allNames = CollectAllDefinedNames(dWb, dSheet, autoRebuildIfEmpty: false);
                // 再次构建有效箱柜列表
                validCabinets = BuildCabinetMap(allNames, sheetName, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
            }

            // 返回构建结果 (保障非空)
            return validCabinets ?? new List<KeyValuePair<int, Models.CabinetAnchorModel>>();
        }


        /// <summary>
        /// 根据指定的物理行号智能匹配其所属的箱柜锚点对象 (命中汇总行或明细大标题至总计行区间)
        /// 遵循规则 6 架构标准
        /// </summary>
        /// <param name="validCabinets">有效箱柜映射集合</param>
        /// <param name="row">待匹配的物理行号</param>
        /// <returns>匹配命中的箱柜键值对实体，未命中返回 null</returns>
        public static KeyValuePair<int, Models.CabinetAnchorModel>? FindCabinetByRow(
            IEnumerable<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets,
            int row)
        {
            // 校验入参与行号有效性
            if (validCabinets == null || row <= 0) return null;

            // 遍历所有有效箱柜进行区间命中判定
            foreach (var cab in validCabinets)
            {
                // 读取汇总行行号
                int sumR = cab.Value.Sum != null ? Convert.ToInt32(cab.Value.Sum.Row) : 0;
                // 读取箱柜信息行行号
                int detR = cab.Value.Det != null ? Convert.ToInt32(cab.Value.Det.Row) : 0;
                // 读取总计行行号 (若总计行为空按默认 27 行估算)
                int tolR = cab.Value.Tolsum != null ? Convert.ToInt32(cab.Value.Tolsum.Row) : (detR + 27);

                // 判定行号是否命中汇总行，或落在明细大标题至总计行下方3行报价人信息完整区间 (规则 6 扩展)
                if (row == sumR || (detR > 0 && row >= (detR - 3) && row <= (tolR + 3)))
                {
                    // 返回命中的箱柜键值对
                    return cab;
                }
            }

            // 未命中返回 null
            return null;
        }

        /// <summary>
        /// 从 Excel Application 当前选区 (Selection) 智能扫描并提取所有命中的箱柜实体列表 (支持单选、跨行选区及离散 Areas)
        /// 遵循规则 6 架构标准，按汇总行物理行号升序排列并去重
        /// </summary>
        /// <param name="app">Excel Application COM 接口实例</param>
        /// <param name="validCabinets">当前工作表有效箱柜映射列表</param>
        /// <param name="fallbackActiveCell">当选区未匹配到任何箱柜时，是否尝试回退检查 ActiveCell (默认 true)</param>
        /// <param name="fallbackSingle">当仍未命中且当前工作表仅有 1 台箱柜时是否自动回退该箱柜 (默认 true)</param>
        /// <returns>命中的有效箱柜键值对集合 (保证非空，按行号升序排列)</returns>
        public static List<KeyValuePair<int, Models.CabinetAnchorModel>> GetSelectedCabinets(
            object app,
            List<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets,
            bool fallbackActiveCell = true,
            bool fallbackSingle = true)
        {
            // 校验入参有效性
            if (app == null || validCabinets == null || validCabinets.Count == 0)
            {
                // 返回空列表
                return new List<KeyValuePair<int, Models.CabinetAnchorModel>>();
            }

            // 存储用户选区覆盖的所有物理行号集合 (使用 HashSet 防重复)
            var selectedRows = new HashSet<int>();
            try
            {
                // 转换 COM 句柄
                dynamic dApp = app;
                // 获取当前活动工作表的选区对象
                dynamic? selection = dApp.Selection;
                if (selection != null)
                {
                    // 获取 Areas 区域集合 (支持 Ctrl 离散多选与多区域)
                    dynamic areas = selection.Areas;
                    int areaCount = 1;
                    try { areaCount = Convert.ToInt32(areas.Count); } catch { }

                    // 遍历所有选区 Area
                    for (int a = 1; a <= areaCount; a++)
                    {
                        // 获取单个 Area 对象
                        dynamic area = areaCount > 1 ? areas[a] : selection;
                        int startRow = Convert.ToInt32(area.Row);
                        int rowCount = Convert.ToInt32(area.Rows.Count);

                        // 将该区域覆盖的所有行号加入集合
                        for (int r = startRow; r < startRow + rowCount; r++)
                        {
                            selectedRows.Add(r);
                        }
                    }
                }
            }
            catch { }

            // 存储命中匹配的箱柜字典 (以箱柜序号 K 作为键去重)
            var matchedDict = new Dictionary<int, KeyValuePair<int, Models.CabinetAnchorModel>>();

            // 1. 若提取到了选区行号，逐行调用 FindCabinetByRow 匹配箱柜
            if (selectedRows.Count > 0)
            {
                // 遍历所有选中的物理行号
                foreach (int row in selectedRows)
                {
                    // 查找命中的箱柜实体
                    var matched = FindCabinetByRow(validCabinets, row);
                    if (matched.HasValue && !matchedDict.ContainsKey(matched.Value.Key))
                    {
                        // 记录命中的箱柜
                        matchedDict.Add(matched.Value.Key, matched.Value);
                    }
                }
            }

            // 2. 若选区未匹配到任何箱柜且允许回退 ActiveCell 光标
            if (matchedDict.Count == 0 && fallbackActiveCell)
            {
                int activeRow = 0;
                try
                {
                    // 获取活动焦点单元格行号
                    dynamic dApp = app;
                    dynamic? activeCell = dApp.ActiveCell;
                    if (activeCell != null) activeRow = Convert.ToInt32(activeCell.Row);
                }
                catch { }

                // 若活动行号有效进行匹配
                if (activeRow > 0)
                {
                    // 查找当前光标命中的箱柜
                    var matched = FindCabinetByRow(validCabinets, activeRow);
                    if (matched.HasValue && !matchedDict.ContainsKey(matched.Value.Key))
                    {
                        // 加入字典
                        matchedDict.Add(matched.Value.Key, matched.Value);
                    }
                }
            }

            // 3. 若仍未匹配到且允许单箱柜自动回退 (当且仅当当前表仅有 1 台箱柜)
            if (matchedDict.Count == 0 && fallbackSingle && validCabinets.Count == 1)
            {
                // 回退返回当前唯一的箱柜
                return new List<KeyValuePair<int, Models.CabinetAnchorModel>> { validCabinets[0] };
            }

            // 4. 将命中箱柜集合按汇总行物理行号从上到下升序排序输出
            var resultList = new List<KeyValuePair<int, Models.CabinetAnchorModel>>(matchedDict.Values);
            resultList.Sort((a, b) =>
            {
                int rowA = a.Value.Sum != null ? Convert.ToInt32(a.Value.Sum.Row) : 0;
                int rowB = b.Value.Sum != null ? Convert.ToInt32(b.Value.Sum.Row) : 0;
                return rowA.CompareTo(rowB);
            });

            // 返回排序后的箱柜列表
            return resultList;
        }

        /// <summary>
        /// 从 Excel Application 活动单元格光标或选区智能匹配所属箱柜 (支持单箱柜自动兜底)
        /// </summary>
        /// <param name="app">Excel Application COM 接口实例</param>
        /// <param name="validCabinets">当前工作表有效箱柜映射列表</param>
        /// <param name="fallbackSingle">当未命中且当前表仅有 1 台箱柜时是否自动回退该箱柜 (默认 true)</param>
        /// <returns>匹配到的箱柜实体，未命中返回 null</returns>
        public static KeyValuePair<int, Models.CabinetAnchorModel>? GetActiveCabinet(
            object app,
            List<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets,
            bool fallbackSingle = true)
        {
            // 校验入参有效性
            if (app == null || validCabinets == null || validCabinets.Count == 0) return null;

            // 调用通用选区与光标扫描方法
            var list = GetSelectedCabinets(app, validCabinets, fallbackActiveCell: true, fallbackSingle: fallbackSingle);

            // 若命中返回首个匹配的箱柜
            if (list != null && list.Count > 0)
            {
                // 返回首个命中实体
                return list[0];
            }

            // 未匹配到返回 null
            return null;
        }



        /// <summary>
        /// 遍历当前工作簿中的所有工作表，根据顶部汇总与明细特征自动校准补齐 4 个定义名称
        /// 遵循规则 6 架构与规则 7 内存批量读入
        /// </summary>
        /// <param name="targetWb">目标工作簿 COM 对象，若为空则自动使用当前活动工作簿</param>
        /// <returns>修复/校准的箱柜总数</returns>
        public static int FixAndFillCabinetNamesForAllSheets(dynamic? targetWb = null)
        {
            // 记录全局处理的箱柜累计总数
            int totalFixedCabinets = 0;
            try
            {
                // 获取 Excel 应用程序实例 (安全调用)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return 0;

                // 若未传入工作簿则获取当前激活的工作簿
                if (targetWb == null) targetWb = app.ActiveWorkbook;
                if (targetWb == null) return 0;

                // 暂存屏刷、警告与事件响应状态以提升执行效率
                bool prevUpdating = app.ScreenUpdating;
                bool prevAlerts = app.DisplayAlerts;
                bool prevEvents = app.EnableEvents;

                // 关闭界面交互刷新提效
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                try
                {
                    // 遍历工作簿中的每一个工作表
                    foreach (dynamic sheet in targetWb.Worksheets)
                    {
                        // 针对单张工作表执行定义名称补齐与校准
                        totalFixedCabinets += FixAndFillCabinetNamesForSheet(sheet);
                    }
                }
                finally
                {
                    // 恢复原始运行状态
                    app.ScreenUpdating = prevUpdating;
                    app.DisplayAlerts = prevAlerts;
                    app.EnableEvents = prevEvents;
                }
            }
            catch (Exception ex)
            {
                // 记录遍历补齐定义名称异常日志
                LogHelper.WriteLog($"遍历补齐工作簿定义名称失败: {ex.Message}");
            }

            // 返回累计修复的箱柜数量
            return totalFixedCabinets;
        }

        // 线程安全警告收集集合：记录最近一次箱柜校准/识别中未匹配预设方案等警告
        private static readonly System.Collections.Concurrent.ConcurrentBag<string> _cabinetWarnings = new();

        /// <summary>
        /// 清空最近一次收集的箱柜校准与费用匹配警告信息
        /// </summary>
        public static void ClearCabinetWarnings()
        {
            // 清空线程安全集合中所有历史警告
            while (_cabinetWarnings.TryTake(out _)) { }
        }

        /// <summary>
        /// 获取当前收集到的所有去重箱柜校准与费用匹配警告信息列表
        /// </summary>
        /// <returns>去重后的警告提示文本列表</returns>
        public static List<string> GetCabinetWarnings()
        {
            // 返回去重后的警告集合列表
            return _cabinetWarnings.Distinct().ToList();
        }

        /// <summary>
        /// 记录单条箱柜警告信息
        /// </summary>
        /// <param name="warning">警告信息内容</param>
        public static void AddCabinetWarning(string warning)
        {
            // 校验入参有效性
            if (!string.IsNullOrWhiteSpace(warning))
            {
                // 添加到线程安全集合中
                _cabinetWarnings.Add(warning.Trim());
            }
        }

        /// <summary>
        /// 针对单张工作表，根据顶部汇总与明细区域特征校准补齐 4 个定义名称
        /// 规则 6: Cab_Sum_k (汇总行), Cab_Det_k (信息行), Cab_Subsum_k (小计行), Cab_Tolsum_k (总计行)
        /// 规则 7: 采用数组一次性读到内存
        /// </summary>
        /// <param name="sheet">目标工作表 COM 引用</param>
        /// <returns>当前工作表修复的箱柜数量</returns>
        public static int FixAndFillCabinetNamesForSheet(dynamic sheet)
        {
            // 校验工作表入参有效性
            if (sheet == null) return 0;

            try
            {
                // 获取工作表名称
                string sheetName = Convert.ToString(sheet.Name) ?? "";
                if (string.IsNullOrWhiteSpace(sheetName)) return 0;

                // 读取 4 种定义名称前缀配置项 (零堆分配)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

                // 读取顶部汇总行基准起始物理行号配置项 (默认 7)
                int cabSumStartRow = ConfigManager.Instance.Current.Excel.CabSumRowIndex;

                // 获取工作表已用区域 UsedRange
                dynamic usedRange = sheet.UsedRange;
                if (usedRange == null) return 0;

                // 获取已用区域起始行与总行数
                int usedStartRow = Convert.ToInt32(usedRange.Row);
                int totalRows = Convert.ToInt32(usedRange.Rows.Count);
                int usedEndRow = usedStartRow + totalRows - 1;
                if (totalRows <= 0) return 0;

                // 规则 7: 一次性读取已用区域的数值与公式数组到内存
                object[,]? valArray = null;
                object[,]? formulaArray = null;
                try { valArray = usedRange.Value2 as object[,]; } catch { }
                try { formulaArray = usedRange.Formula as object[,]; } catch { }
                if (valArray == null) return 0;

                // 获取内存二维数组的行列边界
                int arrRows = valArray.GetLength(0);
                int arrCols = valArray.GetLength(1);

                // 本地辅助函数：安全获取指定物理行和列(1-based)的纯文本
                string GetText(int r, int c)
                {
                    // 计算在二维数组中的相对行索引
                    int ar = r - usedStartRow + 1;
                    // 边界越界校验
                    if (ar < 1 || ar > arrRows || c < 1 || c > arrCols) return "";
                    // 提取并返回修剪后的单元格文本
                    return Convert.ToString(valArray[ar, c])?.Trim() ?? "";
                }

                // 本地辅助函数：安全获取指定物理行和列(1-based)的公式字符串
                string GetFormula(int r, int c)
                {
                    // 校验公式数组有效性
                    if (formulaArray == null) return "";
                    // 计算相对行索引
                    int ar = r - usedStartRow + 1;
                    // 边界越界校验
                    if (ar < 1 || ar > arrRows || c < 1 || c > arrCols) return "";
                    // 提取并返回修剪后的单元格公式
                    return Convert.ToString(formulaArray[ar, c])?.Trim() ?? "";
                }

                // 1. 【扫描明细区域中的所有箱柜信息行 Cab_Det】
                // 特征条件：A 列包含“柜号”或“箱柜”，且下一行 A 列包含“序号”或“项次”等表头标记
                var detRows = new List<int>();
                for (int r = cabSumStartRow + 1; r < usedEndRow; r++)
                {
                    // 提取当前行与下一行的 A 列文本
                    string aText = GetText(r, 1);
                    string nextAText = GetText(r + 1, 1);

                    // 匹配明细大标题与表头特征 (容错序号、项次、NO等)
                    if ((aText.Contains("柜号") || aText.Contains("箱柜")) &&
                        (nextAText.Contains("序号") || nextAText.Contains("项次") || nextAText.Contains("NO") || nextAText.Contains("No")))
                    {
                        // 记录识别到的箱柜信息行行号
                        detRows.Add(r);
                    }
                }

                // 2. 【扫描顶部汇总行 Cab_Sum】
                // 若存在明细行，则终止于首个明细行之前；若无明细行，则扫描至已用区域末尾
                var sumRows = new List<int>();
                // 获取首个明细行的分界位置
                int firstDetRow = detRows.Count > 0 ? detRows[0] : (usedEndRow + 1);
                // 遍历扫描顶部汇总行
                for (int r = cabSumStartRow; r < firstDetRow; r++)
                {
                    // 检查 B 列或 A 列是否有箱柜编号/名称
                    string bVal = GetText(r, 2);
                    string aVal = GetText(r, 1);

                    // 若存在非空内容则判定为有效汇总行
                    if (!string.IsNullOrWhiteSpace(bVal) || !string.IsNullOrWhiteSpace(aVal))
                    {
                        // 记录识别到的有效汇总行物理行号
                        sumRows.Add(r);
                    }
                }

                // 若明细块与汇总行均未识别出任何箱柜，判定为非标准表，跳过
                if (detRows.Count == 0 && sumRows.Count == 0) return 0;

                // 箱柜总数取明细块与汇总行两者的较大值，全面兼容纯汇总无明细箱柜
                int cabCount = Math.Max(detRows.Count, sumRows.Count);

                // 3. 【精准配对顶部汇总行与底表明细块，全面支持穿插纯汇总无明细箱柜】
                // 建立每台箱柜分配到的明细行映射 (数组索引 i 对应第 k = i + 1 台箱柜，值为 0 代表纯汇总无明细)
                int[] matchedDetRows = new int[cabCount];

                // 若同时存在汇总行与明细行，进行智能特征与柜号双向对齐
                if (sumRows.Count > 0 && detRows.Count > 0)
                {
                    // 记录底表明细行是否已被配对使用
                    bool[] detUsed = new bool[detRows.Count];

                    // 第一轮：通过柜号/箱柜名称进行 100% 完全精确对齐 (优先原则)
                    for (int i = 0; i < sumRows.Count; i++)
                    {
                        // 提取汇总行柜号文本 (优先 B 列，若 B 列为空则取 A 列)
                        string sumCabNo = GetText(sumRows[i], 2);
                        if (string.IsNullOrEmpty(sumCabNo)) sumCabNo = GetText(sumRows[i], 1);
                        // 若汇总行未填写柜号则跳过本轮匹配
                        if (string.IsNullOrEmpty(sumCabNo)) continue;

                        // 遍历底表所有未使用的明细块寻找对应箱柜
                        for (int j = 0; j < detRows.Count; j++)
                        {
                            // 跳过已被配对的明细行
                            if (detUsed[j]) continue;
                            // 提取底表明细信息行柜号 (优先 B 列，若 B 列为空则容错提取 A 列合并单元格文本)
                            string detCabNo = GetText(detRows[j], 2);
                            if (string.IsNullOrEmpty(detCabNo)) detCabNo = GetText(detRows[j], 1);
                            string detTextA = GetText(detRows[j], 1);

                            // 只要柜号完全一致，或明细信息行 A 列明确包含汇总柜号，立即建立 1 对 1 精准绑定
                            if ((!string.IsNullOrEmpty(detCabNo) && string.Equals(sumCabNo, detCabNo, StringComparison.OrdinalIgnoreCase)) ||
                                (!string.IsNullOrEmpty(detTextA) && detTextA.Contains(sumCabNo)))
                            {
                                // 记录当前汇总行精准绑定的明细行行号
                                matchedDetRows[i] = detRows[j];
                                // 标记该明细行已被占用
                                detUsed[j] = true;
                                break;
                            }
                        }
                    }

                    // 第二轮：通过汇总行 G 列单价公式引用进行底层关联匹配 (容错柜号名称不完全一致的情况)
                    for (int i = 0; i < sumRows.Count; i++)
                    {
                        // 已通过柜号配对成功的箱柜直接跳过
                        if (matchedDetRows[i] > 0) continue;
                        // 提取汇总行 G 列单价公式
                        string gFormula = GetFormula(sumRows[i], 7);
                        // 校验公式是否为标准引用表达式
                        if (!string.IsNullOrEmpty(gFormula) && gFormula.StartsWith("="))
                        {
                            // 检查公式引用的行号是否落在某个未使用的明细块范围内
                            for (int j = 0; j < detRows.Count; j++)
                            {
                                if (detUsed[j]) continue;
                                // 获取当前明细块起始与结束区间
                                int detStart = detRows[j];
                                int detNext = (j + 1 < detRows.Count) ? detRows[j + 1] : (usedEndRow + 1);
                                // 探测公式中是否显式包含了当前明细块内部的行号引用
                                bool matchRef = System.Text.RegularExpressions.Regex.Matches(gFormula, @"\d+")
                                    .Cast<System.Text.RegularExpressions.Match>()
                                    .Any(m => int.TryParse(m.Value, out int r) && r >= detStart && r < detNext);

                                if (matchRef)
                                {
                                    // 容错匹配成功，锁定明细行
                                    matchedDetRows[i] = detRows[j];
                                    detUsed[j] = true;
                                    break;
                                }
                            }
                        }
                    }

                    // 第三轮：剩余未匹配箱柜按拓扑顺序保底对齐 (精准排除纯数字单价或空白无公式的纯汇总柜)
                    int nextDetIdx = 0;
                    for (int i = 0; i < sumRows.Count; i++)
                    {
                        // 已经配对的箱柜跳过
                        if (matchedDetRows[i] > 0) continue;
                        // 寻找下一个尚未使用的底表明细行
                        while (nextDetIdx < detRows.Count && detUsed[nextDetIdx]) nextDetIdx++;
                        // 若所有底表明细行均已分配完毕则终止
                        if (nextDetIdx >= detRows.Count) break;

                        // 提取汇总行 G 列单价数值与公式特征
                        string gVal = GetText(sumRows[i], 7);
                        string gFormula = GetFormula(sumRows[i], 7);
                        // 判断是否为纯数字直接录入或空白无公式 (纯汇总无明细箱柜的核心业务特征)
                        bool isPureNumberOrBlank = (double.TryParse(gVal, out _) || string.IsNullOrEmpty(gVal)) && string.IsNullOrEmpty(gFormula);

                        // 只有非纯数字且非空白单价的柜子才分配明细行；若存在多余的纯数字/空白单价且无对应底表，则视为纯汇总无明细柜保持为 0
                        if (!isPureNumberOrBlank || detRows.Count == sumRows.Count)
                        {
                            // 按顺序分配明细行
                            matchedDetRows[i] = detRows[nextDetIdx];
                            detUsed[nextDetIdx] = true;
                            nextDetIdx++;
                        }
                    }
                }
                else if (sumRows.Count == 0 && detRows.Count > 0)
                {
                    // 若无顶部汇总行，直接按明细行顺序对齐
                    for (int i = 0; i < detRows.Count; i++) matchedDetRows[i] = detRows[i];
                }

                // 规则 7: 循环外预加载系统公式方案库并倒序排列，避免每台箱柜重复反序列化磁盘 JSON
                List<Controllers.FormulaGroupModel>? sortedFeeGroups = null;
                try
                {
                    // 实例化公式控制器提取系统已登记的所有公式组方案
                    var feeCtrl = new Controllers.FormulaAdjustFeeController();
                    var allGroups = feeCtrl.GetFormulaGroups();
                    // 优先按明细项数量倒序比对，确保优先匹配更具体更长且完全吻合的方案
                    sortedFeeGroups = allGroups?.Where(g => g.Details != null && g.Details.Count >= 3)
                                               .OrderByDescending(g => g.Details!.Count)
                                               .ToList();
                }
                catch { }

                // 本地辅助函数：验证指定方案在当前总计行下是否 100% 逐项完全吻合 (严禁模糊猜测)
                bool TryMatchFormulaGroup(Controllers.FormulaGroupModel g, int tolsumRow, int detRow, out int feeStartRow)
                {
                    feeStartRow = 0;
                    // 校验方案明细项有效性
                    if (g.Details == null || g.Details.Count < 3) return false;
                    int N = g.Details.Count;
                    int candidateStart = tolsumRow - N + 1;
                    // 校验起始行合法性 (不能越过箱柜信息行和元器件起始行)
                    if (candidateStart < detRow + 2) return false;

                    int matchCount = 0;
                    // 逐行比对 B 列费用名称与序号特征
                    for (int idx = 0; idx < N; idx++)
                    {
                        // 方案预设费用名称与序号
                        string expName = g.Details[idx].Name?.Trim() ?? "";
                        string expNo = g.Details[idx].No?.Trim() ?? "";
                        // 单元格实际费用名称与序号 (覆盖 B 列与 A 列)
                        string actName = GetText(candidateStart + idx, 2);
                        string actNo = GetText(candidateStart + idx, 1);

                        // 判定单项是否 100% 吻合
                        bool isMatch = false;
                        // 规则 1: 若方案定义了费用名称且与实际单元格 B 列一致
                        if (!string.IsNullOrEmpty(expName) && string.Equals(actName, expName, StringComparison.OrdinalIgnoreCase))
                        {
                            isMatch = true;
                        }
                        // 规则 2: 总计行匹配 (方案项或序号含总计，或为末尾总计行，且实际 A/B 列含总计)
                        else if ((expNo.Contains("总计") || expName.Contains("总计") || idx == N - 1) &&
                                 (actNo.Contains("总计") || actName.Contains("总计")))
                        {
                            isMatch = true;
                        }
                        // 规则 3: 序号或项目代号严格完全一致
                        else if (!string.IsNullOrEmpty(expNo) && string.Equals(actNo, expNo, StringComparison.OrdinalIgnoreCase))
                        {
                            isMatch = true;
                        }

                        // 匹配计数累加
                        if (isMatch) matchCount++;
                    }

                    // 用户指示：严禁模糊猜测，必须 100% 逐项完全吻合 (matchCount == N) 才确认命中该方案
                    if (matchCount == N)
                    {
                        feeStartRow = candidateStart;
                        return true;
                    }
                    return false;
                }

                // 维护最近命中方案缓存 (高速验证通道)，在多台箱柜连续匹配时实现微秒级瞬时命中
                Controllers.FormulaGroupModel? lastMatchedGroup = null;

                // 4. 【逐个箱柜定位 Subsum (小计) 与 Tolsum (总计) 并覆盖绑定定义名称】
                for (int i = 0; i < cabCount; i++)
                {
                    // 箱柜序号从 1 开始递增
                    int k = i + 1;

                    // 确定当前箱柜对应的汇总行（若汇总行充足则对应取，否则按默认顺序排列）
                    int curSumRow = (i < sumRows.Count) ? sumRows[i] : (cabSumStartRow + i);
                    // 只要存在有效汇总行，无论是否有明细，均全量校准绑定 Cab_Sum_k
                    SafeSetSheetName(sheet, sheetName, $"{sumPrefix}{k}", curSumRow);

                    // 获取当前箱柜精准匹配到的明细行行号 (为 0 说明当前箱柜为纯汇总无明细箱柜)
                    int curDetRow = (i < matchedDetRows.Length) ? matchedDetRows[i] : 0;

                    // 只有当精准匹配到底表明细块时，才定位并绑定明细块的 3 个定义名称
                    if (curDetRow > 0)
                    {
                        // 动态计算当前明细块搜索下边界 (下一个大于当前行号的明细行，或者已用区域末尾)
                        int nextBoundaryRow = usedEndRow + 1;
                        foreach (int dr in detRows)
                        {
                            // 寻找大于当前明细行且距离最近的下一个明细行
                            if (dr > curDetRow && dr < nextBoundaryRow)
                            {
                                nextBoundaryRow = dr;
                            }
                        }

                        // 1. 在当前箱柜区间 [curDetRow + 2, nextBoundaryRow - 1] 内部由底向上寻找总计行 Cab_Tolsum
                        int curTolsumRow = 0;
                        int scanLimitRow = Math.Min(nextBoundaryRow - 1, usedEndRow);

                        // 优先按文本特征倒序扫描总计行 (支持 A/B/C 列合并单元格或文字位于 B 列的情况)
                        for (int r = scanLimitRow; r >= curDetRow + 2; r--)
                        {
                            // 提取前 3 列文本 (覆盖 A、B、C 列)
                            string aText = GetText(r, 1);
                            string bText = GetText(r, 2);
                            string cText = GetText(r, 3);

                            // 排除制表人、审核人、日期等落款说明行
                            bool isSignOff = aText.Contains("人") || bText.Contains("人") || aText.Contains("期") || bText.Contains("期") || aText.Contains("注") || bText.Contains("注");

                            // 只要 A/B/C 列明确包含“总计”且非落款说明行，立即锁定为总计行
                            if ((aText.Contains("总计") || bText.Contains("总计") || cText.Contains("总计")) && !isSignOff)
                            {
                                // 成功锁定总计行物理行号
                                curTolsumRow = r;
                                break;
                            }
                        }

                        // 辅助容错：若未标注“总计”文字，依据销售总价或单价公式的汇总乘算特征探测总计行
                        if (curTolsumRow == 0)
                        {
                            // 倒序扫描寻找包含公式引用的汇总行
                            for (int r = scanLimitRow; r >= curDetRow + 2; r--)
                            {
                                // 获取 H 列销售总价公式与 G 列单价公式
                                string hFormula = GetFormula(r, 8);
                                string gFormula = GetFormula(r, 7);

                                // 判定是否具备总计行的公式特征 (单台合计乘台数或引用前置合计行)
                                if ((!string.IsNullOrEmpty(hFormula) && hFormula.StartsWith("=") && hFormula.Contains("*")) ||
                                    (!string.IsNullOrEmpty(gFormula) && gFormula.StartsWith("=") && (gFormula.Contains("H") || gFormula.Contains("h"))))
                                {
                                    // 容错锁定总计行物理行号
                                    curTolsumRow = r;
                                    break;
                                }
                            }
                        }

                        // 2. 方案库指纹反查锁定计费首行 (第一重保险：100% 逐项完全吻合，支持高速缓存通道)
                        int curFeeStartRow = 0;

                        if (curTolsumRow > 0 && sortedFeeGroups != null && sortedFeeGroups.Count > 0)
                        {
                            try
                            {
                                // 高速通道：优先验证上一个箱柜命中的方案 (同一张图纸大多数箱柜方案完全一致)
                                if (lastMatchedGroup != null && TryMatchFormulaGroup(lastMatchedGroup, curTolsumRow, curDetRow, out int fastFeeStart))
                                {
                                    // 命中最近方案，直接采用
                                    curFeeStartRow = fastFeeStart;
                                }
                                else
                                {
                                    // 高速通道未命中或首次匹配，遍历按项数倒序排列的方案库
                                    foreach (var g in sortedFeeGroups)
                                    {
                                        // 跳过已验证过的最近方案
                                        if (g == lastMatchedGroup) continue;
                                        // 逐项严格 100% 完全比对方案
                                        if (TryMatchFormulaGroup(g, curTolsumRow, curDetRow, out int matchedFeeStart))
                                        {
                                            curFeeStartRow = matchedFeeStart;
                                            // 成功命中，更新最近方案缓存
                                            lastMatchedGroup = g;
                                            break;
                                        }
                                    }
                                }
                            }
                            catch { }

                            // 用户指示：若未能匹配到预设公式方案，则在 ABC 列包含“小计”的行作为计费区的第一行
                            if (curFeeStartRow == 0)
                            {
                                // 在当前箱柜区间 [curTolsumRow - 1 到 curDetRow + 2] 内部由底向上寻找小计行
                                for (int r = curTolsumRow - 1; r >= curDetRow + 2; r--)
                                {
                                    // 提取当前行 A、B、C 列文本
                                    string aText = GetText(r, 1);
                                    string bText = GetText(r, 2);
                                    string cText = GetText(r, 3);

                                    // 只要 A、B、C 列中任意一列包含“小计”，立即锁定为计费区第一行
                                    if (aText.Contains("小计") || bText.Contains("小计") || cText.Contains("小计"))
                                    {
                                        // 记录识别到的小计起始行
                                        curFeeStartRow = r;
                                        break;
                                    }
                                }

                                // 针对通过“小计”关键字兜底锁定的场景记录调试日志
                                if (curFeeStartRow > 0)
                                {
                                    LogHelper.WriteLog($"[小计特征锁定] 箱柜 [{k}] 未能匹配到预设方案，通过 ABC 列“小计”特征成功锁定计费首行为第 {curFeeStartRow} 行。");
                                }
                                else
                                {
                                    // 组装未匹配方案且未找到小计行的静默警告日志信息
                                    string warnMsg = $"箱柜 [{k}] (总计位于第 {curTolsumRow} 行) 未能匹配到任何预设方案且未找到“小计”行。";
                                    // 仅记录警告日志，严禁在此调用模态 MessageBox.Show 避免死锁
                                    LogHelper.WriteLog($"[费用方案匹配警告] {warnMsg}");
                                    // 将警告信息加入收集器，供上层 UI 统一优雅提示
                                    AddCabinetWarning($"【{sheetName}】箱柜 [{k}]：未能匹配到系统预设费用公式方案且未找到小计行");
                                }
                            }
                        }

                        // 3. 规范绑定定义名称与刷新公式
                        SafeSetSheetName(sheet, sheetName, $"{detPrefix}{k}", curDetRow);
                        if (curTolsumRow > 0)
                        {
                            SafeSetSheetName(sheet, sheetName, $"{tolsumPrefix}{k}", curTolsumRow);
                        }

                        int curSubsumRow = curFeeStartRow;
                        // 只有当精准匹配到方案起始行时，才安全校准 Cab_Subsum 定义名称
                        if (curFeeStartRow > 0 && curTolsumRow > 0)
                        {
                            SafeSetSheetName(sheet, sheetName, $"{subsumPrefix}{k}", curSubsumRow);
                        }

                        // 4. 建立/自愈汇总行与明细行双向超链接并保护居中与虚线框样式 (规则 6 架构规范)
                        try
                        {
                            // 汇总行 A 列单元格句柄
                            dynamic sumAnchor = sheet.Cells[curSumRow, 1];
                            string detTarget = $"'{sheetName}'!{detPrefix}{k}";

                            // 若已有超链接，就地更新目标地址，杜绝调用 Delete/Add 导致单元格边框与居中格式被 Excel 抹除
                            if (sumAnchor.Hyperlinks.Count > 0)
                            {
                                // 直接就地更新超链接跳转子地址与屏幕提示
                                dynamic hl = sumAnchor.Hyperlinks[1];
                                hl.SubAddress = detTarget;
                                hl.ScreenTip = "点击进入本箱柜明细表"; // --硬编码: 屏幕提示文本--
                            }
                            else
                            {
                                // 仅当单元格无超链接时才挂载新超链接
                                sheet.Hyperlinks.Add(
                                    Anchor: sumAnchor,
                                    Address: "",
                                    SubAddress: detTarget,
                                    ScreenTip: "点击进入本箱柜明细表" // --硬编码: 屏幕提示文本--
                                );
                            }

                            // 恢复汇总行 A 列自适应动态序号公式
                            sumAnchor.Formula = "=ROW()-ROW(A$6)"; // --硬编码: 公式表达式--

                            // 明细行 A 列单元格句柄
                            dynamic detAnchor = sheet.Cells[curDetRow, 1];
                            string sumTarget = $"'{sheetName}'!{sumPrefix}{k}";

                            // 就地更新明细行超链接目标
                            if (detAnchor.Hyperlinks.Count > 0)
                            {
                                dynamic hl = detAnchor.Hyperlinks[1];
                                hl.SubAddress = sumTarget;
                                hl.ScreenTip = "返回汇总行"; // --硬编码: 屏幕提示文本--
                            }
                            else
                            {
                                // 挂载明细行返回顶部超链接
                                sheet.Hyperlinks.Add(
                                    Anchor: detAnchor,
                                    Address: "",
                                    SubAddress: sumTarget,
                                    ScreenTip: "返回汇总行" // --硬编码: 屏幕提示文本--
                                );
                            }

                            // 明细行 A 列格式保护：去除下划线与恢复自动黑色
                            try
                            {
                                detAnchor.Font.Underline = -4142;  // --硬编码: xlUnderlineStyleNone 去除下划线--
                            }
                            catch { }
                        }
                        catch { }

                        // 确保汇总行 B 列与明细行 B 列无任何超链接 (严格遵守规则 6 规范，仅当存在时安全删除)
                        try { if (sheet.Cells[curSumRow, 2].Hyperlinks.Count > 0) sheet.Cells[curSumRow, 2].Hyperlinks.Delete(); } catch { }
                        try { if (sheet.Cells[curDetRow, 2].Hyperlinks.Count > 0) sheet.Cells[curDetRow, 2].Hyperlinks.Delete(); } catch { }
                    }
                    else
                    {
                        // 针对纯汇总无明细箱柜：安全清理可能残留的明细定义名称，保证定义名称纯净
                        SafeDeleteSheetName(sheet, $"{detPrefix}{k}");
                        SafeDeleteSheetName(sheet, $"{subsumPrefix}{k}");
                        SafeDeleteSheetName(sheet, $"{tolsumPrefix}{k}");

                        // 纯汇总无明细箱柜：清除超链接后同步自愈还原居中与虚线边框
                        try
                        {
                            dynamic sumAnchor = sheet.Cells[curSumRow, 1];
                            if (sumAnchor.Hyperlinks.Count > 0)
                            {
                                sumAnchor.Hyperlinks.Delete();
                            }
                            // 确保自适应序号动态公式正常
                            sumAnchor.Formula = "=ROW()-ROW(A$6)"; // --硬编码: 公式表达式--

                            // 格式自愈与保护：还原居中、去除下划线、恢复自动黑色字体
                            sumAnchor.Font.Underline = -4142;      // --硬编码: xlUnderlineStyleNone 去除下划线--

                            // 同步同行 B 列的虚线边框
                            try
                            {
                                dynamic bBorders = sheet.Cells[curSumRow, 2].Borders;
                                int bLineStyle = Convert.ToInt32(bBorders.LineStyle);
                                if (bLineStyle != 0 && bLineStyle != -4142)
                                {
                                    sumAnchor.Borders.LineStyle = bLineStyle;
                                    try { sumAnchor.Borders.Weight = bBorders.Weight; } catch { }
                                }
                                else
                                {
                                    sumAnchor.Borders.LineStyle = -4115; // --硬编码: xlDot 点线虚线--
                                    sumAnchor.Borders.Weight = 1;        // --硬编码: xlHairline 极细线--
                                }
                            }
                            catch { }
                        }
                        catch { }
                        // 安全清理 B 列超链接
                        try { if (sheet.Cells[curSumRow, 2].Hyperlinks.Count > 0) sheet.Cells[curSumRow, 2].Hyperlinks.Delete(); } catch { }
                    }
                }

                // 5. 【清理大于 cabCount 的多余旧箱柜定义名称，防止幽灵定义名称残留】
                for (int oldK = cabCount + 1; oldK <= cabCount + 30; oldK++)
                {
                    // 安全删除残留的汇总行与明细行定义名称
                    SafeDeleteSheetName(sheet, $"{sumPrefix}{oldK}");
                    SafeDeleteSheetName(sheet, $"{detPrefix}{oldK}");
                    SafeDeleteSheetName(sheet, $"{subsumPrefix}{oldK}");
                    SafeDeleteSheetName(sheet, $"{tolsumPrefix}{oldK}");
                }

                // 返回当前工作表校准绑定的箱柜总数量
                return cabCount;
            }
            catch (Exception ex)
            {
                // 记录工作表定义名称校准异常
                LogHelper.WriteLog($"工作表校准定义名称异常: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 安全删除工作表级别的指定定义名称
        /// </summary>
        /// <param name="sheet">目标工作表 COM 对象</param>
        /// <param name="tagName">定义名称字符串 (如 Cab_Det_1)</param>
        public static void SafeDeleteSheetName(dynamic sheet, string tagName)
        {
            // 校验入参工作表与标签有效性
            if (sheet == null || string.IsNullOrWhiteSpace(tagName)) return;
            try
            {
                // 尝试从工作表定义名称集合提取
                dynamic? existing = SafeGetSheetName(sheet, tagName);
                // 若存在则安全执行物理删除
                if (existing != null) existing.Delete();
            }
            catch { }
        }

        /// <summary>
        /// 安全设置/校准工作表级别的定义名称（若已存在则覆盖）
        /// </summary>
        public static void SafeSetSheetName(dynamic sheet, string sheetName, string tagName, int row)
        {
            try
            {
                // 尝试删除已有同名工作表级定义名称以实现干净校准覆盖
                try
                {
                    dynamic existing = sheet.Names.Item(tagName);
                    if (existing != null) existing.Delete();
                }
                catch { }

                // 添加工作表级别定义名称
                sheet.Names.Add(tagName, $"='{sheetName}'!$A${row}");
            }
            catch { }
        }

        /// <summary>
        /// 安全获取工作表级别的定义名称对象 (若不存在则安全返回 null)
        /// </summary>
        /// <param name="sheet">目标工作表 COM 对象</param>
        /// <param name="tagName">定义名称字符串 (如 Cab_Det_1)</param>
        /// <returns>找到的 Name 对象或 null</returns>
        public static dynamic? SafeGetSheetName(dynamic sheet, string tagName)
        {
            // 校验入参工作表与标签有效性
            if (sheet == null || string.IsNullOrWhiteSpace(tagName)) return null;
            try
            {
                // 尝试从工作表名称集合中提取
                return sheet.Names.Item(tagName);
            }
            catch
            {
                // 找不到或异常时安全回退 null
                return null;
            }
        }

        /// <summary>
        /// 安全从工作表与工作簿中删除指定的定义名称
        /// </summary>
        /// <param name="sheet">目标工作表 COM 对象</param>
        /// <param name="wb">目标工作簿 COM 对象 (可选)</param>
        /// <param name="tagName">待删除的定义名称字符串 (如 Cab_Sum_2)</param>
        public static void SafeDeleteName(dynamic sheet, dynamic? wb, string tagName)
        {
            // 校验待删除标签名是否有效
            if (string.IsNullOrWhiteSpace(tagName)) return;
            try
            {
                // 1. 尝试从工作表级定义名称集合中删除
                if (sheet != null && sheet.Names != null)
                {
                    // 查找工作表同名定义名称
                    try
                    {
                        // 获取工作表级名称项
                        dynamic existingSheetName = sheet.Names.Item(tagName);
                        // 若存在则执行删除
                        if (existingSheetName != null) existingSheetName.Delete();
                    }
                    catch { }
                }

                // 2. 尝试从工作簿级定义名称集合中删除
                if (wb != null && wb.Names != null)
                {
                    // 查找工作簿同名定义名称
                    try
                    {
                        // 获取工作簿级名称项
                        dynamic existingWbName = wb.Names.Item(tagName);
                        // 若存在则执行删除
                        if (existingWbName != null) existingWbName.Delete();
                    }
                    catch { }
                }
            }
            catch { }
        }


        /// <summary>
        /// 动态扫描工作表，根据箱柜序号 K 智能探测并返回该箱柜的标准行号分布 (优先匹配指定 K，未完整覆盖时直接回退取最后一个 K)
        /// 避免硬编码行号因不同模板产生偏移与 +1 错误
        /// </summary>
        /// <param name="sheet">目标工作表对象</param>
        /// <param name="cabinetK">目标箱柜序号 (默认 1)</param>
        /// <returns>指定箱柜的汇总行、明细信息行、小计行、总计行元组 (cabSumRow, cabDetRow, cabSubsumRow, cabTolsumRow)</returns>
        public static (int cabSumRow, int cabDetRow, int cabSubsumRow, int cabTolsumRow) FindStandardCategoryRowIndexes(dynamic sheet, int cabinetK = 1)
        {
            // 读取配置中的默认兜底值
            var cfg = ConfigManager.Instance.Current.Excel;
            int defSum = cfg.CabSumRowIndex;
            int defDet = cfg.CabDetRowIndex;
            int defTol = cfg.CabTolsumRowIndex;
            int defSub = defTol - 5;

            // 若工作表为空直接返回兜底默认值
            if (sheet == null) return (defSum, defDet, defSub, defTol);

            try
            {
                // 1. 获取工作表中已识别的所有有效箱柜集合 (显式强类型接收，避免 dynamic 传染)
                List<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets = GetSheetValidCabinets((object)sheet);
                // 校验是否存在有效箱柜
                if (validCabinets != null && validCabinets.Count > 0)
                {
                    // 1. 全表末尾探测模式 (cabinetK <= 0): 分别求取汇总行绝对最大值与明细块绝对最大值
                    if (cabinetK <= 0)
                    {
                        // 记录全表汇总行最大物理行号
                        int maxSum = 0;
                        // 记录全表明细信息行最大物理行号
                        int maxDet = 0;
                        // 记录全表总计行最大物理行号
                        int maxTol = 0;
                        // 记录全表小计行最大物理行号
                        int maxSub = 0;

                        // 遍历当前工作表所有已识别的有效箱柜 (无论有无明细)
                        foreach (var pair in validCabinets)
                        {
                            var anc = pair.Value;
                            // 探测汇总行最大值 (普通箱柜与无明细箱柜均参与)
                            if (anc.Sum != null)
                            {
                                int r = Convert.ToInt32(anc.Sum.Row);
                                if (r > maxSum) maxSum = r;
                            }
                            // 探测箱柜信息行最大值
                            if (anc.Det != null)
                            {
                                int r = Convert.ToInt32(anc.Det.Row);
                                if (r > maxDet) maxDet = r;
                            }
                            // 探测小计行最大值
                            if (anc.Subsum != null)
                            {
                                int r = Convert.ToInt32(anc.Subsum.Row);
                                if (r > maxSub) maxSub = r;
                            }
                            // 探测总计行最大值
                            if (anc.Tolsum != null)
                            {
                                int r = Convert.ToInt32(anc.Tolsum.Row);
                                if (r > maxTol) maxTol = r;
                            }
                        }

                        // 汇总行优先取全表最大值，确保新增箱柜始终在全表最后（无明细箱柜之后）追加
                        int resSum = maxSum > 0 ? maxSum : defSum;
                        int resDet = maxDet > 0 ? maxDet : defDet;
                        int resTol = maxTol > 0 ? maxTol : defTol;
                        int resSub = maxSub > 0 ? maxSub : (resTol - 5);
                        // 返回末尾行号元组
                        return (resSum, resDet, resSub, resTol);
                    }

                    // 2. 指定箱柜精确查询模式 (cabinetK > 0)
                    Models.CabinetAnchorModel? targetAnchor = null;
                    // 遍历寻找序号匹配的箱柜
                    foreach (var pair in validCabinets)
                    {
                        if (pair.Key == cabinetK)
                        {
                            targetAnchor = pair.Value;
                            break;
                        }
                    }

                    // 提取目标箱柜的 4 个物理行号
                    if (targetAnchor != null)
                    {
                        int sumR = targetAnchor.Sum != null ? Convert.ToInt32(targetAnchor.Sum.Row) : 0;
                        int detR = targetAnchor.Det != null ? Convert.ToInt32(targetAnchor.Det.Row) : 0;
                        int tolR = targetAnchor.Tolsum != null ? Convert.ToInt32(targetAnchor.Tolsum.Row) : 0;
                        int subR = targetAnchor.Subsum != null ? Convert.ToInt32(targetAnchor.Subsum.Row) : (tolR > 0 ? tolR - 5 : 0);

                        // 只要汇总行或明细行任一有效即返回，无明细箱柜 detR/tolR 保持为 0
                        if (sumR > 0 || detR > 0)
                        {
                            return (sumR > 0 ? sumR : defSum, detR, subR, tolR);
                        }
                    }
                }

                // 3. 若未识别到任何有效箱柜定义，直接回退配置默认基准值
                return (defSum, defDet, defSub, defTol);
            }
            catch (Exception ex)
            {
                // 记录异常并回退默认值
                LogHelper.WriteLog($"探测箱柜 {cabinetK} 标准行号分布异常: {ex.Message}");
                return (defSum, defDet, defSub, defTol);
            }
        }
    }
}
