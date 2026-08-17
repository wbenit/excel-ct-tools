using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ExcelAddInDemo.Controllers;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 系统通用工具类，提供全局通用的路径获取、目录检索等工具方法
    /// </summary>
    internal static class Tool
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

            // 提取或回退默认前缀
            sumPrefix = sumPrefix ?? ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
            detPrefix = detPrefix ?? ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";
            subsumPrefix = subsumPrefix ?? ConfigManager.Instance.Current.Excel.SubsumNamePrefix ?? "Cab_Subsum_";
            tolsumPrefix = tolsumPrefix ?? ConfigManager.Instance.Current.Excel.TolsumNamePrefix ?? "Cab_Tolsum_";

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

                // F 列 (索引 5): 数量 =IF(AND(B{row}="",C{row}=""),"",1)
                matrix[r, 5] = $"=IF(AND(B{currPhysicalRow}=\"\",C{currPhysicalRow}=\"\"),\"\",1)";

                // G 列 (索引 6): 销售单价 =IF(AND(B{row}="",C{row}=""),"",ROUND(M{row}*L{row}*N{row},2))
                matrix[r, 6] = $"=IF(AND(B{currPhysicalRow}=\"\",C{currPhysicalRow}=\"\"),\"\",ROUND(M{currPhysicalRow}*L{currPhysicalRow}*N{currPhysicalRow},2))";

                // H 列 (索引 7): 销售总价 =IF(AND(B{row}="",C{row}=""),"",ROUND(F{row}*G{row},2))
                matrix[r, 7] = $"=IF(AND(B{currPhysicalRow}=\"\",C{currPhysicalRow}=\"\"),\"\",ROUND(F{currPhysicalRow}*G{currPhysicalRow},2))";

                // I 列 (索引 8): 备注 (保留空字符串)
                matrix[r, 8] = string.Empty;

                // J 列 (索引 9): 成本单价 =IF(AND(B{row}="",C{row}=""),"",ROUND(M{row}*N{row},2))
                matrix[r, 9] = $"=IF(AND(B{currPhysicalRow}=\"\",C{currPhysicalRow}=\"\"),\"\",ROUND(M{currPhysicalRow}*N{currPhysicalRow},2))";

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
            }

            // 返回构建完成的元器件二维矩阵
            return matrix;
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

            // 过滤出至少拥有 Det 和 Sum 两个锚点的有效箱柜，按汇总行物理行号升序返回
            return cabinetDict
                .Where(x => x.Value.Det != null && x.Value.Sum != null)
                .OrderBy(x => (int)x.Value.Sum.Row)
                .ToList();
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
                // 获取 Excel 应用程序实例
                dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
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

                // 读取 4 种定义名称前缀配置项
                string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";
                string subsumPrefix = ConfigManager.Instance.Current.Excel.SubsumNamePrefix ?? "Cab_Subsum_";
                string tolsumPrefix = ConfigManager.Instance.Current.Excel.TolsumNamePrefix ?? "Cab_Tolsum_";

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
                // 特征条件：A 列包含“柜号”（或“箱柜”），且下一行 A 列包含“序号”
                var detRows = new List<int>();
                for (int r = cabSumStartRow + 1; r < usedEndRow; r++)
                {
                    // 提取当前行与下一行的 A 列文本
                    string aText = GetText(r, 1);
                    string nextAText = GetText(r + 1, 1);

                    // 匹配明细大标题与表头特征
                    if ((aText.Contains("柜号") || aText.Contains("箱柜") || aText.Contains("设备")) &&
                        (nextAText.Contains("序号") || nextAText.Contains("编号")))
                    {
                        // 记录识别到的箱柜信息行行号
                        detRows.Add(r);
                    }
                }

                // 若未识别出任何明细块，说明非标准分类表，跳过
                if (detRows.Count == 0) return 0;

                // 2. 【扫描顶部汇总行 Cab_Sum】
                // 起始于 cabSumStartRow，终止于首个明细行 detRows[0] 之前
                var sumRows = new List<int>();
                int firstDetRow = detRows[0];
                for (int r = cabSumStartRow; r < firstDetRow; r++)
                {
                    // 检查 B 列或 A 列是否有箱柜编号/名称
                    string bVal = GetText(r, 2);
                    string aVal = GetText(r, 1);

                    // 若存在非空内容则判定为有效汇总行
                    if (!string.IsNullOrWhiteSpace(bVal) || !string.IsNullOrWhiteSpace(aVal))
                    {
                        sumRows.Add(r);
                    }
                }

                // 箱柜总数以识别到的明细块数量为基准
                int cabCount = detRows.Count;

                // 3. 【逐个箱柜定位 Subsum (小计) 与 Tolsum (总计) 并覆盖绑定定义名称】
                for (int i = 0; i < cabCount; i++)
                {
                    // 箱柜序号从 1 开始递增
                    int k = i + 1;
                    int curDetRow = detRows[i];
                    int nextBoundaryRow = (i + 1 < detRows.Count) ? detRows[i + 1] : (usedEndRow + 1);

                    // 确定当前箱柜对应的汇总行（若汇总行充足则对应取，否则按默认顺序排列）
                    int curSumRow = (i < sumRows.Count) ? sumRows[i] : (cabSumStartRow + i);

                    // 寻找小计行 Cab_Subsum (规则: 含有公式且公式包含 SUM)
                    int curSubsumRow = 0;
                    // 寻找总计行 Cab_Tolsum (规则: A 列包含总计)
                    int curTolsumRow = 0;

                    // 在明细块区间内部寻找小计行与总计行
                    for (int r = curDetRow + 2; r < nextBoundaryRow; r++)
                    {
                        // 提取 A 列文本
                        string aText = GetText(r, 1);

                        // 优先检查小计行 (若未找到小计行且本行任意单元格公式含 SUM)
                        if (curSubsumRow == 0)
                        {
                            // 扫描前 12 列的公式内容
                            for (int c = 1; c <= Math.Min(arrCols, 12); c++)
                            {
                                string f = GetFormula(r, c);
                                // 判定公式中是否含有 SUM
                                if (!string.IsNullOrEmpty(f) && f.IndexOf("SUM", StringComparison.OrdinalIgnoreCase) >= 0 && f.IndexOf("INDEX", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    curSubsumRow = r;
                                    break;
                                }
                            }
                        }

                        // 检查总计行 (判断条件: 自身 A 列不含公式，且上一行 A 列公式包含 "ROW()-ROW(")
                        if (curTolsumRow == 0)
                        {
                            // 获取当前行 A 列公式文本
                            string curAFormula = GetFormula(r, 1);
                            // 获取上一行 A 列公式文本
                            string prevAFormula = GetFormula(r - 1, 1);

                            // 判断当前行 A 列无公式且上一行包含 ROW()-ROW(
                            if (string.IsNullOrEmpty(curAFormula) &&
                                !string.IsNullOrEmpty(prevAFormula) &&
                                prevAFormula.IndexOf("ROW()-ROW(", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // 记录当前识别到的总计行物理行号
                                curTolsumRow = r;
                            }
                        }

                        // 若小计与总计行均已确定，可提前结束当前箱柜区间的扫描
                        if (curSubsumRow > 0 && curTolsumRow > 0)
                        {
                            break;
                        }
                    }

                    // 兜底策略：若未识别到小计或总计行，按标准模板间距估算 --硬编码--
                    if (curTolsumRow == 0) curTolsumRow = curDetRow + 27;
                    if (curSubsumRow == 0) curSubsumRow = curTolsumRow - 3;

                    // 4. 【在工作表级别校准覆盖绑定 4 个定义名称】
                    SafeSetSheetName(sheet, sheetName, $"{sumPrefix}{k}", curSumRow);
                    SafeSetSheetName(sheet, sheetName, $"{detPrefix}{k}", curDetRow);
                    SafeSetSheetName(sheet, sheetName, $"{subsumPrefix}{k}", curSubsumRow);
                    SafeSetSheetName(sheet, sheetName, $"{tolsumPrefix}{k}", curTolsumRow);
                }

                // 返回当前工作表校准绑定的箱柜数量
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
        /// 动态扫描工作表，智能探测并返回首台箱柜的标准行号分布 (避免硬编码行号因不同模板产生偏移与 +1 错误)
        /// </summary>
        /// <param name="sheet">目标工作表对象</param>
        /// <returns>首台箱柜汇总行、明细信息行、小计行、总计行元组 (cabSumRow, cabDetRow, cabSubsumRow, cabTolsumRow)</returns>
        public static (int cabSumRow, int cabDetRow, int cabSubsumRow, int cabTolsumRow) FindStandardCategoryRowIndexes(dynamic sheet)
        {
            // 读取配置中的默认兜底值
            var cfg = ConfigManager.Instance.Current.Excel;
            int defSum = cfg.CabSumRowIndex;
            int defDet = cfg.CabDetRowIndex;
            int defTol = cfg.CabTolsumRowIndex;
            int defSub = defTol - 5;

            if (sheet == null) return (defSum, defDet, defSub, defTol);

            try
            {
                // 获取工作表已用区域
                dynamic usedRange = sheet.UsedRange;
                if (usedRange == null) return (defSum, defDet, defSub, defTol);

                // 提取已用区域起始行与总行数
                int startRow = Convert.ToInt32(usedRange.Row);
                int rowCount = Convert.ToInt32(usedRange.Rows.Count);
                int endRow = startRow + rowCount - 1;

                // 读取数值二维数组
                object[,]? valArray = usedRange.Value2 as object[,];
                if (valArray == null) return (defSum, defDet, defSub, defTol);

                int arrRows = valArray.GetLength(0);
                int arrCols = valArray.GetLength(1);

                // 本地快速文本获取辅助函数
                string GetText(int r, int c)
                {
                    int ar = r - startRow + 1;
                    if (ar < 1 || ar > arrRows || c < 1 || c > arrCols) return "";
                    return Convert.ToString(valArray[ar, c])?.Trim() ?? "";
                }

                int foundSumRow = 0;
                int foundDetRow = 0;
                int foundSubsumRow = 0;
                int foundTolsumRow = 0;

                // 1. 扫描顶部汇总表表头 (寻找包含 "序号" 且第2列包含 "柜号" 或 "箱柜" 的行)
                for (int r = startRow; r <= endRow; r++)
                {
                    string c1 = GetText(r, 1);
                    string c2 = GetText(r, 2);
                    if (c1.Contains("序号") && (c2.Contains("柜号") || c2.Contains("箱柜") || c2.Contains("产品")))
                    {
                        // 首个箱柜汇总行位于表头下一行
                        foundSumRow = r + 1;
                        break;
                    }
                }

                // 2. 扫描底部明细表箱柜信息行 (寻找 A 列包含 "柜号" 且下一行 A 列包含 "序号" 的行)
                for (int r = (foundSumRow > 0 ? foundSumRow + 1 : startRow + 1); r <= endRow; r++)
                {
                    string aText = GetText(r, 1);
                    string nextAText = GetText(r + 1, 1);
                    if ((aText.Contains("柜号") || aText.Contains("箱柜") || aText.Contains("设备")) &&
                        (nextAText.Contains("序号") || nextAText.Contains("编号") || nextAText.Contains("元件")))
                    {
                        foundDetRow = r;
                        break;
                    }
                }

                // 3. 在明细块内部寻找小计行与总计行
                if (foundDetRow > 0)
                {
                    for (int r = foundDetRow + 2; r <= Math.Min(endRow, foundDetRow + 60); r++)
                    {
                        string aText = GetText(r, 1);
                        string bText = GetText(r, 2);

                        // 识别小计行
                        if (foundSubsumRow == 0 && (aText.Contains("小计") || bText.Contains("小计")))
                        {
                            foundSubsumRow = r;
                        }

                        // 识别总计行
                        if (foundTolsumRow == 0 && (aText.Contains("总计") || bText.Contains("总计")))
                        {
                            foundTolsumRow = r;
                            break;
                        }
                    }
                }

                // 兜底与有效性校验
                if (foundSumRow <= 0) foundSumRow = defSum;
                if (foundDetRow <= 0) foundDetRow = defDet;
                if (foundTolsumRow <= 0) foundTolsumRow = defTol;
                if (foundSubsumRow <= 0) foundSubsumRow = foundTolsumRow - 5;

                // 返回识别出的 4 个基准行号
                return (foundSumRow, foundDetRow, foundSubsumRow, foundTolsumRow);
            }
            catch (Exception ex)
            {
                // 记录异常并回退默认值
                LogHelper.WriteLog($"探测分类标准行号分布异常: {ex.Message}");
                return (defSum, defDet, defSub, defTol);
            }
        }
    }
}
