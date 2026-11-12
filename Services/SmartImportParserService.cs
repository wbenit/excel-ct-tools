using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// 外部 Excel 标书智能解析引擎
    /// 基于【双基线动态依附法则】执行高性能内存二维矩阵扫描与推导
    /// </summary>
    public static class SmartImportParserService
    {
        /// <summary>
        /// 后台静默读取外部 Excel 文件的工作表名称集合以及目标工作表网格数据
        /// </summary>
        /// <param name="filePath">外部 Excel 物理路径</param>
        /// <param name="sheetName">指定的工作表名称 (可为空，默认取第一个)</param>
        /// <param name="maxRows">最大读取行数 (预览模式限制行数，0为全量)</param>
        /// <param name="maxCols">最大读取列数 (默认设为 40 列，完整覆盖到 AF 列第 32 列)</param>
        /// <returns>网格预览数据包</returns>
        public static SmartImportPreviewResult ReadPreviewGrid(string filePath, string sheetName = "", int maxRows = 40, int maxCols = 40)
        {
            var result = new SmartImportPreviewResult();
            // 调用静默读取辅助引擎一次性拉取有效边界矩阵
            object[,]? fullMatrix = ReadSheetMatrixSilently(
                filePath,
                sheetName,
                out List<string> sheetNames,
                out string currentSheet,
                out int totalRows,
                out int totalCols,
                out string errorMessage);

            // 校验读取结果
            if (fullMatrix == null && !string.IsNullOrEmpty(errorMessage))
            {
                result.Success = false;
                result.Message = errorMessage;
                return result;
            }

            result.SheetNames = sheetNames;
            result.CurrentSheet = currentSheet;
            // 计算前端预览网格的截取行列上限 (仅传前 maxRows 行至前端渲染，降低 IPC 与 DOM 开销)
            int limitRows = maxRows > 0 ? Math.Min(totalRows, maxRows) : totalRows;
            int limitCols = maxCols > 0 ? Math.Min(totalCols, maxCols) : totalCols;

            result.RowCount = limitRows;
            result.ColCount = limitCols;
            // 规则 7：保存全量内存二维矩阵，用于后续零 I/O 毫秒级推导
            result.FullMatrix = fullMatrix;
            result.FullRowCount = totalRows;
            result.FullColCount = totalCols;

            // 转化为前端轻量字符串二维网格
            for (int r = 1; r <= limitRows; r++)
            {
                var rowList = new List<string>(limitCols);
                // 逐列提取预览文本
                for (int c = 1; c <= limitCols; c++)
                {
                    string cellStr = string.Empty;
                    if (fullMatrix != null && r <= totalRows && c <= totalCols)
                    {
                        object val = fullMatrix[r, c];
                        cellStr = val?.ToString()?.Trim() ?? string.Empty;
                    }
                    rowList.Add(cellStr);
                }
                result.GridData.Add(rowList);
            }

            result.Success = true;
            return result;
        }

        /// <summary>
        /// 统一静默读取外部 Excel 工作表的真实有效数据二维矩阵 (复用公共引擎，杜绝重复代码与 UsedRange 空白行膨胀卡顿)
        /// </summary>
        private static object[,]? ReadSheetMatrixSilently(
            string filePath,
            string sheetName,
            out List<string> sheetNames,
            out string currentSheet,
            out int totalRows,
            out int totalCols,
            out string errorMessage)
        {
            sheetNames = new List<string>();
            currentSheet = string.Empty;
            totalRows = 0;
            totalCols = 0;
            errorMessage = string.Empty;

            // 校验物理文件是否存在
            if (string.IsNullOrWhiteSpace(filePath) || !System.IO.File.Exists(filePath))
            {
                errorMessage = "指定的 Excel 标书文件不存在！";
                return null;
            }

            dynamic? app = null;
            dynamic? tempWb = null;
            // 备份 Excel 宿主原始环境配置 (包含自动计算状态，防止打开复杂标书触发全量计算卡顿)
            bool origDisplayAlerts = true;
            bool origAskToUpdateLinks = true;
            bool origScreenUpdating = true;
            bool origEnableEvents = true;
            int origCalculation = -4105; // 默认 xlCalculationAutomatic

            try
            {
                // 获取 Excel COM Application 接口实例
                app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    errorMessage = "无法连接当前 Excel 应用程序实例！";
                    return null;
                }

                // 备份原始属性
                try { origDisplayAlerts = app.DisplayAlerts; } catch { }
                try { origAskToUpdateLinks = app.AskToUpdateLinks; } catch { }
                try { origScreenUpdating = app.ScreenUpdating; } catch { }
                try { origEnableEvents = app.EnableEvents; } catch { }
                try { origCalculation = Convert.ToInt32(app.Calculation); } catch { }

                // 开启全静默极速模式：禁用弹窗、链接询问、屏幕刷新、事件以及自动重算
                try { app.DisplayAlerts = false; } catch { }
                try { app.AskToUpdateLinks = false; } catch { }
                try { app.ScreenUpdating = false; } catch { }
                try { app.EnableEvents = false; } catch { }
                try { app.Calculation = -4135; /* xlCalculationManual */ } catch { }

                // 静默只读打开外部 Excel 工作簿，UpdateLinks: 0 彻底忽略外部链接更新
                tempWb = app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);
                if (tempWb == null)
                {
                    errorMessage = "无法打开外部工作簿！";
                    return null;
                }

                // 收集所有有效工作表名称
                foreach (dynamic ws in tempWb.Worksheets)
                {
                    try { sheetNames.Add(Convert.ToString(ws.Name) ?? "Sheet"); } catch { }
                }

                // 确定目标读取的工作表
                dynamic? targetSheet = null;
                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    try { targetSheet = tempWb.Worksheets[sheetName]; } catch { }
                }
                if (targetSheet == null && tempWb.Worksheets.Count > 0)
                {
                    targetSheet = tempWb.Worksheets[1];
                }

                if (targetSheet == null)
                {
                    errorMessage = "未找到有效的工作表！";
                    return null;
                }

                currentSheet = Convert.ToString(targetSheet.Name) ?? "";

                // 精准探测真实非空数据边界，防范空白格式化导致 UsedRange 膨胀到数十万行引起卡死
                (int effRows, int effCols) = GetEffectiveSheetBounds((object)targetSheet);
                totalRows = effRows;
                totalCols = effCols;
                if (totalRows <= 0 || totalCols <= 0)
                {
                    // 空工作表直接返回 1x1 占位矩阵
                    totalRows = 0;
                    totalCols = 0;
                    return new object[0, 0];
                }

                // 规则 7：从 A1 起一次性将有效矩形数据全量读入内存二维数组，保持 1-based 行列坐标严格对齐
                dynamic readRange = targetSheet.Range[targetSheet.Cells[1, 1], targetSheet.Cells[totalRows, totalCols]];
                object? rawVal = readRange.Value2;
                if (rawVal is object[,] mat)
                {
                    return mat;
                }

                // 单单元格 (1x1) 兼容包装为 1-based 二维数组
                object[,] singleMat = (object[,])Array.CreateInstance(typeof(object), new int[] { 1, 1 }, new int[] { 1, 1 });
                singleMat[1, 1] = rawVal!;
                return singleMat;
            }
            catch (Exception ex)
            {
                errorMessage = $"读取外部 Excel 发生异常: {ex.Message}";
                LogHelper.WriteLog($"[SmartImportParser] ReadSheetMatrixSilently 异常: {ex.Message}");
                return null;
            }
            finally
            {
                // 读取完毕立即关闭外部工作簿，严禁句柄泄露
                if (tempWb != null)
                {
                    try { tempWb.Close(false); } catch { }
                }
                // 安全恢复用户原有的 Excel 运行环境设置
                if (app != null)
                {
                    try { app.Calculation = origCalculation; } catch { }
                    try { app.ScreenUpdating = origScreenUpdating; } catch { }
                    try { app.EnableEvents = origEnableEvents; } catch { }
                    try { app.DisplayAlerts = origDisplayAlerts; } catch { }
                    try { app.AskToUpdateLinks = origAskToUpdateLinks; } catch { }
                }
            }
        }

        /// <summary>
        /// 精准探测工作表中实际存在非空数据的最大行列边界，避免 UsedRange 含巨量空白格式化行导致读取卡死
        /// </summary>
        private static (int maxRow, int maxCol) GetEffectiveSheetBounds(dynamic sheet)
        {
            try
            {
                // 先读取 UsedRange 基础尺寸作为参考
                dynamic usedRange = sheet.UsedRange;
                int usedRows = Convert.ToInt32(usedRange.Row) + Convert.ToInt32(usedRange.Rows.Count) - 1;
                int usedCols = Convert.ToInt32(usedRange.Column) + Convert.ToInt32(usedRange.Columns.Count) - 1;

                // 当 UsedRange 尺寸在合理可控范围内时直接使用，减少额外的 Find COM 调用
                if (usedRows <= 2000 && usedCols <= 80)
                {
                    return (usedRows, usedCols);
                }

                // 超出常规尺寸时，使用 Excel 原生 Find 逆向探测最后一个真实非空单元格
                int lastRow = 0;
                int lastCol = 0;
                try
                {
                    // 按行逆向查找最后一个非空单元格 (xlValues=-4163, xlPart=2, xlByRows=1, xlPrevious=2)
                    dynamic? lastRowCell = sheet.Cells.Find(
                        What: "*",
                        LookIn: -4163,
                        LookAt: 2,
                        SearchOrder: 1,
                        SearchDirection: 2);
                    if (lastRowCell != null) lastRow = Convert.ToInt32(lastRowCell.Row);

                    // 按列逆向查找最后一个非空单元格 (xlByColumns=2, xlPrevious=2)
                    dynamic? lastColCell = sheet.Cells.Find(
                        What: "*",
                        LookIn: -4163,
                        LookAt: 2,
                        SearchOrder: 2,
                        SearchDirection: 2);
                    if (lastColCell != null) lastCol = Convert.ToInt32(lastColCell.Column);
                }
                catch
                {
                    // Find 失败时对超大区域做安全兜底截断，防止内存溢出
                    lastRow = Math.Min(usedRows, 20000);
                    lastCol = Math.Min(usedCols, 100);
                }

                // 若工作表完全为空
                if (lastRow <= 0 || lastCol <= 0) return (0, 0);
                // 返回收敛后的真实数据边界 (列数最多保底 100 列)
                return (lastRow, Math.Min(lastCol, 100));
            }
            catch
            {
                return (0, 0);
            }
        }

        /// <summary>
        /// 全量读取外部工作簿并执行双基线动态推导解析
        /// </summary>
        /// <param name="filePath">外部 Excel 物理路径</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="config">用户确认的解析规则配置</param>
        /// <returns>解析出的结构化箱柜列表</returns>
        public static List<ParsedCabinetModel> ParseAllCabinets(string filePath, string sheetName, SmartImportTemplateConfig config)
        {
            var cabinetList = new List<ParsedCabinetModel>();
            if (string.IsNullOrWhiteSpace(filePath) || !System.IO.File.Exists(filePath) || config == null)
            {
                return cabinetList;
            }

            // 复用统一静默矩阵读取引擎
            object[,]? matrix = ReadSheetMatrixSilently(
                filePath,
                sheetName,
                out _,
                out _,
                out int totalRows,
                out int totalCols,
                out _);

            if (matrix == null || totalRows <= 0 || totalCols <= 0)
            {
                return cabinetList;
            }

            // 核心推导算法：定位所有箱柜的表头行与结束行
            return ScanCabinetsFromMatrix(matrix, totalRows, totalCols, config);
        }

        /// <summary>
        /// 双基线动态依附扫描核心推导算法 (直接操作内存二维矩阵，零 I/O 毫秒级极速响应)
        /// </summary>
        public static List<ParsedCabinetModel> ScanCabinetsFromMatrix(object[,] matrix, int totalRows, int totalCols, SmartImportTemplateConfig config)
        {
            var list = new List<ParsedCabinetModel>();
            var colMap = config.ColumnMapping;

            // 1. 提取样本表头行的特征文本 (用于精准识别后续每个箱柜的表头行)
            string sampleItemNameHeader = GetMatrixString(matrix, config.SampleHeaderRow, colMap.ItemNameCol);
            string sampleItemSpecHeader = GetMatrixString(matrix, config.SampleHeaderRow, colMap.ItemSpecCol);
            string endKeyword = string.IsNullOrWhiteSpace(config.EndKeyword) ? "小计" : config.EndKeyword.Trim();
            // 提前在扫描循环外一次性预处理表头匹配关键字，避免逐行扫描时反复切割字符串分配内存
            List<string> nameKeys = SplitKeywords(colMap.ItemNameKeywords);
            List<string> specKeys = SplitKeywords(colMap.ItemSpecKeywords);

            // 2. 扫描整张二维表，寻找各个箱柜块
            int r = 1;
            while (r <= totalRows)
            {
                // 判断当前行是否匹配【表头行】特征
                if (IsMatchingHeaderPattern(matrix, r, colMap, sampleItemNameHeader, sampleItemSpecHeader, nameKeys, specKeys))
                {
                    int currentHeaderRow = r;
                    // 从表头下一行开始向下寻找【元器件结束行 (如小计)】
                    int currentEndRow = FindEndRow(matrix, currentHeaderRow + 1, totalRows, colMap, endKeyword);

                    if (currentEndRow > currentHeaderRow)
                    {
                        // 3. 根据双基线法则提取【柜号】
                        int cabNameRow = config.CabinetAnchor.CalculateRow(currentHeaderRow, currentEndRow);
                        string rawCabText = GetMatrixString(matrix, cabNameRow, config.CabinetAnchor.TargetCol);
                        string cleanCabName = CleanCabinetName(rawCabText, config.CabinetAnchor.PrefixKeyword);
                        if (string.IsNullOrWhiteSpace(cleanCabName))
                        {
                            cleanCabName = $"箱柜_{list.Count + 1}";
                        }

                        // 4. 根据双基线法则提取【箱柜台数】
                        int qtyRow = config.QuantityAnchor.CalculateRow(currentHeaderRow, currentEndRow);
                        string rawQtyText = GetMatrixString(matrix, qtyRow, config.QuantityAnchor.TargetCol);
                        decimal cabQuantity = ParseQuantity(rawQtyText, 1);

                        // 5. 提取可选箱柜属性：名称、型号、备注 (TargetCol <= 0 时自动跳过并返回空字符串)
                        string cabDisplayName = ExtractAnchorText(matrix, currentHeaderRow, currentEndRow, config.CabinetNameAnchor);
                        // 提取箱柜型号文本
                        string cabModel = ExtractAnchorText(matrix, currentHeaderRow, currentEndRow, config.CabinetModelAnchor);
                        // 提取箱柜备注文本
                        string cabRemark = ExtractAnchorText(matrix, currentHeaderRow, currentEndRow, config.CabinetRemarkAnchor);

                        // 5. 提取元器件明细：从 [表头行 + 1] 到 [结束行 - 1]
                        var components = new List<ParsedComponentModel>();
                        decimal totalMarked = 0;
                        int compIndex = 1;

                        for (int compRow = currentHeaderRow + 1; compRow < currentEndRow; compRow++)
                        {
                            string itemName = GetMatrixString(matrix, compRow, colMap.ItemNameCol);
                            string itemSpec = GetMatrixString(matrix, compRow, colMap.ItemSpecCol);
                            string qtyStr = GetMatrixString(matrix, compRow, colMap.QuantityCol);
                            decimal qty = ParseQuantity(qtyStr, 1);

                            // 品名或规格非空，且不包含小计/合计杂质
                            if ((!string.IsNullOrEmpty(itemName) || !string.IsNullOrEmpty(itemSpec)) && !itemName.Contains(endKeyword))
                            {
                                string unit = (colMap.UnitCol > 0) ? GetMatrixString(matrix, compRow, colMap.UnitCol) : "台";
                                if (string.IsNullOrWhiteSpace(unit)) unit = "台";

                                decimal price = (colMap.PriceCol > 0) ? ParseDecimal(GetMatrixString(matrix, compRow, colMap.PriceCol)) : 0;
                                string brand = (colMap.BrandCol > 0) ? GetMatrixString(matrix, compRow, colMap.BrandCol) : "";
                                string remark = (colMap.RemarkCol > 0) ? GetMatrixString(matrix, compRow, colMap.RemarkCol) : "";

                                // 提取动态扩展列文本集合 (最多 8 列)
                                var extraValues = new List<string>();
                                // 获取配置的动态扩展列集合
                                var extraCols = colMap.ExtraColumns;

                                // 若已配置动态扩展列，逐列安全读取
                                if (extraCols != null && extraCols.Count > 0)
                                {
                                    // 遍历扩展列配置集合
                                    for (int ei = 0; ei < extraCols.Count; ei++)
                                    {
                                        // 获取当前扩展列配置
                                        var extraCol = extraCols[ei];
                                        string currentVal = "";
                                        // 当来源列大于 0 时从矩阵提取文本
                                        if (extraCol != null && extraCol.SourceCol > 0)
                                        {
                                            // 从二维矩阵安全提取文本
                                            currentVal = GetMatrixString(matrix, compRow, extraCol.SourceCol);
                                            // 若目标列为单位列 (第 5 列) 且值有效，同步校准 unit
                                            if (extraCol.TargetCol == 5 && !string.IsNullOrWhiteSpace(currentVal))
                                            {
                                                unit = currentVal.Trim();
                                            }
                                            // 若目标列为备注列 (第 9 列) 且值有效，同步校准 remark
                                            else if (extraCol.TargetCol == 9 && !string.IsNullOrWhiteSpace(currentVal))
                                            {
                                                remark = currentVal.Trim();
                                            }
                                        }
                                        // 存入当前元件的动态扩展值列表
                                        extraValues.Add(currentVal);
                                    }
                                }

                                components.Add(new ParsedComponentModel
                                {
                                    Index = compIndex++,
                                    ItemName = itemName,
                                    ItemSpec = itemSpec,
                                    Quantity = qty,
                                    Unit = unit,
                                    MarkedPrice = price,
                                    Brand = brand,
                                    Remark = remark,
                                    // 记录所有动态扩展列提取的值集合 (最多 8 列)
                                    ExtraValues = extraValues
                                });

                                totalMarked += (qty * price);
                            }
                        }

                        // 封装箱柜对象并加入结果集
                        list.Add(new ParsedCabinetModel
                        {
                            SourceStartRow = Math.Min(cabNameRow, currentHeaderRow),
                            SourceEndRow = Math.Max(qtyRow, currentEndRow),
                            CabinetName = cleanCabName,
                            // 写入可选箱柜扩展字段（名称、型号、备注）
                            CabinetDisplayName = cabDisplayName,
                            CabinetModel = cabModel,
                            CabinetRemark = cabRemark,
                            CabinetQuantity = cabQuantity,
                            Components = components,
                            // 显式记录元器件真实总项数
                            TotalComponentCount = components.Count,
                            TotalMarkedPrice = totalMarked
                        });

                        // 游标推进至当前箱柜最底端行，避免重复扫描
                        r = Math.Max(currentEndRow, qtyRow);
                    }
                }

                r++;
            }

            return list;
        }

        /// <summary>
        /// 判断某一行是否为箱柜的表头行
        /// 优先比对样本表头原文本，其次结合预处理后的名称与规格关键字动态匹配
        /// </summary>
        private static bool IsMatchingHeaderPattern(
            object[,] matrix,
            int row,
            SmartImportColumnMapping colMap,
            string sampleNameHeader,
            string sampleSpecHeader,
            List<string> nameKeys,
            List<string> specKeys)
        {
            // 提取当前行在配置的名称列与型号列的文本
            string nameVal = GetMatrixString(matrix, row, colMap.ItemNameCol);
            string specVal = GetMatrixString(matrix, row, colMap.ItemSpecCol);

            // 若与样本表头行原文字严格一致，直接判定命中
            if (!string.IsNullOrEmpty(sampleNameHeader) && nameVal.Equals(sampleNameHeader, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // 检查当前行单元格是否同时满足名称与型号关键词特征 (直接使用预拆分的列表，零内存重复分配)
            bool nameMatched = MatchesAny(nameVal, nameKeys);
            bool specMatched = MatchesAny(specVal, specKeys);

            return nameMatched && specMatched;
        }

        /// <summary>
        /// 根据用户配置的关键词词典，从表头行各单元格文本中自动启发式嗅探 5 列目标列号
        /// </summary>
        public static SmartImportColumnMapping AutoDetectColumnsFromHeaderRow(List<string> headerRowTexts, SmartImportColumnMapping? baseMap = null)
        {
            // 初始化或复用基准列映射对象
            var map = baseMap ?? new SmartImportColumnMapping();
            if (headerRowTexts == null || headerRowTexts.Count == 0) return map;

            // 拆分各列关键词列表
            var nameKeys = SplitKeywords(map.ItemNameKeywords);
            var specKeys = SplitKeywords(map.ItemSpecKeywords);
            var qtyKeys = SplitKeywords(map.QuantityKeywords);
            var priceKeys = SplitKeywords(map.PriceKeywords);
            var brandKeys = SplitKeywords(map.BrandKeywords);

            // 遍历表头行各单元格执行智能匹配
            for (int c = 0; c < headerRowTexts.Count; c++)
            {
                int colIndex = c + 1; // 1-based 列号
                string text = headerRowTexts[c]?.Trim() ?? "";
                if (string.IsNullOrEmpty(text)) continue;

                // 优先匹配单价 (主动避开合价与总价干扰)
                if (MatchesAny(text, priceKeys) && !text.Contains("合价") && !text.Contains("总价"))
                {
                    map.PriceCol = colIndex;
                }
                // 匹配品牌/厂家列
                else if (MatchesAny(text, brandKeys))
                {
                    map.BrandCol = colIndex;
                }
                // 匹配数量列
                else if (MatchesAny(text, qtyKeys))
                {
                    map.QuantityCol = colIndex;
                }
                // 匹配规格型号列
                else if (MatchesAny(text, specKeys))
                {
                    map.ItemSpecCol = colIndex;
                }
                // 匹配元件名称列
                else if (MatchesAny(text, nameKeys))
                {
                    map.ItemNameCol = colIndex;
                }
            }

            return map;
        }

        /// <summary>
        /// 拆分逗号、分号分隔的关键字配置字符串
        /// </summary>
        private static List<string> SplitKeywords(string raw)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(raw)) return list;
            // 支持中英文逗号、分号与坚线分隔符
            var parts = raw.Split(new char[] { ',', '，', '|', ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                string t = p.Trim();
                if (!string.IsNullOrEmpty(t)) list.Add(t);
            }
            return list;
        }

        /// <summary>
        /// 检查文本是否命中任意关键字
        /// </summary>
        private static bool MatchesAny(string text, List<string> keywords)
        {
            if (string.IsNullOrEmpty(text) || keywords == null) return false;
            foreach (var k in keywords)
            {
                if (text.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            return false;
        }

        /// <summary>
        /// 从表头行下一行向下寻找元器件结束行 (如小计行)
        /// </summary>
        private static int FindEndRow(object[,] matrix, int startRow, int totalRows, SmartImportColumnMapping colMap, string endKeyword)
        {
            for (int r = startRow; r <= totalRows; r++)
            {
                // 检查品名列、序号列、型号列是否出现结束关键词
                string valName = GetMatrixString(matrix, r, colMap.ItemNameCol);
                string valIndex = (colMap.IndexCol > 0) ? GetMatrixString(matrix, r, colMap.IndexCol) : "";
                string valSpec = GetMatrixString(matrix, r, colMap.ItemSpecCol);

                if (valName.Contains(endKeyword) || valIndex.Contains(endKeyword) || valSpec.Contains(endKeyword))
                {
                    return r;
                }

                // 若遇到下一个明显的新表头，提前终止防呆
                if (valName.Contains("元件名称") || valName.Contains("品名"))
                {
                    return r;
                }
            }

            // 未找到明确小计，返回总行数防呆
            return totalRows;
        }

        /// <summary>
        /// 统一根据可选动态锚点提取并清洗文本字段
        /// </summary>
        private static string ExtractAnchorText(object[,] matrix, int headerRow, int endRow, DynamicFieldAnchor? anchor)
        {
            // 未配置锚点或目标列号无效时，直接返回空字符串
            if (anchor == null || anchor.TargetCol <= 0) return string.Empty;
            // 计算当前箱柜对应的源表行号
            int targetRow = anchor.CalculateRow(headerRow, endRow);
            // 从内存二维数组中安全读取原始字符串
            string rawText = GetMatrixString(matrix, targetRow, anchor.TargetCol);
            // 复用柜号清洗逻辑剥离冒号及指定前缀
            return CleanCabinetName(rawText, anchor.PrefixKeyword);
        }

        /// <summary>
        /// 清洗剥离柜号前缀
        /// </summary>
        private static string CleanCabinetName(string rawText, string prefix)
        {
            if (string.IsNullOrWhiteSpace(rawText)) return string.Empty;

            string text = rawText.Trim();
            // 若带有冒号，剥离冒号前缀
            if (text.Contains("：") || text.Contains(":"))
            {
                var parts = text.Split(new char[] { '：', ':' }, 2);
                if (parts.Length > 1 && !string.IsNullOrWhiteSpace(parts[1]))
                {
                    return parts[1].Trim();
                }
            }

            // 若指定了前缀文本
            if (!string.IsNullOrWhiteSpace(prefix) && text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                text = text.Substring(prefix.Length).Trim();
            }

            return text;
        }

        /// <summary>
        /// 安全读取二维矩阵字符串值
        /// </summary>
        private static string GetMatrixString(object[,] matrix, int row, int col)
        {
            if (matrix == null || row < 1 || col < 1) return string.Empty;
            int maxR = matrix.GetLength(0);
            int maxC = matrix.GetLength(1);
            if (row > maxR || col > maxC) return string.Empty;

            try
            {
                object val = matrix[row, col];
                return val?.ToString()?.Trim() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 解析数量数值，带有单位剥离与安全默认值
        /// </summary>
        private static decimal ParseQuantity(string text, decimal defaultValue = 1)
        {
            if (string.IsNullOrWhiteSpace(text)) return defaultValue;

            // 通过正则提取纯数字 (包括浮点数)
            var match = Regex.Match(text, @"[0-9]+(\.[0-9]+)?");
            if (match.Success && decimal.TryParse(match.Value, out decimal val))
            {
                return val > 0 ? val : defaultValue;
            }

            return defaultValue;
        }

        /// <summary>
        /// 解析金额数值
        /// </summary>
        private static decimal ParseDecimal(string text, decimal defaultValue = 0)
        {
            if (string.IsNullOrWhiteSpace(text)) return defaultValue;
            var match = Regex.Match(text, @"[0-9]+(\.[0-9]+)?");
            if (match.Success && decimal.TryParse(match.Value, out decimal val))
            {
                return val;
            }
            return defaultValue;
        }
    }
}
