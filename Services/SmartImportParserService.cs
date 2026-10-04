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
        /// <param name="maxCols">最大读取列数 (限制列数以节省内存)</param>
        /// <returns>网格预览数据包</returns>
        public static SmartImportPreviewResult ReadPreviewGrid(string filePath, string sheetName = "", int maxRows = 40, int maxCols = 15)
        {
            var result = new SmartImportPreviewResult();
            // 校验文件是否存在
            if (string.IsNullOrWhiteSpace(filePath) || !System.IO.File.Exists(filePath))
            {
                result.Success = false;
                result.Message = "指定的 Excel 标书文件不存在！";
                return result;
            }

            dynamic? app = null;
            dynamic? tempWb = null;
            // 备份原 Excel 弹窗与更新链接配置状态
            bool origDisplayAlerts = true;
            bool origAskToUpdateLinks = true;
            try
            {
                // 获取 Excel COM Application 接口实例
                app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    result.Success = false;
                    result.Message = "无法连接当前 Excel 应用程序实例！";
                    return result;
                }

                // 备份当前 Excel 提示与弹窗设置
                try
                {
                    origDisplayAlerts = app.DisplayAlerts;
                    origAskToUpdateLinks = app.AskToUpdateLinks;
                }
                catch { }

                // 强制静默禁用任何弹窗与外部链接更新询问，彻底杜绝“不安全外界源链接”弹窗
                try
                {
                    app.DisplayAlerts = false;
                    app.AskToUpdateLinks = false;
                }
                catch { }

                // 静默只读打开外部 Excel 工作簿，UpdateLinks: 0 彻底忽略外部链接更新
                tempWb = app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);
                if (tempWb == null)
                {
                    result.Success = false;
                    result.Message = "无法打开外部工作簿！";
                    return result;
                }

                // 收集所有有效工作表名称
                foreach (dynamic ws in tempWb.Worksheets)
                {
                    try { result.SheetNames.Add(Convert.ToString(ws.Name) ?? "Sheet"); } catch { }
                }

                // 确定目标读取的工作表
                dynamic targetSheet = null;
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
                    result.Success = false;
                    result.Message = "未找到有效的工作表！";
                    return result;
                }

                result.CurrentSheet = Convert.ToString(targetSheet.Name) ?? "";

                // 获取已使用数据区域
                dynamic usedRange = targetSheet.UsedRange;
                int totalRows = Convert.ToInt32(usedRange.Rows.Count);
                int totalCols = Convert.ToInt32(usedRange.Columns.Count);

                int limitRows = maxRows > 0 ? Math.Min(totalRows, maxRows) : totalRows;
                int limitCols = maxCols > 0 ? Math.Min(totalCols, maxCols) : totalCols;

                result.RowCount = limitRows;
                result.ColCount = limitCols;

                // 规则 7：一次性将整表数据全量读入内存二维数组，用于后续零 I/O 毫秒级推导
                object[,] fullMatrix = usedRange.Value2 as object[,];
                result.FullMatrix = fullMatrix;
                result.FullRowCount = totalRows;
                result.FullColCount = totalCols;

                // 转化为前端轻量字符串二维网格 (仅截取前 40 行展示，节省传输与渲染开销)
                for (int r = 1; r <= limitRows; r++)
                {
                    var rowList = new List<string>();
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
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"读取外部 Excel 发生异常: {ex.Message}";
                LogHelper.WriteLog($"[SmartImportParser] ReadPreviewGrid 异常: {ex.Message}");
            }
            finally
            {
                // 读取完毕立即关闭外部工作簿，严禁句柄泄露
                if (tempWb != null)
                {
                    try { tempWb.Close(false); } catch { }
                }
                // 安全恢复用户原有的 Excel 弹窗与设置
                if (app != null)
                {
                    try
                    {
                        app.DisplayAlerts = origDisplayAlerts;
                        app.AskToUpdateLinks = origAskToUpdateLinks;
                    }
                    catch { }
                }
            }

            return result;
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

            dynamic? app = null;
            dynamic? tempWb = null;
            // 备份原 Excel 弹窗与更新链接配置状态
            bool origDisplayAlerts = true;
            bool origAskToUpdateLinks = true;
            try
            {
                app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return cabinetList;

                // 备份当前状态
                try
                {
                    origDisplayAlerts = app.DisplayAlerts;
                    origAskToUpdateLinks = app.AskToUpdateLinks;
                }
                catch { }

                // 强制静默禁用任何弹窗与外部链接更新询问
                try
                {
                    app.DisplayAlerts = false;
                    app.AskToUpdateLinks = false;
                }
                catch { }

                // 静默只读打开外部工作簿
                tempWb = app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);
                dynamic targetSheet = string.IsNullOrWhiteSpace(sheetName) ? tempWb.Worksheets[1] : tempWb.Worksheets[sheetName];
                if (targetSheet == null) return cabinetList;

                // 获取全量已用区域
                dynamic usedRange = targetSheet.UsedRange;
                int totalRows = Convert.ToInt32(usedRange.Rows.Count);
                int totalCols = Convert.ToInt32(usedRange.Columns.Count);

                // 一次性读取整表内存二维数组 (规则 7)
                object[,] matrix = usedRange.Value2 as object[,];

                if (matrix == null) return cabinetList;

                // 核心推导算法：定位所有箱柜的表头行与结束行
                cabinetList = ScanCabinetsFromMatrix(matrix, totalRows, totalCols, config);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[SmartImportParser] ParseAllCabinets 异常: {ex.Message}");
            }
            finally
            {
                // 关闭外部临时工作簿
                if (tempWb != null)
                {
                    try { tempWb.Close(false); } catch { }
                }
                // 安全恢复用户原有的 Excel 弹窗与设置
                if (app != null)
                {
                    try
                    {
                        app.DisplayAlerts = origDisplayAlerts;
                        app.AskToUpdateLinks = origAskToUpdateLinks;
                    }
                    catch { }
                }
            }

            return cabinetList;
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

            // 2. 扫描整张二维表，寻找各个箱柜块
            int r = 1;
            while (r <= totalRows)
            {
                // 判断当前行是否匹配【表头行】特征
                if (IsMatchingHeaderPattern(matrix, r, colMap, sampleItemNameHeader, sampleItemSpecHeader))
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

                                components.Add(new ParsedComponentModel
                                {
                                    Index = compIndex++,
                                    ItemName = itemName,
                                    ItemSpec = itemSpec,
                                    Quantity = qty,
                                    Unit = unit,
                                    MarkedPrice = price,
                                    Brand = brand,
                                    Remark = remark
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
                            CabinetQuantity = cabQuantity,
                            Components = components,
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
        /// 优先比对样本表头原文本，其次结合配置的名称与规格关键字动态匹配
        /// </summary>
        private static bool IsMatchingHeaderPattern(object[,] matrix, int row, SmartImportColumnMapping colMap, string sampleNameHeader, string sampleSpecHeader)
        {
            // 提取当前行在配置的名称列与型号列的文本
            string nameVal = GetMatrixString(matrix, row, colMap.ItemNameCol);
            string specVal = GetMatrixString(matrix, row, colMap.ItemSpecCol);

            // 若与样本表头行原文字严格一致，直接判定命中
            if (!string.IsNullOrEmpty(sampleNameHeader) && nameVal.Equals(sampleNameHeader, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // 读取自定义关键字列表
            var nameKeys = SplitKeywords(colMap.ItemNameKeywords);
            var specKeys = SplitKeywords(colMap.ItemSpecKeywords);

            // 检查当前行单元格是否同时满足名称与型号关键词特征
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
