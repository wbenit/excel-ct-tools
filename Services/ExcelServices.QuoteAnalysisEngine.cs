using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 报价全景分析高级解析引擎
    /// 支持多工作表智能探测、深度明细块提取、双环旭日图与矩形树图多维度数据建模
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 扫描当前活动工作簿中所有具备报价清单结构的工作表信息
        /// </summary>
        /// <returns>可用报价工作表列表</returns>
        public static List<QuoteSheetInfoDto> ListQuoteSheets()
        {
            var sheetList = new List<QuoteSheetInfoDto>();
            try
            {
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return sheetList;

                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null) return sheetList;

                string activeSheetName = Convert.ToString(app.ActiveSheet?.Name)?.Trim() ?? string.Empty;

                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    string sName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(sName) || string.Equals(sName, "项目信息", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    dynamic used = sheet.UsedRange;
                    if (used == null) continue;

                    object[,] values = used.Value2 as object[,];
                    if (values == null) continue;

                    int rowCount = values.GetLength(0);
                    int colCount = values.GetLength(1);

                    // 检查前 15 行是否存在“柜号”或“箱柜名称”表头特征
                    int headerRow = FindQuoteHeaderRow(values, rowCount, colCount);
                    if (headerRow > 0)
                    {
                        // 统计该表有效箱柜数量
                        int count = CountValidCabinetsInSheet(values, rowCount, colCount, headerRow);
                        sheetList.Add(new QuoteSheetInfoDto
                        {
                            Name = sName,
                            CabinetCount = count,
                            IsActive = string.Equals(sName, activeSheetName, StringComparison.OrdinalIgnoreCase)
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[报价分析引擎] 扫描报价工作表异常: {ex.Message}");
            }
            return sheetList;
        }

        /// <summary>
        /// 核心分析调度入口：根据选中的工作表列表执行全景数据解析与多层级图表建模
        /// </summary>
        /// <param name="selectedSheets">用户选定的工作表名称列表，为 null 或空时默认采用当前活动工作表</param>
        /// <returns>报价全景分析传输对象</returns>
        public static QuoteAnalysisDto ExecuteQuoteAnalysis(List<string>? selectedSheets = null)
        {
            var dto = new QuoteAnalysisDto();

            try
            {
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return dto;

                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null) return dto;

                // 1. 扫描工作簿全部可用报价工作表
                var allAvailable = ListQuoteSheets();
                dto.AvailableSheets = allAvailable;

                // 2. 确定本次实际参与解析的目标工作表集合
                string activeSheetName = Convert.ToString(app.ActiveSheet?.Name)?.Trim() ?? string.Empty;
                var targetSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (selectedSheets != null && selectedSheets.Count > 0)
                {
                    foreach (var s in selectedSheets)
                    {
                        if (!string.IsNullOrWhiteSpace(s)) targetSheetNames.Add(s.Trim());
                    }
                }

                // 若未指定或指定列表为空，默认选取当前活动表；若活动表不是有效报价表，则选第一个报价表
                if (targetSheetNames.Count == 0)
                {
                    var activeMatch = allAvailable.FirstOrDefault(x => x.IsActive);
                    if (activeMatch != null)
                    {
                        targetSheetNames.Add(activeMatch.Name);
                    }
                    else if (allAvailable.Count > 0)
                    {
                        targetSheetNames.Add(allAvailable[0].Name);
                    }
                }

                dto.SelectedSheets = targetSheetNames.ToList();

                // 3. 读取项目宏观概况 (从【项目信息】工作表提取项目名、报价单号与公司名称)
                string fallbackProjName = Path.GetFileNameWithoutExtension(Convert.ToString(activeWb.Name) ?? "");
                dto.Summary.ProjectName = fallbackProjName;

                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    string sName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;
                    if (string.Equals(sName, "项目信息", StringComparison.OrdinalIgnoreCase))
                    {
                        ParseProjectInfoSheet(sheet, dto.Summary);
                        break;
                    }
                }

                // 4. 扫描定义名称映射 (Cab_Det_K 快速定位)
                var definedNameRowMap = ExtractDefinedNamesMap(activeWb);

                // 5. 遍历并解析选中的每一个报价工作表
                var allCabinetItems = new List<CabinetAnalysisItemDto>();

                // 全局明细汇总聚合容器
                // 生产厂家 -> 元件名称 -> 采购合价
                var vendorComponentMap = new Dictionary<string, Dictionary<string, double>>(StringComparer.OrdinalIgnoreCase);
                // 类别Q列 -> 元件名称 -> (总价, 数量)
                var categoryTreemapDict = new Dictionary<string, Dictionary<string, (double TotalMoney, double Qty)>>(StringComparer.OrdinalIgnoreCase);
                // 核心元器件物料列表
                var allComponentList = new List<ComponentRankingDto>();

                // 各种成本费用全局累计
                double totalTopCost = 0;       // 顶部J列成本总价
                double totalTopGrossProfit = 0; // 顶部K列毛利
                double totalTopQuote = 0;      // 顶部H列报出总价

                double detailCompCost = 0;     // Q=元件成本
                double detailBoxCost = 0;      // Q=箱体成本
                double detailLaborCost = 0;    // Q=人工费成本
                double detailWireCost = 0;     // Q=电线/材料成本
                double detailAuxCost = 0;      // Q=辅材成本
                double detailCompleteSet = 0;  // 综合成套费
                double detailTax = 0;          // 规费与税金

                // 方案B：严格按照Q列字面量累计成本，Q列为“费用”则采用B列名称累计
                var costByCategoryMap = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
                var feeByNameMap = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    string sName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;
                    if (!targetSheetNames.Contains(sName)) continue;

                    dynamic used = sheet.UsedRange;
                    if (used == null) continue;

                    object[,] values = used.Value2 as object[,];
                    if (values == null) continue;

                    int rowCount = values.GetLength(0);
                    int colCount = values.GetLength(1);

                    int headerRow = FindQuoteHeaderRow(values, rowCount, colCount);
                    if (headerRow <= 0) continue;

                    // 5.1 解析该表的顶部箱柜汇总行 (过滤数量为0的箱柜)
                    var sheetCabinets = ParseTopSummaryCabinets(values, rowCount, colCount, headerRow, sName, definedNameRowMap);

                    // 5.2 扫描并关联下方每台箱柜的明细块
                    ParseSheetDetailBlocks(values, rowCount, colCount, headerRow, sName, sheetCabinets,
                        vendorComponentMap, categoryTreemapDict, allComponentList,
                        ref detailCompCost, ref detailBoxCost, ref detailLaborCost, ref detailWireCost,
                        ref detailAuxCost, ref detailCompleteSet, ref detailTax,
                        costByCategoryMap, feeByNameMap);

                    // 累计顶部金额
                    foreach (var cab in sheetCabinets)
                    {
                        totalTopCost += cab.CostPrice;
                        totalTopGrossProfit += cab.GrossProfit;
                        totalTopQuote += cab.TotalPrice;
                    }

                    allCabinetItems.AddRange(sheetCabinets);
                }

                dto.Cabinets = allCabinetItems;

                // 6. 聚合财务核心 KPI
                dto.Summary.TotalCabinetCount = allCabinetItems.Count;
                dto.Summary.TotalQuotePrice = Math.Round(totalTopQuote, 2);
                dto.Summary.TotalCostPrice = Math.Round(totalTopCost, 2);
                dto.Summary.TotalGrossProfit = Math.Round(totalTopGrossProfit, 2);
                if (dto.Summary.TotalQuotePrice > 0)
                {
                    dto.Summary.GrossProfitRate = Math.Round(dto.Summary.TotalGrossProfit / dto.Summary.TotalQuotePrice, 4);
                }

                // 7. 组装成本构成五大支柱
                dto.CostStructure = new CostBreakdownDto
                {
                    ComponentsTotal = Math.Round(detailCompCost, 2),
                    CabinetBoxTotal = Math.Round(detailBoxCost, 2),
                    LaborTotal = Math.Round(detailLaborCost, 2),
                    AuxMaterialsTotal = Math.Round(detailAuxCost + detailWireCost, 2),
                    CompleteSetTotal = Math.Round(detailCompleteSet, 2),
                    ProfitTotal = Math.Round(dto.Summary.TotalGrossProfit, 2)
                };

                // 8. 计算明细成本与顶部J列汇总之间的核对差额 (基于严格Q列成本字典总和)
                double totalDetailCostSum = costByCategoryMap.Values.Sum();
                if (totalDetailCostSum <= 0)
                {
                    totalDetailCostSum = detailCompCost + detailBoxCost + detailLaborCost + detailWireCost + detailAuxCost;
                }
                dto.ReconcileDiff = Math.Round(totalTopCost - totalDetailCostSum, 2);

                // 9. 构建双环嵌套旭日图：成本结构与利润解构 (方案B：外环严格对齐Q列字面量，费用项采用B列名称)
                dto.CostSunburst = BuildCostProfitSunburst(totalTopCost, totalTopGrossProfit,
                    costByCategoryMap, feeByNameMap);

                // 10. 构建双环嵌套旭日图：生产厂家与元件分布 (内环D列厂家，外环B列元件名称)
                dto.VendorSunburst = BuildVendorSunburst(vendorComponentMap);

                // 11. 构建供应链核心物料矩形树图 (第一层Q列类别，第二层B列元件名称)
                dto.ComponentTreemap = BuildComponentTreemap(categoryTreemapDict);

                // 12. 抽取前 30 核心物料排行
                dto.TopComponents = BuildTopComponentsRanking(allComponentList, out int totalCount);
                dto.Summary.TotalComponentCount = totalCount;

                // 13. 统计品牌体量分布
                dto.BrandDistributions = BuildBrandDistributions(vendorComponentMap);

                // 14. 统计箱柜功能分类分布
                dto.CategoryDistributions = BuildCategoryDistributions(allCabinetItems, dto.Summary.TotalQuotePrice);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[报价分析引擎] 全景数据解析异常: {ex.Message}\r\n{ex.StackTrace}");
            }

            return dto;
        }

        /// <summary>
        /// 在工作表前 15 行中探测包含“柜号”或“箱柜名称”的表头行号
        /// </summary>
        private static int FindQuoteHeaderRow(object[,] values, int rowCount, int colCount)
        {
            for (int r = 1; r <= Math.Min(rowCount, 15); r++)
            {
                string colB = Convert.ToString(values[r, 2])?.Trim() ?? "";
                string colC = Convert.ToString(values[r, 3])?.Trim() ?? "";
                if (colB == "柜号" || colC == "箱柜名称" || colB == "箱柜名称" || (colB.Contains("柜号") && colB.Contains("名称")))
                {
                    return r;
                }
            }
            return 0;
        }

        /// <summary>
        /// 快速统计工作表顶部有效箱柜台数 (过滤数量为0的行)
        /// </summary>
        private static int CountValidCabinetsInSheet(object[,] values, int rowCount, int colCount, int headerRow)
        {
            int count = 0;
            for (int r = headerRow + 1; r <= rowCount; r++)
            {
                string cabName = Convert.ToString(values[r, 2])?.Trim() ?? "";
                string seqStr = Convert.ToString(values[r, 1])?.Trim() ?? "";

                // 遇到空行或明细起点或小计行结束
                if (string.IsNullOrEmpty(cabName) || cabName.StartsWith("柜号:") || cabName.StartsWith("柜号：") ||
                    seqStr == "小计" || cabName == "小计" || seqStr.Contains("合计") || cabName.Contains("合计"))
                {
                    break;
                }

                double qty = ConvertToDouble(values[r, 6]);
                if (qty > 0)
                {
                    count++;
                }
            }
            return count;
        }

        /// <summary>
        /// 解析工作表顶部箱柜汇总行 (严格过滤数量为0的箱柜)
        /// </summary>
        private static List<CabinetAnalysisItemDto> ParseTopSummaryCabinets(object[,] values, int rowCount, int colCount,
            int headerRow, string sheetName, Dictionary<string, int> definedNameRowMap)
        {
            var list = new List<CabinetAnalysisItemDto>();
            int cabSeq = 1;

            for (int r = headerRow + 1; r <= rowCount; r++)
            {
                string seqStr = Convert.ToString(values[r, 1])?.Trim() ?? "";
                string cabCode = Convert.ToString(values[r, 2])?.Trim() ?? ""; // B 列柜号
                string cabNameC = Convert.ToString(values[r, 3])?.Trim() ?? ""; // C 列箱柜名称

                // 汇总区结束条件
                if (string.IsNullOrEmpty(cabCode) || cabCode.StartsWith("柜号:") || cabCode.StartsWith("柜号：") ||
                    seqStr == "小计" || cabCode == "小计" || seqStr.Contains("合计") || cabCode.Contains("合计") ||
                    seqStr.Contains("包装费") || seqStr.Contains("运输费"))
                {
                    break;
                }

                // 数量 F 列 (第 6 列)
                double qty = ConvertToDouble(values[r, 6]);
                // 严格规则：过滤数量等于 0 的未选或作废箱柜
                if (qty <= 0)
                {
                    continue;
                }

                if (!int.TryParse(seqStr, out int seqNum))
                {
                    seqNum = cabSeq;
                }

                string cabModel = Convert.ToString(values[r, 4])?.Trim() ?? "";
                if (string.IsNullOrEmpty(cabModel)) cabModel = cabNameC;

                double unitPrice = ConvertToDouble(values[r, 7]);
                double totalPrice = ConvertToDouble(values[r, 8]);
                if (totalPrice <= 0 && unitPrice > 0) totalPrice = unitPrice * qty;

                double costPrice = ConvertToDouble(values[r, 10]);   // J 列成本总价
                double grossProfit = ConvertToDouble(values[r, 11]); // K 列毛利
                double profitRate = ConvertToDouble(values[r, 12]);  // L 列毛利率
                string size = Convert.ToString(values[r, 13])?.Trim() ?? ""; // M 列尺寸

                // 用户明确规则：按 C 列原始名称分类，不进行任何人工逻辑判断或篡改
                string categoryC = !string.IsNullOrWhiteSpace(cabNameC) ? cabNameC : (!string.IsNullOrWhiteSpace(cabCode) ? cabCode : "未命名箱柜");

                var cabItem = new CabinetAnalysisItemDto
                {
                    Index = seqNum,
                    SheetName = sheetName,
                    Name = cabCode, // B 列柜号 (如 1AP1) 用于明细块匹配与帕累托排布
                    Model = cabModel,
                    Quantity = qty,
                    UnitPrice = unitPrice,
                    TotalPrice = totalPrice,
                    CostPrice = costPrice,
                    GrossProfit = grossProfit,
                    GrossProfitRate = profitRate,
                    Size = size,
                    InstallType = DetectInstallType(cabCode, size),
                    FunctionalCategory = categoryC // 直接采用 C 列原始名称分类
                };

                // 关联定义名称行号锚点
                string detTagKey = $"{sheetName}!Cab_Det_{seqNum}";
                if (definedNameRowMap.TryGetValue(detTagKey, out int detRow))
                {
                    cabItem.DetRow = detRow;
                }
                else if (definedNameRowMap.TryGetValue($"Cab_Det_{seqNum}", out int globalDetRow))
                {
                    cabItem.DetRow = globalDetRow;
                }

                list.Add(cabItem);
                cabSeq++;
            }

            return list;
        }

        /// <summary>
        /// 扫描工作表下方全部“柜号:”明细块，提取各物料明细行并累计五大维度数据 (方案B：严格支持Q列字面量与B列费用名称聚合)
        /// </summary>
        private static void ParseSheetDetailBlocks(object[,] values, int rowCount, int colCount, int headerRow,
            string sheetName, List<CabinetAnalysisItemDto> sheetCabinets,
            Dictionary<string, Dictionary<string, double>> vendorComponentMap,
            Dictionary<string, Dictionary<string, (double TotalMoney, double Qty)>> categoryTreemapDict,
            List<ComponentRankingDto> allComponentList,
            ref double detailCompCost, ref double detailBoxCost, ref double detailLaborCost,
            ref double detailWireCost, ref double detailAuxCost, ref double detailCompleteSet, ref double detailTax,
            Dictionary<string, double> costByCategoryMap,
            Dictionary<string, double> feeByNameMap)
        {
            // 1. 扫描所有明细块的起点行
            var blockStarts = new List<(int Row, string CabCode)>();
            for (int r = headerRow + 1; r <= rowCount; r++)
            {
                string colA = Convert.ToString(values[r, 1])?.Trim() ?? "";
                if (colA.StartsWith("柜号:") || colA.StartsWith("柜号：") || colA == "柜号:")
                {
                    string code = Convert.ToString(values[r, 2])?.Trim() ?? "";
                    if (string.IsNullOrEmpty(code))
                    {
                        int colonIdx = colA.IndexOfAny(new[] { ':', '：' });
                        if (colonIdx >= 0 && colonIdx < colA.Length - 1)
                        {
                            code = colA.Substring(colonIdx + 1).Trim();
                        }
                    }
                    blockStarts.Add((r, code));
                }
            }

            if (blockStarts.Count == 0) return;

            // 建立柜号到箱柜对象的映射字典
            var cabCodeMap = new Dictionary<string, CabinetAnalysisItemDto>(StringComparer.OrdinalIgnoreCase);
            foreach (var cab in sheetCabinets)
            {
                if (!cabCodeMap.ContainsKey(cab.Name))
                {
                    cabCodeMap[cab.Name] = cab;
                }
            }

            // 2. 逐块遍历明细区
            for (int b = 0; b < blockStarts.Count; b++)
            {
                int startRow = blockStarts[b].Row;
                int endRow = (b + 1 < blockStarts.Count) ? blockStarts[b + 1].Row - 1 : rowCount;
                string blockCabCode = blockStarts[b].CabCode;

                // 寻找匹配的箱柜对象
                CabinetAnalysisItemDto? matchedCab = null;
                if (!string.IsNullOrEmpty(blockCabCode) && cabCodeMap.TryGetValue(blockCabCode, out var mCab))
                {
                    matchedCab = mCab;
                }
                else if (b < sheetCabinets.Count)
                {
                    matchedCab = sheetCabinets[b];
                }

                if (matchedCab != null)
                {
                    if (matchedCab.DetRow <= 0) matchedCab.DetRow = startRow;
                    matchedCab.DetBlockStartRow = startRow;
                    matchedCab.DetBlockEndRow = endRow;
                }

                // 寻找明细表头行 (通常在 startRow+1 行，B列为“元件名称”)
                int itemHeaderRow = 0;
                for (int r = startRow; r <= Math.Min(endRow, startRow + 5); r++)
                {
                    string bText = Convert.ToString(values[r, 2])?.Trim() ?? "";
                    if (bText == "元件名称" || bText.Contains("元件名称"))
                    {
                        itemHeaderRow = r;
                        break;
                    }
                }

                if (itemHeaderRow <= 0) continue;

                // 扫描块内明细物料行
                for (int r = itemHeaderRow + 1; r <= endRow; r++)
                {
                    string name = Convert.ToString(values[r, 2])?.Trim() ?? ""; // B 列元件名称
                    if (string.IsNullOrEmpty(name)) continue;

                    // 严格过滤小计、单台合计、总计等汇总行
                    if (name == "小计" || name == "单台合计" || name == "总计" || name == "合计" ||
                        name.StartsWith("总计") || name.StartsWith("单台"))
                    {
                        continue;
                    }

                    // Q 列类别 (第 17 列)
                    string categoryQ = colCount >= 17 ? (Convert.ToString(values[r, 17])?.Trim() ?? "") : "";

                    // 严格规则：过滤 Q 列为空的行 (样本中空值行多为小计中间金额，排除后彻底杜绝重复)
                    if (string.IsNullOrEmpty(categoryQ))
                    {
                        continue;
                    }

                    string model = Convert.ToString(values[r, 3])?.Trim() ?? "";      // C 列型号规格
                    string vendor = Convert.ToString(values[r, 4])?.Trim() ?? "";     // D 列生产厂家
                    if (string.IsNullOrEmpty(vendor)) vendor = "自制/通用";

                    double qty = ConvertToDouble(values[r, 6]);                        // F 列数量
                    double unitPrice = ConvertToDouble(values[r, 7]);                  // G 列单价
                    double totPrice = ConvertToDouble(values[r, 8]);                   // H 列总价
                    if (totPrice <= 0 && qty > 0 && unitPrice > 0) totPrice = qty * unitPrice;

                    double costUnitPrice = ConvertToDouble(values[r, 10]);             // J 列成本单价
                    double costTotalPrice = ConvertToDouble(values[r, 11]);            // K 列成本总价
                    if (costTotalPrice <= 0 && costUnitPrice > 0 && qty > 0) costTotalPrice = costUnitPrice * qty;

                    // 1. 累计到具体箱柜的五维成本
                    if (matchedCab != null)
                    {
                        if (categoryQ == "元件") matchedCab.ComponentsCost += costTotalPrice;
                        else if (categoryQ == "箱体") matchedCab.CabinetBoxCost += costTotalPrice;
                        else if (categoryQ == "辅材") matchedCab.AuxMaterialsCost += costTotalPrice;
                        else if (categoryQ == "人工费") matchedCab.LaborCost += costTotalPrice;
                        else if (categoryQ == "费用" && name.Contains("成套")) matchedCab.CompleteSetFee += totPrice;
                    }

                    // 2. 累计全局成本费用指标 (保持向后兼容)
                    if (categoryQ == "元件") detailCompCost += costTotalPrice;
                    else if (categoryQ == "箱体") detailBoxCost += costTotalPrice;
                    else if (categoryQ == "人工费") detailLaborCost += costTotalPrice;
                    else if (categoryQ == "电线" || categoryQ == "材料") detailWireCost += costTotalPrice;
                    else if (categoryQ == "辅材") detailAuxCost += costTotalPrice;
                    else if (categoryQ == "费用")
                    {
                        if (name.Contains("税金")) detailTax += totPrice;
                        else detailCompleteSet += totPrice;
                    }

                    // 方案B：严格对齐Q列字面量；若Q列为“费用”，采用B列名称
                    if (categoryQ == "费用")
                    {
                        string feeName = string.IsNullOrWhiteSpace(name) ? "其它费用" : name.Trim();
                        double feeAmount = totPrice > 0 ? totPrice : costTotalPrice;
                        if (!feeByNameMap.ContainsKey(feeName)) feeByNameMap[feeName] = 0;
                        feeByNameMap[feeName] += feeAmount;
                    }
                    else
                    {
                        // 成本项严格采用Q列原始文本 (如“元件”、“箱体”、“人工费”、“电线”、“材料”、“辅材”)
                        if (!costByCategoryMap.ContainsKey(categoryQ)) costByCategoryMap[categoryQ] = 0;
                        costByCategoryMap[categoryQ] += costTotalPrice;
                    }

                    // 3. 累计生产厂家 -> 元件名称字典 (用于 VendorSunburst 旭日图 / 玫瑰图)
                    if (!vendorComponentMap.ContainsKey(vendor))
                    {
                        vendorComponentMap[vendor] = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
                    }
                    if (!vendorComponentMap[vendor].ContainsKey(name))
                    {
                        vendorComponentMap[vendor][name] = 0;
                    }
                    vendorComponentMap[vendor][name] += totPrice;

                    // 4. 累计类别Q -> 元件名称字典 (用于 ComponentTreemap 矩形树图)
                    if (!categoryTreemapDict.ContainsKey(categoryQ))
                    {
                        categoryTreemapDict[categoryQ] = new Dictionary<string, (double, double)>(StringComparer.OrdinalIgnoreCase);
                    }
                    if (!categoryTreemapDict[categoryQ].ContainsKey(name))
                    {
                        categoryTreemapDict[categoryQ][name] = (0, 0);
                    }
                    var curTreeVal = categoryTreemapDict[categoryQ][name];
                    categoryTreemapDict[categoryQ][name] = (curTreeVal.TotalMoney + totPrice, curTreeVal.Qty + qty);

                    // 5. 记录核心器件清单
                    if (categoryQ == "元件" || categoryQ == "材料")
                    {
                        allComponentList.Add(new ComponentRankingDto
                        {
                            Name = name,
                            Model = model,
                            Manufacturer = vendor,
                            TotalQuantity = qty,
                            UnitPrice = unitPrice,
                            TotalPrice = totPrice
                        });
                    }
                }
            }
        }

        /// <summary>
        /// 构建成本结构与利润解构双环嵌套旭日图树形模型 (方案B：外环严格对齐Q列字面量，费用项采用B列名称)
        /// </summary>
        private static SunburstNodeDto BuildCostProfitSunburst(
            double totalCost, 
            double totalGrossProfit,
            Dictionary<string, double> costByCategoryMap,
            Dictionary<string, double> feeByNameMap)
        {
            var root = new SunburstNodeDto { Name = "全案价值体系", Value = Math.Round(totalCost + totalGrossProfit, 2) };

            // 内环 1: 制造成本 (J列成本总价合计)
            var costNode = new SunburstNodeDto
            {
                Name = "制造成本",
                Value = Math.Round(totalCost, 2)
            };

            // 外环：严格对齐 Q 列字面量 (如 "元件"、"箱体"、"电线"、"材料"、"人工费"、"辅材")，按金额降序排列
            foreach (var kvp in costByCategoryMap.OrderByDescending(x => x.Value))
            {
                if (kvp.Value > 0)
                {
                    costNode.Children.Add(new SunburstNodeDto
                    {
                        Name = kvp.Key, // 严格使用Q列原始文本
                        Value = Math.Round(kvp.Value, 2)
                    });
                }
            }

            // 若明细拆解为0兜底推导 (保持Q列字面量)
            if (costNode.Children.Count == 0 && totalCost > 0)
            {
                costNode.Children.Add(new SunburstNodeDto { Name = "元件", Value = Math.Round(totalCost * 0.65, 2) });
                costNode.Children.Add(new SunburstNodeDto { Name = "箱体", Value = Math.Round(totalCost * 0.18, 2) });
                costNode.Children.Add(new SunburstNodeDto { Name = "人工费", Value = Math.Round(totalCost * 0.10, 2) });
                costNode.Children.Add(new SunburstNodeDto { Name = "材料", Value = Math.Round(totalCost * 0.05, 2) });
                costNode.Children.Add(new SunburstNodeDto { Name = "辅材", Value = Math.Round(totalCost * 0.02, 2) });
            }

            // 内环 2: 毛利润总额 (K列毛利合计)
            double cleanProfit = Math.Max(0, totalGrossProfit);
            var profitNode = new SunburstNodeDto
            {
                Name = "毛利润总额",
                Value = Math.Round(cleanProfit, 2)
            };

            // 外环：Q列为“费用”时采用B列名称 (如“综合成套服务费”、“规费及税金”)
            double totalFees = 0;
            foreach (var kvp in feeByNameMap.OrderByDescending(x => x.Value))
            {
                if (kvp.Value > 0)
                {
                    totalFees += kvp.Value;
                    profitNode.Children.Add(new SunburstNodeDto
                    {
                        Name = kvp.Key, // 严格采用B列名称
                        Value = Math.Round(kvp.Value, 2)
                    });
                }
            }

            // 扣除费用后的净毛利
            double netProfit = Math.Max(0, cleanProfit - totalFees);
            if (netProfit > 0 || profitNode.Children.Count == 0)
            {
                profitNode.Children.Add(new SunburstNodeDto
                {
                    Name = "净毛利",
                    Value = Math.Round(netProfit, 2)
                });
            }

            root.Children.Add(costNode);
            root.Children.Add(profitNode);

            return root;
        }

        /// <summary>
        /// 构建生产厂家与元件分布双环旭日图数据树 (内环D列厂家，外环B列元件名称)
        /// </summary>
        private static SunburstNodeDto BuildVendorSunburst(Dictionary<string, Dictionary<string, double>> vendorMap)
        {
            var root = new SunburstNodeDto { Name = "全部生产厂家" };

            // 计算所有厂家的总额
            double grandTotal = vendorMap.Values.SelectMany(v => v.Values).Sum();
            if (grandTotal <= 0) grandTotal = 1;

            // 按厂家总合价降序排序
            var sortedVendors = vendorMap
                .Select(kv => new
                {
                    Vendor = kv.Key,
                    Total = kv.Value.Values.Sum(),
                    Items = kv.Value
                })
                .OrderByDescending(v => v.Total)
                .ToList();

            foreach (var vGroup in sortedVendors)
            {
                if (vGroup.Total <= 0) continue;

                var vendorNode = new SunburstNodeDto
                {
                    Name = vGroup.Vendor,
                    Value = Math.Round(vGroup.Total, 2)
                };

                // 该厂家下的元件排序
                var sortedItems = vGroup.Items.OrderByDescending(x => x.Value).ToList();
                double otherSum = 0;

                foreach (var item in sortedItems)
                {
                    // 小于该厂家总额 2% 且超过前 8 项的合并为其他
                    if (item.Value / vGroup.Total < 0.02 && vendorNode.Children.Count >= 6)
                    {
                        otherSum += item.Value;
                    }
                    else
                    {
                        vendorNode.Children.Add(new SunburstNodeDto
                        {
                            Name = item.Key,
                            Value = Math.Round(item.Value, 2)
                        });
                    }
                }

                if (otherSum > 0)
                {
                    vendorNode.Children.Add(new SunburstNodeDto
                    {
                        Name = "其它元件",
                        Value = Math.Round(otherSum, 2)
                    });
                }

                root.Children.Add(vendorNode);
            }

            root.Value = Math.Round(grandTotal, 2);
            return root;
        }

        /// <summary>
        /// 构建供应链元件名称矩形树图 (第一层Q列类别，第二层B列元件名称)
        /// </summary>
        private static SunburstNodeDto BuildComponentTreemap(Dictionary<string, Dictionary<string, (double TotalMoney, double Qty)>> treemapDict)
        {
            var root = new SunburstNodeDto { Name = "供应链物料" };

            // 类别名称中文化润色映射
            var catNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "元件", "电气元器件" },
                { "箱体", "箱体壳体结构" },
                { "材料", "主材材料" },
                { "电线", "铜排一次导线" },
                { "辅材", "二次配线辅料" },
                { "人工费", "工时制作装配" },
                { "费用", "综合管理服务" }
            };

            foreach (var catKv in treemapDict)
            {
                string catKey = catKv.Key;
                string displayCat = catNameMap.TryGetValue(catKey, out var friendly) ? friendly : catKey;

                double catTotal = catKv.Value.Values.Sum(x => x.TotalMoney);
                if (catTotal <= 0) continue;

                var catNode = new SunburstNodeDto
                {
                    Name = displayCat,
                    Value = Math.Round(catTotal, 2)
                };

                var sortedItems = catKv.Value.OrderByDescending(x => x.Value.TotalMoney).ToList();
                double otherSum = 0;
                double otherQty = 0;

                foreach (var item in sortedItems)
                {
                    // 占比极小项合并为“其他”
                    if (item.Value.TotalMoney / catTotal < 0.015 && catNode.Children.Count >= 8)
                    {
                        otherSum += item.Value.TotalMoney;
                        otherQty += item.Value.Qty;
                    }
                    else
                    {
                        catNode.Children.Add(new SunburstNodeDto
                        {
                            Name = item.Key,
                            Value = Math.Round(item.Value.TotalMoney, 2),
                            Quantity = Math.Round(item.Value.Qty, 1),
                            Extra = $"{displayCat} ｜ {item.Key}"
                        });
                    }
                }

                if (otherSum > 0)
                {
                    catNode.Children.Add(new SunburstNodeDto
                    {
                        Name = "其它细项",
                        Value = Math.Round(otherSum, 2),
                        Quantity = Math.Round(otherQty, 1),
                        Extra = $"{displayCat} 其它细碎物料"
                    });
                }

                root.Children.Add(catNode);
            }

            root.Value = Math.Round(root.Children.Sum(c => c.Value), 2);
            return root;
        }

        /// <summary>
        /// 聚合全案核心器件排行榜并按采购合价降序排序
        /// </summary>
        private static List<ComponentRankingDto> BuildTopComponentsRanking(List<ComponentRankingDto> allList, out int totalCount)
        {
            totalCount = allList.Count;
            if (allList.Count == 0) return new List<ComponentRankingDto>();

            // 按元件名称 + 型号进行合并汇总
            var aggregated = allList
                .GroupBy(c => $"{c.Name}|{c.Model}|{c.Manufacturer}")
                .Select(g => new ComponentRankingDto
                {
                    Name = g.First().Name,
                    Model = g.First().Model,
                    Manufacturer = g.First().Manufacturer,
                    TotalQuantity = Math.Round(g.Sum(x => x.TotalQuantity), 1),
                    UnitPrice = g.First().UnitPrice,
                    TotalPrice = Math.Round(g.Sum(x => x.TotalPrice), 2)
                })
                .OrderByDescending(c => c.TotalPrice)
                .Take(30)
                .ToList();

            double grandSum = aggregated.Sum(c => c.TotalPrice);
            if (grandSum <= 0) grandSum = 1;

            int rank = 1;
            foreach (var item in aggregated)
            {
                item.Rank = rank++;
                item.Percentage = Math.Round(item.TotalPrice / grandSum, 4);
            }

            return aggregated;
        }

        /// <summary>
        /// 构建品牌体量采购分布列表
        /// </summary>
        private static List<BrandDistributionDto> BuildBrandDistributions(Dictionary<string, Dictionary<string, double>> vendorMap)
        {
            var list = new List<BrandDistributionDto>();
            double total = vendorMap.Values.SelectMany(v => v.Values).Sum();
            if (total <= 0) total = 1;

            foreach (var kv in vendorMap)
            {
                double sum = kv.Value.Values.Sum();
                if (sum <= 0) continue;

                list.Add(new BrandDistributionDto
                {
                    BrandName = kv.Key,
                    TotalPrice = Math.Round(sum, 2),
                    Percentage = Math.Round(sum / total, 4)
                });
            }

            return list.OrderByDescending(b => b.TotalPrice).ToList();
        }

        /// <summary>
        /// 构建箱柜名称分类分布统计 (按 C 列原始名称分类，不作任何逻辑修改)
        /// </summary>
        private static List<CategoryDistributionDto> BuildCategoryDistributions(List<CabinetAnalysisItemDto> cabinets, double totalMoney)
        {
            if (totalMoney <= 0) totalMoney = 1;

            return cabinets
                .GroupBy(c => c.FunctionalCategory)
                .Select(g => new CategoryDistributionDto
                {
                    CategoryName = g.Key,
                    CabinetCount = (int)Math.Round(g.Sum(c => c.Quantity)), // 物理总台数
                    TotalPrice = Math.Round(g.Sum(c => c.TotalPrice), 2),
                    Percentage = Math.Round(g.Sum(c => c.TotalPrice) / totalMoney, 4)
                })
                .OrderByDescending(c => c.TotalPrice)
                .ToList();
        }
    }
}
