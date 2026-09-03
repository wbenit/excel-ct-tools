using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using ExcelAddInDemo.Models;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 核心业务服务分部类：智能辅材、壳体、人工与铜排计算引擎
    /// </summary>
    public static partial class ExcelServices
    {
        // 规则配置文件默认存储路径
        private static readonly string RulesConfigFilePath = Path.Combine(Tool.GetAppDataDirectory(), "quotation_rules.json");

        // 缓存的计算定额规则实例
        private static QuotationRules? _cachedRules;

        // JSON 序列化选项 (支持中文美化缩进)
        private static readonly JsonSerializerOptions RuleJsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// 加载计算定额与规则配置 (优先从磁盘 JSON 读取，不存在则创建默认配置)
        /// </summary>
        /// <returns>QuotationRules 配置实例</returns>
        public static QuotationRules LoadQuotationRules()
        {
            // 判断内存缓存是否已存在
            if (_cachedRules != null)
            {
                // 直接返回已缓存的规则实例
                return _cachedRules;
            }

            try
            {
                // 判断磁盘上是否存在配置文件
                if (File.Exists(RulesConfigFilePath))
                {
                    // 读取磁盘文件中的 JSON 文本
                    string json = File.ReadAllText(RulesConfigFilePath);
                    // 反序列化为强类型配置对象
                    var rules = JsonSerializer.Deserialize<QuotationRules>(json, RuleJsonOptions);
                    // 校验解析结果
                    if (rules != null)
                    {
                        // 缓存并返回规则对象
                        _cachedRules = rules;
                        return _cachedRules;
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录配置文件读取异常信息
                System.Diagnostics.Debug.WriteLine($"读取规则配置发生异常: {ex.Message}");
            }

            // 若读取失败或文件不存在，生成并持久化默认规则配置
            _cachedRules = new QuotationRules();
            // 保存默认规则至磁盘
            SaveQuotationRules(_cachedRules);
            // 返回默认规则对象
            return _cachedRules;
        }

        /// <summary>
        /// 保存并持久化计算定额规则配置至磁盘
        /// </summary>
        /// <param name="rules">待保存的规则模型</param>
        public static void SaveQuotationRules(QuotationRules rules)
        {
            // 校验输入对象有效性
            if (rules == null) return;
            try
            {
                // 将配置对象序列化为格式化 JSON 字符串
                string json = JsonSerializer.Serialize(rules, RuleJsonOptions);
                // 确保 AppData 所在目录存在
                string dir = Path.GetDirectoryName(RulesConfigFilePath) ?? Tool.GetAppDataDirectory();
                if (!Directory.Exists(dir))
                {
                    // 创建目录
                    Directory.CreateDirectory(dir);
                }
                // 写入磁盘文件
                File.WriteAllText(RulesConfigFilePath, json);
                // 同步更新内存缓存
                _cachedRules = rules;
            }
            catch (Exception ex)
            {
                // 记录保存异常信息
                System.Diagnostics.Debug.WriteLine($"保存规则配置发生异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 扫描指定分类工作表中的特定箱柜元器件数据 (采用 2D 数组一次性批量读入内存)
        /// </summary>
        /// <param name="ws">目标工作表</param>
        /// <param name="cabDetName">箱柜 Det 定义名称 (如 Cab_Det_1)</param>
        /// <returns>CabinetScanData 扫描结果</returns>
        public static CabinetScanData? ScanCabinetData(Worksheet ws, string cabDetName)
        {
            // 校验工作表与名称有效性
            if (ws == null || string.IsNullOrWhiteSpace(cabDetName)) return null;

            try
            {
                // 从工作表中收集并获取所有有效箱柜锚点
                var validCabinets = Tool.GetSheetValidCabinets(ws);
                CabinetAnchorModel? anchor = null;
                int cabIndex = 0;

                // 遍历寻找匹配的目标箱柜
                foreach (var kvp in validCabinets)
                {
                    if (string.Equals($"Cab_Det_{kvp.Key}", cabDetName, StringComparison.OrdinalIgnoreCase))
                    {
                        cabIndex = kvp.Key;
                        anchor = kvp.Value;
                        break;
                    }
                }

                // 校验关键行号的合法性
                if (anchor == null || anchor.Det == null || anchor.Subsum == null || anchor.Tolsum == null) return null;

                int detRow = Convert.ToInt32(anchor.Det.Row);
                int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                int subsumRow = Convert.ToInt32(anchor.Subsum.Row);
                int tolsumRow = Convert.ToInt32(anchor.Tolsum.Row);

                // 计算元器件有效起始行与终止行 (规则: detRow+2 到 subsumRow-1)
                int compStartRow = detRow + 2;
                int compEndRow = subsumRow - 1;

                // 初始化扫描实体
                var scanData = new CabinetScanData
                {
                    SheetName = ws.Name,
                    CabinetIndex = cabIndex,
                    SumRow = sumRow,
                    DetRow = detRow,
                    CompStartRow = compStartRow,
                    CompEndRow = compEndRow,
                    SubsumRow = subsumRow,
                    TolsumRow = tolsumRow,
                    CabinetName = ws.Range[$"A{detRow}"].Value?.ToString() ?? $"箱柜{cabIndex}",
                    Quantity = 1
                };

                // 若元器件区域行数有效，采用 2D 数组一次性批量读入内存 (覆盖 A 到 AB 列)
                if (compEndRow >= compStartRow)
                {
                    // 获取元器件区域的 Range 引用 (覆盖至 AB 列即第 28 列)
                    Range compRange = ws.Range[$"A{compStartRow}:AB{compEndRow}"];
                    // 一次性读取为二维对象数组
                    object[,] compMatrix = compRange.Value2 as object[,];

                    if (compMatrix != null)
                    {
                        // 获取二维数组行数
                        int rowCount = compMatrix.GetLength(0);
                        // 获取二维数组列数
                        int colCount = compMatrix.GetLength(1);
                        // 遍历每一行元器件数据
                        for (int r = 1; r <= rowCount; r++)
                        {
                            // 物理行号
                            int currentRow = compStartRow + r - 1;
                            // 提取 B 列名称 (第 2 列)
                            string name = compMatrix[r, 2]?.ToString()?.Trim() ?? string.Empty;
                            // 提取 C 列型号 (第 3 列)
                            string model = compMatrix[r, 3]?.ToString()?.Trim() ?? string.Empty;

                            // 排除空行或表头无效行
                            if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(model)) continue;
                            if (name == "名称" || name == "配电箱" || name == "箱体") continue;

                            // 提取 F 列数量 (第 6 列)
                            int qty = 1;
                            if (compMatrix[r, 6] != null)
                            {
                                int.TryParse(compMatrix[r, 6].ToString(), out qty);
                                if (qty <= 0) qty = 1;
                            }

                            // 提取 W 列电流 (第 23 列，若为空则从型号中自动正则识别)
                            int current = 0;
                            if (colCount >= 23 && compMatrix[r, 23] != null)
                            {
                                int.TryParse(compMatrix[r, 23].ToString(), out current);
                            }
                            if (current <= 0)
                            {
                                current = ParseCurrentFromModel(model);
                            }

                            // 提取 X 列极数 (第 24 列)
                            string poles = (colCount >= 24 ? compMatrix[r, 24]?.ToString()?.Trim() : null) ?? "3";
                            if (string.IsNullOrWhiteSpace(poles))
                            {
                                poles = ParsePolesFromModel(model);
                            }
                            int poleCount = ParsePoleNumber(poles);

                            // 提取 Y 列脱扣类型/脱扣方式 (第 25 列)
                            string trip = colCount >= 25 ? compMatrix[r, 25]?.ToString()?.Trim() ?? string.Empty : string.Empty;

                            // 提取 Z 列附件描述 (第 26 列)
                            string accessory = colCount >= 26 ? compMatrix[r, 26]?.ToString()?.Trim() ?? string.Empty : string.Empty;

                            // 提取 AA 列图块名称 (第 27 列)
                            string blockName = colCount >= 27 ? compMatrix[r, 27]?.ToString()?.Trim() ?? string.Empty : string.Empty;

                            // 提取 AB 列图块类别 (第 28 列)
                            string blockCategory = colCount >= 28 ? compMatrix[r, 28]?.ToString()?.Trim() ?? string.Empty : string.Empty;

                            // 构造元器件条目实体
                            var compItem = new CabinetComponentItem
                            {
                                RowIndex = currentRow,
                                Name = name,
                                Model = model,
                                Quantity = qty,
                                Current = current,
                                Poles = poles,
                                PoleCount = poleCount,
                                Trip = trip,
                                Accessory = accessory,
                                BlockName = blockName,
                                BlockCategory = blockCategory,
                                IsAts = name.Contains("双电源") || model.Contains("双电源") || model.Contains("ATS") || model.Contains("NZ7") || model.Contains("WATSN"),
                                IsFireTransformer = name.Contains("火灾") || model.Contains("火灾") || name.Contains("漏电互感器"),
                                IsCurrentTransformer = (name.Contains("互感器") || model.Contains("互感器")) && !name.Contains("火灾"),
                                IsReserved = name.Contains("预留") || model.Contains("预留") || name.Contains("备用")
                            };

                            // 加入元器件列表
                            scanData.Components.Add(compItem);
                        }
                    }
                }

                // 返回完整扫描结果
                return scanData;
            }
            catch (Exception ex)
            {
                // 记录扫描异常信息
                System.Diagnostics.Debug.WriteLine($"扫描箱柜发生异常: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 执行箱柜智能推导计算：推导壳体选型、铜排用量、一次/二次辅材与装配人工费
        /// </summary>
        /// <param name="scanData">箱柜扫描数据</param>
        /// <param name="rules">计算规则配置</param>
        /// <returns>CabinetCalcResult 计算结果</returns>
        public static CabinetCalcResult CalculateCabinetAuxAndShell(CabinetScanData scanData, QuotationRules rules)
        {
            // 校验入参
            if (scanData == null) return new CabinetCalcResult();
            if (rules == null) rules = LoadQuotationRules();

            // 提取通用系数
            double xishu = rules.General.ElementMarkupRatio;
            double taxRatio = rules.General.TaxAndManageRatio;
            double copperPrice = rules.General.CopperPricePerKg;

            // 记录最大电流、塑壳开关数、总面积等
            int maxCurrent = 0;
            int plasticCaseCount = 0;
            int totalShuntBreakers = 0;
            int transformerSets = 0;
            bool hasAts = false;
            bool hasFireTransformer = false;
            bool hasReserved = false;
            double totalComponentArea = 0.0;
            int mainSwitchHeight = 0;

            // 电流 -> 线头数汇总字典 (key: 电流, value: 线头数)
            Dictionary<int, int> currentWireMap = new Dictionary<int, int>();
            // 元件名称 -> 数量汇总字典
            Dictionary<string, int> componentNameCountMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // 遍历所有元器件统计基础特征
            for (int i = 0; i < scanData.Components.Count; i++)
            {
                var comp = scanData.Components[i];
                // 更新最大电流
                if (comp.Current > maxCurrent) maxCurrent = comp.Current;

                // 统计特征标记
                if (comp.IsAts) hasAts = true;
                if (comp.IsFireTransformer) hasFireTransformer = true;
                if (comp.IsReserved) hasReserved = true;
                if (comp.IsCurrentTransformer) transformerSets += Math.Max(1, comp.Quantity / 3);

                // 统计塑壳断路器数量
                if (comp.Name.Contains("塑壳") || comp.Model.Contains("塑壳") || comp.Current >= 100)
                {
                    plasticCaseCount += comp.Quantity;
                }
                if (comp.Name.Contains("断路器") || comp.Name.Contains("漏电") || comp.Name.Contains("微断"))
                {
                    totalShuntBreakers += comp.Quantity;
                }

                // 汇总元件名称数量
                if (componentNameCountMap.ContainsKey(comp.Name))
                    componentNameCountMap[comp.Name] += comp.Quantity;
                else
                    componentNameCountMap[comp.Name] = comp.Quantity;

                // 若存在断路器附件，单独累加附件数量
                if (!string.IsNullOrWhiteSpace(comp.Accessory))
                {
                    string accName = "断路器附件";
                    if (componentNameCountMap.ContainsKey(accName))
                        componentNameCountMap[accName] += comp.Quantity;
                    else
                        componentNameCountMap[accName] = comp.Quantity;
                }

                // 提取并校验有效额定电流 (用于接线空间高度估算)
                int effectiveCurrent = comp.Current > 0 ? comp.Current : 25;
                // 保障最小电流基数不低于 25A
                if (effectiveCurrent < 25) effectiveCurrent = 25;

                // 电流线头数统计 (核心约束：首个主元器件 i == 0 不计算分支排也不计算导线，仅出线分路 i > 0 统计导线)
                if (i > 0 && !comp.IsFireTransformer && !comp.IsCurrentTransformer)
                {
                    // 计算出线分路导线线头数 (极数 * 数量)
                    int wireJointCount = comp.PoleCount * comp.Quantity;

                    // 累加对应电流档位的出线线头数
                    if (currentWireMap.ContainsKey(effectiveCurrent))
                        currentWireMap[effectiveCurrent] += wireJointCount;
                    else
                        currentWireMap[effectiveCurrent] = wireJointCount;
                }

                // 计算元件占用面积与接线空间
                int wireSpace = GetWiringSpace(effectiveCurrent, comp.Name, rules.ShellRules.WiringSpaceGradients);
                // 获取元件外形物理尺寸 (宽*高)
                var (compWidth, compHeight) = EstimateComponentDimensions(comp);

                // 首个总开关特殊处理
                if (i == 0 && mainSwitchHeight == 0)
                {
                    mainSwitchHeight = compHeight + wireSpace;
                }
                else if (!comp.IsFireTransformer && !comp.IsCurrentTransformer)
                {
                    // 分支元件面积累加
                    double singleArea = compWidth * (compHeight + wireSpace) * comp.Quantity;
                    totalComponentArea += singleArea;
                }
            }

            // 互感器对总开关高度加成
            if (hasFireTransformer)
            {
                mainSwitchHeight += rules.ShellRules.TransformerSpacing.FireTransformer;
            }
            if (transformerSets == 1)
            {
                mainSwitchHeight += rules.ShellRules.TransformerSpacing.OneSet;
            }
            else if (transformerSets >= 2 && transformerSets <= 5)
            {
                mainSwitchHeight += rules.ShellRules.TransformerSpacing.TwoToFiveSets;
            }
            else if (transformerSets > 5)
            {
                mainSwitchHeight += rules.ShellRules.TransformerSpacing.OverFiveSets;
            }

            // 判定是否为落地柜体
            bool isCabinet = false;
            if (maxCurrent >= rules.ShellRules.CabinetCurrentThreshold) isCabinet = true;
            if (plasticCaseCount >= 6 && maxCurrent >= 160) isCabinet = true;
            if (plasticCaseCount >= 8 && maxCurrent >= 125) isCabinet = true;

            // 智能推导匹配壳体尺寸
            string recommendedSize = MatchOptimalShellSize(
                totalComponentArea,
                isCabinet,
                maxCurrent,
                mainSwitchHeight,
                rules.ShellRules
            );

            // 解析推荐壳体的宽高 (mm)
            int shellWidth = 600;
            int shellHeight = 800;
            if (recommendedSize.Contains("*"))
            {
                var parts = recommendedSize.Split('*');
                if (parts.Length >= 2)
                {
                    int.TryParse(parts[0].Trim(), out shellWidth);
                    int.TryParse(parts[1].Trim(), out shellHeight);
                }
            }
            if (shellHeight > 1000) isCabinet = true;

            // -------------------------------------------------------------
            // 2. 铜排 (TMY) 基于 tmy.DrawIO 全新制作规则与定额计算
            // -------------------------------------------------------------
            double copperWeight = 0.0;
            // 记录铜排各分项计算公式明细列表 (主母排、垂直N排、零地排、出线分支排)
            var copperFormulaDetails = new List<string>();

            // 提取主进线开关 (约定第一行元器件为主进线开关)
            CabinetComponentItem? mainSwitchComp = scanData.Components.Count > 0 ? scanData.Components[0] : null;
            // 主进线额定电流 (若第一行存在则取其电流，否则取整柜最大电流)
            int mainSwitchCurrent = (mainSwitchComp != null && mainSwitchComp.Current > 0) ? mainSwitchComp.Current : maxCurrent;
            // 主进线开关极数 (3P 或 4P，默认 3P)
            int mainSwitchPoleCount = mainSwitchComp != null ? mainSwitchComp.PoleCount : (mainSwitchComp?.Poles?.Contains("4") == true ? 4 : 3);

            // 统计出线/分路元器件特征 (从第 2 行开始，排除主开关)
            int branchMccbCount = 0;              // 出线塑壳断路器数量
            int branchMccbCurrentSum = 0;         // 出线塑壳断路器电流之和
            int branch4PoleMccbCount = 0;         // 出线 4 极塑壳断路器数量
            int branchTotalCurrentSum = 0;        // 出线全部元件电流之和
            // 记录大电流出线分支排按规格分组统计的字典 (Key: 规格名称, Value: (规格条目, 累计台数, 电流档位列表))
            var branchBusGroupMap = new Dictionary<string, (MainBusSpecItem SpecItem, int TotalCount, List<int> Currents)>(StringComparer.OrdinalIgnoreCase);

            // 特殊元器件匹配判定 (满足配置的特殊关键字或原有 ATS/火灾互感器标记)
            bool hasSpecialComponents = false;

            // 遍历箱柜所有元器件检查特殊元器件关键字 (双电源、ATS、火灾探测器等)
            foreach (var comp in scanData.Components)
            {
                // 检查是否命中特殊关键字列表
                if (rules.CopperRules.SpecialComponentKeywords != null)
                {
                    // 遍历配置中的每个特殊元器件关键字
                    foreach (var kw in rules.CopperRules.SpecialComponentKeywords)
                    {
                        if (string.IsNullOrWhiteSpace(kw)) continue;
                        // 匹配元件名称、型号规格、图块类别或图块名称
                        if ((!string.IsNullOrEmpty(comp.Name) && comp.Name.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (!string.IsNullOrEmpty(comp.Model) && comp.Model.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (!string.IsNullOrEmpty(comp.BlockCategory) && comp.BlockCategory.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (!string.IsNullOrEmpty(comp.BlockName) && comp.BlockName.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            hasSpecialComponents = true;
                            break;
                        }
                    }
                }
                // 结合内建布尔标志
                if (comp.IsAts || comp.IsFireTransformer) hasSpecialComponents = true;
            }

            // 遍历出线分路元件 (从索引 1 开始，排除第 0 项主进线开关)
            for (int i = 1; i < scanData.Components.Count; i++)
            {
                var comp = scanData.Components[i];
                // 提取单只额定电流
                int compCur = comp.Current > 0 ? comp.Current : 0;
                // 累加分路总电流 (额定电流 × 数量)
                branchTotalCurrentSum += compCur * Math.Max(1, comp.Quantity);

                // 判断是否为塑壳断路器 (名称或型号含塑壳，或回路电流 >= 100A 且非微断)
                bool isMccb = comp.Name.Contains("塑壳") || comp.Model.Contains("塑壳") ||
                             (!comp.Name.Contains("微断") && compCur >= 100);

                if (isMccb)
                {
                    // 累加出线塑壳台数
                    branchMccbCount += comp.Quantity;
                    // 累加出线塑壳电流总和
                    branchMccbCurrentSum += compCur * comp.Quantity;

                    // 统计极数为 4 的塑壳断路器数量
                    if (comp.PoleCount >= 4 || comp.Poles.Contains("4"))
                    {
                        branch4PoleMccbCount += comp.Quantity;
                    }

                    // 统计大电流出线分支排 (出线电流 > 设定起算门限)
                    if (compCur > rules.CopperRules.BranchMinCurrent)
                    {
                        // 核心重构：出线分支排不能按主排来，根据回路自身额定电流独立匹配母排规则表
                        var branchSpecItem = GetBusbarSpecItem(compCur, rules.CopperRules.MainBusSpecTable);
                        // 检查当前规格是否已存在于分组字典中
                        if (branchBusGroupMap.ContainsKey(branchSpecItem.Spec))
                        {
                            // 提取已有规格分组信息
                            var existing = branchBusGroupMap[branchSpecItem.Spec];
                            // 记录新出现的回路电流档位
                            if (!existing.Currents.Contains(compCur)) existing.Currents.Add(compCur);
                            // 累加该规格断路器数量
                            branchBusGroupMap[branchSpecItem.Spec] = (existing.SpecItem, existing.TotalCount + comp.Quantity, existing.Currents);
                        }
                        else
                        {
                            // 初始化并记录新规格分组
                            branchBusGroupMap[branchSpecItem.Spec] = (branchSpecItem, comp.Quantity, new List<int> { compCur });
                        }
                    }
                }
            }

            // 获取主母排规格条目与理论单重 (kg/m)
            var mainBusSpecItem = GetBusbarSpecItem(mainSwitchCurrent > 0 ? mainSwitchCurrent : maxCurrent, rules.CopperRules.MainBusSpecTable);
            // 提取主母排每米理论单重
            double mainBusWeightPerMeter = mainBusSpecItem.WeightPerMeter;

            // 计算扣除余量后的有效柜宽与有效柜高 (单位: mm)
            int effectiveWidth = Math.Max(0, shellWidth - rules.CopperRules.WidthDeduction);
            // 计算扣除余量后的有效柜高 (单位: mm)
            int effectiveHeight = Math.Max(0, shellHeight - rules.CopperRules.HeightDeduction);

            // ==========================================
            // 分支一：主母排形态判定 (水平排 vs 垂直母排)
            // 依据 tmy.DrawIO 最新流程图严格执行
            // ==========================================
            bool hasHorizontalBus = false;
            // 垂直母排命中标志
            bool hasVerticalBus = false;

            // 1.1 判定是否满足水平排 (节点 5 塑壳数量 >= 门限 且 节点 6 塑壳电流和 >= 门限)
            if (branchMccbCount >= rules.CopperRules.MccbCountThreshold &&
                branchMccbCurrentSum >= rules.CopperRules.MccbCurrentSumThreshold)
            {
                // 满足水平排触发条件 (节点 12)
                hasHorizontalBus = true;
                // 判断 4 极塑壳数量是否达到门限 (节点 42: true 采用 4 根水平排，false 采用 3 根)
                int horizontalPoleCount = (branch4PoleMccbCount >= rules.CopperRules.FourPoleMccbThreshold) ? 4 : 3;
                // 水平排总展开长度 (mm) = (柜宽 - 边距) × 排数
                int horizontalBusLen = effectiveWidth * horizontalPoleCount;
                // 计算水平排理论重量 (kg)
                double horizWeight = (mainBusWeightPerMeter * horizontalBusLen) / 1000.0;
                // 累加铜排总重量
                copperWeight += horizWeight;
                // 记录透明的水平主排计算明细
                copperFormulaDetails.Add($"水平主排 ({horizontalPoleCount}根 | {mainBusSpecItem.Spec} | {mainBusWeightPerMeter:F3}kg/m): {mainBusWeightPerMeter:F3} × [({shellWidth}-{rules.CopperRules.WidthDeduction}) × {horizontalPoleCount}] / 1000 = {horizWeight:F2} KG");
            }
            // 1.2 塑壳台数不足(节点 5 false) 或 塑壳电流和不足(节点 6 false)，均流转至垂直母排判定 (节点 16)
            else if (branchTotalCurrentSum >= rules.CopperRules.BranchTotalCurrentThreshold &&
                     mainSwitchCurrent > rules.CopperRules.MainSwitchCurrentThreshold)
            {
                // 满足垂直母排触发条件 (节点 24)
                hasVerticalBus = true;
                // 判断主开关极数是否为 4 极 (节点 64)
                int vertPoleCount = (mainSwitchPoleCount >= 4) ? 4 : 3;
                // 垂直母排单根展开长 (米) = [基准长 + 延伸系数 × (分路电流和 / 步长基数)]
                double stepBase = rules.CopperRules.LoadExtensionStepCurrent > 0 ? (double)branchTotalCurrentSum / rules.CopperRules.LoadExtensionStepCurrent : 0;
                // 计算垂直母排单根米数
                double vertSingleLenMeters = rules.CopperRules.VerticalBaseLength + (rules.CopperRules.LoadExtensionRatio * stepBase);
                // 垂直母排总长度 (米) = 单根长 × 极数
                double vertTotalLenMeters = vertSingleLenMeters * vertPoleCount;
                // 计算垂直母排理论重量 (kg)
                double vertWeight = mainBusWeightPerMeter * vertTotalLenMeters;
                // 累加铜排总重量
                copperWeight += vertWeight;
                // 记录垂直母排计算明细
                copperFormulaDetails.Add($"垂直母排 ({vertPoleCount}根 | {mainBusSpecItem.Spec} | {mainBusWeightPerMeter:F3}kg/m): {mainBusWeightPerMeter:F3} × [{rules.CopperRules.VerticalBaseLength:F1} + {rules.CopperRules.LoadExtensionRatio:F2}×({branchTotalCurrentSum}/{rules.CopperRules.LoadExtensionStepCurrent})] × {vertPoleCount} = {vertWeight:F2} KG");
            }

            // ==========================================
            // 分支二：垂直 N 排计算 (满足特殊元件或主开关4极)
            // ==========================================
            if (hasSpecialComponents || mainSwitchPoleCount >= 4)
            {
                // 垂直 N 排展开长度 (mm) = 柜高 - 柜高上下边距
                int vertNBusLen = effectiveHeight;
                // 垂直 N 排理论重量 (kg)
                double vertNWeight = (mainBusWeightPerMeter * vertNBusLen) / 1000.0;
                // 累加铜排重量
                copperWeight += vertNWeight;
                // 记录触发原因
                string reason = hasSpecialComponents ? "包含特殊元器件" : "主开关为4极";
                // 记录垂直 N 排计算明细
                copperFormulaDetails.Add($"垂直N排 ({reason} | {mainBusSpecItem.Spec}): {mainBusWeightPerMeter:F3} × ({shellHeight}-{rules.CopperRules.HeightDeduction}) / 1000 = {vertNWeight:F2} KG");
            }

            // ==========================================
            // 分支三：零地排计算 (标配 1 根宽边距排)
            // ==========================================
            if (effectiveWidth > 0)
            {
                // 零地排展开长度 (mm) = 柜宽 - 柜宽边距
                int groundBusLen = effectiveWidth;
                // 零地排理论重量 (kg)
                double groundWeight = (mainBusWeightPerMeter * groundBusLen) / 1000.0;
                // 累加铜排重量
                copperWeight += groundWeight;
                // 记录零地排计算明细
                copperFormulaDetails.Add($"零地排 (标配 | {mainBusSpecItem.Spec}): {mainBusWeightPerMeter:F3} × ({shellWidth}-{rules.CopperRules.WidthDeduction}) / 1000 = {groundWeight:F2} KG");
            }

            // ==========================================
            // 分支四：出线分支铜排计算 (根据各出线回路额定电流独立选型)
            // 规则约束：只有存在水平排时才做出线分支排；若无水平排，出线分支不做排，只能做线
            // ==========================================
            if (hasHorizontalBus && branchBusGroupMap.Count > 0)
            {
                // 获取单台出线分支排基准展开长 (单位: 米，配置默认 1.0m)
                double branchUnitLen = rules.CopperRules.BranchBusUnitLength > 0 ? rules.CopperRules.BranchBusUnitLength : 1.0;

                // 遍历每个规格分组分别计算理论重量并生成透明算式
                foreach (var kvp in branchBusGroupMap)
                {
                    var group = kvp.Value;
                    // 分项理论重量 (kg) = 台数 × 单台基准长 × 对应规格理论每米单重
                    double groupWeight = group.TotalCount * branchUnitLen * group.SpecItem.WeightPerMeter;
                    // 累加铜排总重量
                    copperWeight += groupWeight;
                    // 汇总涉及的回路额定电流描述 (如 "160" 或 "125/160")
                    string currentDesc = string.Join("/", group.Currents);
                    // 记录该规格出线分支排的透明推导算式
                    copperFormulaDetails.Add($"出线分支排 (出线{currentDesc}A共{group.TotalCount}台 | {group.SpecItem.Spec} | {group.SpecItem.WeightPerMeter:F3}kg/m): {group.TotalCount}台 × {branchUnitLen:F1}m × {group.SpecItem.WeightPerMeter:F3}kg/m = {groupWeight:F2} KG");
                }
            }

            // 铜排重量四舍五入保留 1 位小数
            copperWeight = Math.Round(copperWeight, 1);
            // 构造铜排写入单元格的数量公式
            string copperQtyFormula = copperWeight > 0 ? $"=ROUND({copperWeight}*{xishu}*1,1)" : string.Empty;

            // -------------------------------------------------------------
            // 3. 辅材计算 (一次连接导线 + 二次元件接线辅材 + 结构补贴)
            // -------------------------------------------------------------
            double auxiliaryCost = rules.AuxRules.BaseFee;

            // 记录一次导线用量明细字典 (规格名称 -> 用量明细实体)
            var primaryWireMap = new Dictionary<string, PrimaryWireUsageItem>(StringComparer.OrdinalIgnoreCase);

            // 获取一次导线长度计算配置对象
            var wireLenCfg = rules.AuxRules.WireLengthConfig ?? new PrimaryWireLengthConfig();

            // 一次连接导线垂直基准高度计算 (基础基准 + 火灾互感器加成 + 普通互感器加成)
            int verticalLength = wireLenCfg.BaseVerticalHeight;
            if (hasFireTransformer) verticalLength += wireLenCfg.FireTransformerExtraHeight;
            if (transformerSets > 0) verticalLength += wireLenCfg.NormalTransformerExtraHeight;

            // 遍历各电流回路计算一次配线用量
            foreach (var kvp in currentWireMap)
            {
                int cur = kvp.Key;
                int wireCount = kvp.Value;
                // 核心规则联动：若有水平排，电流小于分支门限计算一次导线 (大电流走分支铜排)；
                // 若无水平排，出线分支不做排只能做线，因此所有电流回路全部走一次导线！
                if (!hasHorizontalBus || cur < rules.CopperRules.BranchMinCurrent)
                {
                    // 匹配对应电流的一次线规格与单价
                    var specItem = FindPrimaryWireSpec(cur, rules.AuxRules.PrimaryWireSpecTable);
                    string specName = specItem.Spec;
                    double pricePerMeter = specItem.PricePerMeter;
                    double crossSection = specItem.CrossSection;

                    // 根据箱体尺寸与高度推导单回路导线长度
                    double wireLenMeters;
                    if (shellHeight >= wireLenCfg.CabinetMinHeight)
                    {
                        // 落地柜导线长度公式: 线头数 * (柜宽系数 * 柜宽 + 垂直高度) * 裕量系数 / 1000
                        wireLenMeters = wireCount * (shellWidth * wireLenCfg.CabinetWidthFactor + verticalLength) * wireLenCfg.CabinetLengthMargin / 1000.0;
                    }
                    else
                    {
                        // 配电箱导线长度公式: 线头数 * (箱宽系数 * 柜宽 + 垂直高度) / 1000
                        wireLenMeters = wireCount * (shellWidth * wireLenCfg.BoxWidthFactor + verticalLength) / 1000.0;
                    }

                    // 累加该规格导线的消耗米数
                    if (primaryWireMap.ContainsKey(specName))
                    {
                        primaryWireMap[specName].LengthMeters += wireLenMeters;
                    }
                    else
                    {
                        primaryWireMap[specName] = new PrimaryWireUsageItem
                        {
                            Spec = specName,
                            CrossSection = crossSection,
                            LengthMeters = wireLenMeters,
                            PricePerMeter = pricePerMeter,
                            SubtotalCost = 0
                        };
                    }
                }
            }

            // 汇总一次导线总费用并生成用量明细列表
            var primaryWireDetails = new List<PrimaryWireUsageItem>();
            foreach (var kvp in primaryWireMap)
            {
                var item = kvp.Value;
                // 保留 1 位小数
                item.LengthMeters = Math.Round(item.LengthMeters, 1);
                // 计算该线规金额小计
                item.SubtotalCost = Math.Round(item.LengthMeters * item.PricePerMeter, 1);
                // 累加到辅材总费用
                auxiliaryCost += item.SubtotalCost;
                primaryWireDetails.Add(item);
            }

            // 小箱小零地排补贴 (最大电流 < 140A)
            if (maxCurrent < rules.CopperRules.IStructureCurrent)
            {
                auxiliaryCost += rules.AuxRules.SmallBoxGroundBarFee;
            }
            // 高柜辅材补贴 (高度 > 1500mm)
            if (shellHeight > 1500)
            {
                auxiliaryCost += rules.AuxRules.HighCabinetExtraFee;
            }

            // 二次元件接线辅材
            foreach (var kvp in componentNameCountMap)
            {
                string compName = kvp.Key;
                int count = kvp.Value;
                var secRule = FindSecondaryRule(compName, rules.AuxRules.SecondaryElements);
                if (secRule != null)
                {
                    double secWireCost = (secRule.WireCount * (shellWidth + 300) * secRule.WirePrice / 1000.0) * count;
                    auxiliaryCost += secWireCost;
                }
            }

            auxiliaryCost = Math.Round(auxiliaryCost, 1);
            string auxFormula = $"=ROUND({auxiliaryCost}*{xishu}*1,1)";

            // -------------------------------------------------------------
            // 4. 人工费计算 (箱体平铺面积制作工价 + 二次元件安装接线工价)
            // -------------------------------------------------------------
            double widthDm = shellWidth / 100.0;
            double heightDm = shellHeight / 100.0;
            string areaLaborFormula;

            if (hasReserved && totalShuntBreakers <= 3)
            {
                areaLaborFormula = $"{widthDm:F1}*{heightDm:F1}*{rules.LaborRules.AreaBaseRate:F2}*{rules.LaborRules.ReservedCircuitDiscount:F1}";
            }
            else
            {
                areaLaborFormula = $"{widthDm:F1}*{heightDm:F1}*{rules.LaborRules.AreaBaseRate:F2}";
            }

            // 二次元件装配工价累加
            List<string> laborTerms = new List<string> { areaLaborFormula };
            double totalLaborCost = (widthDm * heightDm * rules.LaborRules.AreaBaseRate) * (hasReserved && totalShuntBreakers <= 3 ? rules.LaborRules.ReservedCircuitDiscount : 1.0);

            foreach (var kvp in componentNameCountMap)
            {
                string compName = kvp.Key;
                int count = kvp.Value;
                var secRule = FindSecondaryRule(compName, rules.AuxRules.SecondaryElements);
                if (secRule != null && secRule.LaborPrice > 0)
                {
                    laborTerms.Add($"{secRule.LaborPrice}*{count}");
                    totalLaborCost += secRule.LaborPrice * count;
                }
            }

            string combinedLaborExpr = string.Join("+", laborTerms);
            string laborFormula = $"=ROUND(({combinedLaborExpr})*{xishu}*{taxRatio},1)";
            totalLaborCost = Math.Round(totalLaborCost * xishu * taxRatio, 1);

            // 构造并返回结果模型
            return new CabinetCalcResult
            {
                CabinetName = scanData.CabinetName,
                DetRow = scanData.DetRow,
                SubsumRow = scanData.SubsumRow,
                TolsumRow = scanData.TolsumRow,
                ComponentArea = Math.Round(totalComponentArea, 0),
                MaxCurrent = maxCurrent,
                IsCabinet = isCabinet,
                RecommendedShellSize = recommendedSize,
                CopperWeight = copperWeight,
                CopperQtyFormula = copperQtyFormula,
                AuxiliaryCost = auxiliaryCost,
                AuxiliaryFormula = auxFormula,
                LaborCost = totalLaborCost,
                LaborFormula = laborFormula,
                PrimaryWireDetails = primaryWireDetails,
                CopperFormulaDetails = copperFormulaDetails,
                Description = $"推导完成: 最大电流 {maxCurrent}A, 判定为{(isCabinet ? "落地柜" : "配电箱")}, 推荐尺寸 {recommendedSize}"
            };
        }

        /// <summary>
        /// 将智能推导结果回写至当前分类工作表 (严格遵循用户指定的壳体回写规则与计费区域规范)
        /// </summary>
        /// <param name="ws">目标工作表</param>
        /// <param name="scanData">箱柜扫描定位数据</param>
        /// <param name="result">推导计算结果</param>
        /// <param name="rules">规则配置</param>
        /// <returns>操作是否成功</returns>
        public static bool WriteCabinetCalcResultToSheet(Worksheet ws, CabinetScanData scanData, CabinetCalcResult result, QuotationRules rules)
        {
            // 校验基础对象
            if (ws == null || scanData == null || result == null) return false;
            if (rules == null) rules = LoadQuotationRules();

            try
            {
                int detRow = scanData.DetRow;
                int subsumRow = scanData.SubsumRow;
                int tolsumRow = scanData.TolsumRow;
                string shellMatchName = rules.ShellRules.ShellMatchName?.Trim() ?? "箱体";

                // ---------------------------------------------------------
                // 1. 壳体写入规则判定与执行
                // 规则: 如果在计费区域的 B 列能找到同名(shellMatchName)，写入该行 C 列；
                // 找不到则将尺寸写入 Cab_Det 的 C 列，B 列名称设为匹配名称
                // ---------------------------------------------------------
                bool matchedShellInFeeArea = false;
                int feeStartRow = subsumRow;
                int feeEndRow = tolsumRow - 1;

                if (feeEndRow >= feeStartRow)
                {
                    // 规则 7: 2D 数组一次性读取计费区域 (A 到 G 列)
                    Range feeRange = ws.Range[$"A{feeStartRow}:G{feeEndRow}"];
                    object[,] feeMatrix = feeRange.Formula as object[,];

                    if (feeMatrix != null)
                    {
                        int feeRowCount = feeMatrix.GetLength(0);
                        for (int r = 1; r <= feeRowCount; r++)
                        {
                            string bName = feeMatrix[r, 2]?.ToString()?.Trim() ?? string.Empty;
                            // 匹配壳体同名项
                            if (string.Equals(bName, shellMatchName, StringComparison.OrdinalIgnoreCase))
                            {
                                feeMatrix[r, 3] = result.RecommendedShellSize;
                                matchedShellInFeeArea = true;
                                result.ShellMatchedInFeeArea = true;
                                result.ShellTargetLocation = $"计费区域第 {feeStartRow + r - 1} 行 (B列: {shellMatchName})";
                                break;
                            }
                        }

                        // 如果在计费区匹配到了壳体，批量写回计费区
                        if (matchedShellInFeeArea)
                        {
                            feeRange.Formula = feeMatrix;
                        }
                    }
                }

                // 若计费区域未匹配到，将尺寸写入 Cab_Det 的 C 列，B 列名称写入匹配名称
                if (!matchedShellInFeeArea)
                {
                    Range detRange = ws.Range[$"A{detRow}:E{detRow}"];
                    object[,] detMatrix = detRange.Formula as object[,];
                    if (detMatrix != null)
                    {
                        detMatrix[1, 2] = shellMatchName;
                        detMatrix[1, 3] = result.RecommendedShellSize;
                        detRange.Formula = detMatrix;
                    }
                    else
                    {
                        ws.Range[$"B{detRow}"].Value2 = shellMatchName;
                        ws.Range[$"C{detRow}"].Value2 = result.RecommendedShellSize;
                    }
                    result.ShellMatchedInFeeArea = false;
                    result.ShellTargetLocation = $"箱柜信息行 Cab_Det (第 {detRow} 行)";
                }

                // ---------------------------------------------------------
                // 2. 铜排、辅材、人工费回写至计费区域
                // ---------------------------------------------------------
                // 重新读取最新的计费区域
                feeStartRow = subsumRow;
                feeEndRow = tolsumRow - 1;

                if (feeEndRow >= feeStartRow)
                {
                    Range feeRange = ws.Range[$"A{feeStartRow}:G{feeEndRow}"];
                    object[,] feeMatrix = feeRange.Formula as object[,];

                    if (feeMatrix != null)
                    {
                        int feeRowCount = feeMatrix.GetLength(0);
                        bool hasCopperRow = false;

                        for (int r = 1; r <= feeRowCount; r++)
                        {
                            string bName = feeMatrix[r, 2]?.ToString()?.Trim() ?? string.Empty;

                            // 辅材行回写
                            if (bName == "辅材" && !string.IsNullOrWhiteSpace(result.AuxiliaryFormula))
                            {
                                feeMatrix[r, 7] = result.AuxiliaryFormula;
                            }
                            // 人工费行回写
                            else if ((bName == "人工费" || bName == "人工") && !string.IsNullOrWhiteSpace(result.LaborFormula))
                            {
                                feeMatrix[r, 7] = result.LaborFormula;
                            }
                            // 已有铜排行回写
                            else if (bName == "铜排" && result.CopperWeight > 0)
                            {
                                feeMatrix[r, 3] = "TMY";
                                feeMatrix[r, 4] = result.CopperQtyFormula;
                                feeMatrix[r, 5] = "KG";
                                feeMatrix[r, 6] = rules.General.CopperPricePerKg;
                                int curPhysicalRow = feeStartRow + r - 1;
                                feeMatrix[r, 7] = $"=F{curPhysicalRow}*D{curPhysicalRow}";
                                hasCopperRow = true;
                            }
                        }

                        // 一次性写回计费区域
                        feeRange.Formula = feeMatrix;

                        // 若计算出铜排重量 > 0 且原计费区中无“铜排”行，需在计费区第一行前安全插入铜排行
                        if (result.CopperWeight > 0 && !hasCopperRow)
                        {
                            int insertRow = feeStartRow;
                            dynamic insertRange = ws.Rows[insertRow];
                            insertRange.Insert(XlInsertShiftDirection.xlShiftDown);

                            // 填充新插入的铜排行
                            ws.Range[$"B{insertRow}"].Value2 = "铜排";
                            ws.Range[$"C{insertRow}"].Value2 = "TMY";
                            ws.Range[$"D{insertRow}"].Formula = result.CopperQtyFormula;
                            ws.Range[$"E{insertRow}"].Value2 = "KG";
                            ws.Range[$"F{insertRow}"].Value2 = rules.General.CopperPricePerKg;
                            ws.Range[$"G{insertRow}"].Formula = $"=F{insertRow}*D{insertRow}";

                            // 重新刷新计费区域公式与小计
                            RefreshCabinetFeeAreaFormulas(ws, detRow, scanData.CompStartRow, subsumRow + 1, tolsumRow + 1);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"回写箱柜计算结果发生异常: {ex.Message}");
                return false;
            }
        }

        // =========================================================
        // 私有辅助推导算法
        // =========================================================

        /// <summary>
        /// 匹配最优标准壳体尺寸
        /// </summary>
        private static string MatchOptimalShellSize(
            double compArea,
            bool isCabinet,
            int maxCurrent,
            int mainSwitchReserveHeight,
            ShellConfig shellConfig)
        {
            if (shellConfig.StandardSizes == null || shellConfig.StandardSizes.Count == 0)
            {
                return isCabinet ? "800*1800" : "600*800";
            }

            foreach (string sizeStr in shellConfig.StandardSizes)
            {
                if (!sizeStr.Contains("*")) continue;
                var parts = sizeStr.Split('*');
                if (parts.Length < 2) continue;

                if (!int.TryParse(parts[0].Trim(), out int w) || !int.TryParse(parts[1].Trim(), out int h)) continue;

                bool isCandidateCabinet = isCabinet || h > 1000;

                if (isCandidateCabinet)
                {
                    // 落地柜可用面积计算
                    double availableArea = (w - 120) * (h - 250 - mainSwitchReserveHeight) - 20000;
                    if (compArea <= 0) return sizeStr;
                    if (availableArea / compArea >= shellConfig.CabinetAreaSafetyFactor && h >= shellConfig.CabinetMinHeight)
                    {
                        return sizeStr;
                    }
                }
                else
                {
                    // 配电箱可用面积计算
                    double availableArea = (w - 100) * (h - 80 - mainSwitchReserveHeight);
                    if (compArea <= 0) return sizeStr;
                    if (availableArea / compArea >= shellConfig.BoxAreaSafetyFactor)
                    {
                        return sizeStr;
                    }
                }
            }

            // 兜底返回列表中的最后一项或标准尺寸
            return shellConfig.StandardSizes[shellConfig.StandardSizes.Count - 1];
        }

        /// <summary>
        /// 根据电流大小获取接线空间预留高度
        /// </summary>
        private static int GetWiringSpace(int current, string compName, List<WiringSpaceItem> gradients)
        {
            if (compName == "多功能表" || compName == "电度表") return 0;
            if (gradients == null || gradients.Count == 0) return 110;

            foreach (var item in gradients)
            {
                if (current <= item.MaxCurrent)
                {
                    int space = item.Space;
                    if (compName.Contains("塑壳") && space < 150) space = 150;
                    return space;
                }
            }
            return 370;
        }

        /// <summary>
        /// 获取电流对应的母排规格条目与每米单重
        /// </summary>
        private static MainBusSpecItem GetBusbarSpecItem(int current, List<MainBusSpecItem> specTable)
        {
            // 默认兜底规格为 TMY-30*4 单重 1.068 kg/m --硬编码--
            if (specTable == null || specTable.Count == 0) return new MainBusSpecItem { Spec = "TMY-30*4", WeightPerMeter = 1.068, MaxCurrent = 250 };

            // 遍历规格表阶梯逐级比对电流上限
            foreach (var item in specTable)
            {
                // 若回路电流小于等于该档电流上限则命中
                if (current <= item.MaxCurrent)
                {
                    return item;
                }
            }
            // 超出上限时返回最大档位规格
            return specTable[specTable.Count - 1];
        }

        /// <summary>
        /// 获取电流对应的母排理论每米单重
        /// </summary>
        private static double GetBusbarWeightPerMeter(int current, List<MainBusSpecItem> specTable)
        {
            // 调用 GetBusbarSpecItem 获取条目并提取单重
            var item = GetBusbarSpecItem(current, specTable);
            // 返回理论每米单重
            return item != null ? item.WeightPerMeter : 1.068;
        }

        /// <summary>
        /// 估算元器件外形尺寸 (宽*高，单位 mm)
        /// </summary>
        private static (int Width, int Height) EstimateComponentDimensions(CabinetComponentItem comp)
        {
            // 若为塑壳断路器
            if (comp.Name.Contains("塑壳") || comp.Current >= 100)
            {
                if (comp.Current <= 125) return (90, 150);
                if (comp.Current <= 250) return (105, 165);
                if (comp.Current <= 400) return (140, 255);
                if (comp.Current <= 630) return (185, 270);
                return (210, 300);
            }
            // 若为双电源 ATS
            if (comp.IsAts)
            {
                if (comp.Current <= 100) return (220, 180);
                if (comp.Current <= 250) return (300, 240);
                return (400, 300);
            }
            // 若为微型断路器
            if (comp.Name.Contains("微断") || comp.Name.Contains("断路器") || comp.Name.Contains("小型"))
            {
                int p = comp.PoleCount > 0 ? comp.PoleCount : 1;
                return (18 * p, 80);
            }
            // 若为接触器 / 继电器
            if (comp.Name.Contains("接触器") || comp.Name.Contains("继电器"))
            {
                return (55, 85);
            }
            // 若为仪表
            if (comp.Name.Contains("表") || comp.Name.Contains("多功能"))
            {
                return (96, 96);
            }
            // 默认基准尺寸
            return (50, 80);
        }

        /// <summary>
        /// 匹配电流对应的一次配线规格与单价条目
        /// </summary>
        /// <param name="current">回路额定电流 (A)</param>
        /// <param name="specTable">一次配线规格选型对照表</param>
        /// <returns>匹配到的 PrimaryWireSpecItem 实体</returns>
        private static PrimaryWireSpecItem FindPrimaryWireSpec(int current, List<PrimaryWireSpecItem> specTable)
        {
            // 校验规格表有效性
            if (specTable == null || specTable.Count == 0)
            {
                // 默认返回 4.0 平方导线作为兜底
                return new PrimaryWireSpecItem { MaxCurrent = 999, Spec = "BV-4.0", CrossSection = 4.0, PricePerMeter = 2.5 };
            }

            // 遍历规格阶梯寻找首个满足电流上限的规格
            foreach (var item in specTable)
            {
                if (current <= item.MaxCurrent)
                {
                    return item;
                }
            }

            // 若超出所有设定电流，返回表中最大规格
            return specTable[specTable.Count - 1];
        }

        /// <summary>
        /// 匹配二次元件定额规则
        /// </summary>
        private static SecondaryElementRule? FindSecondaryRule(string compName, List<SecondaryElementRule> rules)
        {
            if (string.IsNullOrWhiteSpace(compName) || rules == null) return null;
            foreach (var r in rules)
            {
                if (compName.Contains(r.Keyword)) return r;
            }
            return null;
        }

        /// <summary>
        /// 从型号文本中提取电流数字
        /// </summary>
        private static int ParseCurrentFromModel(string model)
        {
            if (string.IsNullOrWhiteSpace(model)) return 25;
            var match = Regex.Match(model, @"(?:C|D|In=|/|\s)(\d{1,4})(?:A|\b)", RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out int cur))
            {
                return cur;
            }
            var numMatch = Regex.Match(model, @"\b(\d{2,4})\b");
            if (numMatch.Success && int.TryParse(numMatch.Groups[1].Value, out int cur2))
            {
                if (cur2 >= 6 && cur2 <= 6300) return cur2;
            }
            return 25;
        }

        /// <summary>
        /// 从型号文本中提取极数
        /// </summary>
        private static string ParsePolesFromModel(string model)
        {
            if (string.IsNullOrWhiteSpace(model)) return "3";
            if (Regex.IsMatch(model, @"4P|/4300|/4\b|\b4极", RegexOptions.IgnoreCase)) return "4";
            if (Regex.IsMatch(model, @"3P\+N|3N", RegexOptions.IgnoreCase)) return "3+N";
            if (Regex.IsMatch(model, @"3P|/3300|/3\b|\b3极", RegexOptions.IgnoreCase)) return "3";
            if (Regex.IsMatch(model, @"2P|1P\+N|1N", RegexOptions.IgnoreCase)) return "2";
            if (Regex.IsMatch(model, @"1P|\b1极", RegexOptions.IgnoreCase)) return "1";
            return "3";
        }

        /// <summary>
        /// 解析极数字符串为数值
        /// </summary>
        private static int ParsePoleNumber(string poles)
        {
            if (string.IsNullOrWhiteSpace(poles)) return 3;
            if (poles.Contains("+N") || poles.Contains("＋N") || poles.Contains("1N")) return 4;
            if (poles.Contains("4")) return 4;
            if (poles.Contains("3")) return 3;
            if (poles.Contains("2")) return 2;
            if (poles.Contains("1")) return 1;
            return 3;
        }
    }
}
