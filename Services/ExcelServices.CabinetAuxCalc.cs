using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;
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

                // 若元器件区域行数有效，采用 2D 数组一次性批量读入内存 (覆盖 A 到 AF 列即第 32 列)
                if (compEndRow >= compStartRow)
                {
                    // 获取元器件区域的 Range 引用 (覆盖至 AF 列第 32 列以提取二次方案绑定图号)
                    Range compRange = ws.Range[$"A{compStartRow}:AF{compEndRow}"];
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

                            // 提取 W 列电流 (第 23 列，若为空则保持 0，不做自动提取，由后续计算引擎统一做告警提醒)
                            int current = 0;
                            if (colCount >= 23 && compMatrix[r, 23] != null)
                            {
                                // 尝试将 W 列读取为整数电流
                                int.TryParse(compMatrix[r, 23].ToString(), out current);
                            }
                            // 规则约束：电流列为空时不调用自动正则提取，保持 current 为 0，避免误将型号代号当成电流值

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

                            // 提取 AF 列绑定的二次回路代号或图号 (第 32 列)
                            string boundDwgCode = colCount >= 32 ? compMatrix[r, 32]?.ToString()?.Trim() ?? string.Empty : string.Empty;

                            // 判别是否为二次元件组：B列='元件组' 或 C列以'*'开头，或第32列存在绑定图号
                            bool isCategoryGroup = string.Equals(name, "元件组", StringComparison.OrdinalIgnoreCase);
                            // 型号以 * 开头
                            bool isModelStar = string.IsNullOrWhiteSpace(name) && model.StartsWith("*");
                            // 综合二次元件组判定
                            bool isComponentGroup = isCategoryGroup || isModelStar || !string.IsNullOrWhiteSpace(boundDwgCode);

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
                                BoundDwgCode = boundDwgCode,
                                IsComponentGroup = isComponentGroup,
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

            // 记录计算过程中的警告与未填电流提醒列表
            var calcWarnings = new List<string>();

            // 遍历所有元器件统计基础特征
            for (int i = 0; i < scanData.Components.Count; i++)
            {
                var comp = scanData.Components[i];
                // 判断当前元器件是否属于一次导线与铜排计算自定义集合
                bool inPrimarySet = IsComponentInPrimaryCalcSet(comp, rules.General?.PrimaryCalcComponents);

                // 若在集合内但电流列为空 (<= 0)，按用户需求记录警告提醒，不做自动提取
                if (inPrimarySet && comp.Current <= 0)
                {
                    // 记录提醒文本 (包含行号、名称和型号)
                    calcWarnings.Add($"第 {comp.RowIndex} 行【{comp.Name} {comp.Model}】W列电流为空，不做自动提取，已跳过一次线与分支铜排计算！");
                }

                // 更新最大电流 (必须属于一次计算集合且具有有效正电流)
                if (inPrimarySet && comp.Current > maxCurrent) maxCurrent = comp.Current;

                // 统计特征标记
                if (comp.IsAts) hasAts = true;
                if (comp.IsFireTransformer) hasFireTransformer = true;
                if (comp.IsReserved) hasReserved = true;
                if (comp.IsCurrentTransformer) transformerSets += Math.Max(1, comp.Quantity / 3);

                // 统计塑壳断路器数量 (必须属于一次计算集合)
                if (inPrimarySet && (comp.Name.Contains("塑壳") || comp.Model.Contains("塑壳") || (comp.Current >= 100 && !comp.Name.Contains("微断"))))
                {
                    plasticCaseCount += comp.Quantity;
                }
                // 统计整柜有效分路断路器总台数 (包含断路器、微断、微型漏电、漏电开关等)
                string compName = comp.Name ?? string.Empty;
                string compModel = comp.Model ?? string.Empty;
                if (compName.Contains("断路器") || compName.Contains("漏电") || compName.Contains("微断") ||
                    compModel.Contains("断路器") || compModel.Contains("漏电") || compModel.Contains("微断"))
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

                // 电流线头数统计 (核心约束：首个主元器件 i == 0 不统计导线，仅出线分路 i > 0 且必须在一次计算集合内、电流有效才统计导线)
                if (i > 0 && inPrimarySet && comp.Current > 0 && !comp.IsFireTransformer && !comp.IsCurrentTransformer)
                {
                    // 计算出线分路导线线头数 (极数 * 数量)
                    int wireJointCount = comp.PoleCount * comp.Quantity;

                    // 累加对应电流档位的出线线头数 (以真实填写的电流为准)
                    if (currentWireMap.ContainsKey(comp.Current))
                        currentWireMap[comp.Current] += wireJointCount;
                    else
                        currentWireMap[comp.Current] = wireJointCount;
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
                // 门禁过滤：只有在自定义集合内的元器件，才计算分支排与出线电流
                if (!IsComponentInPrimaryCalcSet(comp, rules.General?.PrimaryCalcComponents))
                {
                    // 非一次计算集合内的元件 (如热继电器、接触器、电涌保护器等) 直接跳过
                    continue;
                }

                // 提取单只额定电流 (电流列为空时跳过分支铜排计算)
                int compCur = comp.Current > 0 ? comp.Current : 0;
                if (compCur <= 0) continue;

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

            // 提取小箱零地排电流门限 (若整柜最大电流 < 该门限，判定为小箱，免计任何铜排，自动享受零地排补贴)
            int smallBoxCurrentThreshold = rules.CopperRules != null && rules.CopperRules.IStructureCurrent > 0
                ? rules.CopperRules.IStructureCurrent
                : 140; // --硬编码-- 兜底默认 140A

            // 判定是否为小箱免铜排
            bool isSmallBox = maxCurrent < smallBoxCurrentThreshold;

            // 水平母排命中标志 (作用域提升至小箱判定前供后续导线逻辑判断)
            bool hasHorizontalBus = false;
            // 垂直母排命中标志
            bool hasVerticalBus = false;

            if (isSmallBox)
            {
                // 小箱免计任何铜排 (零大母排、零垂直排、零零地排、零分支排)
                copperWeight = 0.0;
                // 记录清晰透明的小箱免排原因
                copperFormulaDetails.Add($"小箱免铜排 (整柜最大电流 {maxCurrent}A < 门限 {smallBoxCurrentThreshold}A，自动计入小箱零地排固定补贴，不计任何大铜排)");
            }
            else
            {
                // ==========================================
                // 分支一：主母排形态判定 (水平排 vs 垂直母排)
                // 依据 tmy.DrawIO 最新流程图严格执行
                // ==========================================

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

            // 一次连接导线垂直基准高度计算 (基础垂直预留 + 元器件映射表规则加成，单柜每类去重只加一次)
            int verticalLength = wireLenCfg.BaseVerticalHeight;
            // 校验配置中的映射规则列表有效性
            if (wireLenCfg.ExtraHeightRules != null && wireLenCfg.ExtraHeightRules.Count > 0)
            {
                // 遍历每条垂直高度加成规则条目
                foreach (var rule in wireLenCfg.ExtraHeightRules)
                {
                    // 忽略空关键字或无效高度增量
                    if (string.IsNullOrWhiteSpace(rule.Keyword) || rule.ExtraHeight <= 0) continue;
                    // 检查当前箱柜中是否包含匹配该关键字的元器件 (单柜同类元件去重，无论多少只/套仅累加一次)
                    bool isHit = scanData.Components.Any(c =>
                        (!string.IsNullOrEmpty(c.Name) && c.Name.IndexOf(rule.Keyword, StringComparison.OrdinalIgnoreCase) >= 0) ||
                        (!string.IsNullOrEmpty(c.Model) && c.Model.IndexOf(rule.Keyword, StringComparison.OrdinalIgnoreCase) >= 0));
                    // 若命中规则项，单柜叠加一次加成高度
                    if (isHit)
                    {
                        verticalLength += rule.ExtraHeight;
                    }
                }
            }
            else
            {
                // 兼容历史老版本配置兜底
                if (hasFireTransformer) verticalLength += wireLenCfg.FireTransformerExtraHeight;
                if (transformerSets > 0) verticalLength += wireLenCfg.NormalTransformerExtraHeight;
            }

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

                    // 提取柜高比例参数 (落地柜与配电箱柜高折算系数)
                    double cabHeightFactor = wireLenCfg.CabinetHeightFactor > 0 ? wireLenCfg.CabinetHeightFactor : 0.4;
                    // 配电箱柜高比例系数
                    double boxHeightFactor = wireLenCfg.BoxHeightFactor > 0 ? wireLenCfg.BoxHeightFactor : 0.3;

                    // 根据箱体尺寸与高度推导单回路导线长度 (综合柜宽与柜高双维空间走线)
                    double wireLenMeters;
                    if (shellHeight >= wireLenCfg.CabinetMinHeight)
                    {
                        // 落地柜导线长度公式: 线头数 * (柜宽系数 * 柜宽 + 柜高系数 * 柜高 + 垂直预留) * 裕量放大系数 / 1000
                        wireLenMeters = wireCount * (shellWidth * wireLenCfg.CabinetWidthFactor + shellHeight * cabHeightFactor + verticalLength) * wireLenCfg.CabinetLengthMargin / 1000.0;
                    }
                    else
                    {
                        // 配电箱导线长度公式: 线头数 * (箱宽系数 * 柜宽 + 箱高系数 * 柜高 + 垂直预留) * 配电箱裕量放大系数 / 1000
                        double boxMargin = wireLenCfg.BoxLengthMargin > 0 ? wireLenCfg.BoxLengthMargin : 1.05;
                        wireLenMeters = wireCount * (shellWidth * wireLenCfg.BoxWidthFactor + shellHeight * boxHeightFactor + verticalLength) * boxMargin / 1000.0;
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

            // -------------------------------------------------------------
            // 固定长度元器件接线导线计算 (短跳线免计算箱体宽高，如接触器 300mm)
            // -------------------------------------------------------------
            if (rules.AuxRules.FixedLengthRules != null && rules.AuxRules.FixedLengthRules.Count > 0)
            {
                // 遍历箱柜所有出线分路元件 (排除主进线总开关)
                for (int i = 1; i < scanData.Components.Count; i++)
                {
                    var comp = scanData.Components[i];
                    // 检查是否命中固定长度映射规则列表
                    var hitRule = rules.AuxRules.FixedLengthRules.FirstOrDefault(r =>
                        !string.IsNullOrWhiteSpace(r.Keyword) && (
                            (!string.IsNullOrEmpty(comp.Name) && comp.Name.IndexOf(r.Keyword, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (!string.IsNullOrEmpty(comp.Model) && comp.Model.IndexOf(r.Keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                        ));

                    // 若命中规则条目且长度有效
                    if (hitRule != null && hitRule.FixedLengthMm > 0)
                    {
                        // 提取该元件的有效额定电流 (用于匹配导线规格截面)
                        int compCurrent = comp.Current > 0 ? comp.Current : 25;
                        // 匹配对应电流档位的一次导线规格
                        var specItem = FindPrimaryWireSpec(compCurrent, rules.AuxRules.PrimaryWireSpecTable);
                        // 获取元件极数 (默认 3 极)
                        int poleCount = comp.PoleCount > 0 ? comp.PoleCount : 3;
                        // 获取元件数量
                        int qty = Math.Max(1, comp.Quantity);
                        // 计算短跳线总长度 (米) = 数量 × 极数 × (固定毫米数 / 1000)
                        double fixedWireMeters = qty * poleCount * (hitRule.FixedLengthMm / 1000.0);

                        // 累加到对应线径规格的一次导线用量明细中
                        if (primaryWireMap.ContainsKey(specItem.Spec))
                        {
                            primaryWireMap[specItem.Spec].LengthMeters += fixedWireMeters;
                        }
                        else
                        {
                            primaryWireMap[specItem.Spec] = new PrimaryWireUsageItem
                            {
                                Spec = specItem.Spec,
                                CrossSection = specItem.CrossSection,
                                LengthMeters = fixedWireMeters,
                                PricePerMeter = specItem.PricePerMeter,
                                SubtotalCost = 0
                            };
                        }
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

            // 小箱小零地排补贴 (最大电流 < 设定门限时补贴，且不计大铜排)
            if (maxCurrent < smallBoxCurrentThreshold)
            {
                auxiliaryCost += rules.AuxRules.SmallBoxGroundBarFee;
            }
            // 高柜辅材补贴 (高度 > 1500mm)
            if (shellHeight > 1500)
            {
                auxiliaryCost += rules.AuxRules.HighCabinetExtraFee;
            }

            // -------------------------------------------------------------
            // 二次方案参数对接：从本地数据库查询二次方案并计算二次配线与装配工价
            // -------------------------------------------------------------
            List<SecondarySchemeEntity> allSchemes;
            try
            {
                // 一次性加载本地全量二次方案以供快速匹配
                allSchemes = PersonalComponentDbService.GetAllSecondarySchemes() ?? new List<SecondarySchemeEntity>();
            }
            catch (Exception exScheme)
            {
                // 异常安全兜底防崩溃
                LogHelper.WriteLog($"[CabinetAuxCalc] 获取二次方案列表异常: {exScheme.Message}");
                allSchemes = new List<SecondarySchemeEntity>();
            }

            // 记录命中的二次方案计算明细列表
            var secondarySchemeDetails = new List<SecondarySchemeCalcItem>();
            // 记录未绑定的二次元件组提示
            var unboundGroupWarnings = new List<string>();

            // 二次配线每米单价 (单位: 元/米，配置优先，默认 0.8) --硬编码--
            double secWirePrice = rules.AuxRules.SecondaryWirePrice > 0 ? rules.AuxRules.SecondaryWirePrice : 0.8;
            // 二次跨门柜宽折算系数 (默认 0.8) --硬编码--
            double secWidthFactor = rules.AuxRules.SecondaryWidthFactor > 0 ? rules.AuxRules.SecondaryWidthFactor : 0.8;
            // 二次跨门柜高比例折算系数 (默认 0.3) --硬编码--
            double secHeightFactor = rules.AuxRules.SecondaryHeightFactor > 0 ? rules.AuxRules.SecondaryHeightFactor : 0.3;
            // 二次端头走向预留裕量 (单位: mm，默认 300) --硬编码--
            int secMarginLen = rules.AuxRules.SecondaryMarginLength > 0 ? rules.AuxRules.SecondaryMarginLength : 300;

            // 单根跨门二次线基准推导长度 (单位: 米，综合柜宽与柜高比例以及端头走向余量)
            double secSingleWireLen = Math.Round((shellWidth * secWidthFactor + shellHeight * secHeightFactor + secMarginLen) / 1000.0, 2);
            if (secSingleWireLen <= 0) secSingleWireLen = 1.0;

            // 遍历箱柜元器件行，优先识别第 32 列绑定方案或二次元件组
            foreach (var comp in scanData.Components)
            {
                // 若既非元件组且未绑定回路代号，则跳过二次方案匹配
                if (!comp.IsComponentGroup && string.IsNullOrWhiteSpace(comp.BoundDwgCode))
                {
                    continue;
                }

                // 执行二次方案匹配：优先回路代号、次选CAD图名、再选方案名/型号
                SecondarySchemeEntity? matchedScheme = null;
                if (!string.IsNullOrWhiteSpace(comp.BoundDwgCode) && allSchemes.Count > 0)
                {
                    string bound = comp.BoundDwgCode.Trim();
                    // 1. 匹配适用回路代号
                    matchedScheme = allSchemes.FirstOrDefault(s => s.ApplicableCodes != null &&
                        s.ApplicableCodes.Any(c => string.Equals(c, bound, StringComparison.OrdinalIgnoreCase)));

                    // 2. 匹配 CAD 图纸名称 (去除后缀对比)
                    if (matchedScheme == null)
                    {
                        string pureBound = Path.GetFileNameWithoutExtension(bound);
                        matchedScheme = allSchemes.FirstOrDefault(s =>
                            string.Equals(Path.GetFileNameWithoutExtension(s.CadDrawingName), pureBound, StringComparison.OrdinalIgnoreCase));
                    }

                    // 3. 匹配方案名称
                    if (matchedScheme == null)
                    {
                        matchedScheme = allSchemes.FirstOrDefault(s =>
                            string.Equals(s.SchemeName, bound, StringComparison.OrdinalIgnoreCase) || s.SchemeName.Contains(bound));
                    }
                }

                // 4. 若第 32 列为空，但标记为元件组，尝试用型号去*号模糊匹配
                if (matchedScheme == null && comp.IsComponentGroup && !string.IsNullOrWhiteSpace(comp.Model) && allSchemes.Count > 0)
                {
                    string cleanModel = comp.Model.TrimStart('*').Trim();
                    if (!string.IsNullOrWhiteSpace(cleanModel))
                    {
                        matchedScheme = allSchemes.FirstOrDefault(s =>
                            (s.ApplicableCodes != null && s.ApplicableCodes.Any(c => c.IndexOf(cleanModel, StringComparison.OrdinalIgnoreCase) >= 0)) ||
                            s.SchemeName.IndexOf(cleanModel, StringComparison.OrdinalIgnoreCase) >= 0);
                    }
                }

                // 若命中二次方案，计算该方案对应的二次配线与人工费
                if (matchedScheme != null)
                {
                    int qty = Math.Max(1, comp.Quantity);

                    // 单套回路二次导线推导长度 (米) = 方案跨门根数 × 单根米数
                    double singleCircuitWireLen = Math.Round(matchedScheme.CrossDoorCount * secSingleWireLen, 1);
                    // 该回路二次导线累计推导总长度 (米) = 单套米数 × 套数
                    double totalCircuitWireLen = Math.Round(singleCircuitWireLen * qty, 1);

                    // 计算方案级二次跨门线辅材费 (元) = 累计总米数 × 线单价
                    double circuitWireCost = Math.Round(totalCircuitWireLen * secWirePrice, 1);
                    auxiliaryCost += circuitWireCost;

                    // 计算方案级二次装配与接线人工工费 (元) = 单套工价 × 套数
                    double circuitLaborCost = Math.Round(matchedScheme.LaborCost * qty, 1);

                    // 记录计算明细项供前端展示与核对
                    secondarySchemeDetails.Add(new SecondarySchemeCalcItem
                    {
                        SchemeId = matchedScheme.Id,
                        SchemeName = matchedScheme.SchemeName,
                        CircuitCode = !string.IsNullOrWhiteSpace(comp.BoundDwgCode) ? comp.BoundDwgCode : comp.Model,
                        Quantity = qty,
                        CrossDoorCount = matchedScheme.CrossDoorCount,
                        SingleWireLength = singleCircuitWireLen,
                        TotalWireLength = totalCircuitWireLen,
                        WireCost = circuitWireCost,
                        UnitLaborCost = matchedScheme.LaborCost,
                        LaborCost = circuitLaborCost
                    });
                }
                else
                {
                    // 记录未绑定方案的提示警告
                    string codeDesc = !string.IsNullOrWhiteSpace(comp.BoundDwgCode) ? comp.BoundDwgCode : comp.Model;
                    if (!string.IsNullOrWhiteSpace(codeDesc) && !unboundGroupWarnings.Contains(codeDesc))
                    {
                        unboundGroupWarnings.Add(codeDesc);
                    }
                }
            }

            auxiliaryCost = Math.Round(auxiliaryCost, 1);
            string auxFormula = $"=ROUND({auxiliaryCost}*{xishu}*1,1)";

            // -------------------------------------------------------------
            // 4. 人工费计算 (箱体平铺面积制作工价 + 绑定的二次方案装配工费)
            // -------------------------------------------------------------
            double widthDm = shellWidth / 100.0;
            double heightDm = shellHeight / 100.0;
            string areaLaborFormula;

            // 预留回路打折判定的断路器台数上限门限 (从配置中动态获取，默认 3 台)
            int maxBreakersThreshold = rules.LaborRules != null && rules.LaborRules.ReservedMaxBreakersThreshold > 0
                ? rules.LaborRules.ReservedMaxBreakersThreshold
                : 3; // --硬编码-- 兜底默认 3 台
            // 判定是否同时满足含预留回路且整柜断路器数量未超门限
            bool isReservedDiscounted = hasReserved && totalShuntBreakers <= maxBreakersThreshold;

            if (isReservedDiscounted)
            {
                areaLaborFormula = $"{widthDm:F1}*{heightDm:F1}*{rules.LaborRules.AreaBaseRate:F2}*{rules.LaborRules.ReservedCircuitDiscount:F1}";
            }
            else
            {
                areaLaborFormula = $"{widthDm:F1}*{heightDm:F1}*{rules.LaborRules.AreaBaseRate:F2}";
            }

            // 基础壳体制作工价
            double totalLaborCost = (widthDm * heightDm * rules.LaborRules.AreaBaseRate) * (isReservedDiscounted ? rules.LaborRules.ReservedCircuitDiscount : 1.0);
            List<string> laborTerms = new List<string> { areaLaborFormula };

            // 累加命中的二次方案装配工价
            foreach (var secItem in secondarySchemeDetails)
            {
                if (secItem.UnitLaborCost > 0)
                {
                    // 追加至算式表达式 (若套数大于1则展示 乘套数)
                    if (secItem.Quantity > 1)
                    {
                        laborTerms.Add($"{secItem.UnitLaborCost:F1}*{secItem.Quantity}");
                    }
                    else
                    {
                        laborTerms.Add($"{secItem.UnitLaborCost:F1}");
                    }
                    // 累加人工总额
                    totalLaborCost += secItem.LaborCost;
                }
            }

            string combinedLaborExpr = string.Join("+", laborTerms);
            string laborFormula = $"=ROUND(({combinedLaborExpr})*{xishu}*{taxRatio},1)";
            totalLaborCost = Math.Round(totalLaborCost * xishu * taxRatio, 1);

            // 汇总二次导线全局统计指标
            double totalCrossDoor = secondarySchemeDetails.Sum(s => s.CrossDoorCount * s.Quantity);
            double totalSecWireLength = Math.Round(secondarySchemeDetails.Sum(s => s.TotalWireLength), 1);

            // 组合推导说明描述
            string desc = $"推导完成: 最大电流 {maxCurrent}A, 判定为{(isCabinet ? "落地柜" : "配电箱")}, 推荐尺寸 {recommendedSize}";
            if (secondarySchemeDetails.Count > 0)
            {
                double totalSecLabor = secondarySchemeDetails.Sum(s => s.LaborCost);
                desc += $" | 已命中 {secondarySchemeDetails.Count} 项二次方案(跨门线共{totalCrossDoor:F0}根, 推导总长{totalSecWireLength:F1}米, 二次工价{totalSecLabor:F1}元)";
            }
            if (unboundGroupWarnings.Count > 0)
            {
                desc += $" | 提示: [{string.Join(", ", unboundGroupWarnings)}] 未绑定二次方案";
            }
            if (calcWarnings.Count > 0)
            {
                desc += $" | ⚠️ 未填电流提醒: 有 {calcWarnings.Count} 项一次元器件W列电流为空，未计入计算";
            }

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
                SecondarySchemeDetails = secondarySchemeDetails,
                SecondarySingleWireLength = secSingleWireLen,
                SecondaryTotalWireCount = totalCrossDoor,
                SecondaryTotalWireLength = totalSecWireLength,
                CopperFormulaDetails = copperFormulaDetails,
                Description = desc,
                Warnings = calcWarnings
            };
        }

        /// <summary>
        /// 判定元器件是否属于参与一次导线与铜排计算的自定义集合
        /// </summary>
        /// <param name="comp">元器件条目</param>
        /// <param name="allowedKeywords">允许的关键字列表</param>
        /// <returns>是否命中集合</returns>
        public static bool IsComponentInPrimaryCalcSet(CabinetComponentItem comp, List<string>? allowedKeywords)
        {
            // 校验元件有效性
            if (comp == null) return false;
            // 若未配置或列表为空，使用默认一次元件关键字列表兜底 --硬编码--
            if (allowedKeywords == null || allowedKeywords.Count == 0)
            {
                // 默认包含主流断路器与开关关键字
                allowedKeywords = new List<string> { "塑壳断路器", "塑壳", "微断", "小型断路器", "断路器", "隔离开关", "负荷开关", "框架断路器", "双电源", "ATS" };
            }
            // 遍历配置中的每个有效关键字
            foreach (var kw in allowedKeywords)
            {
                // 忽略空白关键字
                if (string.IsNullOrWhiteSpace(kw)) continue;
                // 逐项匹配元件名称、型号规格、图块类别或图块名称
                if ((!string.IsNullOrEmpty(comp.Name) && comp.Name.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    (!string.IsNullOrEmpty(comp.Model) && comp.Model.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    (!string.IsNullOrEmpty(comp.BlockCategory) && comp.BlockCategory.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    (!string.IsNullOrEmpty(comp.BlockName) && comp.BlockName.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    // 命中一次计算集合
                    return true;
                }
            }
            // 未匹配任何一次元件关键字
            return false;
        }

        /// <summary>
        /// 将智能推导结果回写至当前分类工作表 (壳体、辅材、人工双区查找，铜排填入元器件最末行)
        /// </summary>
        /// <param name="ws">目标工作表</param>
        /// <param name="scanData">箱柜扫描定位数据</param>
        /// <param name="result">推导计算结果</param>
        /// <param name="rules">规则配置</param>
        /// <returns>操作是否成功</returns>
        public static bool WriteCabinetCalcResultToSheet(Worksheet ws, CabinetScanData scanData, CabinetCalcResult result, QuotationRules rules)
        {
            // 校验基础对象有效性
            if (ws == null || scanData == null || result == null) return false;
            // 若规则为空，加载默认规则配置
            if (rules == null) rules = LoadQuotationRules();

            try
            {
                // 获取箱柜基本物理行号
                int detRow = scanData.DetRow;
                int subsumRow = scanData.SubsumRow;
                int tolsumRow = scanData.TolsumRow;
                int compStartRow = scanData.CompStartRow;

                // 提取用户配置的各项回写匹配名称 (默认兜底) --硬编码--
                string shellMatchName = rules.ShellRules.ShellMatchName?.Trim() ?? "箱体";
                string auxMatchName = rules.AuxRules.AuxMatchName?.Trim() ?? "辅材";
                string laborMatchName = rules.LaborRules.LaborMatchName?.Trim() ?? "人工费";

                // ---------------------------------------------------------
                // 1. 优先扫描并回写计费区域 (涵盖壳体、辅材、人工匹配)
                // 计费区域定义: subsumRow (小计行) 至 tolsumRow - 1 (单台合计前一行)
                // ---------------------------------------------------------
                bool matchedShellInFeeArea = false;
                bool matchedAuxInFeeArea = false;
                bool matchedLaborInFeeArea = false;

                int feeStartRow = subsumRow;
                int feeEndRow = tolsumRow - 1;

                if (feeEndRow >= feeStartRow)
                {
                    // 规则 7: 采用 2D 数组一次性批量读取计费区域 (A 到 H 列覆盖至销售总价列)
                    Range feeRange = ws.Range[$"A{feeStartRow}:H{feeEndRow}"];
                    object[,] feeMatrix = feeRange.Formula as object[,];

                    if (feeMatrix != null)
                    {
                        int feeRowCount = feeMatrix.GetLength(0);
                        bool feeMatrixModified = false;

                        // 遍历计费区域的每一行检查匹配项
                        for (int r = 1; r <= feeRowCount; r++)
                        {
                            // 提取 B 列 (第 2 列) 名称
                            string bName = feeMatrix[r, 2]?.ToString()?.Trim() ?? string.Empty;
                            int currentPhysRow = feeStartRow + r - 1;

                            // 1.1 壳体匹配: 在计费区 B 列匹配同名，命中则写入该行 C 列
                            if (!matchedShellInFeeArea && string.Equals(bName, shellMatchName, StringComparison.OrdinalIgnoreCase))
                            {
                                // 将推荐尺寸写入 C 列 (规格型号)
                                feeMatrix[r, 3] = result.RecommendedShellSize;
                                matchedShellInFeeArea = true;
                                feeMatrixModified = true;
                                result.ShellMatchedInFeeArea = true;
                                result.ShellTargetLocation = $"计费区域第 {currentPhysRow} 行 (B列: {shellMatchName})";
                            }

                            // 1.2 辅材匹配: 优先在计费区查找匹配名称
                            if (!matchedAuxInFeeArea && (string.Equals(bName, auxMatchName, StringComparison.OrdinalIgnoreCase) ||
                                (auxMatchName == "辅材" && bName == "辅材")))
                            {
                                if (!string.IsNullOrWhiteSpace(result.AuxiliaryFormula))
                                {
                                    // 销售总价写入 H 列 (第 8 列)
                                    feeMatrix[r, 8] = result.AuxiliaryFormula;
                                    matchedAuxInFeeArea = true;
                                    feeMatrixModified = true;
                                    result.AuxTargetLocation = $"计费区域第 {currentPhysRow} 行 (B列: {bName})";
                                }
                            }

                            // 1.3 人工匹配: 优先在计费区查找匹配名称 (支持"人工费"或"人工")
                            if (!matchedLaborInFeeArea && (string.Equals(bName, laborMatchName, StringComparison.OrdinalIgnoreCase) ||
                                (laborMatchName == "人工费" && (bName == "人工费" || bName == "人工"))))
                            {
                                if (!string.IsNullOrWhiteSpace(result.LaborFormula))
                                {
                                    // 销售总价写入 H 列 (第 8 列)
                                    feeMatrix[r, 8] = result.LaborFormula;
                                    matchedLaborInFeeArea = true;
                                    feeMatrixModified = true;
                                    result.LaborTargetLocation = $"计费区域第 {currentPhysRow} 行 (B列: {bName})";
                                }
                            }

                            // 1.4 清理历史遗留在计费区的铜排行，避免与元器件区域铜排重复计费
                            if (bName == "铜排")
                            {
                                // 清空历史铜排行的数量与单价合价
                                feeMatrix[r, 4] = string.Empty;
                                feeMatrix[r, 6] = string.Empty;
                                feeMatrix[r, 7] = string.Empty;
                                feeMatrix[r, 8] = string.Empty;
                                feeMatrixModified = true;
                            }
                        }

                        // 若计费区域有更新，一次性批量写回
                        if (feeMatrixModified)
                        {
                            feeRange.Formula = feeMatrix;
                        }
                    }
                }

                // ---------------------------------------------------------
                // 2. 壳体兜底写入: 若计费区未匹配到，回退写入 Cab_Det 信息行
                // ---------------------------------------------------------
                if (!matchedShellInFeeArea)
                {
                    Range detRange = ws.Range[$"A{detRow}:E{detRow}"];
                    object[,] detMatrix = detRange.Formula as object[,];
                    if (detMatrix != null)
                    {
                        // B 列写入壳体匹配名称
                        detMatrix[1, 2] = shellMatchName;
                        // C 列写入推荐壳体尺寸
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
                // 3. 辅材与人工二级查找: 若计费区未找到，到元器件区域查找
                // 元器件区域定义: compStartRow (detRow + 2) 至 subsumRow - 1
                // ---------------------------------------------------------
                int compEndRow = subsumRow - 1;
                if ((!matchedAuxInFeeArea || !matchedLaborInFeeArea) && compEndRow >= compStartRow)
                {
                    // 规则 7: 2D 数组读取元器件区域 A 到 H 列
                    Range compRange = ws.Range[$"A{compStartRow}:H{compEndRow}"];
                    object[,] compMatrix = compRange.Formula as object[,];

                    if (compMatrix != null)
                    {
                        int compRowCount = compMatrix.GetLength(0);
                        bool compMatrixModified = false;

                        for (int r = 1; r <= compRowCount; r++)
                        {
                            string bName = compMatrix[r, 2]?.ToString()?.Trim() ?? string.Empty;
                            int currentPhysRow = compStartRow + r - 1;

                            // 辅材在元器件区域匹配
                            if (!matchedAuxInFeeArea && (string.Equals(bName, auxMatchName, StringComparison.OrdinalIgnoreCase) ||
                                (auxMatchName == "辅材" && bName == "辅材")))
                            {
                                if (!string.IsNullOrWhiteSpace(result.AuxiliaryFormula))
                                {
                                    // F 列数量置为 1
                                    compMatrix[r, 6] = 1;
                                    // H 列写入销售总价公式
                                    compMatrix[r, 8] = result.AuxiliaryFormula;
                                    matchedAuxInFeeArea = true;
                                    compMatrixModified = true;
                                    result.AuxTargetLocation = $"元器件区域第 {currentPhysRow} 行 (B列: {bName})";
                                }
                            }

                            // 人工在元器件区域匹配
                            if (!matchedLaborInFeeArea && (string.Equals(bName, laborMatchName, StringComparison.OrdinalIgnoreCase) ||
                                (laborMatchName == "人工费" && (bName == "人工费" || bName == "人工"))))
                            {
                                if (!string.IsNullOrWhiteSpace(result.LaborFormula))
                                {
                                    // F 列数量置为 1
                                    compMatrix[r, 6] = 1;
                                    // H 列写入装配工费公式
                                    compMatrix[r, 8] = result.LaborFormula;
                                    matchedLaborInFeeArea = true;
                                    compMatrixModified = true;
                                    result.LaborTargetLocation = $"元器件区域第 {currentPhysRow} 行 (B列: {bName})";
                                }
                            }
                        }

                        // 若元器件区域有命中修改，批量写回
                        if (compMatrixModified)
                        {
                            compRange.Formula = compMatrix;
                        }
                    }
                }

                // ---------------------------------------------------------
                // 4. 铜排写入元器件最下面一行 (若无空位置则自动插入一行)
                // ---------------------------------------------------------
                if (result.CopperWeight > 0)
                {
                    // 重新计算最新的元器件终止行 (subsumRow - 1)
                    compEndRow = subsumRow - 1;
                    int targetCopperRow = -1;
                    bool needInsertRow = false;

                    // 4.1 检查元器件区域内是否已存在铜排行 (幂等性保护，防止重复插入多行铜排)
                    if (compEndRow >= compStartRow)
                    {
                        Range checkRange = ws.Range[$"B{compStartRow}:C{compEndRow}"];
                        object[,] checkMatrix = checkRange.Value2 as object[,];
                        if (checkMatrix != null)
                        {
                            int rowCount = checkMatrix.GetLength(0);
                            for (int r = rowCount; r >= 1; r--)
                            {
                                string bName = checkMatrix[r, 1]?.ToString()?.Trim() ?? string.Empty;
                                if (bName == "铜排")
                                {
                                    // 命中既有铜排行，直接就地更新
                                    targetCopperRow = compStartRow + r - 1;
                                    break;
                                }
                            }
                        }
                    }

                    // 4.2 若不存在既有铜排行，检查元器件区域最后一行是否为空位置
                    if (targetCopperRow <= 0)
                    {
                        if (compEndRow >= compStartRow)
                        {
                            // 检查最后一行是否为空 (B列名称与C列型号均为空)
                            string lastB = ws.Range[$"B{compEndRow}"].Text?.ToString()?.Trim() ?? string.Empty;
                            string lastC = ws.Range[$"C{compEndRow}"].Text?.ToString()?.Trim() ?? string.Empty;

                            if (string.IsNullOrWhiteSpace(lastB) && string.IsNullOrWhiteSpace(lastC))
                            {
                                // 最后一行本身即为空行，直接使用该空行填入铜排
                                targetCopperRow = compEndRow;
                            }
                            else
                            {
                                // 最后一行已有元件，属于无空位置，需插入新行
                                needInsertRow = true;
                            }
                        }
                        else
                        {
                            // 当前元器件区域完全没有可用行，需插入新行
                            needInsertRow = true;
                        }

                        // 4.3 若没有空位置，在小计行 (subsumRow) 位置向下推移插入一行
                        if (needInsertRow)
                        {
                            int insertRow = subsumRow;
                            dynamic insertRange = ws.Rows[insertRow];
                            insertRange.Insert(XlInsertShiftDirection.xlShiftDown);

                            // 新插入的行即为元器件区域的最末行
                            targetCopperRow = insertRow;

                            // 插入行后，小计行与总计行物理行号相应后移 1 行
                            subsumRow++;
                            tolsumRow++;
                        }
                    }

                    // 4.4 批量填充元器件最下面一行的铜排明细数据 (规则 7: 2D数组一次性写入 A 到 Q 列)
                    if (targetCopperRow > 0)
                    {
                        Range copperRowRange = ws.Range[$"A{targetCopperRow}:Q{targetCopperRow}"];
                        object[,] copperMatrix = new object[1, 17];

                        // A 列 (索引 0): 动态序号公式
                        copperMatrix[0, 0] = $"=ROW()-ROW(A${detRow + 1})";
                        // B 列 (索引 1): 元件名称
                        copperMatrix[0, 1] = "铜排";
                        // C 列 (索引 2): 规格型号
                        copperMatrix[0, 2] = "TMY";
                        // D 列 (索引 3): 生产厂家
                        copperMatrix[0, 3] = string.Empty;
                        // E 列 (索引 4): 计量单位
                        copperMatrix[0, 4] = "KG";
                        // F 列 (索引 5): 数量 (数量公式)
                        copperMatrix[0, 5] = result.CopperQtyFormula;
                        // G 列 (索引 6): 销售单价
                        copperMatrix[0, 6] = rules.General.CopperPricePerKg;
                        // H 列 (索引 7): 销售总价公式
                        copperMatrix[0, 7] = $"=ROUND(F{targetCopperRow}*G{targetCopperRow},2)";
                        // I 列 (索引 8): 备注
                        copperMatrix[0, 8] = string.Empty;
                        // J 列 (索引 9): 成本单价
                        copperMatrix[0, 9] = rules.General.CopperPricePerKg;
                        // K 列 (索引 10): 成本总价公式
                        copperMatrix[0, 10] = $"=ROUND(F{targetCopperRow}*J{targetCopperRow},2)";
                        // L 列 (索引 11): 加价系数
                        copperMatrix[0, 11] = 1;
                        // M 列 (索引 12): 面价/基准单价
                        copperMatrix[0, 12] = rules.General.CopperPricePerKg;
                        // N 列 (索引 13): 采购折扣系数
                        copperMatrix[0, 13] = 1;
                        // O 列 (索引 14): 预留
                        copperMatrix[0, 14] = string.Empty;
                        // P 列 (索引 15): 预留
                        copperMatrix[0, 15] = string.Empty;
                        // Q 列 (索引 16): 类别
                        copperMatrix[0, 16] = "材料";

                        // 一次性写入目标铜排行
                        copperRowRange.Formula = copperMatrix;
                        result.CopperTargetLocation = $"元器件区域第 {targetCopperRow} 行" + (needInsertRow ? " (自动插入行)" : "");

                        // 4.5 若发生了插入行，必须重新刷新小计行、计费区 A 列序号与单台合计
                        if (needInsertRow)
                        {
                            RefreshCabinetFeeAreaFormulas(ws, detRow, compStartRow, subsumRow, tolsumRow);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                // 记录回写异常日志
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

        /// <summary>
        /// 针对单个指定箱柜执行智能推导分析并写入 Excel 工作表
        /// </summary>
        /// <param name="ws">目标工作表</param>
        /// <param name="detName">箱柜 Det 定义名称 (如 Cab_Det_1)</param>
        /// <param name="rules">定额与规则实体</param>
        /// <returns>包含操作结果、箱柜名称与反馈信息的元组</returns>
        public static (bool Success, string CabinetName, string Message) WriteSingleCabinetAuxAndShell(
            Worksheet ws,
            string detName,
            QuotationRules rules)
        {
            // 基础空值防御校验
            if (ws == null || string.IsNullOrWhiteSpace(detName))
            {
                // 参数缺失时直接返回失败
                return (false, string.Empty, "工作表或箱柜标识为空。");
            }

            try
            {
                // 扫描目标箱柜的元器件区域与行号信息
                var scanData = ScanCabinetData(ws, detName);
                if (scanData == null)
                {
                    // 扫描失败返回明确提示信息
                    return (false, string.Empty, $"未找到箱柜【{detName}】的有效数据或锚点定义。");
                }

                // 执行壳体、铜排、辅材与装配人工费的智能推导
                var result = CalculateCabinetAuxAndShell(scanData, rules);
                if (result == null)
                {
                    // 推导失败提示
                    return (false, scanData.CabinetName, $"箱柜【{scanData.CabinetName}】推导计算失败。");
                }

                // 将计算结果与动态算式公式批量回写至 Excel
                bool ok = WriteCabinetCalcResultToSheet(ws, scanData, result, rules);
                if (ok)
                {
                    // 回写成功返回箱柜名称
                    return (true, scanData.CabinetName, $"成功写入箱柜【{scanData.CabinetName}】的推导数据与公式！");
                }
                else
                {
                    // 回写异常提示
                    return (false, scanData.CabinetName, $"箱柜【{scanData.CabinetName}】回写数据至工作表失败。");
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志并安全返回
                System.Diagnostics.Debug.WriteLine($"[WriteSingleCabinetAuxAndShell] 写入单柜异常: {ex.Message}");
                return (false, string.Empty, $"写入过程发生异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 更新当前指定分类表中的所有箱柜 (自底向上倒序更新，保障行号绝对稳定)
        /// </summary>
        /// <param name="ws">目标分类工作表</param>
        /// <param name="rules">定额与规则实体</param>
        /// <returns>包含是否成功、更新箱柜数量与反馈消息的元组</returns>
        public static (bool Success, int UpdatedCabinets, string Message) UpdateCurrentCategoryAuxAndShell(
            Worksheet ws,
            QuotationRules rules)
        {
            // 基础空值校验
            if (ws == null) return (false, 0, "工作表对象为空。");

            try
            {
                // 获取当前工作表的所属工作簿引用
                Workbook wb = ws.Parent as Workbook;
                // 探测当前分类表下的所有有效箱柜定义
                var validCabinets = Tool.GetSheetValidCabinets(ws, wb);

                // 若未识别到箱柜锚点，尝试执行一次定义名称补齐自愈
                if (validCabinets.Count == 0)
                {
                    // 触发工作表定义名称自动探测补齐
                    Tool.FixAndFillCabinetNamesForSheet(ws);
                    // 重新提取有效箱柜列表
                    validCabinets = Tool.GetSheetValidCabinets(ws, wb);
                }

                // 筛选出包含 Det 明细行的有效箱柜
                var targetCabinets = new List<KeyValuePair<int, CabinetAnchorModel>>();
                foreach (var kvp in validCabinets)
                {
                    // 仅处理具备明细块的箱柜
                    if (kvp.Value?.Det != null)
                    {
                        targetCabinets.Add(kvp);
                    }
                }

                // 若无有效箱柜，直接安全返回
                if (targetCabinets.Count == 0)
                {
                    // 返回无箱柜提示
                    return (false, 0, $"工作表【{ws.Name}】中未检测到有效的明细箱柜。");
                }

                // 核心法则: 按照 Det 物理行号从大到小倒序排列 (自底向上遍历)
                // 优势: 即使下方箱柜因缺少空行插入了新铜排行，绝不影响上方任何箱柜的物理行号！
                targetCabinets.Sort((a, b) =>
                {
                    int rowA = Convert.ToInt32(a.Value.Det.Row);
                    int rowB = Convert.ToInt32(b.Value.Det.Row);
                    return rowB.CompareTo(rowA);
                });

                int successCount = 0;
                // 遍历倒序箱柜集合执行推导与回写
                foreach (var kvp in targetCabinets)
                {
                    // 拼接当前箱柜的 Det 定义名称
                    string detName = $"Cab_Det_{kvp.Key}";
                    // 动态重新扫描该箱柜数据 (每次读取最新实际行号)
                    var scanData = ScanCabinetData(ws, detName);
                    if (scanData == null) continue;

                    // 计算推导结果
                    var result = CalculateCabinetAuxAndShell(scanData, rules);
                    if (result == null) continue;

                    // 回写计算结果至当前工作表
                    bool ok = WriteCabinetCalcResultToSheet(ws, scanData, result, rules);
                    if (ok)
                    {
                        // 递增成功计数器
                        successCount++;
                    }
                }

                // 返回成功更新的箱柜总台数
                return (successCount > 0, successCount, $"成功更新分类表【{ws.Name}】共 {successCount} 台箱柜的数据与公式！");
            }
            catch (Exception ex)
            {
                // 记录异常日志
                System.Diagnostics.Debug.WriteLine($"[UpdateCurrentCategoryAuxAndShell] 更新当前分类异常: {ex.Message}");
                return (false, 0, $"更新分类表【{ws.Name}】发生异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 更新当前工作簿中所有分类表的全部箱柜 (全工作簿多表倒序更新，带屏幕刷新性能保护)
        /// </summary>
        /// <param name="rules">定额与规则实体</param>
        /// <returns>包含是否成功、更新工作表总数、更新箱柜总数与提示消息的元组</returns>
        public static (bool Success, int UpdatedSheets, int UpdatedCabinets, string Message) UpdateAllCategoriesAuxAndShell(
            QuotationRules rules)
        {
            // 获取当前 Excel Application COM 接口实例
            dynamic? app = ExcelDnaSafeAccessor.GetApplication();
            if (app == null) return (false, 0, 0, "无法连接到 Excel 应用程序。");

            // 获取当前活动工作簿
            Workbook activeWb = app.ActiveWorkbook;
            if (activeWb == null) return (false, 0, 0, "当前没有打开的 Excel 工作簿。");

            // 记录原始活动工作表，以便处理完成后平滑恢复激活
            Worksheet? origActiveSheet = null;
            try { origActiveSheet = app.ActiveSheet as Worksheet; } catch { }

            // 临时关闭屏幕刷新与系统提示以提升 5~10 倍批量更新吞吐性能
            app.ScreenUpdating = false;
            app.DisplayAlerts = false;

            int totalUpdatedSheets = 0;
            int totalUpdatedCabinets = 0;

            try
            {
                // 遍历活动工作簿下的每一个 Worksheet
                foreach (Worksheet ws in activeWb.Worksheets)
                {
                    try
                    {
                        // 跳过隐藏工作表
                        if (ws.Visible != XlSheetVisibility.xlSheetVisible) continue;

                        // 提取纯文本工作表名称
                        string wsName = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;

                        // 排除明确的系统非分类辅助表 (如 项目信息、元件汇总表、元器件数据管理) --硬编码: 系统保留辅助工作表名称--
                        if (string.Equals(wsName, "项目信息", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(wsName, Models.ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(wsName, "元器件数据管理", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        // 激活工作表以避免跨表操作时的 Excel 内部 1004 COM 异常
                        ws.Activate();

                        // 调用当前分类表的批量倒序更新方法
                        var categoryResult = UpdateCurrentCategoryAuxAndShell(ws, rules);
                        if (categoryResult.Success && categoryResult.UpdatedCabinets > 0)
                        {
                            // 累加成功更新的工作表数量与箱柜数量
                            totalUpdatedSheets++;
                            totalUpdatedCabinets += categoryResult.UpdatedCabinets;
                        }
                    }
                    catch (Exception exSheet)
                    {
                        // 记录单表处理异常，不阻断后续工作表处理
                        System.Diagnostics.Debug.WriteLine($"[UpdateAllCategoriesAuxAndShell] 处理表【{ws.Name}】异常: {exSheet.Message}");
                    }
                }

                // 构建成功返回消息
                string msg = totalUpdatedSheets > 0
                    ? $"成功更新 {totalUpdatedSheets} 个分类表，共 {totalUpdatedCabinets} 台箱柜的数据与公式！"
                    : "未在工作簿中检测到可更新的有效分类表。";

                return (totalUpdatedSheets > 0, totalUpdatedSheets, totalUpdatedCabinets, msg);
            }
            finally
            {
                // 恢复原始活动工作表焦点
                try
                {
                    origActiveSheet?.Activate();
                }
                catch { }

                // 强制恢复系统屏幕刷新与警告提示，杜绝界面冻结
                app.ScreenUpdating = true;
                app.DisplayAlerts = true;
            }
        }
    }
}
