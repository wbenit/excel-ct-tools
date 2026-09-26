using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;
using ExcelAddInDemo.Models;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：三箱型号智能管道判定与批量添写功能
    /// 严格遵循规则 6 (4个定义名称)、规则 7 (数组一次性吞吐)、规则 8 (FixAndFill 自愈) 与每 3 行 1 行中文注释规范
    /// </summary>
    public static partial class ExcelServices
    {
        // 管道规则配置文件持久化名称
        private const string PipelineConfigFileName = "cabinet_pipeline_rules.json"; // --硬编码: 规则配置文件名--

        // 三箱型号添写窗体静态引用 (非模态单例保持)
        private static Forms.CabinetModelPipelineForm? _cabinetModelPipelineForm = null;

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“三箱型号添写向导”窗口 (非模态，保持 Excel 可交互)
        /// </summary>
        public static void ShowCabinetModelPipelineDialog()
        {
            try
            {
                // 以非模态方式展示向导窗口
                ShowModelessForm(ref _cabinetModelPipelineForm, () => new Forms.CabinetModelPipelineForm());
            }
            catch (Exception ex)
            {
                // 捕获异常防止 Excel 崩溃闪退
                LogHelper.WriteLog($"[三箱管道] 弹出向导窗口异常: {ex.Message}");
            }
        }

        // JSON 序列化配置选项
        private static readonly JsonSerializerOptions PipelineJsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// 获取三箱型号管道初始配置与默认规则
        /// 若用户本地无自定义配置文件，则自动生成符合行业标准的默认规则集合
        /// </summary>
        /// <returns>管道全局配置模型</returns>
        public static CabinetPipelineConfig LoadCabinetPipelineConfig()
        {
            try
            {
                // 获取用户自定义数据目录物理路径
                string dataDir = Tool.GetCustomDataDirectoryFromGlobalConfig();
                if (string.IsNullOrWhiteSpace(dataDir))
                {
                    // 回退至应用程序默认数据目录
                    dataDir = Tool.GetAppDataDirectory();
                }

                // 拼接配置文件的绝对路径
                string configFilePath = Path.Combine(dataDir, PipelineConfigFileName);

                // 若本地存在配置文件，直接反序列化读取
                if (File.Exists(configFilePath))
                {
                    string jsonContent = File.ReadAllText(configFilePath);
                    var loadedConfig = JsonSerializer.Deserialize<CabinetPipelineConfig>(jsonContent, PipelineJsonOptions);
                    if (loadedConfig != null && loadedConfig.Rules != null && loadedConfig.Rules.Count > 0)
                    {
                        return loadedConfig;
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[三箱型号管道] 读取本地配置文件异常: {ex.Message}");
            }

            // 本地无有效配置时，构建并返回默认标准预设规则集
            return CreateDefaultPipelineConfig();
        }

        /// <summary>
        /// 将用户修改后的管道规则及配置持久化保存至本地 JSON 文件
        /// </summary>
        /// <param name="config">待保存的管道配置实体</param>
        /// <returns>保存是否成功</returns>
        public static bool SaveCabinetPipelineConfig(CabinetPipelineConfig config)
        {
            if (config == null) return false;

            try
            {
                // 获取用户数据目录路径
                string dataDir = Tool.GetCustomDataDirectoryFromGlobalConfig();
                if (string.IsNullOrWhiteSpace(dataDir))
                {
                    dataDir = Tool.GetAppDataDirectory();
                }

                // 确保数据目录物理存在
                if (!Directory.Exists(dataDir))
                {
                    Directory.CreateDirectory(dataDir);
                }

                // 拼接并写入目标文件
                string configFilePath = Path.Combine(dataDir, PipelineConfigFileName);
                string jsonString = JsonSerializer.Serialize(config, PipelineJsonOptions);
                File.WriteAllText(configFilePath, jsonString);
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[三箱型号管道] 保存配置文件异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 构建开箱即用的行业经典三箱型号预设规则集合
        /// 覆盖 ATS双电源箱、XL-21动力配电柜、PZ30照明配电箱、JXF基业箱及兜底规则
        /// </summary>
        /// <returns>标准预设配置实体</returns>
        public static CabinetPipelineConfig CreateDefaultPipelineConfig()
        {
            var config = new CabinetPipelineConfig
            {
                FallbackStrategy = "original", // --硬编码: 默认策略--
                FallbackCustomModel = "待确认", // --硬编码: 默认兜底型号--
                WriteToSummaryN = true,
                WriteToSummaryD = false,
                WriteToDetailD = false,
                Rules = new List<CabinetPipelineRule>()
            };

            // Stage 1: ATS 双电源转换箱 (默认无尺寸限制，仅保留特征词)
            config.Rules.Add(new CabinetPipelineRule
            {
                Id = "rule_ats",
                Name = "Stage 1: 双电源箱 (ATS)",
                TargetModel = "ATS", // --硬编码: 目标型号--
                IsEnabled = true,
                SortOrder = 1,
                HeightCondition = "none",
                HeightMin = 0,
                HeightMax = 0,
                DepthCondition = "none",
                DepthMin = 0,
                DepthMax = 0,
                ComponentKeywords = new List<string> { "双电源", "ATS" }, // --硬编码--
                ComponentMatchMode = "any",
                ExcludeComponentKeywords = new List<string>(),
                NameKeywords = new List<string>(),
                Description = "双电源箱 (默认无尺寸限制，按需配置)"
            });

            // Stage 2: XL-21 动力配电柜 (去除默认尺寸和排除项，完全放权给用户自由设定)
            config.Rules.Add(new CabinetPipelineRule
            {
                Id = "rule_xl21",
                Name = "Stage 2: 落地动力柜 (XL-21)",
                TargetModel = "XL-21", // --硬编码: 目标型号--
                IsEnabled = true,
                SortOrder = 2,
                HeightCondition = "none",
                HeightMin = 0,
                HeightMax = 0,
                DepthCondition = "none",
                DepthMin = 0,
                DepthMax = 0,
                ComponentKeywords = new List<string>(),
                ComponentMatchMode = "none",
                ExcludeComponentKeywords = new List<string>(),
                NameKeywords = new List<string>(),
                Description = "落地动力柜 (默认无限制，完全由用户自由设定)"
            });

            // Stage 3: PZ30 模数化终端照明箱 (彻底去除140/900默认限制及塑壳排除词，避免打架)
            config.Rules.Add(new CabinetPipelineRule
            {
                Id = "rule_pz30",
                Name = "Stage 3: 终端照明箱 (PZ30)",
                TargetModel = "PZ30", // --硬编码: 目标型号--
                IsEnabled = true,
                SortOrder = 3,
                HeightCondition = "none",
                HeightMin = 0,
                HeightMax = 0,
                DepthCondition = "none",
                DepthMin = 0,
                DepthMax = 0,
                ComponentKeywords = new List<string>(),
                ComponentMatchMode = "none",
                ExcludeComponentKeywords = new List<string>(),
                NameKeywords = new List<string>(),
                Description = "终端照明箱 (默认无限制，完全由用户自由设定)"
            });

            // Stage 4: JXF 基业箱 / 壁挂动力箱 (彻底去除默认区间，避免打架)
            config.Rules.Add(new CabinetPipelineRule
            {
                Id = "rule_jxf",
                Name = "Stage 4: 基业控制箱 (JXF)",
                TargetModel = "JXF", // --硬编码: 目标型号--
                IsEnabled = true,
                SortOrder = 4,
                HeightCondition = "none",
                HeightMin = 0,
                HeightMax = 0,
                DepthCondition = "none",
                DepthMin = 0,
                DepthMax = 0,
                ComponentKeywords = new List<string>(),
                ComponentMatchMode = "none",
                ExcludeComponentKeywords = new List<string>(),
                NameKeywords = new List<string>(),
                Description = "通用基业箱 (默认无限制，完全由用户自由设定)"
            });

            return config;
        }

        /// <summary>
        /// 获取三箱型号添写初始化数据包
        /// 自动扫描当前工作簿中的所有分类表，并对当前选中的分类表执行管道识别推导
        /// </summary>
        /// <param name="targetSheetName">可选指定的工作表名称，为空时默认取当前活动工作表</param>
        /// <returns>初始化数据包模型</returns>
        public static CabinetPipelineInitDataDto GetCabinetPipelineInitData(string? targetSheetName = null)
        {
            // 初始化返回的数据实体包
            var initData = new CabinetPipelineInitDataDto();

            try
            {
                // 获取 Excel 宿主应用程序实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    LogHelper.WriteLog("[三箱管道] 获取 Excel Application 失败");
                    return initData;
                }

                // 获取活动算价工作簿
                dynamic? wb = app.ActiveWorkbook;
                if (wb == null)
                {
                    LogHelper.WriteLog("[三箱管道] 未检测到活动工作簿");
                    return initData;
                }

                // 加载或创建当前管道规则配置
                initData.Config = LoadCabinetPipelineConfig();
                // 防御性检查：若规则集为空，自动重建默认四级规则并持久化
                if (initData.Config.Rules == null || initData.Config.Rules.Count == 0)
                {
                    initData.Config = CreateDefaultPipelineConfig();
                    SaveCabinetPipelineConfig(initData.Config);
                }
                // 按优先级排序规则
                initData.Rules = initData.Config.Rules.OrderBy(r => r.SortOrder).ToList();

                // 搜集工作表候选集合
                var candidateSheetNames = new List<string>();
                var sheetsWithCabinets = new List<string>();

                // 1. 若当前工作簿为成套报价工程工作簿，优先读取项目信息登记的分类表白名单
                HashSet<string> projectCategories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                try
                {
                    if (Tool.IsProjectWorkbook(wb))
                    {
                        projectCategories = Tool.GetProjectCategorySheetNames(wb);
                    }
                }
                catch { }

                // 获取当前活动工作表名称
                string currentActiveSheetName = "";
                try
                {
                    dynamic? activeWs = app.ActiveSheet;
                    currentActiveSheetName = activeWs != null ? Convert.ToString(activeWs.Name)?.Trim() ?? "" : "";
                }
                catch { }

                // 遍历工作簿中的所有工作表 (使用 dynamic 避免图表工作表等强转 COM 失败)
                foreach (dynamic ws in wb.Worksheets)
                {
                    string sName = "";
                    try { sName = Convert.ToString(ws.Name)?.Trim() ?? ""; } catch { }
                    if (string.IsNullOrWhiteSpace(sName)) continue;

                    // 排除系统保留与报表工作表 (项目信息、清单、封面、调价等)
                    if (Tool.IsReservedOrReportSheet(sName) ||
                        sName.Contains("清单") ||
                        sName.Contains("项目信息") ||
                        sName.Contains("成品交接") ||
                        sName.Contains("调价") ||
                        sName.Contains("封面"))
                    {
                        continue;
                    }

                    // 若登记在工程白名单中，或者未定义白名单 (普通工作簿放行所有非保留表)
                    if (projectCategories.Count == 0 || projectCategories.Contains(sName))
                    {
                        candidateSheetNames.Add(sName);
                        // 探测该工作表是否有箱柜锚点 (规则 6 架构)
                        try
                        {
                            var validCabs = Tool.GetSheetValidCabinets(ws, wb);
                            if (validCabs != null && validCabs.Count > 0)
                            {
                                sheetsWithCabinets.Add(sName);
                            }
                        }
                        catch { }
                    }
                }

                // 容错兜底：若白名单过滤后无任何表，但工作簿存在非保留表，全量放行非保留表
                if (candidateSheetNames.Count == 0)
                {
                    foreach (dynamic ws in wb.Worksheets)
                    {
                        string sName = "";
                        try { sName = Convert.ToString(ws.Name)?.Trim() ?? ""; } catch { }
                        if (string.IsNullOrWhiteSpace(sName)) continue;
                        if (!Tool.IsReservedOrReportSheet(sName))
                        {
                            candidateSheetNames.Add(sName);
                        }
                    }
                }

                // 回填工作表候选列表至 DTO
                initData.SheetNames = candidateSheetNames;

                // 确定当前选中的工作表 (优先级：指定表 > 当前活动表 > 含有箱柜的表 > 首个候选表)
                string chosenSheetName = "";
                if (!string.IsNullOrWhiteSpace(targetSheetName) && candidateSheetNames.Contains(targetSheetName))
                {
                    chosenSheetName = targetSheetName;
                }
                else if (!string.IsNullOrWhiteSpace(currentActiveSheetName) && candidateSheetNames.Contains(currentActiveSheetName))
                {
                    chosenSheetName = currentActiveSheetName;
                }
                else if (sheetsWithCabinets.Count > 0)
                {
                    chosenSheetName = sheetsWithCabinets[0];
                }
                else if (candidateSheetNames.Count > 0)
                {
                    chosenSheetName = candidateSheetNames[0];
                }

                // 记录最终选中的表名
                initData.SelectedSheetName = chosenSheetName;

                // 若定位到了有效的目标工作表，执行箱柜扫描与管道判定
                if (!string.IsNullOrWhiteSpace(chosenSheetName))
                {
                    try
                    {
                        dynamic targetWs = wb.Worksheets[chosenSheetName];
                        initData.Cabinets = ScanAndEvaluateCabinetsForSheet(targetWs, wb, initData.Config);
                    }
                    catch (Exception ex)
                    {
                        LogHelper.WriteLog($"[三箱管道] 扫描工作表 [{chosenSheetName}] 失败: {ex.Message}");
                    }
                }

                // 输出详细的成功日志
                LogHelper.WriteLog($"[三箱管道] 初始化就绪: 候选表数={candidateSheetNames.Count}, 选定表=[{chosenSheetName}], 识别箱柜数={initData.Cabinets.Count}");
            }
            catch (Exception ex)
            {
                // 异常统一捕获与记录
                LogHelper.WriteLog($"[三箱型号管道] 获取初始化数据异常: {ex.Message}");
            }

            // 返回初始化实体
            return initData;
        }

        /// <summary>
        /// 针对单个工作表执行箱柜扫描、尺寸解析、元器件特征提取与管道型号判定
        /// 严格遵循规则 6、7、8
        /// </summary>
        /// <param name="ws">目标工作表 (支持 dynamic COM 对象)</param>
        /// <param name="wb">工作簿实例</param>
        /// <param name="config">管道配置规则集</param>
        /// <returns>扫描出的箱柜判定列表</returns>
        public static List<CabinetPipelineItemDto> ScanAndEvaluateCabinetsForSheet(
            dynamic ws,
            dynamic wb,
            CabinetPipelineConfig config)
        {
            var resultList = new List<CabinetPipelineItemDto>();
            if (ws == null || wb == null) return resultList;

            try
            {
                // 1. 遵循规则 8: 优先执行自愈与拓扑校验，确保规则 6 锚点名称规范且存在
                Tool.FixAndFillCabinetNamesForSheet(ws);

                // 2. 提取当前工作表所有有效的箱柜锚点字典 (有序排列)
                var validCabinets = Tool.GetSheetValidCabinets(ws, wb);
                if (validCabinets == null || validCabinets.Count == 0) return resultList;

                string sheetName = ws.Name;
                // 获取当前工作表的已用行范围
                int usedEndRow = ws.UsedRange.Rows.Count;

                // 3. 遍历每一个有效箱柜进行特征提取与判定
                foreach (var kvp in validCabinets)
                {
                    int cabK = kvp.Key;
                    var anchor = kvp.Value;
                    if (anchor == null) continue;

                    // 提取各个锚点行号 (规则 6 架构)
                    int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                    int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;
                    int subsumRow = anchor.Subsum != null ? Convert.ToInt32(anchor.Subsum.Row) : 0;
                    int tolsumRow = anchor.Tolsum != null ? Convert.ToInt32(anchor.Tolsum.Row) : 0;

                    var item = new CabinetPipelineItemDto
                    {
                        SheetName = sheetName,
                        CabinetIndex = cabK,
                        SumRow = sumRow,
                        DetRow = detRow,
                        CabinetNo = $"箱柜{cabK}"
                    };

                    // -------------------------------------------------------------
                    // 步骤 A: 从顶部汇总行提取基础柜号、名称、原型号与箱柜尺寸 (直接使用 sum 行 M 列)
                    // -------------------------------------------------------------
                    if (sumRow > 0)
                    {
                        try
                        {
                            // 规则 7: 数组一次性读取顶部汇总行 A 到 N 列 (覆盖至第 14 列 N 列)
                            Range sumRange = ws.Range[$"A{sumRow}:N{sumRow}"];
                            object[,] sumMatrix = sumRange.Value2 as object[,];
                            if (sumMatrix != null)
                            {
                                int colLen = sumMatrix.GetLength(1);

                                // B 列为柜号 (第 2 列)
                                string bNo = sumMatrix[1, 2]?.ToString()?.Trim() ?? "";
                                if (!string.IsNullOrWhiteSpace(bNo)) item.CabinetNo = bNo;

                                // C 列为箱柜名称 (第 3 列)
                                string cName = sumMatrix[1, 3]?.ToString()?.Trim() ?? "";
                                if (!string.IsNullOrWhiteSpace(cName)) item.CabinetName = cName;

                                // D 列为原箱柜型号 (第 4 列)
                                string dModel = sumMatrix[1, 4]?.ToString()?.Trim() ?? "";

                                // N 列为现存回填的三箱型号 (第 14 列，若已回填过则优先展示)
                                string nModel = colLen >= 14 ? sumMatrix[1, 14]?.ToString()?.Trim() ?? "" : "";
                                item.CurrentModel = !string.IsNullOrWhiteSpace(nModel) ? nModel : dModel;

                                // M 列为箱柜尺寸 (第 13 列，用户明确指定尺寸直接取 sum 行 M 列)
                                string mSize = colLen >= 13 ? sumMatrix[1, 13]?.ToString()?.Trim() ?? "" : "";
                                if (!string.IsNullOrWhiteSpace(mSize) && TryParseShellDimensions(mSize, out int mw, out int mh, out int md))
                                {
                                    item.Width = mw;
                                    item.Height = mh;
                                    item.Depth = md;
                                    item.RawDimensionText = mSize;
                                }
                            }
                        }
                        catch (Exception exSum)
                        {
                            LogHelper.WriteLog($"[三箱管道] 提取汇总行箱柜{cabK}异常: {exSum.Message}");
                        }
                    }

                    // -------------------------------------------------------------
                    // 步骤 C: 依据规则 6 扫描元器件区域 (Cab_Det+2 到 Cab_Subsum-1)
                    // -------------------------------------------------------------
                    int compStartRow = detRow > 0 ? detRow + 2 : 0;
                    int compEndRow = subsumRow > 0 ? subsumRow - 1 : 0;

                    var compNameList = new List<string>();
                    var compModelList = new List<string>();
                    var compCountDict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    if (compStartRow > 0 && compEndRow >= compStartRow)
                    {
                        try
                        {
                            // 规则 7: 二维数组一次性读取元器件区域 B 列(名称)与 C 列(型号)
                            Range compRange = ws.Range[$"B{compStartRow}:C{compEndRow}"];
                            object[,] compMatrix = compRange.Value2 as object[,];
                            if (compMatrix != null)
                            {
                                int cRowCount = compMatrix.GetLength(0);
                                for (int r = 1; r <= cRowCount; r++)
                                {
                                    string cName = compMatrix[r, 1]?.ToString()?.Trim() ?? "";
                                    string cModel = compMatrix[r, 2]?.ToString()?.Trim() ?? "";

                                    // 排除空行或表头无效行
                                    if (string.IsNullOrWhiteSpace(cName) && string.IsNullOrWhiteSpace(cModel)) continue;
                                    if (cName == "元件名称" || cName == "名称" || cName == "型号规格") continue;

                                    if (!string.IsNullOrWhiteSpace(cName))
                                    {
                                        compNameList.Add(cName);
                                        // 聚合计数用于摘要展示
                                        string shortName = SimplifyComponentName(cName);
                                        if (compCountDict.ContainsKey(shortName)) compCountDict[shortName]++;
                                        else compCountDict[shortName] = 1;
                                    }
                                    if (!string.IsNullOrWhiteSpace(cModel))
                                    {
                                        compModelList.Add(cModel);
                                    }
                                }
                            }
                        }
                        catch (Exception exComp)
                        {
                            LogHelper.WriteLog($"[三箱管道] 提取元器件区域异常: {exComp.Message}");
                        }
                    }

                    item.ComponentNames = compNameList;
                    item.ComponentModels = compModelList;

                    // 构造简洁易懂的元器件摘要 (例如: "微断*8, 塑壳*1, 浪涌*1")
                    if (compCountDict.Count > 0)
                    {
                        item.ComponentsSummary = string.Join(", ", compCountDict.Take(5).Select(kv => $"{kv.Key}*{kv.Value}"));
                        if (compCountDict.Count > 5)
                        {
                            item.ComponentsSummary += "...";
                        }
                    }
                    else
                    {
                        item.ComponentsSummary = "无元件明细"; // --硬编码--
                    }

                    // -------------------------------------------------------------
                    // 步骤 D: 将箱柜数据注入管道规则引擎执行匹配推导
                    // -------------------------------------------------------------
                    EvaluateSingleCabinetInPipeline(item, config);

                    resultList.Add(item);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[三箱型号管道] ScanAndEvaluateCabinetsForSheet 异常: {ex.Message}");
            }

            return resultList;
        }

        /// <summary>
        /// 核心管道决策引擎：让单台箱柜依次通过管道各个阶段，产出命中结果与推荐型号
        /// </summary>
        /// <param name="item">箱柜数据实体</param>
        /// <param name="config">管道配置</param>
        public static void EvaluateSingleCabinetInPipeline(CabinetPipelineItemDto item, CabinetPipelineConfig config)
        {
            if (item == null || config == null) return;

            // 获取按优先级排序的有效规则流
            var activeRules = (config.Rules ?? new List<CabinetPipelineRule>())
                .Where(r => r.IsEnabled)
                .OrderBy(r => r.SortOrder)
                .ToList();

            bool isMatched = false;

            // 依次流经各个规则阶段
            foreach (var rule in activeRules)
            {
                if (IsCabinetMatchRule(item, rule))
                {
                    // 命中第一条规则，立即中断流转 (Match and Stop)
                    item.HitRuleId = rule.Id;
                    item.HitRuleName = rule.Name;
                    item.SuggestedModel = rule.TargetModel;
                    item.FinalModel = rule.TargetModel;
                    isMatched = true;
                    break;
                }
            }

            // 若所有管道阶段均未命中，执行兜底策略
            if (!isMatched)
            {
                item.HitRuleId = "fallback";
                item.HitRuleName = "未命中管道规则 (兜底)";

                if (config.FallbackStrategy == "original" && !string.IsNullOrWhiteSpace(item.CurrentModel))
                {
                    // 保留工作表中原有的型号
                    item.SuggestedModel = item.CurrentModel;
                    item.FinalModel = item.CurrentModel;
                }
                else if (config.FallbackStrategy == "custom" && !string.IsNullOrWhiteSpace(config.FallbackCustomModel))
                {
                    // 指定的自定义兜底型号
                    item.SuggestedModel = config.FallbackCustomModel;
                    item.FinalModel = config.FallbackCustomModel;
                }
                else
                {
                    // 标记为待确认
                    item.SuggestedModel = "待确认"; // --硬编码--
                    item.FinalModel = !string.IsNullOrWhiteSpace(item.CurrentModel) ? item.CurrentModel : "待确认";
                }
            }
        }

        /// <summary>
        /// 判断单台箱柜是否满足某条特定管道规则的所有前置约束条件
        /// </summary>
        /// <param name="item">箱柜明细数据</param>
        /// <param name="rule">规则实体</param>
        /// <returns>是否匹配成功</returns>
        private static bool IsCabinetMatchRule(CabinetPipelineItemDto item, CabinetPipelineRule rule)
        {
            // 1. 检查排除元器件关键词 (一票否决)
            if (rule.ExcludeComponentKeywords != null && rule.ExcludeComponentKeywords.Count > 0)
            {
                foreach (var exKey in rule.ExcludeComponentKeywords)
                {
                    if (string.IsNullOrWhiteSpace(exKey)) continue;
                    // 若任何元器件名称或型号包含排他关键词，则直接不匹配
                    if (item.ComponentNames.Any(n => n.IndexOf(exKey, StringComparison.OrdinalIgnoreCase) >= 0) ||
                        item.ComponentModels.Any(m => m.IndexOf(exKey, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        return false;
                    }
                }
            }

            // 2. 检查高度条件
            if (!CheckNumericCondition(item.Height, rule.HeightCondition, rule.HeightMin, rule.HeightMax))
            {
                return false;
            }

            // 3. 检查深度条件
            if (!CheckNumericCondition(item.Depth, rule.DepthCondition, rule.DepthMin, rule.DepthMax))
            {
                return false;
            }

            // 4. 检查宽度条件
            if (!CheckNumericCondition(item.Width, rule.WidthCondition, rule.WidthMin, rule.WidthMax))
            {
                return false;
            }

            // 5. 检查元器件包含关键词条件
            if (rule.ComponentKeywords != null && rule.ComponentKeywords.Count > 0)
            {
                var validKeys = rule.ComponentKeywords.Where(k => !string.IsNullOrWhiteSpace(k)).ToList();
                if (validKeys.Count > 0)
                {
                    if (rule.ComponentMatchMode == "all")
                    {
                        // 必须全部包含
                        foreach (var key in validKeys)
                        {
                            bool hasKey = item.ComponentNames.Any(n => n.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0) ||
                                          item.ComponentModels.Any(m => m.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0);
                            if (!hasKey) return false;
                        }
                    }
                    else if (rule.ComponentMatchMode == "any")
                    {
                        // 包含任意一个即可
                        bool hasAny = validKeys.Any(key =>
                            item.ComponentNames.Any(n => n.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            item.ComponentModels.Any(m => m.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0)
                        );
                        if (!hasAny) return false;
                    }
                }
            }

            // 6. 检查柜号/箱柜名称辅助包含关键词 (若配置了则参与校验)
            if (rule.NameKeywords != null && rule.NameKeywords.Count > 0)
            {
                var validNameKeys = rule.NameKeywords.Where(k => !string.IsNullOrWhiteSpace(k)).ToList();
                if (validNameKeys.Count > 0)
                {
                    bool nameMatch = validNameKeys.Any(k =>
                        item.CabinetNo.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        item.CabinetName.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0
                    );
                    // 若规则配置了纯名称强校验，则需匹配
                    // 这里作为加分/协同校验：如果规则既无尺寸又无元件约束，必须满足名称条件
                    bool isPureNameRule = (rule.HeightCondition == "none" && rule.DepthCondition == "none" && (rule.ComponentKeywords == null || rule.ComponentKeywords.Count == 0));
                    if (isPureNameRule && !nameMatch)
                    {
                        return false;
                    }
                }
            }

            // 7. 兜底保护：检查规则是否配置了至少一项过滤判定条件
            bool hasHeightCond = !string.IsNullOrWhiteSpace(rule.HeightCondition) && !rule.HeightCondition.Equals("none", StringComparison.OrdinalIgnoreCase);
            // 检查是否存在深度判定
            bool hasDepthCond = !string.IsNullOrWhiteSpace(rule.DepthCondition) && !rule.DepthCondition.Equals("none", StringComparison.OrdinalIgnoreCase);
            // 检查是否存在宽度判定
            bool hasWidthCond = !string.IsNullOrWhiteSpace(rule.WidthCondition) && !rule.WidthCondition.Equals("none", StringComparison.OrdinalIgnoreCase);
            // 检查是否存在元器件包含关键词
            bool hasComponentKws = rule.ComponentKeywords != null && rule.ComponentKeywords.Any(k => !string.IsNullOrWhiteSpace(k));
            // 检查是否存在元器件排除关键词
            bool hasExcludeKws = rule.ExcludeComponentKeywords != null && rule.ExcludeComponentKeywords.Any(k => !string.IsNullOrWhiteSpace(k));
            // 检查是否存在箱柜名称关键词
            bool hasNameKws = rule.NameKeywords != null && rule.NameKeywords.Any(k => !string.IsNullOrWhiteSpace(k));

            // 若规则完全未配置任何条件，则视为未配置阶段，不盲目命中全部，避免吞噬下游打架
            if (!hasHeightCond && !hasDepthCond && !hasWidthCond && !hasComponentKws && !hasExcludeKws && !hasNameKws)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 数值条件比较通用辅助函数 (支持 none, >=, <=, between, ==)
        /// </summary>
        /// <summary>
        /// 数值条件比较通用辅助函数 (支持 none, >=, <=, between, ==)
        /// 具备高度容错：单输入框无论存于 min 还是 max 均能精准完成逻辑匹配
        /// </summary>
        private static bool CheckNumericCondition(int value, string condition, int min, int max)
        {
            // 条件为空或 none 直接判定为无限制放行
            if (string.IsNullOrWhiteSpace(condition) || condition.Equals("none", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // 若被检测的箱体尺寸为 0 (未提取到有效尺寸)，无法满足严格的尺寸判断条件
            if (value <= 0) return false;

            // 分支处理各类条件逻辑
            switch (condition.ToLowerInvariant())
            {
                // 大于等于门限
                case ">=":
                    int lowerThreshold = min > 0 ? min : (max > 0 ? max : 0);
                    return lowerThreshold > 0 ? value >= lowerThreshold : true;

                // 小于等于上限 (优先取 max，若用户在单输入框输入则取 min)
                case "<=":
                    int upperThreshold = max > 0 ? max : (min > 0 ? min : 0);
                    return upperThreshold > 0 ? value <= upperThreshold : true;

                // 介于区间
                case "between":
                    if (min > 0 && max > 0) return value >= min && value <= max;
                    if (min > 0) return value >= min;
                    if (max > 0) return value <= max;
                    return true;

                // 精确等于
                case "==":
                    int eq = min > 0 ? min : max;
                    return eq > 0 ? value == eq : true;

                default:
                    return true;
            }
        }

        /// <summary>
        /// 简化元件名称为易于阅读的标签短词 (如 "小型断路器" -> "微断")
        /// </summary>
        private static string SimplifyComponentName(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName)) return "";
            if (fullName.Contains("小型断路器") || fullName.Contains("微型断路器")) return "微断"; // --硬编码--
            if (fullName.Contains("塑壳断路器")) return "塑壳"; // --硬编码--
            if (fullName.Contains("双电源")) return "双电源"; // --硬编码--
            if (fullName.Contains("刀开关") || fullName.Contains("刀熔")) return "刀开关"; // --硬编码--
            if (fullName.Contains("浪涌")) return "浪涌"; // --硬编码--
            if (fullName.Contains("接触器")) return "接触器"; // --硬编码--
            if (fullName.Contains("热继电器")) return "热继"; // --硬编码--
            if (fullName.Contains("电度表") || fullName.Contains("电表")) return "电表"; // --硬编码--
            if (fullName.Contains("指示灯")) return "指示灯"; // --硬编码--
            if (fullName.Length > 6) return fullName.Substring(0, 5) + "..";
            return fullName;
        }

        /// <summary>
        /// 批量将用户确认的三箱型号安全回写至 Excel 工作表
        /// 遵循用户明确指示：直接回填到顶部汇总行 (Cab_Sum) 的 N 列 (第 14 列)
        /// </summary>
        /// <param name="items">待回写的箱柜条目列表</param>
        /// <param name="writeSummaryN">是否写入顶部汇总行 N 列 (用户指定目标列，默认 true)</param>
        /// <param name="writeSummaryD">是否写入顶部汇总行 D 列</param>
        /// <param name="writeDetailD">是否写入底部明细行 D 列</param>
        /// <returns>批量回写执行结果</returns>
        public static CabinetPipelineApplyResult BatchApplyCabinetModelsToExcel(
            List<CabinetPipelineApplyItem> items,
            bool writeSummaryN = true,
            bool writeSummaryD = false,
            bool writeDetailD = false)
        {
            var result = new CabinetPipelineApplyResult();

            if (items == null || items.Count == 0)
            {
                result.Success = false;
                result.Message = "没有需要回写的箱柜数据！";
                return result;
            }

            try
            {
                var context = Tool.GetActiveExcelContext();
                if (context == null)
                {
                    result.Success = false;
                    result.Message = "无法连接至 Excel 运行环境！";
                    return result;
                }

                dynamic app = context.App;
                dynamic wb = context.Wb;

                // 关闭屏幕刷新与事件广播，防止界面闪烁并大幅提升 COM 执行性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                int totalUpdated = 0;

                try
                {
                    // 按工作表分组批量写入
                    var groupedBySheet = items.GroupBy(i => i.SheetName);

                    foreach (var group in groupedBySheet)
                    {
                        string sheetName = group.Key;
                        Worksheet ws = null;
                        try { ws = wb.Worksheets[sheetName]; } catch { }
                        if (ws == null) continue;

                        // 检测汇总表表头行 (通常为第 6 行) 的 N 列，若未填表头文本，自动补齐为“箱柜型号”
                        if (writeSummaryN)
                        {
                            try
                            {
                                Range nHeaderCell = ws.Range["N6"];
                                string nHeader = nHeaderCell.Text?.ToString()?.Trim() ?? "";
                                if (string.IsNullOrWhiteSpace(nHeader))
                                {
                                    nHeaderCell.Value2 = "箱柜型号"; // --硬编码: 表头名--
                                }
                            }
                            catch { }
                        }

                        foreach (var cab in group)
                        {
                            string targetModel = cab.Model?.Trim() ?? string.Empty;
                            // 排除标记为待确认的项，避免写入无效内容
                            if (string.Equals(targetModel, "待确认", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            // 1. 核心回写：用户明确指示直接回填至 sum 行的 N 列 (第 14 列)
                            if (writeSummaryN && cab.SumRow > 0)
                            {
                                try
                                {
                                    Range sumNCell = ws.Range[$"N{cab.SumRow}"];
                                    sumNCell.Value2 = targetModel;
                                }
                                catch (Exception exWn)
                                {
                                    LogHelper.WriteLog($"[三箱管道] 回写汇总行 N 列异常 ({sheetName} Row {cab.SumRow}): {exWn.Message}");
                                }
                            }

                            // 2. 兼容回写：回写顶部汇总行 Cab_Sum 的 D 列 (第 4 列)
                            if (writeSummaryD && cab.SumRow > 0)
                            {
                                try
                                {
                                    Range sumCell = ws.Range[$"D{cab.SumRow}"];
                                    sumCell.Value2 = targetModel;
                                }
                                catch (Exception exW1)
                                {
                                    LogHelper.WriteLog($"[三箱管道] 回写汇总行 D 列异常 ({sheetName} Row {cab.SumRow}): {exW1.Message}");
                                }
                            }

                            // 3. 兼容回写：回写底部明细行 Cab_Det 的 D 列 (第 4 列: 型号值)
                            if (writeDetailD && cab.DetRow > 0)
                            {
                                try
                                {
                                    Range detCell = ws.Range[$"D{cab.DetRow}"];
                                    detCell.Value2 = targetModel;
                                }
                                catch (Exception exW2)
                                {
                                    LogHelper.WriteLog($"[三箱管道] 回写明细行 D 列异常 ({sheetName} Row {cab.DetRow}): {exW2.Message}");
                                }
                            }

                            totalUpdated++;
                        }
                    }

                    result.Success = true;
                    result.UpdatedCount = totalUpdated;
                    result.Message = $"成功将型号批量添写至 {totalUpdated} 台箱柜中！";
                }
                finally
                {
                    // 恢复 Excel 屏幕刷新与系统事件广播
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"批量回写发生异常: {ex.Message}";
                LogHelper.WriteLog($"[三箱型号管道] BatchApplyCabinetModelsToExcel 异常: {ex.Message}");
            }

            return result;
        }
    }
}
