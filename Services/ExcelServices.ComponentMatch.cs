using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using ExcelAddInDemo.Forms;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 公共服务类分部类: 基于名称、电流、极数、脱扣、品牌及必含约束反查物料库并批量回填与智能下拉联想
    /// </summary>
    public static partial class ExcelServices
    {
        // 本地保存物料匹配过滤配置的文件路径
        private static readonly string FilterConfigFilePath = Path.Combine(Tool.GetAppDataDirectory(), "component_match_filter_config.json");

        // 内存配置缓存
        private static ComponentMatchFilterConfig? _cachedFilterConfig;

        // 智能联想下拉悬浮窗全局静态单例
        private static ComponentMatchOverlayForm? _matchOverlayForm;

        // 物料匹配设置窗口静态单例引用 (可空)
        private static ComponentMatchForm? _matchSettingForm;

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“元器件物料匹配与品牌规则设置”窗口
        /// </summary>
        public static void ShowComponentMatchDialog()
        {
            try
            {
                // 以非模态方式展示物料匹配设置窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _matchSettingForm, () => new ComponentMatchForm());

                // 强制将设置窗口设为最前端并激活，确保不被置顶下拉悬浮窗遮蔽
                if (_matchSettingForm != null && !_matchSettingForm.IsDisposed)
                {
                    // 开启置顶
                    _matchSettingForm.TopMost = true;
                    // 置于前台
                    _matchSettingForm.BringToFront();
                    // 激活焦点
                    _matchSettingForm.Activate();
                }
            }
            catch (Exception ex)
            {
                // 记录打开弹窗异常日志
                LogHelper.WriteLog($"ShowComponentMatchDialog 异常: {ex.Message}");
            }
        }

        // 多工作表元器件行区间高速内存缓存 (零 COM 耗时，毫秒级瞬发)
        // Key: 工作表名称，Value: (缓存产生时间戳, 该表所有箱柜的 (起始行, 终止行) 区间列表)
        private static readonly Dictionary<string, (DateTime CacheTime, List<(int StartRow, int EndRow)> Ranges)> _categoryRangesSheetCache =
            new Dictionary<string, (DateTime, List<(int, int)>)>(StringComparer.OrdinalIgnoreCase);

        // 缓存有效期设为 10 分钟 (工作表结构稳定，切换或连续点击 100% 内存瞬发命中)
        private static readonly TimeSpan CategoryRangesCacheExpiry = TimeSpan.FromMinutes(10);

        /// <summary>
        /// 清空分类表元器件行区间缓存 (当工作表插入行、删除分类或新建箱柜时主动调用)
        /// </summary>
        /// <param name="sheetName">指定失效的工作表名称 (传空则清空全簿所有工作表缓存)</param>
        public static void InvalidateCategoryRowCache(string? sheetName = null)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                // 清空全簿所有工作表缓存
                _categoryRangesSheetCache.Clear();
                // 同步清空项目分类工作表白名单缓存
                Tool.InvalidateProjectCategorySheetCache();
            }
            else
            {
                // 仅移除指定工作表缓存
                _categoryRangesSheetCache.Remove(sheetName);
            }
        }

        /// <summary>
        /// 在后台或空闲时预热物料智能联想悬浮窗单例，使 WebView2 内核与前端页面提前就绪 (消除用户首次点击冷启动延迟)
        /// </summary>
        public static void PreloadComponentMatchOverlay()
        {
            try
            {
                // 仅在单例未创建或已释放时执行预热
                if (_matchOverlayForm == null || _matchOverlayForm.IsDisposed)
                {
                    // 实例化窗口单例
                    _matchOverlayForm = new Forms.ComponentMatchOverlayForm();
                    // 启动后台静默热备
                    _matchOverlayForm.WarmUp();
                }
            }
            catch (Exception ex)
            {
                // 记录预热异常日志
                LogHelper.WriteLog($"PreloadComponentMatchOverlay 预热异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 判定指定工作表的某一行是否处于分类明细表的有效元器件行区间 (严格遵循规则 6 架构)
        /// 规则 6: Cab_Det_k.Row + 2 为元器件起始行，Cab_Subsum_k.Row - 1 为元器件终止行
        /// 内置 10 分钟多工作表内存高速缓存与轻量名称短路，单表多次点击耗时 0 毫秒
        /// </summary>
        /// <param name="sheet">目标工作表 COM 句柄</param>
        /// <param name="row">待检测的目标行物理行号</param>
        /// <returns>若处于有效箱柜元器件行区间返回 true，否则返回 false</returns>
        public static bool IsCategoryComponentRow(dynamic sheet, int row)
        {
            // 校验工作表与行号有效性
            if (sheet == null || row <= 0) return false;

            try
            {
                // 获取当前工作表名称
                string sheetName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(sheetName)) return false;

                // 快速过滤系统隐藏表、字典表或选择表
                if (sheetName.StartsWith("_") || string.Equals(sheetName, "选择表", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                // 核心安全守门：优先白名单校验，若非白名单则做系统保留表安全过滤
                bool isCategory = Tool.IsProjectCategorySheet(sheet);
                if (!isCategory)
                {
                    // 若白名单未命中，进一步做保留表安全排除：排除系统保留表 (封面/项目信息/元件汇总表等)
                    if (Tool.IsReservedOrReportSheet(sheetName))
                    {
                        return false;
                    }
                    // 非系统保留表且未在白名单登记时，允许进入下方箱柜定义名称或结构嗅探，防止独立工作簿被误杀
                }

                var now = DateTime.UtcNow;

                // 1. 【高速多表内存缓存命中 (0ms)】：若在 10 分钟内且已有缓存记录，直接比对内存区间！
                if (_categoryRangesSheetCache.TryGetValue(sheetName, out var cachedEntry) &&
                    (now - cachedEntry.CacheTime) < CategoryRangesCacheExpiry)
                {
                    // 纯内存比较，0 次 COM 调用
                    foreach (var range in cachedEntry.Ranges)
                    {
                        if (row >= range.StartRow && row <= range.EndRow)
                        {
                            return true;
                        }
                    }
                    return false;
                }

                // 2. 缓存未命中：采用轻量级名称前缀短路扫描 (避免遍历无关名称导致上百次 COM 调用)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;
                var detDict = new Dictionary<int, int>();
                var subsumDict = new Dictionary<int, int>();

                // 内部辅助方法：快速解析名称集合中的 Det 和 Subsum
                // 参数 isSheetScope 标明是否为工作表级局部定义名称
                void QuickScanNames(dynamic? namesCollection, bool isSheetScope)
                {
                    if (namesCollection == null) return;
                    try
                    {
                        foreach (dynamic n in namesCollection)
                        {
                            try
                            {
                                // 提取纯文本名称字符串 (纯内存操作，不触发 COM Range 创建)
                                string rawName = Convert.ToString(n.Name) ?? string.Empty;
                                string clean = Tool.ExtractCleanNameStr(rawName);

                                // 关键短路：若既不含 Det 也不含 Subsum 前缀，立即跳过！绝不调用 RefersToRange！
                                bool isDet = clean.StartsWith(detPrefix, StringComparison.OrdinalIgnoreCase);
                                bool isSubsum = clean.StartsWith(subsumPrefix, StringComparison.OrdinalIgnoreCase);
                                if (!isDet && !isSubsum) continue;

                                // 提取箱柜序号 k
                                int k = Tool.ExtractIndexFromName(clean, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                                if (k <= 0) continue;

                                // 检查是否属于当前工作表:
                                // 工作簿级全局名称 (isSheetScope == false) 其 RefersTo 必须包含本工作表名
                                // 工作表级局部名称 (isSheetScope == true) 其 RefersTo 常常仅为 =$A$5，天然归属于本表
                                if (!isSheetScope)
                                {
                                    string refersTo = Convert.ToString(n.RefersTo) ?? string.Empty;
                                    if (!string.IsNullOrEmpty(refersTo) && !refersTo.Contains(sheetName))
                                    {
                                        // 公式明确属于其他工作表，跳过
                                        continue;
                                    }
                                }

                                // 仅针对命中的箱柜锚点安全读取行号
                                dynamic? refRange = null;
                                try { refRange = n.RefersToRange; } catch { }
                                if (refRange == null) continue;

                                int targetR = 0;
                                try { targetR = Convert.ToInt32(refRange.Row); } catch { continue; }
                                if (targetR <= 0) continue;

                                if (isDet && !detDict.ContainsKey(k)) detDict[k] = targetR;
                                if (isSubsum && !subsumDict.ContainsKey(k)) subsumDict[k] = targetR;
                            }
                            catch { }
                        }
                    }
                    catch { }
                }

                // 优先从工作表级定义名称集合快速扫描 (标记为工作表级局部作用域)
                try { QuickScanNames(sheet.Names, true); } catch { }

                // 再次从工作簿级定义名称集合快速扫描 (标记为工作簿级全局作用域)
                try
                {
                    dynamic? wb = sheet.Parent;
                    if (wb != null) QuickScanNames(wb.Names, false);
                }
                catch { }

                // 组装箱柜元器件行区间 (规则 6: Det + 2 到 Subsum - 1)
                var newRanges = new List<(int StartRow, int EndRow)>();
                foreach (var kvp in detDict)
                {
                    int k = kvp.Key;
                    int detRow = kvp.Value;
                    if (subsumDict.TryGetValue(k, out int subsumRow))
                    {
                        int compStartRow = detRow + 2;
                        int compEndRow = subsumRow - 1;
                        if (compEndRow >= compStartRow)
                        {
                            newRanges.Add((compStartRow, compEndRow));
                        }
                    }
                }

                // 3. 容错回退检查：若轻量快速名称扫描未发现有效箱柜，调用只读箱柜嗅探兜底
                if (newRanges.Count == 0)
                {
                    var validCabinets = Tool.GetSheetValidCabinets((object)sheet);
                    if (validCabinets != null && validCabinets.Count > 0)
                    {
                        foreach (var kvp in validCabinets)
                        {
                            var anc = kvp.Value;
                            if (anc?.Det == null || anc?.Subsum == null) continue;
                            int detR = 0, subR = 0;
                            try
                            {
                                detR = Convert.ToInt32(anc.Det.Row);
                                subR = Convert.ToInt32(anc.Subsum.Row);
                            }
                            catch { continue; }
                            int cStart = detR + 2, cEnd = subR - 1;
                            if (cEnd >= cStart) newRanges.Add((cStart, cEnd));
                        }
                    }
                }

                // 写入多工作表内存长效缓存 (即便是空列表也缓存，防止非箱柜表反复暴力扫描)
                _categoryRangesSheetCache[sheetName] = (now, newRanges);

                // 纯内存比对当前行
                foreach (var range in newRanges)
                {
                    if (row >= range.StartRow && row <= range.EndRow)
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录判定异常日志
                LogHelper.WriteLog($"[ComponentMatch] IsCategoryComponentRow 判定异常: {ex.Message}");
            }

            // 落在汇总区、信息行、计费区或小计总计行时返回 false
            return false;
        }

        /// <summary>
        /// 从本地持久化文件中加载物料匹配过滤配置
        /// </summary>
        public static ComponentMatchFilterConfig LoadComponentMatchFilterConfig()
        {
            try
            {
                if (_cachedFilterConfig != null) return _cachedFilterConfig;

                // 判断本地配置文件是否存在
                if (File.Exists(FilterConfigFilePath))
                {
                    string json = File.ReadAllText(FilterConfigFilePath);
                    var cfg = JsonSerializer.Deserialize<ComponentMatchFilterConfig>(json);
                    if (cfg != null)
                    {
                        _cachedFilterConfig = cfg;
                        return cfg;
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"加载物料匹配配置异常: {ex.Message}");
            }
            // 默认返回初始配置
            _cachedFilterConfig = ComponentMatchFilterConfig.CreateDefault();
            return _cachedFilterConfig;
        }

        /// <summary>
        /// 将物料匹配过滤配置保存至本地磁盘
        /// </summary>
        public static void SaveComponentMatchFilterConfig(ComponentMatchFilterConfig config)
        {
            try
            {
                _cachedFilterConfig = config;

                // 确保父目录存在
                string dir = Path.GetDirectoryName(FilterConfigFilePath)!;
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                // 序列化并写入文件
                string json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(FilterConfigFilePath, json);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"保存物料匹配配置异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 热重载物料匹配浮窗配置缓存
        /// </summary>
        /// <param name="config">最新的物料匹配配置对象</param>
        public static void ReloadComponentMatchOverlayConfig(ComponentMatchFilterConfig config)
        {
            // 更新服务层静态内存配置缓存
            _cachedFilterConfig = config;
            try
            {
                // 若浮窗已实例化且未释放，记录配置更新日志
                if (_matchOverlayForm != null && !_matchOverlayForm.IsDisposed)
                {
                    // 记录配置已成功热重载日志
                    LogHelper.WriteLog("物料匹配浮窗配置已成功热重载");
                }
            }
            catch (Exception ex)
            {
                // 记录热重载浮窗配置异常
                LogHelper.WriteLog($"ReloadComponentMatchOverlayConfig 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取当前生效的列映射配置（优先联动读取用户在“识别极数电流”中配置的输出列，保证两处统一联动）
        /// </summary>
        public static ComponentMatchColumnConfig GetEffectiveColumnConfig(ComponentMatchFilterConfig? filterConfig = null)
        {
            var activeFilterCfg = filterConfig ?? LoadComponentMatchFilterConfig();
            var colCfg = activeFilterCfg.ColumnConfig ?? new ComponentMatchColumnConfig();

            try
            {
                // 读取用户在“识别极数电流”界面保存的 ModelParserConfig
                var parserCtrl = new Controllers.ModelParamParserController();
                var parserCfg = parserCtrl.LoadConfig();
                if (parserCfg != null)
                {
                    // 若 ModelParserConfig 中配置了最小电流列，以其为准联动
                    if (!string.IsNullOrWhiteSpace(parserCfg.MinCurrentColumn))
                    {
                        colCfg.CurrentColumn = parserCfg.MinCurrentColumn.Trim().ToUpper();
                    }
                    // 若 ModelParserConfig 中配置了极数列，以其为准联动
                    if (!string.IsNullOrWhiteSpace(parserCfg.PoleColumn))
                    {
                        colCfg.PoleColumn = parserCfg.PoleColumn.Trim().ToUpper();
                    }
                    // 若 ModelParserConfig 中配置了脱扣列，以其为准联动
                    if (!string.IsNullOrWhiteSpace(parserCfg.TripModeColumn))
                    {
                        colCfg.TripModeColumn = parserCfg.TripModeColumn.Trim().ToUpper();
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"GetEffectiveColumnConfig 联动异常: {ex.Message}");
            }

            return colCfg;
        }

        /// <summary>
        /// 核心批处理方法: 直读选区已有的 B 列(名称)、T 列(额定电流)、U 列(极数)、V 列(脱扣方式)，应用品牌与必含字段约束反查 WebAPI 并批量回填
        /// </summary>
        /// <param name="filterConfig">多维匹配过滤配置 (包含品牌与必含字段规则，若为空则自动加载本地保存配置)</param>
        /// <returns>执行统计结果报告</returns>
        public static BatchMatchExecuteResult ExecuteBatchMatchWithDb(ComponentMatchFilterConfig? filterConfig = null)
        {
            // 初始化计时器
            var stopwatch = Stopwatch.StartNew();
            // 初始化返回结果
            var result = new BatchMatchExecuteResult();

            try
            {
                // 获取当前正在运行的 Excel 顶级 Application COM 句柄
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null)
                {
                    result.Success = false;
                    result.Message = "未检测到运行中的 Excel 应用程序实例";
                    return result;
                }

                // 获取当前活动工作簿与工作表
                dynamic? activeSheet = app.ActiveSheet;
                if (activeSheet == null)
                {
                    result.Success = false;
                    result.Message = "请先在 Excel 中打开或激活一个工作表";
                    return result;
                }

                // 获取当前工作表名称
                string sheetName = Convert.ToString(activeSheet.Name)?.Trim() ?? string.Empty;
                // 判断当前工作表是否为“元件汇总表”
                bool isSummarySheet = string.Equals(sheetName, ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase);

                // 规则 8 强约束: 在操作/提取分类明细表前，必须先调用 FixAndFillCabinetNamesForSheet 确保规则 6 定义名称与架构正确
                if (!isSummarySheet)
                {
                    Tool.FixAndFillCabinetNamesForSheet(activeSheet);
                }

                // 获取当前用户选区 Selection
                dynamic? selection = app.Selection;
                if (selection == null)
                {
                    result.Success = false;
                    result.Message = "请先在 Excel 中选择需要反查物料的数据行";
                    return result;
                }

                // 加载生效的过滤规则与列映射配置
                var activeFilterCfg = filterConfig ?? LoadComponentMatchFilterConfig();

                // 提取多选品牌限定与动态必含字段规则
                var selectedBrands = activeFilterCfg.GetEffectiveBrands();
                var mustContainRules = activeFilterCfg.MustContainRules ?? new List<MustContainRule>();

                // 双表自适应列映射路由:
                // 【元件汇总表】: B=名称, T=电流, U=极数, V=脱扣 | D=型号, I=品牌, L=表价, X=Param1, Y=Param2
                // 【分类明细表】: B=名称, W=电流, X=极数, Y=脱扣 | C=型号, D=品牌, M=表价, AA=Param1, AB=Param2
                string colName = "B";
                string colCur = isSummarySheet ? "T" : "W";
                string colPole = isSummarySheet ? "U" : "X";
                string colTrip = isSummarySheet ? "V" : "Y";

                string colModel = isSummarySheet ? "D" : "C";
                string colBrand = isSummarySheet ? "I" : "D";
                string colPrice = isSummarySheet ? "L" : "M";
                string colParam1 = isSummarySheet ? "X" : "AA";
                string colParam2 = isSummarySheet ? "Y" : "AB";

                // 初始化统计计数器
                int totalRows = 0;
                int uniqueCount = 0;
                int multipleCount = 0;
                int noneCount = 0;

                // 收集所有子选区的可撤销差量切片列表
                var undoSlices = new List<RangeDeltaSlice>();

                // 遍历当前选区的所有子区域 (支持连续区域及按住 Ctrl 的多选区)
                foreach (dynamic area in selection.Areas)
                {
                    // 获取当前区域的起始行与总行数
                    int startRow = (int)area.Row;
                    int rowCount = (int)area.Rows.Count;
                    int endRow = startRow + rowCount - 1;

                    // 若行数无效则跳过
                    if (rowCount <= 0) continue;

                    // ==================== 1. 一次性从输入列读入内存 ====================
                    // 一次性读入 B 列 (名称)
                    dynamic nameRange = activeSheet.Range[$"{colName}{startRow}:{colName}{endRow}"];
                    object[,] nameRawArray = ConvertTo2DArray(nameRange.Value2, rowCount);

                    // 一次性读入电流列 (汇总表 T 列 / 分类表 W 列)
                    dynamic curRange = activeSheet.Range[$"{colCur}{startRow}:{colCur}{endRow}"];
                    object[,] curRawArray = ConvertTo2DArray(curRange.Value2, rowCount);

                    // 一次性读入极数列 (汇总表 U 列 / 分类表 X 列)
                    dynamic poleRange = activeSheet.Range[$"{colPole}{startRow}:{colPole}{endRow}"];
                    object[,] poleRawArray = ConvertTo2DArray(poleRange.Value2, rowCount);

                    // 一次性读入脱扣方式列 (汇总表 V 列 / 分类表 Y 列)
                    dynamic tripRange = activeSheet.Range[$"{colTrip}{startRow}:{colTrip}{endRow}"];
                    object[,] tripRawArray = ConvertTo2DArray(tripRange.Value2, rowCount);

                    // ==================== 2. 读入输出列已有的原值作为安全底稿 (确保跳过的非元器件行与空行原样保留) ====================
                    dynamic origModelRange = activeSheet.Range[$"{colModel}{startRow}:{colModel}{endRow}"];
                    object[,] origModelRaw = ConvertTo2DArray(origModelRange.Value2, rowCount);

                    dynamic origBrandRange = activeSheet.Range[$"{colBrand}{startRow}:{colBrand}{endRow}"];
                    object[,] origBrandRaw = ConvertTo2DArray(origBrandRange.Value2, rowCount);

                    dynamic origPriceRange = activeSheet.Range[$"{colPrice}{startRow}:{colPrice}{endRow}"];
                    object[,] origPriceRaw = ConvertTo2DArray(origPriceRange.Value2, rowCount);

                    dynamic origParam1Range = activeSheet.Range[$"{colParam1}{startRow}:{colParam1}{endRow}"];
                    object[,] origParam1Raw = ConvertTo2DArray(origParam1Range.Value2, rowCount);

                    dynamic origParam2Range = activeSheet.Range[$"{colParam2}{startRow}:{colParam2}{endRow}"];
                    object[,] origParam2Raw = ConvertTo2DArray(origParam2Range.Value2, rowCount);

                    // 在内存中分配 6 个回填目标列的二维数组，初始拷贝底稿原值
                    object[,] nameArray = new object[rowCount, 1];
                    object[,] modelArray = new object[rowCount, 1];
                    object[,] brandArray = new object[rowCount, 1];
                    object[,] priceArray = new object[rowCount, 1];
                    object[,] param1Array = new object[rowCount, 1];
                    object[,] param2Array = new object[rowCount, 1];

                    for (int k = 0; k < rowCount; k++)
                    {
                        nameArray[k, 0] = nameRawArray[k + 1, 1];
                        modelArray[k, 0] = origModelRaw[k + 1, 1];
                        brandArray[k, 0] = origBrandRaw[k + 1, 1];
                        priceArray[k, 0] = origPriceRaw[k + 1, 1];
                        param1Array[k, 0] = origParam1Raw[k + 1, 1];
                        param2Array[k, 0] = origParam2Raw[k + 1, 1];
                    }

                    // 汇总表 M 列折扣底稿
                    object[,] discountMArray = null;
                    object[,] origDiscountMArray = null;
                    if (isSummarySheet)
                    {
                        dynamic discountMRange = activeSheet.Range[$"M{startRow}:M{endRow}"];
                        object[,] discRaw = ConvertTo2DArray(discountMRange.Value2, rowCount);
                        discountMArray = new object[rowCount, 1];
                        origDiscountMArray = new object[rowCount, 1];
                        for (int k = 0; k < rowCount; k++)
                        {
                            discountMArray[k, 0] = discRaw[k + 1, 1];
                            origDiscountMArray[k, 0] = discRaw[k + 1, 1];
                        }
                    }

                    // 记录型号列单元格的原始底色 (用于撤销时无损精准还原)
                    int[,] origModelColorIndexes = new int[rowCount, 1];
                    int[,] origModelColors = new int[rowCount, 1];
                    for (int k = 0; k < rowCount; k++)
                    {
                        dynamic mCell = activeSheet.Range[$"{colModel}{startRow + k}"];
                        origModelColorIndexes[k, 0] = (int)mCell.Interior.ColorIndex;
                        origModelColors[k, 0] = (int)mCell.Interior.Color;
                    }

                    // 收集当前区域中需要高亮淡黄底色的行号集合
                    var yellowHighlightRowList = new List<int>();
                    // 收集不需要高亮/需清除底色的行号集合
                    var clearHighlightRowList = new List<int>();

                    // ==================== 3. 内存逐行匹配 ====================
                    for (int i = 0; i < rowCount; i++)
                    {
                        // 计算当前内存行在 Excel 中的绝对物理行号
                        int currentRealRow = startRow + i;

                        // 门控 1: 在分类明细表中，严格校验是否为有效元器件行 (规则 6: Cab_Det+2 至 Cab_Subsum-1)
                        if (!isSummarySheet && !IsCategoryComponentRow(activeSheet, currentRealRow))
                        {
                            // 落在汇总区、信息行、计费区或小计总计行时，严格保留底稿原值不修改
                            continue;
                        }

                        // 直接获取当前行已存在的名称、电流、极数、脱扣方式内容
                        string rawName = nameRawArray[i + 1, 1]?.ToString()?.Trim() ?? string.Empty;
                        string minCur = curRawArray[i + 1, 1]?.ToString()?.Trim() ?? string.Empty;
                        string pole = poleRawArray[i + 1, 1]?.ToString()?.Trim() ?? string.Empty;
                        string tripMode = tripRawArray[i + 1, 1]?.ToString()?.Trim() ?? string.Empty;

                        // 若该行名称、电流与极数均为空，视为空白元器件行，清除底色并保持原样
                        if (string.IsNullOrWhiteSpace(rawName) && string.IsNullOrWhiteSpace(minCur) && string.IsNullOrWhiteSpace(pole))
                        {
                            clearHighlightRowList.Add(currentRealRow);
                            continue;
                        }

                        totalRows++;

                        // 提取该行已有品牌 (优先锁定单元格已有品牌，若为空则采用用户偏好设置)
                        string existingBrand = origBrandRaw[i + 1, 1]?.ToString()?.Trim() ?? string.Empty;
                        var rowBrands = new List<string>();
                        if (!string.IsNullOrEmpty(existingBrand))
                        {
                            rowBrands.Add(existingBrand);
                        }
                        else
                        {
                            rowBrands.AddRange(selectedBrands);
                        }

                        // 调用 WebAPI 客户端或本地 SQLite 个人物料库反查真实数据库 (支持多选品牌)
                        var matchedItems = string.Equals(activeFilterCfg.DataSource, "personal", StringComparison.OrdinalIgnoreCase)
                            ? Services.PersonalComponentDbService.SearchComponents(null, rawName, minCur, pole, tripMode, rowBrands, mustContainRules)
                            : ComponentApiClient.QueryComponents(
                                rawName,
                                minCur,
                                pole,
                                tripMode,
                                rowBrands,
                                mustContainRules
                            );

                        // =================================================================
                        // 分支 1: 查询到唯一值 (Count == 1) -> 完整自动回填各字段
                        // =================================================================
                        if (matchedItems.Count == 1)
                        {
                            var item = matchedItems[0];
                            // B 列 (标准名称)
                            nameArray[i, 0] = !string.IsNullOrEmpty(item.Name) ? item.Name : rawName;
                            // 型号列 (汇总表 D 列 / 分类明细表 C 列)
                            modelArray[i, 0] = item.Model ?? string.Empty;
                            // 品牌/生产厂家列 (汇总表 I 列 / 分类明细表 D 列)
                            brandArray[i, 0] = item.Brand ?? string.Empty;
                            // 表价列 (汇总表 L 列本体表价 / 分类明细表 M 列基准表价)
                            priceArray[i, 0] = item.Price > 0 ? (object)(double)item.Price : string.Empty;
                            // 扩展参数1列 (汇总表 X 列 / 分类明细表 AA 列)
                            param1Array[i, 0] = item.Param1 ?? string.Empty;
                            // 扩展参数2列 (汇总表 Y 列 / 分类明细表 AB 列)
                            param2Array[i, 0] = item.Param2 ?? string.Empty;

                            // 汇总表特殊维护: 若 M 列折扣为空或0，补为 1
                            if (isSummarySheet && discountMArray != null)
                            {
                                string curM = discountMArray[i, 0]?.ToString()?.Trim() ?? string.Empty;
                                if (string.IsNullOrEmpty(curM) || curM == "0")
                                {
                                    discountMArray[i, 0] = 1;
                                }
                            }

                            clearHighlightRowList.Add(currentRealRow);
                            uniqueCount++;
                        }
                        // =================================================================
                        // 分支 2: 查出来多个匹配 (Count > 1) -> 保留原有名称，型号列填入“点击查询(条数)”并设淡黄底色
                        // =================================================================
                        else if (matchedItems.Count > 1)
                        {
                            // 型号列填入带匹配条数的提示文本 (如: 点击查询(7)) --硬编码--
                            modelArray[i, 0] = $"{ComponentMatchDefaults.MultipleCandidatesText}({matchedItems.Count})";
                            yellowHighlightRowList.Add(currentRealRow);
                            multipleCount++;
                        }
                        // =================================================================
                        // 分支 3: 没找到任何匹配 (Count == 0) -> 保留原样并清除底色
                        // =================================================================
                        else
                        {
                            clearHighlightRowList.Add(currentRealRow);
                            noneCount++;
                        }
                    }

                    // ==================== 4. 一次性将二维数组整块写回 Excel 目标字段列 (绝不覆盖电流/极数/脱扣输入列) ====================
                    // 写回 B 列 (名称)
                    activeSheet.Range[$"{colName}{startRow}:{colName}{endRow}"].Value2 = nameArray;
                    // 写回型号列 (汇总表 D 列 / 分类明细表 C 列)
                    activeSheet.Range[$"{colModel}{startRow}:{colModel}{endRow}"].Value2 = modelArray;
                    // 写回品牌列 (汇总表 I 列 / 分类明细表 D 列)
                    activeSheet.Range[$"{colBrand}{startRow}:{colBrand}{endRow}"].Value2 = brandArray;
                    // 写回表价列 (汇总表 L 列 / 分类明细表 M 列)
                    activeSheet.Range[$"{colPrice}{startRow}:{colPrice}{endRow}"].Value2 = priceArray;
                    // 写回参数1列 (汇总表 X 列 / 分类明细表 AA 列)
                    activeSheet.Range[$"{colParam1}{startRow}:{colParam1}{endRow}"].Value2 = param1Array;
                    // 写回参数2列 (汇总表 Y 列 / 分类明细表 AB 列)
                    activeSheet.Range[$"{colParam2}{startRow}:{colParam2}{endRow}"].Value2 = param2Array;

                    // 汇总表写回 M 列折扣
                    if (isSummarySheet && discountMArray != null)
                    {
                        activeSheet.Range[$"M{startRow}:M{endRow}"].Value2 = discountMArray;
                    }

                    // ==================== 5. 针对“点击查询”单元格统一应用淡黄底色 ====================
                    // 统一设置多条匹配行的型号列单元格背景颜色为淡黄色
                    foreach (int r in yellowHighlightRowList)
                    {
                        dynamic targetCell = activeSheet.Range[$"{colModel}{r}"];
                        // 设置淡黄底色 (RGB: 255, 242, 204)
                        targetCell.Interior.Color = ComponentMatchDefaults.LightYellowOleColor;
                    }

                    // 统一清除唯一匹配或无匹配行的型号列底色
                    foreach (int r in clearHighlightRowList)
                    {
                        dynamic targetCell = activeSheet.Range[$"{colModel}{r}"];
                        // 仅当已有底色时重置为无填充
                        if ((int)targetCell.Interior.ColorIndex != ComponentMatchDefaults.XlNoneColorIndex)
                        {
                            targetCell.Interior.ColorIndex = ComponentMatchDefaults.XlNoneColorIndex;
                        }
                    }

                    // 记录型号列最新底色 (用于重做时准确恢复淡黄等样式)
                    int[,] newModelColorIndexes = new int[rowCount, 1];
                    int[,] newModelColors = new int[rowCount, 1];
                    for (int k = 0; k < rowCount; k++)
                    {
                        dynamic mCell = activeSheet.Range[$"{colModel}{startRow + k}"];
                        newModelColorIndexes[k, 0] = (int)mCell.Interior.ColorIndex;
                        newModelColors[k, 0] = (int)mCell.Interior.Color;
                    }

                    // 构造原值二维切片数组
                    object[,] origNameArr = new object[rowCount, 1];
                    object[,] origModelArr = new object[rowCount, 1];
                    object[,] origBrandArr = new object[rowCount, 1];
                    object[,] origPriceArr = new object[rowCount, 1];
                    object[,] origParam1Arr = new object[rowCount, 1];
                    object[,] origParam2Arr = new object[rowCount, 1];
                    for (int k = 0; k < rowCount; k++)
                    {
                        origNameArr[k, 0] = nameRawArray[k + 1, 1];
                        origModelArr[k, 0] = origModelRaw[k + 1, 1];
                        origBrandArr[k, 0] = origBrandRaw[k + 1, 1];
                        origPriceArr[k, 0] = origPriceRaw[k + 1, 1];
                        origParam1Arr[k, 0] = origParam1Raw[k + 1, 1];
                        origParam2Arr[k, 0] = origParam2Raw[k + 1, 1];
                    }

                    // 登记各输出列差量切片至撤销集合
                    undoSlices.Add(new RangeDeltaSlice { SheetName = sheetName, RangeAddress = $"{colName}{startRow}:{colName}{endRow}", OldValues = origNameArr, NewValues = nameArray });
                    undoSlices.Add(new RangeDeltaSlice { SheetName = sheetName, RangeAddress = $"{colModel}{startRow}:{colModel}{endRow}", OldValues = origModelArr, NewValues = modelArray, OldColorIndexes = origModelColorIndexes, NewColorIndexes = newModelColorIndexes, OldColors = origModelColors, NewColors = newModelColors });
                    undoSlices.Add(new RangeDeltaSlice { SheetName = sheetName, RangeAddress = $"{colBrand}{startRow}:{colBrand}{endRow}", OldValues = origBrandArr, NewValues = brandArray });
                    undoSlices.Add(new RangeDeltaSlice { SheetName = sheetName, RangeAddress = $"{colPrice}{startRow}:{colPrice}{endRow}", OldValues = origPriceArr, NewValues = priceArray });
                    undoSlices.Add(new RangeDeltaSlice { SheetName = sheetName, RangeAddress = $"{colParam1}{startRow}:{colParam1}{endRow}", OldValues = origParam1Arr, NewValues = param1Array });
                    undoSlices.Add(new RangeDeltaSlice { SheetName = sheetName, RangeAddress = $"{colParam2}{startRow}:{colParam2}{endRow}", OldValues = origParam2Arr, NewValues = param2Array });

                    // 汇总表特殊维护：登记 M 列折扣切片
                    if (isSummarySheet && discountMArray != null && origDiscountMArray != null)
                    {
                        undoSlices.Add(new RangeDeltaSlice { SheetName = sheetName, RangeAddress = $"M{startRow}:M{endRow}", OldValues = origDiscountMArray, NewValues = discountMArray });
                    }
                }

                // 若有有效修改切片且实际处理了数据行，打包生成撤销命令并入栈
                if (undoSlices.Count > 0 && totalRows > 0)
                {
                    // 创建批量物料匹配的差量撤销命令
                    var batchCmd = new RangeDeltaCommand($"批量匹配物料 ({totalRows}行)", undoSlices);
                    // 压入全局撤销重做中心
                    UndoRedoManager.Instance.PushCommand(batchCmd);
                }

                // 停止计时并汇总执行结果
                stopwatch.Stop();
                result.Success = true;
                result.TotalRows = totalRows;
                result.UniqueMatchCount = uniqueCount;
                result.MultipleMatchCount = multipleCount;
                result.NoneMatchCount = noneCount;
                result.ElapsedMilliseconds = stopwatch.ElapsedMilliseconds;
                result.Message = $"批量处理完成：共处理 {totalRows} 行，唯一精确匹配 {uniqueCount} 行，多条待选 {multipleCount} 行(已高亮淡黄)，未匹配 {noneCount} 行，耗时 {stopwatch.ElapsedMilliseconds} ms";
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                result.Success = false;
                result.Message = $"反查匹配异常: {ex.Message}";
            }

            // 返回最终执行统计结果
            return result;
        }

        /// <summary>
        /// 兼容旧版仅传入列配置的重载方法
        /// </summary>
        public static BatchMatchExecuteResult ExecuteBatchMatchWithColumnConfig(ComponentMatchColumnConfig? colConfig)
        {
            var filterConfig = LoadComponentMatchFilterConfig();
            if (colConfig != null) filterConfig.ColumnConfig = colConfig;
            return ExecuteBatchMatchWithDb(filterConfig);
        }

        /// <summary>
        /// 在活动单元格 (D 列) 位置智能激活并贴合弹出物料联想下拉框 (类似 SmartInput 交互)
        /// </summary>
        /// <param name="activeCell">当前选中的活动单元格 COM 句柄</param>
        /// <param name="isCategoryRowValidated">是否已经前置校验过属于有效元器件行 (默认 false，为 true 时跳过二次判定)</param>
        public static void ShowComponentMatchOverlay(dynamic activeCell, bool isCategoryRowValidated = false)
        {
            if (activeCell == null) return;

            // 核心门控: 若当前下拉悬浮窗已处于显示且被用户“固定置顶”，绝不重新搜索覆盖现有结果
            if (_matchOverlayForm != null && !_matchOverlayForm.IsDisposed && _matchOverlayForm.Visible && _matchOverlayForm.IsPinned)
            {
                // 仅更新目标活动单元格引用，确保后续直接点击候选条目时精准回填至最新行
                _matchOverlayForm.ShowAtCell(activeCell, null, new CellParamsContext(), null!);
                return;
            }

            try
            {
                // 1. 获取当前活动单元格的行号、列号与所属工作表
                int col = 0;
                try { col = Convert.ToInt32(activeCell.Column); } catch { }
                int row = 0;
                try { row = Convert.ToInt32(activeCell.Row); } catch { }
                if (row <= 0 || col <= 0) return;

                // 获取所属工作表对象
                dynamic sheet = activeCell.Worksheet;
                if (sheet == null) return;

                // 判断是否为“元件汇总表”
                string sheetName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;
                bool isSummarySheet = string.Equals(sheetName, ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase);

                // 2. 校验触发列与工作表类型匹配:
                //    - 元件汇总表: 必须在 D 列 (第 4 列: 规格型号)
                //    - 常规分类明细表: 必须在 C 列 (第 3 列: 规格型号)，且必须处于有效箱柜元器件行区间 (规则 6)
                if (isSummarySheet)
                {
                    // 汇总表非 D 列不弹出
                    if (col != 4)
                    {
                        HideComponentMatchOverlay();
                        return;
                    }
                }
                else
                {
                    // 分类明细表非 C 列不弹出
                    if (col != 3)
                    {
                        HideComponentMatchOverlay();
                        return;
                    }

                    // 严格门控：校验是否处于当前表箱柜元器件插槽行 (Cab_Det+2 至 Cab_Subsum-1)
                    // 若调用方已前置校验通过，直接跳过；否则校验
                    if (!isCategoryRowValidated && !IsCategoryComponentRow(sheet, row))
                    {
                        // 落在汇总区、信息行或计费区域时隐藏浮窗
                        HideComponentMatchOverlay();
                        return;
                    }
                }

                // 3. 加载当前生效的全局过滤管道配置
                var filterConfig = LoadComponentMatchFilterConfig();

                // 核心门控: 若用户明确关闭了搜索浮窗，且当前单元格内容不含“点击查询”，才静默隐藏
                // 若单元格内容包含“点击查询”，始终允许弹出浮窗供用户选择候选物料
                string activeCellVal = Convert.ToString(activeCell.Value2)?.Trim() ?? string.Empty;
                bool isMultipleCandidates = activeCellVal.Contains(ComponentMatchDefaults.MultipleCandidatesText);
                if (!filterConfig.EnableSearchOverlay && !isMultipleCandidates)
                {
                    HideComponentMatchOverlay();
                    return;
                }

                // 4. 提取当前行参数与上下文
                string rawName = string.Empty;
                string rawCur = string.Empty;
                string rawPole = string.Empty;
                string rawTrip = string.Empty;
                string rawModel = string.Empty;
                string rawPriceFormula = string.Empty;
                string rawParam1 = string.Empty;
                string rawParam2 = string.Empty;
                var brandsToUse = new List<string>();

                if (isSummarySheet)
                {
                    // 【元件汇总表】一次性向量读取 B~Y 列 (第 2 列至第 25 列，仅 1 次 COM 往返)
                    dynamic summaryRowRange = sheet.Range[$"B{row}:Y{row}"];
                    object[,] sData = ConvertTo2DArray(summaryRowRange.Value2, 1);

                    // B 列 (偏移 1) = 名称
                    rawName = sData[1, 1]?.ToString()?.Trim() ?? string.Empty;
                    // D 列 (偏移 3: 4-2+1=3) = 规格型号
                    rawModel = sData[1, 3]?.ToString()?.Trim() ?? string.Empty;
                    // C 列 (偏移 2: 3-2+1=2) = 原型号参考
                    string refModel = sData[1, 2]?.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(rawModel) || string.Equals(rawModel, ComponentMatchDefaults.MultipleCandidatesText, StringComparison.OrdinalIgnoreCase))
                    {
                        rawModel = refModel;
                    }

                    // I 列 (偏移 8: 9-2+1=8) = 品牌/厂家
                    string brandVal = sData[1, 8]?.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(brandVal))
                    {
                        brandsToUse.Add(brandVal);
                    }
                    else
                    {
                        brandsToUse.AddRange(filterConfig.GetEffectiveBrands());
                    }

                    // T/U/V 列 (偏移 19, 20, 21: 20-2+1=19) = 电流、极数、脱扣
                    rawCur = sData[1, 19]?.ToString()?.Trim() ?? string.Empty;
                    rawPole = sData[1, 20]?.ToString()?.Trim() ?? string.Empty;
                    rawTrip = sData[1, 21]?.ToString()?.Trim() ?? string.Empty;

                    // X/Y 列 (偏移 23, 24: 24-2+1=23) = Param1, Param2
                    rawParam1 = sData[1, 23]?.ToString()?.Trim() ?? string.Empty;
                    rawParam2 = sData[1, 24]?.ToString()?.Trim() ?? string.Empty;

                    // L 列 (偏移 11: 12-2+1=11) = 本体表价
                    rawPriceFormula = sData[1, 11]?.ToString()?.Trim() ?? string.Empty;
                }
                else
                {
                    // 【分类明细表】一次性向量读取 B~AB 列 (第 2 列至第 28 列，仅 1 次 COM 往返，提速 90%！)
                    dynamic catRowRange = sheet.Range[$"B{row}:AB{row}"];
                    object[,] cData = ConvertTo2DArray(catRowRange.Value2, 1);

                    // B 列 (偏移 1) = 名称
                    rawName = cData[1, 1]?.ToString()?.Trim() ?? string.Empty;
                    // C 列 (偏移 2) = 规格型号
                    rawModel = cData[1, 2]?.ToString()?.Trim() ?? string.Empty;

                    // D 列 (偏移 3) = 品牌/厂家
                    string brandVal = cData[1, 3]?.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(brandVal))
                    {
                        brandsToUse.Add(brandVal);
                    }
                    else
                    {
                        brandsToUse.AddRange(filterConfig.GetEffectiveBrands());
                    }

                    // M 列 (偏移 12: 13-2+1=12) = 基准表价
                    rawPriceFormula = cData[1, 12]?.ToString()?.Trim() ?? string.Empty;

                    // W/X/Y 列 (偏移 22, 23, 24: 23-2+1=22) = 电流、极数、脱扣
                    rawCur = cData[1, 22]?.ToString()?.Trim() ?? string.Empty;
                    rawPole = cData[1, 23]?.ToString()?.Trim() ?? string.Empty;
                    rawTrip = cData[1, 24]?.ToString()?.Trim() ?? string.Empty;

                    // AA/AB 列 (偏移 26, 27: 27-2+1=26) = Param1, Param2
                    rawParam1 = cData[1, 26]?.ToString()?.Trim() ?? string.Empty;
                    rawParam2 = cData[1, 27]?.ToString()?.Trim() ?? string.Empty;
                }

                // 构造上下文参数 (带上 Param1、Param2 与工作表标识)
                var cellParams = new CellParamsContext
                {
                    Name = rawName,
                    Current = rawCur,
                    Pole = rawPole,
                    TripMode = rawTrip,
                    CurrentModel = rawModel,
                    CurrentPrice = rawPriceFormula,
                    Param1 = rawParam1,
                    Param2 = rawParam2,
                    IsCategorySheet = !isSummarySheet,
                    Brands = brandsToUse
                };

                // 5. 初始化或复用下拉悬浮窗实例
                if (_matchOverlayForm == null || _matchOverlayForm.IsDisposed)
                {
                    // 创建悬浮窗单例
                    _matchOverlayForm = new ComponentMatchOverlayForm();
                }

                // 6. 在当前单元格下方贴合弹出 (传 null 触发后台异步非阻塞拉取数据)
                _matchOverlayForm.ShowAtCell(activeCell, null, cellParams, filterConfig);
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"ShowComponentMatchOverlay 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 在活动单元格 (或当前行 D 列) 位置一键弹出配套附件选配浮窗 (供右键菜单【选配配套附件...】调用)
        /// </summary>
        /// <param name="targetCell">目标单元格 COM 句柄 (为空则自动采用 Application.ActiveCell)</param>
        public static void ShowComponentAttachmentOverlay(dynamic? targetCell = null)
        {
            try
            {
                // 获取 Excel 应用程序实例
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null) return;

                // 获取活动单元格
                dynamic activeCell = targetCell ?? app.ActiveCell;
                if (activeCell == null) return;

                // 获取当前行号与所属工作表
                int row = 0;
                try { row = Convert.ToInt32(activeCell.Row); } catch { }
                if (row <= 0) return;

                dynamic sheet = activeCell.Worksheet;
                if (sheet == null) return;

                // 判断是否为“元件汇总表”
                string sheetName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;
                bool isSummarySheet = string.Equals(sheetName, ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase);

                // 获取当前活动行型号所在单元格 (汇总表对齐 D 列，分类明细表对齐 C 列)
                string modelCol = isSummarySheet ? "D" : "C";
                dynamic targetModelCell = sheet.Range[$"{modelCol}{row}"];

                // 加载当前生效的全局过滤管道配置
                var filterConfig = LoadComponentMatchFilterConfig();

                string rawName = string.Empty;
                string rawCur = string.Empty;
                string rawPole = string.Empty;
                string rawTrip = string.Empty;
                string rawModel = string.Empty;
                string rawPriceFormula = string.Empty;
                string rawParam1 = string.Empty;
                string rawParam2 = string.Empty;
                var brandsToUse = new List<string>();

                if (isSummarySheet)
                {
                    // 汇总表提取
                    rawName = Convert.ToString(sheet.Range[$"B{row}"].Value2)?.Trim() ?? string.Empty;
                    rawModel = Convert.ToString(sheet.Range[$"D{row}"].Value2)?.Trim() ?? string.Empty;
                    string refModel = Convert.ToString(sheet.Range[$"C{row}"].Value2)?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(rawModel) || string.Equals(rawModel, ComponentMatchDefaults.MultipleCandidatesText, StringComparison.OrdinalIgnoreCase))
                    {
                        rawModel = refModel;
                    }

                    string brandVal = Convert.ToString(sheet.Range[$"I{row}"].Value2)?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(brandVal))
                    {
                        brandsToUse.Add(brandVal);
                    }
                    else
                    {
                        brandsToUse.AddRange(filterConfig.GetEffectiveBrands());
                    }

                    rawCur = Convert.ToString(sheet.Range[$"T{row}"].Value2)?.Trim() ?? string.Empty;
                    rawPole = Convert.ToString(sheet.Range[$"U{row}"].Value2)?.Trim() ?? string.Empty;
                    rawTrip = Convert.ToString(sheet.Range[$"V{row}"].Value2)?.Trim() ?? string.Empty;
                    rawParam1 = Convert.ToString(sheet.Range[$"X{row}"].Value2)?.Trim() ?? string.Empty;
                    rawParam2 = Convert.ToString(sheet.Range[$"Y{row}"].Value2)?.Trim() ?? string.Empty;

                    rawPriceFormula = Convert.ToString(sheet.Range[$"L{row}"].Formula)?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(rawPriceFormula))
                    {
                        rawPriceFormula = Convert.ToString(sheet.Range[$"L{row}"].Value2)?.Trim() ?? string.Empty;
                    }
                }
                else
                {
                    // 分类明细表提取
                    rawName = Convert.ToString(sheet.Range[$"B{row}"].Value2)?.Trim() ?? string.Empty;
                    rawModel = Convert.ToString(sheet.Range[$"C{row}"].Value2)?.Trim() ?? string.Empty;

                    string brandVal = Convert.ToString(sheet.Range[$"D{row}"].Value2)?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(brandVal))
                    {
                        brandsToUse.Add(brandVal);
                    }
                    else
                    {
                        brandsToUse.AddRange(filterConfig.GetEffectiveBrands());
                    }

                    rawCur = Convert.ToString(sheet.Range[$"W{row}"].Value2)?.Trim() ?? string.Empty;
                    rawPole = Convert.ToString(sheet.Range[$"X{row}"].Value2)?.Trim() ?? string.Empty;
                    rawTrip = Convert.ToString(sheet.Range[$"Y{row}"].Value2)?.Trim() ?? string.Empty;
                    rawParam1 = Convert.ToString(sheet.Range[$"AA{row}"].Value2)?.Trim() ?? string.Empty;
                    rawParam2 = Convert.ToString(sheet.Range[$"AB{row}"].Value2)?.Trim() ?? string.Empty;

                    rawPriceFormula = Convert.ToString(sheet.Range[$"M{row}"].Formula)?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(rawPriceFormula))
                    {
                        rawPriceFormula = Convert.ToString(sheet.Range[$"M{row}"].Value2)?.Trim() ?? string.Empty;
                    }
                }

                // 构造上下文参数
                var cellParams = new CellParamsContext
                {
                    Name = rawName,
                    Current = rawCur,
                    Pole = rawPole,
                    TripMode = rawTrip,
                    CurrentModel = rawModel,
                    CurrentPrice = rawPriceFormula,
                    Param1 = rawParam1,
                    Param2 = rawParam2,
                    IsCategorySheet = !isSummarySheet,
                    Brands = brandsToUse
                };

                // 若当前型号为空，友好提示用户先填写或选择型号
                if (string.IsNullOrWhiteSpace(rawModel))
                {
                    System.Windows.Forms.MessageBox.Show(
                        "当前行尚未填写或匹配元器件规格型号，请先输入或选择物料型号后再配置配套附件！",
                        "提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information
                    );
                    // 顺畅弹起常规物料选择框引导用户先选型号
                    ShowComponentMatchOverlay(targetModelCell);
                    return;
                }

                // 初始化或复用下拉悬浮窗实例
                if (_matchOverlayForm == null || _matchOverlayForm.IsDisposed)
                {
                    _matchOverlayForm = new ComponentMatchOverlayForm();
                }

                // 直接进入配套附件选配模式展示 (贴合在目标型号所在单元格)
                _matchOverlayForm.ShowAttachmentsAtCell(targetModelCell, cellParams, filterConfig);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ShowComponentAttachmentOverlay 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 隐藏物料联想下拉悬浮框
        /// </summary>
        public static void HideComponentMatchOverlay()
        {
            try
            {
                if (_matchOverlayForm != null && !_matchOverlayForm.IsDisposed && _matchOverlayForm.Visible)
                {
                    // 若已被用户“固定置顶”在前端，绝不自动隐藏
                    if (_matchOverlayForm.IsPinned)
                    {
                        return;
                    }
                    _matchOverlayForm.Hide();
                }
            }
            catch { }
        }

        /// <summary>
        /// 将选中的标准物料项自动回填至当前活动行
        /// 在“元件汇总表”中：汇总表自身已内嵌所有公式，仅将本体价格填在 L 列 (本体表价)，M 列折扣补 1，绝不重写公式
        /// </summary>
        /// <param name="item">用户选中的标准元器件物料 DTO</param>
        /// <param name="targetCell">目标单元格 COM 句柄 (为空则使用当前活动单元格)</param>
        public static bool FillSelectedComponentToActiveRow(ComponentApiDto item, dynamic? targetCell = null)
        {
            if (item == null) return false;

            try
            {
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null) return false;

                // 确定目标单元格
                dynamic cell = targetCell ?? app.ActiveCell;
                if (cell == null) return false;

                dynamic sheet = cell.Worksheet;
                if (sheet == null) return false;

                int row = Convert.ToInt32(cell.Row);
                if (row <= 0) return false;

                // 判断当前是否在“元件汇总表”中
                string sheetName = Convert.ToString(sheet.Name) ?? string.Empty;
                bool isSummarySheet = string.Equals(sheetName.Trim(), ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase);

                // 定义当前操作将涉及的所有列字段及是否捕获底色的映射配置
                var targetColumns = isSummarySheet
                    ? new (string Col, bool HasColor)[]
                    {
                        ("B", false), ("D", true), ("I", false), ("L", false), ("M", false),
                        ("P", false), ("T", false), ("U", false), ("V", false), ("W", false),
                        ("X", false), ("Y", false)
                    }
                    : new (string Col, bool HasColor)[]
                    {
                        ("B", false), ("C", true), ("D", true), ("M", false), ("W", false),
                        ("X", false), ("Y", false), ("Z", false), ("AA", false), ("AB", false)
                    };

                // 在修改前集中捕获涉及字段的原值与底色切片
                var singleRowSlices = new List<RangeDeltaSlice>();
                foreach (var itemCol in targetColumns)
                {
                    singleRowSlices.Add(CaptureSingleCellSlice(sheet, sheetName, itemCol.Col, row, itemCol.HasColor));
                }

                // 1. 回填 B 列 (标准名称)
                if (!string.IsNullOrEmpty(item.Name))
                {
                    // 将标准名称写入 B 列
                    sheet.Range[$"B{row}"].Value2 = item.Name;
                }

                // 2. 根据工作表类型执行差异化回填
                if (isSummarySheet)
                {
                    // 在“元件汇总表”中：
                    // D 列 (第 4 列) = 标准型号 (覆盖原有提示)
                    sheet.Range[$"D{row}"].Value2 = item.Model ?? string.Empty;

                    // I 列 (第 9 列) = 品牌/生产厂家
                    sheet.Range[$"I{row}"].Value2 = item.Brand ?? string.Empty;

                    // L 列 (第 12 列) = 本体表价 (汇总表自带公式计算销售单价与总价)
                    sheet.Range[$"L{row}"].Value2 = item.Price > 0 ? (object)(double)item.Price : string.Empty;

                    // 若 M 列 (本体折扣) 当前为空或 0，则默认补 1 保障公式有效
                    string curM = Convert.ToString(sheet.Range[$"M{row}"].Value2)?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(curM) || curM == "0")
                    {
                        // 补齐本体折扣为 1
                        sheet.Range[$"M{row}"].Value2 = 1;
                    }

                    // P 列 (第 16 列) = 备注
                    if (!string.IsNullOrEmpty(item.Remark))
                    {
                        // 写入备注信息至 P 列
                        sheet.Range[$"P{row}"].Value2 = item.Remark;
                    }

                    // T 列 (第 20 列) = 额定电流数值
                    sheet.Range[$"T{row}"].Value2 = item.Current.HasValue && item.Current.Value > 0 ? (object)item.Current.Value : string.Empty;

                    // U 列 (第 21 列) = 极数 (如 "3", "4", "3P")
                    sheet.Range[$"U{row}"].Value2 = item.Poles ?? string.Empty;

                    // V 列 (第 22 列) = 脱扣方式 (如 "TM", "C", "D")
                    sheet.Range[$"V{row}"].Value2 = item.Tripping ?? string.Empty;

                    // W 列 (第 23 列) = 附件列 (重置清空，等待选附件时填入附件脱扣方式)
                    sheet.Range[$"W{row}"].Value2 = string.Empty;

                    // X 列 (第 24 列) = 扩展参数1 (如 "6kA", 结构尺寸等)
                    sheet.Range[$"X{row}"].Value2 = item.Param1 ?? string.Empty;

                    // Y 列 (第 25 列) = 扩展参数2 (如 "400V", 分类说明等)
                    sheet.Range[$"Y{row}"].Value2 = item.Param2 ?? string.Empty;

                    // 清除 D 列单元格的淡黄底色 (重置为无填充 xlNone)
                    dynamic cellD = sheet.Range[$"D{row}"];
                    cellD.Interior.ColorIndex = ComponentMatchDefaults.XlNoneColorIndex;
                }
                else
                {
                    // 在【常规分类明细表】中（按图片标准）：
                    // C 列 (第 3 列) = 标准规格型号
                    sheet.Range[$"C{row}"].Value2 = item.Model ?? string.Empty;

                    // D 列 (第 4 列) = 生产厂家/品牌
                    sheet.Range[$"D{row}"].Value2 = item.Brand ?? string.Empty;

                    // M 列 (第 13 列) = 表价 (本体基准价)
                    sheet.Range[$"M{row}"].Value2 = item.Price > 0 ? (object)(double)item.Price : string.Empty;

                    // W 列 (第 23 列) = Current 额定电流
                    sheet.Range[$"W{row}"].Value2 = item.Current.HasValue && item.Current.Value > 0 ? (object)item.Current.Value : string.Empty;

                    // X 列 (第 24 列) = Poles 极数 (如 "3", "4", "1P")
                    sheet.Range[$"X{row}"].Value2 = item.Poles ?? string.Empty;

                    // Y 列 (第 25 列) = trip 脱扣方式 (如 "D", "C")
                    sheet.Range[$"Y{row}"].Value2 = item.Tripping ?? string.Empty;

                    // Z 列 (第 26 列) = Accessory 附件 (置空，等待选附件时填入)
                    sheet.Range[$"Z{row}"].Value2 = string.Empty;

                    // AA 列 (第 27 列) = BlockName 块名/扩展参数1
                    sheet.Range[$"AA{row}"].Value2 = item.Param1 ?? string.Empty;

                    // AB 列 (第 28 列) = BlockCategory 块类别/扩展参数2
                    sheet.Range[$"AB{row}"].Value2 = item.Param2 ?? string.Empty;

                    // 清除 C 列与 D 列的可能残留底色
                    try
                    {
                        // 重置 C 列背景为无填充
                        sheet.Range[$"C{row}"].Interior.ColorIndex = ComponentMatchDefaults.XlNoneColorIndex;
                        // 重置 D 列背景为无填充
                        sheet.Range[$"D{row}"].Interior.ColorIndex = ComponentMatchDefaults.XlNoneColorIndex;
                    }
                    catch { }
                }

                // 统一补充并完成新值与新底色切片收集
                for (int i = 0; i < targetColumns.Length; i++)
                {
                    FinalizeSingleCellSlice(sheet, singleRowSlices[i], targetColumns[i].HasColor);
                }

                // 创建单行物料回填的撤销命令并推入撤销栈
                string modelDesc = !string.IsNullOrWhiteSpace(item.Model) ? item.Model : item.Name;
                var fillCommand = new RangeDeltaCommand($"物料回填: {modelDesc}", singleRowSlices);
                UndoRedoManager.Instance.PushCommand(fillCommand);

                return true;
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"FillSelectedComponentToActiveRow 回填异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 捕获单个单元格的原始差量切片数据包 (支持底色快照)
        /// </summary>
        private static RangeDeltaSlice CaptureSingleCellSlice(dynamic sheet, string sheetName, string col, int row, bool includeColor = false)
        {
            // 获取目标单元格对象
            dynamic cell = sheet.Range[$"{col}{row}"];
            // 提取修改前的值
            object oldVal = cell.Value2;
            // 实例化切片数据对象
            var slice = new RangeDeltaSlice
            {
                SheetName = sheetName,
                RangeAddress = $"{col}{row}",
                OldValues = new object[,] { { oldVal } }
            };

            // 若需捕获底色，记录原始 ColorIndex 与 Color
            if (includeColor)
            {
                slice.OldColorIndexes = new int[,] { { (int)cell.Interior.ColorIndex } };
                slice.OldColors = new int[,] { { (int)cell.Interior.Color } };
            }

            return slice;
        }

        /// <summary>
        /// 封装并补充单个单元格的新值与新底色切片
        /// </summary>
        private static void FinalizeSingleCellSlice(dynamic sheet, RangeDeltaSlice slice, bool includeColor = false)
        {
            // 获取目标单元格对象
            dynamic cell = sheet.Range[slice.RangeAddress];
            // 写入最新的值
            slice.NewValues = new object[,] { { cell.Value2 } };

            // 若需捕获底色，记录最新 ColorIndex 与 Color
            if (includeColor)
            {
                slice.NewColorIndexes = new int[,] { { (int)cell.Interior.ColorIndex } };
                slice.NewColors = new int[,] { { (int)cell.Interior.Color } };
            }
        }

        /// <summary>
        /// 用户选中配套附件后：D 列追加“+附件型号”（若数量大于1则为“+附件型号*数量”）
        /// 汇总表中附件填 N 列（多附件加法公式），常规表中本体和附件在 M 列进行加法公式连加
        /// </summary>
        /// <param name="attachment">选中的附件元器件物料 DTO</param>
        /// <param name="targetCell">目标单元格 COM 句柄</param>
        /// <param name="quantity">选定的附件数量，默认为 1</param>
        /// <returns>回填是否成功</returns>
        public static bool FillSelectedAttachmentToActiveRow(ComponentApiDto attachment, dynamic? targetCell = null, int quantity = 1)
        {
            // 校验附件物料对象非空
            if (attachment == null) return false;
            // 约束数量最小为 1
            if (quantity <= 0) quantity = 1;

            try
            {
                // 获取 Excel 宿主应用程序实例
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null) return false;

                // 获取活动单元格 COM 句柄
                dynamic cell = targetCell ?? app.ActiveCell;
                if (cell == null) return false;

                // 获取所在工作表
                dynamic sheet = cell.Worksheet;
                if (sheet == null) return false;

                // 获取当前物理行号
                int row = Convert.ToInt32(cell.Row);
                if (row <= 0) return false;

                // 判断是否在“元件汇总表”工作表中
                string sheetName = Convert.ToString(sheet.Name) ?? string.Empty;
                bool isSummarySheet = string.Equals(sheetName.Trim(), ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase);

                // 1. 获取型号所在列 (汇总表为 D 列，常规明细表为 C 列)
                string modelCol = isSummarySheet ? "D" : "C";
                string oldModel = Convert.ToString(sheet.Range[$"{modelCol}{row}"].Value2)?.Trim() ?? string.Empty;
                // 若原单元格内容为初始提示词“点击查询”，则清空
                if (string.Equals(oldModel, "点击查询", StringComparison.OrdinalIgnoreCase))
                {
                    // 重置为空文本
                    oldModel = string.Empty;
                }

                // 附件基础型号
                string attachModel = attachment.Model?.Trim() ?? string.Empty;
                // 若数量大于 1，则格式化为 "附件型号*数量" (如 "MX*2")，否则保持 "附件型号"
                string attachModelWithQty = quantity > 1 ? $"{attachModel}*{quantity}" : attachModel;

                // 拼接新型号文本：“原内容+附件型号[*数量]”
                string newModel;
                if (string.IsNullOrEmpty(oldModel))
                {
                    // 原内容为空时直接作为新型号
                    newModel = attachModelWithQty;
                }
                else
                {
                    // 消除首尾已有的加号字符，防止出现连续的 "++" 符号
                    string trimmedOld = oldModel.TrimEnd('+', '＋');
                    string trimmedAttach = attachModelWithQty.TrimStart('+', '＋');
                    newModel = $"{trimmedOld}+{trimmedAttach}";
                }

                // 回填到对应型号列
                sheet.Range[$"{modelCol}{row}"].Value2 = newModel;

                // 2. 处理附件价格回填逻辑
                decimal attachPrice = attachment.Price > 0 ? attachment.Price : 0m;
                if (attachPrice > 0)
                {
                    // 格式化价格单项表达式：数量大于 1 时为 "单价*数量" (如 "336.2*2")，否则为纯单价 (如 "336.2")
                    string priceTerm = quantity > 1 ? $"{attachPrice:0.##}*{quantity}" : $"{attachPrice:0.##}";

                    if (isSummarySheet)
                    {
                        // 在“元件汇总表”中：附件价格填入 N 列 (附件表价)，多个附件用加法公式连接
                        string oldFormulaN = Convert.ToString(sheet.Range[$"N{row}"].Formula)?.Trim() ?? string.Empty;
                        string oldValN = Convert.ToString(sheet.Range[$"N{row}"].Value2)?.Trim() ?? string.Empty;

                        // 若此前 N 列没有任何价格数据
                        if (string.IsNullOrEmpty(oldFormulaN) && string.IsNullOrEmpty(oldValN))
                        {
                            if (quantity > 1)
                            {
                                // 数量大于 1 时以乘法公式写入，例如 "=336.2*2"
                                sheet.Range[$"N{row}"].Formula = $"={priceTerm}";
                            }
                            else
                            {
                                // 单件直接填入数值
                                sheet.Range[$"N{row}"].Value2 = (double)attachPrice;
                            }
                        }
                        // 若此前已有公式 (以 '=' 开头)，直接在原公式末尾累加 "+价格表达式"
                        else if (!string.IsNullOrEmpty(oldFormulaN) && oldFormulaN.StartsWith("="))
                        {
                            sheet.Range[$"N{row}"].Formula = $"{oldFormulaN}+{priceTerm}";
                        }
                        // 若此前为纯数值 (例如 49)，升级为加法公式 (例如 =49+336.2*2)
                        else
                        {
                            string baseNum = !string.IsNullOrEmpty(oldValN) ? oldValN : oldFormulaN;
                            sheet.Range[$"N{row}"].Formula = $"={baseNum}+{priceTerm}";
                        }

                        // 若 O 列 (附件折扣) 为空或 0，默认填入 1，保障表格自带内置公式计算有效
                        string curO = Convert.ToString(sheet.Range[$"O{row}"].Value2)?.Trim() ?? string.Empty;
                        if (string.IsNullOrEmpty(curO) || curO == "0")
                        {
                            sheet.Range[$"O{row}"].Value2 = 1;
                        }
                    }
                    else
                    {
                        // 常规分类明细表中：本体和附件在 M 列 (表价) 连加 (如 =447.01+336.2*2 或 =447.01+150+60*2)
                        string oldFormulaM = Convert.ToString(sheet.Range[$"M{row}"].Formula)?.Trim() ?? string.Empty;
                        string oldValM = Convert.ToString(sheet.Range[$"M{row}"].Value2)?.Trim() ?? string.Empty;

                        string newFormulaM;
                        // 若此前 M 列无任何数据，直接写公式 "=价格表达式"
                        if (string.IsNullOrEmpty(oldFormulaM) && string.IsNullOrEmpty(oldValM))
                        {
                            newFormulaM = $"={priceTerm}";
                        }
                        // 若此前已有公式 (以 '=' 开头)，直接在原公式末尾连加 "+价格表达式"
                        else if (!string.IsNullOrEmpty(oldFormulaM) && oldFormulaM.StartsWith("="))
                        {
                            newFormulaM = $"{oldFormulaM}+{priceTerm}";
                        }
                        // 若此前为本体纯数值 (如 447.01)，升级为连加公式 (如 =447.01+336.2*2)
                        else
                        {
                            string basePriceStr = !string.IsNullOrEmpty(oldValM) ? oldValM : oldFormulaM;
                            newFormulaM = $"={basePriceStr}+{priceTerm}";
                        }

                        // 写入 M 列公式
                        sheet.Range[$"M{row}"].Formula = newFormulaM;
                    }
                }

                // 3. 处理附件信息列回填 (汇总表填入 W 列，常规分类明细表填入 Z 列 Accessory)
                string attachCol = isSummarySheet ? "W" : "Z";
                // 提取当前选中附件的脱扣方式 (如 "MX", "OF", "MN" 等)
                string attachTripping = attachment.Tripping?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(attachTripping))
                {
                    // 获取当前行附件列原有内容
                    string oldTrip = Convert.ToString(sheet.Range[$"{attachCol}{row}"].Value2)?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(oldTrip))
                    {
                        // 若附件列为空则直接写入附件脱扣方式
                        sheet.Range[$"{attachCol}{row}"].Value2 = attachTripping;
                    }
                    else
                    {
                        // 若已有内容则以“+”连接追加
                        string trimmedOldTrip = oldTrip.TrimEnd('+', '＋');
                        string trimmedAttachTrip = attachTripping.TrimStart('+', '＋');
                        sheet.Range[$"{attachCol}{row}"].Value2 = $"{trimmedOldTrip}+{trimmedAttachTrip}";
                    }
                }

                // 4. 清除底色 (C 列与 D 列)
                try
                {
                    // 清除型号列底色
                    sheet.Range[$"{modelCol}{row}"].Interior.ColorIndex = ComponentMatchDefaults.XlNoneColorIndex;
                    // 清除 D 列底色
                    sheet.Range[$"D{row}"].Interior.ColorIndex = ComponentMatchDefaults.XlNoneColorIndex;
                }
                catch { }

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"FillSelectedAttachmentToActiveRow 异常: {ex.Message}");
                return false;
            }
        }
    }
}
