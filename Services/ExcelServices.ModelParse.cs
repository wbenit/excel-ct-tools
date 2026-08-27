using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 公共服务类: 元器件型号参数（极数与电流）双通道流水线解析与内存批量回填
    /// </summary>
    public static partial class ExcelServices
    {
        // 预定义排除占位掩码字符串
        private const string MASK_PLACEHOLDER = " ___MASKED___ ";

        // 缓存当前已打开的型号参数识别配置窗体实例
        private static ModelParamParserForm? _modelParamParserForm = null;

        /// <summary>
        /// 弹出基于 WebView2 + Vue 3 的“元器件型号参数识别设置”窗口
        /// </summary>
        public static void ShowModelParamParserDialog()
        {
            try
            {
                // 以标准非模态方式展示型号参数识别窗口 (挂载 Excel 主句柄并保持置顶交互)
                ShowModelessForm(ref _modelParamParserForm, () => new ModelParamParserForm());
            }
            catch (Exception ex)
            {
                // 弹出异常提示
                System.Windows.Forms.MessageBox.Show($"打开型号参数识别窗口失败: {ex.Message}", "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 极数解析核心逻辑: 经历前置排除过滤 -> 多级顺位流水线 -> 后置标称白名单校验
        /// </summary>
        /// <param name="rawModel">原始型号文本</param>
        /// <param name="config">配置对象</param>
        /// <param name="hitRuleTitle">命中的规则标题</param>
        /// <returns>提取出的极数结果（如 3P 或 3）</returns>
        public static string ParsePoles(string rawModel, ModelParserConfig config, out string hitRuleTitle)
        {
            // 初始化命中规则输出变量
            hitRuleTitle = string.Empty;
            // 校验输入字符串是否为空
            if (string.IsNullOrWhiteSpace(rawModel)) return string.Empty;

            // 1. 前置必去项与负向排除过滤 (如将 IP65, IP20 中的 P 过滤，避免误识别为极数)
            string cleanText = MaskPoleExclusions(rawModel, config.PoleExcludeKeywords);

            // 2. 依次遍历用户配置的极数顺位流水线规则
            if (config.PolePipeline != null)
            {
                foreach (var rule in config.PolePipeline)
                {
                    // 若当前顺位规则被用户禁用，则跳过
                    if (!rule.Enabled) continue;

                    // 尝试使用当前规则提取极数
                    string extracted = TryExtractPoleByRule(cleanText, rule);

                    // 检查是否提取到了非空内容
                    if (!string.IsNullOrWhiteSpace(extracted))
                    {
                        // 3. 后置标称白名单校验 (如过滤 8P, 99P 等非标极数)
                        if (ValidatePoleWhitelist(extracted, config))
                        {
                            // 记录命中的顺位规则标题
                            hitRuleTitle = rule.Title;
                            // 格式化输出 (3P 或 3)
                            return FormatPoleOutput(extracted, config.PoleFormat);
                        }
                    }
                }
            }

            // 所有顺位均未匹配或未通过白名单，返回空
            return string.Empty;
        }

        /// <summary>
        /// 电流解析核心逻辑: 提取出最小电流与最大电流 (支持区间 2.5-4A 分流及单值 100A 同值)
        /// </summary>
        /// <param name="rawModel">原始型号文本</param>
        /// <param name="config">配置对象</param>
        /// <param name="hitRuleTitle">命中的规则标题</param>
        /// <param name="candidateList">所有被成功提取并符合白名单的候选电流值</param>
        /// <returns>提取出的最小电流与最大电流元组 (MinCurrent, MaxCurrent)</returns>
        public static (string MinCurrent, string MaxCurrent) ParseCurrentMinMax(string rawModel, ModelParserConfig config, out string hitRuleTitle, out List<string> candidateList)
        {
            // 初始化命中规则输出变量
            hitRuleTitle = string.Empty;
            // 初始化候选值输出列表
            candidateList = new List<string>();

            // 校验输入字符串有效性
            if (string.IsNullOrWhiteSpace(rawModel)) return (string.Empty, string.Empty);

            // 1. 前置必去项与负向排除过滤 (如将 100mA, 10kA, 220V 掩码，避免误提取为额定电流)
            string cleanText = MaskCurrentExclusions(rawModel, config.CurrentExcludeKeywords);

            // 存储收集到的所有合法数值及其规则来源
            var validCandidates = new List<(double Value, string OriginalStr, string RuleTitle)>();

            // 2. 依次遍历用户配置的所有启用电流顺位规则
            if (config.CurrentPipeline != null)
            {
                foreach (var rule in config.CurrentPipeline)
                {
                    // 跳过未启用的顺位规则
                    if (!rule.Enabled) continue;

                    // 从清洗后的文本中提取符合该规则的所有候选原始字符串
                    var rawMatches = ExtractRawCurrentCandidatesByRule(cleanText, rule);

                    foreach (var rawMatch in rawMatches)
                    {
                        // 若无法转为 int，固定提取数字；若为区间，则展开两个端点数字
                        var numericValues = ExtractPureNumbersFromMatch(rawMatch);

                        foreach (var numVal in numericValues)
                        {
                            string numStr = numVal.ToString("0.###", CultureInfo.InvariantCulture);

                            // 3. 后置标称范围白名单校验 (如过滤 99 等非标数值)
                            if (ValidateCurrentWhitelist(numStr, config))
                            {
                                // 防止同一规则重复添加完全相同的值
                                if (!validCandidates.Exists(c => Math.Abs(c.Value - numVal) < 0.0001))
                                {
                                    validCandidates.Add((numVal, numStr, rule.Title));
                                }
                            }
                        }
                    }
                }
            }

            // 若未找到任何合规的候选电流，返回空元组
            if (validCandidates.Count == 0)
            {
                return (string.Empty, string.Empty);
            }

            // 填充输出的候选字符串列表
            foreach (var item in validCandidates)
            {
                candidateList.Add(item.OriginalStr);
            }

            // 4. 计算最小电流与最大电流候选
            var minItem = validCandidates[0];
            var maxItem = validCandidates[0];
            for (int i = 1; i < validCandidates.Count; i++)
            {
                // 寻找最小值
                if (validCandidates[i].Value < minItem.Value)
                {
                    minItem = validCandidates[i];
                }
                // 寻找最大值
                if (validCandidates[i].Value > maxItem.Value)
                {
                    maxItem = validCandidates[i];
                }
            }

            // 记录命中的规则名称
            hitRuleTitle = minItem.RuleTitle;

            // 5. 格式化输出 (根据配置决定是否带 A)
            string minFormatted = FormatCurrentOutput(minItem.OriginalStr, config.CurrentFormat);
            string maxFormatted = FormatCurrentOutput(maxItem.OriginalStr, config.CurrentFormat);

            return (minFormatted, maxFormatted);
        }

        /// <summary>
        /// 兼容保留原 ParseCurrent 接口 (返回最小电流)
        /// </summary>
        public static string ParseCurrent(string rawModel, ModelParserConfig config, out string hitRuleTitle, out List<string> candidateList)
        {
            // 调用最小最大解析接口
            var (minCur, _) = ParseCurrentMinMax(rawModel, config, out hitRuleTitle, out candidateList);
            // 返回最小电流
            return minCur;
        }

        /// <summary>
        /// 兼容保留原单输出 ParseCurrent 接口
        /// </summary>
        public static string ParseCurrent(string rawModel, ModelParserConfig config, out string hitRuleTitle)
        {
            // 调用重载方法
            return ParseCurrent(rawModel, config, out hitRuleTitle, out _);
        }

        /// <summary>
        /// 脱扣方式解析核心逻辑: 经历前置排除过滤 -> 多级顺位流水线 -> 后置标称白名单校验
        /// </summary>
        /// <param name="rawModel">原始型号文本</param>
        /// <param name="config">配置对象</param>
        /// <param name="hitRuleTitle">命中的规则标题</param>
        /// <returns>提取出的脱扣方式简写代号（如 TM, C, D, MA, Elec 等）</returns>
        public static string ParseTripMode(string rawModel, ModelParserConfig config, out string hitRuleTitle)
        {
            // 初始化命中规则输出变量
            hitRuleTitle = string.Empty;
            // 校验输入字符串有效性
            if (string.IsNullOrWhiteSpace(rawModel)) return string.Empty;

            // 1. 前置必去项排除过滤
            string cleanText = MaskTripModeExclusions(rawModel, config.TripModeExcludeKeywords);

            // 2. 依次遍历用户配置的脱扣方式顺位流水线规则
            if (config.TripModePipeline != null)
            {
                foreach (var rule in config.TripModePipeline)
                {
                    // 跳过未启用的顺位规则
                    if (!rule.Enabled) continue;

                    // 尝试使用当前顺位模式提取脱扣方式
                    string extracted = TryExtractTripModeByRule(cleanText, rule);

                    // 判断是否提取到了有效简写代号
                    if (!string.IsNullOrWhiteSpace(extracted))
                    {
                        // 3. 后置白名单校验 (若开启严格白名单)
                        if (ValidateTripModeWhitelist(extracted, config))
                        {
                            // 记录命中的顺位规则名称
                            hitRuleTitle = rule.Title;
                            return extracted.Trim();
                        }
                    }
                }
            }

            // 未成功提取，返回空
            return string.Empty;
        }

        /// <summary>
        /// 单个型号解析综合测试接口 (用于前端沙盒实时预览)
        /// </summary>
        public static ParseResultDto ParseSingleModel(string rawModel, ModelParserConfig config)
        {
            // 创建解析结果 DTO 实例
            var result = new ParseResultDto
            {
                // 记录原始输入
                RawModel = rawModel ?? string.Empty
            };

            // 针对空输入直接返回 Failed 状态
            if (string.IsNullOrWhiteSpace(rawModel))
            {
                result.Status = "Failed";
                return result;
            }

            // 1. 执行极数流水线解析
            result.Pole = ParsePoles(rawModel ?? string.Empty, config, out string hitPole);
            result.HitPoleRule = hitPole;

            // 2. 执行电流流水线解析 (分别提取最小电流与最大电流，同时获取所有合法候选列表)
            var (minCur, maxCur) = ParseCurrentMinMax(rawModel ?? string.Empty, config, out string hitCurrent, out var candidateList);
            result.MinCurrent = minCur;
            result.MaxCurrent = maxCur;
            result.Current = minCur; // 兼容旧字段
            result.HitCurrentRule = hitCurrent;
            result.CandidateCurrents = candidateList;

            // 3. 执行脱扣方式流水线解析
            result.TripMode = ParseTripMode(rawModel ?? string.Empty, config, out string hitTrip);
            result.HitTripModeRule = hitTrip;

            // 综合评估解析结果状态
            bool hasPole = !string.IsNullOrWhiteSpace(result.Pole);
            bool hasCur = !string.IsNullOrWhiteSpace(result.MinCurrent) || !string.IsNullOrWhiteSpace(result.MaxCurrent);
            bool hasTrip = !string.IsNullOrWhiteSpace(result.TripMode);

            // 三项均成功提取记为 Success (若未配置脱扣顺位则按两项判定)
            bool expectTrip = config.TripModePipeline != null && config.TripModePipeline.Exists(r => r.Enabled);
            if (hasPole && hasCur && (!expectTrip || hasTrip))
            {
                result.Status = "Success";
            }
            // 任意提取到其中一项记为 Partial
            else if (hasPole || hasCur || hasTrip)
            {
                result.Status = "Partial";
            }
            // 均未提取到记为 Failed
            else
            {
                result.Status = "Failed";
            }

            // 返回综合分析结果
            return result;
        }

        /// <summary>
        /// 执行 Excel 批量解析与回填任务 (严格采用二维数组一次性读入与写入)
        /// </summary>
        public static BatchParseResult ExecuteBatchModelParse(ModelParserConfig config)
        {
            // 初始化计时器精确计算总耗时
            var stopwatch = Stopwatch.StartNew();
            // 初始化批量执行结果返回对象
            var batchResult = new BatchParseResult();

            try
            {
                // 获取当前正在运行的 Excel 顶级 Application 实例
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null)
                {
                    batchResult.Success = false;
                    batchResult.Message = "未检测到运行中的 Excel 应用程序实例";
                    return batchResult;
                }

                // 获取当前活动工作簿
                dynamic? activeWb = app.ActiveWorkbook;
                // 获取当前活动工作表
                dynamic? activeSheet = app.ActiveSheet;

                if (activeWb == null || activeSheet == null)
                {
                    batchResult.Success = false;
                    batchResult.Message = "请先在 Excel 中打开或激活一个工作表";
                    return batchResult;
                }

                // 规范化列名 (转为大写且去除空格)
                string colSrc = string.IsNullOrWhiteSpace(config.SourceColumn) ? "C" : config.SourceColumn.Trim().ToUpper();
                string colMinCur = string.IsNullOrWhiteSpace(config.MinCurrentColumn) ? string.Empty : config.MinCurrentColumn.Trim().ToUpper();
                string colMaxCur = string.IsNullOrWhiteSpace(config.MaxCurrentColumn) ? string.Empty : config.MaxCurrentColumn.Trim().ToUpper();
                string colPole = string.IsNullOrWhiteSpace(config.PoleColumn) ? "T" : config.PoleColumn.Trim().ToUpper();
                string colTrip = string.IsNullOrWhiteSpace(config.TripModeColumn) ? "U" : config.TripModeColumn.Trim().ToUpper();

                // ==================== 0. 自动根据配置在工作表第 5 行对应列写入表头 ====================
                AddModelParserHeadersToExcel(config);

                // 获取用户在 Excel 中当前选中的区域 (Selection)
                dynamic? selection = app.Selection;
                if (selection == null)
                {
                    batchResult.Success = false;
                    batchResult.Message = "请先在 Excel 中选择需要识别的单元格或行区域";
                    return batchResult;
                }

                // 初始化统计计数器
                int totalRows = 0;
                int successCount = 0;
                int failedCount = 0;

                // 遍历当前选择的所有选区 (支持单个连续选区或按住 Ctrl 的多选区)
                foreach (dynamic area in selection.Areas)
                {
                    // 获取当前选区的起始物理行号与总行数
                    int startRow = (int)area.Row;
                    int rowCount = (int)area.Rows.Count;
                    int endRow = startRow + rowCount - 1;

                    // 若选区行数无效则跳过
                    if (rowCount <= 0) continue;
                    totalRows += rowCount;

                    // ==================== 1. 一次性将选中区域对应源型号整块读入二维数组 ====================
                    dynamic srcRange = activeSheet.Range[$"{colSrc}{startRow}:{colSrc}{endRow}"];
                    object[,] rawArray = ConvertTo2DArray(srcRange.Value2, rowCount);

                    // ==================== 2. 在内存中创建当前选区对应的目标二维数组 ====================
                    object[,] minCurrentArray = new object[rowCount, 1];
                    object[,] maxCurrentArray = new object[rowCount, 1];
                    object[,] poleArray = new object[rowCount, 1];
                    object[,] tripArray = new object[rowCount, 1];

                    // 循环遍历当前选区内存中的二维数组行
                    for (int i = 0; i < rowCount; i++)
                    {
                        // 获取当前行源型号内容
                        object val = rawArray[i + 1, 1];
                        string rawModel = val?.ToString()?.Trim() ?? string.Empty;

                        // 若为空行则置空保留
                        if (string.IsNullOrWhiteSpace(rawModel))
                        {
                            minCurrentArray[i, 0] = string.Empty;
                            maxCurrentArray[i, 0] = string.Empty;
                            poleArray[i, 0] = string.Empty;
                            tripArray[i, 0] = string.Empty;
                            continue;
                        }

                        // 调用三通道流水线解析极数、最小/最大电流与脱扣方式
                        string pole = ParsePoles(rawModel, config, out _);
                        var (minCurrent, maxCurrent) = ParseCurrentMinMax(rawModel, config, out _, out _);
                        string tripMode = ParseTripMode(rawModel, config, out _);

                        // 填充内存二维数组
                        minCurrentArray[i, 0] = minCurrent;
                        maxCurrentArray[i, 0] = maxCurrent;
                        poleArray[i, 0] = pole;
                        tripArray[i, 0] = tripMode;

                        // 统计识别成功与失败数
                        if (!string.IsNullOrWhiteSpace(pole) && (!string.IsNullOrWhiteSpace(minCurrent) || !string.IsNullOrWhiteSpace(maxCurrent)))
                        {
                            successCount++;
                        }
                        else
                        {
                            failedCount++;
                        }
                    }

                    // ==================== 3. 一次性将二维数组整块写回当前选区目标列 ====================
                    // 写回最小电流列
                    if (!string.IsNullOrWhiteSpace(colMinCur))
                    {
                        dynamic minCurRange = activeSheet.Range[$"{colMinCur}{startRow}:{colMinCur}{endRow}"];
                        minCurRange.Value2 = minCurrentArray;
                    }

                    // 写回最大电流列 (若配置了独立列)
                    if (!string.IsNullOrWhiteSpace(colMaxCur) && colMaxCur != colMinCur)
                    {
                        dynamic maxCurRange = activeSheet.Range[$"{colMaxCur}{startRow}:{colMaxCur}{endRow}"];
                        maxCurRange.Value2 = maxCurrentArray;
                    }

                    // 写回极数列
                    if (!string.IsNullOrWhiteSpace(colPole))
                    {
                        dynamic poleRange = activeSheet.Range[$"{colPole}{startRow}:{colPole}{endRow}"];
                        poleRange.Value2 = poleArray;
                    }

                    // 写回脱扣方式列
                    if (!string.IsNullOrWhiteSpace(colTrip))
                    {
                        dynamic tripRange = activeSheet.Range[$"{colTrip}{startRow}:{colTrip}{endRow}"];
                        tripRange.Value2 = tripArray;
                    }
                }

                // 检查选区内是否有有效数据行
                if (totalRows == 0)
                {
                    batchResult.Success = true;
                    batchResult.TotalRows = 0;
                    batchResult.Message = "当前选择区域内无有效数据行";
                    return batchResult;
                }

                // 停止计时并汇总结果
                stopwatch.Stop();
                batchResult.Success = true;
                batchResult.TotalRows = totalRows;
                batchResult.SuccessCount = successCount;
                batchResult.FailedCount = failedCount;
                batchResult.ElapsedMilliseconds = stopwatch.ElapsedMilliseconds;
                batchResult.Message = $"选中区域处理完成：共处理 {totalRows} 行，有效识别 {successCount} 行，耗时 {stopwatch.ElapsedMilliseconds} ms";
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                batchResult.Success = false;
                batchResult.Message = $"批量执行异常: {ex.Message}";
            }

            // 返回最终执行统计报告
            return batchResult;
        }

        /// <summary>
        /// 根据配置的列映射在当前工作表第 5 行添加表头 (型号、最小电流、最大电流、极数、脱扣方式)
        /// </summary>
        /// <param name="config">型号参数识别配置对象</param>
        /// <returns>操作结果元组 (Success: 是否成功, Message: 提示信息)</returns>
        public static (bool Success, string Message) AddModelParserHeadersToExcel(ModelParserConfig config)
        {
            try
            {
                // 获取当前正在运行的 Excel 顶级 Application 实例
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null)
                {
                    // 未检测到运行中的 Excel 实例
                    return (false, "未检测到运行中的 Excel 应用程序实例");
                }

                // 获取当前活动工作表
                dynamic? activeSheet = app.ActiveSheet;
                if (activeSheet == null)
                {
                    // 未激活工作表提示
                    return (false, "请先在 Excel 中打开或激活一个工作表");
                }

                // 固定设置表头所在的行号为第 5 行
                const int HEADER_ROW = 5; // --硬编码-- 固定表头设置在第 5 行

                // 规范化列名 (转为大写且去除多余首尾空格)
                string colSrc = string.IsNullOrWhiteSpace(config.SourceColumn) ? string.Empty : config.SourceColumn.Trim().ToUpper();
                string colMinCur = string.IsNullOrWhiteSpace(config.MinCurrentColumn) ? string.Empty : config.MinCurrentColumn.Trim().ToUpper();
                string colMaxCur = string.IsNullOrWhiteSpace(config.MaxCurrentColumn) ? string.Empty : config.MaxCurrentColumn.Trim().ToUpper();
                string colPole = string.IsNullOrWhiteSpace(config.PoleColumn) ? string.Empty : config.PoleColumn.Trim().ToUpper();
                string colTrip = string.IsNullOrWhiteSpace(config.TripModeColumn) ? string.Empty : config.TripModeColumn.Trim().ToUpper();

                // 存储待写入表头的列与对应标题文本列表
                var headersToSet = new List<(string Col, string HeaderText)>();

                // 1. 添加源型号列表头
                if (!string.IsNullOrWhiteSpace(colSrc))
                {
                    // 源型号列默认表头为“型号”
                    headersToSet.Add((colSrc, "型号")); // --硬编码-- 默认源型号列表头
                }

                // 2. 添加最小电流列表头
                if (!string.IsNullOrWhiteSpace(colMinCur))
                {
                    // 最小电流列默认表头为“最小电流”
                    headersToSet.Add((colMinCur, "最小电流")); // --硬编码-- 默认最小电流列表头
                }

                // 3. 添加最大电流列表头 (若已配置且与最小电流列不重复)
                if (!string.IsNullOrWhiteSpace(colMaxCur) && colMaxCur != colMinCur)
                {
                    // 最大电流列默认表头为“最大电流”
                    headersToSet.Add((colMaxCur, "最大电流")); // --硬编码-- 默认最大电流列表头
                }

                // 4. 添加极数输出列表头
                if (!string.IsNullOrWhiteSpace(colPole))
                {
                    // 极数输出列默认表头为“极数”
                    headersToSet.Add((colPole, "极数")); // --硬编码-- 默认极数输出列表头
                }

                // 5. 添加脱扣方式输出列表头
                if (!string.IsNullOrWhiteSpace(colTrip))
                {
                    // 脱扣方式输出列默认表头为“脱扣方式”
                    headersToSet.Add((colTrip, "脱扣方式")); // --硬编码-- 默认脱扣方式列表头
                }

                // 检查是否有有效列配置
                if (headersToSet.Count == 0)
                {
                    // 无有效列配置提示
                    return (false, "未配置任何有效的列映射，无法添加表头");
                }

                // 记录成功写入的表头描述清单
                var successList = new List<string>();

                // 逐个向第 5 行对应列写入表头并应用格式
                foreach (var (col, headerText) in headersToSet)
                {
                    // 获取目标第 5 行单元格 Range 对象
                    dynamic cell = activeSheet.Range[$"{col}{HEADER_ROW}"];

                    // 写入表头文本内容
                    cell.Value2 = headerText;

                    // 设置水平居中对齐 (xlCenter = -4108)
                    cell.HorizontalAlignment = -4108; // --硬编码-- 居中对齐常量

                    // 设置垂直居中对齐 (xlCenter = -4108)
                    cell.VerticalAlignment = -4108; // --硬编码-- 居中对齐常量

                    // 设置字体加粗
                    cell.Font.Bold = true;

                    // 记录写入描述
                    successList.Add($"{col}列 [{headerText}]");
                }

                // 返回成功消息
                return (true, $"已在第 {HEADER_ROW} 行成功添加表头：{string.Join("，", successList)}");
            }
            catch (Exception ex)
            {
                // 捕获并返回异常信息
                return (false, $"添加表头失败: {ex.Message}");
            }
        }

        // ===================================================================================
        // 内部辅助方法与正则提取引擎
        // ===================================================================================

        /// <summary>
        /// 屏蔽极数前置排除关键词 (如将 IP65, DPN, PF 等干扰项掩码)
        /// </summary>
        private static string MaskPoleExclusions(string input, List<string>? exclusions)
        {
            if (string.IsNullOrWhiteSpace(input) || exclusions == null || exclusions.Count == 0)
                return input;

            string result = input;
            foreach (var word in exclusions)
            {
                if (string.IsNullOrWhiteSpace(word)) continue;
                // 正则匹配排除词后跟数字或字母的组合 (例如 IP\s*\d+)
                string pattern = $@"(?i)\b{Regex.Escape(word.Trim())}\s*\d*\b";
                result = Regex.Replace(result, pattern, MASK_PLACEHOLDER);
            }
            return result;
        }

        /// <summary>
        /// 屏蔽电流前置排除关键词 (如将 100mA, 10kA, 220VAC, 50Hz 等干扰项掩码)
        /// </summary>
        private static string MaskCurrentExclusions(string input, List<string>? exclusions)
        {
            if (string.IsNullOrWhiteSpace(input) || exclusions == null || exclusions.Count == 0)
                return input;

            string result = input;
            foreach (var word in exclusions)
            {
                if (string.IsNullOrWhiteSpace(word)) continue;
                // 正则匹配紧跟排除单位的数值 (例如 100mA, 6kA, 220V)
                string pattern = $@"(?i)\d+(?:\.\d+)?\s*{Regex.Escape(word.Trim())}\b";
                result = Regex.Replace(result, pattern, MASK_PLACEHOLDER);
            }
            return result;
        }

        /// <summary>
        /// 根据指定的顺位规则尝试提取极数
        /// </summary>
        private static string TryExtractPoleByRule(string text, PipelineRuleItem rule)
        {
            switch (rule.Mode)
            {
                // 顺位模式 1: 寻找 "P" 前面的数字 (支持 1P, 2P, 3P, 4P, 1P+N, 3P+N 等)
                case "FindBeforeP":
                case "NumberWithP":
                    var matchP = Regex.Match(text, @"(?i)(?:^|[^a-zA-Z0-9])([1-4]\s*P(?:\s*\+\s*N)?)(?:[^a-zA-Z0-9]|$)");
                    if (matchP.Success)
                    {
                        return matchP.Groups[1].Value.Replace(" ", "").ToUpper();
                    }
                    break;

                // 顺位模式: 空格+数字+空格 (如独立的 " 3 " 或 " 4 ")
                case "SpaceNumberSpace":
                    var matchSpace = Regex.Match(text, @"(?:^|\s)([1-4])(?:\s|$)");
                    if (matchSpace.Success)
                    {
                        return matchSpace.Groups[1].Value + "P";
                    }
                    break;

                // 顺位模式 2: 脱扣代号或特殊特征对照表映射 (如 /3300 -> 3P)，区分大小写
                case "CodeMapping":
                    if (rule.Mapping != null)
                    {
                        foreach (var kvp in rule.Mapping)
                        {
                            // 严格区分大小写匹配代号关键词
                            if (!string.IsNullOrWhiteSpace(kvp.Key) &&
                                text.IndexOf(kvp.Key, StringComparison.Ordinal) >= 0)
                            {
                                return kvp.Value?.Trim() ?? string.Empty;
                            }
                        }
                    }
                    break;

                // 顺位模式 3: 寻找中文“极”前面的数字 (如 3极, 4极)
                case "FindBeforeJi":
                    var matchJi = Regex.Match(text, @"([1-4])\s*极");
                    if (matchJi.Success)
                    {
                        return matchJi.Groups[1].Value + "P";
                    }
                    break;

                // 顺位模式 4: 斜杠后的单数字极数 (如 /3, /4)
                case "FrameCode":
                    var matchFrame = Regex.Match(text, @"\/([1-4])(?:\s|$|[A-Za-z])");
                    if (matchFrame.Success)
                    {
                        return matchFrame.Groups[1].Value + "P";
                    }
                    break;

                // 顺位模式 5: 固定极数值 (如固定为 1P+N, 3P)
                case "FixedValue":
                    return rule.FixedValue?.Trim() ?? string.Empty;

                // 顺位模式 6: 自定义正则表达式 (支持 (?<pole>...) 命名捕获组或首个括号捕获)
                case "CustomRegex":
                    if (!string.IsNullOrWhiteSpace(rule.CustomRegex))
                    {
                        try
                        {
                            var matchCustom = Regex.Match(text, rule.CustomRegex);
                            if (matchCustom.Success)
                            {
                                if (matchCustom.Groups["pole"].Success)
                                {
                                    return matchCustom.Groups["pole"].Value.Trim();
                                }
                                if (matchCustom.Groups.Count > 1)
                                {
                                    return matchCustom.Groups[1].Value.Trim();
                                }
                                return matchCustom.Value.Trim();
                            }
                        }
                        catch { }
                    }
                    break;
            }

            return string.Empty;
        }

        /// <summary>
        /// 根据指定的顺位规则尝试提取电流（返回首个匹配，保持旧接口兼容）
        /// </summary>
        private static string TryExtractCurrentByRule(string text, PipelineRuleItem rule)
        {
            // 提取该规则下的所有候选原始匹配项
            var candidates = ExtractRawCurrentCandidatesByRule(text, rule);
            if (candidates.Count > 0)
            {
                // 提取首个候选中的纯数字部分
                var numbers = ExtractPureNumbersFromMatch(candidates[0]);
                if (numbers.Count > 0)
                {
                    return numbers[0].ToString("0.###", CultureInfo.InvariantCulture);
                }
                return candidates[0];
            }
            return string.Empty;
        }

        /// <summary>
        /// 根据规则提取清洗后文本中的所有原始候选匹配项
        /// </summary>
        private static List<string> ExtractRawCurrentCandidatesByRule(string text, PipelineRuleItem rule)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(text) || rule == null) return list;

            switch (rule.Mode)
            {
                // 顺位模式 1: 寻找带单位 "A" 的数字 (如 100A, 63A, 0.03A)
                case "NumberWithA":
                    var matchesA = Regex.Matches(text, @"(?i)(?:^|[^a-zA-Z0-9\.])(\d+(?:\.\d+)?)\s*A(?:[^a-zA-Z0-9]|$)");
                    foreach (Match m in matchesA)
                    {
                        if (m.Success && m.Groups.Count > 1)
                        {
                            list.Add(m.Groups[1].Value.Trim());
                        }
                    }
                    break;

                // 顺位模式 2: 寻找脱扣曲线字母 [C/D/K/Z] 后面的数字 (如 C32, D16)
                case "CurveLetterNumber":
                    var matchesCurve = Regex.Matches(text, @"(?i)(?:^|[^a-zA-Z0-9])[CDKZ](\d+(?:\.\d+)?)(?:[^\d\.]|$)");
                    foreach (Match m in matchesCurve)
                    {
                        if (m.Success && m.Groups.Count > 1)
                        {
                            list.Add(m.Groups[1].Value.Trim());
                        }
                    }
                    break;

                // 顺位模式 3: 寻找整定电流可调区间 (如 2.5-4A, 0.63-1A)
                case "CurrentRange":
                    var matchesRange = Regex.Matches(text, @"(?i)(\d+(?:\.\d+)?\s*[-~～至]\s*\d+(?:\.\d+)?)\s*A?");
                    foreach (Match m in matchesRange)
                    {
                        if (m.Success && m.Groups.Count > 1)
                        {
                            list.Add(m.Groups[1].Value.Replace(" ", "").Trim());
                        }
                    }
                    break;

                // 顺位模式: 空格+数字+空格 (如独立的纯数字电流 " 100 " 或 " 63 ")
                case "SpaceNumberSpace":
                    var matchesSpace = Regex.Matches(text, @"(?:^|\s)(\d+(?:\.\d+)?)(?:\s|$)");
                    foreach (Match m in matchesSpace)
                    {
                        if (m.Success && m.Groups.Count > 1)
                        {
                            list.Add(m.Groups[1].Value.Trim());
                        }
                    }
                    break;

                // 顺位模式 4: 寻找末尾纯数字 (如 NM1-125 100)
                case "TrailingNumber":
                    var matchTail = Regex.Match(text, @"(?i)(?:^|\s|-|_|/)(\d+(?:\.\d+)?)\s*$");
                    if (matchTail.Success && matchTail.Groups.Count > 1)
                    {
                        list.Add(matchTail.Groups[1].Value.Trim());
                    }
                    break;

                // 顺位模式 5: 固定电流值 (如固定 100A)
                case "FixedValue":
                    if (!string.IsNullOrWhiteSpace(rule.FixedValue))
                    {
                        list.Add(rule.FixedValue.Trim());
                    }
                    break;

                // 顺位模式 6: 自定义正则表达式 (支持 (?<current>...) 命名捕获组或首个括号捕获)
                case "CustomRegex":
                    if (!string.IsNullOrWhiteSpace(rule.CustomRegex))
                    {
                        try
                        {
                            var matchesCustom = Regex.Matches(text, rule.CustomRegex);
                            foreach (Match m in matchesCustom)
                            {
                                if (!m.Success) continue;
                                if (m.Groups["current"].Success)
                                {
                                    list.Add(m.Groups["current"].Value.Trim());
                                }
                                else if (m.Groups.Count > 1)
                                {
                                    list.Add(m.Groups[1].Value.Trim());
                                }
                                else
                                {
                                    list.Add(m.Value.Trim());
                                }
                            }
                        }
                        catch { }
                    }
                    break;
            }

            return list;
        }

        /// <summary>
        /// 从正则匹配到的候选文本中提取纯数字；若无法直接转为 int，则固定提取其中包含的数字；若是区间则展开端点
        /// </summary>
        /// <param name="rawMatch">匹配到的原始文本 (如 200A, C32, 2.5-4A)</param>
        /// <returns>提取出的浮点数值集合</returns>
        private static List<double> ExtractPureNumbersFromMatch(string rawMatch)
        {
            var results = new List<double>();
            if (string.IsNullOrWhiteSpace(rawMatch)) return results;

            string clean = rawMatch.Trim();

            // 1. 先检查是否为区间格式 (如 2.5-4A, 6~10A)
            var rangeMatch = Regex.Match(clean, @"(\d+(?:\.\d+)?)\s*[-~～至]\s*(\d+(?:\.\d+)?)");
            if (rangeMatch.Success && rangeMatch.Groups.Count > 2)
            {
                if (double.TryParse(rangeMatch.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double lowVal))
                {
                    results.Add(lowVal);
                }
                if (double.TryParse(rangeMatch.Groups[2].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double highVal))
                {
                    results.Add(highVal);
                }
                return results;
            }

            // 2. 检查是否可直接转为整数 int
            if (int.TryParse(clean, out int intVal))
            {
                results.Add(intVal);
                return results;
            }

            // 3. 检查是否可直接转为浮点数 double
            if (double.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleVal))
            {
                results.Add(doubleVal);
                return results;
            }

            // 4. 若无法转为 int/double (如匹配到 200A, C32 等非纯数字)，固定正则提取其中的纯数字
            var numMatches = Regex.Matches(clean, @"\d+(?:\.\d+)?");
            foreach (Match nm in numMatches)
            {
                if (nm.Success && double.TryParse(nm.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedVal))
                {
                    results.Add(parsedVal);
                }
            }

            return results;
        }

        /// <summary>
        /// 校验极数是否命中白名单
        /// </summary>
        private static bool ValidatePoleWhitelist(string pole, ModelParserConfig config)
        {
            // 若未启用严格白名单，直接通过
            if (!config.EnableStrictPoleWhitelist) return true;
            if (string.IsNullOrWhiteSpace(pole)) return false;

            // 规范化待校验极数字符串
            string cleanPole = pole.Trim().ToUpper();
            string numOnly = cleanPole.Replace("P", "").Trim();

            // 遍历白名单项校验
            if (config.PoleAllowedValues != null)
            {
                foreach (var item in config.PoleAllowedValues)
                {
                    if (string.IsNullOrWhiteSpace(item)) continue;
                    string standardItem = item.Trim().ToUpper();
                    if (cleanPole == standardItem || numOnly == standardItem)
                    {
                        return true;
                    }
                }
            }

            // 白名单未命中
            return false;
        }

        /// <summary>
        /// 校验电流是否命中标称白名单 (如 99 不在白名单内返回 false)
        /// </summary>
        private static bool ValidateCurrentWhitelist(string current, ModelParserConfig config)
        {
            // 若未启用严格白名单，直接通过
            if (!config.EnableStrictCurrentWhitelist) return true;
            if (string.IsNullOrWhiteSpace(current)) return false;

            // 若为区间型电流 (如 2.5-4)，视为合法
            if (current.Contains("-") || current.Contains("~") || current.Contains("～"))
            {
                return true;
            }

            // 尝试将电流转为浮点数
            if (double.TryParse(current, NumberStyles.Any, CultureInfo.InvariantCulture, out double curVal))
            {
                if (config.CurrentAllowedValues != null)
                {
                    foreach (var item in config.CurrentAllowedValues)
                    {
                        if (double.TryParse(item, NumberStyles.Any, CultureInfo.InvariantCulture, out double standardVal))
                        {
                            // 浮点数近零差值比较
                            if (Math.Abs(curVal - standardVal) < 0.001)
                            {
                                return true;
                            }
                        }
                    }
                }
            }

            // 未通过白名单校验
            return false;
        }

        /// <summary>
        /// 格式化极数输出 (根据配置决定是否带 P，保护 1P+N 等复合极数)
        /// </summary>
        private static string FormatPoleOutput(string pole, string format)
        {
            if (string.IsNullOrWhiteSpace(pole)) return string.Empty;
            string upper = pole.Trim().ToUpper();

            // 若为 1P+N, 3P+N 等包含 + 或 N 的复合极数，保持原样输出，避免被截断
            if (upper.Contains("+") || upper.Contains("N"))
            {
                return upper;
            }

            if (format == "NumberOnly")
            {
                return upper.Replace("P", "").Trim();
            }
            else
            {
                return upper.EndsWith("P") ? upper : upper + "P";
            }
        }

        /// <summary>
        /// 格式化电流输出 (根据配置决定是否带 A)
        /// </summary>
        private static string FormatCurrentOutput(string current, string format)
        {
            if (string.IsNullOrWhiteSpace(current)) return string.Empty;
            string clean = current.Trim();

            if (format == "WithA")
            {
                return clean.EndsWith("A", StringComparison.OrdinalIgnoreCase) ? clean : clean + "A";
            }
            else
            {
                return clean.EndsWith("A", StringComparison.OrdinalIgnoreCase)
                    ? clean.Substring(0, clean.Length - 1).Trim()
                    : clean;
            }
        }

        /// <summary>
        /// 屏蔽脱扣方式前置排除关键词
        /// </summary>
        private static string MaskTripModeExclusions(string input, List<string>? exclusions)
        {
            if (string.IsNullOrWhiteSpace(input) || exclusions == null || exclusions.Count == 0)
                return input;

            string result = input;
            foreach (var word in exclusions)
            {
                if (string.IsNullOrWhiteSpace(word)) continue;
                string pattern = $@"(?i)\b{Regex.Escape(word.Trim())}\b";
                result = Regex.Replace(result, pattern, MASK_PLACEHOLDER);
            }
            return result;
        }

        /// <summary>
        /// 根据指定的顺位规则尝试提取脱扣方式简写代号
        /// </summary>
        private static string TryExtractTripModeByRule(string text, PipelineRuleItem rule)
        {
            switch (rule.Mode)
            {
                // 顺位模式 1 & 2: 代号与特征对照表映射 (如 /3300->TM, TMD->TMD)，区分大小写
                case "CodeMapping":
                    if (rule.Mapping != null)
                    {
                        foreach (var kvp in rule.Mapping)
                        {
                            // 严格区分大小写匹配代号关键词
                            if (!string.IsNullOrWhiteSpace(kvp.Key) &&
                                text.IndexOf(kvp.Key, StringComparison.Ordinal) >= 0)
                            {
                                return kvp.Value?.Trim() ?? string.Empty;
                            }
                        }
                    }
                    break;

                // 顺位模式 3: 微断脱扣特性曲线提取 (C/D/B/K/Z)
                case "CurveLetter":
                    var matchCurve = Regex.Match(text, @"(?i)(?:^|[^a-zA-Z0-9])([CDBKZ])\d+(?:[^\d\.]|$)");
                    if (matchCurve.Success && matchCurve.Groups.Count > 1)
                    {
                        return matchCurve.Groups[1].Value.ToUpper();
                    }
                    break;

                // 顺位模式: 固定脱扣方式简写
                case "FixedValue":
                    return rule.FixedValue?.Trim() ?? string.Empty;

                // 顺位模式: 自定义正则表达式提取
                case "CustomRegex":
                    if (!string.IsNullOrWhiteSpace(rule.CustomRegex))
                    {
                        try
                        {
                            var matchCustom = Regex.Match(text, rule.CustomRegex);
                            if (matchCustom.Success)
                            {
                                if (matchCustom.Groups["trip"].Success)
                                {
                                    return matchCustom.Groups["trip"].Value.Trim();
                                }
                                if (matchCustom.Groups.Count > 1)
                                {
                                    return matchCustom.Groups[1].Value.Trim();
                                }
                                return matchCustom.Value.Trim();
                            }
                        }
                        catch { }
                    }
                    break;
            }

            return string.Empty;
        }

        /// <summary>
        /// 校验脱扣方式简写代号是否命中白名单
        /// </summary>
        private static bool ValidateTripModeWhitelist(string tripMode, ModelParserConfig config)
        {
            // 若未启用严格白名单校验，直接通过
            if (!config.EnableStrictTripModeWhitelist) return true;
            if (string.IsNullOrWhiteSpace(tripMode)) return false;

            string clean = tripMode.Trim().ToUpper();
            if (config.TripModeAllowedValues != null)
            {
                foreach (var item in config.TripModeAllowedValues)
                {
                    if (string.IsNullOrWhiteSpace(item)) continue;
                    if (string.Equals(clean, item.Trim().ToUpper(), StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// 将 COM 读取的单个对象或二维数组统一转换为标准 object[,] 二维数组
        /// </summary>
        private static object[,] ConvertTo2DArray(object? value, int expectedRows)
        {
            // 如果本来就是二维数组，直接强制转换
            if (value is object[,] arr2D)
            {
                return arr2D;
            }

            // 如果是单个单元格对象，构造 1x1 二维数组
            object[,] result = new object[expectedRows + 1, 2];
            result[1, 1] = value ?? string.Empty;
            return result;
        }
    }
}
