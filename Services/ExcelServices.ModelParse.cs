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
                // 若窗体已存在且未被销毁，则直接激活并带到最前
                if (_modelParamParserForm != null && !_modelParamParserForm.IsDisposed)
                {
                    _modelParamParserForm.WindowState = System.Windows.Forms.FormWindowState.Normal;
                    _modelParamParserForm.Activate();
                    return;
                }

                // 实例化新窗体
                _modelParamParserForm = new ModelParamParserForm();
                // 窗体关闭时清空引用
                _modelParamParserForm.FormClosed += (s, e) => _modelParamParserForm = null;
                // 显示窗体
                _modelParamParserForm.Show();
            }
            catch (Exception ex)
            {
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
        /// 电流解析核心逻辑: 经历前置排除过滤 -> 多级顺位流水线 -> 后置标称白名单校验
        /// </summary>
        /// <param name="rawModel">原始型号文本</param>
        /// <param name="config">配置对象</param>
        /// <param name="hitRuleTitle">命中的规则标题</param>
        /// <returns>提取出的电流结果（如 100 或 100A）</returns>
        public static string ParseCurrent(string rawModel, ModelParserConfig config, out string hitRuleTitle)
        {
            // 初始化命中规则输出变量
            hitRuleTitle = string.Empty;
            // 校验输入字符串有效性
            if (string.IsNullOrWhiteSpace(rawModel)) return string.Empty;

            // 1. 前置必去项与负向排除过滤 (如将 100mA, 10kA, 220V 掩码，避免误提取为额定电流)
            string cleanText = MaskCurrentExclusions(rawModel, config.CurrentExcludeKeywords);

            // 2. 依次遍历用户配置的电流顺位流水线规则
            if (config.CurrentPipeline != null)
            {
                foreach (var rule in config.CurrentPipeline)
                {
                    // 跳过未启用的顺位规则
                    if (!rule.Enabled) continue;

                    // 尝试使用当前规则模式提取电流
                    string extracted = TryExtractCurrentByRule(cleanText, rule);

                    // 判断是否提取到了候选电流值
                    if (!string.IsNullOrWhiteSpace(extracted))
                    {
                        // 3. 后置标称范围白名单校验 (如过滤 99 等非标数值)
                        if (ValidateCurrentWhitelist(extracted, config))
                        {
                            // 记录命中的顺位规则名称
                            hitRuleTitle = rule.Title;
                            // 格式化输出 (100 或 100A)
                            return FormatCurrentOutput(extracted, config.CurrentFormat);
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

            // 执行极数双通道流水线解析
            result.Pole = ParsePoles(rawModel ?? string.Empty, config, out string hitPole);
            result.HitPoleRule = hitPole;

            // 执行电流双通道流水线解析
            result.Current = ParseCurrent(rawModel ?? string.Empty, config, out string hitCurrent);
            result.HitCurrentRule = hitCurrent;

            // 综合评估解析结果状态
            bool hasPole = !string.IsNullOrWhiteSpace(result.Pole);
            bool hasCur = !string.IsNullOrWhiteSpace(result.Current);

            // 两者均成功提取记为 Success
            if (hasPole && hasCur)
            {
                result.Status = "Success";
            }
            // 仅提取到其中一项记为 Partial
            else if (hasPole || hasCur)
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
                string colCur = string.IsNullOrWhiteSpace(config.CurrentColumn) ? "S" : config.CurrentColumn.Trim().ToUpper();
                string colPole = string.IsNullOrWhiteSpace(config.PoleColumn) ? "T" : config.PoleColumn.Trim().ToUpper();
                // 确保起始行号不小于 1
                int startRow = config.StartRow < 1 ? 2 : config.StartRow;

                // 计算当前工作表中已使用区域的最大物理行数
                int maxUsedRow = 0;
                try
                {
                    dynamic usedRange = activeSheet.UsedRange;
                    maxUsedRow = (int)(usedRange.Row + usedRange.Rows.Count - 1);
                }
                catch
                {
                    maxUsedRow = startRow + 100;
                }

                // 若有效行数小于起始行，说明表为空
                if (maxUsedRow < startRow)
                {
                    batchResult.Success = true;
                    batchResult.TotalRows = 0;
                    batchResult.Message = "指定起始行下方无有效数据";
                    return batchResult;
                }

                // 计算待处理的总行数
                int rowCount = maxUsedRow - startRow + 1;
                batchResult.TotalRows = rowCount;

                // ==================== 1. 一次性将源型号整列读入二维数组 ====================
                dynamic srcRange = activeSheet.Range[$"{colSrc}{startRow}:{colSrc}{maxUsedRow}"];
                object[,] rawArray = ConvertTo2DArray(srcRange.Value2, rowCount);

                // ==================== 2. 在内存中创建电流与极数的目标二维数组 ====================
                object[,] currentArray = new object[rowCount, 1];
                object[,] poleArray = new object[rowCount, 1];

                // 计数器统计识别情况
                int successCount = 0;
                int failedCount = 0;

                // 循环遍历内存中的二维数组行
                for (int i = 0; i < rowCount; i++)
                {
                    // 获取当前行的原始型号文本
                    object val = rawArray[i + 1, 1];
                    string rawModel = val?.ToString()?.Trim() ?? string.Empty;

                    // 若本行为空行，则保持空值
                    if (string.IsNullOrWhiteSpace(rawModel))
                    {
                        currentArray[i, 0] = string.Empty;
                        poleArray[i, 0] = string.Empty;
                        continue;
                    }

                    // 调用双通道流水线解析极数与电流
                    string pole = ParsePoles(rawModel, config, out _);
                    string current = ParseCurrent(rawModel, config, out _);

                    // 写入内存结果数组
                    currentArray[i, 0] = current;
                    poleArray[i, 0] = pole;

                    // 统计成功率
                    if (!string.IsNullOrWhiteSpace(pole) && !string.IsNullOrWhiteSpace(current))
                    {
                        successCount++;
                    }
                    else
                    {
                        failedCount++;
                    }
                }

                // ==================== 3. 一次性将二维数组整块写入目标列 ====================
                dynamic curRange = activeSheet.Range[$"{colCur}{startRow}:{colCur}{maxUsedRow}"];
                curRange.Value2 = currentArray;

                dynamic poleRange = activeSheet.Range[$"{colPole}{startRow}:{colPole}{maxUsedRow}"];
                poleRange.Value2 = poleArray;

                // 停止计时并汇总结果
                stopwatch.Stop();
                batchResult.Success = true;
                batchResult.SuccessCount = successCount;
                batchResult.FailedCount = failedCount;
                batchResult.ElapsedMilliseconds = stopwatch.ElapsedMilliseconds;
                batchResult.Message = $"批量处理完成：共处理 {rowCount} 行，完全识别 {successCount} 行，耗时 {stopwatch.ElapsedMilliseconds} ms";
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

                // 顺位模式 2: 脱扣代号或特殊特征对照表映射 (如 /3300 -> 3P)
                case "CodeMapping":
                    if (rule.Mapping != null)
                    {
                        foreach (var kvp in rule.Mapping)
                        {
                            if (!string.IsNullOrWhiteSpace(kvp.Key) &&
                                text.IndexOf(kvp.Key, StringComparison.OrdinalIgnoreCase) >= 0)
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
        /// 根据指定的顺位规则尝试提取电流
        /// </summary>
        private static string TryExtractCurrentByRule(string text, PipelineRuleItem rule)
        {
            switch (rule.Mode)
            {
                // 顺位模式 1: 寻找带单位 "A" 的数字 (如 100A, 63A, 0.03A)
                case "NumberWithA":
                    var matchA = Regex.Match(text, @"(?i)(?:^|[^a-zA-Z0-9\.])(\d+(?:\.\d+)?)\s*A(?:[^a-zA-Z0-9]|$)");
                    if (matchA.Success)
                    {
                        return matchA.Groups[1].Value;
                    }
                    break;

                // 顺位模式 2: 寻找脱扣曲线字母 [C/D/K/Z] 后面的数字 (如 C32, D16)
                case "CurveLetterNumber":
                    var matchCurve = Regex.Match(text, @"(?i)(?:^|[^a-zA-Z0-9])[CDKZ](\d+(?:\.\d+)?)(?:[^\d\.]|$)");
                    if (matchCurve.Success)
                    {
                        return matchCurve.Groups[1].Value;
                    }
                    break;

                // 顺位模式 3: 寻找整定电流可调区间 (如 2.5-4A, 0.63-1A)
                case "CurrentRange":
                    var matchRange = Regex.Match(text, @"(?i)(\d+(?:\.\d+)?\s*[-~～至]\s*\d+(?:\.\d+)?)\s*A?");
                    if (matchRange.Success)
                    {
                        return matchRange.Groups[1].Value.Replace(" ", "");
                    }
                    break;

                // 顺位模式: 空格+数字+空格 (如独立的纯数字电流 " 100 " 或 " 63 ")
                case "SpaceNumberSpace":
                    var matchSpaceCur = Regex.Match(text, @"(?:^|\s)(\d+(?:\.\d+)?)(?:\s|$)");
                    if (matchSpaceCur.Success)
                    {
                        return matchSpaceCur.Groups[1].Value;
                    }
                    break;

                // 顺位模式 4: 寻找末尾纯数字 (如 NM1-125 100)
                case "TrailingNumber":
                    var matchTail = Regex.Match(text, @"(?i)(?:^|\s|-|_|/)(\d+(?:\.\d+)?)\s*$");
                    if (matchTail.Success)
                    {
                        return matchTail.Groups[1].Value;
                    }
                    break;

                // 顺位模式 5: 固定电流值 (如固定 100A)
                case "FixedValue":
                    return rule.FixedValue?.Trim() ?? string.Empty;

                // 顺位模式 6: 自定义正则表达式 (支持 (?<current>...) 命名捕获组或首个括号捕获)
                case "CustomRegex":
                    if (!string.IsNullOrWhiteSpace(rule.CustomRegex))
                    {
                        try
                        {
                            var matchCustomCur = Regex.Match(text, rule.CustomRegex);
                            if (matchCustomCur.Success)
                            {
                                if (matchCustomCur.Groups["current"].Success)
                                {
                                    return matchCustomCur.Groups["current"].Value.Trim();
                                }
                                if (matchCustomCur.Groups.Count > 1)
                                {
                                    return matchCustomCur.Groups[1].Value.Trim();
                                }
                                return matchCustomCur.Value.Trim();
                            }
                        }
                        catch { }
                    }
                    break;
            }

            return string.Empty;
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
