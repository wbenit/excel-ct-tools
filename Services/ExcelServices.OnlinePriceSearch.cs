using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;
using ExcelDna.Integration;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 在线查价与批量静默回写业务服务层 (严格遵守规则 2、3、4、7)
    /// </summary>
    public static partial class ExcelServices
    {
        // 剥离括号及其内部参数的正则表达式（兼容中英文圆括号）
        private static readonly Regex ParenthesesRegex = new Regex(@"[\(（][^\)）]*[\)）]", RegexOptions.Compiled);

        /// <summary>
        /// 智能清理并提取纯净的检索规格型号
        /// </summary>
        /// <param name="raw">单元格原始文字</param>
        /// <param name="trimParentheses">是否剥离括号参数</param>
        /// <returns>清洗后的规格型号</returns>
        public static string CleanSearchModel(string raw, bool trimParentheses)
        {
            // 空值保护
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;

            // 清理首尾空格与制表符
            string cleaned = raw.Trim();
            // 若开启剥离括号
            if (trimParentheses)
            {
                // 去除所有中英文括号及其内部文字 (例如 "NM1-125S/3300 (100A)" -> "NM1-125S/3300")
                cleaned = ParenthesesRegex.Replace(cleaned, "");
            }

            // 再次去除多余首尾空白并返回
            return cleaned.Trim();
        }

        /// <summary>
        /// 判定当前行是否属于需要过滤的保留行 (表头、小计、合计等)
        /// </summary>
        private static bool IsIgnoredRow(string text)
        {
            // 空文本不单独作为保留特征
            if (string.IsNullOrWhiteSpace(text)) return false;

            // 包含常见汇总统计关键字的行进行排除 --硬编码: 保留行关键字--
            return text.Contains("小计") ||
                   text.Contains("合计") ||
                   text.Contains("总计") ||
                   text.Contains("箱柜") ||
                   text.Contains("名称") ||
                   text.Contains("型号") ||
                   text.Contains("单台合计");
        }

        /// <summary>
        /// 快速探测当前 Excel 选区概况并返回预览数据
        /// </summary>
        /// <param name="trimParentheses">是否剥离括号</param>
        /// <returns>选区探测结果包</returns>
        public static SelectionDetectDto DetectSelectionForPriceSearch(bool trimParentheses = true)
        {
            // 初始化返回对象
            var result = new SelectionDetectDto();
            try
            {
                // 获取 Excel 顶级应用程序接口
                dynamic? app = ExcelDnaUtil.Application;
                if (app == null) return result;

                // 获取当前活动工作表
                dynamic? activeSheet = app.ActiveSheet;
                // 获取当前活动选区
                dynamic? selection = app.Selection;
                if (activeSheet == null || selection == null) return result;

                // 记录工作表名称
                result.SheetName = activeSheet.Name?.ToString() ?? string.Empty;

                // 获取选区行数与列数
                int rowCount = selection.Rows.Count;
                int colCount = selection.Columns.Count;
                int startRow = selection.Row;
                int startCol = selection.Column;

                // 记录起始与终止行
                result.StartRow = startRow;
                result.EndRow = startRow + rowCount - 1;

                // 规则 7: 二维数组一次性读取选区内存快照
                object[,] rawValues;
                if (rowCount == 1 && colCount == 1)
                {
                    // 单个单元格时适配为 1x1 二维矩阵
                    rawValues = new object[1, 1];
                    rawValues[0, 0] = selection.Value2;
                }
                else
                {
                    // 多单元格转换为标准二维数组
                    rawValues = (object[,])selection.Value2;
                }

                // 确定选区中哪个相对列代表型号
                // 默认取第 1 列；若选区跨越多列且当前包含了 C 列 (Column 3)，则优先取 C 列
                int targetRelativeCol = 1;
                if (colCount > 1)
                {
                    // 计算 C 列在选区中的相对列索引 (C列绝对列号为 3)
                    int cColRel = 3 - startCol + 1;
                    if (cColRel >= 1 && cColRel <= colCount)
                    {
                        targetRelativeCol = cColRel;
                    }
                }

                // 遍历内存二维数组识别有效元器件行
                int validCount = 0;
                var previewList = new List<string>();

                for (int r = 1; r <= rowCount; r++)
                {
                    // 提取目标单元格文本
                    string cellText = rawValues[r, targetRelativeCol]?.ToString() ?? string.Empty;
                    // 若为空或为排除行则跳过
                    if (string.IsNullOrWhiteSpace(cellText) || IsIgnoredRow(cellText)) continue;

                    // 清洗提取纯型号
                    string model = CleanSearchModel(cellText, trimParentheses);
                    // 型号长度至少大于等于 2 个字符
                    if (model.Length >= 2)
                    {
                        validCount++;
                        // 仅收录前 5 个作为预览
                        if (previewList.Count < 5)
                        {
                            previewList.Add($"行 {startRow + r - 1}: {model}");
                        }
                    }
                }

                // 组装最终结果
                result.HasSelection = validCount > 0;
                result.ValidRowCount = validCount;
                result.PreviewModels = previewList;
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"DetectCurrentSelection 选区探测异常: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 从当前 Excel 选区中一次性提取待查价的元器件条目列表 (遵循规则 7)
        /// </summary>
        /// <param name="trimParentheses">是否剥离括号参数</param>
        /// <returns>待查价项集合</returns>
        public static List<OnlinePriceItemDto> GetSelectedItemsForSearch(bool trimParentheses = true)
        {
            // 初始化条目列表
            var items = new List<OnlinePriceItemDto>();
            try
            {
                // 获取 Excel 宿主句柄
                dynamic? app = ExcelDnaUtil.Application;
                if (app == null) return items;

                // 获取当前活动选区
                dynamic? selection = app.Selection;
                if (selection == null) return items;

                // 选区几何参数
                int rowCount = selection.Rows.Count;
                int colCount = selection.Columns.Count;
                int startRow = selection.Row;
                int startCol = selection.Column;

                // 规则 7: 二维数组一次性读入内存
                object[,] rawValues;
                if (rowCount == 1 && colCount == 1)
                {
                    // 单格封装
                    rawValues = new object[1, 1];
                    rawValues[0, 0] = selection.Value2;
                }
                else
                {
                    // 矩阵强转
                    rawValues = (object[,])selection.Value2;
                }

                // 识别型号所在的相对列
                int targetRelativeCol = 1;
                if (colCount > 1)
                {
                    // 若选区跨越了 C 列 (Col 3)，优先取 C 列
                    int cColRel = 3 - startCol + 1;
                    if (cColRel >= 1 && cColRel <= colCount)
                    {
                        targetRelativeCol = cColRel;
                    }
                }

                // 遍历二维矩阵生成待查传输项
                for (int r = 1; r <= rowCount; r++)
                {
                    // 物理行号
                    int physicalRow = startRow + r - 1;
                    // 读取单元格文字
                    string cellText = rawValues[r, targetRelativeCol]?.ToString() ?? string.Empty;

                    // 过滤空行或表头合计
                    if (string.IsNullOrWhiteSpace(cellText) || IsIgnoredRow(cellText)) continue;

                    // 清洗并提取有效型号
                    string searchModel = CleanSearchModel(cellText, trimParentheses);
                    // 字符长度过滤
                    if (searchModel.Length < 2) continue;

                    // 加入集合
                    items.Add(new OnlinePriceItemDto
                    {
                        Row = physicalRow,
                        RawText = cellText,
                        SearchModel = searchModel,
                        Status = "Pending"
                    });
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"GetSelectedItemsForSearch 提取选区条目异常: {ex.Message}");
            }

            return items;
        }

        /// <summary>
        /// 规范化用户选择或手填的 Excel 列字母标识符 (如 "C", "C 列", "d" 统一转为大写 "C", "D")
        /// </summary>
        private static string NormalizeColLetter(string? rawCol)
        {
            // 空值保护处理
            if (string.IsNullOrWhiteSpace(rawCol)) return string.Empty;
            // 转换为全大写并裁剪空格
            string trimmed = rawCol.Trim().ToUpper();
            // 若为显式不回填标记
            if (trimmed == "NONE" || trimmed.Contains("不回填")) return "NONE";
            // 正则匹配提取纯英文字母序列 (支持单字母 C 或双字母 AA 等)
            var match = Regex.Match(trimmed, @"[A-Z]+");
            // 匹配成功返回提取大写字母，失败返回空
            return match.Success ? match.Value : string.Empty;
        }

        /// <summary>
        /// 将在线查价结果批量静默写回 Excel (严格遵循规则 7，支持用户自选/自填列与撤销栈)
        /// </summary>
        /// <param name="items">已完成查价的结果列表</param>
        /// <param name="config">查价与回填列配置</param>
        /// <returns>回写执行状态元组 (成功状态, 更新行数, 提示消息)</returns>
        public static (bool success, int updatedCount, string message) BatchWriteBackPrices(
            List<OnlinePriceItemDto> items,
            OnlinePriceConfig config)
        {
            // 过滤出查价成功且有有效单价的数据项
            var successItems = items?
                .Where(i => i.Status == "Success" && i.Price > 0)
                .OrderBy(i => i.Row)
                .ToList();

            // 若无有效数据
            if (successItems == null || successItems.Count == 0)
            {
                return (false, 0, "没有成功查得价格的有效元器件项");
            }

            // 获取 Excel 宿主对象
            dynamic? app = ExcelDnaUtil.Application;
            if (app == null) return (false, 0, "无法获取 Excel 应用程序实例");

            // 获取当前活动工作表
            dynamic? activeSheet = app.ActiveSheet;
            if (activeSheet == null) return (false, 0, "当前无有效活动工作表");

            // 原始状态记录
            bool oldScreenUpdating = true;
            bool oldEnableEvents = true;

            try
            {
                // 记录原始重绘与事件开关
                oldScreenUpdating = app.ScreenUpdating;
                oldEnableEvents = app.EnableEvents;

                // 挂起屏幕重绘与自动事件，防止卡顿与事件雪崩
                app.ScreenUpdating = false;
                app.EnableEvents = false;

                // 解析目标价格列字母 (可选手填，默认 "M")
                string priceColLetter = NormalizeColLetter(config.PriceTargetCol);
                if (string.IsNullOrWhiteSpace(priceColLetter) || priceColLetter == "NONE") priceColLetter = "M";

                // 解析目标型号列字母 (可选手填，若为 NONE 或不回填则不回写型号)
                string modelColLetter = NormalizeColLetter(config.ModelTargetCol);
                bool needWriteModel = !config.OnlyUpdatePrice && !string.IsNullOrWhiteSpace(modelColLetter) && modelColLetter != "NONE";

                // 解析目标品牌厂商列字母 (可选手填，新增支持品牌回写)
                string brandColLetter = NormalizeColLetter(config.BrandTargetCol);
                bool needWriteBrand = !string.IsNullOrWhiteSpace(brandColLetter) && brandColLetter != "NONE";

                // 计算受影响的行范围 (起始行与终止行)
                int minRow = successItems.Min(i => i.Row);
                int maxRow = successItems.Max(i => i.Row);
                int totalSpanRows = maxRow - minRow + 1;

                // 构造行映射字典，便于快速索引
                var rowItemMap = successItems.ToDictionary(i => i.Row, i => i);

                // 准备撤销命令切片集合
                var undoSlices = new List<RangeDeltaSlice>();

                // ==================== 1. 批量回写价格列 (规则 7: 二维数组整块写入) ====================
                if (!string.IsNullOrWhiteSpace(priceColLetter))
                {
                    // 锁定目标区域标准 A1 表达式
                    string priceRangeAddr = $"{priceColLetter}{minRow}:{priceColLetter}{maxRow}";
                    dynamic priceRange = activeSheet.Range[priceRangeAddr];

                    // 一次性读取原有二维数组快照
                    object[,] oldPriceMatrix;
                    if (totalSpanRows == 1)
                    {
                        oldPriceMatrix = new object[1, 1];
                        oldPriceMatrix[0, 0] = priceRange.Value2;
                    }
                    else
                    {
                        oldPriceMatrix = (object[,])priceRange.Value2;
                    }

                    // 实例化新的二维数据矩阵
                    object[,] newPriceMatrix = new object[totalSpanRows, 1];

                    // 复制原有数据并填充新查得的单价
                    for (int r = 1; r <= totalSpanRows; r++)
                    {
                        int currentRow = minRow + r - 1;
                        // 若当前行在命中查价结果中
                        if (rowItemMap.TryGetValue(currentRow, out var matchDto))
                        {
                            // 填入最新单价数值 (保留2位小数)
                            newPriceMatrix[r - 1, 0] = Math.Round(matchDto.Price, 2);
                        }
                        else
                        {
                            // 保持原样不作变动
                            newPriceMatrix[r - 1, 0] = oldPriceMatrix[r, 1];
                        }
                    }

                    // 记录撤销快照切片
                    undoSlices.Add(new RangeDeltaSlice
                    {
                        SheetName = activeSheet.Name,
                        RangeAddress = priceRangeAddr,
                        OldValues = oldPriceMatrix,
                        NewValues = newPriceMatrix
                    });

                    // 规则 7: 单次 COM 调用将新矩阵批量写回 Excel
                    priceRange.Value2 = newPriceMatrix;
                }

                // ==================== 2. 批量回写型号列 (规则 7: 若配置启用) ====================
                if (needWriteModel)
                {
                    // 锁定型号区域 A1 表达式
                    string modelRangeAddr = $"{modelColLetter}{minRow}:{modelColLetter}{maxRow}";
                    dynamic modelRange = activeSheet.Range[modelRangeAddr];

                    // 一次性读取型号原有快照
                    object[,] oldModelMatrix;
                    if (totalSpanRows == 1)
                    {
                        oldModelMatrix = new object[1, 1];
                        oldModelMatrix[0, 0] = modelRange.Value2;
                    }
                    else
                    {
                        oldModelMatrix = (object[,])modelRange.Value2;
                    }

                    // 实例化新的型号二维数据矩阵
                    object[,] newModelMatrix = new object[totalSpanRows, 1];

                    // 复制原有数据并填充官方标准型号
                    for (int r = 1; r <= totalSpanRows; r++)
                    {
                        int currentRow = minRow + r - 1;
                        if (rowItemMap.TryGetValue(currentRow, out var matchDto) && !string.IsNullOrWhiteSpace(matchDto.MatchedModel))
                        {
                            // 回填官方匹配标准型号
                            newModelMatrix[r - 1, 0] = matchDto.MatchedModel;
                        }
                        else
                        {
                            // 保持原样
                            newModelMatrix[r - 1, 0] = oldModelMatrix[r, 1];
                        }
                    }

                    // 记录撤销快照切片
                    undoSlices.Add(new RangeDeltaSlice
                    {
                        SheetName = activeSheet.Name,
                        RangeAddress = modelRangeAddr,
                        OldValues = oldModelMatrix,
                        NewValues = newModelMatrix
                    });

                    // 规则 7: 批量写回型号
                    modelRange.Value2 = newModelMatrix;
                }

                // ==================== 3. 批量回写品牌厂商列 (规则 7: 若配置启用) ====================
                if (needWriteBrand)
                {
                    // 锁定品牌区域 A1 表达式
                    string brandRangeAddr = $"{brandColLetter}{minRow}:{brandColLetter}{maxRow}";
                    dynamic brandRange = activeSheet.Range[brandRangeAddr];

                    // 一次性读取品牌原有快照
                    object[,] oldBrandMatrix;
                    if (totalSpanRows == 1)
                    {
                        oldBrandMatrix = new object[1, 1];
                        oldBrandMatrix[0, 0] = brandRange.Value2;
                    }
                    else
                    {
                        oldBrandMatrix = (object[,])brandRange.Value2;
                    }

                    // 实例化新的品牌二维数据矩阵
                    object[,] newBrandMatrix = new object[totalSpanRows, 1];

                    // 复制原有数据并填充查得的厂商或指定品牌
                    for (int r = 1; r <= totalSpanRows; r++)
                    {
                        int currentRow = minRow + r - 1;
                        if (rowItemMap.TryGetValue(currentRow, out var matchDto))
                        {
                            // 优先回写返回的实际品牌名称 (排除“全部”占位符)
                            string brandVal = !string.IsNullOrWhiteSpace(matchDto.Vendor) && matchDto.Vendor != "全部" ? matchDto.Vendor : string.Empty;
                            // 若厂商字段为空，次之取用户界面指定的明确品牌
                            if (string.IsNullOrWhiteSpace(brandVal) && config.Brand != "全部")
                            {
                                brandVal = config.Brand;
                            }
                            // 若依然为空，从匹配官方型号或检索型号特征前缀智能推导实际品牌
                            if (string.IsNullOrWhiteSpace(brandVal))
                            {
                                brandVal = OnlinePriceSearchClient.DeduceBrandFromModel(!string.IsNullOrWhiteSpace(matchDto.MatchedModel) ? matchDto.MatchedModel : matchDto.SearchModel);
                            }
                            // 最终若提炼到实际品牌则批量更新，若完全无法识别则保留原单元格内容
                            newBrandMatrix[r - 1, 0] = !string.IsNullOrWhiteSpace(brandVal) ? brandVal : oldBrandMatrix[r, 1];
                        }
                        else
                        {
                            // 保持原样
                            newBrandMatrix[r - 1, 0] = oldBrandMatrix[r, 1];
                        }
                    }

                    // 记录撤销快照切片
                    undoSlices.Add(new RangeDeltaSlice
                    {
                        SheetName = activeSheet.Name,
                        RangeAddress = brandRangeAddr,
                        OldValues = oldBrandMatrix,
                        NewValues = newBrandMatrix
                    });

                    // 规则 7: 批量写回品牌厂商列
                    brandRange.Value2 = newBrandMatrix;
                }

                // ==================== 3. 注册撤销栈命令 ====================
                if (undoSlices.Count > 0)
                {
                    // 压入撤销栈
                    var cmd = new RangeDeltaCommand($"在线查价批量回写 ({successItems.Count} 项)", undoSlices);
                    UndoRedoManager.Instance.PushCommand(cmd);
                }

                // 成功返回更新统计
                return (true, successItems.Count, $"成功更新 {successItems.Count} 个元器件价格");
            }
            catch (Exception ex)
            {
                // 记录错误日志
                LogHelper.WriteLog($"BatchWriteBackPrices 回写异常: {ex.Message}");
                return (false, 0, $"回写 Excel 异常: {ex.Message}");
            }
            finally
            {
                // 严格恢复屏幕重绘与全局事件
                try
                {
                    app.ScreenUpdating = oldScreenUpdating;
                    app.EnableEvents = oldEnableEvents;
                }
                catch { }
            }
        }

        /// <summary>
        /// 弹出基于 WebView2 + Vue 3 的在线查价与静默回写窗口 (非模态置顶)
        /// </summary>
        public static void ShowOnlinePriceSearchDialog()
        {
            try
            {
                // 实例化在线查价窗体
                var form = new Forms.OnlinePriceSearchForm();
                // 非模态显示窗体，不阻塞 Excel 宿主消息泵
                form.Show();
            }
            catch (Exception ex)
            {
                // 记录错误日志
                LogHelper.WriteLog($"ShowOnlinePriceSearchDialog 弹窗异常: {ex.Message}");
            }
        }
    }
}
