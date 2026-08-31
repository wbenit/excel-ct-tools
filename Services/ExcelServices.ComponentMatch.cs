using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using ExcelAddInDemo.Forms;
using ExcelAddInDemo.Models;

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
            }
            catch (Exception ex)
            {
                // 记录打开弹窗异常日志
                LogHelper.WriteLog($"ShowComponentMatchDialog 异常: {ex.Message}");
            }
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
        /// 核心批处理方法: 直读选区已有的 B 列(名称)、S 列(电流)、T 列(极数)、U 列(脱扣)，应用品牌与必含字段约束反查 WebAPI 并批量回填
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
                var activeColCfg = activeFilterCfg.ColumnConfig ?? new ComponentMatchColumnConfig();

                // 提取品牌限定与动态必含字段规则
                string selectedBrand = activeFilterCfg.SelectedBrand ?? string.Empty;
                var mustContainRules = activeFilterCfg.MustContainRules ?? new List<MustContainRule>();

                // 规范化列名 (转为大写且去除空格)
                string colName = string.IsNullOrWhiteSpace(activeColCfg.NameColumn) ? "B" : activeColCfg.NameColumn.Trim().ToUpper();
                string colCur = string.IsNullOrWhiteSpace(activeColCfg.CurrentColumn) ? "S" : activeColCfg.CurrentColumn.Trim().ToUpper();
                string colPole = string.IsNullOrWhiteSpace(activeColCfg.PoleColumn) ? "T" : activeColCfg.PoleColumn.Trim().ToUpper();
                string colTrip = string.IsNullOrWhiteSpace(activeColCfg.TripModeColumn) ? "U" : activeColCfg.TripModeColumn.Trim().ToUpper();
                string colModel = string.IsNullOrWhiteSpace(activeColCfg.ModelColumn) ? "D" : activeColCfg.ModelColumn.Trim().ToUpper();
                string colPrice = string.IsNullOrWhiteSpace(activeColCfg.PriceColumn) ? "G" : activeColCfg.PriceColumn.Trim().ToUpper();
                string colRemark = string.IsNullOrWhiteSpace(activeColCfg.RemarkColumn) ? "I" : activeColCfg.RemarkColumn.Trim().ToUpper();
                string colParam1 = string.IsNullOrWhiteSpace(activeColCfg.Param1Column) ? "W" : activeColCfg.Param1Column.Trim().ToUpper();
                string colParam2 = string.IsNullOrWhiteSpace(activeColCfg.Param2Column) ? "X" : activeColCfg.Param2Column.Trim().ToUpper();

                // 初始化统计计数器
                int totalRows = 0;
                int uniqueCount = 0;
                int multipleCount = 0;
                int noneCount = 0;

                // 遍历当前选区的所有子区域 (支持连续区域及按住 Ctrl 的多选区)
                foreach (dynamic area in selection.Areas)
                {
                    // 获取当前区域的起始行与总行数
                    int startRow = (int)area.Row;
                    int rowCount = (int)area.Rows.Count;
                    int endRow = startRow + rowCount - 1;

                    // 若行数无效则跳过
                    if (rowCount <= 0) continue;
                    totalRows += rowCount;

                    // ==================== 1. 一次性从已有的 B列(名称)、S列(电流)、T列(极数)、U列(脱扣)读入内存 ====================
                    // 一次性读入 B 列 (名称)
                    dynamic nameRange = activeSheet.Range[$"{colName}{startRow}:{colName}{endRow}"];
                    object[,] nameRawArray = ConvertTo2DArray(nameRange.Value2, rowCount);

                    // 一次性读入 S 列 (电流)
                    dynamic curRange = activeSheet.Range[$"{colCur}{startRow}:{colCur}{endRow}"];
                    object[,] curRawArray = ConvertTo2DArray(curRange.Value2, rowCount);

                    // 一次性读入 T 列 (极数)
                    dynamic poleRange = activeSheet.Range[$"{colPole}{startRow}:{colPole}{endRow}"];
                    object[,] poleRawArray = ConvertTo2DArray(poleRange.Value2, rowCount);

                    // 一次性读入 U 列 (脱扣方式)
                    dynamic tripRange = activeSheet.Range[$"{colTrip}{startRow}:{colTrip}{endRow}"];
                    object[,] tripRawArray = ConvertTo2DArray(tripRange.Value2, rowCount);

                    // ==================== 2. 在内存中分配 6 个回填目标列的二维数组 ====================
                    object[,] nameArray = new object[rowCount, 1];    // B 列 (名称)
                    object[,] modelArray = new object[rowCount, 1];   // D 列 (型号)
                    object[,] priceArray = new object[rowCount, 1];   // G 列 (单价)
                    object[,] remarkArray = new object[rowCount, 1];  // I 列 (备注)
                    object[,] param1Array = new object[rowCount, 1];  // W 列 (参数1)
                    object[,] param2Array = new object[rowCount, 1];  // X 列 (参数2)

                    // 收集当前区域中需要高亮淡黄底色的行号集合 (相对于工作表的绝对物理行号)
                    var yellowHighlightRowList = new List<int>();
                    // 收集不需要高亮/需清除底色的行号集合
                    var clearHighlightRowList = new List<int>();

                    // ==================== 3. 内存循环：调用带品牌与必含字段约束的 WebAPI 反查物料库 ====================
                    for (int i = 0; i < rowCount; i++)
                    {
                        // 计算当前内存行在 Excel 中的绝对物理行号
                        int currentRealRow = startRow + i;

                        // 直接获取当前行已存在的名称、电流、极数、脱扣方式内容
                        string rawName = nameRawArray[i + 1, 1]?.ToString()?.Trim() ?? string.Empty;
                        string minCur = curRawArray[i + 1, 1]?.ToString()?.Trim() ?? string.Empty;
                        string pole = poleRawArray[i + 1, 1]?.ToString()?.Trim() ?? string.Empty;
                        string tripMode = tripRawArray[i + 1, 1]?.ToString()?.Trim() ?? string.Empty;

                        // 若该行名称、电流与极数均为空，则判定为空白行，各字段留空并清除高亮
                        if (string.IsNullOrWhiteSpace(rawName) && string.IsNullOrWhiteSpace(minCur) && string.IsNullOrWhiteSpace(pole))
                        {
                            nameArray[i, 0] = string.Empty;
                            modelArray[i, 0] = string.Empty;
                            priceArray[i, 0] = string.Empty;
                            remarkArray[i, 0] = string.Empty;
                            param1Array[i, 0] = string.Empty;
                            param2Array[i, 0] = string.Empty;
                            clearHighlightRowList.Add(currentRealRow);
                            continue;
                        }

                        // 调用 WebAPI 客户端根据【名称、电流、极数、脱扣、品牌、必含字段】反查真实数据库
                        var matchedItems = ComponentApiClient.QueryComponents(
                            rawName,
                            minCur,
                            pole,
                            tripMode,
                            selectedBrand,
                            mustContainRules
                        );

                        // =================================================================
                        // 分支 1: 查询到唯一值 (Count == 1) -> 完整自动回填各字段
                        // =================================================================
                        if (matchedItems.Count == 1)
                        {
                            var item = matchedItems[0];
                            nameArray[i, 0] = !string.IsNullOrEmpty(item.Name) ? item.Name : rawName;       // B 列 (标准名称)
                            modelArray[i, 0] = item.Model ?? string.Empty;                                 // D 列 (标准型号)
                            priceArray[i, 0] = item.Price > 0 ? (object)item.Price : string.Empty;         // G 列 (单价)
                            remarkArray[i, 0] = item.Remark ?? string.Empty;                               // I 列 (备注)
                            param1Array[i, 0] = item.Param1 ?? string.Empty;                               // W 列 (参数1)
                            param2Array[i, 0] = item.Param2 ?? string.Empty;                               // X 列 (参数2)

                            clearHighlightRowList.Add(currentRealRow);
                            uniqueCount++;
                        }
                        // =================================================================
                        // 分支 2: 查出来多个匹配 (Count > 1) -> 保留B列原名，D列填入“点击查询(条数)”并设淡黄底色
                        // =================================================================
                        else if (matchedItems.Count > 1)
                        {
                            nameArray[i, 0] = rawName;         // B 列 (保留原有名称)
                            // D 列填入带匹配条数的提示文本 (如: 点击查询(7)) --硬编码--
                            modelArray[i, 0] = $"{ComponentMatchDefaults.MultipleCandidatesText}({matchedItems.Count})";
                            priceArray[i, 0] = string.Empty;
                            remarkArray[i, 0] = string.Empty;
                            param1Array[i, 0] = string.Empty;
                            param2Array[i, 0] = string.Empty;

                            yellowHighlightRowList.Add(currentRealRow);
                            multipleCount++;
                        }
                        // =================================================================
                        // 分支 3: 没找到任何匹配 (Count == 0) -> 保留B列原名，其余置空并清除底色
                        // =================================================================
                        else
                        {
                            nameArray[i, 0] = rawName;         // B 列 (保留原有名称)
                            modelArray[i, 0] = string.Empty;
                            priceArray[i, 0] = string.Empty;
                            remarkArray[i, 0] = string.Empty;
                            param1Array[i, 0] = string.Empty;
                            param2Array[i, 0] = string.Empty;

                            clearHighlightRowList.Add(currentRealRow);
                            noneCount++;
                        }
                    }

                    // ==================== 4. 一次性将二维数组整块写回 Excel 目标字段列 (不覆盖S/T/U列) ====================
                    // 写回 B 列 (名称)
                    activeSheet.Range[$"{colName}{startRow}:{colName}{endRow}"].Value2 = nameArray;
                    // 写回 D 列 (型号)
                    activeSheet.Range[$"{colModel}{startRow}:{colModel}{endRow}"].Value2 = modelArray;
                    // 写回 G 列 (单价)
                    activeSheet.Range[$"{colPrice}{startRow}:{colPrice}{endRow}"].Value2 = priceArray;
                    // 写回 I 列 (备注)
                    activeSheet.Range[$"{colRemark}{startRow}:{colRemark}{endRow}"].Value2 = remarkArray;
                    // 写回 W 列 (参数1)
                    activeSheet.Range[$"{colParam1}{startRow}:{colParam1}{endRow}"].Value2 = param1Array;
                    // 写回 X 列 (参数2)
                    activeSheet.Range[$"{colParam2}{startRow}:{colParam2}{endRow}"].Value2 = param2Array;

                    // ==================== 5. 针对“点击查询”单元格统一应用淡黄底色 ====================
                    // 统一设置多条匹配行的 D 列单元格背景颜色为淡黄色
                    foreach (int r in yellowHighlightRowList)
                    {
                        dynamic targetCell = activeSheet.Range[$"{colModel}{r}"];
                        // 设置淡黄底色 (RGB: 255, 242, 204)
                        targetCell.Interior.Color = ComponentMatchDefaults.LightYellowOleColor;
                    }

                    // 统一清除唯一匹配或无匹配行的 D 列底色 (避免残留黄色背景)
                    foreach (int r in clearHighlightRowList)
                    {
                        dynamic targetCell = activeSheet.Range[$"{colModel}{r}"];
                        // 仅当已有底色时重置为无填充
                        if ((int)targetCell.Interior.ColorIndex != ComponentMatchDefaults.XlNoneColorIndex)
                        {
                            targetCell.Interior.ColorIndex = ComponentMatchDefaults.XlNoneColorIndex;
                        }
                    }
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
        public static void ShowComponentMatchOverlay(dynamic activeCell)
        {
            if (activeCell == null) return;

            try
            {
                // 1. 校验是否处于 D 列 (第 4 列: 规格型号)
                int col = 0;
                try { col = Convert.ToInt32(activeCell.Column); } catch { }
                if (col != 4)
                {
                    HideComponentMatchOverlay();
                    return;
                }

                // 2. 获取当前行号与所属工作表
                int row = 0;
                try { row = Convert.ToInt32(activeCell.Row); } catch { }
                if (row <= 0) return;

                dynamic sheet = activeCell.Worksheet;
                if (sheet == null) return;

                // 3. 读取当前行已有的 B 列(名称)、S 列(电流)、T 列(极数)、U 列(脱扣)、D 列(型号)、C 列(原型号)、G 列(单价)
                string rawName = Convert.ToString(sheet.Range[$"B{row}"].Value2)?.Trim() ?? string.Empty;
                string rawCur = Convert.ToString(sheet.Range[$"S{row}"].Value2)?.Trim() ?? string.Empty;
                string rawPole = Convert.ToString(sheet.Range[$"T{row}"].Value2)?.Trim() ?? string.Empty;
                string rawTrip = Convert.ToString(sheet.Range[$"U{row}"].Value2)?.Trim() ?? string.Empty;

                // 读取当前 D 列实际内容（型号规格）与 C 列内容（原型号规格）
                string rawModel = Convert.ToString(sheet.Range[$"D{row}"].Value2)?.Trim() ?? string.Empty;
                string refModel = Convert.ToString(sheet.Range[$"C{row}"].Value2)?.Trim() ?? string.Empty;
                // 若 D 列为空或为占位提示“点击查询”，优先采用 C 列原型号作为主体型号匹配附件
                if (string.IsNullOrEmpty(rawModel) || string.Equals(rawModel, "点击查询", StringComparison.OrdinalIgnoreCase))
                {
                    rawModel = refModel;
                }

                // 读取当前 G 列单价或公式（优先通过 Formula 或 Value2 获取）
                string rawPriceFormula = Convert.ToString(sheet.Range[$"G{row}"].Formula)?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(rawPriceFormula))
                {
                    rawPriceFormula = Convert.ToString(sheet.Range[$"G{row}"].Value2)?.Trim() ?? string.Empty;
                }

                // 4. 加载当前生效的全局过滤管道配置 (品牌 + 必含字段约束 + 搜索开关)
                var filterConfig = LoadComponentMatchFilterConfig();

                // 构造上下文参数 (带上原型号、原单价与所属品牌)
                var cellParams = new CellParamsContext
                {
                    Name = rawName,
                    Current = rawCur,
                    Pole = rawPole,
                    TripMode = rawTrip,
                    CurrentModel = rawModel,
                    CurrentPrice = rawPriceFormula,
                    Brand = filterConfig.SelectedBrand ?? string.Empty
                };

                // 核心门控: 只有当用户在设置面板中勾选开启了“搜索”时，点击 D 列才弹起搜索框
                if (!filterConfig.EnableSearchOverlay)
                {
                    HideComponentMatchOverlay();
                    return;
                }

                // 5. 按照已有参数初筛候选列表
                var initialItems = ComponentApiClient.QueryComponents(
                    rawName,
                    rawCur,
                    rawPole,
                    rawTrip,
                    filterConfig.SelectedBrand,
                    filterConfig.MustContainRules
                );

                // 6. 初始化或复用下拉悬浮窗实例
                if (_matchOverlayForm == null || _matchOverlayForm.IsDisposed)
                {
                    _matchOverlayForm = new ComponentMatchOverlayForm();
                }

                // 在当前单元格下方贴合弹出
                _matchOverlayForm.ShowAtCell(activeCell, initialItems, cellParams, filterConfig);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ShowComponentMatchOverlay 异常: {ex.Message}");
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

                // 1. 回填 B 列 (标准名称)
                if (!string.IsNullOrEmpty(item.Name))
                {
                    sheet.Range[$"B{row}"].Value2 = item.Name;
                }

                // 2. 回填 D 列 (标准型号，覆盖掉“点击查询”)
                sheet.Range[$"D{row}"].Value2 = item.Model ?? string.Empty;

                // 3. 回填 I 列 (品牌/生产厂家)
                sheet.Range[$"I{row}"].Value2 = item.Remark ?? item.Brand ?? string.Empty;

                // 4. 回填价格逻辑
                if (isSummarySheet)
                {
                    // 在“元件汇总表”中：汇总表自带公式，仅填 L 列本体表价
                    sheet.Range[$"L{row}"].Value2 = item.Price > 0 ? (object)(double)item.Price : string.Empty;

                    // 若 M 列 (本体折扣) 当前为空或 0，则默认补 1
                    string curM = Convert.ToString(sheet.Range[$"M{row}"].Value2)?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(curM) || curM == "0")
                    {
                        sheet.Range[$"M{row}"].Value2 = 1;
                    }
                }
                else
                {
                    // 在常规分类明细表中：本体表价填在 M 列 (表价)
                    sheet.Range[$"M{row}"].Value2 = item.Price > 0 ? (object)item.Price : string.Empty;

                    // 回填 W 列 (扩展参数1)
                    sheet.Range[$"W{row}"].Value2 = item.Param1 ?? string.Empty;

                    // 回填 X 列 (扩展参数2)
                    sheet.Range[$"X{row}"].Value2 = item.Param2 ?? string.Empty;
                }

                // 5. 清除 D 列单元格的淡黄底色 (重置为无填充 xlNone)
                dynamic cellD = sheet.Range[$"D{row}"];
                cellD.Interior.ColorIndex = ComponentMatchDefaults.XlNoneColorIndex;

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"FillSelectedComponentToActiveRow 回填异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 用户选中配套附件后：D 列追加“+附件型号”
        /// 汇总表中附件填 N 列（多附件加法公式），常规表中本体和附件在 M 列进行加法公式连加
        /// </summary>
        /// <param name="attachment">选中的附件元器件物料 DTO</param>
        /// <param name="targetCell">目标单元格 COM 句柄</param>
        /// <returns>回填是否成功</returns>
        public static bool FillSelectedAttachmentToActiveRow(ComponentApiDto attachment, dynamic? targetCell = null)
        {
            if (attachment == null) return false;

            try
            {
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null) return false;

                dynamic cell = targetCell ?? app.ActiveCell;
                if (cell == null) return false;

                dynamic sheet = cell.Worksheet;
                if (sheet == null) return false;

                int row = Convert.ToInt32(cell.Row);
                if (row <= 0) return false;

                // 判断是否在“元件汇总表”工作表中
                string sheetName = Convert.ToString(sheet.Name) ?? string.Empty;
                bool isSummarySheet = string.Equals(sheetName.Trim(), ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase);

                // 1. 获取 D 列原有型号内容
                string oldModel = Convert.ToString(sheet.Range[$"D{row}"].Value2)?.Trim() ?? string.Empty;
                // 若原单元格内容为初始提示词“点击查询”，则清空
                if (string.Equals(oldModel, "点击查询", StringComparison.OrdinalIgnoreCase))
                {
                    oldModel = string.Empty;
                }

                // 附件型号
                string attachModel = attachment.Model?.Trim() ?? string.Empty;

                // 拼接新型号文本：“原内容+附件型号”
                string newModel;
                if (string.IsNullOrEmpty(oldModel))
                {
                    newModel = attachModel;
                }
                else
                {
                    // 消除首尾已有的加号字符，防止出现连续的 "++" 符号
                    string trimmedOld = oldModel.TrimEnd('+', '＋');
                    string trimmedAttach = attachModel.TrimStart('+', '＋');
                    newModel = $"{trimmedOld}+{trimmedAttach}";
                }

                // 回填到 D 列
                sheet.Range[$"D{row}"].Value2 = newModel;

                // 2. 处理附件价格回填逻辑
                decimal attachPrice = attachment.Price > 0 ? attachment.Price : 0m;
                if (attachPrice > 0)
                {
                    if (isSummarySheet)
                    {
                        // 在“元件汇总表”中：附件价格填入 N 列 (附件表价)，多个附件用加法公式连接
                        string oldFormulaN = Convert.ToString(sheet.Range[$"N{row}"].Formula)?.Trim() ?? string.Empty;
                        string oldValN = Convert.ToString(sheet.Range[$"N{row}"].Value2)?.Trim() ?? string.Empty;

                        // 若此前 N 列没有任何价格数据，直接填入单个附件数值
                        if (string.IsNullOrEmpty(oldFormulaN) && string.IsNullOrEmpty(oldValN))
                        {
                            sheet.Range[$"N{row}"].Value2 = (double)attachPrice;
                        }
                        // 若此前已有公式 (以 '=' 开头)，直接在原公式末尾累加 "+附件价格"
                        else if (!string.IsNullOrEmpty(oldFormulaN) && oldFormulaN.StartsWith("="))
                        {
                            sheet.Range[$"N{row}"].Formula = $"{oldFormulaN}+{attachPrice:0.##}";
                        }
                        // 若此前为纯数值 (例如 49)，升级为加法公式 (例如 =49+60)
                        else
                        {
                            string baseNum = !string.IsNullOrEmpty(oldValN) ? oldValN : oldFormulaN;
                            sheet.Range[$"N{row}"].Formula = $"={baseNum}+{attachPrice:0.##}";
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
                        // 常规分类明细表中：本体和附件在 M 列 (表价) 连加 (如 =447.01+150 或 =447.01+150+60)
                        string oldFormulaM = Convert.ToString(sheet.Range[$"M{row}"].Formula)?.Trim() ?? string.Empty;
                        string oldValM = Convert.ToString(sheet.Range[$"M{row}"].Value2)?.Trim() ?? string.Empty;

                        string newFormulaM;
                        // 若此前 M 列无任何数据，直接写公式 "=附件价格"
                        if (string.IsNullOrEmpty(oldFormulaM) && string.IsNullOrEmpty(oldValM))
                        {
                            newFormulaM = $"={attachPrice:0.##}";
                        }
                        // 若此前已有公式 (以 '=' 开头)，直接在原公式末尾连加 "+附件价格"
                        else if (!string.IsNullOrEmpty(oldFormulaM) && oldFormulaM.StartsWith("="))
                        {
                            newFormulaM = $"{oldFormulaM}+{attachPrice:0.##}";
                        }
                        // 若此前为本体纯数值 (如 447.01)，升级为连加公式 (如 =447.01+150)
                        else
                        {
                            string basePriceStr = !string.IsNullOrEmpty(oldValM) ? oldValM : oldFormulaM;
                            newFormulaM = $"={basePriceStr}+{attachPrice:0.##}";
                        }

                        // 写入 M 列公式
                        sheet.Range[$"M{row}"].Formula = newFormulaM;
                    }
                }

                // 3. 清除 D 列底色
                dynamic cellD = sheet.Range[$"D{row}"];
                cellD.Interior.ColorIndex = ComponentMatchDefaults.XlNoneColorIndex;

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
