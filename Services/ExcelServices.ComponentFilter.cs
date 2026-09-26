using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelAddInDemo.Services;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 公共服务分部类: 元器件多选行当前列筛选与同箱柜共存标记服务
    /// </summary>
    public static partial class ExcelServices
    {
        // 单元格原始颜色快照数据模型，用于在取消筛选时 100% 还原用户自定义标记色
        private class CellColorSnapshot
        {
            // 单元格所在物理行号
            public int Row { get; set; }
            // 单元格所在物理列号
            public int Col { get; set; }
            // 原先是否为无填充色 (xlNone: -4142)
            public bool HasNoColor { get; set; }
            // 原先的具体 OLE 颜色数值
            public object? OriginalColor { get; set; }
        }

        // 缓存工作表在筛选前的单元格原始背景色快照字典 (Key: 工作表名称，确保跨表独立)
        private static readonly Dictionary<string, List<CellColorSnapshot>> _filterOriginalColorSnapshots 
            = new Dictionary<string, List<CellColorSnapshot>>(StringComparer.OrdinalIgnoreCase);

        // 虚拟表头备份快照数据模型，用于在取消筛选时 100% 还原第 1 行原本的行高、单元格内容与排版格式 (方案 A)
        private class VirtualHeaderSnapshot
        {
            // 工作表名称
            public string SheetName { get; set; } = string.Empty;
            // 第 1 行原始行高数值
            public object? OriginalRowHeight { get; set; }
            // 第 1 行原始单元格数据二维数组 (object[,])
            public object[,]? OriginalValues { get; set; }
            // 第 1 行原始单元格公式二维数组 (object[,])
            public object[,]? OriginalFormulas { get; set; }
            // 第 1 行原始背景色 OLE 数值
            public object? OriginalInteriorColor { get; set; }
            // 第 1 行原始背景色索引 (如 xlNone 等)
            public object? OriginalInteriorColorIndex { get; set; }
            // 第 1 行原始字体加粗属性
            public object? OriginalFontBold { get; set; }
            // 第 1 行原始字体字号
            public object? OriginalFontSize { get; set; }
            // 第 1 行原始水平对齐方式
            public object? OriginalHorizontalAlignment { get; set; }
            // 第 1 行原始垂直对齐方式
            public object? OriginalVerticalAlignment { get; set; }
            // 原始是否开启了窗口冻结窗格
            public bool? OriginalFreezePanes { get; set; }
            // 原始冻结窗口拆分行数
            public int OriginalSplitRow { get; set; }
            // 原始冻结窗口拆分列数
            public int OriginalSplitColumn { get; set; }
            // 原始第 1 行是否存在单元格合并
            public bool OriginalMergeCells { get; set; }
            // 快照记录的有效列数
            public int ColumnCount { get; set; }
        }

        // 缓存工作表在筛选前第 1 行原始状态快照字典 (Key: 工作表名称，确保跨表独立)
        private static readonly Dictionary<string, VirtualHeaderSnapshot> _virtualHeaderSnapshots 
            = new Dictionary<string, VirtualHeaderSnapshot>(StringComparer.OrdinalIgnoreCase);

        // 分类明细表标准默认列名数组 (用于在未能从箱柜明细行提取到时的强壮兜底)
        // --硬编码: 分类明细表默认标准列名--
        private static readonly string[] DefaultCategoryDetailHeaders = new string[]
        {
            "序号", "元件名称", "型号规格", "生产厂家", "单位", "数量", "单   价", "总   价", "备注",
            "成本单价", "成本总价", "报出系数", "表价", "折扣系数", "取费系数", "成套费", "类别", "加工费"
        };

        // 虚拟表头标准显示行高 (24.0pt，充裕包容文字与筛选箭头)
        // --硬编码: 虚拟表头标准行高 24.0--
        private const double VirtualHeaderStandardRowHeight = 24.0;

        // 虚拟表头背景底色 Hex #E8F4F2 (淡青微灰，完美呼应成套 #009688 主题色调)
        // --硬编码: 虚拟表头背景色 RGB--
        private static readonly Color VirtualHeaderBgColor = Color.FromArgb(232, 244, 242);

        // 相邻箱柜区分：白底 (第一台/奇数台)
        // --硬编码: 第一台/奇数台白底色--
        private static readonly Color AlternateWhiteColor = Color.White;

        // 相邻箱柜区分：青底 (第二台/偶数台)，采用柔和淡青绿 Hex #E0F2F1，对齐 #009688 主题
        // --硬编码: 第二台/偶数台淡青底色 RGB--
        private static readonly Color AlternateCyanColor = Color.FromArgb(224, 242, 241);

        // Excel 常数: 无填充色 ColorIndex (xlNone)
        // --硬编码: Excel xlNone 常数--
        private const int XlNoneColorIndex = -4142;

        // Excel 常数: 手动计算模式 (xlCalculationManual)
        // --硬编码: Excel 手动计算常数--
        private const int XlCalcManual = -4135;

        // Excel 常数: 自动计算模式 (xlCalculationAutomatic)
        // --硬编码: Excel 自动计算常数--
        private const int XlCalcAutomatic = -4105;

        /// <summary>
        /// 执行 Excel 系统原生自动筛选：按当前选区（支持单选或多选值）的值筛选 (100% 系统级 AutoFilter)
        /// </summary>
        public static void ExecuteNativeFilterBySelection()
        {
            try
            {
                // 获取 Excel Application 动态实例
                dynamic? app = ExcelDnaUtil.Application;
                // 校验 Excel 全局实例有效性
                if (app == null) return;

                // 获取当前活动单元格与活动工作表
                dynamic? activeCell = app.ActiveCell;
                // 获取当前活动工作表
                dynamic? activeSheet = app.ActiveSheet;
                // 获取当前工作簿选区 Selection
                dynamic? selection = app.Selection;

                // 若未检测到有效活动单元格，友好提示并返回
                if (activeCell == null || activeSheet == null)
                {
                    // 弹窗提示需要先选中单元格
                    MessageBox.Show("未检测到有效活动单元格，请先选中需要筛选的单元格！", "原生筛选提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    // 中断返回
                    return;
                }

                // 提取活动单元格所在的绝对物理列号
                int activeCol = (int)activeCell.Column;

                // 从用户当前选区中提取目标列的所有非空文本（去重保序，支持单选与多选）
                List<string> filterKeywords = ExtractFilterKeywordsFromSelection(selection, activeCol, activeCell);

                // 标记原生自动筛选是否成功应用
                bool isFilterApplied = false;

                // 核心业务升级 (方案 A)：智能判定工作表是否需要挂载虚拟表头 (零依赖项目信息白名单，基于排版与内容特征自适应)
                bool needsVirtualHeader = NeedsVirtualHeader(activeSheet);
                if (needsVirtualHeader)
                {
                    // 挂载虚拟表头并调整行高样式与冻结首行
                    EnsureVirtualHeaderForCategorySheet(activeSheet, app);
                }

                // 场景 A: 当前工作表已经处于 AutoFilter 开启状态
                if (activeSheet.AutoFilterMode == true && activeSheet.AutoFilter != null)
                {
                    // 若需要虚拟表头，确保第 1 行行高被强制恢复为 24pt 且处于非隐藏状态
                    if (needsVirtualHeader)
                    {
                        try
                        {
                            activeSheet.Rows[1].Hidden = false;
                            activeSheet.Rows[1].RowHeight = VirtualHeaderStandardRowHeight;
                        }
                        catch { }
                    }

                    // 提取现有的筛选区域 Range
                    dynamic filterRange = activeSheet.AutoFilter.Range;
                    int filterStartRow = (int)filterRange.Row;

                    // 关键防御：若需要虚拟表头但旧筛选未从第 1 行开始（如在第 6 行或中间某行），重置旧 AutoFilter 重新以第 1 行开启
                    if (needsVirtualHeader && filterStartRow != 1)
                    {
                        try { activeSheet.AutoFilterMode = false; } catch { }
                    }
                    else
                    {
                        // 提取筛选区域起始列
                        int startCol = (int)filterRange.Column;
                        // 计算筛选区域结束列
                        int endCol = startCol + (int)filterRange.Columns.Count - 1;

                        // 若活动单元格列落在该筛选列范围内
                        if (activeCol >= startCol && activeCol <= endCol)
                        {
                            // 计算列在筛选区域内部的相对列索引 (1-based)
                            int fieldIndex = activeCol - startCol + 1;
                            // 调用多值原生 AutoFilter 执行筛选
                            ApplyNativeAutoFilter(filterRange, fieldIndex, filterKeywords);
                            // 标记已成功应用筛选
                            isFilterApplied = true;
                        }
                        else if (needsVirtualHeader)
                        {
                            // 若活动列超出原筛选列范围，重置旧 AutoFilter 并以整表重新开启
                            try { activeSheet.AutoFilterMode = false; } catch { }
                        }
                    }
                }

                // 场景 B: 若为普通平铺表且未应用筛选，智能定位当前单元格所在的连续数据块 CurrentRegion
                if (!isFilterApplied && !needsVirtualHeader)
                {
                    dynamic? targetRegion = null;
                    try
                    {
                        // 读取活动单元格连续区域
                        targetRegion = activeCell.CurrentRegion;
                    }
                    catch { }

                    // 校验连续区域行数有效性 (至少包含表头与一行数据，>= 2 行)
                    if (targetRegion != null && (int)targetRegion.Rows.Count >= 2)
                    {
                        // 提取连续区域起始列
                        int startCol = (int)targetRegion.Column;
                        // 计算连续区域结束列
                        int endCol = startCol + (int)targetRegion.Columns.Count - 1;

                        // 检查活动列是否在连续区域内
                        if (activeCol >= startCol && activeCol <= endCol)
                        {
                            // 计算相对字段列号
                            int fieldIndex = activeCol - startCol + 1;
                            // 在连续数据块上开启并执行原生 AutoFilter (单值或多值)
                            ApplyNativeAutoFilter(targetRegion, fieldIndex, filterKeywords);
                            // 标记已成功应用筛选
                            isFilterApplied = true;
                        }
                    }
                }

                // 场景 C: 多箱柜成套明细表 (方案 A 核心承载) 或普通表兜底：以第 1 行为表头开启全局原生自动筛选
                if (!isFilterApplied)
                {
                    // 获取包含虚拟表头的整表已用区域 UsedRange
                    dynamic usedRange = activeSheet.UsedRange;
                    dynamic targetFilterRange = usedRange;

                    // 若挂载了虚拟表头，显式构建以 A1 为顶点的标准连续矩形区域，确保 100% 将第 1 行作为表头行挂载漏斗箭头
                    if (needsVirtualHeader)
                    {
                        int maxCols = 25;
                        int maxRows = 100;
                        try
                        {
                            if (usedRange != null)
                            {
                                maxCols = Math.Max(maxCols, (int)usedRange.Column + (int)usedRange.Columns.Count - 1);
                                maxRows = Math.Max(maxRows, (int)usedRange.Row + (int)usedRange.Rows.Count - 1);
                            }
                        }
                        catch { }
                        targetFilterRange = activeSheet.Range[activeSheet.Cells[1, 1], activeSheet.Cells[maxRows, maxCols]];
                    }

                    // 校验目标筛选区域行数
                    if (targetFilterRange != null && (int)targetFilterRange.Rows.Count >= 2)
                    {
                        // 提取起始列
                        int startCol = (int)targetFilterRange.Column;
                        // 计算相对列号
                        int fieldIndex = activeCol - startCol + 1;
                        // 校验列索引范围
                        if (fieldIndex >= 1 && fieldIndex <= (int)targetFilterRange.Columns.Count)
                        {
                            // 对整表执行自动筛选 (第 1 行作为具有 24pt 行高与清晰列名的表头挂载漏斗箭头)
                            ApplyNativeAutoFilter(targetFilterRange, fieldIndex, filterKeywords);
                            // 标记已成功应用筛选
                            isFilterApplied = true;
                        }
                    }
                }

                // 核心视口与吸顶保障：筛选完成后，强制将视口垂直滚动到第 1 行，确保虚拟表头第一眼清晰可见
                if (isFilterApplied && needsVirtualHeader)
                {
                    try
                    {
                        if (app.ActiveWindow != null)
                        {
                            // 视口置顶滚动至第 1 行
                            app.ActiveWindow.ScrollRow = 1;
                            app.ActiveWindow.ScrollColumn = 1;
                        }
                    }
                    catch { }
                }

                // 若未能成功应用筛选，友好弹窗提示
                if (!isFilterApplied)
                {
                    // 弹出友好提示
                    MessageBox.Show("当前选区未处于可识别的数据表格区域内，无法开启原生自动筛选！", "原生筛选提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    // 退出
                    return;
                }

                // 核心业务增强：若当前表为分类明细表，自动对命中行执行“相邻箱柜”白/淡青底交替分色并备份原色快照
                ApplyAdjacentCabinetColorsAfterFilter(app, activeSheet);

                // 更新底部状态栏提示
                ShowNativeFilterStatusBar(app, filterKeywords);

                // 准备关键词摘要
                string kwSummary = (filterKeywords != null && filterKeywords.Count > 0)
                    ? string.Join("、", filterKeywords.Take(2))
                    : "空值";
                if (filterKeywords != null && filterKeywords.Count > 2) kwSummary += "...";

                // 推入通用撤销命令：支持按 Ctrl+Z 一键撤销原生筛选、解除隐藏并 100% 还原用户原始底色
                UndoRedoManager.Instance.PushCommand(new ActionUndoableCommand(
                    $"原生筛选 [{kwSummary}]",
                    () => ClearComponentFilter(activeSheet),
                    () => ExecuteNativeFilterBySelection()
                ));
            }
            catch (Exception ex)
            {
                // 记录原生筛选异常日志
                LogHelper.WriteLog($"执行原生筛选异常: {ex.Message}\r\n{ex.StackTrace}");
                // 弹出异常提示框
                MessageBox.Show($"原生筛选执行失败: {ex.Message}", "原生筛选提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// 在原生筛选后，若当前表为多箱柜分类明细表，自动对命中可见行执行相邻箱柜淡青/白底交替分色并备份原色快照
        /// </summary>
        private static void ApplyAdjacentCabinetColorsAfterFilter(dynamic app, dynamic activeSheet)
        {
            try
            {
                // 获取工作簿句柄
                dynamic? activeWb = null;
                try { activeWb = activeSheet.Parent; } catch { }

                // 提取合规箱柜列表
                var validCabinets = Tool.GetSheetValidCabinets((object)activeSheet, (object?)activeWb);
                // 若当前工作表不存在箱柜定义名称（如普通平铺表），直接退出无需箱柜分色
                if (validCabinets == null || validCabinets.Count == 0) return;

                // 收集各命中的箱柜及其当前可见元器件行
                var hitCabinetsOrdered = new List<(int CabIdx, List<int> VisibleRows)>();
                // 收集全表所有命中可见物理行号 (使用 HashSet 自动去重)
                HashSet<int> allVisibleHitRows = new HashSet<int>();

                // 遍历工作表中每一个合规箱柜
                foreach (var kv in validCabinets)
                {
                    var anchor = kv.Value;
                    if (anchor?.Det == null || anchor?.Subsum == null) continue;

                    // 根据规则 6：Cab_Det.row+2 为元器件起始行，Cab_Subsum.row-1 为元器件终止行
                    int compStartRow = anchor.Det.Row + 2;
                    int compEndRow = anchor.Subsum.Row - 1;
                    if (compEndRow < compStartRow) continue;

                    // 收集当前箱柜内处于可见状态的元器件行
                    List<int> visibleRowsInCab = new List<int>();

                    try
                    {
                        // 优先通过 SpecialCells(12 即 xlCellTypeVisible) 一次性极速获取该箱柜元器件区所有可见行
                        // --硬编码: Excel 常数 xlCellTypeVisible 为 12--
                        dynamic compRange = activeSheet.Range[$"A{compStartRow}:A{compEndRow}"];
                        dynamic visibleCells = compRange.SpecialCells(12);
                        if (visibleCells != null)
                        {
                            // 遍历可见离散块区域 Areas
                            foreach (dynamic area in visibleCells.Areas)
                            {
                                int areaRow = (int)area.Row;
                                int areaCount = (int)area.Rows.Count;
                                // 登记各可见物理行号
                                for (int r = areaRow; r < areaRow + areaCount; r++)
                                {
                                    visibleRowsInCab.Add(r);
                                    allVisibleHitRows.Add(r);
                                }
                            }
                        }
                    }
                    catch
                    {
                        // 若 SpecialCells 抛出异常 (说明该区间内无任何可见行，全部被过滤隐藏)，静默跳过
                    }

                    // 容错兜底：若 SpecialCells 未获取到且包含行数时，逐行校验 Hidden
                    if (visibleRowsInCab.Count == 0)
                    {
                        for (int r = compStartRow; r <= compEndRow; r++)
                        {
                            try
                            {
                                bool isHidden = true;
                                try { isHidden = Convert.ToBoolean(activeSheet.Rows[r].Hidden); } catch { }
                                // 若当前行未被隐藏，登记该可见行
                                if (!isHidden)
                                {
                                    visibleRowsInCab.Add(r);
                                    allVisibleHitRows.Add(r);
                                }
                            }
                            catch { }
                        }
                    }

                    // 若当前箱柜存在可见行，登记该箱柜
                    if (visibleRowsInCab.Count > 0)
                    {
                        hitCabinetsOrdered.Add((kv.Key, visibleRowsInCab));
                    }
                }

                // 若存在需要上色的可见行
                if (allVisibleHitRows.Count > 0)
                {
                    // 挂起屏幕更新保证备份与上色丝滑流畅
                    bool prevUpdating = true;
                    try
                    {
                        prevUpdating = app.ScreenUpdating;
                        app.ScreenUpdating = false;
                    }
                    catch { }

                    try
                    {
                        // 步骤 1: 若先前已有快照，先还原旧快照，杜绝用户原色丢失
                        RestoreColorSnapshotsForSheet(activeSheet);

                        // 步骤 2: 对即将上色的所有可见行 A~M 列执行原始背景色快照备份（100% 保护用户自定义标记色）
                        BackupColorSnapshotsForRows(activeSheet, allVisibleHitRows);

                        // 步骤 3: 相邻箱柜斑马纹交替上色 (第一台淡青底，第二台白底...)
                        int cyanOle = ColorTranslator.ToOle(AlternateCyanColor);
                        int whiteOle = ColorTranslator.ToOle(AlternateWhiteColor);

                        // 遍历命中箱柜序列
                        for (int cabOrder = 0; cabOrder < hitCabinetsOrdered.Count; cabOrder++)
                        {
                            var cabItem = hitCabinetsOrdered[cabOrder];
                            // 偶数序号赋予淡青底，奇数序号赋予白底，相邻箱柜边界极其鲜明
                            int targetOle = (cabOrder % 2 == 0) ? cyanOle : whiteOle;

                            // 遍历该箱柜内的可见行
                            foreach (int r in cabItem.VisibleRows)
                            {
                                try
                                {
                                    // 为 A 列至 M 列赋予对应箱柜的交替背景底色
                                    activeSheet.Range[$"A{r}:M{r}"].Interior.Color = targetOle;
                                }
                                catch { }
                            }
                        }
                    }
                    finally
                    {
                        // 恢复屏幕更新
                        try { app.ScreenUpdating = prevUpdating; } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录相邻箱柜上色异常日志
                LogHelper.WriteLog($"原生筛选后相邻箱柜分色异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 核心辅助方法：对指定 Range 区域应用单值或多值 Excel 系统原生 AutoFilter
        /// </summary>
        private static void ApplyNativeAutoFilter(dynamic targetRange, int fieldIndex, List<string> filterKeywords)
        {
            // 若未提取到有效值，按空白单元格筛选
            if (filterKeywords == null || filterKeywords.Count == 0)
            {
                // 等号表示空白筛选
                targetRange.AutoFilter(fieldIndex, "=");
            }
            else if (filterKeywords.Count == 1)
            {
                // 单值筛选：直接传入单一字符串准则
                targetRange.AutoFilter(fieldIndex, filterKeywords[0]);
            }
            else
            {
                // 多值联合筛选：必须将值列表转换为一维 object[] 数组
                object[] criteriaArray = filterKeywords.Cast<object>().ToArray();
                // Operator 传入 7 即 Excel 常数 XlAutoFilterOperator.xlFilterValues (支持同时勾选多个值)
                // --硬编码: Excel xlFilterValues 枚举数值为 7--
                targetRange.AutoFilter(fieldIndex, criteriaArray, (dynamic)7);
            }
        }

        /// <summary>
        /// 在 Excel 底部状态栏展示原生筛选统计提示
        /// </summary>
        private static void ShowNativeFilterStatusBar(dynamic app, List<string> filterKeywords)
        {
            try
            {
                // 拼接关键词摘要 (最多展示前 3 个)
                string summary = (filterKeywords != null && filterKeywords.Count > 0)
                    ? string.Join("、", filterKeywords.Take(3))
                    : "空白";
                // 若超过 3 个追加省略号
                if (filterKeywords != null && filterKeywords.Count > 3) summary += "...";
                // 写入 Excel 状态栏提示信息
                int count = filterKeywords?.Count ?? 0;
                app.StatusBar = $"[原生筛选] 已对当前列按 [{summary}] (共 {count} 个值) 应用系统自动筛选";
            }
            catch { }
        }

        /// <summary>
        /// 根据用户当前选区（单选或多选行）在当前激活列的内容，执行多选联合筛选并进行同箱柜共存标记
        /// </summary>
        public static void FilterComponentsBySelection()
        {
            try
            {
                // 获取 Excel Application 动态实例
                dynamic? app = ExcelDnaUtil.Application;
                // 校验 Excel 全局实例有效性
                if (app == null) return;

                // 获取当前工作簿活动选区 Range
                dynamic? selection = app.Selection;
                // 获取当前活动单元格 ActiveCell
                dynamic? activeCell = app.ActiveCell;
                // 获取当前活动工作表 ActiveSheet
                dynamic? activeSheet = app.ActiveSheet;

                // 若选区或工作表为空，直接退出
                if (selection == null || activeSheet == null)
                {
                    // 弹出友好提示告知用户未选定单元格
                    MessageBox.Show("未检测到有效单元格选区，请先选中需要筛选的单元格或行！", "筛选提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    // 中断退出
                    return;
                }

                // 确定目标物理列号（以当前活动单元格所在列为准，支持用户跨列拖选时聚焦当前列）
                int targetCol = (activeCell != null) ? (int)activeCell.Column : 2;

                // 准备去重保存用户选中的筛选关键字列表
                List<string> filterKeywords = ExtractFilterKeywordsFromSelection(selection, targetCol, activeCell);

                // 若未能提取到任何有效文本，友好提示并返回
                if (filterKeywords.Count == 0)
                {
                    // 弹出友好提示告知用户所选单元格为空
                    MessageBox.Show("所选单元格在当前列的内容为空，请选择包含有效元器件名称或型号的单元格后重试！", "筛选提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    // 中断退出
                    return;
                }

                // 规则 8 架构安全防护：在操作前自愈与校验当前工作表的定义名称拓扑结构
                Tool.FixAndFillCabinetNamesForSheet(activeSheet);

                // 获取所属工作簿对象
                dynamic? activeWb = null;
                // 向上追溯所属工作簿句柄
                try { activeWb = activeSheet.Parent; } catch { }

                // 获取当前工作表中所有合规箱柜锚点 (按物理行号升序排列)
                var validCabinets = Tool.GetSheetValidCabinets((object)activeSheet, (object?)activeWb);

                // 若当前工作表不存在箱柜定义名称（可能为元件汇总表或普通数据表），走通用平铺行筛选降级方案
                if (validCabinets == null || validCabinets.Count == 0)
                {
                    // 执行普通扁平数据表的当前列多选筛选降级逻辑
                    FilterFlatSheetByKeywords(app, activeSheet, targetCol, filterKeywords);
                }
                else
                {
                    // 执行分类明细表专属的箱柜区间多选筛选与同柜标记逻辑
                    FilterCategorySheetByCabinets(app, activeSheet, validCabinets, targetCol, filterKeywords);
                }

                // 准备筛选关键词摘要描述
                string kwSummary = string.Join(", ", filterKeywords.Take(2));
                if (filterKeywords.Count > 2) kwSummary += "...";

                // 创建并推入撤销命令：撤销时清除筛选恢复全貌
                UndoRedoManager.Instance.PushCommand(new ActionUndoableCommand(
                    $"成套查件定位 [{kwSummary}]",
                    () => ClearComponentFilter(activeSheet),
                    () => FilterComponentsBySelection()
                ));
            }
            catch (Exception ex)
            {
                // 记录筛选处理异常日志
                LogHelper.WriteLog($"执行多选行当前列筛选异常: {ex.Message}\r\n{ex.StackTrace}");
                // 弹出异常提示信息
                MessageBox.Show($"筛选执行失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 从用户选区中提取目标列的非空文本作为筛选关键词集合（去重保序）
        /// </summary>
        private static List<string> ExtractFilterKeywordsFromSelection(dynamic selection, int targetCol, dynamic? activeCell)
        {
            // 用于去重判定的哈希集合 (忽略大小写)
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            // 保存有序的关键词结果列表
            List<string> result = new List<string>();

            try
            {
                // 获取选区中的子区域 Areas
                dynamic areas = selection.Areas;
                // 获取子区域数量
                int areaCount = areas.Count;

                // 遍历每一个选区 Area
                for (int a = 1; a <= areaCount; a++)
                {
                    // 获取当前子区域 Range
                    dynamic area = areas[a];
                    // 获取子区域行数与列数
                    int rowCount = area.Rows.Count;
                    // 获取子区域起始物理列与结束物理列
                    int startCol = area.Column;
                    // 计算结束列号
                    int endCol = startCol + area.Columns.Count - 1;

                    // 若该子区域未覆盖目标列 targetCol，则跳过
                    if (targetCol < startCol || targetCol > endCol)
                    {
                        // 不在目标列范围内跳过
                        continue;
                    }

                    // 计算目标列在当前 Area 内部的相对列索引 (1-based)
                    int relCol = targetCol - startCol + 1;

                    // 若行数较多采用数组批量读入内存 (规则 7)
                    if (rowCount > 1)
                    {
                        // 提取目标列在该 Area 中的一维/二维值数组
                        dynamic colRange = area.Columns[relCol];
                        // 一次性读入内存
                        object[,] valMatrix = colRange.Value2 as object[,];

                        // 遍历二维数组行
                        if (valMatrix != null)
                        {
                            // 遍历提取文本
                            for (int r = 1; r <= rowCount; r++)
                            {
                                // 转换为纯文本并去除首尾空白
                                string text = Convert.ToString(valMatrix[r, 1])?.Trim() ?? "";
                                // 若非空且未记录过，加入列表
                                if (!string.IsNullOrWhiteSpace(text) && seen.Add(text))
                                {
                                    // 添加到关键词列表
                                    result.Add(text);
                                }
                            }
                        }
                    }
                    else
                    {
                        // 单行单单元格读取
                        dynamic cell = area.Cells[1, relCol];
                        // 转换为纯文本并去除首尾空白
                        string text = Convert.ToString(cell.Value2)?.Trim() ?? "";
                        // 若非空且未记录过，加入列表
                        if (!string.IsNullOrWhiteSpace(text) && seen.Add(text))
                        {
                            // 添加到关键词列表
                            result.Add(text);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录选区提取异常
                LogHelper.WriteLog($"提取选区关键词异常: {ex.Message}");
            }

            // 若选区遍历后仍为空，尝试读取当前活动单元格 ActiveCell 的值作为兜底
            if (result.Count == 0 && activeCell != null)
            {
                // 读取活动单元格文本
                string fallbackText = Convert.ToString(activeCell.Value2)?.Trim() ?? "";
                // 若非空则加入结果
                if (!string.IsNullOrWhiteSpace(fallbackText))
                {
                    // 加入兜底关键词
                    result.Add(fallbackText);
                }
            }

            // 返回提取出的有效关键词列表
            return result;
        }

        /// <summary>
        /// 分类明细表专属逻辑：遍历箱柜明细块执行多选筛选，识别同柜共存并打上淡青绿高亮标记
        /// </summary>
        private static void FilterCategorySheetByCabinets(
            dynamic app,
            dynamic activeSheet,
            List<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets,
            int targetCol,
            List<string> filterKeywords)
        {
            // 汇总全表所有命中的元器件物理行号
            HashSet<int> allHitRows = new HashSet<int>();
            // 汇总命中了【同箱柜共存】（同时集齐全部筛选关键词）的元器件物理行号
            HashSet<int> sameCabinetMatchedRows = new HashSet<int>();
            // 记录同时集齐全部筛选词的箱柜数量
            int sameCabinetCount = 0;

            // 按照从上到下的顺序收集包含命中行的箱柜数据模型列表: (箱柜键, 箱柜锚点, 命中物理行列表, 是否集齐全部已选元件)
            var hitCabinetsOrdered = new List<(int CabIdx, Models.CabinetAnchorModel Anchor, List<int> HitRows, bool IsFullMatch)>();

            // 遍历工作表中每一个合法箱柜
            foreach (var kv in validCabinets)
            {
                // 提取箱柜锚点数据模型
                var anchor = kv.Value;
                // 校验 Det 与 Subsum 锚点 Range 是否有效
                if (anchor?.Det == null || anchor?.Subsum == null) continue;

                // 提取箱柜标题信息行物理行号
                int detRow = anchor.Det.Row;
                // 提取底部明细小计行物理行号
                int subsumRow = anchor.Subsum.Row;

                // 根据规则 6：Cab_Det.row + 2 为元器件起始行
                int compStartRow = detRow + 2;
                // 根据规则 6：Cab_Subsum.row - 1 为元器件终止行
                int compEndRow = subsumRow - 1;

                // 若该箱柜元器件区间无有效行，跳过当前箱柜
                if (compEndRow < compStartRow) continue;

                // 计算该箱柜元器件总行数
                int compRowCount = compEndRow - compStartRow + 1;

                // 规则 7：采用二维数组一次性将该箱柜目标列的所有单元格数据读取到内存中
                dynamic colRange = activeSheet.Range[activeSheet.Cells[compStartRow, targetCol], activeSheet.Cells[compEndRow, targetCol]];
                // 转换为 object[,] 二维矩阵
                object[,] valMatrix = colRange.Value2 as object[,];
                // 校验数组有效性
                if (valMatrix == null) continue;

                // 记录当前箱柜命中的关键词集合 (去重统计命中种类数)
                HashSet<string> hitKwsInCabinet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                // 记录当前箱柜中所有命中的元器件物理行号
                List<int> hitRowsInCabinet = new List<int>();

                // 遍历该箱柜每一行元器件
                for (int r = 1; r <= compRowCount; r++)
                {
                    // 获取当前行单元格纯文本
                    string cellValue = Convert.ToString(valMatrix[r, 1])?.Trim() ?? "";
                    // 若内容为空则跳过空行
                    if (string.IsNullOrWhiteSpace(cellValue)) continue;

                    // 计算在 Excel 工作表中的真实绝对物理行号
                    int physicalRow = compStartRow + r - 1;

                    // 检查是否命中了用户选中的任一关键词 (必须单元格全匹配，忽略大小写)
                    foreach (string kw in filterKeywords)
                    {
                        // 进行忽略大小写的单元格内容完全精准匹配
                        if (string.Equals(cellValue, kw, StringComparison.OrdinalIgnoreCase))
                        {
                            // 记录命中的关键词类别
                            hitKwsInCabinet.Add(kw);
                            // 记录命中的行号
                            hitRowsInCabinet.Add(physicalRow);
                            // 命中后跳出关键字循环，避免单行重复记录
                            break;
                        }
                    }
                }

                // 若当前箱柜存在命中行
                if (hitRowsInCabinet.Count > 0)
                {
                    // 将当前箱柜命中行并入全表总命中集合
                    foreach (int r in hitRowsInCabinet) allHitRows.Add(r);

                    // 核心业务判定：如果用户选择了 2 个或以上关键词，且当前箱柜【同时集齐了所有筛选关键词】
                    bool isFullSameCabinetMatch = (filterKeywords.Count >= 2 && hitKwsInCabinet.Count >= filterKeywords.Count);

                    // 若判定为同箱柜全部共存匹配
                    if (isFullSameCabinetMatch)
                    {
                        // 累加同柜箱柜计数
                        sameCabinetCount++;
                        // 将该箱柜所有命中行加入同柜高亮行集合
                        foreach (int r in hitRowsInCabinet) sameCabinetMatchedRows.Add(r);
                    }

                    // 登记该命中箱柜在连续有序列表中 (用于相邻箱柜斑马纹交替区分)
                    hitCabinetsOrdered.Add((kv.Key, anchor, hitRowsInCabinet, isFullSameCabinetMatch));
                }
            }

            // 若全表扫描完毕未找到任何匹配行，弹出友好提示并返回
            if (allHitRows.Count == 0)
            {
                // 拼接所选关键词字符串
                string kwStr = string.Join("、", filterKeywords);
                // 弹出提示告知用户未找到匹配项
                MessageBox.Show($"未在当前分类表中检索到匹配 [{kwStr}] 的元器件行！", "筛选结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // 中断退出
                return;
            }

            // 计算整表需要扫描管理的最大物理行号 (取 UsedRange 与各箱柜 Tolsum 的最大值)
            int maxRow = 100;
            try
            {
                // 获取工作表已用区域 UsedRange
                dynamic used = activeSheet.UsedRange;
                // 计算 UsedRange 底部行号
                maxRow = Math.Max(maxRow, (int)used.Row + (int)used.Rows.Count);
            }
            catch { }

            // 遍历箱柜提取最大的 Tolsum 行号
            foreach (var kv in validCabinets)
            {
                // 若 Tolsum 有效，更新最大行号
                if (kv.Value?.Tolsum != null)
                {
                    // 取当前最大值与 Tolsum 行号的最大值
                    maxRow = Math.Max(maxRow, (int)kv.Value.Tolsum.Row + 5);
                }
            }

            // 记录当前的屏幕更新与事件状态，并在进入大批量操作前挂起以保证极速流畅
            bool prevUpdating = true;
            // 记录先前的自动重算模式
            int prevCalculation = XlCalcAutomatic;
            try
            {
                // 保存先前屏幕更新状态
                prevUpdating = app.ScreenUpdating;
                // 冻结屏幕重绘
                app.ScreenUpdating = false;
                // 冻结 COM 事件触发
                app.EnableEvents = false;
                // 读取原计算模式
                prevCalculation = (int)app.Calculation;
                // 切换为手动重算模式，消除万次公式级联重算风暴
                app.Calculation = (dynamic)XlCalcManual;
            }
            catch { }

            try
            {
                // 步骤 0：成套查件智能挂载虚拟表头，无论点击哪个筛选均统一呈现吸顶表头
                bool needsVirtualHeader = NeedsVirtualHeader(activeSheet);
                if (needsVirtualHeader)
                {
                    // 挂载第 1 行虚拟明细表头并开启 24pt 行高与首行吸顶冻结
                    EnsureVirtualHeaderForCategorySheet(activeSheet, app);
                }

                // 步骤 1：若本工作表先前已有快照未还原（如多次连续筛选），先还原旧快照，杜绝用户原色丢失
                RestoreColorSnapshotsForSheet(activeSheet);

                // 步骤 2：对本次命中的所有元器件行 A~M 列执行原始背景色快照备份（100% 保护用户自定义标记色）
                BackupColorSnapshotsForRows(activeSheet, allHitRows);

                // 步骤 3：汇总需要保留显示的物理行集合 (包含第 1 行表头、命中元器件行与所属箱柜的标题行、表头行)
                HashSet<int> keepVisibleRows = new HashSet<int>(allHitRows);
                // 关键点：若挂载了虚拟表头，第 1 行必须加入保留可见集合，杜绝被批量隐藏
                if (needsVirtualHeader)
                {
                    keepVisibleRows.Add(1);
                }

                // 遍历各个命中的箱柜
                foreach (var cabItem in hitCabinetsOrdered)
                {
                    // 校验箱柜 Det 锚点有效性
                    if (cabItem.Anchor?.Det != null)
                    {
                        // 提取箱柜标题信息行物理行号 (如 "1AA1 进线柜")
                        int detRow = cabItem.Anchor.Det.Row;
                        // 保留箱柜信息标题行，让用户一眼看清元器件归属哪台柜
                        keepVisibleRows.Add(detRow);
                        // 保留箱柜表头行 (序号/名称/型号...)，保持列位直观对齐
                        keepVisibleRows.Add(detRow + 1);
                    }
                }

                // 执行“隐藏非保留行，只留下命中元器件及其所属箱柜标题行”
                // 将 1 到 maxRow 中所有不属于 keepVisibleRows 的行批量设置 Hidden = true
                ApplyBatchRowVisibility(activeSheet, 1, maxRow, keepVisibleRows);

                // 核心保障：若挂载了虚拟表头，强制恢复第 1 行行高并解除隐藏且置顶视口
                if (needsVirtualHeader)
                {
                    try
                    {
                        activeSheet.Rows[1].Hidden = false;
                        activeSheet.Rows[1].RowHeight = VirtualHeaderStandardRowHeight;
                        if (app.ActiveWindow != null)
                        {
                            app.ActiveWindow.ScrollRow = 1;
                            app.ActiveWindow.ScrollColumn = 1;
                        }
                    }
                    catch { }
                }

                // 步骤 4：相邻箱柜斑马纹交替上色（第一台淡青底，第二台白底，第三台淡青底...）
                int cyanOle = ColorTranslator.ToOle(AlternateCyanColor);
                // 获取白底 OLE 数值
                int whiteOle = ColorTranslator.ToOle(AlternateWhiteColor);

                // 遍历命中箱柜序列
                for (int cabOrder = 0; cabOrder < hitCabinetsOrdered.Count; cabOrder++)
                {
                    // 提取当前命中箱柜数据
                    var cabItem = hitCabinetsOrdered[cabOrder];
                    // 偶数序号 (0, 2, 4...) 赋予淡青底，奇数序号 (1, 3, 5...) 赋予白底，相邻箱柜边界极其鲜明
                    int targetOleColor = (cabOrder % 2 == 0) ? cyanOle : whiteOle;

                    // 遍历该箱柜内的所有命中行
                    foreach (int row in cabItem.HitRows)
                    {
                        try
                        {
                            // 选取该行 A 列至 M 列核心区域
                            dynamic rowRange = activeSheet.Range[$"A{row}:M{row}"];
                            // 赋予对应箱柜的交替背景底色
                            rowRange.Interior.Color = targetOleColor;
                        }
                        catch { }
                    }
                }

                // 步骤 5：在 Excel 底部状态栏展示醒目的筛选汇总统计
                string statusMsg = (sameCabinetCount > 0)
                    ? $"[筛选完成] 已筛选出 {allHitRows.Count} 行元器件 (共 {hitCabinetsOrdered.Count} 台箱柜，已按白/青底交替区分，其中 {sameCabinetCount} 台同时包含已选元器件)"
                    : $"[筛选完成] 已筛选出 {allHitRows.Count} 行元器件 (共 {hitCabinetsOrdered.Count} 台箱柜，已按白/青底交替区分)";
                // 写入 Excel 状态栏
                try { app.StatusBar = statusMsg; } catch { }
            }
            finally
            {
                // 无论是否发生异常，100% 恢复 Excel 原始渲染、计算与事件环境
                try
                {
                    // 恢复公式计算模式
                    app.Calculation = (dynamic)prevCalculation;
                    // 恢复事件侦听
                    app.EnableEvents = true;
                    // 恢复屏幕刷新
                    app.ScreenUpdating = prevUpdating;
                }
                catch { }
            }
        }

        /// <summary>
        /// 针对普通平铺工作表（无箱柜拓扑）的当前列多选筛选降级逻辑
        /// </summary>
        private static void FilterFlatSheetByKeywords(dynamic app, dynamic activeSheet, int targetCol, List<string> filterKeywords)
        {
            // 获取 UsedRange 范围
            dynamic used = activeSheet.UsedRange;
            // 获取起始行与总行数
            int startRow = used.Row;
            // 计算总行数
            int rowCount = used.Rows.Count;
            // 计算结束行
            int endRow = startRow + rowCount - 1;

            // 若行数极少则不处理
            if (rowCount <= 1) return;

            // 规则 7：一次性读入目标列数据矩阵
            dynamic colRange = activeSheet.Range[activeSheet.Cells[startRow, targetCol], activeSheet.Cells[endRow, targetCol]];
            // 转换为二维矩阵
            object[,] matrix = colRange.Value2 as object[,];
            // 校验矩阵
            if (matrix == null) return;

            // 收集命中的物理行号
            HashSet<int> hitRows = new HashSet<int>();

            // 遍历所有数据行
            for (int r = 1; r <= rowCount; r++)
            {
                // 获取当前行文本
                string val = Convert.ToString(matrix[r, 1])?.Trim() ?? "";
                // 若空则跳过
                if (string.IsNullOrWhiteSpace(val)) continue;

                // 绝对物理行号
                int physicalRow = startRow + r - 1;

                // 检查是否完全匹配任一关键字 (必须单元格全匹配，忽略大小写)
                foreach (string kw in filterKeywords)
                {
                    // 进行忽略大小写的单元格内容完全精准匹配
                    if (string.Equals(val, kw, StringComparison.OrdinalIgnoreCase))
                    {
                        // 记录命中行
                        hitRows.Add(physicalRow);
                        // 结束当前行检查
                        break;
                    }
                }
            }

            // 若无命中行
            if (hitRows.Count == 0)
            {
                // 提示未找到
                MessageBox.Show("未在当前工作表中检索到符合条件的行！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // 退出
                return;
            }

            // 挂起刷新执行批量行显隐
            bool prevUpdating = true;
            try
            {
                // 记录并挂起刷新
                prevUpdating = app.ScreenUpdating;
                // 关闭重绘
                app.ScreenUpdating = false;
                // 关闭事件
                app.EnableEvents = false;

                // 将 1 到 endRow 批量隐藏除 hitRows 以外的所有行
                ApplyBatchRowVisibility(activeSheet, 1, endRow, hitRows);

                // 状态栏提示
                try { app.StatusBar = $"[筛选完成] 已筛选出 {hitRows.Count} 行数据"; } catch { }
            }
            finally
            {
                // 恢复环境
                try
                {
                    // 开启事件
                    app.EnableEvents = true;
                    // 开启屏幕更新
                    app.ScreenUpdating = prevUpdating;
                }
                catch { }
            }
        }

        /// <summary>
        /// 批量应用行可见性：将非保留行聚合为连续物理区间批量执行 EntireRow.Hidden = true，杜绝逐行 COM 卡顿
        /// </summary>
        private static void ApplyBatchRowVisibility(dynamic sheet, int minRow, int maxRow, HashSet<int> keepVisibleRows)
        {
            try
            {
                // 记录当前连续待隐藏区间的起始行号 (-1 表示当前不在隐藏区间内)
                int hideStart = -1;

                // 从 minRow 线性扫描至 maxRow
                for (int r = minRow; r <= maxRow; r++)
                {
                    // 判断当前行是否需要隐藏 (即不属于 keepVisibleRows)
                    bool shouldHide = !keepVisibleRows.Contains(r);

                    // 若当前行需要隐藏
                    if (shouldHide)
                    {
                        // 若尚未开启连续区间，以此行为起点
                        if (hideStart == -1) hideStart = r;
                    }
                    else
                    {
                        // 当前行需要保持显示：若先前存在未闭合的隐藏区间，在此闭合执行
                        if (hideStart != -1)
                        {
                            // 批量设置该连续区间的整行隐藏
                            sheet.Range[$"A{hideStart}:A{r - 1}"].EntireRow.Hidden = true;
                            // 重置区间标记
                            hideStart = -1;
                        }
                        // 确保该保留行处于显示状态
                        sheet.Rows[r].Hidden = false;
                    }
                }

                // 循环结束后，若尾部仍有未闭合的隐藏区间，执行最终批量隐藏
                if (hideStart != -1)
                {
                    // 批量隐藏尾部连续区间
                    sheet.Range[$"A{hideStart}:A{maxRow}"].EntireRow.Hidden = true;
                }
            }
            catch (Exception ex)
            {
                // 记录批量显隐异常
                LogHelper.WriteLog($"批量设置行显隐异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 清除当前工作表的元器件筛选状态：一键恢复所有行显示并精准无损还原用户原本的标记底色
        /// </summary>
        public static void ClearComponentFilter(object? sheet = null)
        {
            try
            {
                // 获取 Excel Application 动态实例
                dynamic? app = ExcelDnaUtil.Application;
                // 校验有效性
                if (app == null) return;

                // 若未显式传入工作表，使用当前活动工作表
                dynamic? targetSheet = sheet ?? app.ActiveSheet;
                // 校验工作表有效性
                if (targetSheet == null) return;

                // 挂起屏幕刷新以保证瞬发恢复
                bool prevUpdating = true;
                try
                {
                    // 保存刷新状态
                    prevUpdating = app.ScreenUpdating;
                    // 暂停屏幕更新
                    app.ScreenUpdating = false;
                    // 暂停事件处理
                    app.EnableEvents = false;

                    // 1. 核心业务升级 (方案 A)：若存在第 1 行虚拟表头快照，无损还原原始行高、内容、样式与窗口冻结
                    bool headerRestored = RestoreVirtualHeaderForSheet(targetSheet, app);

                    // 2. 双轨联动：若当前工作表处于 Excel 系统原生 AutoFilter 筛选状态
                    try
                    {
                        // 若成功还原了虚拟表头，彻底退出原生 AutoFilter 筛选模式，消除残留箭头并让表格回到未筛选原貌
                        if (headerRestored)
                        {
                            // 彻底关闭自动筛选模式
                            targetSheet.AutoFilterMode = false;
                        }
                        else if (targetSheet.FilterMode == true)
                        {
                            // 若为普通平铺表，清除系统筛选条件，显示全部数据并恢复漏斗箭头
                            targetSheet.ShowAllData();
                        }
                    }
                    catch (Exception filterEx)
                    {
                        // 记录清除系统原生 AutoFilter 异常日志
                        LogHelper.WriteLog($"清除原生 AutoFilter 异常: {filterEx.Message}");
                    }

                    // 2. 解除可能由成套查件产生的整表隐藏状态 (安全容错包装，杜绝因 AutoFilter 冲突抛出 1004 阻断后续流程)
                    try
                    {
                        // 一行代码极速解除整表所有行的隐藏状态，100% 恢复全貌
                        targetSheet.Rows.EntireRow.Hidden = false;
                    }
                    catch (Exception hideEx)
                    {
                        // 记录解除隐藏异常日志
                        LogHelper.WriteLog($"解除行隐藏异常: {hideEx.Message}");
                    }

                    // 3. 核心技术点：无损还原单元格原始背景色（完美保留用户原本所有的自定义标记色）
                    bool restoredBySnapshot = RestoreColorSnapshotsForSheet(targetSheet);

                    // 4. 容错兜底：若无快照存在（例如首次打开直接点清除），才执行常规清除
                    if (!restoredBySnapshot)
                    {
                        dynamic? targetWb = null;
                        // 追溯所属工作簿
                        try { targetWb = targetSheet.Parent; } catch { }
                        // 获取有效箱柜列表
                        var validCabinets = Tool.GetSheetValidCabinets((object)targetSheet, (object?)targetWb);

                        // 若存在箱柜，清除箱柜元器件区间背景色
                        if (validCabinets != null && validCabinets.Count > 0)
                        {
                            // 清除箱柜高亮底色
                            ClearCabinetHighlightColors(targetSheet, validCabinets);
                        }
                        else
                        {
                            // 若为普通表，清除已用区域的背景色
                            try { targetSheet.UsedRange.Interior.ColorIndex = XlNoneColorIndex; } catch { }
                        }
                    }

                    // 5. 在状态栏给出恢复提示
                    try { app.StatusBar = "已清除筛选，已恢复显示全部行并还原原有标记底色。"; } catch { }
                }
                finally
                {
                    // 恢复事件与屏幕更新
                    try
                    {
                        // 开启事件
                        app.EnableEvents = true;
                        // 开启屏幕更新
                        app.ScreenUpdating = prevUpdating;
                    }
                    catch { }
                }

                // 5. 交互体验优化：将垂直滚动条移动到活动单元格处于视口中间位置
                try
                {
                    // 获取当前焦点活动单元格 ActiveCell
                    dynamic? activeCell = app.ActiveCell;
                    // 获取 Excel 活动窗口 Window
                    dynamic? win = app.ActiveWindow;
                    // 校验活动单元格与窗口对象有效性
                    if (activeCell != null && win != null)
                    {
                        // 获取活动单元格所在工作表的名称
                        string activeCellSheet = Convert.ToString(activeCell.Worksheet?.Name) ?? "";
                        // 获取当前被取消筛选的目标工作表名称
                        string targetSheetName = Convert.ToString(targetSheet.Name) ?? "";
                        // 确保活动单元格归属于当前目标工作表，杜绝跨表误滚动
                        if (string.Equals(activeCellSheet, targetSheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            // 提取活动单元格所在的绝对物理行号
                            int activeRow = (int)activeCell.Row;
                            // 获取当前视口可见区域的总行数 (备用默认 25 行)
                            // --硬编码: 视口可见行数备用默认值--
                            int visibleRowCount = 25;
                            try
                            {
                                // 动态读取当前窗口可视范围内的总行数
                                visibleRowCount = (int)win.VisibleRange.Rows.Count;
                            }
                            catch { }

                            // 计算使活动单元格居中显示的视口首行号 (最小为 1)
                            int targetScrollRow = Math.Max(1, activeRow - (visibleRowCount / 2));
                            // 将垂直滚动条起始行 ScrollRow 定位到目标居中行
                            win.ScrollRow = targetScrollRow;
                        }
                    }
                }
                catch (Exception scrollEx)
                {
                    // 记录垂直视口居中异常日志 (不阻断正常业务流)
                    LogHelper.WriteLog($"取消筛选调整视口居中异常: {scrollEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // 记录清除筛选异常日志
                LogHelper.WriteLog($"清除筛选异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 对指定的所有行 A~M 列执行原始背景色快照备份（记录 ColorIndex 与 Color，保证无损还原）
        /// </summary>
        private static void BackupColorSnapshotsForRows(dynamic sheet, IEnumerable<int> rows)
        {
            try
            {
                // 获取工作表名称作为快照隔离标识
                string sheetKey = Convert.ToString(sheet.Name) ?? "DefaultSheet";
                // 准备单元格快照列表
                var list = new List<CellColorSnapshot>();

                // 遍历所有待上色的物理行号 (使用 Distinct 规避可能传入的重复行)
                foreach (int row in rows.Distinct())
                {
                    // 遍历 A 列至 M 列 (1 到 13 列) --硬编码: 13 列覆盖成套报价核心列--
                    for (int col = 1; col <= 13; col++)
                    {
                        try
                        {
                            // 获取单元格对象
                            dynamic cell = sheet.Cells[row, col];
                            // 安全读取背景色索引 (转为 int32 避免 COM VARIANT 异常)
                            int cIdx = Convert.ToInt32(cell.Interior.ColorIndex);

                            // 若原先为无填充色
                            if (cIdx == XlNoneColorIndex)
                            {
                                // 登记无填充色记录
                                list.Add(new CellColorSnapshot { Row = row, Col = col, HasNoColor = true });
                            }
                            else
                            {
                                // 登记具体 OLE 颜色数值
                                list.Add(new CellColorSnapshot { Row = row, Col = col, HasNoColor = false, OriginalColor = cell.Interior.Color });
                            }
                        }
                        catch { }
                    }
                }

                // 保存至静态快照缓存
                _filterOriginalColorSnapshots[sheetKey] = list;
            }
            catch (Exception ex)
            {
                // 记录备份异常
                LogHelper.WriteLog($"备份原始底色快照异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 从快照中无损精准还原工作表单元格原始背景色（完美保留用户原本的所有自定义标记色）
        /// </summary>
        private static bool RestoreColorSnapshotsForSheet(dynamic sheet)
        {
            try
            {
                // 获取工作表名称
                string sheetKey = Convert.ToString(sheet.Name) ?? "DefaultSheet";

                // 检查是否存在当前工作表的快照
                if (_filterOriginalColorSnapshots.TryGetValue(sheetKey, out var list) && list != null && list.Count > 0)
                {
                    // 遍历快照记录
                    foreach (var snap in list)
                    {
                        try
                        {
                            // 获取单元格
                            dynamic cell = sheet.Cells[snap.Row, snap.Col];
                            // 若原本为无填充色
                            if (snap.HasNoColor)
                            {
                                // 还原为无填充色
                                cell.Interior.ColorIndex = XlNoneColorIndex;
                            }
                            else if (snap.OriginalColor != null)
                            {
                                // 100% 原样精准还原用户原本自定义的标记色
                                cell.Interior.Color = snap.OriginalColor;
                            }
                        }
                        catch { }
                    }

                    // 清除已还原的快照缓存
                    _filterOriginalColorSnapshots.Remove(sheetKey);
                    // 标记成功还原
                    return true;
                }
            }
            catch (Exception ex)
            {
                // 记录还原异常
                LogHelper.WriteLog($"还原原始底色快照异常: {ex.Message}");
            }

            // 未命中快照
            return false;
        }

        /// <summary>
        /// 清除所有箱柜元器件区间的临时高亮背景色 (安全还原为 xlNone 无填充色)
        /// </summary>
        private static void ClearCabinetHighlightColors(dynamic sheet, List<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets)
        {
            // 遍历每一个箱柜
            foreach (var kv in validCabinets)
            {
                try
                {
                    // 提取锚点
                    var anchor = kv.Value;
                    // 校验锚点 Range
                    if (anchor?.Det == null || anchor?.Subsum == null) continue;

                    // 提取起始行与结束行
                    int compStartRow = anchor.Det.Row + 2;
                    // 计算结束行
                    int compEndRow = anchor.Subsum.Row - 1;

                    // 若行数合法
                    if (compEndRow >= compStartRow)
                    {
                        // 选取该箱柜元器件区间的 A 列至 M 列
                        dynamic compRange = sheet.Range[$"A{compStartRow}:M{compEndRow}"];
                        // 设置背景色为无填充色 (xlNone: -4142)
                        compRange.Interior.ColorIndex = XlNoneColorIndex;
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// 判定当前工作表是否需要挂载第 1 行虚拟明细表头 (零依赖白名单，直接基于排版与内容特征嗅探)
        /// </summary>
        private static bool NeedsVirtualHeader(dynamic sheet)
        {
            try
            {
                // 1. 若第 1 行行高小于 18.0pt (典型如 0~6pt 压扁留白行)，必然需要展开表头
                double r1Height = 0;
                try { r1Height = Convert.ToDouble(sheet.Rows[1].RowHeight); } catch { }
                if (r1Height < 18.0) return true;

                // 2. 检查第 1 行 A~C 列内容是否为常规数据/空白而不是表头
                string a1 = Convert.ToString(sheet.Cells[1, 1].Value2)?.Trim() ?? "";
                string b1 = Convert.ToString(sheet.Cells[1, 2].Value2)?.Trim() ?? "";
                string c1 = Convert.ToString(sheet.Cells[1, 3].Value2)?.Trim() ?? "";

                // 若第 1 行 A、B 列为空，显然是空白行，需要挂载表头
                if (string.IsNullOrWhiteSpace(a1) && string.IsNullOrWhiteSpace(b1)) return true;

                // 若第 1 行 A 列不是“序号/NO”且 B 列不是“名称/元件”，判定为非表头行
                bool isStandardHeaderRow1 = (a1.Contains("序号") || a1.Contains("NO") || a1.Contains("No")) &&
                                            (b1.Contains("名称") || b1.Contains("元件") || c1.Contains("型号") || c1.Contains("规格"));
                if (!isStandardHeaderRow1) return true;
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 从工作表内部动态扫描嗅探真实的元器件明细表头所在物理行号 (零依赖白名单)
        /// </summary>
        private static int FindDetailHeaderRowInSheet(dynamic sheet, int maxScanRows = 120)
        {
            try
            {
                // 限制最大扫描行数
                int endRow = Math.Min(maxScanRows, 150);
                // 选取前 endRow 行的 A:E 列区域
                dynamic scanRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[endRow, 5]];
                // 规则 7: 一次性读取到二维数组
                object[,] scanVals = scanRange.Value2 as object[,];
                if (scanVals == null) return 0;

                int rowCount = scanVals.GetLength(0);
                // 从第 2 行开始往下逐行扫描特征
                for (int r = 2; r <= rowCount; r++)
                {
                    string colA = Convert.ToString(scanVals[r, 1])?.Trim() ?? "";
                    string colB = Convert.ToString(scanVals[r, 2])?.Trim() ?? "";
                    string colC = Convert.ToString(scanVals[r, 3])?.Trim() ?? "";
                    string colD = Convert.ToString(scanVals[r, 4])?.Trim() ?? "";

                    // 特征匹配：A 列含“序号/项次/NO”且 B 列含“名称/元件”或 C 列含“型号/规格”
                    bool isHeader = (colA.Contains("序号") || colA.Contains("项次") || colA.Contains("NO") || colA.Contains("No")) &&
                                    (colB.Contains("名称") || colB.Contains("元件") || colC.Contains("型号") || colC.Contains("规格") || colD.Contains("厂家"));

                    // 容错特征：某行直接包含“元件名称”与“型号规格”
                    if (!isHeader && (colB.Contains("元件") || colB.Contains("名称")) && (colC.Contains("型号") || colC.Contains("规格")))
                    {
                        isHeader = true;
                    }

                    // 若匹配成功，返回在工作表中的真实物理行号
                    if (isHeader)
                    {
                        return r;
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"嗅探明细表头行异常: {ex.Message}");
            }
            return 0;
        }

        /// <summary>
        /// 为成套分类表构建并呈现第 1 行虚拟明细表头，并设置自适应 24pt 行高与吸顶冻结窗格 (方案 A 核心承载，零依赖白名单)
        /// </summary>
        /// <param name="activeSheet">当前活动工作表 COM 实例</param>
        /// <param name="app">Excel Application COM 实例</param>
        private static void EnsureVirtualHeaderForCategorySheet(dynamic activeSheet, dynamic app)
        {
            try
            {
                // 获取当前工作表名称
                string sheetKey = Convert.ToString(activeSheet.Name) ?? "DefaultSheet";
                LogHelper.WriteLog($"[虚拟表头] 触发 EnsureVirtualHeaderForCategorySheet，目标表: [{sheetKey}]");

                // 计算整表需要覆盖的最大列数 (至少 25 列，优先探测已用区域列数)
                int maxCols = 25;
                try
                {
                    // 获取已用区域
                    dynamic used = activeSheet.UsedRange;
                    // 校验已用区域列边界
                    if (used != null)
                    {
                        // 计算已用区域最大物理列号
                        int usedEndCol = (int)used.Column + (int)used.Columns.Count - 1;
                        // 取较大值确保整表各列全覆盖
                        maxCols = Math.Max(maxCols, usedEndCol);
                    }
                }
                catch { }

                // 选取第 1 行 A 列至 maxCols 列对应的单元格区域 Range
                dynamic row1Range = activeSheet.Range[activeSheet.Cells[1, 1], activeSheet.Cells[1, maxCols]];

                // 若尚未记录该表的快照，执行完整备份
                if (!_virtualHeaderSnapshots.ContainsKey(sheetKey))
                {
                    // 检查原始第 1 行是否存在单元格合并
                    bool isMerged = false;
                    try { isMerged = Convert.ToBoolean(row1Range.MergeCells); } catch { }

                    // 构建快照对象并完整记录第 1 行原本的全部状态与排版属性
                    var snapshot = new VirtualHeaderSnapshot
                    {
                        // 记录工作表名
                        SheetName = sheetKey,
                        // 记录原始行高
                        OriginalRowHeight = activeSheet.Rows[1].RowHeight,
                        // 规则 7: 一次性批量读入第 1 行原始值
                        OriginalValues = row1Range.Value2 as object[,],
                        // 规则 7: 一次性批量读入第 1 行原始公式
                        OriginalFormulas = row1Range.Formula as object[,],
                        // 记录原始背景色
                        OriginalInteriorColor = row1Range.Interior.Color,
                        // 记录原始背景色索引
                        OriginalInteriorColorIndex = row1Range.Interior.ColorIndex,
                        // 记录原始字体粗细
                        OriginalFontBold = row1Range.Font.Bold,
                        // 记录原始字体大小
                        OriginalFontSize = row1Range.Font.Size,
                        // 记录原始水平对齐
                        OriginalHorizontalAlignment = row1Range.HorizontalAlignment,
                        // 记录原始垂直对齐
                        OriginalVerticalAlignment = row1Range.VerticalAlignment,
                        // 记录原始是否合并
                        OriginalMergeCells = isMerged,
                        // 记录最大列数
                        ColumnCount = maxCols
                    };

                    // 备份当前的窗口冻结窗格配置
                    try
                    {
                        // 获取当前活动窗口句柄
                        dynamic win = app.ActiveWindow;
                        // 校验活动窗口句柄有效性
                        if (win != null)
                        {
                            // 记录原始冻结状态
                            snapshot.OriginalFreezePanes = win.FreezePanes;
                            // 记录原始拆分行
                            snapshot.OriginalSplitRow = win.SplitRow;
                            // 记录原始拆分列
                            snapshot.OriginalSplitColumn = win.SplitColumn;
                        }
                    }
                    catch { }

                    // 保存快照至全局字典
                    _virtualHeaderSnapshots[sheetKey] = snapshot;
                }

                // 核心安全保障：若第 1 行存在合并单元格，先取消合并以允许向各列填入独立列名
                try
                {
                    if (Convert.ToBoolean(row1Range.MergeCells))
                    {
                        row1Range.UnMerge();
                    }
                }
                catch { }

                // 动态探测当前表中真实的元器件明细表头所在物理行 (零依赖白名单)
                int detectedHeaderRow = FindDetailHeaderRowInSheet(activeSheet, 120);
                object[,] newHeaders = new object[1, maxCols];
                bool extractedFromDetail = false;

                // 若成功探测到明细表头物理行
                if (detectedHeaderRow > 0)
                {
                    try
                    {
                        // 选取探测到的明细表头 Range
                        dynamic srcHeaderRange = activeSheet.Range[activeSheet.Cells[detectedHeaderRow, 1], activeSheet.Cells[detectedHeaderRow, maxCols]];
                        // 规则 7: 一次性读取真实表头文本到内存
                        object[,] srcVals = srcHeaderRange.Value2 as object[,];

                        // 校验表头数组有效性
                        if (srcVals != null)
                        {
                            // 遍历各列提取真实列名
                            for (int c = 1; c <= maxCols; c++)
                            {
                                // 转换为纯文本并去除首尾空白
                                string colText = Convert.ToString(srcVals[1, c])?.Trim() ?? "";
                                // 若 A 列带有动态序号公式结果(如 "序号1" 或 "序号Cab_Sum_1")，规整为标准纯文字“序号”
                                if (c == 1 && colText.Contains("序号"))
                                {
                                    colText = "序号";
                                }
                                // 回填至新表头二维数组
                                newHeaders[0, c - 1] = colText;
                            }
                            // 标记提取成功
                            extractedFromDetail = true;
                            LogHelper.WriteLog($"[虚拟表头] 成功从第 {detectedHeaderRow} 行动态抓取真实明细表头！");
                        }
                    }
                    catch (Exception extractEx)
                    {
                        // 记录提取异常日志
                        LogHelper.WriteLog($"从检测行 {detectedHeaderRow} 提取表头异常: {extractEx.Message}");
                    }
                }

                // 容错兜底：若未能成功提取，使用成套标准默认列名数组填充
                if (!extractedFromDetail)
                {
                    // 遍历填充标准列名
                    for (int c = 0; c < maxCols; c++)
                    {
                        // 优先填充标准预设列名
                        newHeaders[0, c] = (c < DefaultCategoryDetailHeaders.Length) ? DefaultCategoryDetailHeaders[c] : "";
                    }
                    LogHelper.WriteLog($"[虚拟表头] 未嗅探到明细表头行，采用成套标准默认列名回填。");
                }

                // 规则 7: 一次性将虚拟表头二维数组批量写入第 1 行单元格
                row1Range.Value2 = newHeaders;

                // 强制解除第 1 行隐藏状态
                activeSheet.Rows[1].Hidden = false;
                // 强制调整第 1 行行高为标准 24pt (充裕容纳文字与下拉箭头)
                activeSheet.Rows[1].RowHeight = VirtualHeaderStandardRowHeight;

                // 赋予第 1 行浅灰青底色，与成套主题色 #009688 深度协调
                row1Range.Interior.Color = ColorTranslator.ToOle(VirtualHeaderBgColor);

                // 设置字体加粗
                row1Range.Font.Bold = true;
                // 设置字体大小为 10pt
                row1Range.Font.Size = 10;
                // 设置水平居中 (-4108: xlCenter)
                // --硬编码: Excel xlCenter 常数--
                row1Range.HorizontalAlignment = -4108;
                // 设置垂直居中 (-4108: xlCenter)
                // --硬编码: Excel xlCenter 常数--
                row1Range.VerticalAlignment = -4108;

                // 视觉增强：开启首行冻结 (Freeze Top Row)，向下滚动查阅数据时表头始终吸顶固定
                try
                {
                    // 确保活动工作表被激活
                    activeSheet.Activate();
                    // 获取当前窗口
                    dynamic win = app.ActiveWindow;
                    // 校验窗口有效性
                    if (win != null)
                    {
                        // 先解除可能存在的旧冻结
                        win.FreezePanes = false;
                        // 核心保障：必须先将视口滚动到第 1 行第 1 列，让第 1 行稳稳处于可视区域最顶端！
                        win.ScrollRow = 1;
                        win.ScrollColumn = 1;
                        // 设置拆分行为第 1 行
                        win.SplitRow = 1;
                        // 拆分列设为 0
                        win.SplitColumn = 0;
                        // 开启窗口冻结
                        win.FreezePanes = true;
                    }
                }
                catch (Exception freezeEx)
                {
                    // 记录冻结异常日志
                    LogHelper.WriteLog($"开启首行冻结窗格异常: {freezeEx.Message}");
                }

                LogHelper.WriteLog($"[虚拟表头] 第 1 行虚拟表头挂载完成！行高: {activeSheet.Rows[1].RowHeight}pt");
            }
            catch (Exception ex)
            {
                // 记录构建虚拟表头全局异常日志
                LogHelper.WriteLog($"构建虚拟表头全局异常: {ex.Message}\r\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// 从快照中无损精准还原工作表第 1 行原始状态（行高、内容、公式、样式与冻结窗格）
        /// </summary>
        /// <param name="sheet">目标工作表 COM 实例</param>
        /// <param name="app">Excel Application COM 实例</param>
        /// <returns>若成功执行还原返回 true，否则返回 false</returns>
        private static bool RestoreVirtualHeaderForSheet(dynamic sheet, dynamic? app = null)
        {
            try
            {
                // 获取工作表名称
                string sheetKey = Convert.ToString(sheet.Name) ?? "DefaultSheet";

                // 检查是否存在该工作表的虚拟表头快照
                if (_virtualHeaderSnapshots.TryGetValue(sheetKey, out var snapshot) && snapshot != null)
                {
                    // 获取记录的最大列数
                    int cols = snapshot.ColumnCount > 0 ? snapshot.ColumnCount : 25;
                    // 选取第 1 行对应区域 Range
                    dynamic row1Range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, cols]];

                    // 1. 还原公式或原始数值
                    if (snapshot.OriginalFormulas != null)
                    {
                        // 规则 7: 批量还原原始公式
                        row1Range.Formula = snapshot.OriginalFormulas;
                    }
                    else if (snapshot.OriginalValues != null)
                    {
                        // 规则 7: 批量还原原始值
                        row1Range.Value2 = snapshot.OriginalValues;
                    }
                    else
                    {
                        // 清空第 1 行写入的表头文本
                        row1Range.ClearContents();
                    }

                    // 2. 还原背景底色
                    if (snapshot.OriginalInteriorColorIndex != null && Convert.ToInt32(snapshot.OriginalInteriorColorIndex) == XlNoneColorIndex)
                    {
                        // 还原为无填充色
                        row1Range.Interior.ColorIndex = XlNoneColorIndex;
                    }
                    else if (snapshot.OriginalInteriorColor != null)
                    {
                        // 还原为原始具体颜色
                        row1Range.Interior.Color = snapshot.OriginalInteriorColor;
                    }

                    // 3. 还原字体加粗与大小
                    if (snapshot.OriginalFontBold != null) row1Range.Font.Bold = snapshot.OriginalFontBold;
                    if (snapshot.OriginalFontSize != null) row1Range.Font.Size = snapshot.OriginalFontSize;
                    // 还原水平与垂直对齐方式
                    if (snapshot.OriginalHorizontalAlignment != null) row1Range.HorizontalAlignment = snapshot.OriginalHorizontalAlignment;
                    if (snapshot.OriginalVerticalAlignment != null) row1Range.VerticalAlignment = snapshot.OriginalVerticalAlignment;

                    // 4. 若原本存在合并单元格，还原合并
                    if (snapshot.OriginalMergeCells)
                    {
                        try { row1Range.Merge(); } catch { }
                    }

                    // 5. 还原第 1 行原始行高 (如原本只有几像素的极窄行高)
                    if (snapshot.OriginalRowHeight != null)
                    {
                        // 设置回原始行高
                        sheet.Rows[1].RowHeight = snapshot.OriginalRowHeight;
                    }

                    // 6. 还原原始窗口冻结窗格配置
                    if (app != null)
                    {
                        try
                        {
                            // 获取活动窗口句柄
                            dynamic win = app.ActiveWindow;
                            // 校验窗口句柄有效性
                            if (win != null)
                            {
                                // 若原本处于冻结状态
                                if (snapshot.OriginalFreezePanes == true)
                                {
                                    // 先关闭当前冻结
                                    win.FreezePanes = false;
                                    // 还原拆分行
                                    win.SplitRow = snapshot.OriginalSplitRow;
                                    // 还原拆分列
                                    win.SplitColumn = snapshot.OriginalSplitColumn;
                                    // 重新激活冻结
                                    win.FreezePanes = true;
                                }
                                else
                                {
                                    // 若原本无冻结，彻底关闭冻结窗格
                                    win.FreezePanes = false;
                                    // 重置拆分行
                                    win.SplitRow = 0;
                                    // 重置拆分列
                                    win.SplitColumn = 0;
                                }
                            }
                        }
                        catch (Exception winEx)
                        {
                            // 记录还原窗口冻结异常日志
                            LogHelper.WriteLog($"还原窗口冻结异常: {winEx.Message}");
                        }
                    }

                    // 从静态字典中移除已还原的快照缓存
                    _virtualHeaderSnapshots.Remove(sheetKey);
                    // 标记还原成功
                    return true;
                }
            }
            catch (Exception ex)
            {
                // 记录还原虚拟表头异常日志
                LogHelper.WriteLog($"还原虚拟表头异常: {ex.Message}");
            }

            // 未命中快照返回 false
            return false;
        }
    }
}
