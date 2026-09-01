using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel COM 事件与上下文菜单生命周期统一管理器
    /// 集中管理 SheetChange、SheetFollowHyperlink、SheetBeforeRightClick 以及右键菜单 CommandBars 控制
    /// </summary>
    public static class ExcelEventManager
    {
        // 保持对 Excel Application 实例的静态强引用，防止 GC 回收 COM 事件下沉节点
        private static Microsoft.Office.Interop.Excel.Application? _excelApp = null;

        // 保存对右键菜单按钮的强引用集合，防止 COM 事件下沉委托被 GC 提前回收导致点击无响应
        private static readonly List<dynamic> _contextMenuButtons = new List<dynamic>();

        /// <summary>
        /// 注册 Excel 全局事件（SheetChange, SheetFollowHyperlink, SheetBeforeRightClick）
        /// </summary>
        public static void RegisterEvents()
        {
            try
            {
                // 获取并保持 Excel Application 静态引用，避免被 GC 回收 (安全调用)
                _excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaSafeAccessor.GetApplication();

                // 校验 _excelApp 对象有效性
                if (_excelApp == null) return;

                // 强行开启 Excel 系统级 EnableEvents 选项
                _excelApp.EnableEvents = true;

                // 先解除已有委托绑定，避免叠加；再重新绑定 SheetChange 事件处理函数
                _excelApp.SheetChange -= OnSheetChange;
                _excelApp.SheetChange += OnSheetChange;

                // 解除已有的 SheetFollowHyperlink 事件处理委托绑定
                _excelApp.SheetFollowHyperlink -= OnSheetFollowHyperlink;
                // 重新绑定 SheetFollowHyperlink 事件处理委托，实现跳转后的 ScrollRow 偏移定位
                _excelApp.SheetFollowHyperlink += OnSheetFollowHyperlink;

                // 解除已有的 SheetBeforeRightClick 事件处理委托绑定，避免重复挂载
                _excelApp.SheetBeforeRightClick -= OnSheetBeforeRightClick;
                // 重新绑定 SheetBeforeRightClick 事件处理委托，实现在第一个 Cab_Det 上方右击添加“新建箱柜”按钮
                _excelApp.SheetBeforeRightClick += OnSheetBeforeRightClick;

                // 解除已有的 SheetSelectionChange 事件处理委托绑定，避免重复挂载
                _excelApp.SheetSelectionChange -= OnSheetSelectionChange;
                // 重新绑定 SheetSelectionChange 事件处理委托，实现进入 C 列智能输入或 D 列物料联想下拉
                _excelApp.SheetSelectionChange += OnSheetSelectionChange;

                // 解除已有的 SheetBeforeDoubleClick 事件绑定
                _excelApp.SheetBeforeDoubleClick -= OnSheetBeforeDoubleClick;
                // 重新绑定 SheetBeforeDoubleClick 事件，双击 D 列单元格可快速触发物料智能联想
                _excelApp.SheetBeforeDoubleClick += OnSheetBeforeDoubleClick;
            }
            catch (Exception ex)
            {
                // 弹出注册异常提示帮助诊断 (--硬编码: 弹窗标题与提示文本--)
                System.Windows.Forms.MessageBox.Show($"注册 Excel 事件失败: {ex.Message}", "系统提示");
            }
        }

        /// <summary>
        /// 注销所有已注册的 Excel 全局事件，并清理自定义右键菜单项
        /// </summary>
        public static void UnregisterEvents()
        {
            try
            {
                if (_excelApp != null)
                {
                    // 解除 SheetChange 事件绑定
                    _excelApp.SheetChange -= OnSheetChange;
                    // 解除 SheetFollowHyperlink 事件绑定
                    _excelApp.SheetFollowHyperlink -= OnSheetFollowHyperlink;
                    // 解除 SheetBeforeRightClick 事件绑定
                    _excelApp.SheetBeforeRightClick -= OnSheetBeforeRightClick;
                    // 解除 SheetSelectionChange 事件绑定
                    _excelApp.SheetSelectionChange -= OnSheetSelectionChange;
                    // 解除 SheetBeforeDoubleClick 事件绑定
                    _excelApp.SheetBeforeDoubleClick -= OnSheetBeforeDoubleClick;
                }

                // 隐藏覆盖输入框与物料联想下拉浮窗
                ExcelServices.HideSmartInputOverlay();
                ExcelServices.HideComponentMatchOverlay();

                // 彻底清理注册的右键菜单控件
                RemoveContextMenuControls();
            }
            catch { }
        }

        /// <summary>
        /// 响应工作表单元格焦点切换事件，实现选中 C 列智能输入或选中 D 列“点击查询”时激活智能下拉
        /// </summary>
        private static void OnSheetSelectionChange(object shObj, Microsoft.Office.Interop.Excel.Range target)
        {
            try
            {
                // 校验目标单元格与全局 Application
                if (target == null || _excelApp == null) return;

                // 1. 判断选区是否包含 C 列 (第 3 列: 规格型号) 或选中了整行
                int startCol = target.Column;
                int endCol = startCol + target.Columns.Count - 1;
                bool containsColumnC = (3 >= startCol && 3 <= endCol);

                // 若选区包含 C 列且开启了 CAD 联动：提取选区内所有行的 AD 列 (第 30 列) / AA 列 (第 27 列兼容) 句柄并防抖推送
                if (containsColumnC && Services.CadSyncClient.SyncToCadEnabled)
                {
                    try
                    {
                        // 获取当前工作表引用
                        var ws = (shObj as Microsoft.Office.Interop.Excel.Worksheet) ?? target.Worksheet;
                        if (ws != null)
                        {
                            int startRow = target.Row;
                            int endRow = startRow + target.Rows.Count - 1;
                            // 限制单次最大多选行数上限（50 行），防止误选全表引起额外开销
                            if (target.Rows.Count > 50) endRow = startRow + 49;

                            List<string> handleList = new List<string>();

                            // 遍历所选的所有行
                            for (int r = startRow; r <= endRow; r++)
                            {
                                // 读取 AD 列 (第 30 列) 的 CAD 句柄字符串
                                string rawHandles = Convert.ToString(ws.Range[$"AD{r}"].Value2)?.Trim() ?? string.Empty;
                                if (!string.IsNullOrEmpty(rawHandles))
                                {
                                    // 兼容两级分隔符（逗号“,”与连字符“-”）进行拆分提取
                                    string[] parts = rawHandles.Split(new[] { ',', ';', '-', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (var p in parts)
                                    {
                                        string cleanP = p.Trim();
                                        if (!string.IsNullOrEmpty(cleanP) && !handleList.Contains(cleanP))
                                        {
                                            handleList.Add(cleanP);
                                        }
                                    }
                                }
                            }

                            // 带有 50ms 防抖与自动缩放视角（autoZoom=true）推送至 CAD 管道
                            Services.CadSyncClient.SendHandlesDebounced(handleList, true);
                        }
                    }
                    catch { }
                }

                // 2. 处理 UI 浮窗交互：若选中的是单个单元格
                if (target.Rows.Count == 1 && target.Columns.Count == 1)
                {
                    int col = target.Column;

                    // 2.1 若选中的是 C 列 (第 3 列: 规格型号 / 元件汇总表原型号规格)
                    if (col == 3)
                    {
                        // 隐藏物料智能联想下拉浮窗
                        ExcelServices.HideComponentMatchOverlay();
                        // 获取当前工作表名称
                        string curSheetName = (shObj as Microsoft.Office.Interop.Excel.Worksheet)?.Name ?? target.Worksheet?.Name ?? string.Empty;
                        // 仅在非“元件汇总表”的普通分类表中激活智能输入覆盖框 (元件汇总表 C 列专用于基准查看与 CAD 夹点联动)
                        if (curSheetName != ComponentMatchDefaults.ComponentSummarySheetName)
                        {
                            // 弹出智能输入覆盖框
                            ExcelServices.ShuRu(target);
                        }
                        else
                        {
                            // 在元件汇总表中隐藏智能输入覆盖框，确保纯粹的 CAD 夹点联动体验
                            ExcelServices.HideSmartInputOverlay();
                        }
                    }
                    // 2.2 若选中的是 D 列 (第 4 列: 规格型号/点击查询) -> 检查是否在“元件汇总表”中并触发物料智能联想下拉
                    else if (col == 4)
                    {
                        // 隐藏智能输入覆盖框
                        ExcelServices.HideSmartInputOverlay();

                        // 获取当前工作表名称
                        string curSheetName = (shObj as Microsoft.Office.Interop.Excel.Worksheet)?.Name ?? target.Worksheet?.Name ?? string.Empty;

                        // 仅限定在“元件汇总表”下触发
                        if (curSheetName == ComponentMatchDefaults.ComponentSummarySheetName)
                        {
                            // 读取当前单元格文本值
                            string cellVal = Convert.ToString(target.Value2)?.Trim() ?? string.Empty;

                            // 若单元格内容包含“点击查询”或者为空/可查询时弹出物料智能联想下拉
                            if (cellVal.Contains(ComponentMatchDefaults.MultipleCandidatesText) || string.IsNullOrEmpty(cellVal))
                            {
                                ExcelServices.ShowComponentMatchOverlay(target);
                            }
                            else
                            {
                                ExcelServices.HideComponentMatchOverlay();
                            }
                        }
                        else
                        {
                            // 非“元件汇总表”隐藏物料下拉浮窗
                            ExcelServices.HideComponentMatchOverlay();
                        }
                    }
                    else
                    {
                        // 离开 C/D 列时全部隐藏
                        ExcelServices.HideSmartInputOverlay();
                        ExcelServices.HideComponentMatchOverlay();
                    }
                }
                else
                {
                    // 选中多单元格区域时隐藏浮窗
                    ExcelServices.HideSmartInputOverlay();
                    ExcelServices.HideComponentMatchOverlay();
                }
            }
            catch
            {
                // 异常时安全兜底隐藏
                ExcelServices.HideSmartInputOverlay();
                ExcelServices.HideComponentMatchOverlay();
            }
        }

        /// <summary>
        /// 响应工作表单元格双击事件，在“元件汇总表”且开启搜索时双击 D 列单元格弹出物料联想下拉
        /// </summary>
        private static void OnSheetBeforeDoubleClick(object shObj, Microsoft.Office.Interop.Excel.Range target, ref bool cancel)
        {
            try
            {
                // 校验有效性
                if (target == null || _excelApp == null) return;

                // 获取当前双击所在的工作表名称
                string sheetName = (shObj as Microsoft.Office.Interop.Excel.Worksheet)?.Name ?? target.Worksheet?.Name ?? string.Empty;

                // 核心条件：必须在“元件汇总表”工作表中，且双击的是 D 列单个单元格
                if (sheetName == ComponentMatchDefaults.ComponentSummarySheetName &&
                    target.Rows.Count == 1 && target.Columns.Count == 1 && target.Column == 4)
                {
                    // 先校验当前是否开启了“搜索”功能
                    var cfg = ExcelServices.LoadComponentMatchFilterConfig();
                    if (cfg != null && cfg.EnableSearchOverlay)
                    {
                        // 仅在开启“搜索”时拦截双击并弹起智能联想下拉
                        cancel = true;
                        ExcelServices.ShowComponentMatchOverlay(target);
                    }
                    // 若未开启“搜索”，保持 cancel = false，允许 Excel 正常双击进入单元格进行文本编辑
                }
            }
            catch { }
        }

        /// <summary>
        /// 响应点击超链接事件，实现跳转后视图 ScrollRow 偏移定位
        /// </summary>
        private static void OnSheetFollowHyperlink(object shObj, Microsoft.Office.Interop.Excel.Hyperlink target)
        {
            try
            {
                // 校验全局 _excelApp 句柄有效性
                if (_excelApp == null) return;

                // 获取当前活动窗口强类型对象
                Microsoft.Office.Interop.Excel.Window win = (Microsoft.Office.Interop.Excel.Window)_excelApp.ActiveWindow;
                // 校验窗口句柄有效性
                if (win == null) return;

                // 获取跳转后选中的焦点单元格 Range 强类型对象
                Microsoft.Office.Interop.Excel.Range activeCell = (Microsoft.Office.Interop.Excel.Range)_excelApp.ActiveCell;
                // 校验焦点单元格句柄有效性
                if (activeCell == null) return;

                // 获取目标单元格的物理行号
                int targetRow = activeCell.Row;
                // 从 ConfigManager 全局配置中读取 ScrollRowOffset 偏移行数修正值 (默认 -3)
                int scrollOffset = ConfigManager.Instance.Current.Excel.ScrollRowOffset;

                // 使用 config 中的 ScrollRowOffset 修正计算视图首行行号 (兜底保障行号不小于 1)
                int targetScrollRow = Math.Max(1, targetRow + scrollOffset);

                // 将计算后的修正行号赋予窗口可视起始行 ScrollRow
                win.ScrollRow = targetScrollRow;
            }
            catch { }
        }

        /// <summary>
        /// 响应单元格修改事件，实现箱柜关键行之间的双向数据同步与规格型号智能联动回填
        /// </summary>
        private static void OnSheetChange(object shObj, Microsoft.Office.Interop.Excel.Range target)
        {
            try
            {
                // 校验 app 与 target 句柄有效性
                if (_excelApp == null || target == null) return;

                // 转换目标 Sheet 工作表对象
                Microsoft.Office.Interop.Excel.Worksheet? sh = shObj as Microsoft.Office.Interop.Excel.Worksheet;
                // 校验 sh 对象有效性
                if (sh == null) return;

                // 安全获取改动单元格所在工作簿对象
                Microsoft.Office.Interop.Excel.Workbook? wb = sh.Parent as Microsoft.Office.Interop.Excel.Workbook;
                // 校验 wb 句柄有效性
                if (wb == null) return;

                // 1. 优先尝试处理箱柜关键行（Sum 汇总行、Det 明细行、Tolsum 总计行）之间的 8 组双向数据绑定联动
                if (TryHandleCabinetBiDirectionalSync(wb, sh, target))
                {
                    // 若命中双向绑定同步，直接退出避免重复触发
                    return;
                }

                // 2. 处理第 3 列 (C列 - 元器件规格型号) 修改时的智能属性联动回填
                if (target.Column == 3 && target.Cells.Count == 1)
                {
                    // 读取 C 列最新输入的规格型号字符串
                    string newModel = Convert.ToString(target.Value)?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(newModel))
                    {
                        // 异步安全触发或同步读取智能输入控制器配置
                        var ctrl = new ExcelAddInDemo.Controllers.SmartInputController();
                        var config = ctrl.GetConfig();

                        // 若未勾选任何回填字段则不进行联动
                        if (config != null && (config.FillName || config.FillManufacturer || config.FillUnit || config.FillUnitPrice))
                        {
                            // 读取元器件缓存数据
                            var storage = ctrl.GetStoredComponents();
                            if (storage != null && storage.Sheets != null)
                            {
                                // 筛选已选工作表中的元器件
                                var selectedSheets = config.SelectedSheets != null && config.SelectedSheets.Count > 0
                                    ? config.SelectedSheets
                                    : storage.Sheets.Select(s => s.SheetName).ToList();

                                // 查找匹配的规格型号
                                ExcelAddInDemo.Models.SmartComponentItem? matchedItem = null;
                                foreach (var sData in storage.Sheets.Where(s => selectedSheets.Contains(s.SheetName)))
                                {
                                    matchedItem = sData.Components?.FirstOrDefault(c => string.Equals(c.Model, newModel, StringComparison.OrdinalIgnoreCase));
                                    if (matchedItem != null) break;
                                }

                                // 若找到了对应的物料属性
                                if (matchedItem != null)
                                {
                                    // 暂停事件触发避免循环调用
                                    _excelApp.EnableEvents = false;
                                    try
                                    {
                                        int r = target.Row;
                                        dynamic dynSh = sh;
                                        // 联动 B列 (元件名称)
                                        if (config.FillName && !string.IsNullOrEmpty(matchedItem.Name))
                                        {
                                            dynSh.Cells[r, 2].Value = matchedItem.Name;
                                        }
                                        // 联动 D列 (生产厂家)
                                        if (config.FillManufacturer && !string.IsNullOrEmpty(matchedItem.Manufacturer))
                                        {
                                            dynSh.Cells[r, 4].Value = matchedItem.Manufacturer;
                                        }
                                        // 联动 E列 (计量单位)
                                        if (config.FillUnit && !string.IsNullOrEmpty(matchedItem.Unit))
                                        {
                                            dynSh.Cells[r, 5].Value = matchedItem.Unit;
                                        }
                                        // 联动 G列 (销售单价)
                                        if (config.FillUnitPrice && matchedItem.UnitPrice > 0)
                                        {
                                            dynSh.Cells[r, 7].Value = matchedItem.UnitPrice;
                                        }
                                    }
                                    catch { }
                                    finally
                                    {
                                        // 恢复事件触发机制
                                        _excelApp.EnableEvents = true;
                                    }
                                }
                            }
                        }
                    }
                    return;
                }
            }
            catch { }
        }

        /// <summary>
        /// 尝试处理箱柜关键行（Sum 汇总行、Det 明细信息行、Tolsum 总计行）之间的 8 组双向数据同步
        /// 包含:
        /// 1. Sum B (2) <-> Det B (2) [柜号]
        /// 2. Sum C (3) <-> Det G (7) [箱柜名称，Det G:H合并]
        /// 3. Sum D (4) <-> Det D (4) [箱柜型号，Det D:E合并]
        /// 4. Sum F (6) <-> Tolsum F (6) [数量/台数]
        /// 5. Sum I (9) <-> Det I (9) [备注]
        /// 6. Sum M (13) <-> Det K (11) [箱柜尺寸]
        /// 7. Sum N (14) <-> Det M (13) [类别]
        /// 8. Sum Q (17) <-> Det O (15) [图号]
        /// </summary>
        /// <param name="wb">当前工作簿对象</param>
        /// <param name="sh">当前工作表对象</param>
        /// <param name="target">被修改的目标单元格 Range</param>
        /// <returns>若成功命中并处理了双向同步返回 true，否则返回 false</returns>
        private static bool TryHandleCabinetBiDirectionalSync(
            Microsoft.Office.Interop.Excel.Workbook wb,
            Microsoft.Office.Interop.Excel.Worksheet sh,
            Microsoft.Office.Interop.Excel.Range target)
        {
            try
            {
                // 校验关键句柄有效性
                if (wb == null || sh == null || target == null || _excelApp == null) return false;

                // 获取改动单元格的物理列号与行号
                int col = target.Column;
                int row = target.Row;

                // 从 CabinetPrefixConfig 读取当前配置的定义名称前缀
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

                // 尝试获取当前行第 1 列 (A 列) 单元格
                Microsoft.Office.Interop.Excel.Range? aCell = null;
                try { aCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[row, 1]; } catch { }

                // 尝试提取 A 列单元格的超链接子地址 SubAddress
                string subAddr = "";
                if (aCell != null && aCell.Hyperlinks != null && aCell.Hyperlinks.Count > 0)
                {
                    try { subAddr = aCell.Hyperlinks[1].SubAddress ?? ""; } catch { }
                }

                // ----------------------------------------------------
                // 情况 1: 修改的是顶部汇总行 (Sum 行)
                // 判定标准: A 列超链接指向 Det 明细前缀 (包含 detPrefix)
                // ----------------------------------------------------
                if (!string.IsNullOrEmpty(subAddr) && subAddr.Contains(detPrefix))
                {
                    // 提取目标 Det 标签与箱柜序号 K
                    string detTag = ExtractTag(subAddr, detPrefix);
                    int k = Tool.ExtractIndexFromName(detTag, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                    if (k <= 0 && !string.IsNullOrEmpty(subAddr))
                    {
                        // 兜底从完整超链接地址解析箱柜序号 K
                        k = Tool.ExtractIndexFromName(subAddr, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                    }

                    // 1.1 若修改的是 Sum 行 F 列 (第 6 列: 数量) -> 同步至 Tolsum 行 F 列 (第 6 列)
                    if (col == 6)
                    {
                        if (k > 0)
                        {
                            // 查找对应的 Tolsum 总计行定义名称 Range
                            string tolsumTag = $"{tolsumPrefix}{k}";
                            Microsoft.Office.Interop.Excel.Range? tolsumAnchorA = FindRangeByTag(wb, sh, tolsumTag);
                            if (tolsumAnchorA != null)
                            {
                                // 获取 Tolsum 行的 F 列单元格 (第 6 列)
                                Microsoft.Office.Interop.Excel.Range tolsumCellF = (Microsoft.Office.Interop.Excel.Range)sh.Cells[tolsumAnchorA.Row, 6];
                                // 执行安全同步赋值
                                SyncCellValue(tolsumCellF, target.Value);
                                return true;
                            }
                        }
                        return false;
                    }

                    // 1.2 查找底部明细行 (Det) 的 A 列锚点 Range
                    Microsoft.Office.Interop.Excel.Range? detAnchorA = !string.IsNullOrEmpty(detTag) ? FindRangeByTag(wb, sh, detTag) : null;
                    if (detAnchorA == null && k > 0)
                    {
                        // 兜底使用 detPrefix + k 寻找
                        detAnchorA = FindRangeByTag(wb, sh, $"{detPrefix}{k}");
                    }

                    // 若成功定位到 Det 明细行
                    if (detAnchorA != null)
                    {
                        int detRow = detAnchorA.Row;
                        Microsoft.Office.Interop.Excel.Range? targetDetCell = null;

                        // 依据修改的 Sum 行列号匹配对应的 Det 目标数据列
                        switch (col)
                        {
                            case 2:  // Sum 行 B 列 (2: 柜号) -> Det 行 B 列 (2: 柜号)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 2];
                                break;
                            case 3:  // Sum 行 C 列 (3: 箱柜名称) -> Det 行 G 列 (7: 箱柜名称, G:H合并)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 7];
                                break;
                            case 4:  // Sum 行 D 列 (4: 箱柜型号) -> Det 行 D 列 (4: 箱柜型号, D:E合并)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 4];
                                break;
                            case 9:  // Sum 行 I 列 (9: 备注) -> Det 行 I 列 (9: 备注)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 9];
                                break;
                            case 13: // Sum 行 M 列 (13: 箱柜尺寸) -> Det 行 K 列 (11: 箱柜尺寸)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 11];
                                break;
                            case 14: // Sum 行 N 列 (14: 类别) -> Det 行 M 列 (13: 类别)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 13];
                                break;
                            case 17: // Sum 行 Q 列 (17: 图号) -> Det 行 O 列 (15: 图号)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 15];
                                break;
                        }

                        // 若命中同步目标单元格，执行安全赋值并返回成功
                        if (targetDetCell != null)
                        {
                            SyncCellValue(targetDetCell, target.Value);
                            return true;
                        }
                    }
                    return false;
                }

                // ----------------------------------------------------
                // 情况 2: 修改的是底部明细行 (Det 行)
                // 判定标准: A 列超链接指向 Sum 汇总前缀 (包含 sumPrefix)
                // ----------------------------------------------------
                if (!string.IsNullOrEmpty(subAddr) && subAddr.Contains(sumPrefix))
                {
                    // 提取目标 Sum 标签与箱柜序号 K
                    string sumTag = ExtractTag(subAddr, sumPrefix);
                    int k = Tool.ExtractIndexFromName(sumTag, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                    if (k <= 0 && !string.IsNullOrEmpty(subAddr))
                    {
                        // 兜底从完整超链接解析箱柜序号 K
                        k = Tool.ExtractIndexFromName(subAddr, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                    }

                    // 查找顶部汇总行 (Sum) 的 A 列锚点 Range
                    Microsoft.Office.Interop.Excel.Range? sumAnchorA = !string.IsNullOrEmpty(sumTag) ? FindRangeByTag(wb, sh, sumTag) : null;
                    if (sumAnchorA == null && k > 0)
                    {
                        // 兜底使用 sumPrefix + k 寻找
                        sumAnchorA = FindRangeByTag(wb, sh, $"{sumPrefix}{k}");
                    }

                    // 若成功定位到 Sum 汇总行
                    if (sumAnchorA != null)
                    {
                        int sumRow = sumAnchorA.Row;
                        Microsoft.Office.Interop.Excel.Range? targetSumCell = null;

                        // 依据修改的 Det 行列号匹配对应的 Sum 目标数据列
                        switch (col)
                        {
                            case 2:  // Det 行 B 列 (2: 柜号) -> Sum 行 B 列 (2: 柜号)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 2];
                                break;
                            case 4:  // Det 行 D 列 (4: 箱柜型号) -> Sum 行 D 列 (4: 箱柜型号)
                            case 5:  // Det 行 E 列 (5: 箱柜型号合并区) -> Sum 行 D 列 (4: 箱柜型号)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 4];
                                break;
                            case 7:  // Det 行 G 列 (7: 箱柜名称) -> Sum 行 C 列 (3: 箱柜名称)
                            case 8:  // Det 行 H 列 (8: 箱柜名称合并区) -> Sum 行 C 列 (3: 箱柜名称)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 3];
                                break;
                            case 9:  // Det 行 I 列 (9: 备注) -> Sum 行 I 列 (9: 备注)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 9];
                                break;
                            case 11: // Det 行 K 列 (11: 箱柜尺寸) -> Sum 行 M 列 (13: 箱柜尺寸)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 13];
                                break;
                            case 13: // Det 行 M 列 (13: 类别) -> Sum 行 N 列 (14: 类别)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 14];
                                break;
                            case 15: // Det 行 O 列 (15: 图号) -> Sum 行 Q 列 (17: 图号)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 17];
                                break;
                        }

                        // 若命中同步目标单元格，执行安全赋值并返回成功
                        if (targetSumCell != null)
                        {
                            SyncCellValue(targetSumCell, target.Value);
                            return true;
                        }
                    }
                    return false;
                }

                // ----------------------------------------------------
                // 情况 3: 修改的是总计行 (Tolsum 行) 的 F 列 (第 6 列: 数量)
                // 判定标准: col == 6 且当前行匹配 Cab_Tolsum_k
                // ----------------------------------------------------
                if (col == 6)
                {
                    // 尝试通过定义名称识别当前行是否为 Tolsum 行并提取序号 K
                    int tolsumK = FindCabinetIndexByRowAndPrefix(wb, sh, row, tolsumPrefix);
                    if (tolsumK > 0)
                    {
                        // 找到对应的 Sum 汇总行锚点
                        string sumTag = $"{sumPrefix}{tolsumK}";
                        Microsoft.Office.Interop.Excel.Range? sumAnchorA = FindRangeByTag(wb, sh, sumTag);
                        if (sumAnchorA != null)
                        {
                            // 获取 Sum 汇总行的 F 列单元格 (第 6 列)
                            Microsoft.Office.Interop.Excel.Range sumCellF = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumAnchorA.Row, 6];
                            // 执行反向同步赋值
                            SyncCellValue(sumCellF, target.Value);
                            return true;
                        }
                    }
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 在挂起 Excel 全局事件的前提下将新值安全同步给目标单元格，并恢复事件机制
        /// </summary>
        /// <param name="targetCell">目标单元格 Range</param>
        /// <param name="newValue">待同步的新值 (支持文本、数值、空值)</param>
        private static void SyncCellValue(Microsoft.Office.Interop.Excel.Range targetCell, object? newValue)
        {
            if (_excelApp == null || targetCell == null) return;
            // 挂起 Excel 全局事件触发避免递归死循环
            _excelApp.EnableEvents = false;
            try
            {
                // 将新值同步写入目标单元格
                targetCell.Value = newValue;
            }
            catch { }
            finally
            {
                // 恢复 Excel 全局事件触发机制
                _excelApp.EnableEvents = true;
            }
        }

        /// <summary>
        /// 根据物理行号与定义名称前缀查找当前工作表中匹配的箱柜序号 K
        /// </summary>
        /// <param name="wb">工作簿对象</param>
        /// <param name="sh">工作表对象</param>
        /// <param name="row">目标物理行号</param>
        /// <param name="prefix">定义名称前缀 (如 Cab_Tolsum_)</param>
        /// <returns>匹配到的箱柜序号 K，未匹配返回 0</returns>
        private static int FindCabinetIndexByRowAndPrefix(
            Microsoft.Office.Interop.Excel.Workbook wb,
            Microsoft.Office.Interop.Excel.Worksheet sh,
            int row,
            string prefix)
        {
            try
            {
                string curSheetName = Convert.ToString(sh.Name) ?? "";

                // 1. 优先扫描工作表级别定义名称
                if (sh.Names != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Name n in sh.Names)
                    {
                        try
                        {
                            string clean = Tool.ExtractCleanNameStr(Convert.ToString(n.Name) ?? "");
                            if (clean.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                            {
                                if (n.RefersToRange != null && n.RefersToRange.Row == row)
                                {
                                    int k = Tool.ExtractIndexFromName(clean);
                                    if (k > 0) return k;
                                }
                            }
                        }
                        catch { }
                    }
                }

                // 2. 扫描工作簿级别定义名称
                if (wb != null && wb.Names != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Name n in wb.Names)
                    {
                        try
                        {
                            string clean = Tool.ExtractCleanNameStr(Convert.ToString(n.Name) ?? "");
                            if (clean.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                            {
                                var r = n.RefersToRange;
                                if (r != null && r.Row == row &&
                                    string.Equals(Convert.ToString(r.Worksheet?.Name), curSheetName, StringComparison.OrdinalIgnoreCase))
                                {
                                    int k = Tool.ExtractIndexFromName(clean);
                                    if (k > 0) return k;
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return 0;
        }

        /// <summary>
        /// 响应工作表右键点击事件：实现【WebView2 业务专属菜单】与【Excel 原生右键菜单】的彻底二选一隔离切换
        /// </summary>
        private static void OnSheetBeforeRightClick(object shObj, Microsoft.Office.Interop.Excel.Range target, ref bool cancel)
        {
            try
            {
                // 校验全局 _excelApp 与 target 句柄有效性
                if (_excelApp == null || target == null) return;

                // 读取当前配置：是否启用自定义业务右键菜单
                bool useCustomMenu = ConfigManager.Instance.Current.Excel.UseCustomContextMenu;

                // 转换当前触发右击的工作表强类型对象
                Microsoft.Office.Interop.Excel.Worksheet? activeSheet = shObj as Microsoft.Office.Interop.Excel.Worksheet;
                if (activeSheet == null) return;

                // 获取所属工作簿对象
                Microsoft.Office.Interop.Excel.Workbook? wb = null;
                try { wb = activeSheet.Parent as Microsoft.Office.Interop.Excel.Workbook; } catch { }

                // 若处于【Excel 原生右键菜单模式】：完全放行，不做任何拦截，并清理 CommandBars 注入项
                if (!useCustomMenu)
                {
                    // 彻底放行原生右键菜单
                    cancel = false;
                    // 清理任何历史上可能遗留在 CommandBars 中的自定义控件，保持 Excel 原厂菜单 100% 纯净
                    RemoveContextMenuControls();
                    // 直接返回
                    return;
                }

                // ------------------ 【以下为 WebView 2 业务专属菜单模式】 ------------------
                // 1. 完全拦截 Excel 原生右键菜单弹窗
                cancel = true;

                // 2. 清理 CommandBars，确保后台无冗余控件注入
                RemoveContextMenuControls();

                // 3. 计算箱柜特征：箱柜明细前缀
                string detPrefix = CabinetPrefixConfig.Current.DetPrefix;
                // 记录当前工作表中第一个（行号最小的）Cab_Det 物理行号
                int minDetRow = int.MaxValue;
                string currentSheetName = Convert.ToString(activeSheet.Name) ?? "";

                // 4. 扫描工作簿级别定义名称，寻找属于当前 Sheet 且匹配 detPrefix 的最小行号
                if (wb != null && wb.Names != null)
                {
                    try
                    {
                        foreach (Microsoft.Office.Interop.Excel.Name n in wb.Names)
                        {
                            try
                            {
                                string nName = Convert.ToString(n.Name) ?? "";
                                if (nName.Contains("!")) nName = nName.Substring(nName.IndexOf("!") + 1);
                                nName = nName.Trim('\'', '=', ' ', '"');
                                if (nName.StartsWith(detPrefix, StringComparison.OrdinalIgnoreCase))
                                {
                                    Microsoft.Office.Interop.Excel.Range? r = null;
                                    try { r = n.RefersToRange; } catch { }
                                    if (r != null && r.Worksheet != null && string.Equals(Convert.ToString(r.Worksheet.Name), currentSheetName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (r.Row < minDetRow) minDetRow = r.Row;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }

                // 5. 扫描工作表级别定义名称，寻找匹配 detPrefix 的最小行号
                if (activeSheet.Names != null)
                {
                    try
                    {
                        foreach (Microsoft.Office.Interop.Excel.Name n in activeSheet.Names)
                        {
                            try
                            {
                                string nName = Convert.ToString(n.Name) ?? "";
                                if (nName.Contains("!")) nName = nName.Substring(nName.IndexOf("!") + 1);
                                nName = nName.Trim('\'', '=', ' ', '"');
                                if (nName.StartsWith(detPrefix, StringComparison.OrdinalIgnoreCase))
                                {
                                    Microsoft.Office.Interop.Excel.Range? r = null;
                                    try { r = n.RefersToRange; } catch { }
                                    if (r != null)
                                    {
                                        if (r.Row < minDetRow) minDetRow = r.Row;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }

                // 获取用户当前右击的物理行号与列号
                int rightClickRow = target.Row;
                int rightClickCol = target.Column;
                // 获取当前单元格的绝对地址 (如 $C$15)
                string cellAddress = "";
                try { cellAddress = Convert.ToString(target.get_Address(Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing)) ?? ""; } catch { }

                // 判断是否在第一个 Cab_Det 行上方：
                // 若未识别到任何 Cab_Det 行（如新表），或者右击行号小于最小明细行号，则判定为在上方
                bool isAboveFirstDet = (minDetRow == int.MaxValue) || (rightClickRow < minDetRow);

                // 获取当前屏幕物理鼠标坐标
                System.Drawing.Point screenPos = System.Windows.Forms.Cursor.Position;

                // 在鼠标所在屏幕位置调起基于 WebView2 + Vue 3 的业务专属右键菜单
                ExcelAddInDemo.Forms.CustomContextMenuForm.ShowMenu(
                    screenPos,
                    currentSheetName,
                    cellAddress,
                    rightClickRow,
                    rightClickCol,
                    isAboveFirstDet
                );
            }
            catch (Exception ex)
            {
                // 记录右键菜单处理异常日志
                LogHelper.WriteLog($"右键菜单调度异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 动态更新或添加 Excel 单元格右键上下文菜单中的“识别参数并匹配物料”与“物料匹配规则设置...”按钮
        /// </summary>
        private static void UpdateParseMatchComponentsContextMenu(string menuCaption, string menuTag)
        {
            try
            {
                // 校验 Excel 全局应用实例
                if (_excelApp == null) return;

                // 将 Excel Application 转换为 dynamic 动态对象
                dynamic dynApp = (dynamic)_excelApp;
                dynamic commandBars = dynApp.CommandBars;
                if (commandBars == null) return;

                const string settingTag = "CT_BTN_OPEN_MATCH_SETTING";

                // 遍历 Excel CommandBars 中所有名称为 "Cell" 的上下文右键菜单
                foreach (dynamic bar in commandBars)
                {
                    if (bar.Name == "Cell")
                    {
                        // 1. 挂载/更新“识别参数并匹配物料”按钮
                        dynamic? existingCtrl = null;
                        try { existingCtrl = bar.FindControl(1, Type.Missing, menuTag); } catch { }

                        if (existingCtrl != null)
                        {
                            existingCtrl.Caption = menuCaption;
                            existingCtrl.Visible = true;
                            existingCtrl.Enabled = true;
                        }
                        else
                        {
                            dynamic btn = bar.Controls.Add(1, Type.Missing, Type.Missing, 2, true);
                            btn.Caption = menuCaption;
                            btn.Tag = menuTag;
                            btn.BeginGroup = true;
                            btn.OnAction = "MacroParseAndMatchComponents";
                            btn.Visible = true;
                        }

                        // 2. 挂载/更新“物料匹配与品牌规则设置...”窗口调起按钮
                        dynamic? existingSettingCtrl = null;
                        try { existingSettingCtrl = bar.FindControl(1, Type.Missing, settingTag); } catch { }

                        if (existingSettingCtrl != null)
                        {
                            existingSettingCtrl.Caption = "物料匹配规则与品牌设置...";
                            existingSettingCtrl.Visible = true;
                            existingSettingCtrl.Enabled = true;
                        }
                        else
                        {
                            dynamic btnSetting = bar.Controls.Add(1, Type.Missing, Type.Missing, 3, true);
                            btnSetting.Caption = "物料匹配规则与品牌设置...";
                            btnSetting.Tag = settingTag;
                            btnSetting.BeginGroup = false;
                            btnSetting.OnAction = "MacroOpenComponentMatchDialog";
                            btnSetting.Visible = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录更新识别参数并匹配物料右键菜单异常日志
                LogHelper.WriteLog($"更新识别参数并匹配物料右键菜单异常: {ex.Message}");
            }
        }

        /// <summary>
        /// Excel-DNA 宏入口：打开“元器件物料匹配与品牌规则设置”现代化 Vue 3 浮窗
        /// </summary>
        [ExcelCommand]
        public static void MacroOpenComponentMatchDialog()
        {
            try
            {
                // 打开或激活物料匹配设置弹窗
                ExcelServices.ShowComponentMatchDialog();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"打开物料匹配设置弹窗异常: {ex.Message}");
            }
        }

        /// <summary>
        /// Excel-DNA 宏入口：响应右键菜单中“识别参数并匹配物料”按钮指令
        /// </summary>
        [ExcelCommand]
        public static void MacroParseAndMatchComponents()
        {
            try
            {
                // 调度执行业务服务层：根据选区已有的 S/T/U 列参数反查物料库并回填
                var result = ExcelServices.ExecuteBatchMatchWithDb(null);

                // 弹出执行结果反馈提示
                if (result != null)
                {
                    if (result.Success)
                    {
                        // 提示处理完成
                        System.Windows.Forms.MessageBox.Show(
                            result.Message,
                            "识别与匹配物料完成",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information
                        );
                    }
                    else
                    {
                        // 提示失败信息
                        System.Windows.Forms.MessageBox.Show(
                            $"识别与匹配失败: {result.Message}",
                            "提示",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"右键识别参数并匹配物料执行异常: {ex.Message}");
                // 弹出异常弹窗
                System.Windows.Forms.MessageBox.Show($"执行异常: {ex.Message}", "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 动态更新或添加 Excel 单元格右键上下文菜单（CommandBars["Cell"]）中的“新建箱柜”按钮
        /// </summary>
        private static void UpdateNewCabinetContextMenu(bool isAboveFirstDet, string menuCaption, string menuTag)
        {
            try
            {
                // 校验 Excel 全局应用实例
                if (_excelApp == null) return;

                // 将 Excel Application 转换为 dynamic 动态对象，避开对 office.dll PIA 的静态类型编译依赖
                dynamic dynApp = (dynamic)_excelApp;
                dynamic commandBars = dynApp.CommandBars;
                if (commandBars == null) return;

                // 遍历 Excel CommandBars 中所有名称为 "Cell" 的上下文右键菜单 (兼容普通视图与分页预览视图)
                foreach (dynamic bar in commandBars)
                {
                    if (bar.Name == "Cell")
                    {
                        // 尝试根据 Tag 查找是否已经存在该右键按钮控件 (Type 1 为 msoControlButton)
                        dynamic? existingCtrl = null;
                        try
                        {
                            existingCtrl = bar.FindControl(1, Type.Missing, menuTag);
                        }
                        catch { }

                        // 若处于第一个 Cab_Det 上方，添加或显示按钮
                        if (isAboveFirstDet)
                        {
                            if (existingCtrl != null)
                            {
                                // 控件已存在，确保标题正确并设置为可见与启用
                                existingCtrl.Caption = menuCaption;
                                existingCtrl.Visible = true;
                                existingCtrl.Enabled = true;
                            }
                            else
                            {
                                // 控件不存在，在右键菜单顶部位置添加临时按钮 (Type 1 为 msoControlButton)
                                dynamic btn = bar.Controls.Add(
                                    1, // msoControlButton
                                    Type.Missing,
                                    Type.Missing,
                                    1, // 放置在右键菜单最顶部首项
                                    true // Temporary: true (关闭 Excel 自动清理)
                                );
                                // 设置按钮显示文本
                                btn.Caption = menuCaption;
                                // 设置按钮唯一 Tag 标识
                                btn.Tag = menuTag;
                                // 开启分组分割线
                                btn.BeginGroup = true;
                                // 设置点击时调用的 Excel-DNA 宏名称，确保点击响应 100% 可靠稳定
                                btn.OnAction = "MacroCreateNewCabinet";
                                // 显示按钮
                                btn.Visible = true;
                            }
                        }
                        else
                        {
                            // 若不在第一个 Cab_Det 上方，将该按钮隐藏
                            if (existingCtrl != null)
                            {
                                existingCtrl.Visible = false;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录更新菜单控件异常日志
                LogHelper.WriteLog($"更新右键菜单按钮异常: {ex.Message}");
            }
        }

        /// <summary>
        /// Excel-DNA 宏入口：响应右键菜单中“新建箱柜”按钮指令
        /// </summary>
        [ExcelCommand]
        public static void MacroCreateNewCabinet()
        {
            try
            {
                // 调度至主线程执行“新建箱柜”业务逻辑
                ExcelServices.CreateNewCabinetFromSelection();
            }
            catch (Exception ex)
            {
                // 记录右键新建箱柜执行异常日志
                LogHelper.WriteLog($"右键新建箱柜执行异常: {ex.Message}");
            }
        }

        /// <summary>
        /// Excel-DNA 宏入口：一键切换【业务专属右键菜单】与【Excel 原生右键菜单】模式
        /// </summary>
        [ExcelCommand]
        public static void MacroToggleContextMenuMode()
        {
            try
            {
                // 执行切换并获取最新的模式布尔值
                bool newMode = ConfigManager.Instance.ToggleCustomContextMenuMode();

                // 彻底清理 CommandBars 遗留，保证原生环境无污染
                RemoveContextMenuControls();

                // 准备提示文本
                string modeText = newMode ? "【业务专属右键菜单 (WebView2)】" : "【Excel 纯净原生右键菜单】";
                string detailText = newMode
                    ? "在工作表中右键将弹出极简现代化的业务菜单！"
                    : "在工作表中右键将直接弹出 Excel 原厂自带的默认菜单！";

                // 弹出切换成功提示
                System.Windows.Forms.MessageBox.Show(
                    $"已成功切换至：{modeText}\n\n{detailText}",
                    "右键菜单模式切换",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                // 记录切换异常日志
                LogHelper.WriteLog($"切换右键菜单模式异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 清理插件注册的所有右键菜单控件，保持 Excel 原始环境干净整洁
        /// </summary>
        public static void RemoveContextMenuControls()
        {
            try
            {
                // 校验 Excel 全局应用实例
                if (_excelApp == null) return;
                // 读取配置中定义的唯一 Tag 标识
                string menuTag = ConfigManager.Instance.Current.Excel.NewCabinetMenuTag ?? "CT_BTN_NEW_CABINET";
                string matchTag = ConfigManager.Instance.Current.Excel.ParseMatchComponentsMenuTag ?? "CT_BTN_PARSE_MATCH_COMPONENTS";
                const string settingTag = "CT_BTN_OPEN_MATCH_SETTING";

                // 将 Excel Application 转换为 dynamic 动态对象
                dynamic dynApp = (dynamic)_excelApp;
                dynamic commandBars = dynApp.CommandBars;
                if (commandBars == null) return;

                // 遍历所有 Cell 上下文菜单并安全删除自定义控件
                foreach (dynamic bar in commandBars)
                {
                    if (bar.Name == "Cell")
                    {
                        try
                        {
                            dynamic? ctrl = bar.FindControl(1, Type.Missing, menuTag);
                            if (ctrl != null) ctrl.Delete(true);

                            dynamic? matchCtrl = bar.FindControl(1, Type.Missing, matchTag);
                            if (matchCtrl != null) matchCtrl.Delete(true);

                            dynamic? settingCtrl = bar.FindControl(1, Type.Missing, settingTag);
                            if (settingCtrl != null) settingCtrl.Delete(true);
                        }
                        catch { }
                    }
                }
                // 清空引用集合
                _contextMenuButtons.Clear();
            }
            catch { }
        }


        /// <summary>
        /// 从超链接子地址中截取精准的定义名称标签字符串
        /// </summary>
        private static string ExtractTag(string subAddr, string prefix)
        {
            try
            {
                // 定位前缀关键字位置
                int idx = subAddr.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                {
                    // 截取前缀及其之后的所有字符串
                    string tag = subAddr.Substring(idx);
                    // 清理单引号分隔符
                    int endIdx = tag.IndexOf('\'');
                    if (endIdx > 0) tag = tag.Substring(0, endIdx);
                    // 返回修剪后的纯净标签名
                    return tag.Trim();
                }
            }
            catch { }
            return "";
        }

        /// <summary>
        /// 依据定义名称标签精准获取 Range 对象，带多层 Safe-Lookup 兜底 (优先工作表级，后回退工作簿级)
        /// </summary>
        private static Microsoft.Office.Interop.Excel.Range? FindRangeByTag(Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.Worksheet sh, string tagName)
        {
            try
            {
                string curSheetName = Convert.ToString(sh.Name) ?? "";

                // 1. 优先遍历当前工作表中的定义名称
                if (sh.Names != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Name n in sh.Names)
                    {
                        try
                        {
                            string nStr = Tool.ExtractCleanNameStr(Convert.ToString(n.Name) ?? "");
                            if (string.Equals(nStr, tagName, StringComparison.OrdinalIgnoreCase) || nStr.EndsWith(tagName, StringComparison.OrdinalIgnoreCase))
                            {
                                if (n.RefersToRange != null) return n.RefersToRange;
                            }
                        }
                        catch { }
                    }
                }

                // 2. 安全遍历工作簿中的定义名称进行精确比对与工作表校验
                if (wb != null && wb.Names != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Name n in wb.Names)
                    {
                        try
                        {
                            string nStr = Tool.ExtractCleanNameStr(Convert.ToString(n.Name) ?? "");
                            if (string.Equals(nStr, tagName, StringComparison.OrdinalIgnoreCase) || nStr.EndsWith(tagName, StringComparison.OrdinalIgnoreCase))
                            {
                                var r = n.RefersToRange;
                                if (r != null && string.Equals(Convert.ToString(r.Worksheet?.Name), curSheetName, StringComparison.OrdinalIgnoreCase))
                                {
                                    return r;
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return null;
        }
    }
}
