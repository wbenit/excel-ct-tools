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
                // 重新绑定 SheetSelectionChange 事件处理委托，实现进入 C 列元器件行自动触发覆盖式智能输入 (方案 B)
                _excelApp.SheetSelectionChange += OnSheetSelectionChange;
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
                }

                // 隐藏方案 B 覆盖输入框
                ExcelServices.HideSmartInputOverlay();

                // 彻底清理注册的右键菜单控件
                RemoveContextMenuControls();
            }
            catch { }
        }

        /// <summary>
        /// 响应工作表单元格焦点切换事件，实现选中 C 列元器件行时自动激活覆盖式智能输入 (方案 B / 对应 ZhiNengEn.ShuRu)
        /// </summary>
        private static void OnSheetSelectionChange(object shObj, Microsoft.Office.Interop.Excel.Range target)
        {
            try
            {
                // 校验目标单元格与全局 Application
                if (target == null || _excelApp == null) return;

                // 若选中的是单个单元格
                if (target.Rows.Count == 1 && target.Columns.Count == 1)
                {
                    // 尝试激活覆盖式智能输入框 (内部自动校验 C 列与箱柜元器件行区间)
                    ExcelServices.ShuRu(target);
                }
                else
                {
                    // 选中多单元格区域时隐藏覆盖输入框
                    ExcelServices.HideSmartInputOverlay();
                }
            }
            catch
            {
                // 异常时安全兜底隐藏覆盖输入框
                ExcelServices.HideSmartInputOverlay();
            }
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
        /// 1. Sum B (2) <-> Det B (2) [箱柜名称]
        /// 2. Sum C (3) <-> Det G (7) [柜型]
        /// 3. Sum D (4) <-> Det D (4) [外形尺寸]
        /// 4. Sum F (6) <-> Tolsum F (6) [数量/台数]
        /// 5. Sum I (9) <-> Det I (9) [备注]
        /// 6. Sum M (13) <-> Det K (11) [重量]
        /// 7. Sum N (14) <-> Det M (13) [排产周期]
        /// 8. Sum O (15) <-> Det O (15) [开标日期]
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

                        // 依据修改的 Sum 行列号匹配对应的 Det 目标列
                        switch (col)
                        {
                            case 2:  // Sum 行 B 列 (2: 箱柜名称) -> Det 行 B 列 (2)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 2];
                                break;
                            case 3:  // Sum 行 C 列 (3: 柜型) -> Det 行 G 列 (7)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 7];
                                break;
                            case 4:  // Sum 行 D 列 (4: 外形尺寸) -> Det 行 D 列 (4)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 4];
                                break;
                            case 9:  // Sum 行 I 列 (9: 备注) -> Det 行 I 列 (9)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 9];
                                break;
                            case 13: // Sum 行 M 列 (13: 重量) -> Det 行 K 列 (11)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 11];
                                break;
                            case 14: // Sum 行 N 列 (14: 排产周期) -> Det 行 M 列 (13)
                                targetDetCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[detRow, 13];
                                break;
                            case 15: // Sum 行 O 列 (15: 开标日期) -> Det 行 O 列 (15)
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

                        // 依据修改的 Det 行列号匹配对应的 Sum 目标列
                        switch (col)
                        {
                            case 2:  // Det 行 B 列 (2: 箱柜名称) -> Sum 行 B 列 (2)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 2];
                                break;
                            case 4:  // Det 行 D 列 (4: 外形尺寸) -> Sum 行 D 列 (4)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 4];
                                break;
                            case 7:  // Det 行 G 列 (7: 柜型) -> Sum 行 C 列 (3)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 3];
                                break;
                            case 9:  // Det 行 I 列 (9: 备注) -> Sum 行 I 列 (9)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 9];
                                break;
                            case 11: // Det 行 K 列 (11: 重量) -> Sum 行 M 列 (13)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 13];
                                break;
                            case 13: // Det 行 M 列 (13: 排产周期) -> Sum 行 N 列 (14)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 14];
                                break;
                            case 15: // Det 行 O 列 (15: 开标日期) -> Sum 行 O 列 (15)
                                targetSumCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[sumRow, 15];
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
        /// 响应工作表右键点击事件：在第一个 Cab_Det 行上方右击时，动态添加/显示“新建箱柜”右键菜单按钮
        /// </summary>
        private static void OnSheetBeforeRightClick(object shObj, Microsoft.Office.Interop.Excel.Range target, ref bool cancel)
        {
            try
            {
                // 校验全局 _excelApp 与 target 句柄有效性
                if (_excelApp == null || target == null) return;

                // 转换当前触发右击的工作表强类型对象
                Microsoft.Office.Interop.Excel.Worksheet? activeSheet = shObj as Microsoft.Office.Interop.Excel.Worksheet;
                if (activeSheet == null) return;

                // 获取所属工作簿对象
                Microsoft.Office.Interop.Excel.Workbook? wb = null;
                try { wb = activeSheet.Parent as Microsoft.Office.Interop.Excel.Workbook; } catch { }

                // 【配置文件替代硬编码列举】
                // 1. 箱柜明细前缀: CabinetPrefixConfig.Current.DetPrefix
                // 2. 右键菜单文本: NewCabinetMenuCaption (默认 新建箱柜)
                // 3. 右键菜单Tag标识: NewCabinetMenuTag (默认 CT_BTN_NEW_CABINET)
                string detPrefix = CabinetPrefixConfig.Current.DetPrefix;
                string menuCaption = ConfigManager.Instance.Current.Excel.NewCabinetMenuCaption ?? "新建箱柜";
                string menuTag = ConfigManager.Instance.Current.Excel.NewCabinetMenuTag ?? "CT_BTN_NEW_CABINET";

                // 记录当前工作表中第一个（行号最小的）Cab_Det 物理行号
                int minDetRow = int.MaxValue;
                string currentSheetName = Convert.ToString(activeSheet.Name) ?? "";

                // 1. 扫描工作簿级别定义名称，寻找属于当前 Sheet 且匹配 detPrefix 的最小行号
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

                // 2. 扫描工作表级别定义名称，寻找匹配 detPrefix 的最小行号
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

                // 获取用户当前右击的物理行号
                int rightClickRow = target.Row;

                // 判断是否在第一个 Cab_Det 行上方：
                // 若未识别到任何 Cab_Det 行（如新表），或者右击行号小于最小明细行号，则判定为在上方
                bool isAboveFirstDet = (minDetRow == int.MaxValue) || (rightClickRow < minDetRow);

                // 动态更新 Excel 单元格右键菜单中“新建箱柜”按钮的状态
                UpdateNewCabinetContextMenu(isAboveFirstDet, menuCaption, menuTag);
            }
            catch (Exception ex)
            {
                // 记录右键菜单更新异常日志
                LogHelper.WriteLog($"右键菜单状态更新异常: {ex.Message}");
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
                            if (ctrl != null)
                            {
                                ctrl.Delete(true);
                            }
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
