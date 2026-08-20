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
        /// 响应单元格修改事件，实现箱柜名称在顶部与底部的纯文本双向同步更新
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

                // 处理第 3 列 (C列 - 规格型号) 修改时的智能属性联动回填
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

                // 限制仅处理第 2 列 (B列) 的修改
                if (target.Column != 2) return;

                // 读取改动单元格的新纯文本内容
                string newName = Convert.ToString(target.Value) ?? "";
                // 若新文本为空白则直接退出
                if (string.IsNullOrWhiteSpace(newName)) return;

                // 获取改动单元格所在行第 1 列 (A列) 单元格
                Microsoft.Office.Interop.Excel.Range aCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[target.Row, 1];

                // 校验 A列单元格是否包含超链接锚点
                if (aCell != null && aCell.Hyperlinks != null && aCell.Hyperlinks.Count > 0)
                {
                    // 获取超链接跳转目标子地址 (SubAddress)
                    string subAddr = "";
                    try { subAddr = aCell.Hyperlinks[1].SubAddress ?? ""; } catch { }

                    // 【配置文件替代硬编码列举】
                    // 1. 原硬编码前缀校验: "Cab_Sum_" (汇总标签检测) 和 "Cab_Det_" (明细标签检测)
                    // 2. 替代配置项: CabinetPrefixConfig.Current
                    var prefixes = CabinetPrefixConfig.Current;
                    string sumPrefix = prefixes.SumPrefix;
                    string detPrefix = prefixes.DetPrefix;

                    // 若超链接指向明细前缀 (表明当前修改的是顶部汇总行的 B列名称)
                    if (subAddr.Contains(detPrefix))
                    {
                        // 提取对应的目标明细定义名称标签 (例如 Cab_Det_1)
                        string targetTag = ExtractTag(subAddr, detPrefix);
                        if (!string.IsNullOrEmpty(targetTag))
                        {
                            // 查找对应的底部明细 A列单元格 Range
                            Microsoft.Office.Interop.Excel.Range? detAnchorA = FindRangeByTag(wb, sh, targetTag);
                            if (detAnchorA != null)
                            {
                                // 获取底部明细 B列名称单元格 (右移1列)
                                Microsoft.Office.Interop.Excel.Range detNameB = (Microsoft.Office.Interop.Excel.Range)detAnchorA.Offset[0, 1];

                                // 暂停事件触发防止无限递归
                                _excelApp.EnableEvents = false;
                                try
                                {
                                    // 1. 同步更新底部明细表头 B列纯文本数值
                                    detNameB.Value = newName;
                                }
                                catch { }
                                finally
                                {
                                    // 恢复事件触发机制
                                    _excelApp.EnableEvents = true;
                                }
                                return;
                            }
                        }
                    }
                    // 若超链接指向汇总前缀 (表明当前修改的是底部明细行的 B列名称)
                    else if (subAddr.Contains(sumPrefix))
                    {
                        // 提取对应的目标汇总定义名称标签 (例如 Cab_Sum_1)
                        string targetTag = ExtractTag(subAddr, sumPrefix);
                        if (!string.IsNullOrEmpty(targetTag))
                        {
                            // 查找对应的顶部汇总 A列单元格 Range
                            Microsoft.Office.Interop.Excel.Range? sumAnchorA = FindRangeByTag(wb, sh, targetTag);
                            if (sumAnchorA != null)
                            {
                                // 获取顶部汇总 B列名称单元格 (右移1列)
                                Microsoft.Office.Interop.Excel.Range sumNameB = (Microsoft.Office.Interop.Excel.Range)sumAnchorA.Offset[0, 1];

                                // 暂停事件触发防止无限递归
                                _excelApp.EnableEvents = false;
                                try
                                {
                                    // 1. 反向同步更新顶部汇总行 B列纯文本数值
                                    sumNameB.Value = newName;
                                }
                                catch { }
                                finally
                                {
                                    // 恢复事件触发机制
                                    _excelApp.EnableEvents = true;
                                }
                                return;
                            }
                        }
                    }
                }
            }
            catch { }
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
        /// 依据定义名称标签精准获取 Range 对象，带多层 Safe-Lookup 兜底
        /// </summary>
        private static Microsoft.Office.Interop.Excel.Range? FindRangeByTag(Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.Worksheet sh, string tagName)
        {
            try
            {
                // 安全遍历工作簿中的定义名称进行精确名称比对与后缀匹配
                foreach (Microsoft.Office.Interop.Excel.Name n in wb.Names)
                {
                    string nStr = Convert.ToString(n.Name) ?? "";
                    if (nStr.Contains("!"))
                    {
                        nStr = nStr.Substring(nStr.IndexOf("!") + 1);
                    }
                    nStr = nStr.Trim('\'', '=', ' ', '"');
                    if (string.Equals(nStr, tagName, StringComparison.OrdinalIgnoreCase) || nStr.EndsWith(tagName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (n.RefersToRange != null) return n.RefersToRange;
                    }
                }
            }
            catch { }
            return null;
        }
    }
}
