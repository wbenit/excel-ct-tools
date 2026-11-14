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

        // 记录上一次激活的工作簿唯一标识，避免相同工作簿重复激活产生冗余开销
        private static string _lastActivatedWorkbookKey = string.Empty;

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

                // 解除已有的 SheetActivate 事件绑定，避免重复挂载
                _excelApp.SheetActivate -= OnSheetActivate;
                // 重新绑定 SheetActivate 事件，新建表格或切换工作表时即时唤醒聚光灯
                _excelApp.SheetActivate += OnSheetActivate;

                // 解除已有的 WorkbookActivate 事件绑定
                _excelApp.WorkbookActivate -= OnWorkbookActivate;
                // 重新绑定 WorkbookActivate 事件，激活或新建工作簿时自动挂载视口
                _excelApp.WorkbookActivate += OnWorkbookActivate;

                // 解除已有的 WorkbookDeactivate 事件绑定，避免重复挂载
                _excelApp.WorkbookDeactivate -= OnWorkbookDeactivate;
                // 重新绑定 WorkbookDeactivate 事件，监控工作簿失去焦点时的定义名称现场
                _excelApp.WorkbookDeactivate += OnWorkbookDeactivate;

                // 解除已有的 WorkbookBeforeClose 事件绑定，避免重复挂载
                _excelApp.WorkbookBeforeClose -= OnWorkbookBeforeClose;
                // 重新绑定 WorkbookBeforeClose 事件，在工作簿关闭前执行聚光灯安全避让
                _excelApp.WorkbookBeforeClose += OnWorkbookBeforeClose;

                // 解除已有的 WorkbookAfterSave 事件绑定
                _excelApp.WorkbookAfterSave -= OnWorkbookAfterSave;
                // 重新绑定 WorkbookAfterSave 事件，文件保存后即时刷新 Ribbon 本机项目收录状态
                _excelApp.WorkbookAfterSave += OnWorkbookAfterSave;

                // 贯彻核心指令：这里的撤回不使用快捷键，彻底归还 Ctrl+Z 与 Ctrl+Y 原生控制权
                try
                {
                    // 显式解绑 Ctrl+Z，确保 100% 保持 Excel 原生撤销行为
                    _excelApp.OnKey("^z");
                    // 显式解绑 Ctrl+Y，确保 100% 保持 Excel 原生还原行为
                    _excelApp.OnKey("^y");

                    // 挂接成套元器件专属全局快捷键 (Ctrl+Shift+C 复制, Ctrl+Shift+V 插入, Ctrl+Shift+D 删除, Ctrl+Shift+X 剪切)
                    _excelApp.OnKey("^+c", "MacroCopyComponentAction");
                    _excelApp.OnKey("^+v", "MacroInsertCopiedComponentAction");
                    _excelApp.OnKey("^+d", "MacroDeleteComponentAction");
                    _excelApp.OnKey("^+x", "MacroCutComponentAction");
                }
                catch { }

                // 启动后台空闲静默预热：提前加载 WebView2 内核与物料下拉界面，彻底消除用户首次点击冷启动延迟
                try
                {
                    ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        ExcelServices.PreloadComponentMatchOverlay();
                    });
                }
                catch { }
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
                    // 解除 SheetActivate 事件绑定
                    _excelApp.SheetActivate -= OnSheetActivate;
                    // 解除 WorkbookActivate 事件绑定
                    _excelApp.WorkbookActivate -= OnWorkbookActivate;
                    // 解除 WorkbookDeactivate 事件绑定
                    _excelApp.WorkbookDeactivate -= OnWorkbookDeactivate;
                    // 解除 WorkbookBeforeClose 事件绑定
                    _excelApp.WorkbookBeforeClose -= OnWorkbookBeforeClose;
                    // 解除 WorkbookAfterSave 事件绑定
                    _excelApp.WorkbookAfterSave -= OnWorkbookAfterSave;

                    // 释放快捷键挂钩，恢复 Excel 原生按键映射
                    try
                    {
                        _excelApp.OnKey("^z");
                        _excelApp.OnKey("^y");

                        // 释放成套元器件专属快捷键
                        _excelApp.OnKey("^+c");
                        _excelApp.OnKey("^+v");
                        _excelApp.OnKey("^+d");
                        _excelApp.OnKey("^+x");
                    }
                    catch { }
                }

                // 隐藏覆盖输入框与物料联想下拉浮窗
                ExcelServices.HideSmartInputOverlay();
                ExcelServices.HideComponentMatchOverlay();

                // 释放聚光灯浮窗与视口消息钩子
                ExcelServices.DisableSpotlight();

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

                // 核心加固：若工作表已处于筛选激活状态 (FilterMode == true) 且仍残留 CutCopyMode 跑马灯虚线框
                // 说明用户刚刚执行了筛选，自动自愈清除挂起的复制剪切状态，防止输入卡死
                try
                {
                    // 检测是否处于复制剪切状态
                    if (_excelApp.CutCopyMode != 0)
                    {
                        // 提取所属工作表强类型对象
                        var activeWs = (shObj as Microsoft.Office.Interop.Excel.Worksheet) ?? target.Worksheet;
                        // 若工作表处于筛选激活模式
                        if (activeWs != null && activeWs.FilterMode)
                        {
                            // 自动清除挂起的复制状态，彻底消灭残留虚线框
                            _excelApp.CutCopyMode = (Microsoft.Office.Interop.Excel.XlCutCopyMode)0;
                        }
                    }
                }
                catch { }

                // 0. 若开启了聚光灯高亮功能，即时计算并更新当前选中单元格的行列十字高亮区域
                if (ExcelServices.IsSpotlightEnabled)
                {
                    // 调用聚光灯公共服务更新浮窗高亮区域
                    ExcelServices.UpdateSpotlightPosition(target);
                }

                // 提取所属工作表与工作簿 COM 对象
                var ws = (shObj as Microsoft.Office.Interop.Excel.Worksheet) ?? target.Worksheet;
                var wb = ws?.Parent as Microsoft.Office.Interop.Excel.Workbook;

                // 核心安全守门一：若当前工作簿不是成套工程工作簿 (不含【项目信息】表)，0ms 极速放行并退出
                // 彻底杜绝在外部公司的普通报价单中盲目读取 AD/AA 列与调用箱柜检索引起的卡顿
                if (!Tool.IsProjectWorkbook(wb))
                {
                    // 安全隐藏浮窗，避免遮挡外部常规表格视线
                    ExcelServices.HideSmartInputOverlay();
                    // 隐藏物料联想覆盖下拉框
                    ExcelServices.HideComponentMatchOverlay();
                    return;
                }


                // 1. 若开启了 CAD 联动，只要选区触及常规核心列且具备联动价值，即时向 CAD 发起联动
                int startCol = target.Column;
                int endCol = startCol + target.Columns.Count - 1;
                // 获取当前活动工作表名称
                string sheetTitle = Convert.ToString(ws?.Name)?.Trim() ?? string.Empty;
                // 判断当前活动工作表是否为元件汇总表 (重命名为 isSummaryWorksheet 避免与后续局部变量冲突)
                bool isSummaryWorksheet = string.Equals(sheetTitle, Models.ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase);
                // 选区列判断：汇总表中点击 A~Z 任意数据列均触发，普通分类表中覆盖 A~G 列或包含 C 列或整行
                bool isComponentRowArea = isSummaryWorksheet
                    ? (startCol <= 26 && endCol >= 1)
                    : ((startCol <= 7 && endCol >= 1) || (3 >= startCol && 3 <= endCol));

                // 若处于有效列范围且全局开启了 CAD 联动：提取选区内所有行的 AD 列 (第 30 列) / AA 列 (第 27 列兼容) 句柄并防抖推送
                if (isComponentRowArea && Services.CadSyncClient.SyncToCadEnabled)
                {
                    try
                    {
                        // 校验工作表有效性
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
                                // 读取 AD 列 (第 30 列) 的 CAD 句柄字符串，若为空容错读取 AA 列 (第 27 列)
                                string rawHandles = Convert.ToString(ws.Range[$"AD{r}"].Value2)?.Trim() ?? string.Empty;
                                if (string.IsNullOrEmpty(rawHandles))
                                {
                                    rawHandles = Convert.ToString(ws.Range[$"AA{r}"].Value2)?.Trim() ?? string.Empty;
                                }
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

                            // 智能解析提取当前选区行所属的箱柜名称 (柜号)
                            string activeCabinetName = ExcelServices.GetActiveCabinetName(ws, startRow);
                            // 智能解析提取当前选区行所属箱柜的源 DWG 图纸路径 (从 Cab_Det 行 AE 列读取)
                            string activeSourceDwg = ExcelServices.GetActiveCabinetSourceDwg(ws, startRow);

                            // 提取当前选中行的规格型号 (支持普通分类表 C 列与元件汇总表 D/C/R 列)
                            string selectedModel = string.Empty;
                            // 校验是否为元件汇总表
                            if (isSummaryWorksheet)
                            {
                                // 元件汇总表：优先提取 D 列型号规格，若为空回退提取 C 列原型号，再回退 R 列原始型号
                                selectedModel = Convert.ToString(ws.Range[$"D{startRow}"].Value2)?.Trim() ?? string.Empty;
                                if (string.IsNullOrEmpty(selectedModel))
                                {
                                    // 兜底读取 C 列原型号规格
                                    selectedModel = Convert.ToString(ws.Range[$"C{startRow}"].Value2)?.Trim() ?? string.Empty;
                                }
                                if (string.IsNullOrEmpty(selectedModel))
                                {
                                    // 再次兜底读取 R 列原始型号 (可能包含多型号)
                                    selectedModel = Convert.ToString(ws.Range[$"R{startRow}"].Value2)?.Trim() ?? string.Empty;
                                }
                            }
                            else
                            {
                                // 分类明细表：提取 C 列型号规格
                                selectedModel = Convert.ToString(ws.Range[$"C{startRow}"].Value2)?.Trim() ?? string.Empty;
                            }

                            // 带有 50ms 防抖与自动缩放视角推送至 CAD 管道 (附带箱柜名称、型号及源图纸路径以支持跨图自动切换定位)
                            Services.CadSyncClient.SendHandlesDebounced(handleList, true, cabinetName: activeCabinetName, model: selectedModel, sourceDwgPath: activeSourceDwg);
                        }
                    }
                    catch { }
                }

                // 2. 处理 UI 浮窗交互：若选中的是单个单元格
                if (target.Rows.Count == 1 && target.Columns.Count == 1)
                {
                    int col = target.Column;

                    // 2.1 若选中的是 C 列 (第 3 列: 规格型号 / 汇总表原型号)
                    if (col == 3)
                    {
                        // 获取当前工作表名称
                        string curSheetName = (shObj as Microsoft.Office.Interop.Excel.Worksheet)?.Name ?? target.Worksheet?.Name ?? string.Empty;
                        bool isSummarySheet = string.Equals(curSheetName.Trim(), ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase);

                        if (!isSummarySheet)
                        {
                            // 校验是否落在箱柜元器件有效行或计费区域有效行区间 (涵盖 Cab_Det+2 至 Cab_Tolsum)
                            dynamic curSheet = target.Worksheet;
                            int targetRow = target.Row;
                            if (ExcelServices.IsCategoryComponentOrFeeRow(curSheet, targetRow))
                            {
                                // 优先读取元器件物料匹配与品牌规则配置 (包含“搜索”开关状态)
                                var matchCfg = ExcelServices.LoadComponentMatchFilterConfig();

                                // 依据业务优先级执行智能分流：
                                // 若用户在【元器件物料匹配与品牌规则设置】中明确勾选了“搜索”功能 (默认开启)
                                // 则绝对优先弹出全新的云端/本地物料模糊联想下拉悬浮框 (贴合 C 列下方)
                                if (matchCfg != null && matchCfg.EnableSearchOverlay)
                                {
                                    // 平滑隐藏原生覆盖输入框
                                    ExcelServices.HideSmartInputOverlay();
                                    // 弹出物料智能联想下拉悬浮框 (贴合 C 列下方，标记已校验跳过二次计算)
                                    ExcelServices.ShowComponentMatchOverlay(target, isCategoryRowValidated: true);
                                }
                                else
                                {
                                    // 用户未开启物料匹配搜索时，再回退读取智能输入配置 (内存直出 0ms)
                                    var smartCtrl = new Controllers.SmartInputController();
                                    var smartCfg = smartCtrl.GetConfig();

                                    // 仅当开启了覆盖式智能输入且未开启物料匹配时，才唤起原生覆盖输入窗体
                                    if (smartCfg != null && smartCfg.AutoPopupFloatWindow)
                                    {
                                        // 隐藏物料匹配下拉悬浮框
                                        ExcelServices.HideComponentMatchOverlay();
                                        // 激活原生覆盖输入框 (传入已校验标志跳过二次计算)
                                        ExcelServices.ShuRu(target, isCategoryRowValidated: true);
                                    }
                                    else
                                    {
                                        // 两者均未开启时静默隐藏全部浮窗
                                        ExcelServices.HideSmartInputOverlay();
                                        ExcelServices.HideComponentMatchOverlay();
                                    }
                                }
                            }
                            else
                            {
                                // 落在汇总区、信息行、计费区或小计总计行时静默隐藏全部浮窗
                                ExcelServices.HideSmartInputOverlay();
                                ExcelServices.HideComponentMatchOverlay();
                            }
                        }
                        else
                        {
                            // 在“元件汇总表”中: C 列专用于基准查看与 CAD 夹点联动，隐藏全部浮窗
                            ExcelServices.HideSmartInputOverlay();
                            ExcelServices.HideComponentMatchOverlay();
                        }
                    }
                    // 2.2 若选中的是 D 列 (第 4 列: 规格型号 / 分类表明细厂家品牌)
                    else if (col == 4)
                    {
                        // 隐藏智能输入覆盖框
                        ExcelServices.HideSmartInputOverlay();

                        // 获取当前工作表名称
                        string curSheetName = (shObj as Microsoft.Office.Interop.Excel.Worksheet)?.Name ?? target.Worksheet?.Name ?? string.Empty;
                        bool isSummarySheet = string.Equals(curSheetName.Trim(), ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase);

                        // 仅在“元件汇总表”中 D 列触发物料联想下拉
                        if (isSummarySheet)
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
                            // 分类表中 D 列为厂家品牌列，隐藏物料下拉浮窗
                            ExcelServices.HideComponentMatchOverlay();
                        }
                    }
                    else
                    {
                        // 离开 C/D 列时全部隐藏浮窗
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
        /// 响应工作表激活事件，当新建表格或切换工作表时自动同步聚光灯
        /// </summary>
        /// <param name="sh">激活的工作表对象</param>
        private static void OnSheetActivate(object sh)
        {
            try
            {
                // 若聚光灯功能处于激活状态，刷新聚光灯高亮位置
                if (ExcelServices.IsSpotlightEnabled)
                {
                    // 调用聚光灯公共服务更新浮窗高亮区域
                    ExcelServices.UpdateSpotlightPosition(null);
                }
            }
            catch { }
        }

        /// <summary>
        /// 响应工作簿激活事件，当新建或打开工作簿时自愈点亮聚光灯
        /// </summary>
        /// <param name="wb">激活的工作簿对象</param>
        private static void OnWorkbookActivate(Microsoft.Office.Interop.Excel.Workbook wb)
        {
            try
            {
                // 提取当前激活工作簿的唯一路径或名称标识
                string currentKey = string.Empty;
                try { currentKey = wb?.FullName ?? wb?.Name ?? string.Empty; } catch { }

                // 检查是否为同一个工作簿重复触发激活事件
                bool isSameWorkbook = !string.IsNullOrEmpty(currentKey) &&
                    string.Equals(currentKey, _lastActivatedWorkbookKey, StringComparison.OrdinalIgnoreCase);

                // 记录当前最新激活的工作簿标识
                _lastActivatedWorkbookKey = currentKey;

                // 若聚光灯功能处于激活状态，安全异步恢复聚光灯服务（避开工作簿切换临界期的 COM 重入阻塞）
                if (ExcelServices.IsSpotlightEnabled)
                {
                    // 延迟异步调度聚光灯恢复，彻底避开 SDI 窗口注销与重绘冲突
                    System.Windows.Forms.Timer activateTimer = new System.Windows.Forms.Timer { Interval = 150 };
                    // 绑定单次触发回调
                    activateTimer.Tick += (s, e) =>
                    {
                        // 停止并清理临时单次定时器
                        activateTimer.Stop();
                        activateTimer.Dispose();
                        // 异步安全恢复聚光灯高亮与守护定时器
                        ExcelServices.ResumeSpotlightAfterTransition();
                    };
                    // 启动临时延时定时器
                    activateTimer.Start();
                }

                // 仅当切换至不同工作簿时，才触发功能区按压状态重绘
                if (!isSameWorkbook)
                {
                    // 刷新 Ribbon 控件状态，依靠短效纯内存缓存实现平滑响应
                    RibbonController.InvalidateRibbon();
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[EVENT] OnWorkbookActivate 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 响应工作簿失活事件，监控焦点离开工作簿瞬间的安全避让
        /// </summary>
        /// <param name="wb">失活的工作簿对象</param>
        private static void OnWorkbookDeactivate(Microsoft.Office.Interop.Excel.Workbook wb)
        {
            try
            {
                // 核心防护：工作簿失活切出期间临时隐匿聚光灯避让，防止跨窗口销毁时死锁
                if (ExcelServices.IsSpotlightEnabled)
                {
                    // 临时隐匿浮窗并安全断开视口钩子
                    ExcelServices.TemporarilyHideSpotlightForTransition();
                }
            }
            catch (Exception ex)
            {
                // 记录失活事件异常
                LogHelper.WriteLog($"[EVENT] OnWorkbookDeactivate 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 响应工作簿关闭前事件，保护工作簿安全关闭与聚光灯避让
        /// </summary>
        /// <param name="wb">即将关闭的工作簿对象</param>
        /// <param name="cancel">是否取消关闭</param>
        private static void OnWorkbookBeforeClose(Microsoft.Office.Interop.Excel.Workbook wb, ref bool cancel)
        {
            try
            {
                // 核心防护：若开启了聚光灯，在工作簿关闭临界期立即执行安全隐匿避让，消除 COM 竞争与鼠标转圈卡顿
                if (ExcelServices.IsSpotlightEnabled)
                {
                    // 临时隐匿浮窗并安全断开视口钩子
                    ExcelServices.TemporarilyHideSpotlightForTransition();
                }
            }
            catch (Exception ex)
            {
                // 记录关闭前事件异常
                LogHelper.WriteLog($"[EVENT] OnWorkbookBeforeClose 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 响应工作簿保存成功事件，刷新 Ribbon 本机项目勾选按压状态
        /// </summary>
        /// <param name="wb">已保存的工作簿对象</param>
        /// <param name="success">是否保存成功</param>
        private static void OnWorkbookAfterSave(Microsoft.Office.Interop.Excel.Workbook wb, bool success)
        {
            try
            {
                // 若保存成功，使本地项目缓存失效并刷新 Ribbon 状态
                if (success)
                {
                    // 清除路径比对短效缓存
                    ExcelServices.InvalidateLocalProjectCache();
                    // 即时重绘功能区控件按压高亮状态
                    RibbonController.InvalidateRibbon();
                }
            }
            catch { }
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

                // 条件 1: “元件汇总表”中双击 D 列 (规格型号)
                bool isSummaryColD = sheetName == ComponentMatchDefaults.ComponentSummarySheetName &&
                    target.Rows.Count == 1 && target.Columns.Count == 1 && target.Column == 4;

                // 条件 2: 常规分类明细表中双击 C 列 (规格型号)
                bool isCategoryColC = sheetName != ComponentMatchDefaults.ComponentSummarySheetName &&
                    target.Rows.Count == 1 && target.Columns.Count == 1 && target.Column == 3;

                if (isSummaryColD || isCategoryColC)
                {
                    // 方案 A: 双击完全归还 Excel 原生文本就地编辑
                    // 先平滑隐藏可能已弹出的物料联想悬浮框与原生覆盖输入框，避免遮挡单元格视线
                    ExcelServices.HideComponentMatchOverlay();
                    // 隐藏原生覆盖输入框
                    ExcelServices.HideSmartInputOverlay();

                    // 保持 cancel = false，100% 允许 Excel 正常双击进入单元格进行文本光标编辑
                    cancel = false;
                    return;
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

                // 核心安全守门：若当前工作簿不是成套工程工作簿 (不含【项目信息】表)，直接退出，杜绝常规表格受任何联动干扰
                if (!Tool.IsProjectWorkbook(wb)) return;

                // 核心感知：只要工作表发生单元格或行列变动，立即失效该表的元器件行缓存，防止删行插行后行号错位
                ExcelServices.InvalidateCategoryRowCache(sh.Name);

                // 1. 优先尝试处理箱柜关键行（Sum 汇总行、Det 明细行、Tolsum 总计行）之间的 8 组双向数据绑定联动
                if (TryHandleCabinetBiDirectionalSync(wb, sh, target))
                {
                    // 若命中双向绑定同步，直接退出避免重复触发
                    return;
                }

                // 2. 处理第 3 列 (C列 - 元器件规格型号) 修改时的智能属性联动回填
                if (target.Column == 3 && target.Cells.Count == 1)
                {
                    // 手动完成编辑后，平滑隐藏物料联想下拉框
                    ExcelServices.HideComponentMatchOverlay();
                    // 手动完成编辑后，同步平滑隐藏原生覆盖输入框
                    ExcelServices.HideSmartInputOverlay();

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
                    // 尝试学习当前行元器件数据 (若为新物料或属性更全则增量沉淀至候选库)
                    ExcelServices.CheckAndLearnComponentFromRow(sh, target.Row);
                    return;
                }

                // 3. 处理元件汇总表 D 列手动输入或就地编辑完成时的收尾联动
                if (target.Column == 4 && target.Cells.Count == 1)
                {
                    // 手动完成编辑后，平滑隐藏物料联想下拉框
                    ExcelServices.HideComponentMatchOverlay();
                    return;
                }

                // 4. 处理分类明细表元器件同行属性列 (B 列:名称, D 列:厂家, E 列:单位, G 列:单价) 修改时的自动学习自愈
                int modCol = target.Column;
                // 校验改动列是否属于元器件核心属性列且为单单元格
                if ((modCol == 2 || modCol == 4 || modCol == 5 || modCol == 7) && target.Cells.Count == 1)
                {
                    // 尝试对变动行进行元器件质量门控检测与增量学习
                    ExcelServices.CheckAndLearnComponentFromRow(sh, target.Row);
                }
            }
            catch { }
        }

        // 顶部汇总行 (Sum) 同步至底部明细信息行 (Det) 的固定列号映射 (--硬编码: 模板固定列号映射--)
        private static readonly Dictionary<int, int> _sumToDetColumnMap = new Dictionary<int, int>
        {
            { 2, 2 },   // 柜号: Sum B(2) -> Det B(2)
            { 3, 7 },   // 箱柜名称: Sum C(3) -> Det G(7)
            { 4, 4 },   // 箱柜型号: Sum D(4) -> Det D(4)
            { 9, 9 },   // 备注: Sum I(9) -> Det I(9)
            { 13, 11 }, // 箱柜尺寸: Sum M(13) -> Det K(11)
            { 14, 13 }, // 类别: Sum N(14) -> Det M(13)
            { 17, 15 }  // 图号: Sum Q(17) -> Det O(15)
        };

        // 底部明细信息行 (Det) 反向同步至顶部汇总行 (Sum) 的固定列号映射 (--硬编码: 模板固定列号映射--)
        private static readonly Dictionary<int, int> _detToSumColumnMap = new Dictionary<int, int>
        {
            { 2, 2 },   // 柜号: Det B(2) -> Sum B(2)
            { 4, 4 },   // 箱柜型号: Det D(4) -> Sum D(4)
            { 5, 4 },   // 箱柜型号合并区: Det E(5) -> Sum D(4)
            { 7, 3 },   // 箱柜名称: Det G(7) -> Sum C(3)
            { 8, 3 },   // 箱柜名称合并区: Det H(8) -> Sum C(3)
            { 9, 9 },   // 备注: Det I(9) -> Sum I(9)
            { 11, 13 }, // 箱柜尺寸: Det K(11) -> Sum M(13)
            { 13, 14 }, // 类别: Det M(13) -> Sum N(14)
            { 15, 17 }  // 图号: Det O(15) -> Sum Q(17)
        };

        // 数量固定列号 F 列(6) 与双向同步最大列号 Q 列(17) (--硬编码: 模板固定列号--)
        private const int CabinetQtyColIndex = 6;
        private const int CabinetSyncMaxColIndex = 17;

        /// <summary>
        /// 单台箱柜关键物理行号快照 (Sum 汇总行 / Det 明细信息行 / Tolsum 总计行)
        /// </summary>
        private readonly struct CabinetRowAnchor
        {
            // 箱柜唯一序号 K
            public readonly int CabinetK;
            // 顶部汇总行物理行号
            public readonly int SumRow;
            // 底部明细信息行物理行号
            public readonly int DetRow;
            // 底部总计行物理行号
            public readonly int TolsumRow;

            /// <summary>
            /// 初始化单台箱柜行号锚点快照
            /// </summary>
            public CabinetRowAnchor(int cabinetK, int sumRow, int detRow, int tolsumRow)
            {
                // 赋值箱柜序号与三大关键物理行号
                CabinetK = cabinetK;
                SumRow = sumRow;
                DetRow = detRow;
                TolsumRow = tolsumRow;
            }
        }

        /// <summary>
        /// 单张工作表箱柜行号 O(1) 快速索引容器
        /// </summary>
        private sealed class SheetCabinetRowLookup
        {
            // 按 Sum 行号索引箱柜锚点
            public readonly Dictionary<int, CabinetRowAnchor> SumMap = new Dictionary<int, CabinetRowAnchor>();
            // 按 Det 行号索引箱柜锚点
            public readonly Dictionary<int, CabinetRowAnchor> DetMap = new Dictionary<int, CabinetRowAnchor>();
            // 按 Tolsum 行号索引箱柜锚点
            public readonly Dictionary<int, CabinetRowAnchor> TolsumMap = new Dictionary<int, CabinetRowAnchor>();
            // 按箱柜序号 K 索引箱柜锚点 (供超链接兜底查询使用)
            public readonly Dictionary<int, CabinetRowAnchor> ByIndex = new Dictionary<int, CabinetRowAnchor>();
            // 当前工作表涉及箱柜锚点的最大物理行号 (用于整列选中时裁剪扫描边界)
            public int MaxRow;
            // 当前工作表首个底部明细信息行号 (用于判定超链接兜底是否位于顶部汇总区)
            public int MinDetRow = int.MaxValue;
        }

        /// <summary>
        /// 尝试处理箱柜关键行（Sum 汇总行、Det 明细信息行、Tolsum 总计行）之间的 8 组双向数据同步
        /// 支持单格编辑、多行/多列拖动填充、区域复制粘贴、批量清空及 Ctrl 多选区同步
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
        /// <param name="target">被修改的目标单元格或区域 Range</param>
        /// <returns>若成功命中箱柜双向同步列返回 true，否则返回 false</returns>
        private static bool TryHandleCabinetBiDirectionalSync(
            Microsoft.Office.Interop.Excel.Workbook wb,
            Microsoft.Office.Interop.Excel.Worksheet sh,
            Microsoft.Office.Interop.Excel.Range target)
        {
            try
            {
                // 校验关键 COM 句柄有效性
                if (wb == null || sh == null || target == null || _excelApp == null) return false;

                // 提取工作表名称并拦截系统保留表与汇总报表
                string sheetName = Convert.ToString(sh.Name) ?? string.Empty;
                if (Tool.IsReservedOrReportSheet(sheetName)) return false;

                // 1. 快速预检所有选区 Areas 的列交集，未涉及同步列时 0ms 直接退出
                var areaInfos = new List<(Microsoft.Office.Interop.Excel.Range Area, int StartRow, int EndRow, int StartCol, int EndCol)>();
                bool hasSyncableColumn = false;
                foreach (Microsoft.Office.Interop.Excel.Range area in target.Areas)
                {
                    // 提取当前子区域起止行列边界
                    int sRow = area.Row;
                    int eRow = sRow + area.Rows.Count - 1;
                    int sCol = area.Column;
                    int eCol = sCol + area.Columns.Count - 1;
                    areaInfos.Add((area, sRow, eRow, sCol, eCol));

                    // 若尚未命中同步列，则判断当前区域列范围是否覆盖任意监听列
                    if (!hasSyncableColumn)
                    {
                        int cFrom = Math.Max(2, sCol);
                        int cTo = Math.Min(CabinetSyncMaxColIndex, eCol);
                        for (int c = cFrom; c <= cTo; c++)
                        {
                            // 只要命中任意同步列即标记通过
                            if (IsCabinetSyncColumn(c))
                            {
                                hasSyncableColumn = true;
                                break;
                            }
                        }
                    }
                }
                if (!hasSyncableColumn) return false;

                // 2. 一次性构建当前工作表箱柜行号 O(1) 内存映射表，杜绝循环遍历 Names
                var lookup = BuildSheetCabinetRowLookup(wb, sh);
                if (lookup == null || lookup.MaxRow <= 0) return false;

                // 3. 批量解析各区域内存数组并收集去重后的待同步写入队列
                var pendingWrites = new Dictionary<(int Row, int Col), object?>();
                var hyperlinkFallbackCache = new Dictionary<int, CabinetRowAnchor?>();
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;
                bool hitCabinetSyncCell = false;

                foreach (var info in areaInfos)
                {
                    // 裁剪行列至当前工作表箱柜有效边界，防止整行/整列操作读取海量空单元格
                    int effStartRow = Math.Max(1, info.StartRow);
                    int effEndRow = Math.Min(lookup.MaxRow, info.EndRow);
                    int effStartCol = Math.Max(2, info.StartCol);
                    int effEndCol = Math.Min(CabinetSyncMaxColIndex, info.EndCol);
                    if (effStartRow > effEndRow || effStartCol > effEndCol) continue;

                    // 一次性批量读取当前有效矩形区域的 Value 到内存 (单格为标量，多格为 1-based 二维数组)
                    object? rawBlock = ReadAreaValueBlock(
                        sh, info.Area, info.StartRow, info.EndRow, info.StartCol, info.EndCol,
                        effStartRow, effEndRow, effStartCol, effEndCol);

                    // 逐行匹配 Sum / Det / Tolsum 锚点并收集待写入映射
                    for (int r = effStartRow; r <= effEndRow; r++)
                    {
                        int rowOffset = r - effStartRow;

                        // 优先命中顶部汇总行 (Sum -> Det / Tolsum)
                        if (lookup.SumMap.TryGetValue(r, out var sumAnchor) ||
                            (r < lookup.MinDetRow && TryResolveSumAnchorByHyperlink(
                                sh, r, lookup, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix,
                                hyperlinkFallbackCache, out sumAnchor)))
                        {
                            // 收集当前汇总行修改列对应的底部同步写入项
                            if (CollectSumRowWrites(sumAnchor, effStartCol, effEndCol, rawBlock, rowOffset, pendingWrites))
                            {
                                hitCabinetSyncCell = true;
                            }
                            continue;
                        }

                        // 次选命中底部明细信息行 (Det -> Sum)
                        if (lookup.DetMap.TryGetValue(r, out var detAnchor))
                        {
                            // 收集当前明细行修改列对应的顶部同步写入项
                            if (CollectDetRowWrites(sh, r, detAnchor, effStartCol, effEndCol, rawBlock, rowOffset, pendingWrites))
                            {
                                hitCabinetSyncCell = true;
                            }
                            continue;
                        }

                        // 最后命中底部总计行 F 列 (Tolsum F -> Sum F)
                        if (lookup.TolsumMap.TryGetValue(r, out var tolsumAnchor))
                        {
                            // 检查数量 F 列是否落在修改列区间内
                            if (CabinetQtyColIndex >= effStartCol && CabinetQtyColIndex <= effEndCol)
                            {
                                hitCabinetSyncCell = true;
                                if (tolsumAnchor.SumRow > 0)
                                {
                                    // 提取 F 列修改值并加入写入队列
                                    object? qtyVal = ExtractBlockCellValue(rawBlock, rowOffset, CabinetQtyColIndex - effStartCol);
                                    pendingWrites[(tolsumAnchor.SumRow, CabinetQtyColIndex)] = qtyVal;
                                }
                            }
                        }
                    }
                }

                // 4. 在外层统一关闭事件与屏幕刷新后，一次性批量写回所有目标单元格
                if (pendingWrites.Count > 0)
                {
                    BatchSyncCellValues(sh, pendingWrites);
                }

                return hitCabinetSyncCell;
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 判断列号是否属于箱柜双向同步监听列
        /// </summary>
        private static bool IsCabinetSyncColumn(int col)
        {
            // 命中数量 F 列(6)、Sum 同步列或 Det 同步列之一即返回 true
            return col == CabinetQtyColIndex ||
                   _sumToDetColumnMap.ContainsKey(col) ||
                   _detToSumColumnMap.ContainsKey(col);
        }

        /// <summary>
        /// 解析箱柜锚点 Range 的实时物理行号，免疫手动插删行引起的缓存偏移
        /// </summary>
        private static int ResolveAnchorLiveRow(dynamic? anchorRange, int cachedRow)
        {
            // 存在 COM Range 对象时直接读取其随表格自动平移后的实时 Row
            if (anchorRange != null)
            {
                try
                {
                    // 返回最新物理行号
                    return Convert.ToInt32(anchorRange.Row);
                }
                catch
                {
                    // Range 若因整行被删失效 (#REF!) 则返回 0
                    return 0;
                }
            }

            // 无 Range 对象时回退使用静态行号
            return cachedRow;
        }

        /// <summary>
        /// 一次性构建当前工作表箱柜关键物理行号的 O(1) 内存索引表
        /// </summary>
        private static SheetCabinetRowLookup? BuildSheetCabinetRowLookup(
            Microsoft.Office.Interop.Excel.Workbook wb,
            Microsoft.Office.Interop.Excel.Worksheet sh)
        {
            // 校验工作表名并跳过系统保留表与报表
            string sheetName = Convert.ToString(sh.Name) ?? string.Empty;
            if (Tool.IsReservedOrReportSheet(sheetName)) return null;

            // 优先使用已纳管分类表缓存高速获取箱柜锚点
            var validCabinets = Tool.GetSheetValidCabinets((object)sh, (object)wb);

            // 若尚未在【项目信息】登记但表内仍带有定义名称，则按单表兜底解析
            if (validCabinets == null || validCabinets.Count == 0)
            {
                bool hasAnyNames = false;
                try
                {
                    // 快速判断工作表或工作簿是否存在定义名称
                    hasAnyNames = (sh.Names != null && sh.Names.Count > 0) || (wb.Names != null && wb.Names.Count > 0);
                }
                catch { }
                if (!hasAnyNames) return null;

                // 收集并构建当前表箱柜映射
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;
                var rawNames = Tool.CollectAllDefinedNames(wb, sh, autoRebuildIfEmpty: false);
                validCabinets = Tool.BuildCabinetMap(rawNames, sheetName, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
            }

            // 无有效箱柜时直接返回 null
            if (validCabinets == null || validCabinets.Count == 0) return null;

            var lookup = new SheetCabinetRowLookup();
            bool cacheDrifted = false;

            foreach (var kvp in validCabinets)
            {
                // 读取实时物理行号并检测缓存是否因手动插删行产生偏移
                var cab = kvp.Value;
                int sumR = ResolveAnchorLiveRow(cab.Sum, cab.SumRow);
                int detR = ResolveAnchorLiveRow(cab.Det, cab.DetRow);
                int tolsumR = ResolveAnchorLiveRow(cab.Tolsum, cab.TolsumRow);

                // 任意行号偏移即标记需要刷新缓存
                if ((cab.SumRow > 0 && sumR > 0 && sumR != cab.SumRow) ||
                    (cab.DetRow > 0 && detR > 0 && detR != cab.DetRow) ||
                    (cab.TolsumRow > 0 && tolsumR > 0 && tolsumR != cab.TolsumRow))
                {
                    cacheDrifted = true;
                }

                // 生成只读快照并写入序号索引
                var rowAnchor = new CabinetRowAnchor(kvp.Key, sumR, detR, tolsumR);
                lookup.ByIndex[kvp.Key] = rowAnchor;

                // 注册顶部汇总行映射
                if (sumR > 0)
                {
                    lookup.SumMap[sumR] = rowAnchor;
                    if (sumR > lookup.MaxRow) lookup.MaxRow = sumR;
                }

                // 注册底部明细行映射
                if (detR > 0)
                {
                    lookup.DetMap[detR] = rowAnchor;
                    if (detR > lookup.MaxRow) lookup.MaxRow = detR;
                    if (detR < lookup.MinDetRow) lookup.MinDetRow = detR;
                }

                // 注册底部总计行映射
                if (tolsumR > 0)
                {
                    lookup.TolsumMap[tolsumR] = rowAnchor;
                    if (tolsumR > lookup.MaxRow) lookup.MaxRow = tolsumR;
                }
            }

            // 若检测到缓存行号漂移则立即失效，保证下次查询自愈对齐
            if (cacheDrifted)
            {
                Tool.InvalidateSheetCabinetsCache(sheetName);
            }

            return lookup;
        }

        /// <summary>
        /// 兜底方案：当顶部汇总行缺失 Cab_Sum 定义名称时，依据 A 列超链接反查内存箱柜锚点
        /// </summary>
        private static bool TryResolveSumAnchorByHyperlink(
            Microsoft.Office.Interop.Excel.Worksheet sh,
            int row,
            SheetCabinetRowLookup lookup,
            string sumPrefix,
            string detPrefix,
            string subsumPrefix,
            string tolsumPrefix,
            Dictionary<int, CabinetRowAnchor?> rowCache,
            out CabinetRowAnchor anchor)
        {
            // 优先命中同行内存缓存，防止多列循环重复读取 A 列超链接
            if (rowCache.TryGetValue(row, out var cachedAnchor))
            {
                anchor = cachedAnchor ?? default;
                return cachedAnchor.HasValue;
            }

            // 预置空哨兵缓存，避免非超链接行重复查询
            rowCache[row] = null;
            anchor = default;

            try
            {
                // 提取 A 列超链接 SubAddress
                var aCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[row, 1];
                if (aCell.Hyperlinks == null || aCell.Hyperlinks.Count == 0) return false;
                string subAddr = aCell.Hyperlinks[1].SubAddress ?? string.Empty;

                // 仅处理指向明细前缀的超链接并解析箱柜序号 K
                if (string.IsNullOrEmpty(subAddr) || !subAddr.Contains(detPrefix)) return false;
                string detTag = ExtractTag(subAddr, detPrefix);
                int k = Tool.ExtractIndexFromName(detTag, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                if (k <= 0)
                {
                    // 兜底从完整子地址提取序号 K
                    k = Tool.ExtractIndexFromName(subAddr, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                }

                // 在序号字典中 O(1) 定位对应箱柜锚点
                if (k > 0 && lookup.ByIndex.TryGetValue(k, out var matched))
                {
                    rowCache[row] = matched;
                    anchor = matched;
                    return true;
                }
            }
            catch { }

            return false;
        }

        /// <summary>
        /// 一次性批量读取有效边界范围内的单元格 Value 到内存
        /// </summary>
        private static object? ReadAreaValueBlock(
            Microsoft.Office.Interop.Excel.Worksheet sh,
            Microsoft.Office.Interop.Excel.Range area,
            int startRow, int endRow, int startCol, int endCol,
            int effStartRow, int effEndRow, int effStartCol, int effEndCol)
        {
            // 边界未被裁剪时直接读取 area.Value，避免额外创建 Range COM 代理
            if (effStartRow == startRow && effEndRow == endRow && effStartCol == startCol && effEndCol == endCol)
            {
                return area.Value;
            }

            // 按裁剪后的矩形边界构建子区域并一次性读入内存
            var subRange = (Microsoft.Office.Interop.Excel.Range)sh.Range[
                sh.Cells[effStartRow, effStartCol],
                sh.Cells[effEndRow, effEndCol]
            ];
            return subRange.Value;
        }

        /// <summary>
        /// 从单格标量或 Excel 1-based 二维数组中安全读取指定偏移位置的单元格值
        /// </summary>
        private static object? ExtractBlockCellValue(object? rawBlock, int rowOffset, int colOffset)
        {
            // 多格区域为二维数组，结合真实数组下界安全寻址
            if (rawBlock is object[,] matrix)
            {
                int rIdx = matrix.GetLowerBound(0) + rowOffset;
                int cIdx = matrix.GetLowerBound(1) + colOffset;
                if (rIdx <= matrix.GetUpperBound(0) && cIdx <= matrix.GetUpperBound(1))
                {
                    return matrix[rIdx, cIdx];
                }
                return null;
            }

            // 单格区域直接返回标量值
            return (rowOffset == 0 && colOffset == 0) ? rawBlock : null;
        }

        /// <summary>
        /// 收集顶部汇总行 (Sum) 在修改列区间内需要同步至底部明细/总计行的写操作
        /// </summary>
        private static bool CollectSumRowWrites(
            CabinetRowAnchor anchor,
            int effStartCol,
            int effEndCol,
            object? rawBlock,
            int rowOffset,
            Dictionary<(int Row, int Col), object?> pendingWrites)
        {
            bool hit = false;

            // 遍历当前行在修改选区内的每一列
            for (int c = effStartCol; c <= effEndCol; c++)
            {
                int colOffset = c - effStartCol;

                // 1. 若修改的是 F 列(6: 数量)，同步至对应 Tolsum 行 F 列
                if (c == CabinetQtyColIndex)
                {
                    hit = true;
                    if (anchor.TolsumRow > 0)
                    {
                        // 直接按数值/文本覆盖写回，无公式依赖
                        pendingWrites[(anchor.TolsumRow, CabinetQtyColIndex)] = ExtractBlockCellValue(rawBlock, rowOffset, colOffset);
                    }
                    continue;
                }

                // 2. 若命中 Sum -> Det 映射表，同步至对应 Det 行指定列
                if (_sumToDetColumnMap.TryGetValue(c, out int targetDetCol))
                {
                    hit = true;
                    if (anchor.DetRow > 0)
                    {
                        // 写入待同步字典 (相同目标单元格自动去重)
                        pendingWrites[(anchor.DetRow, targetDetCol)] = ExtractBlockCellValue(rawBlock, rowOffset, colOffset);
                    }
                }
            }

            return hit;
        }

        /// <summary>
        /// 收集底部明细信息行 (Det) 在修改列区间内需要反向同步至顶部汇总行的写操作
        /// </summary>
        private static bool CollectDetRowWrites(
            Microsoft.Office.Interop.Excel.Worksheet sh,
            int row,
            CabinetRowAnchor anchor,
            int effStartCol,
            int effEndCol,
            object? rawBlock,
            int rowOffset,
            Dictionary<(int Row, int Col), object?> pendingWrites)
        {
            bool hit = false;

            // 按从左到右顺序遍历修改列，确保合并主列 (D/G) 优先于从属列 (E/H)
            for (int c = effStartCol; c <= effEndCol; c++)
            {
                if (!_detToSumColumnMap.TryGetValue(c, out int targetSumCol)) continue;
                hit = true;
                if (anchor.SumRow <= 0) continue;

                var destKey = (anchor.SumRow, targetSumCol);
                // 若主列 (D/G) 已在同批生成同步项，跳过合并从属列 (E/H) 防止空值覆写主值
                if ((c == 5 || c == 8) && pendingWrites.ContainsKey(destKey)) continue;

                // 若单改合并从属列且值为空，检查是否为合并单元格从属格
                object? val = ExtractBlockCellValue(rawBlock, rowOffset, c - effStartCol);
                if ((c == 5 || c == 8) && val == null && IsCellMerged(sh, row, c)) continue;

                pendingWrites[destKey] = val;
            }

            return hit;
        }

        /// <summary>
        /// 安全判断指定单元格是否属于合并单元格区域
        /// </summary>
        private static bool IsCellMerged(Microsoft.Office.Interop.Excel.Worksheet sh, int row, int col)
        {
            try
            {
                // 读取 COM MergeCells 布尔标识
                object? merged = ((Microsoft.Office.Interop.Excel.Range)sh.Cells[row, col]).MergeCells;
                return merged is bool isMerged && isMerged;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 在挂起 Excel 全局事件 (并在多格写入时挂起屏幕刷新) 的前提下批量写回同步值
        /// </summary>
        private static void BatchSyncCellValues(
            Microsoft.Office.Interop.Excel.Worksheet sh,
            Dictionary<(int Row, int Col), object?> pendingWrites)
        {
            if (_excelApp == null || pendingWrites.Count == 0) return;

            // 保存原有事件与屏幕刷新状态
            bool prevEvents = _excelApp.EnableEvents;
            bool prevUpdating = _excelApp.ScreenUpdating;

            // 关闭全局事件避免递归触发；批量写入大于 1 格时关闭屏幕刷新防止界面闪烁
            _excelApp.EnableEvents = false;
            if (pendingWrites.Count > 1)
            {
                _excelApp.ScreenUpdating = false;
            }

            try
            {
                foreach (var kvp in pendingWrites)
                {
                    try
                    {
                        // 安全写回目标单元格新值
                        ((Microsoft.Office.Interop.Excel.Range)sh.Cells[kvp.Key.Row, kvp.Key.Col]).Value = kvp.Value;
                    }
                    catch { }
                }
            }
            finally
            {
                // 严格成对恢复 Excel 事件与屏幕刷新状态
                _excelApp.EnableEvents = prevEvents;
                _excelApp.ScreenUpdating = prevUpdating;
            }
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

                // 优先读取当前右键菜单配置项：是否启用自定义业务右键菜单
                bool useCustomMenu = ConfigManager.Instance.Current.Excel.UseCustomContextMenu;

                // 若当前使用的是系统菜单 (Excel 原生右键菜单模式)：
                // 彻底放行原生右键菜单，移除同步 CutCopyMode = 0 清除，不做任何拦截与选区修改，保证 0 延迟、0 干扰、100% 原生流畅
                if (!useCustomMenu)
                {
                    // 彻底放行原生右键菜单弹窗
                    cancel = false;
                    // 直接返回，绝不干扰 Excel 原生剪贴板跑马灯状态机与上下文菜单管道
                    return;
                }

                // 转换当前触发右击的工作表强类型对象
                Microsoft.Office.Interop.Excel.Worksheet? activeSheet = shObj as Microsoft.Office.Interop.Excel.Worksheet;
                if (activeSheet == null) return;

                // 核心业务表智能路由：判定当前工作表是否应弹出业务专属右键菜单
                // 规则：外部文件、项目信息表、元件分布表、封面、目录等无关表，自动放行系统原生右键菜单
                if (!ExcelServices.ShouldUseCustomContextMenu(activeSheet))
                {
                    // 彻底放行原生系统右键菜单弹窗
                    cancel = false;
                    // 直接返回，绝不干扰 Excel 原生剪贴板跑马灯状态机与上下文菜单管道
                    return;
                }

                // 获取所属工作簿对象
                Microsoft.Office.Interop.Excel.Workbook? wb = null;
                try { wb = activeSheet.Parent as Microsoft.Office.Interop.Excel.Workbook; } catch { }

                // 核心修复：若当前处于复制或剪切模式 (单元格四周带虚线框)，且用户右击的单元格落在当前被选中的选区内部
                // 说明用户在被复制的源单元格自身上点击右键 (例如准备按当前值进行右键筛选)，绝非要在自身上执行粘贴
                // 此时提前清除挂起的 CutCopyMode，消除跑马灯虚线框，避免带着虚线框触发右键筛选导致状态机挂起、卡死无法输入
                try
                {
                    // 检测是否处于复制或剪切状态
                    if (_excelApp.CutCopyMode != 0)
                    {
                        // 提取当前选区对象
                        Microsoft.Office.Interop.Excel.Range? currentSel = _excelApp.Selection as Microsoft.Office.Interop.Excel.Range;
                        // 若右击目标单元格落在当前选区交集内部，判定为源单元格自身操作
                        if (currentSel != null && _excelApp.Intersect(currentSel, target) != null)
                        {
                            // 清除挂起的复制剪切状态，平滑消除虚线框
                            _excelApp.CutCopyMode = (Microsoft.Office.Interop.Excel.XlCutCopyMode)0;
                        }
                    }
                }
                catch { }

                // ------------------ 【以下为 WebView 2 业务专属菜单模式】 ------------------
                // 确保右键点击的目标单元格处于活动选中状态，对齐 Excel 原生行为
                try
                {
                    // 获取当前活动工作簿选区
                    Microsoft.Office.Interop.Excel.Range? sel = _excelApp.Selection as Microsoft.Office.Interop.Excel.Range;
                    // 若当前无有效选区，或右击的目标单元格不在现有选区交集范围内
                    if (sel == null || _excelApp.Intersect(sel, target) == null)
                    {
                        // 显式激活选中目标单元格，保证后续剪切/复制/插入/删除等原生指令精确生效
                        target.Select();
                    }
                }
                catch { }

                // 1. 完全拦截 Excel 原生右键菜单弹窗
                cancel = true;

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
        /// Excel-DNA 宏入口：响应 Excel 原生右键菜单中“添加到本地个人库”指令
        /// </summary>
        [ExcelCommand]
        public static void MacroAddToPersonalDb()
        {
            try
            {
                // 调度业务服务层打开添加到个人库核对确认窗体
                ExcelServices.OpenAddToPersonalDbDialog();
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"右键添加到本地个人库宏执行异常: {ex.Message}");
            }
        }

        /// <summary>
        /// Excel-DNA 宏入口：响应 Excel 右键菜单中“拆分改型元件”指令
        /// </summary>
        [ExcelCommand]
        public static void MacroSplitComponent()
        {
            try
            {
                // 调度业务服务层打开拆分改型微型核对窗体
                ExcelServices.OpenComponentSplitDialog();
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"右键拆分改型元件宏执行异常: {ex.Message}");
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
        /// 安全清理插件注册的所有右键菜单控件，保持 Excel 原始环境干净整洁
        /// </summary>
        public static void RemoveContextMenuControls()
        {
            try
            {
                // 校验 Excel 全局应用实例是否有效
                if (_excelApp == null) return;
                // 读取配置中定义的唯一 Tag 标识 (新建箱柜按钮标识)
                string menuTag = ConfigManager.Instance.Current.Excel.NewCabinetMenuTag ?? "CT_BTN_NEW_CABINET";
                // 读取配置中定义的唯一 Tag 标识 (物料匹配按钮标识)
                string matchTag = ConfigManager.Instance.Current.Excel.ParseMatchComponentsMenuTag ?? "CT_BTN_PARSE_MATCH_COMPONENTS";
                // 预设匹配规则设置按钮唯一 Tag 标识
                const string settingTag = "CT_BTN_OPEN_MATCH_SETTING";

                // 将 Excel Application 转换为 dynamic 动态对象
                dynamic dynApp = (dynamic)_excelApp;
                // 获取 Excel 顶层 CommandBars 集合句柄
                dynamic commandBars = dynApp.CommandBars;
                // 校验 CommandBars 对象是否就绪
                if (commandBars == null) return;

                // 直接按名称获取 "Cell" CommandBar 进行精准清理，避免遍历 200+ CommandBars 引发 COM 卡顿与死锁
                string[] cellBarNames = new string[] { "Cell" };
                // 遍历预设的右键上下文菜单名称数组
                foreach (string barName in cellBarNames)
                {
                    try
                    {
                        // 直接通过键名安全检索 CommandBar 菜单
                        dynamic bar = commandBars[barName];
                        // 确认右键菜单对象有效
                        if (bar != null)
                        {
                            try
                            {
                                // 查找并删除已挂载的“新建箱柜”按钮控件
                                dynamic? ctrl = bar.FindControl(1, Type.Missing, menuTag);
                                // 若存在则执行删除
                                if (ctrl != null) ctrl.Delete(true);
                            }
                            catch { }

                            try
                            {
                                // 查找并删除已挂载的“识别参数并匹配物料”按钮控件
                                dynamic? matchCtrl = bar.FindControl(1, Type.Missing, matchTag);
                                // 若存在则执行删除
                                if (matchCtrl != null) matchCtrl.Delete(true);
                            }
                            catch { }

                            try
                            {
                                // 查找并删除已挂载的“物料匹配规则与品牌设置”按钮控件
                                dynamic? settingCtrl = bar.FindControl(1, Type.Missing, settingTag);
                                // 若存在则执行删除
                                if (settingCtrl != null) settingCtrl.Delete(true);
                            }
                            catch { }
                        }
                    }
                    catch { }
                }

                // 清空内存缓存中的按钮动态引用集合
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
        /// Excel-DNA 宏入口：撤销上一步操作 (仅供宏过程调用，不占用快捷键)
        /// </summary>
        [ExcelCommand]
        public static void MacroUndoAction()
        {
            try
            {
                // 若插件撤销管理器中存在可撤销记录，优先执行插件撤销
                if (ExcelServices.CanUndo)
                {
                    // 调度插件撤销
                    ExcelServices.Undo();
                }
                else
                {
                    // 若插件撤销栈为空，安全放行给 Excel 原生撤销能力
                    dynamic? app = ExcelDnaUtil.Application;
                    if (app != null)
                    {
                        try
                        {
                            // 执行原生撤销
                            app.Undo();
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录宏执行异常
                LogHelper.WriteLog($"MacroUndoAction 执行异常: {ex.Message}");
            }
        }

        /// <summary>
        /// Excel-DNA 宏入口：重做/还原上一步操作 (仅供宏过程调用，不占用快捷键)
        /// </summary>
        [ExcelCommand]
        public static void MacroRedoAction()
        {
            try
            {
                // 若插件重做栈中存在可还原记录，优先执行插件重做
                if (ExcelServices.CanRedo)
                {
                    // 调度插件重做
                    ExcelServices.Redo();
                }
                else
                {
                    // 若插件重做栈为空，安全放行给 Excel 原生重复/重做
                    dynamic? app = ExcelDnaUtil.Application;
                    if (app != null)
                    {
                        try
                        {
                            // 执行原生 Repeat
                            app.Repeat();
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录宏执行异常
                LogHelper.WriteLog($"MacroRedoAction 执行异常: {ex.Message}");
            }
        }

        /// <summary>
        /// Excel-DNA 宏入口：复制元件 (响应快捷键 Ctrl+Shift+C)
        /// </summary>
        [ExcelCommand]
        public static void MacroCopyComponentAction()
        {
            try
            {
                // 调度执行成套元件复制
                ExcelServices.CopyComponentRow();
            }
            catch (Exception ex)
            {
                // 记录宏执行异常
                LogHelper.WriteLog($"MacroCopyComponentAction 执行异常: {ex.Message}");
            }
        }

        /// <summary>
        /// Excel-DNA 宏入口：插入复制/剪切的元件 (响应快捷键 Ctrl+Shift+V)
        /// </summary>
        [ExcelCommand]
        public static void MacroInsertCopiedComponentAction()
        {
            try
            {
                // 调度执行成套元件插入
                ExcelServices.InsertCopiedOrCutComponentRow();
            }
            catch (Exception ex)
            {
                // 记录宏执行异常
                LogHelper.WriteLog($"MacroInsertCopiedComponentAction 执行异常: {ex.Message}");
            }
        }

        /// <summary>
        /// Excel-DNA 宏入口：删除元件 (响应快捷键 Ctrl+Shift+D)
        /// </summary>
        [ExcelCommand]
        public static void MacroDeleteComponentAction()
        {
            try
            {
                // 调度执行成套元件删除
                ExcelServices.DeleteComponentRow();
            }
            catch (Exception ex)
            {
                // 记录宏执行异常
                LogHelper.WriteLog($"MacroDeleteComponentAction 执行异常: {ex.Message}");
            }
        }

        /// <summary>
        /// Excel-DNA 宏入口：剪切元件 (响应快捷键 Ctrl+Shift+X)
        /// </summary>
        [ExcelCommand]
        public static void MacroCutComponentAction()
        {
            try
            {
                // 调度执行成套元件剪切
                ExcelServices.CutComponentRow();
            }
            catch (Exception ex)
            {
                // 记录宏执行异常
                LogHelper.WriteLog($"MacroCutComponentAction 执行异常: {ex.Message}");
            }
        }
    }
}
