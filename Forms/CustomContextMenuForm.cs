using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo;
using ExcelAddInDemo.Models;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的现代化业务专属右键上下文菜单浮窗
    /// </summary>
    public class CustomContextMenuForm : Form
    {
        // 全局单例句柄，确保内存中保持单个快速响应的菜单实例
        private static CustomContextMenuForm? _instance;

        // WebView2 浏览器控件实例
        private readonly WebView2 _webView;

        // WebView2 是否已完成环境初始化并可接收通信
        private bool _isWebReady = false;

        // 缓存的待发送上下文数据包
        private object? _pendingContextData = null;

        // 鼠标移出菜单巡检定时器 (用于光标移回 Excel 单元格且未点击时自动平滑隐藏菜单)
        private readonly System.Windows.Forms.Timer _mouseTrackerTimer;

        // 记录鼠标连续检测处于外部的计数值
        private int _outOfMenuTicks = 0;

        // JSON 序列化选项
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// 私有构造函数：配置无边框置顶菜单窗体与 WebView2
        /// </summary>
        private CustomContextMenuForm()
        {
            // 初始化 WebView2 控件
            _webView = new WebView2();
            // 设置 WebView2 默认背景色为透明，避免加载时闪白并完美贴边消除多余内圈
            _webView.DefaultBackgroundColor = Color.Transparent;

            // 设置无边框模式
            this.FormBorderStyle = FormBorderStyle.None;
            // 不在 Windows 任务栏中显示图标
            this.ShowInTaskbar = false;
            // 窗体始终保持最前端置顶显示
            this.TopMost = true;
            // 设定符合 Office 原生菜单规格的标准尺寸 (宽 250px，高 585px，容纳全部 22 项业务与原生菜单) --硬编码: 菜单初始规格--
            this.Size = new Size(250, 585);
            // 启用手动绝对坐标定位
            this.StartPosition = FormStartPosition.Manual;
            // 设置白色背景
            this.BackColor = Color.White;

            // WebView2 控件完全填充窗体
            _webView.Dock = DockStyle.Fill;
            // 将控件添加至窗体控件集合
            this.Controls.Add(_webView);

            // 初始化移出巡检定时器 (100ms 周期) --硬编码: 巡检定时器周期 100ms--
            _mouseTrackerTimer = new System.Windows.Forms.Timer { Interval = 100 };
            // 订阅定时器巡检事件
            _mouseTrackerTimer.Tick += OnMouseTrackerTick;

            // 订阅窗体加载事件
            this.Load += OnFormLoadAsync;
            // 订阅失去焦点失活事件 (带鼠标区域防误隐藏校验)
            this.Deactivate += OnOverlayDeactivate;
        }

        /// <summary>
        /// 获取当前屏幕或窗体的 DPI 缩放比率 (以标准 96 DPI 为 1.0 基准)
        /// </summary>
        private float GetDpiScale()
        {
            try
            {
                // 获取当前窗体 GDI 句柄以读取横向物理 DPI
                using (Graphics g = this.CreateGraphics())
                {
                    // 96 DPI 对应 100% 缩放
                    float dpi = g.DpiX;
                    // 返回安全计算比率 (最低保底 1.0)
                    return dpi > 0 ? (dpi / 96.0f) : 1.0f;
                }
            }
            catch
            {
                // 异常回退标准比率 1.0
                return 1.0f;
            }
        }

        /// <summary>
        /// 窗体失去焦点时平滑隐藏 (若鼠标在主菜单或二级子菜单内部点击触发失焦则忽略)
        /// </summary>
        private void OnOverlayDeactivate(object? sender, EventArgs e)
        {
            try
            {
                // 获取当前鼠标指针绝对物理屏幕坐标
                Point cur = Cursor.Position;
                // 若鼠标当前停留在主菜单窗体矩形区域内，说明用户正在操作菜单项，不触发隐藏
                if (this.Bounds.Contains(cur))
                {
                    return;
                }
                // 若鼠标当前停留在独立二级子菜单窗体矩形区域内，说明用户正在选择子菜单，不触发隐藏
                if (CustomContextSubmenuForm.Instance.Visible && CustomContextSubmenuForm.Instance.Bounds.Contains(cur))
                {
                    return;
                }
                // 停止鼠标移出巡检定时器
                _mouseTrackerTimer.Stop();
                // 重置计数值
                _outOfMenuTicks = 0;
                // 鼠标在外部区域点击，平滑隐藏主菜单
                this.Hide();
                // 联动隐藏二级子菜单
                if (CustomContextSubmenuForm.Instance.Visible)
                {
                    CustomContextSubmenuForm.Instance.Hide();
                }
            }
            catch { }
        }

        /// <summary>
        /// 巡检鼠标绝对坐标：当光标移出菜单区域到 Excel 表格且未点击时，自动平滑隐藏主菜单与二级子菜单
        /// </summary>
        private void OnMouseTrackerTick(object? sender, EventArgs e)
        {
            try
            {
                // 若主菜单已不处于显示状态，立即停用巡检定时器避免后台无效占用
                if (!this.Visible)
                {
                    _mouseTrackerTimer.Stop();
                    _outOfMenuTicks = 0;
                    return;
                }

                // 获取鼠标屏幕绝对物理坐标
                Point cur = Cursor.Position;

                // 判断是否在主菜单内部
                bool inMain = this.Bounds.Contains(cur);
                // 判断是否在二级子菜单内部 (仅当子菜单可见时有效)
                bool inSub = CustomContextSubmenuForm.Instance.Visible && CustomContextSubmenuForm.Instance.Bounds.Contains(cur);

                if (inMain || inSub)
                {
                    // 鼠标仍在主菜单或二级子菜单内，重置离开计数
                    _outOfMenuTicks = 0;
                }
                else
                {
                    // 鼠标已移出菜单区域 (移回 Excel 单元格或其它窗口)
                    _outOfMenuTicks++;
                    // 离开达到阈值 (连续检测 2 次，约 200ms)，触发自动隐藏
                    if (_outOfMenuTicks >= 2) // --硬编码: 离开判定阈值 2 次 (约 200ms)--
                    {
                        _mouseTrackerTimer.Stop();
                        _outOfMenuTicks = 0;
                        // 平滑隐藏主菜单
                        this.Hide();
                        // 联动隐藏二级子菜单
                        if (CustomContextSubmenuForm.Instance.Visible)
                        {
                            CustomContextSubmenuForm.Instance.Hide();
                        }
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 异步加载 WebView2 环境并导航至 custom_context_menu.html
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 1. 获取系统本地 LocalApplicationData 专有缓存目录，避免污染用户业务配置或网盘同步
                string userDataDir = Tool.GetWebView2UserDataFolder("WebView2_ContextMenu");
                // 2. 异步创建 WebView2 运行时环境
                var env = await CoreWebView2Environment.CreateAsync(null, userDataDir);
                // 3. 确保 WebView2 控件与环境绑定就绪
                await _webView.EnsureCoreWebView2Async(env);

                // 禁用浏览器默认右键菜单，防止出现套娃右键
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                // 禁用底部状态栏
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

                // 注册 Web 消息接收处理委托
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 确定 HTML 文件路径 (优先读取 DLL 运行目录，回退读取 AppDomain 目录)
                string appDir = Tool.GetAppDirectory();
                string htmlPath = Path.Combine(appDir, "Resources", "custom_context_menu.html");
                if (!File.Exists(htmlPath))
                {
                    htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "custom_context_menu.html");
                }

                // 校验文件存在并导航加载
                if (File.Exists(htmlPath))
                {
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
                }
            }
            catch (Exception ex)
            {
                // 记录初始化异常日志
                LogHelper.WriteLog($"CustomContextMenuForm 初始化异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 接收来自前端 Vue 3 的消息并路由调度执行对应动作
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 兼容读取 String 与 Json 两种数据形态
                string rawJson = "";
                try { rawJson = e.TryGetWebMessageAsString(); } catch { }
                if (string.IsNullOrWhiteSpace(rawJson))
                {
                    try { rawJson = e.WebMessageAsJson; } catch { }
                }
                if (string.IsNullOrWhiteSpace(rawJson)) return;

                // 反序列化为 JsonDocument
                using var doc = JsonDocument.Parse(rawJson);
                var root = doc.RootElement;

                // 读取 action 字符串
                string action = root.TryGetProperty("action", out var actionProp) ? (actionProp.GetString() ?? "") : "";

                // 根据前端指令进行路由分支处理
                switch (action)
                {
                    case "menuReady":
                        // 标记前端已准备好
                        _isWebReady = true;
                        // 若有待发送的上下文数据，立即推送
                        if (_pendingContextData != null)
                        {
                            SendContextToWeb(_pendingContextData);
                            _pendingContextData = null;
                        }
                        break;

                    case "closeMenu":
                        // 隐藏当前右键主菜单及独立二级子菜单
                        SafeInvoke(() =>
                        {
                            // 停止鼠标移出巡检定时器
                            _mouseTrackerTimer.Stop();
                            _outOfMenuTicks = 0;
                            this.Hide();
                            // 同步隐藏二级子菜单
                            if (CustomContextSubmenuForm.Instance.Visible)
                            {
                                CustomContextSubmenuForm.Instance.Hide();
                            }
                        });
                        break;

                    case "openSubmenu":
                        // 呼出独立原生二级子菜单窗体，避免主窗体拓宽导致右上方大背景遮挡与最外层矩形阴影
                        SafeInvoke(() =>
                        {
                            try
                            {
                                // 获取菜单项相对于视口顶部的相对距离
                                double itemTop = root.TryGetProperty("top", out var tProp) ? tProp.GetDouble() : 0;
                                float dpiScale = this.GetDpiScale();
                                // 计算屏幕贴合坐标：X 轴为主菜单右边缘内嵌 1 像素消除接缝，Y 轴为主菜单顶边加上项的相对坐标
                                int anchorX = this.Left + this.Width - (int)Math.Ceiling(1 * dpiScale); // --硬编码: 子菜单贴合X偏移--
                                int anchorY = this.Top + (int)Math.Ceiling(itemTop * dpiScale);
                                // 呼出独立轻量子菜单
                                CustomContextSubmenuForm.Instance.ShowSubmenu(this, new Point(anchorX, anchorY));
                            }
                            catch (Exception ex)
                            {
                                LogHelper.WriteLog($"呼出二级子菜单异常: {ex.Message}");
                            }
                        });
                        break;

                    case "closeSubmenu":
                        // 调度二级子菜单延迟隐藏 (支持移入子菜单防抖取消)
                        SafeInvoke(() =>
                        {
                            try
                            {
                                CustomContextSubmenuForm.Instance.ScheduleHide();
                            }
                            catch { }
                        });
                        break;

                    // 核心业务与原生菜单动作集合
                    case "undoAction":
                    case "redoAction":
                    case "cutComponent":
                    case "copyComponent":
                    case "insertCopiedComponent":
                    case "deleteComponent":
                    case "excelCut":
                    case "excelCopy":
                    case "excelPaste":
                    case "excelInsert":
                    case "excelDelete":
                    // Excel 原生隐藏与取消隐藏动作
                    case "excelHide":
                    case "excelUnhide":
                    // Excel 原生系统自动筛选动作
                    case "excelFilterByValueNative":
                    // 成套多选与同柜查件动作
                    case "excelFilterByValue":
                    case "excelClearFilter":
                    case "createCabinet":
                    case "createCabinetNoDetail":
                    case "batchNewCabinet":
                    case "editCabinet":
                    case "cutCabinet":
                    case "copyCabinet":
                    case "insertCopiedCabinet":
                    case "deleteCabinet":
                    case "reorderCabinet":
                    case "parseAndMatch":
                    case "openComponentAttachment":
                    case "searchPricePersonal":
                    case "searchPriceCloud":
                    case "disablePriceSearch":
                    case "openMatchSetting":
                    case "openSmartInput":
                    case "openSummaryAdjustPrice":
                    case "openComponentManage":
                    case "openCabinetAuxCalc":
                    // 恢复此前误删的元器件图纸参数匹配侧边窗口调起分支
                    case "openComponentParamMatch":
                    case "openOnlinePriceSearch":
                    case "switchToNativeMenu":
                    // 元件汇总分布表专属操作
                    case "insertDistributionComponent":
                    case "deleteDistributionComponent":
                    case "filterDistributionRows":
                    case "clearDistributionFilter":
                    case "addToPersonalDb":
                    case "splitComponent":
                    case "deleteCategory":
                        // 收到菜单点击指令：先隐藏主菜单与二级子菜单，后通过 ExcelAsyncUtil.QueueAsMacro 异步执行
                        SafeInvoke(() =>
                        {
                            // 停止巡检定时器
                            _mouseTrackerTimer.Stop();
                            _outOfMenuTicks = 0;
                            this.Hide();
                            // 同步关闭二级子菜单
                            if (CustomContextSubmenuForm.Instance.Visible)
                            {
                                CustomContextSubmenuForm.Instance.Hide();
                            }

                            // 若是切换为原生模式，彻底关闭并释放当前菜单浮窗
                            if (action == "switchToNativeMenu")
                            {
                                try
                                {
                                    this.Close();
                                    _instance = null;
                                }
                                catch { }
                            }

                            // 关键保障：使用 QueueAsMacro 脱离 WebView2 WebMessage 回调上下文，交由 Excel 纯净主线程调度
                            ExcelAsyncUtil.QueueAsMacro(() =>
                            {
                                ExecuteMenuAction(action);
                            });
                        });
                        break;
                }
            }
            catch (Exception ex)
            {
                // 记录消息处理异常日志
                LogHelper.WriteLog($"右键菜单接收消息异常: {ex.Message}");
            }
        }

        // 记录最近一次执行动作的名称与时间戳，防止快速重复调用
        private string _lastActionName = string.Empty;
        private DateTime _lastActionTime = DateTime.MinValue;

        /// <summary>
        /// 执行具体的菜单业务指令 (在 Excel 纯净宏上下文中执行)
        /// </summary>
        private void ExecuteMenuAction(string actionName)
        {
            try
            {
                // 获取当前时间戳
                DateTime now = DateTime.Now;
                // 若 500ms 内重复收到相同指令，直接忽略
                if (string.Equals(_lastActionName, actionName, StringComparison.OrdinalIgnoreCase) && (now - _lastActionTime).TotalMilliseconds < 500)
                {
                    return;
                }
                // 更新最近一次执行状态
                _lastActionName = actionName;
                _lastActionTime = now;

                switch (actionName)
                {
                    case "undoAction":
                        // 调度执行撤销
                        ExcelServices.Undo();
                        break;

                    case "redoAction":
                        // 调度执行重做/还原
                        ExcelServices.Redo();
                        break;

                    case "cutComponent":
                        // 调度执行成套专属剪切元件
                        ExcelServices.CutComponentRow();
                        break;

                    case "copyComponent":
                        // 调度执行成套专属复制元件 (含完整参数与Handle)
                        ExcelServices.CopyComponentRow();
                        break;

                    case "insertCopiedComponent":
                        // 调度执行成套专属插入复制/剪切的元件 (跨柜自动置空CadHandle)
                        ExcelServices.InsertCopiedOrCutComponentRow();
                        break;

                    case "deleteComponent":
                        // 调度执行成套专属删除元件 (防#REF!损坏并自动重排序号)
                        ExcelServices.DeleteComponentRow();
                        break;

                    case "excelCut":
                        // 调度执行 Excel 原生剪切指令
                        try
                        {
                            // 获取 Excel 宿主动态句柄
                            dynamic? dynApp = ExcelDnaUtil.Application;
                            // 优先通过 CommandBars 执行标准 Cut MSO 动作
                            dynApp?.CommandBars?.ExecuteMso("Cut");
                        }
                        catch
                        {
                            // 容错回退：直接对当前选区执行 Cut
                            try { ((dynamic?)ExcelDnaUtil.Application)?.Selection?.Cut(); } catch { }
                        }
                        break;

                    case "excelCopy":
                        // 调度执行 Excel 原生复制指令
                        try
                        {
                            // 获取 Excel 宿主动态句柄
                            dynamic? dynApp = ExcelDnaUtil.Application;
                            // 优先通过 CommandBars 执行标准 Copy MSO 动作
                            dynApp?.CommandBars?.ExecuteMso("Copy");
                        }
                        catch
                        {
                            // 容错回退：直接对当前选区执行 Copy
                            try { ((dynamic?)ExcelDnaUtil.Application)?.Selection?.Copy(); } catch { }
                        }
                        break;

                    case "excelPaste":
                        // 调度执行 Excel 原生粘贴指令
                        try
                        {
                            // 获取 Excel 宿主动态句柄
                            dynamic? dynApp = ExcelDnaUtil.Application;
                            // 优先通过 CommandBars 执行标准 Paste MSO 动作
                            dynApp?.CommandBars?.ExecuteMso("Paste");
                        }
                        catch
                        {
                            // 容错回退：对当前活动工作表执行 Paste
                            try { ((dynamic?)ExcelDnaUtil.Application)?.ActiveSheet?.Paste(); } catch { }
                        }
                        break;

                    case "excelInsert":
                        // 调度执行 Excel 原生“插入...”对话框指令
                        try
                        {
                            // 获取 Excel 宿主动态句柄
                            dynamic? dynApp = ExcelDnaUtil.Application;
                            // 优先通过 CommandBars 调起插入单元格对话框
                            dynApp?.CommandBars?.ExecuteMso("CellsInsertDialog");
                        }
                        catch
                        {
                            // 容错回退：通过内置对话框枚举展示插入窗口
                            try { ((dynamic?)ExcelDnaUtil.Application)?.Dialogs[Microsoft.Office.Interop.Excel.XlBuiltInDialog.xlDialogInsert].Show(); } catch { }
                        }
                        break;

                    case "excelDelete":
                        // 调度执行 Excel 原生“删除...”对话框指令
                        try
                        {
                            // 获取 Excel 宿主动态句柄
                            dynamic? dynApp = ExcelDnaUtil.Application;
                            // 优先通过 CommandBars 调起删除单元格对话框
                            dynApp?.CommandBars?.ExecuteMso("CellsDeleteDialog");
                        }
                        catch
                        {
                            // 容错回退：通过内置对话框枚举展示删除窗口
                            try { ((dynamic?)ExcelDnaUtil.Application)?.Dialogs[Microsoft.Office.Interop.Excel.XlBuiltInDialog.xlDialogEditDelete].Show(); } catch { }
                        }
                        break;

                    case "deleteCategory":
                        // 调度执行删除分类业务 (在项目信息表中智能弹出批量删除，在分类表中执行当前删除)
                        ExcelServices.DeleteCurrentCategory();
                        break;

                    case "excelHide":
                        // 调度业务服务层执行 Excel 原生隐藏 (严格遵守规则 3：Excel 操作统一收敛于 ExcelServices)
                        ExcelServices.ExecuteNativeHide();
                        break;

                    case "excelUnhide":
                        // 调度业务服务层执行 Excel 原生取消隐藏 (严格遵守规则 3：Excel 操作统一收敛于 ExcelServices)
                        ExcelServices.ExecuteNativeUnhide();
                        break;

                    case "excelFilterByValueNative":
                        // 调度业务服务层执行 Excel 原生自动筛选 (严格遵守规则 3：Excel 操作统一收敛于 ExcelServices)
                        ExcelServices.ExecuteNativeFilterBySelection();
                        break;

                    case "excelFilterByValue":
                        // 调度业务层执行“成套多选行当前列筛选与同箱柜高亮标记”
                        ExcelServices.FilterComponentsBySelection();
                        break;

                    case "excelClearFilter":
                        // 调度业务层执行“清除筛选并双轨恢复全貌”
                        ExcelServices.ClearComponentFilter();
                        break;

                    case "createCabinet":
                        // 调度业务层执行“新建箱柜”
                        ExcelServices.CreateNewCabinetFromSelection();
                        break;

                    case "createCabinetNoDetail":
                        // 调度业务层执行“新建无明细箱柜”
                        ExcelServices.CreateNewCabinetNoDetailFromSelection();
                        break;

                    case "batchNewCabinet":
                        // 调度业务层弹出“批建箱柜”窗口
                        ExcelServices.ShowBatchNewCabinetDialog();
                        break;

                    case "editCabinet":
                        // 调度业务层弹出“编辑箱柜信息”窗口
                        ExcelServices.ShowEditCabinetDialog();
                        break;

                    case "cutCabinet":
                        // 调度业务层执行“剪切箱柜”
                        ExcelServices.CutCurrentCabinet();
                        break;

                    case "copyCabinet":
                        // 调度业务层执行“复制箱柜”
                        ExcelServices.CopyCurrentCabinet();
                        break;

                    case "insertCopiedCabinet":
                        // 调度业务层执行“插入复制的箱柜”
                        ExcelServices.InsertCopiedCabinet();
                        break;

                    case "deleteCabinet":
                        // 调度业务层执行“删除箱柜”
                        ExcelServices.DeleteCabinetFromSelection();
                        break;

                    case "reorderCabinet":
                        // 调度业务层弹出“箱柜调序”窗口
                        ExcelServices.ShowCabinetReorderDialog();
                        break;

                    case "parseAndMatch":
                        // 调度业务层执行“识别参数并匹配物料”
                        var result = ExcelServices.ExecuteBatchMatchWithDb(null);
                        if (result != null)
                        {
                            MessageBox.Show(
                                result.Message,
                                result.Success ? "识别与匹配完成" : "提示",
                                MessageBoxButtons.OK,
                                result.Success ? MessageBoxIcon.Information : MessageBoxIcon.Warning
                            );
                        }
                        break;

                    case "openComponentAttachment":
                        // 调度业务层打开当前行元器件的选配配套附件浮窗
                        ExcelServices.ShowComponentAttachmentOverlay();
                        break;

                    case "searchPricePersonal":
                        // 调度业务层在当前活动行弹出物料搜索下拉悬浮窗 (指定数据源为本地个人库)
                        ExcelServices.ShowPriceSearchOverlay("personal");
                        break;

                    case "searchPriceCloud":
                        // 调度业务层在当前活动行弹出物料搜索下拉悬浮窗 (指定数据源为云端公共库)
                        ExcelServices.ShowPriceSearchOverlay("cloud");
                        break;

                    case "disablePriceSearch":
                        // 调度业务层快速关闭物料搜索功能并隐藏下拉悬浮窗
                        ExcelServices.DisablePriceSearch();
                        break;

                    case "openMatchSetting":
                        // 打开“元器件物料匹配与品牌规则设置”窗口
                        ExcelServices.ShowComponentMatchDialog();
                        break;

                    case "openSmartInput":
                        // 打开“智能输入配置”窗口
                        ExcelServices.ShowSmartInputDialog();
                        break;

                    case "openSummaryAdjustPrice":
                        // 打开“汇总调价”窗口
                        ExcelServices.ShowSummaryAdjustPriceDialog();
                        break;

                    case "openComponentManage":
                        // 打开“元器件数据管理”窗口
                        ExcelServices.ShowComponentManageDialog();
                        break;

                    case "openCabinetAuxCalc":
                        // 打开“智能辅材与壳体计算”窗口
                        ExcelServices.ShowCabinetAuxCalcDialog();
                        break;

                    case "openComponentParamMatch":
                        // 打开“元器件图纸参数匹配 (200x800)”侧边浮窗
                        ExcelServices.ShowComponentParamMatchDialog();
                        break;

                    case "openOnlinePriceSearch":
                        // 打开“在线查价与静默回写 (电气天下/天工)”窗口
                        ExcelServices.ShowOnlinePriceSearchDialog();
                        break;

                    case "insertDistributionComponent":
                        // 调度业务层执行【元件汇总分布表】插入新元件行
                        ExcelServices.InsertDistributionComponentRow();
                        break;

                    case "deleteDistributionComponent":
                        // 调度业务层执行【元件汇总分布表】删除选中元件行
                        ExcelServices.DeleteDistributionComponentRows();
                        break;

                    case "filterDistributionRows":
                        // 调度业务层执行【元件汇总分布表】选中行聚焦筛选 (自动隐藏全空箱柜)
                        ExcelServices.FocusFilterDistributionRows();
                        break;

                    case "clearDistributionFilter":
                        // 调度业务层执行【元件汇总分布表】清除筛选恢复全貌
                        ExcelServices.ClearDistributionFilter();
                        break;

                    case "addToPersonalDb":
                        // 调度业务层执行“添加到本地个人库”核对与入库
                        ExcelServices.OpenAddToPersonalDbDialog();
                        break;

                    case "splitComponent":
                        // 调度业务层执行“拆分改型元件”工作台
                        ExcelServices.OpenComponentSplitDialog();
                        break;

                    case "switchToNativeMenu":
                        // 1. 切换为 Excel 原生右键菜单模式并持久化
                        ConfigManager.Instance.SetCustomContextMenuMode(false);
                        // 2. 彻底安全清理 CommandBars 残留
                        ExcelEventManager.RemoveContextMenuControls();
                        // 3. 在 Excel 底部状态栏给出即时提示
                        try
                        {
                            dynamic? app = ExcelDnaUtil.Application;
                            if (app != null) app.StatusBar = "已切换为【Excel 原生右键菜单】模式";
                        }
                        catch { }
                        // 4. 弹出友好提示告知用户已切回原生模式
                        MessageBox.Show(
                            "已成功切换为【Excel 原生右键菜单】！\n\n下次在工作表中右键将直接弹出 Excel 原生菜单。\n如需切回业务专属菜单，请点击 Excel 顶部功能区“右键菜单模式”按钮。",
                            "右键菜单模式切换",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                        break;
                }
            }
            catch (Exception ex)
            {
                // 记录业务执行异常日志
                LogHelper.WriteLog($"执行菜单动作 [{actionName}] 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 将上下文数据包推送到前端 Vue 3
        /// </summary>
        private void SendContextToWeb(object data)
        {
            SafeInvoke(() =>
            {
                try
                {
                    // 校验 WebView 状态
                    if (_webView?.CoreWebView2 != null)
                    {
                        // 序列化上下文数据为 JSON 字符串
                        string json = JsonSerializer.Serialize(data, JsonOptions);
                        // 发送 Web 消息至前端 (使用 PostWebMessageAsString)
                        _webView.CoreWebView2.PostWebMessageAsString(json);
                    }
                }
                catch (Exception ex)
                {
                    // 记录数据推送异常
                    LogHelper.WriteLog($"推送上下文至右键菜单异常: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 跨线程安全调度
        /// </summary>
        private void SafeInvoke(Action action)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(action);
            }
            else
            {
                action();
            }
        }

        /// <summary>
        /// 在屏幕指定坐标处弹出业务专属右键菜单
        /// </summary>
        public static void ShowMenu(Point screenPos, string sheetName, string cellAddress, int row, int column, bool isAboveFirstDet)
        {
            try
            {
                // 确保实例已创建
                if (_instance == null || _instance.IsDisposed)
                {
                    _instance = new CustomContextMenuForm();
                }

                // 判定当前工作表是否为“元件汇总表”
                bool isSummarySheet = string.Equals(sheetName?.Trim(), ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase);

                // 判定当前工作表是否为“元件汇总分布表”
                bool isDistributionSheet = string.Equals(sheetName?.Trim(), ExcelServices.DistributionSheetName, StringComparison.OrdinalIgnoreCase);

                // 准备上下文传输数据
                var contextData = new
                {
                    action = "initContext",
                    sheetName = sheetName,
                    cellAddress = cellAddress,
                    row = row,
                    column = column,
                    isAboveFirstDet = isAboveFirstDet,
                    // 标记当前活动表是否为元件汇总表
                    isSummarySheet = isSummarySheet,
                    // 标记当前活动表是否为元件汇总分布表
                    isDistributionSheet = isDistributionSheet,
                    canUndo = ExcelServices.CanUndo,
                    canRedo = ExcelServices.CanRedo,
                    undoName = ExcelServices.CurrentUndoName ?? "撤销",
                    redoName = ExcelServices.CurrentRedoName ?? "还原"
                };

                // 获取当前屏幕可用工作区域
                Screen currentScreen = Screen.FromPoint(screenPos);
                Rectangle workArea = currentScreen.WorkingArea;

                // 每次显示前重置为标准尺寸：
                // 1. 元件汇总分布表 (专属6项)：165px
                // 2. 元件汇总表 (全局大单表)：215px
                // 3. 分类表顶部箱柜汇总区 (isAboveFirstDet == true)：385px
                // 4. 分类表明细元器件插槽区 (isAboveFirstDet == false)：445px (隐藏禁用的新建箱柜后紧凑自适应)
                int standardHeight = isDistributionSheet ? 175 : (isSummarySheet ? 215 : (isAboveFirstDet ? 385 : 445)); // --硬编码: 右键菜单标准高度 (分布表 175px, 汇总表 215px, 顶部箱柜 385px, 明细表 445px)--
                // 若工作区高度受限 (如低分辨率笔记本屏幕)，自适应贴合可用工作区
                int targetHeight = Math.Min(standardHeight, workArea.Height - 10);
                // 获取当前窗体 DPI 缩放比率并计算标准物理宽度
                float dpiScale = _instance.GetDpiScale();
                int standardWidth = (int)Math.Ceiling(250 * dpiScale); // --硬编码: 初始标准宽度 250px--
                // 设置窗口实际物理尺寸
                _instance.Size = new Size(standardWidth, targetHeight);

                // 初始坐标偏移 2 像素防止挡住鼠标
                int x = screenPos.X + 2;
                int y = screenPos.Y + 2;

                // 防止右侧超出屏幕边缘
                if (x + _instance.Width > workArea.Right)
                {
                    x = Math.Max(workArea.Left, screenPos.X - _instance.Width - 2);
                }

                // 防止底部超出屏幕边缘
                if (y + _instance.Height > workArea.Bottom)
                {
                    y = Math.Max(workArea.Top, screenPos.Y - _instance.Height - 2);
                }

                // 设置窗体绝对物理坐标
                _instance.Location = new Point(x, y);

                // 若前端已就绪，立即推送上下文数据
                if (_instance._isWebReady)
                {
                    _instance.SendContextToWeb(contextData);
                }
                else
                {
                    // 暂存待页面就绪后推送
                    _instance._pendingContextData = contextData;
                }

                // 每次弹出主菜单前确保隐藏历史残留的二级子菜单
                if (CustomContextSubmenuForm.Instance.Visible)
                {
                    CustomContextSubmenuForm.Instance.Hide();
                }

                // 显示窗口并置顶
                if (!_instance.Visible)
                {
                    _instance.Show();
                }
                _instance.BringToFront();
                _instance.Activate();

                // 每次显示前重置移出离开计数并启动巡检定时器 (提供初始 200ms 防误触缓冲期)
                _instance._outOfMenuTicks = -2; // --硬编码: 初始唤醒缓冲计数--
                _instance._mouseTrackerTimer.Start();
            }
            catch (Exception ex)
            {
                // 记录弹窗异常
                LogHelper.WriteLog($"弹出自定义右键菜单异常: {ex.Message}");
            }
        }
    }
}
