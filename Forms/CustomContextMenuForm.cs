using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo;
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

        // JSON 序列化选项
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// 启用 Windows 系统原生右键菜单阴影效果 (CS_DROPSHADOW)
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                // 系统原生下拉阴影样式常数 --硬编码--
                const int CS_DROPSHADOW = 0x00020000;
                // 获取基础创建参数
                CreateParams cp = base.CreateParams;
                // 附加阴影样式位，提供 Office 原生柔和外阴影
                cp.ClassStyle |= CS_DROPSHADOW;
                // 返回配置后的创建参数
                return cp;
            }
        }

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

            // 订阅窗体加载事件
            this.Load += OnFormLoadAsync;
            // 订阅失去焦点失活事件 (带鼠标区域防误隐藏校验)
            this.Deactivate += OnOverlayDeactivate;
        }

        /// <summary>
        /// 窗体失去焦点时平滑隐藏 (若鼠标在菜单内部点击触发失焦则忽略)
        /// </summary>
        private void OnOverlayDeactivate(object? sender, EventArgs e)
        {
            try
            {
                // 若鼠标当前仍停留在菜单窗体矩形区域内，说明用户正在点击菜单项，不触发隐藏
                if (this.Bounds.Contains(Cursor.Position))
                {
                    return;
                }
                // 鼠标在外部点击，平滑隐藏菜单
                this.Hide();
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
                // 获取用户专属 WebView2 缓存目录
                string userDataDir = Path.Combine(Tool.GetAppDataDirectory(), "WebView2_ContextMenu");
                // 异步创建 WebView2 运行时环境
                var env = await CoreWebView2Environment.CreateAsync(null, userDataDir);
                // 确保 WebView2 控件与环境绑定就绪
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
                        // 隐藏当前右键菜单
                        SafeInvoke(this.Hide);
                        break;

                    // 核心业务与原生菜单动作集合
                    case "undoAction":
                    case "redoAction":
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
                    case "parseAndMatch":
                    case "openComponentAttachment":
                    case "openMatchSetting":
                    case "openSmartInput":
                    case "openSummaryAdjustPrice":
                    case "openComponentManage":
                    case "openCabinetAuxCalc":
                    case "openComponentParamMatch":
                    case "switchToNativeMenu":
                        // 收到菜单点击指令：先隐藏菜单并关闭浮窗，后通过 ExcelAsyncUtil.QueueAsMacro 异步执行
                        SafeInvoke(() =>
                        {
                            this.Hide();

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

                // 准备上下文传输数据
                var contextData = new
                {
                    action = "initContext",
                    sheetName = sheetName,
                    cellAddress = cellAddress,
                    row = row,
                    column = column,
                    isAboveFirstDet = isAboveFirstDet,
                    canUndo = ExcelServices.CanUndo,
                    canRedo = ExcelServices.CanRedo,
                    undoName = ExcelServices.CurrentUndoName ?? "撤销",
                    redoName = ExcelServices.CurrentRedoName ?? "还原"
                };

                // 获取当前屏幕可用工作区域
                Screen currentScreen = Screen.FromPoint(screenPos);
                Rectangle workArea = currentScreen.WorkingArea;

                // 每次显示前重置为标准尺寸，杜绝历史状态残留与 DPI 缩放萎缩
                int standardHeight = 610; // --硬编码: 右键菜单标准高度--
                // 若工作区高度受限 (如低分辨率笔记本屏幕)，自适应贴合可用工作区
                int targetHeight = Math.Min(standardHeight, workArea.Height - 10);
                _instance.Size = new Size(250, targetHeight);

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

                // 显示窗口并置顶
                if (!_instance.Visible)
                {
                    _instance.Show();
                }
                _instance.BringToFront();
                _instance.Activate();
            }
            catch (Exception ex)
            {
                // 记录弹窗异常
                LogHelper.WriteLog($"弹出自定义右键菜单异常: {ex.Message}");
            }
        }
    }
}
