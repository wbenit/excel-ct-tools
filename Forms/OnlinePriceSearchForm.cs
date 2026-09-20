using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的在线查价与静默回写无边框置顶宿主窗口
    /// </summary>
    public class OnlinePriceSearchForm : Form
    {
        // 声明 WebView2 浏览器控件句柄
        private readonly WebView2 _webView;

        // 声明在线查价控制器
        private readonly OnlinePriceSearchController _controller;

        // 异步查价取消令牌源
        private CancellationTokenSource? _cts;

        // Windows 原生 API: 用于无边框窗体拖拽移动
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 系统消息常量: 标题栏左键按下 --硬编码: Win32消息代码--
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 通用 JSON 序列化配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// 构造函数: 初始化控制器与无边框窗体
        /// </summary>
        public OnlinePriceSearchForm()
        {
            // 实例化后端控制器
            _controller = new OnlinePriceSearchController();

            // 实例化 WebView2
            _webView = new WebView2();

            // 配置窗体几何尺寸与样式
            InitializeFormProperties();

            // 挂载 WebView2 并初始化导航
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体几何外观 (940x680 像素，无边框置顶)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 窗体文本
            this.Text = "在线查价与一键静默回写 (电气天下 / 天工矩阵)";

            // 设定宽 940，高 680 --硬编码: 窗体规格尺寸--
            this.ClientSize = new Size(940, 680);

            // 屏幕正中央启动
            this.StartPosition = FormStartPosition.CenterScreen;

            // 无边框模式
            this.FormBorderStyle = FormBorderStyle.None;

            // 保持最前置顶显示
            this.TopMost = true;

            // 禁用原生最大化/最小化按钮
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 背景色设为纯白
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 挂载并初始化 WebView2 控件
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 充满窗体工作区
            _webView.Dock = DockStyle.Fill;

            // 默认透明背景，防止白屏闪烁
            _webView.DefaultBackgroundColor = Color.Transparent;

            // 添加至窗体控件树
            this.Controls.Add(_webView);

            // 挂接窗体加载事件
            this.Load += OnFormLoadAsync;

            // 挂接窗体关闭事件，安全释放
            this.FormClosing += OnFormClosingCleanup;
            this.FormClosed += OnFormClosedDispose;
        }

        /// <summary>
        /// 异步初始化 WebView2 运行时环境并导航至 HTML 页面
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 获取插件专属 AppData 缓存目录
                string userDataDir = Path.Combine(Tool.GetAppDataDirectory(), "WebView2_OnlinePrice");
                // 异步创建核心环境
                var env = await CoreWebView2Environment.CreateAsync(null, userDataDir);
                // 绑定到控件
                await _webView.EnsureCoreWebView2Async(env);

                // 禁用浏览器默认右键菜单
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                // 禁用底部状态栏
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

                // 注册 Web 消息接收处理监听
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 探测定位 HTML 资源绝对路径
                string appDir = Tool.GetAppDirectory();
                string[] possiblePaths = new[]
                {
                    Path.Combine(appDir, "Resources", "online_price_search.html"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "online_price_search.html"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "publish", "Resources", "online_price_search.html")
                };

                string htmlPath = "";
                // 遍历寻找存在的路径
                foreach (var p in possiblePaths)
                {
                    if (File.Exists(p))
                    {
                        htmlPath = p;
                        break;
                    }
                }

                // 校验存在并导航
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
                }
                else
                {
                    MessageBox.Show($"未找到在线查价界面文件: {htmlPath}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                // 记录错误日志
                LogHelper.WriteLog($"OnlinePriceSearchForm 初始化异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 响应来自前端 Vue 3 的 IPC 交互指令
        /// </summary>
        private async void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 读取消息字符串
                string messageJson = "";
                try { messageJson = e.TryGetWebMessageAsString(); } catch { }
                if (string.IsNullOrEmpty(messageJson)) messageJson = e.WebMessageAsJson;
                if (string.IsNullOrEmpty(messageJson)) return;

                // 解析 JSON
                using var doc = JsonDocument.Parse(messageJson);
                var root = doc.RootElement;
                string action = root.TryGetProperty("action", out var actProp) ? actProp.GetString() ?? "" : "";

                switch (action)
                {
                    // 1. 获取当前配置与选区探测数据
                    case "initData":
                        var cfg = _controller.LoadConfig();
                        var selectionInfo = _controller.DetectSelection();
                        PostWebMessageSafe(new
                        {
                            action = "renderInitData",
                            config = cfg,
                            selection = selectionInfo
                        });
                        break;

                    // 2. 重新检测当前 Excel 选区
                    case "refreshSelection":
                        var newSelection = _controller.DetectSelection();
                        PostWebMessageSafe(new
                        {
                            action = "renderSelection",
                            selection = newSelection
                        });
                        break;

                    // 3. 保存配置
                    case "saveConfig":
                        if (root.TryGetProperty("config", out var cfgElem))
                        {
                            var saveObj = JsonSerializer.Deserialize<OnlinePriceConfig>(cfgElem.GetRawText(), JsonOptions);
                            if (saveObj != null)
                            {
                                bool ok = _controller.SaveConfig(saveObj);
                                PostWebMessageSafe(new { action = "saveConfigResult", success = ok });
                            }
                        }
                        break;

                    // 4. 单项精确搜价测试
                    case "searchSingle":
                        string model = root.TryGetProperty("model", out var mProp) ? mProp.GetString() ?? "" : "";
                        int platformVal = root.TryGetProperty("platform", out var pProp) ? pProp.GetInt32() : 0;
                        string brand = root.TryGetProperty("brand", out var bProp) ? bProp.GetString() ?? "全部" : "全部";

                        var singleResult = await _controller.SearchSingleAsync(model, (OnlinePricePlatform)platformVal, brand);
                        PostWebMessageSafe(new
                        {
                            action = "singleSearchResult",
                            result = singleResult
                        });
                        break;

                    // 5. 框选一键静默批量查价并回写 Excel
                    case "startBatchSearch":
                        if (root.TryGetProperty("config", out var batchCfgElem))
                        {
                            // 解析前端提交的最新查价与回填配置
                            var batchCfg = JsonSerializer.Deserialize<OnlinePriceConfig>(batchCfgElem.GetRawText(), JsonOptions) ?? new OnlinePriceConfig();

                            // 步骤 1: 必须在 UI/主线程安全同步读取选区数据 (彻底杜绝跨线程访问 COM 引起的死锁)
                            List<OnlinePriceItemDto> items;
                            try
                            {
                                // 同步提取选区元器件数据
                                items = ExcelServices.GetSelectedItemsForSearch(batchCfg.TrimParentheses);
                            }
                            catch (Exception ex)
                            {
                                // 捕获并通知前端
                                PostWebMessageSafe(new
                                {
                                    action = "batchComplete",
                                    success = false,
                                    total = 0,
                                    updatedCount = 0,
                                    message = $"读取 Excel 选区异常: {ex.Message}",
                                    items = new List<OnlinePriceItemDto>()
                                });
                                break;
                            }

                            // 校验提取结果是否为空
                            if (items == null || items.Count == 0)
                            {
                                // 提示用户选区无有效型号
                                PostWebMessageSafe(new
                                {
                                    action = "batchComplete",
                                    success = false,
                                    total = 0,
                                    updatedCount = 0,
                                    message = "当前选区未识别到有效元器件，请在 Excel 中框选包含规格型号的单元格",
                                    items = new List<OnlinePriceItemDto>()
                                });
                                break;
                            }

                            // 重置并初始化操作取消令牌
                            _cts?.Cancel();
                            _cts = new CancellationTokenSource();
                            var token = _cts.Token;

                            // 进度上报委托，用于向前端实时推流
                            Action<OnlinePriceProgressDto> progressHandler = (p) =>
                            {
                                PostWebMessageSafe(new
                                {
                                    action = "progressUpdate",
                                    progress = p
                                });
                            };

                            // 步骤 2: 纯粹在后台线程池执行并发 HTTP 远端抓取 (不触碰任何 Excel COM 接口)
                            _ = Task.Run(async () =>
                            {
                                try
                                {
                                    // 调度控制器并发拉取远端价格
                                    await _controller.ExecuteRemoteBatchSearchAsync(items, batchCfg, progressHandler, token).ConfigureAwait(false);

                                    // 若检测到操作已被用户中止
                                    if (token.IsCancellationRequested)
                                    {
                                        PostWebMessageSafe(new
                                        {
                                            action = "batchComplete",
                                            success = false,
                                            total = items.Count,
                                            updatedCount = 0,
                                            message = "操作已被用户中止",
                                            items = items
                                        });
                                        return;
                                    }

                                    // 通知前端准备执行单元格批量回写
                                    progressHandler(new OnlinePriceProgressDto
                                    {
                                        Current = items.Count,
                                        Total = items.Count,
                                        Percent = 100,
                                        LogText = "远端价格获取完毕，正在静默批量回写 Excel 数据矩阵..."
                                    });

                                    // 步骤 3: 回到 UI/主线程安全执行 Excel 单元格二维矩阵批量回写
                                    this.Invoke(new Action(() =>
                                    {
                                        try
                                        {
                                            // 主线程执行二维矩阵回写
                                            var writeRes = _controller.WriteBackToExcel(items, batchCfg);
                                            // 格式化完成提示文本
                                            string finalMsg = writeRes.success ?
                                                $"完成！{writeRes.message}" :
                                                $"查价完成，但回写 Excel 时提示: {writeRes.message}";

                                            // 输出最终完成日志
                                            progressHandler(new OnlinePriceProgressDto
                                            {
                                                Current = items.Count,
                                                Total = items.Count,
                                                Percent = 100,
                                                LogText = finalMsg
                                            });

                                            // 通知前端批量任务圆满完成
                                            PostWebMessageSafe(new
                                            {
                                                action = "batchComplete",
                                                success = writeRes.success,
                                                total = items.Count,
                                                updatedCount = writeRes.updatedCount,
                                                message = finalMsg,
                                                items = items
                                            });
                                        }
                                        catch (Exception writeEx)
                                        {
                                            // 捕获回写过程异常
                                            PostWebMessageSafe(new
                                            {
                                                action = "batchComplete",
                                                success = false,
                                                total = items.Count,
                                                updatedCount = 0,
                                                message = $"回写 Excel 异常: {writeEx.Message}",
                                                items = items
                                            });
                                        }
                                    }));
                                }
                                catch (Exception ex)
                                {
                                    // 捕获批量查价主流程异常
                                    PostWebMessageSafe(new
                                    {
                                        action = "batchComplete",
                                        success = false,
                                        total = items.Count,
                                        updatedCount = 0,
                                        message = $"批量查价网络异常: {ex.Message}",
                                        items = items
                                    });
                                }
                            });
                        }
                        break;

                    // 6. 中止查价
                    case "cancelBatchSearch":
                        _cts?.Cancel();
                        PostWebMessageSafe(new { action = "batchCancelled" });
                        break;

                    // 7. 窗口拖拽移动
                    case "dragWindow":
                        ReleaseCapture();
                        SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        break;

                    // 8. 窗口最小化
                    case "minimizeWindow":
                        this.WindowState = FormWindowState.Minimized;
                        break;

                    // 9. 关闭窗口 (落实最佳实践: BeginInvoke 异步脱离 IPC 消息栈，防死锁)
                    case "closeWindow":
                        this.BeginInvoke(new Action(() => this.Close()));
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"OnWebMessageReceived 消息处理异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 安全向前端页面投递 JSON 消息
        /// </summary>
        private void PostWebMessageSafe(object payload)
        {
            try
            {
                // 控件已销毁或核心未就绪则忽略
                if (this.IsDisposed || _webView.CoreWebView2 == null) return;

                // 序列化消息体
                string json = JsonSerializer.Serialize(payload, JsonOptions);

                // 若在后台线程则切回 UI 线程分发
                if (this.InvokeRequired)
                {
                    this.BeginInvoke(new Action(() =>
                    {
                        try { _webView.CoreWebView2?.PostWebMessageAsString(json); } catch { }
                    }));
                }
                else
                {
                    _webView.CoreWebView2.PostWebMessageAsString(json);
                }
            }
            catch { }
        }

        /// <summary>
        /// 窗体即将关闭时取消任务与解绑监听
        /// </summary>
        private void OnFormClosingCleanup(object? sender, FormClosingEventArgs e)
        {
            try
            {
                // 取消正在进行的后台网络任务
                _cts?.Cancel();
                // 解绑 WebMessage 监听
                if (_webView?.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
                }
            }
            catch { }
        }

        /// <summary>
        /// 窗体完全脱离屏幕后安全销毁 WebView2 (防死锁)
        /// </summary>
        private void OnFormClosedDispose(object? sender, FormClosedEventArgs e)
        {
            try
            {
                // 彻底释放 WebView2 实例句柄
                _webView?.Dispose();
            }
            catch { }
        }
    }
}
