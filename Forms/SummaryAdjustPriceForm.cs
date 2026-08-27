using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Controllers;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“汇总调价”无边框模态宿主窗口
    /// </summary>
    public class SummaryAdjustPriceForm : Form
    {
        // 声明 WebView2 浏览器主控件
        private readonly WebView2 _webView;

        // 声明汇总调价 WebAPI 风格控制器
        private readonly SummaryAdjustPriceController _controller;

        // 导入 Windows 原生 user32.dll 接口用于支持拖拽窗体
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 原生消息接口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 常量定义: 标题栏鼠标左键按下消息
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 通用 JSON 序列化配置结构
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化控制器与 WebView2 控件属性
        /// </summary>
        public SummaryAdjustPriceForm()
        {
            // 实例化汇总调价控制器
            _controller = new SummaryAdjustPriceController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 设置 Form 窗体尺寸与显示几何外观
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸 (720x620 像素)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题文本
            this.Text = "汇总调价";

            // 依据设计布局设定尺寸为 720x620 像素
            this.ClientSize = new Size(720, 620);

            // 设置屏幕中央弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 设为无边框样式
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用 WinForm 原生最大化
            this.MaximizeBox = false;

            // 启用最小化
            this.MinimizeBox = false;

            // 设置窗口背景填充色
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件并注册加载回调
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 控件满框充满 Form
            _webView.Dock = DockStyle.Fill;

            // 挂载至 Controls 控件集
            this.Controls.Add(_webView);

            // 绑定 Form Load 异步加载监听
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载触发的异步初始化逻辑
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 窗体已被销毁则退出
                if (this.IsDisposed || this.Disposing) return;

                // 设置本地 AppData 中的 WebView2 缓存生成路径，防止 Excel 进程在 Program Files 下被拒绝访问
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAddInDemo",
                    "WebView2Data"
                );

                // 创建缓存文件夹
                Directory.CreateDirectory(userDataFolder);

                // 异步创建 WebView2 核心环境并指定专属用户数据目录
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 判断窗体有效性
                if (this.IsDisposed || this.Disposing) return;

                // 初始化 CoreWebView2 核心对象
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 注册前端 postMessage 通信监听
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 寻找目标 HTML 资源文件路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string appDir = Tool.GetAppDirectory();

                // 配置多级备选路径集 --避免调试与打包执行环境差异--
                string[] candidatePaths = new string[]
                {
                    Path.Combine(baseDir, "Resources", "summary_adjust_price.html"),
                    Path.Combine(baseDir, "..", "Resources", "summary_adjust_price.html"),
                    Path.Combine(baseDir, "publish", "Resources", "summary_adjust_price.html"),
                    Path.Combine(appDir, "Resources", "summary_adjust_price.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "summary_adjust_price.html")
                };

                // 循环查找首个存在的目标文件
                string htmlPath = string.Empty;
                foreach (string candidate in candidatePaths)
                {
                    if (File.Exists(candidate))
                    {
                        htmlPath = candidate;
                        break;
                    }
                }

                // 导航至前端 Vue 3 页面
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    _webView.Source = new Uri(htmlPath);
                }
                else
                {
                    MessageBox.Show($"未找到汇总调价界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化 WebView2 发生异常: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 安全跨线程调度 UI 动作，防止在句柄未创建或窗体已被释放时调用抛出异常
        /// </summary>
        private void SafeInvoke(Action action)
        {
            if (this.IsDisposed || this.Disposing) return;
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
        /// 处理来自前端 Vue 3 的 postMessage 请求
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 读取原始消息文本（优先读取 String，备选读取 JSON）
                string jsonString = string.Empty;
                try
                {
                    jsonString = e.TryGetWebMessageAsString();
                }
                catch { }

                if (string.IsNullOrEmpty(jsonString))
                {
                    try { jsonString = e.WebMessageAsJson; } catch { }
                }

                if (string.IsNullOrEmpty(jsonString)) return;

                // 解析 JSON 报文为 JsonDocument
                using var doc = JsonDocument.Parse(jsonString);
                var root = doc.RootElement;

                // 读取动作 action 指令
                if (!root.TryGetProperty("action", out var actionElement)) return;
                string action = actionElement.GetString() ?? string.Empty;

                // 响应无边框窗体拖拽指令
                if (action == "dragWindow")
                {
                    SafeInvoke(() =>
                    {
                        // 释放当前鼠标捕获句柄
                        ReleaseCapture();
                        // 发送 WM_NCLBUTTONDOWN (0xA1) 消息触发原生无边框窗口拖拽
                        SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                    });
                }
                // 响应最小化窗口指令
                else if (action == "minimize")
                {
                    SafeInvoke(() => this.WindowState = FormWindowState.Minimized);
                }
                // 响应关闭窗口指令
                else if (action == "close")
                {
                    SafeInvoke(() => this.Close());
                }
                // 响应获取分类列表指令
                else if (action == "getCategories")
                {
                    string resultJson = _controller.GetCategories();
                    // 跨线程安全回发分类列表
                    PostWebMessageAsStringSafe(resultJson);
                }
                // 响应生成元件汇总表指令
                else if (action == "generateSummary")
                {
                    var req = JsonSerializer.Deserialize<GenerateSummaryRequest>(jsonString, JsonOptions);
                    if (req != null)
                    {
                        string resultJson = _controller.GenerateSummary(req);
                        // 跨线程安全回发生成结果
                        PostWebMessageAsStringSafe(resultJson);
                    }
                }
                // 响应调整窗体尺寸指令 (例如切换为图二紧凑编辑条时动态收缩窗口)
                else if (action == "resizeWindow")
                {
                    // 读取目标宽度，默认 720
                    int width = root.TryGetProperty("width", out var wElem) ? wElem.GetInt32() : 720;
                    // 读取目标高度，默认 620
                    int height = root.TryGetProperty("height", out var hElem) ? hElem.GetInt32() : 620;

                    // 调度至主线程执行窗口几何尺寸更新
                    SafeInvoke(() =>
                    {
                        // 检查当前尺寸是否有变化
                        if (this.ClientSize.Width != width || this.ClientSize.Height != height)
                        {
                            // 动态调整窗体工作区尺寸
                            this.ClientSize = new Size(width, height);
                        }
                    });
                }
                // 响应一键更新指令 (预留接口)
                else if (action == "updateFromSummary")
                {
                    // 调用控制器执行一键更新业务
                    string resultJson = _controller.UpdateFromSummary(jsonString);
                    // 跨线程安全向前端回发更新结果
                    PostWebMessageAsStringSafe(resultJson);
                }
                // 响应切换列隐藏状态指令
                else if (action == "toggleColumnsVisibility")
                {
                    // 读取目标列区域
                    string targetRange = root.TryGetProperty("targetRange", out var trElem) ? (trElem.GetString() ?? "") : "";
                    // 读取隐藏标志
                    bool hidden = root.TryGetProperty("hidden", out var hElem) && hElem.GetBoolean();

                    // 调用控制器执行切换
                    string resultJson = _controller.ToggleColumnsVisibility(targetRange, hidden);
                    // 跨线程安全回发执行结果
                    PostWebMessageAsStringSafe(resultJson);
                }
                // 响应获取列隐藏状态指令
                else if (action == "getColumnsHiddenStatus")
                {
                    // 调用控制器读取列隐藏状态
                    string resultJson = _controller.GetColumnsHiddenStatus();
                    // 跨线程安全回发状态报文
                    PostWebMessageAsStringSafe(resultJson);
                }
            }
            catch (Exception ex)
            {
                // 记录异常并回传给前端
                LogHelper.WriteLog($"处理汇总调价前端消息异常: {ex.Message}");
                var errorObj = new
                {
                    action = "onError",
                    message = $"处理消息异常: {ex.Message}"
                };
                // 跨线程安全回发错误消息
                PostWebMessageAsStringSafe(JsonSerializer.Serialize(errorObj));
            }
        }

        /// <summary>
        /// 跨线程安全向 WebView2 发送字符串消息
        /// </summary>
        private void PostWebMessageAsStringSafe(string text)
        {
            // 在 UI 线程中调度发送
            SafeInvoke(() =>
            {
                // 确保 WebView2 控件及其内核有效
                if (!this.IsDisposed && _webView?.CoreWebView2 != null)
                {
                    // 向前端页面发送文本消息
                    _webView.CoreWebView2.PostWebMessageAsString(text);
                }
            });
        }
    }
}
