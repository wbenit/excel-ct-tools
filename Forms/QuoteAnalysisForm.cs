using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 + Apache ECharts 5 的工程报价全景分析大屏窗体
    /// 支持深邃科技与极简白双模视觉、穿透下钻 Excel 与交互式调价测算
    /// </summary>
    public class QuoteAnalysisForm : Form
    {
        // WebView2 核心浏览器控件
        private readonly WebView2 _webView;

        // Windows 原生 API 导入：支持标题栏拖拽移动
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int VK_LBUTTON = 0x01;
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 通用 JSON 序列化配置 (驼峰命名)
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// 构造函数：初始化窗体属性与 WebView2 核心控件
        /// </summary>
        public QuoteAnalysisForm()
        {
            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 初始化窗体基本外观与尺寸
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体尺寸与外观 (1260x820 像素，宽屏自适应大屏体验)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 窗体标题 --硬编码: 窗体标题--
            this.Text = "工程报价全景分析与可视化大屏";
            // 默认大屏视口尺寸
            this.ClientSize = new Size(1280, 820);
            // 限制最小尺寸以保证图表排版美观 (防止被过度压缩)
            this.MinimumSize = new Size(980, 680);
            // 屏幕居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;
            // 采用标准带缩放控制手柄的现代窗体边框
            this.FormBorderStyle = FormBorderStyle.Sizable;
            // 启用最大化与最小化按钮
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            // 默认背景色与科技暗色/浅色呼应
            this.BackColor = Color.FromArgb(11, 19, 32);
        }

        /// <summary>
        /// 挂载 WebView2 控件并注册加载事件
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 铺满整个窗体客户区
            _webView.Dock = DockStyle.Fill;
            // 添加至控件集合
            this.Controls.Add(_webView);
            // 绑定异步加载事件
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载触发异步初始化 WebView2 内核及加载 HTML 资源
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 1. 获取系统 LocalApplicationData 专有隔离缓存目录 (避免污染业务配置与网盘)
                string webViewCacheDir = Tool.GetWebView2UserDataFolder("WebView2_QuoteAnalysis");

                // 2. 异步创建 WebView2 运行环境
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, webViewCacheDir);

                // 3. 确保 CoreWebView2 就绪
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 4. 注册与前端 Vue 3 的双向通信管道
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 5. 检索 HTML 页面资源物理路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string appDir = Tool.GetAppDirectory();

                // 配置多级备用检索路径集 --硬编码: 资源文件名--
                string[] candidatePaths = new string[]
                {
                    Path.Combine(appDir, "Resources", "quote_analysis.html"),
                    Path.Combine(appDir, "quote_analysis.html"),
                    Path.Combine(baseDir, "Resources", "quote_analysis.html"),
                    Path.Combine(baseDir, "quote_analysis.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "quote_analysis.html")
                };

                string htmlPath = string.Empty;
                foreach (string candidate in candidatePaths)
                {
                    if (File.Exists(candidate))
                    {
                        htmlPath = candidate;
                        break;
                    }
                }

                // 6. 导航至大屏前端页面
                if (!string.IsNullOrWhiteSpace(htmlPath))
                {
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
                }
                else
                {
                    MessageBox.Show("未找到报价分析大屏界面资源文件: quote_analysis.html", "资源缺失", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"QuoteAnalysisForm 初始化异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 核心消息分发路由器：响应前端 Vue 3 postMessage 指令
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 双轨解析前端 JSON 报文
                string jsonString = e.WebMessageAsJson;
                using var jsonDoc = JsonDocument.Parse(jsonString);
                var root = jsonDoc.RootElement;

                // 提取 action 动作标识
                if (!root.TryGetProperty("action", out var actionProp)) return;
                string action = actionProp.GetString() ?? string.Empty;

                switch (action)
                {
                    // 1. 请求加载或刷新报价全景数据
                    case "getQuoteData":
                    case "refreshData":
                        LoadAndPushQuoteData();
                        break;

                    // 2. 点击图表或箱柜卡片，下钻穿透定位至 Excel 单元格
                    case "navigateToCabinet":
                        string sheetName = string.Empty;
                        int rowNumber = 0;
                        if (root.TryGetProperty("sheetName", out var sProp)) sheetName = sProp.GetString() ?? "";
                        if (root.TryGetProperty("rowNumber", out var rProp)) rowNumber = rProp.GetInt32();

                        if (!string.IsNullOrEmpty(sheetName) && rowNumber > 0)
                        {
                            ExcelServices.NavigateToCabinetInExcel(sheetName, rowNumber);
                        }
                        break;

                    // 3. 拖拽标题栏移动窗体 (带鼠标物理按压状态防护)
                    case "dragWindow":
                        if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0)
                        {
                            ReleaseCapture();
                            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        }
                        break;

                    // 4. 窗口最大化/还原切换
                    case "toggleMaximize":
                        if (this.WindowState == FormWindowState.Maximized)
                        {
                            this.WindowState = FormWindowState.Normal;
                        }
                        else
                        {
                            this.WindowState = FormWindowState.Maximized;
                        }
                        break;

                    // 5. 窗口最小化
                    case "minimizeWindow":
                        this.WindowState = FormWindowState.Minimized;
                        break;

                    // 6. 关闭大屏窗口
                    case "closeWindow":
                        this.Close();
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"QuoteAnalysisForm 消息处理异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 从 ExcelServices 提取全量业务数据并推送至前端 Vue 3
        /// </summary>
        private void LoadAndPushQuoteData()
        {
            try
            {
                // 从业务服务层批量抽取数据
                var quoteData = ExcelServices.GetQuoteAnalysisData();

                // 序列化为 JSON 字符串
                string payload = JsonSerializer.Serialize(new
                {
                    action = "quoteDataLoaded",
                    data = quoteData
                }, JsonOptions);

                // 线程安全回传给前端
                PostWebMessageSafe(payload);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[报价分析大屏] 推送数据异常: {ex.Message}");
                PostWebMessageSafe(JsonSerializer.Serialize(new
                {
                    action = "error",
                    message = $"提取报价数据失败: {ex.Message}"
                }, JsonOptions));
            }
        }

        /// <summary>
        /// 线程安全向前端 WebView2 投递消息 (防释放与句柄未就绪异常)
        /// </summary>
        private void PostWebMessageSafe(string messageJson)
        {
            if (this.IsDisposed || !this.IsHandleCreated || _webView.IsDisposed) return;

            try
            {
                if (this.InvokeRequired)
                {
                    this.BeginInvoke(new Action(() =>
                    {
                        if (!_webView.IsDisposed && _webView.CoreWebView2 != null)
                        {
                            _webView.CoreWebView2.PostWebMessageAsJson(messageJson);
                        }
                    }));
                }
                else
                {
                    if (_webView.CoreWebView2 != null)
                    {
                        _webView.CoreWebView2.PostWebMessageAsJson(messageJson);
                    }
                }
            }
            catch { }
        }
    }
}
