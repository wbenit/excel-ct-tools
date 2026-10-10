using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelDna.Integration;
using ExcelAddInDemo.Controllers;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“总价一键调整”无边框模态/非模态宿主窗口
    /// </summary>
    public class TotalPriceAdjustForm : Form
    {
        // 声明 WebView2 浏览器主控件
        private readonly WebView2 _webView;

        // 声明总价一键调整控制器
        private readonly TotalPriceAdjustController _controller;

        // 统一 JSON 解析选项
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化窗体属性与 WebView2 控件
        /// </summary>
        public TotalPriceAdjustForm()
        {
            // 实例化控制器
            _controller = new TotalPriceAdjustController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 设置窗体外观与尺寸
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸 (宽 800 x 高 680 像素)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 窗体标题文本
            this.Text = "总价一键调整";

            // 依据弹性布局设计设定尺寸为 800x680 像素
            this.ClientSize = new Size(800, 680);

            // 屏幕中央居中显示
            this.StartPosition = FormStartPosition.CenterScreen;

            // 无边框窗口样式
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用最大化
            this.MaximizeBox = false;

            // 允许最小化
            this.MinimizeBox = true;

            // 设置背景底色为纯白
            this.BackColor = Color.White;

            // 开启双缓冲减少重绘闪烁
            this.DoubleBuffered = true;
        }

        /// <summary>
        /// 初始化 WebView2 控件并绑定加载事件
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 充满整个窗口
            _webView.Dock = DockStyle.Fill;

            // 挂载至控件集
            this.Controls.Add(_webView);

            // 绑定窗体加载完成异步事件
            this.Load += async (sender, e) =>
            {
                await InitializeWebViewAsync();
            };
        }

        /// <summary>
        /// 异步初始化 WebView2 环境与加载 total_price_adjust.html 页面
        /// </summary>
        private async System.Threading.Tasks.Task InitializeWebViewAsync()
        {
            try
            {
                // 创建专用用户数据目录，避免进程冲突
                string userDataFolder = Path.Combine(Path.GetTempPath(), "ExcelAddInDemo_WebView2_TotalPriceAdjust");
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 初始化 CoreWebView2 核心实例
                await _webView.EnsureCoreWebView2Async(env);

                // 配置 WebView2 核心设置 (禁用默认右键菜单，开发模式可按需保留)
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

                // 绑定 WebMessage 消息接收回调
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 多路径智能检索 HTML 模板资源路径 (规则 4)
                string htmlPath = FindHtmlResourcePath("total_price_adjust.html");

                // 若资源文件存在则导航加载
                if (File.Exists(htmlPath))
                {
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
                }
                else
                {
                    MessageBox.Show($"未找到总价一键调整界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[TotalPriceAdjustForm] 初始化 WebView2 发生异常: {ex.Message}");
                MessageBox.Show($"初始化 WebView2 发生异常: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 多重备用路径检索指定名称的 HTML 资源文件物理全路径
        /// </summary>
        private string FindHtmlResourcePath(string fileName)
        {
            // 获取应用程序域基础基准目录
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            // 获取当前插件程序集实际物理加载目录 (兼容 XLL/DLL)
            string appDir = Tool.GetAppDirectory();

            // 候选路径 1: 优先从插件真实物理目录下的 Resources 目录查找
            string path1 = Path.Combine(appDir, "Resources", fileName);
            if (File.Exists(path1)) return path1;

            // 候选路径 2: 应用程序基准目录下的 Resources 目录查找
            string path2 = Path.Combine(baseDir, "Resources", fileName);
            if (File.Exists(path2)) return path2;

            // 候选路径 3: 插件物理目录下的发布目录 publish/Resources
            string path3 = Path.Combine(appDir, "publish", "Resources", fileName);
            if (File.Exists(path3)) return path3;

            // 候选路径 4: 从插件物理目录向上一级回退至源码根目录的 Resources (针对开发调试环境)
            string path4 = Path.Combine(appDir, "..", "..", "Resources", fileName);
            if (File.Exists(path4)) return path4;

            // 候选路径 5: 从基准目录向上一级回退至源码 Resources
            string path5 = Path.Combine(baseDir, "..", "..", "Resources", fileName);
            if (File.Exists(path5)) return path5;

            // 默认回退第一候选路径
            return path1;
        }

        /// <summary>
        /// 安全跨线程调度 UI 动作，防止句柄未创建或窗体已释放抛出异常
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
        /// 安全向前端 WebView2 回送字符串消息
        /// </summary>
        private void PostWebMessageAsStringSafe(string message)
        {
            SafeInvoke(() =>
            {
                try
                {
                    if (_webView != null && _webView.CoreWebView2 != null && !this.IsDisposed)
                    {
                        _webView.CoreWebView2.PostWebMessageAsString(message);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[TotalPriceAdjustForm] PostWebMessageAsStringSafe 异常: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 处理来自前端 Vue 3 的 postMessage 通信指令
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
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

                using var doc = JsonDocument.Parse(jsonString);
                var root = doc.RootElement;

                if (!root.TryGetProperty("action", out var actionElement)) return;
                string action = actionElement.GetString() ?? string.Empty;

                // 响应窗口位移拖拽 (物理增量坐标移动，杜绝 Win32 消息泵卡死)
                if (action == "moveWindow")
                {
                    int deltaX = root.TryGetProperty("deltaX", out var dxProp) ? dxProp.GetInt32() : 0;
                    int deltaY = root.TryGetProperty("deltaY", out var dyProp) ? dyProp.GetInt32() : 0;
                    if (deltaX != 0 || deltaY != 0)
                    {
                        SafeInvoke(() =>
                        {
                            this.Location = new Point(this.Left + deltaX, this.Top + deltaY);
                        });
                    }
                }
                // 响应最小化
                else if (action == "minimize" || action == "minimizeWindow")
                {
                    SafeInvoke(() => this.WindowState = FormWindowState.Minimized);
                }
                // 响应关闭窗体
                else if (action == "close" || action == "closeWindow")
                {
                    SafeInvoke(() => this.Close());
                }
                // 响应获取初始化数据请求
                else if (action == "getInitData")
                {
                    // 在 Excel 主线程宏队列中执行 COM 访问防死锁
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            string resultJson = _controller.GetInitData();
                            PostWebMessageAsStringSafe(resultJson);
                        }
                        catch (Exception ex)
                        {
                            LogHelper.WriteLog($"[TotalPriceAdjustForm] getInitData 异常: {ex.Message}");
                        }
                    });
                }
                // 响应执行总价一键调整请求
                else if (action == "executeAdjust")
                {
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            string resultJson = _controller.ExecuteAdjust(jsonString);
                            PostWebMessageAsStringSafe(resultJson);
                        }
                        catch (Exception ex)
                        {
                            LogHelper.WriteLog($"[TotalPriceAdjustForm] executeAdjust 异常: {ex.Message}");
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[TotalPriceAdjustForm] OnWebMessageReceived 处理消息异常: {ex.Message}");
            }
        }
    }
}
