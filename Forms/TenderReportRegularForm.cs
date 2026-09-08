using System;
using System.IO;
using System.Drawing;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using ExcelAddInDemo.Controllers;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“常规样式投标报表”导出宿主窗口
    /// </summary>
    public class TenderReportRegularForm : Form
    {
        // 声明 WebView2 浏览器控件
        private readonly WebView2 _webView;

        // 声明常规投标报表业务控制器
        private readonly TenderReportRegularController _controller;

        // 导入 Windows 原生 user32.dll 接口用于窗体无边框拖拽移动
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int VK_LBUTTON = 0x01;
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 全局通用的 JSON 序列化规范
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// 构造函数：初始化控制器与 WebView2 控件
        /// </summary>
        public TenderReportRegularForm()
        {
            // 实例化业务交互控制器
            _controller = new TenderReportRegularController(this);

            // 实例化 WebView2 浏览器控件
            _webView = new WebView2();

            // 配置窗体外观与初始尺寸
            InitializeFormProperties();

            // 初始化 WebView2 布局
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体无边框、适中尺寸与屏幕居中
        /// </summary>
        private void InitializeFormProperties()
        {
            // 标题文本
            this.Text = "常规样式投标报表导出";

            // 初始显示尺寸调整为 740x660 像素，留足表格与表单可视空间
            this.ClientSize = new Size(740, 660);

            // 屏幕居中
            this.StartPosition = FormStartPosition.CenterScreen;

            // 无边框窗口样式
            this.FormBorderStyle = FormBorderStyle.None;

            // 隐藏最大化按钮
            this.MaximizeBox = false;

            // 背景色
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件布局与加载监听
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 充满工作区
            _webView.Dock = DockStyle.Fill;

            // 添加至 Controls
            this.Controls.Add(_webView);

            // 注册 Form Load 异步处理
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载异步事件：初始化 CoreWebView2 并导航至 tender_report_regular.html
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                if (this.IsDisposed || this.Disposing) return;

                // 计算本地缓存路径
                string userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAddInDemo", "WebView2Data");
                Directory.CreateDirectory(userDataFolder);

                // 创建 WebView2 运行时环境
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                await _webView.EnsureCoreWebView2Async(env);

                // 禁用默认右键上下文菜单
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;

                // 注册 WebMessageReceived 消息监听
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 检索前端 HTML 资源绝对物理路径
                string appDir = Tool.GetAppDirectory();
                string htmlPath = Path.Combine(appDir, "Resources", "tender_report_regular.html");

                // 若首选路径不存在，尝试从备用路径查找
                if (!File.Exists(htmlPath))
                {
                    htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "tender_report_regular.html");
                }

                // 导航至本地 HTML 页面
                if (File.Exists(htmlPath))
                {
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
                }
                else
                {
                    MessageBox.Show($"未找到报表向导界面文件: {htmlPath}", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[TenderReportForm] 窗体初始化异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 响应前端发送的 Web 消息并进行路由分发
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 提取前端传入的消息文本 (双轨兼容：优先提取纯字符串，若为原生对象则读取 WebMessageAsJson)
                string rawJson = string.Empty;
                try { rawJson = e.TryGetWebMessageAsString(); } catch { }
                if (string.IsNullOrWhiteSpace(rawJson))
                {
                    try { rawJson = e.WebMessageAsJson; } catch { }
                }

                // 消息空校验防护
                if (string.IsNullOrWhiteSpace(rawJson)) return;

                // 初次反序列化 JSON 文档
                using var doc = JsonDocument.Parse(rawJson);
                var root = doc.RootElement;

                // 若根节点是字符串字面量 (如前端使用 postMessage(JSON.stringify(...)))，进行二次解包
                if (root.ValueKind == JsonValueKind.String)
                {
                    string innerJson = root.GetString() ?? "";
                    if (!string.IsNullOrWhiteSpace(innerJson))
                    {
                        using var innerDoc = JsonDocument.Parse(innerJson);
                        root = innerDoc.RootElement.Clone();
                    }
                }

                // 提取动作 action 指令名称
                string action = root.TryGetProperty("action", out var actProp) ? actProp.GetString() ?? "" : "";

                switch (action)
                {
                    // 1. 获取初始数据 (项目信息 + 分类概况)
                    case "getInitialData":
                        {
                            string dataJson = _controller.GetInitialDataJson();
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "initialDataLoaded",
                                payload = JsonDocument.Parse(dataJson).RootElement
                            }, JsonOptions));
                        }
                        break;

                    // 2. 执行常规报表导出
                    case "exportReport":
                        {
                            if (root.TryGetProperty("config", out var configProp))
                            {
                                string configRaw = configProp.GetRawText();
                                string resultJson = _controller.ExportReport(configRaw);

                                PostWebMessageSafe(JsonSerializer.Serialize(new
                                {
                                    action = "exportReportResult",
                                    payload = JsonDocument.Parse(resultJson).RootElement
                                }, JsonOptions));
                            }
                        }
                        break;

                    // 3. 打开导出文件所在文件夹
                    case "openExportedFileLocation":
                        {
                            if (root.TryGetProperty("filePath", out var pathProp))
                            {
                                string path = pathProp.GetString() ?? "";
                                _controller.OpenExportedFileLocation(path);
                            }
                        }
                        break;

                    // 4. 响应平滑增量物理位移 (防止模态死锁)
                    case "moveWindow":
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
                        break;

                    // 5. 标题栏无边框拖拽移动 (兼容旧指令)
                    case "dragWindow":
                        {
                            // 物理按键校验防止死锁
                            if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0)
                            {
                                ReleaseCapture();
                                SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                            }
                        }
                        break;

                    // 6. 关闭窗体
                    case "closeWindow":
                    case "close":
                        _controller.CloseWindow();
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[TenderReportForm] 消息分发异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 线程安全地向前端发送 Web 消息
        /// </summary>
        public void PostWebMessageSafe(string message)
        {
            SafeInvoke(() =>
            {
                if (_webView != null && _webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.PostWebMessageAsString(message);
                }
            });
        }

        /// <summary>
        /// 线程安全调度委托执行
        /// </summary>
        public void SafeInvoke(Action action)
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
        /// 窗体关闭时显式释放 WebView2 控件资源，杜绝进程残留与 Excel 退出阻塞
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try
            {
                // 解绑 WebMessageReceived 事件防止悬空引用
                if (_webView?.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
                }
                // 显式销毁 WebView2 控件释放底层 Chromium 句柄
                _webView?.Dispose();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[TenderReportRegularForm] OnFormClosing 释放异常: {ex.Message}");
            }
            base.OnFormClosing(e);
        }
    }
}
