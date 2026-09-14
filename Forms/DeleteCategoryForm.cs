using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“批量删除分类”宿主窗口
    /// </summary>
    public class DeleteCategoryForm : Form
    {
        // 声明 WebView2 浏览器核心控件
        private readonly WebView2 _webView;

        // 声明分类控制器实例
        private readonly CategoryController _controller;

        // 导入 Windows 原生 user32.dll 接口用于支持原生无边框拖拽
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 原生消息接口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 常量定义: 标题栏鼠标左键按下消息
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 通用 JSON 序列化配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化控制器与 WebView2 控件属性
        /// </summary>
        public DeleteCategoryForm()
        {
            // 实例化分类控制器
            _controller = new CategoryController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 配置 Form 窗体尺寸与外观
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸
        /// </summary>
        private void InitializeFormProperties()
        {
            // 窗体标题 --硬编码: 窗体标题--
            this.Text = "删除分类";
            // 设定适宜列表展示的高宽尺寸
            this.ClientSize = new Size(540, 480);
            // 屏幕居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;
            // 设为无边框现代扁平样式
            this.FormBorderStyle = FormBorderStyle.None;
            // 禁用最大化与最小化
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            // 背景填充纯白
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件并注册加载回调
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 满框铺满 Form
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
                // 窗体已被销毁则直接退出
                if (this.IsDisposed || this.Disposing) return;

                // 设置本地 AppData 中的 WebView2 专属缓存目录
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAddInDemo",
                    "WebView2Data"
                );

                // 自动创建缓存文件夹
                Directory.CreateDirectory(userDataFolder);

                // 异步创建 WebView2 核心环境
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 二次校验窗体生命周期
                if (this.IsDisposed || this.Disposing) return;

                // 确保 CoreWebView2 核心对象就绪
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 注册前端 postMessage 通信监听
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 寻找目标 HTML 静态资源物理路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string appDir = Tool.GetAppDirectory();

                // 配置多级备选路径集以实现高容错加载
                string[] candidatePaths = new string[]
                {
                    // 候选路径 1: 插件真实运行物理目录下的 Resources
                    Path.Combine(appDir, "Resources", "delete_category.html"),
                    // 候选路径 2: 插件运行物理目录根路径
                    Path.Combine(appDir, "delete_category.html"),
                    // 候选路径 3: 当前 AppDomain 根路径/Resources
                    Path.Combine(baseDir, "Resources", "delete_category.html"),
                    // 候选路径 4: 当前 AppDomain 根路径
                    Path.Combine(baseDir, "delete_category.html"),
                    // 候选路径 5: 当前工作目录下的 Resources
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "delete_category.html")
                };

                // 循环检索首个真实存在的目标 HTML 文件
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
                    MessageBox.Show($"未找到删除分类界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                // 记录初始化异常
                LogHelper.WriteLog($"初始化删除分类 WebView2 异常: {ex.Message}");
                MessageBox.Show($"初始化界面发生异常: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 安全跨线程调度 UI 操作
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
        /// 处理来自前端 Vue 3 的 WebMessage 请求
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 解析前端传递的 JSON 消息字符串
                string jsonString = e.TryGetWebMessageAsString();
                if (string.IsNullOrEmpty(jsonString)) return;

                using var doc = JsonDocument.Parse(jsonString);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;

                string action = actionProp.GetString() ?? string.Empty;

                switch (action)
                {
                    // 窗口平滑位移拖拽 (基于物理屏幕增量，纯非模态调度)
                    case "moveWindow":
                        int deltaX = 0;
                        int deltaY = 0;
                        if (root.TryGetProperty("data", out var dataElem))
                        {
                            deltaX = dataElem.TryGetProperty("deltaX", out var dxp) ? dxp.GetInt32() : 0;
                            deltaY = dataElem.TryGetProperty("deltaY", out var dyp) ? dyp.GetInt32() : 0;
                        }
                        else
                        {
                            deltaX = root.TryGetProperty("deltaX", out var dxp) ? dxp.GetInt32() : 0;
                            deltaY = root.TryGetProperty("deltaY", out var dyp) ? dyp.GetInt32() : 0;
                        }
                        if (deltaX != 0 || deltaY != 0)
                        {
                            SafeInvoke(() =>
                            {
                                this.Location = new Point(this.Left + deltaX, this.Top + deltaY);
                            });
                        }
                        break;

                    case "dragWindow":
                        // 原生拖拽兜底
                        ReleaseCapture();
                        SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        break;

                    case "getDeleteCategoriesData":
                        // 请求分类列表与选区感知数据
                        SafeInvoke(() =>
                        {
                            var data = _controller.GetDeleteCategoriesData();
                            var resObj = new
                            {
                                action = "renderDeleteCategoriesData",
                                data = data
                            };
                            string resJson = JsonSerializer.Serialize(resObj, JsonOptions);
                            _webView.CoreWebView2.PostWebMessageAsString(resJson);
                        });
                        break;

                    case "deleteCategories":
                        // 前端提交批量删除分类请求
                        if (root.TryGetProperty("data", out var reqProp))
                        {
                            var req = JsonSerializer.Deserialize<DeleteCategoriesRequest>(reqProp.GetRawText(), JsonOptions);
                            if (req != null)
                            {
                                SafeInvoke(() =>
                                {
                                    // 执行批量删除核心业务
                                    var result = _controller.DeleteCategories(req);

                                    // 回发删除执行结果
                                    var resObj = new
                                    {
                                        action = "deleteCategoriesResult",
                                        data = result
                                    };
                                    string resJson = JsonSerializer.Serialize(resObj, JsonOptions);
                                    _webView.CoreWebView2.PostWebMessageAsString(resJson);

                                    // 若删除成功，延时 300ms 安全关闭窗口
                                    if (result.Success)
                                    {
                                        var closeTimer = new System.Windows.Forms.Timer { Interval = 300 };
                                        closeTimer.Tick += (s, args) =>
                                        {
                                            closeTimer.Stop();
                                            closeTimer.Dispose();
                                            this.Close();
                                        };
                                        closeTimer.Start();
                                    }
                                });
                            }
                        }
                        break;

                    case "cancel":
                    case "close":
                        // 取消或关闭当前窗口
                        SafeInvoke(() => this.Close());
                        break;
                }
            }
            catch (Exception ex)
            {
                // 记录消息处理异常
                LogHelper.WriteLog($"处理删除分类 WebMessage 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 窗体关闭时显式释放 WebView2 控件资源，杜绝句柄泄漏
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try
            {
                // 解绑消息回调
                if (_webView?.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
                }
                // 显式销毁 WebView2 控件
                _webView?.Dispose();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[DeleteCategoryForm] OnFormClosing 释放异常: {ex.Message}");
            }
            base.OnFormClosing(e);
        }
    }
}
