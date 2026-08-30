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
    /// 基于 WebView2 + Vue 3 的“新建分类”宿主窗口
    /// </summary>
    public class CategoryForm : Form
    {
        // 声明 WebView2 浏览器主控件
        private readonly WebView2 _webView;

        // 声明分类管理控制器实例
        private readonly CategoryController _controller;

        // 导入 Windows 原生 user32.dll 接口用于支持无边框拖拽窗体
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
        public CategoryForm()
        {
            // 实例化分类控制器
            _controller = new CategoryController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 设置 Form 窗体尺寸与显示外观
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸 (480x420 像素)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题文本
            this.Text = "新建分类";

            // 依据界面布局与元素高度设定尺寸为 480x420 像素，确保底部确定与取消按钮完整展示
            this.ClientSize = new Size(480, 420);

            // 设置屏幕中央弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 设为无边框样式
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用 WinForm 原生最大化
            this.MaximizeBox = false;

            // 禁用最小化
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

                // 设置本地 AppData 中的 WebView2 缓存生成路径
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

                // 配置多级备选路径集 --优先从插件真实物理安装目录加载--
                string[] candidatePaths = new string[]
                {
                    // 候选路径 1: 插件运行物理目录下的 Resources
                    Path.Combine(appDir, "Resources", "category.html"),
                    // 候选路径 2: 插件运行物理目录
                    Path.Combine(appDir, "category.html"),
                    // 候选路径 3: 当前 AppDomain 根路径/Resources
                    Path.Combine(baseDir, "Resources", "category.html"),
                    // 候选路径 4: 当前 AppDomain 根路径
                    Path.Combine(baseDir, "category.html"),
                    // 候选路径 5: 当前工作目录下的 Resources
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "category.html")
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
                    MessageBox.Show($"未找到新建分类界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化 WebView2 发生异常: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 安全跨线程调度 UI 动作
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
                // 解析前端传递的 JSON 字符串消息
                string jsonString = e.TryGetWebMessageAsString();
                if (string.IsNullOrEmpty(jsonString)) return;

                using var doc = JsonDocument.Parse(jsonString);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;

                string action = actionProp.GetString() ?? string.Empty;

                switch (action)
                {
                    case "dragWindow":
                        // 处理无边框窗体拖拽移动
                        ReleaseCapture();
                        SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        break;

                    case "getSuggestInfo":
                        // 前端页面加载时请求分类建议与公式组数据
                        SafeInvoke(() =>
                        {
                            var suggestInfo = _controller.GetCategorySuggestInfo();
                            var resObj = new
                            {
                                action = "renderSuggestInfo",
                                data = suggestInfo
                            };
                            string resJson = JsonSerializer.Serialize(resObj, JsonOptions);
                            _webView.CoreWebView2.PostWebMessageAsString(resJson);
                        });
                        break;

                    case "createCategory":
                        // 响应前端提交新建分类请求
                        if (root.TryGetProperty("data", out var dataProp))
                        {
                            var req = JsonSerializer.Deserialize<CreateCategoryRequest>(dataProp.GetRawText(), JsonOptions);
                            if (req != null)
                            {
                                SafeInvoke(() =>
                                {
                                    // 调度控制器执行创建操作
                                    var result = _controller.CreateCategory(req);

                                    // 回发操作结果消息
                                    var resObj = new
                                    {
                                        action = "createCategoryResult",
                                        data = result
                                    };
                                    string resJson = JsonSerializer.Serialize(resObj, JsonOptions);
                                    _webView.CoreWebView2.PostWebMessageAsString(resJson);

                                    // 若创建成功，延时关闭窗口
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
                        // 关闭当前新建分类窗口
                        SafeInvoke(() => this.Close());
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"处理新建分类 WebMessage 异常: {ex.Message}");
            }
        }
    }
}
