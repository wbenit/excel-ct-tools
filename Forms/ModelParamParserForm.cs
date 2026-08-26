using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“元器件型号参数识别设置”无边框宿主窗口
    /// </summary>
    public class ModelParamParserForm : Form
    {
        // 声明 WebView2 浏览器主控件
        private readonly WebView2 _webView;

        // 声明后端控制器
        private readonly ModelParamParserController _controller;

        // 导入 Windows 原生 user32.dll 接口用于无边框窗口拖拽
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 常量定义: 标题栏拖拽消息标识
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
        /// 构造函数: 初始化控制器与窗体属性
        /// </summary>
        public ModelParamParserForm()
        {
            // 实例化后端控制器
            _controller = new ModelParamParserController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 配置窗体外观与几何尺寸
            InitializeFormProperties();

            // 配置挂载 WebView2
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体几何外观 (960x700 像素)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题
            this.Text = "型号参数识别设置 (极数与电流)";

            // 设置尺寸为 960x700 像素
            this.ClientSize = new Size(960, 700);

            // 设置屏幕居中显示
            this.StartPosition = FormStartPosition.CenterScreen;

            // 设为无边框样式
            this.FormBorderStyle = FormBorderStyle.None;

            // 窗体强力置顶显示
            this.TopMost = true;

            // 禁用原生最大化
            this.MaximizeBox = false;

            // 禁用最小化
            this.MinimizeBox = false;

            // 设置背景填充色
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件并挂载事件
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 控件满屏停靠
            _webView.Dock = DockStyle.Fill;

            // 挂载至 Controls 控件集
            this.Controls.Add(_webView);

            // 绑定 Load 事件
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载事件: 初始化 WebView2 环境并导航至 HTML 资源
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                if (this.IsDisposed || this.Disposing) return;

                // 设置 WebView2 缓存路径
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAddInDemo",
                    "WebView2Data"
                );
                Directory.CreateDirectory(userDataFolder);

                // 创建 WebView2 运行时环境
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                if (this.IsDisposed || this.Disposing) return;

                // 初始化 CoreWebView2 核心对象
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 挂载前端交互消息监听
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 寻找目标 HTML 资源文件
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string[] candidatePaths = new string[]
                {
                    Path.Combine(baseDir, "Resources", "model_param_parser.html"),
                    Path.Combine(baseDir, "..", "Resources", "model_param_parser.html"),
                    Path.Combine(baseDir, "publish", "Resources", "model_param_parser.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "model_param_parser.html")
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

                // 导航至 HTML 页面
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    _webView.Source = new Uri(htmlPath);
                }
                else
                {
                    MessageBox.Show($"未找到型号参数识别界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化 WebView2 控件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 响应来自前端 Vue 3 发来的 JSON 交互指令
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 获取消息文本
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
                    // 1. 获取当前最新配置
                    case "getConfig":
                        var currentCfg = _controller.LoadConfig();
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "renderConfig",
                            config = currentCfg
                        }, JsonOptions));
                        break;

                    // 2. 保存用户配置
                    case "saveConfig":
                        if (root.TryGetProperty("config", out var cfgProp))
                        {
                            var saveObj = JsonSerializer.Deserialize<ModelParserConfig>(cfgProp.GetRawText(), JsonOptions);
                            bool ok = _controller.SaveConfig(saveObj!);
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "saveResult",
                                success = ok,
                                message = ok ? "配置保存成功！" : "配置保存失败！"
                            }, JsonOptions));
                        }
                        break;

                    // 3. 恢复出厂默认配置
                    case "resetConfig":
                        var defaultCfg = _controller.ResetToDefault();
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "renderConfig",
                            config = defaultCfg,
                            message = "已恢复出厂默认配置"
                        }, JsonOptions));
                        break;

                    // 4. 实时沙盒单条型号测试
                    case "testParse":
                        string rawModel = root.TryGetProperty("rawModel", out var rawProp) ? rawProp.GetString() ?? "" : "";
                        ModelParserConfig? testCfg = null;
                        if (root.TryGetProperty("config", out var testCfgProp))
                        {
                            testCfg = JsonSerializer.Deserialize<ModelParserConfig>(testCfgProp.GetRawText(), JsonOptions);
                        }
                        var testRes = _controller.TestParse(rawModel, testCfg!);
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "testParseResult",
                            result = testRes
                        }, JsonOptions));
                        break;

                    // 5. 执行 Excel 批量识别回填
                    case "executeBatch":
                        ModelParserConfig? execCfg = null;
                        if (root.TryGetProperty("config", out var execCfgProp))
                        {
                            execCfg = JsonSerializer.Deserialize<ModelParserConfig>(execCfgProp.GetRawText(), JsonOptions);
                        }
                        var batchRes = _controller.ExecuteBatch(execCfg!);
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "batchExecuteResult",
                            result = batchRes
                        }, JsonOptions));
                        break;

                    // 6. 窗体窗口拖拽
                    case "dragWindow":
                        SafeInvoke(() =>
                        {
                            ReleaseCapture();
                            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        });
                        break;

                    // 7. 最小化窗口
                    case "minimize":
                        SafeInvoke(() => this.WindowState = FormWindowState.Minimized);
                        break;

                    // 8. 关闭窗口
                    case "close":
                        SafeInvoke(() => this.Close());
                        break;

                    // 9. 响应窗体尺寸调整指令 (支持折叠为单行状态栏或展开完整设置面板)
                    case "resizeWindow":
                        // 提取目标宽度，默认 960 像素
                        int width = root.TryGetProperty("width", out var wElem) ? wElem.GetInt32() : 960;
                        // 提取目标高度，默认 700 像素
                        int height = root.TryGetProperty("height", out var hElem) ? hElem.GetInt32() : 700;
                        // 调度至 UI 线程更新窗体尺寸
                        SafeInvoke(() =>
                        {
                            // 判断尺寸是否有变动，避免无效重绘
                            if (this.ClientSize.Width != width || this.ClientSize.Height != height)
                            {
                                // 设置窗体工作区尺寸
                                this.ClientSize = new Size(width, height);
                            }
                        });
                        break;
                }
            }
            catch (Exception ex)
            {
                PostWebMessageSafe(JsonSerializer.Serialize(new
                {
                    action = "error",
                    message = $"处理异常: {ex.Message}"
                }, JsonOptions));
            }
        }

        /// <summary>
        /// UI 线程安全执行辅助方法
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
        /// 跨线程安全回发消息给前端 WebView2
        /// </summary>
        private void PostWebMessageSafe(string jsonMessage)
        {
            if (this.IsDisposed || this.Disposing) return;

            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        if (_webView?.CoreWebView2 != null)
                        {
                            _webView.CoreWebView2.PostWebMessageAsString(jsonMessage);
                        }
                    }
                    catch { }
                }));
            }
            else
            {
                try
                {
                    if (_webView?.CoreWebView2 != null)
                    {
                        _webView.CoreWebView2.PostWebMessageAsString(jsonMessage);
                    }
                }
                catch { }
            }
        }
    }
}
