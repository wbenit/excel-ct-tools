using System;
using System.Collections.Generic;
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
    /// 基于 WebView2 + Vue 3 的“二次元件组规则管道构建器”无边框宿主窗体
    /// </summary>
    public class ComponentGroupBuilderForm : Form
    {
        // 声明 WebView2 浏览器控件
        private readonly WebView2 _webView;

        // 声明后端控制器
        private readonly ComponentGroupBuilderController _controller;

        // 导入 Windows 原生 user32.dll 接口用于无边框窗口拖拽
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 常量定义: 标题栏拖拽消息标识
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 通用 JSON 序列化配置结构 (驼峰命名)
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化控制器与窗体属性
        /// </summary>
        public ComponentGroupBuilderForm()
        {
            // 实例化后端控制器
            _controller = new ComponentGroupBuilderController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 配置窗体几何外观
            InitializeFormProperties();

            // 配置挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体几何外观 (1360x820 像素大视口)
        /// </summary>
        private void InitializeFormProperties()
        {
            this.Text = "二次元件组规则管道构建器";
            this.ClientSize = new Size(1360, 820);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件并挂载事件
        /// </summary>
        private void InitializeWebViewControl()
        {
            _webView.Dock = DockStyle.Fill;
            this.Controls.Add(_webView);
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
                    Path.Combine(baseDir, "Resources", "component_group_builder.html"),
                    Path.Combine(baseDir, "..", "Resources", "component_group_builder.html"),
                    Path.Combine(baseDir, "publish", "Resources", "component_group_builder.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "component_group_builder.html")
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
                    MessageBox.Show($"未找到二次元件组规则管道界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                            var saveObj = JsonSerializer.Deserialize<ComponentGroupConfig>(cfgProp.GetRawText(), JsonOptions);
                            bool ok = _controller.SaveConfig(saveObj!);
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "saveResult",
                                success = ok,
                                message = ok ? "二次元件组规则管道已成功保存！" : "配置保存失败！"
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
                            message = "已恢复出厂默认二次元件组规则库！"
                        }, JsonOptions));
                        break;

                    // 4. 从 Excel 中抓取当前选中箱柜的真实元件数据
                    case "loadActiveCabinet":
                        ComponentGroupConfig? activeCfg = null;
                        if (root.TryGetProperty("config", out var actCfgProp))
                        {
                            activeCfg = JsonSerializer.Deserialize<ComponentGroupConfig>(actCfgProp.GetRawText(), JsonOptions);
                        }
                        var cabComponents = _controller.GetActiveCabinetComponents(activeCfg!);
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "activeCabinetLoaded",
                            components = cabComponents,
                            count = cabComponents.Count
                        }, JsonOptions));
                        break;

                    // 5. 实时沙盒测试规则管道
                    case "testPipeline":
                        ComponentGroupConfig? testCfg = null;
                        if (root.TryGetProperty("config", out var testCfgProp))
                        {
                            testCfg = JsonSerializer.Deserialize<ComponentGroupConfig>(testCfgProp.GetRawText(), JsonOptions);
                        }
                        List<EleComponentDto>? testComps = null;
                        if (root.TryGetProperty("components", out var testCompsProp))
                        {
                            testComps = JsonSerializer.Deserialize<List<EleComponentDto>>(testCompsProp.GetRawText(), JsonOptions);
                        }
                        var testRes = _controller.RunSandboxTest(testCfg!, testComps ?? new List<EleComponentDto>());
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "pipelineTestResult",
                            result = testRes
                        }, JsonOptions));
                        break;

                    // 6. 执行 Excel 批量生成
                    case "executeBatch":
                        ComponentGroupConfig? execCfg = null;
                        if (root.TryGetProperty("config", out var execCfgProp))
                        {
                            execCfg = JsonSerializer.Deserialize<ComponentGroupConfig>(execCfgProp.GetRawText(), JsonOptions);
                        }
                        bool activeCabinetOnly = true;
                        if (root.TryGetProperty("activeCabinetOnly", out var scopeProp))
                        {
                            activeCabinetOnly = scopeProp.GetBoolean();
                        }
                        var batchRes = _controller.ExecuteBatch(execCfg!, activeCabinetOnly);
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "batchExecuteResult",
                            result = batchRes
                        }, JsonOptions));
                        break;

                    // 7. 表达式转条件树
                    case "parseExpression":
                        string rawExpr = root.TryGetProperty("expression", out var exprProp) ? exprProp.GetString() ?? "" : "";
                        int policyInt = root.TryGetProperty("policy", out var polProp) ? polProp.GetInt32() : 0;
                        var treeRes = _controller.ParseExpression(rawExpr, (QuantityPolicy)policyInt);
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "parseExpressionResult",
                            tree = treeRes
                        }, JsonOptions));
                        break;

                    // 8. 窗体窗口拖拽
                    case "dragWindow":
                        SafeInvoke(() =>
                        {
                            ReleaseCapture();
                            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        });
                        break;

                    // 9. 最小化窗口
                    case "minimize":
                        SafeInvoke(() => this.WindowState = FormWindowState.Minimized);
                        break;

                    // 10. 最大化 / 还原窗口切换
                    case "toggleMaximize":
                        SafeInvoke(() =>
                        {
                            this.WindowState = this.WindowState == FormWindowState.Maximized
                                ? FormWindowState.Normal
                                : FormWindowState.Maximized;
                        });
                        break;

                    // 11. 关闭窗口
                    case "close":
                        SafeInvoke(() => this.Close());
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
