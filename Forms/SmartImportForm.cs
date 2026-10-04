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
    /// 基于 WebView2 + Vue 3 的智能导入箱柜 BOM 模态窗口
    /// 遵循主色调 #009688、弹性布局与无水平滚动条规范
    /// </summary>
    public class SmartImportForm : Form
    {
        // 声明 WebView2 浏览器控件
        private readonly WebView2 _webView;
        // 声明智能导入业务控制器
        private readonly SmartImportController _controller;

        // Windows 原生拖拽 API 导入
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 统一 JSON 序列化设置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        /// <summary>
        /// 构造函数初始化控件与基础外观
        /// </summary>
        public SmartImportForm()
        {
            // 初始化控制器
            _controller = new SmartImportController();
            // 初始化 WebView2 控件
            _webView = new WebView2();

            // 设置窗口基本属性
            this.Text = "智能导入箱柜 BOM 工作台";
            this.ClientSize = new Size(1160, 740);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            // 挂载 WebView2
            _webView.Dock = DockStyle.Fill;
            this.Controls.Add(_webView);

            // 绑定异步加载
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗口加载时初始化 WebView2 环境与本地 HTML 页面
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                if (this.IsDisposed || this.Disposing) return;

                // 设置本地缓存目录
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAddInDemo",
                    "WebView2Data"
                );
                Directory.CreateDirectory(userDataFolder);

                // 创建运行环境并绑定控件
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                if (this.IsDisposed || this.Disposing) return;

                await _webView.EnsureCoreWebView2Async(env);
                if (_webView.CoreWebView2 != null)
                {
                    // 启用开发者调试工具 (支持快捷键 F12 呼出控制台)
                    _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;
                    // 注册前端 WebMessage 消息监听
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                    // 监听页面导航加载结果防范异常白屏
                    _webView.CoreWebView2.NavigationCompleted += (s, args) =>
                    {
                        if (!args.IsSuccess)
                        {
                            LogHelper.WriteLog($"[SmartImportForm] 页面导航异常: {args.WebErrorStatus}");
                        }
                    };
                }

                // 寻找 smart_import.html 模板文件物理路径
                string htmlPath = FindHtmlResourcePath("smart_import.html");
                if (File.Exists(htmlPath))
                {
                    // 提取模板所在的真实文件夹物理路径
                    string resDir = Path.GetDirectoryName(htmlPath)!;
                    // 将本地资源目录安全映射为虚拟主机域名 https://appassets.local (对齐 CabinetAuxCalcForm 架构)
                    _webView.CoreWebView2?.SetVirtualHostNameToFolderMapping(
                        "appassets.local",
                        resDir,
                        Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.Allow);

                    // 优先通过标准安全的虚拟主机域名导航，杜绝 file:/// 协议的跨域与沙箱拦截
                    // 追加当前时间戳动态参数，彻底禁用 WebView2 静态资源本地 HTTP 缓存，确保最新页面即刻生效
                    string freshUrl = $"https://appassets.local/smart_import.html?v={DateTime.Now.Ticks}";
                    _webView.Source = new Uri(freshUrl);
                }
                else
                {
                    MessageBox.Show($"未找到界面资源文件: {htmlPath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[SmartImportForm] OnFormLoadAsync 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 接收前端 postMessage 消息分发处理
        /// <summary>
        /// 接收前端 postMessage 消息分发处理 (双轨兼容字符串与原生 JSON 对象)
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 双轨提取消息文本：优先作为字符串提取，失败则读取 WebMessageAsJson
                string rawJson = string.Empty;
                try { rawJson = e.TryGetWebMessageAsString(); } catch { }
                if (string.IsNullOrWhiteSpace(rawJson))
                {
                    try { rawJson = e.WebMessageAsJson; } catch { }
                }

                if (string.IsNullOrWhiteSpace(rawJson)) return;

                using var doc = JsonDocument.Parse(rawJson);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;

                string action = actionProp.GetString() ?? "";

                switch (action)
                {
                    // 1. 点击选择外部 Excel 文件 (在宿主主线程以 this 为 Owner 弹窗置顶)
                    case "selectFile":
                        SafeInvoke(() =>
                        {
                            try
                            {
                                using var dialog = new OpenFileDialog
                                {
                                    Title = "请选择外部 Excel 标书清单文件",
                                    Filter = "Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*",
                                    CheckFileExists = true,
                                    Multiselect = false,
                                    AutoUpgradeEnabled = true
                                };

                                // 以当前模态窗体作为 Owner，确保 100% 居中置顶，绝不被遮挡
                                if (dialog.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.FileName))
                                {
                                    string chosenPath = dialog.FileName;
                                    // 异步线程后台加载网格数据，防止阻塞界面
                                    System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                                    {
                                        var preview = _controller.LoadPreviewGrid(chosenPath);
                                        SafeInvoke(() =>
                                        {
                                            PostWebMessageSafe(JsonSerializer.Serialize(new
                                            {
                                                action = "fileSelectedResult",
                                                filePath = chosenPath,
                                                preview = preview
                                            }, JsonOptions));
                                        });
                                    });
                                }
                                else
                                {
                                    // 用户取消选择，通知前端重置 loading 状态
                                    PostWebMessageSafe(JsonSerializer.Serialize(new
                                    {
                                        action = "fileSelectedCancel"
                                    }, JsonOptions));
                                }
                            }
                            catch (Exception ex)
                            {
                                LogHelper.WriteLog($"[selectFile] 弹窗异常: {ex.Message}");
                                PostWebMessageSafe(JsonSerializer.Serialize(new
                                {
                                    action = "fileSelectedCancel"
                                }, JsonOptions));
                            }
                        });
                        break;

                    // 2. 切换选中的工作表 Sheet
                    case "changeSheet":
                        if (root.TryGetProperty("filePath", out var fp) && root.TryGetProperty("sheetName", out var sn))
                        {
                            string path = fp.GetString() ?? "";
                            string sheet = sn.GetString() ?? "";
                            // 异步加载选定工作表网格数据
                            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                            {
                                var preview = _controller.LoadPreviewGrid(path, sheet);
                                SafeInvoke(() =>
                                {
                                    PostWebMessageSafe(JsonSerializer.Serialize(new
                                    {
                                        action = "previewGridLoaded",
                                        preview = preview
                                    }, JsonOptions));
                                });
                            });
                        }
                        break;

                    // 3. 用户在样本柜完成 4 步点选后，触发全量推导解析预览
                    case "parsePreview":
                        if (root.TryGetProperty("filePath", out var pFp) &&
                            root.TryGetProperty("sheetName", out var pSn) &&
                            root.TryGetProperty("config", out var pCfg))
                        {
                            string path = pFp.GetString() ?? "";
                            string sheet = pSn.GetString() ?? "";
                            var cfg = JsonSerializer.Deserialize<SmartImportTemplateConfig>(pCfg.GetRawText(), JsonOptions);
                            var cabinets = _controller.ParsePreview(path, sheet, cfg);

                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "parsePreviewResult",
                                success = true,
                                cabinets = cabinets
                            }, JsonOptions));
                        }
                        break;

                    // 4. 执行一键导入 (新建分类表并写入数据)
                    case "executeImport":
                        if (root.TryGetProperty("data", out var dataProp))
                        {
                            var req = JsonSerializer.Deserialize<SmartImportExecuteRequest>(dataProp.GetRawText(), JsonOptions);
                            var result = _controller.ExecuteImport(req);

                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "executeImportResult",
                                result = result
                            }, JsonOptions));

                            // 若导入成功，延迟 1.5 秒自动关闭弹窗
                            if (result != null && result.Success)
                            {
                                var timer = new System.Windows.Forms.Timer { Interval = 1500 };
                                timer.Tick += (s, args) =>
                                {
                                    timer.Stop();
                                    timer.Dispose();
                                    SafeInvoke(() => this.Close());
                                };
                                timer.Start();
                            }
                        }
                        break;

                    // 5. 保存用户点选的规则模板
                    case "saveTemplate":
                        if (root.TryGetProperty("config", out var saveCfg))
                        {
                            var cfg = JsonSerializer.Deserialize<SmartImportTemplateConfig>(saveCfg.GetRawText(), JsonOptions);
                            bool ok = _controller.SaveTemplate(cfg);
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "saveTemplateResult",
                                success = ok
                            }, JsonOptions));
                        }
                        break;

                    // 6. 获取所有已保存的规则模板
                    case "getTemplates":
                        var templates = _controller.GetTemplates();
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "getTemplatesResult",
                            templates = templates
                        }, JsonOptions));
                        break;

                    // 7. 拖拽无边框窗体移动
                    case "dragWindow":
                        ReleaseCapture();
                        SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        break;

                    // 8. 关闭窗口
                    case "closeWindow":
                        this.Close();
                        break;

                    // 9. 接收前端上报的异常与错误，统一持久化至日志文件
                    case "logError":
                        if (root.TryGetProperty("error", out var errProp))
                        {
                            // 记录前端传入的详细异常堆栈
                            string errText = errProp.GetString() ?? "";
                            LogHelper.WriteLog($"[SmartImportForm 前端异常] {errText}");
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[SmartImportForm] OnWebMessageReceived 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 检索 HTML 页面资源物理路径
        /// </summary>
        private static string FindHtmlResourcePath(string fileName)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string appDir = Tool.GetAppDirectory();

            string[] paths = new string[]
            {
                Path.Combine(appDir, "Resources", fileName),
                Path.Combine(baseDir, "Resources", fileName),
                Path.Combine(appDir, "publish", "Resources", fileName),
                Path.Combine(baseDir, "publish", "Resources", fileName)
            };

            foreach (var p in paths)
            {
                if (File.Exists(p)) return p;
            }

            return paths[0];
        }

        /// <summary>
        /// 线程安全向前端 WebView2 发送 JSON 消息
        /// </summary>
        private void PostWebMessageSafe(string messageJson)
        {
            SafeInvoke(() =>
            {
                try
                {
                    if (_webView != null && _webView.CoreWebView2 != null)
                    {
                        _webView.CoreWebView2.PostWebMessageAsString(messageJson);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[PostWebMessageSafe] 发送消息异常: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 跨线程安全 Invoke 辅助方法
        /// </summary>
        private void SafeInvoke(Action action)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            if (this.InvokeRequired)
            {
                this.BeginInvoke(action);
            }
            else
            {
                action();
            }
        }
    }
}
