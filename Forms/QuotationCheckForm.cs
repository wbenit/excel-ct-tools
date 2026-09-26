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

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 + Element Plus 的“报价智能核验向导”现代无边框窗体
    /// 主色调 #009688 绿蓝商务科技感，弹性响应式布局，零水平滚动条，安全拖拽与 IPC 解耦
    /// </summary>
    public class QuotationCheckForm : Form
    {
        // 声明 WebView2 浏览器核心组件
        private readonly WebView2 _webView;

        // 声明报价核验业务控制器
        private readonly QuotationCheckController _controller;

        // P/Invoke 导入 user32.dll 原生接口用于支持无边框拖拽
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // P/Invoke 导入 SendMessage 原生消息分发接口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // P/Invoke 导入 GetAsyncKeyState 接口检测鼠标物理按键状态 (杜绝幽灵拖拽死锁)
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int VK_LBUTTON = 0x01;
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // JSON 序列化通用配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// 构造函数: 初始化控制器、WebView2 宿主容器及窗体属性
        /// </summary>
        public QuotationCheckForm()
        {
            // 实例化业务控制器
            _controller = new QuotationCheckController(this);

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 初始化窗体基本外观与尺寸 (1180x720 像素)
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 初始化窗体尺寸与外观
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设定窗体标题 --硬编码: 窗体标题--
            this.Text = "报价智能核验与防错体检";
            // 设定宽屏比例 1180x720，提供充裕的指标卡片与问题明细展示空间
            this.ClientSize = new Size(1180, 720);
            // 屏幕居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;
            // 无边框现代扁平风格
            this.FormBorderStyle = FormBorderStyle.None;
            // 禁用最小化与最大化按钮
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            // 背景深蓝灰微调
            this.BackColor = Color.FromArgb(240, 244, 248);
            // 显示在 Windows 任务栏
            this.ShowInTaskbar = true;
            // 确保窗口置顶展示，便于与 Excel 并排操作
            this.TopMost = true;
        }

        /// <summary>
        /// 配置并挂载 WebView2 控件
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 充满整个窗体 Client 区域
            _webView.Dock = DockStyle.Fill;
            // 将 WebView2 控件添加至窗体 Controls 集合
            this.Controls.Add(_webView);

            // 订阅核心初始化完成事件
            _webView.CoreWebView2InitializationCompleted += OnWebViewInitialized;

            // 异步初始化 WebView2 运行环境
            InitializeWebViewAsync();
        }

        /// <summary>
        /// 异步初始化 WebView2 核心环境
        /// </summary>
        private async void InitializeWebViewAsync()
        {
            try
            {
                // 创建临时用户数据目录
                string userDataFolder = Path.Combine(Path.GetTempPath(), "ExcelAddIn_QuotationCheck_WebView2");
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                // 确保 CoreWebView2 引擎就绪
                await _webView.EnsureCoreWebView2Async(env);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[QuotationCheckForm] 初始化 WebView2 失败: {ex.Message}");
            }
        }

        /// <summary>
        /// WebView2 初始化完成回调处理
        /// </summary>
        private void OnWebViewInitialized(object? sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            if (!e.IsSuccess)
            {
                LogHelper.WriteLog($"[QuotationCheckForm] CoreWebView2 初始化失败: {e.InitializationException?.Message}");
                return;
            }

            // 禁用默认上下文菜单
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
            // 允许开启 DevTools 便于调试
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

            // 订阅 Web 消息监听器
            _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

            // 探测前端 HTML 资源文件路径
            string htmlPath = LocateHtmlResource("quotation_check.html");
            if (File.Exists(htmlPath))
            {
                // 以本地 File 协议导航进入前端界面
                _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
            }
            else
            {
                // 若未找到资源则呈现友好的降级错误提示 HTML
                _webView.CoreWebView2.NavigateToString(
                    $"<div style='padding:40px;color:red;font-family:sans-serif;'>未找到前端界面资源文件: {htmlPath}</div>");
            }
        }

        /// <summary>
        /// 处理前端 Vue 3 发送的异步 IPC 消息指令
        /// 遵循规范: 严禁在此回调中直接调用 MessageBox.Show 阻塞 Chromium IPC 握手
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string json = e.WebMessageAsJson;
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;

                string action = actionProp.GetString() ?? string.Empty;

                switch (action)
                {
                    // 1. 无边框窗口原生安全拖拽
                    case "dragWindow":
                        // 校验鼠标左键物理状态，彻底杜绝幽灵鼠标捕获死锁
                        if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0)
                        {
                            ReleaseCapture();
                            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        }
                        break;

                    // 2. 关闭向导窗口
                    case "closeWindow":
                        SafeInvoke(() => this.Close());
                        break;

                    // 3. 初始加载 / 重新刷新报价智能核验
                    case "runAudit":
                    case "refresh":
                        RunAuditAsync();
                        break;

                    // 4. Excel 视口定位与高亮跳转
                    case "jumpToCell":
                        if (root.TryGetProperty("sheetName", out var sProp) &&
                            root.TryGetProperty("row", out var rProp) &&
                            root.TryGetProperty("colLetter", out var cProp))
                        {
                            string sName = sProp.GetString() ?? string.Empty;
                            int r = rProp.GetInt32();
                            string cLetter = cProp.GetString() ?? "C";
                            _controller.NavigateToCell(sName, r, cLetter);
                        }
                        break;

                    // 5. 一键批量修复公式损坏行
                    case "fixFormulas":
                        List<string>? issueIds = null;
                        if (root.TryGetProperty("issueIds", out var idsProp) && idsProp.ValueKind == JsonValueKind.Array)
                        {
                            issueIds = new List<string>();
                            foreach (var item in idsProp.EnumerateArray())
                            {
                                string? strId = item.GetString();
                                if (!string.IsNullOrEmpty(strId)) issueIds.Add(strId);
                            }
                        }

                        // 异步执行公式修复
                        var fixRes = _controller.FixFormulas(issueIds);
                        // 回传修复结果通知前端展示 Toast
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "formulaFixed",
                            success = fixRes.Success,
                            message = fixRes.Message,
                            fixedCount = fixRes.FixedCount
                        }, JsonOptions));

                        // 修复完成后自动触发一次重新审计以更新视图
                        if (fixRes.Success)
                        {
                            RunAuditAsync();
                        }
                        break;

                    // 6. 同型号价格批量对齐广播
                    case "alignModelPrice":
                        if (root.TryGetProperty("params", out var paramsProp))
                        {
                            var alignReq = JsonSerializer.Deserialize<AlignModelPriceRequest>(paramsProp.GetRawText(), JsonOptions);
                            if (alignReq != null)
                            {
                                var alignRes = _controller.AlignModelPrice(alignReq);
                                PostWebMessageSafe(JsonSerializer.Serialize(new
                                {
                                    action = "modelAligned",
                                    success = alignRes.Success,
                                    message = alignRes.Message,
                                    updatedCount = alignRes.UpdatedCount
                                }, JsonOptions));

                                if (alignRes.Success)
                                {
                                    RunAuditAsync();
                                }
                            }
                        }
                        break;

                    // 7. 一键清洗文本型数字
                    case "cleanTextNumbers":
                        var cleanRes = _controller.CleanTextNumbers();
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "textNumbersCleaned",
                            success = cleanRes.Success,
                            message = cleanRes.Message,
                            cleanedCount = cleanRes.CleanedCount
                        }, JsonOptions));

                        if (cleanRes.Success)
                        {
                            RunAuditAsync();
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[QuotationCheckForm] 消息处理异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 异步执行报价体检并将结果推送给前端 Vue
        /// </summary>
        private void RunAuditAsync()
        {
            try
            {
                var auditResult = _controller.RunAudit();
                string payload = JsonSerializer.Serialize(new
                {
                    action = "auditDataLoaded",
                    data = auditResult
                }, JsonOptions);

                PostWebMessageSafe(payload);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[QuotationCheckForm] 执行审计异常: {ex.Message}");
                PostWebMessageSafe(JsonSerializer.Serialize(new
                {
                    action = "auditDataLoaded",
                    data = new QuotationCheckResult { Success = false, Message = ex.Message }
                }, JsonOptions));
            }
        }

        /// <summary>
        /// 线程安全向 WebView2 前端回发 JSON 消息
        /// </summary>
        private void PostWebMessageSafe(string json)
        {
            SafeInvoke(() =>
            {
                if (_webView?.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.PostWebMessageAsJson(json);
                }
            });
        }

        /// <summary>
        /// 控件安全跨线程调用封装
        /// </summary>
        private void SafeInvoke(System.Action action)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            if (this.InvokeRequired)
            {
                this.Invoke(action);
            }
            else
            {
                action();
            }
        }

        /// <summary>
        /// 多级检索前端 HTML 资源文件的物理路径
        /// </summary>
        private static string LocateHtmlResource(string fileName)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string[] probePaths = new[]
            {
                Path.Combine(baseDir, "Resources", fileName),
                Path.Combine(baseDir, "publish", "Resources", fileName),
                Path.Combine(baseDir, fileName),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "Resources", fileName)
            };

            foreach (var path in probePaths)
            {
                if (File.Exists(path)) return Path.GetFullPath(path);
            }

            return Path.Combine(baseDir, "Resources", fileName);
        }
    }
}
