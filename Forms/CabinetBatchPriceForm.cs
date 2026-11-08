using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo;
using ExcelAddInDemo.Models;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 箱体一键改价无边框宿主窗体 (C# + WebView2 + Vue3 + Element-Plus)
    /// 遵循规范：新增代码每 3 行至少包含 1 行中文注释
    /// </summary>
    public partial class CabinetBatchPriceForm : Form
    {
        // 核心 WebView2 浏览器控件实例
        private readonly WebView2 _webView = new WebView2();

        // 统一 JSON 序列化选项配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        #region Windows API 物理无边框窗口拖动支撑
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;
        private const int VK_LBUTTON = 0x01;
        #endregion

        /// <summary>
        /// 构造函数：初始化窗体属性与 WebView2 控件
        /// </summary>
        public CabinetBatchPriceForm()
        {
            // 配置窗体外观与尺寸
            InitializeFormProperties();

            // 挂载并初始化 WebView2
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题文本
            this.Text = "箱体一键改价";

            // 设定工作台尺寸为 1120x700 像素 --硬编码: 窗体初始尺寸--
            this.ClientSize = new Size(1120, 700);

            // 设置屏幕中央居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 设为无边框窗口样式
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用 WinForm 原生最大化
            this.MaximizeBox = false;

            // 启用最小化
            this.MinimizeBox = true;

            // 设置窗口背景填充色
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件并注册加载回调
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 控件满屏充满 Form
            _webView.Dock = DockStyle.Fill;

            // 挂载至 Controls 控件集
            this.Controls.Add(_webView);

            // 绑定 Form Load 异步加载监听
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// Form 加载事件: 初始化 WebView2 环境并导航至 cabinet_batch_price.html
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

                // 异步创建 WebView2 核心环境
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 判断窗体有效性
                if (this.IsDisposed || this.Disposing) return;

                // 初始化 CoreWebView2 核心对象
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 挂载前端交互消息回调
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 寻找目标 HTML 资源文件路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // 配置多级备选路径集 --避免调试与发布环境物理路径差异--
                string[] candidatePaths = new string[]
                {
                    Path.Combine(baseDir, "Resources", "cabinet_batch_price.html"),
                    Path.Combine(baseDir, "cabinet_batch_price.html"),
                    Path.Combine(Application.StartupPath, "Resources", "cabinet_batch_price.html"),
                    Path.Combine(Application.StartupPath, "cabinet_batch_price.html"),
                    Path.Combine(Tool.GetAppDirectory(), "Resources", "cabinet_batch_price.html"),
                    Path.Combine(Tool.GetAppDirectory(), "cabinet_batch_price.html"),
                    @"d:\code\excel-ct-tools\Resources\cabinet_batch_price.html",
                    @"d:\code\excel-ct-tools\publish\Resources\cabinet_batch_price.html"
                };

                string htmlPath = string.Empty;
                foreach (var path in candidatePaths)
                {
                    if (File.Exists(path))
                    {
                        htmlPath = path;
                        break;
                    }
                }

                // 若找到有效 HTML 文件则导航
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
                }
                else
                {
                    // 容错展示内联兜底提示
                    _webView.NavigateToString("<h3 style='color:red;padding:20px;'>未找到 cabinet_batch_price.html 资源页面，请确认文件已复制至运行目录。</h3>");
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CabinetBatchPriceForm] 初始化 WebView2 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理来自 WebView2 前端 Vue 3 的 postMessage 指令消息
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 获取前端传递的消息字符串
                string messageJson = "";
                try { messageJson = e.TryGetWebMessageAsString(); } catch { }

                if (string.IsNullOrEmpty(messageJson))
                {
                    try { messageJson = e.WebMessageAsJson; } catch { }
                }

                if (string.IsNullOrWhiteSpace(messageJson)) return;

                // 使用 JsonDocument 极速读取 action 字段
                using var doc = JsonDocument.Parse(messageJson);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;

                string action = actionProp.GetString() ?? "";

                // 根据指令分支分发任务
                switch (action)
                {
                    case "scanCabinets":
                        // 扫描箱体物料
                        HandleScanCabinets(root);
                        break;

                    case "applyPrice":
                        // 执行一键改价回写
                        HandleApplyPrice(root);
                        break;

                    case "dragWindow":
                        // 处理无边框窗体拖拽移动
                        HandleDragWindow();
                        break;

                    case "close":
                        // 安全关闭窗口
                        SafeInvoke(() => this.Close());
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CabinetBatchPriceForm] 处理前端消息异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理箱体物料扫描指令
        /// </summary>
        private void HandleScanCabinets(JsonElement root)
        {
            try
            {
                CabinetBatchPriceConfigDto config = null;
                if (root.TryGetProperty("config", out var configProp))
                {
                    config = JsonSerializer.Deserialize<CabinetBatchPriceConfigDto>(configProp.GetRawText(), JsonOptions);
                }

                // 调用 ExcelServices 扫描全表或当前表箱体物料
                var result = ExcelServices.ScanProjectCabinets(config);

                // 将扫描结果安全序列化并推送给前端
                string responseJson = JsonSerializer.Serialize(new
                {
                    action = "scanResult",
                    data = result
                }, JsonOptions);

                PostWebMessageSafe(responseJson);
            }
            catch (Exception ex)
            {
                // 推送异常消息
                string errorJson = JsonSerializer.Serialize(new
                {
                    action = "scanResult",
                    data = new CabinetBatchPriceScanResultDto
                    {
                        Success = false,
                        Message = $"扫描异常: {ex.Message}"
                    }
                }, JsonOptions);

                PostWebMessageSafe(errorJson);
            }
        }

        /// <summary>
        /// 处理箱体一键改价批量回写指令
        /// </summary>
        private void HandleApplyPrice(JsonElement root)
        {
            try
            {
                CabinetApplyPricePayload payload = null;
                if (root.TryGetProperty("payload", out var payloadProp))
                {
                    payload = JsonSerializer.Deserialize<CabinetApplyPricePayload>(payloadProp.GetRawText(), JsonOptions);
                }

                if (payload == null)
                {
                    PostWebMessageSafe(JsonSerializer.Serialize(new
                    {
                        action = "applyResult",
                        data = new { success = false, message = "改价请求载荷为空" }
                    }, JsonOptions));
                    return;
                }

                // 调用 ExcelServices 批量回写至 Excel
                var (success, message, count) = ExcelServices.ApplyCabinetBatchPriceToExcel(payload);

                // 响应执行结果至前端
                string responseJson = JsonSerializer.Serialize(new
                {
                    action = "applyResult",
                    data = new
                    {
                        success = success,
                        message = message,
                        updatedCount = count
                    }
                }, JsonOptions);

                PostWebMessageSafe(responseJson);
            }
            catch (Exception ex)
            {
                PostWebMessageSafe(JsonSerializer.Serialize(new
                {
                    action = "applyResult",
                    data = new { success = false, message = $"回写异常: {ex.Message}" }
                }, JsonOptions));
            }
        }

        /// <summary>
        /// 处理无边框窗体拖拽移动 (带物理按键防粘死锁保护)
        /// </summary>
        private void HandleDragWindow()
        {
            SafeInvoke(() =>
            {
                try
                {
                    // 检测鼠标左键物理状态 (最高位为 1 表示按下，为 0 表示已抬起)
                    bool isPhysicalDown = (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
                    if (!isPhysicalDown)
                    {
                        // 物理上已抬起按键，立即释放捕获退出，防止粘滞死锁
                        ReleaseCapture();
                        return;
                    }

                    // 释放鼠标捕获并向 Win32 发送标题栏拖拽消息
                    ReleaseCapture();
                    SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                }
                catch { }
            });
        }

        /// <summary>
        /// 安全向前端 WebView2 发送 JSON 字符串消息
        /// </summary>
        private void PostWebMessageSafe(string jsonMessage)
        {
            SafeInvoke(() =>
            {
                try
                {
                    if (_webView != null && _webView.CoreWebView2 != null && !this.IsDisposed)
                    {
                        _webView.CoreWebView2.PostWebMessageAsString(jsonMessage);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[CabinetBatchPriceForm] PostWebMessageSafe 异常: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// UI 线程安全调度辅助工具
        /// </summary>
        private void SafeInvoke(System.Action action)
        {
            try
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
            catch { }
        }

        /// <summary>
        /// 窗体销毁时释放 WebView2 核心资源
        /// </summary>
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try
            {
                _webView?.Dispose();
            }
            catch { }

            base.OnFormClosed(e);
        }
    }
}
