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
    /// 基于 WebView2 + Vue 3 的母排一键改价无边框工作台宿主窗体
    /// </summary>
    public class BusbarBatchPriceForm : Form
    {
        // 声明 WebView2 核心控件
        private readonly WebView2 _webView;

        // 导入 Windows 原生 user32.dll 接口用于标题栏拖拽窗口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 原生系统消息分发接口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 导入 GetAsyncKeyState 接口检测物理鼠标按键是否松开（根除幽灵鼠标捕获死锁）
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        // 常量定义: 标题栏拖拽与左键虚拟键码
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;
        private const int VK_LBUTTON = 0x01;

        // 通用 JSON 序列化配置结构
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        /// <summary>
        /// 构造函数: 初始化窗体属性与 WebView2 控件
        /// </summary>
        public BusbarBatchPriceForm()
        {
            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 设置 Form 窗体尺寸与显示几何外观 (1080x680 像素)
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题文本
            this.Text = "母排一键改价";

            // 设定工作台尺寸为 1080x680 像素 --硬编码: 窗体初始尺寸--
            this.ClientSize = new Size(1080, 680);

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
        /// Form 加载事件: 初始化 WebView2 环境并导航至 busbar_batch_price.html
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
                    Path.Combine(baseDir, "Resources", "busbar_batch_price.html"),
                    Path.Combine(baseDir, "..", "Resources", "busbar_batch_price.html"),
                    Path.Combine(baseDir, "publish", "Resources", "busbar_batch_price.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "busbar_batch_price.html")
                };

                // 保存匹配的目标文件路径
                string htmlPath = string.Empty;

                // 循环查找首个存在的目标文件
                foreach (string candidate in candidatePaths)
                {
                    if (File.Exists(candidate))
                    {
                        htmlPath = candidate;
                        break;
                    }
                }

                // 导航至目标 HTML 资源
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    _webView.Source = new Uri(htmlPath);
                }
                else
                {
                    MessageBox.Show($"未找到母排一键改价界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化 WebView2 控件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 响应来自前端 Vue 3 发来的 JSON 交互消息
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
                    case "scanBusbars":
                        // 扫描母排物料
                        HandleScanBusbars(root);
                        break;

                    case "getMarketPrice":
                        // 在主线程先同步读取 forceRefresh 参数，杜绝跨线程访问已释放的 JsonDocument
                        bool forceRefresh = false;
                        // 若前端传递了 forceRefresh 属性则提取布尔值
                        if (root.TryGetProperty("forceRefresh", out var forceProp))
                        {
                            forceRefresh = forceProp.GetBoolean();
                        }
                        // 触发异步大盘行情拉取
                        HandleGetMarketPrice(forceRefresh);
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
                LogHelper.WriteLog($"处理前端消息异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理母排物料扫描指令
        /// </summary>
        private void HandleScanBusbars(JsonElement root)
        {
            try
            {
                BusbarBatchPriceConfigDto config = null;
                if (root.TryGetProperty("config", out var configProp))
                {
                    config = JsonSerializer.Deserialize<BusbarBatchPriceConfigDto>(configProp.GetRawText(), JsonOptions);
                }

                // 调用 ExcelServices 扫描全表或当前表母排物料
                var result = ExcelServices.ScanProjectBusbars(config);

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
                    data = new BusbarBatchPriceScanResultDto
                    {
                        Success = false,
                        Message = $"扫描异常: {ex.Message}"
                    }
                }, JsonOptions);

                PostWebMessageSafe(errorJson);
            }
        }

        /// <summary>
        /// 处理一键改价批量回写指令
        /// </summary>
        private void HandleApplyPrice(JsonElement root)
        {
            try
            {
                BusbarApplyPricePayload payload = null;
                if (root.TryGetProperty("payload", out var payloadProp))
                {
                    payload = JsonSerializer.Deserialize<BusbarApplyPricePayload>(payloadProp.GetRawText(), JsonOptions);
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
                var (success, message, count) = ExcelServices.ApplyBusbarBatchPriceToExcel(payload);

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
        /// 异步获取当日大盘铜价与铝价最新行情 (与云物料及智能算料完全统一)
        /// </summary>
        /// <param name="forceRefresh">是否强制跳过短周期内存缓存刷新行情</param>
        private void HandleGetMarketPrice(bool forceRefresh)
        {
            // 在后台工作线程执行异步网络调用，避免阻塞 UI 主线程与 Chromium IPC
            System.Threading.Tasks.Task.Run(async () =>
            {
                try
                {
                    // 调用统一的 MarketPriceService 提取长江/上海有色大盘现货行情
                    var result = await Services.MarketPriceService.FetchMarketPricesAsync(forceRefresh);

                    // 折算为元/kg (元/吨除以 1000，四舍五入保留 2 位小数)
                    double copperPricePerKg = Math.Round(result.CopperPrice / 1000.0, 2);
                    // 折算铝价为元/kg
                    double aluminumPricePerKg = Math.Round(result.AluminumPrice / 1000.0, 2);

                    // 封装统一响应报文
                    var response = new
                    {
                        action = "marketPriceResponse",
                        success = result.Success,
                        copperPricePerKg = copperPricePerKg,
                        copperPricePerTon = result.CopperPrice,
                        aluminumPricePerKg = aluminumPricePerKg,
                        aluminumPricePerTon = result.AluminumPrice,
                        source = result.Source,
                        updateTime = result.UpdateTime,
                        isFallback = result.IsFallback,
                        message = result.Message
                    };

                    // 线程安全投递至前端 WebView2
                    PostWebMessageSafe(JsonSerializer.Serialize(response, JsonOptions));
                }
                catch (Exception ex)
                {
                    // 记录异常日志
                    LogHelper.WriteLog($"[BusbarBatchPriceForm] HandleGetMarketPrice 异常: {ex.Message}");
                    PostWebMessageSafe(JsonSerializer.Serialize(new
                    {
                        action = "marketPriceResponse",
                        success = false,
                        message = $"获取大盘行情失败: {ex.Message}"
                    }, JsonOptions));
                }
            });
        }

        /// <summary>
        /// 处理无边框窗体拖拽移动 (带物理按键防粘死锁保护)
        /// </summary>
        private void HandleDragWindow()
        {
            // 校验当前物理鼠标左键是否仍在按下状态，避免幽灵捕获死锁
            if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) == 0) return;

            // 释放当前鼠标捕获
            ReleaseCapture();

            // 发送系统移动窗口消息
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
        }

        /// <summary>
        /// 线程安全向前端 WebView2 投递消息
        /// </summary>
        private void PostWebMessageSafe(string message)
        {
            SafeInvoke(() =>
            {
                if (_webView != null && _webView.CoreWebView2 != null && !this.IsDisposed)
                {
                    _webView.CoreWebView2.PostWebMessageAsString(message);
                }
            });
        }

        /// <summary>
        /// 跨线程安全执行调度器
        /// </summary>
        private void SafeInvoke(Action action)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            if (this.InvokeRequired)
            {
                try
                {
                    this.BeginInvoke(action);
                }
                catch { }
            }
            else
            {
                action();
            }
        }
    }
}

