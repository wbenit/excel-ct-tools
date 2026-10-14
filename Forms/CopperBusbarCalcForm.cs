using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的智能母排/铜排计算器无边框弹窗宿主窗体
    /// </summary>
    public class CopperBusbarCalcForm : Form
    {
        // 声明 WebView2 浏览器核心控件
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
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化窗体属性与 WebView2 控件
        /// </summary>
        public CopperBusbarCalcForm()
        {
            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 设置 Form 窗体尺寸与显示几何外观 (1120x720 像素)
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸 (宽 1120，高 720 像素，宽屏弹性展示)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题文本
            this.Text = "智能母排/铜排计算器";

            // 依据三栏式工作台布局设定尺寸为 1120x720 像素 --硬编码: 窗体初始尺寸--
            this.ClientSize = new Size(1120, 720);

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
        /// Form 加载事件: 初始化 WebView2 环境并导航至 Resources/copper_busbar_calc.html
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
                    Path.Combine(baseDir, "Resources", "copper_busbar_calc.html"),
                    Path.Combine(baseDir, "..", "Resources", "copper_busbar_calc.html"),
                    Path.Combine(baseDir, "publish", "Resources", "copper_busbar_calc.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "copper_busbar_calc.html")
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
                    MessageBox.Show($"未找到母排计算器界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化 WebView2 控件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 响应来自前端 Vue 3 发来的 JSON 交互消息 (双轨通信规范)
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 获取前端传递的消息字符串（优先读取 String，备选读取 Raw JSON）
                string messageJson = "";
                try
                {
                    messageJson = e.TryGetWebMessageAsString();
                }
                catch { }

                // 若尝试读取为空，则获取原生 WebMessageAsJson 内容
                if (string.IsNullOrEmpty(messageJson))
                {
                    try { messageJson = e.WebMessageAsJson; } catch { }
                }

                if (string.IsNullOrWhiteSpace(messageJson)) return;

                // 解析为 JsonDocument 模型
                using var doc = JsonDocument.Parse(messageJson);
                var root = doc.RootElement;

                // 获取指令 Action 名称
                string action = root.TryGetProperty("action", out var actProp) ? actProp.GetString() ?? "" : "";

                // 根据 Action 分发业务
                switch (action)
                {
                    case "getProjectCabinets":
                        // 前端初始化时请求当前工程所有箱柜及 W/X 列信息
                        HandleGetProjectCabinets();
                        break;

                    case "writeBackCopperBusbar":
                        // 接收前端回填请求，将铜排物料插入 Excel
                        HandleWriteBackCopperBusbar(messageJson);
                        break;

                    case "dragWindow":
                        // 响应标题栏无边框拖拽移动
                        HandleDragWindow();
                        break;

                    case "minimizeWindow":
                        // 窗口最小化
                        this.WindowState = FormWindowState.Minimized;
                        break;

                    case "closeWindow":
                        // 关闭窗口
                        this.Close();
                        break;
                }
            }
            catch (Exception ex)
            {
                // 仅静默记录日志，杜绝阻断式异常弹窗
                LogHelper.WriteLog($"[CopperBusbarCalcForm] 处理交互异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理获取当前工作表中所有箱柜及其 W 列电流、X 列极数
        /// </summary>
        private void HandleGetProjectCabinets()
        {
            try
            {
                // 调用公共服务层提取当前分类表箱柜
                var cabinets = ExcelServices.GetCabinetsForBusbarCalc();

                // 组装回发前端的消息载荷
                var responseObj = new
                {
                    action = "projectCabinetsLoaded",
                    success = true,
                    cabinets = cabinets
                };

                // 线程安全回送给前端 WebView2
                PostWebMessageSafe(JsonSerializer.Serialize(responseObj, JsonOptions));
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CopperBusbarCalcForm] 提取箱柜失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理铜排物料行安全插入至 Excel 分类表
        /// </summary>
        private void HandleWriteBackCopperBusbar(string payloadJson)
        {
            try
            {
                // 将回填载荷反序列化为强类型对象
                var payload = JsonSerializer.Deserialize<CopperWriteBackPayload>(payloadJson, JsonOptions);
                if (payload == null) return;

                // 调用 ExcelServices 执行规则 6 安全插行与写入
                bool success = ExcelServices.InsertCopperBusbarBomToExcel(payload);

                if (success)
                {
                    // 延迟 800ms 后安全关闭当前弹窗
                    var closeTimer = new System.Windows.Forms.Timer { Interval = 800 };
                    closeTimer.Tick += (s, ev) =>
                    {
                        closeTimer.Stop();
                        closeTimer.Dispose();
                        SafeInvoke(() => this.Close());
                    };
                    closeTimer.Start();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"写入 Excel 失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
        /// 跨线程安全执行调度器 (遵循经验 10 防范句柄未创建异常)
        /// </summary>
        private void SafeInvoke(Action action)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            if (this.InvokeRequired)
            {
                try { this.Invoke(action); } catch { }
            }
            else
            {
                action();
            }
        }
    }
}
