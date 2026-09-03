using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“智能输入配置”无边框宿主窗口
    /// </summary>
    public class SmartInputForm : Form
    {
        // 声明 WebView2 浏览器主控件句柄
        private readonly WebView2 _webView;

        // 声明智能输入后端 WebAPI 控制器句柄
        private readonly SmartInputController _controller;

        // 导入 Windows 原生 user32.dll 接口用于鼠标拖拽无边框窗口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 原生消息分发接口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // Win32 常量定义: 标题栏拖拽消息标识
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // JSON 序列化与反序列化通用配置结构
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化控制器与 WebView2 控件属性
        /// </summary>
        public SmartInputForm()
        {
            // 实例化后端 WebAPI 风格控制器
            _controller = new SmartInputController();

            // 实例化 WebView2 浏览器控件
            _webView = new WebView2();

            // 设置 Form 窗体几何外观与尺寸
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸 (820x600 像素)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题文本
            this.Text = "智能填写配置";

            // 设置尺寸为 820x600 像素
            this.ClientSize = new Size(820, 600);

            // 设置屏幕居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 设为无边框样式
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用最大化按钮
            this.MaximizeBox = false;

            // 启用最小化按钮
            this.MinimizeBox = true;

            // 设置背景填充色为白色
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件并注册加载回调
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 控件充满整个 Form 容器
            _webView.Dock = DockStyle.Fill;

            // 挂载至 WinForm 控件集合
            this.Controls.Add(_webView);

            // 绑定 Form 异步加载监听
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// Form 加载事件: 初始化 WebView2 环境并导航至 Resources/smart_input.html
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 窗体已释放则退出
                if (this.IsDisposed || this.Disposing) return;

                // 设置本地 AppData 中的 WebView2 专属缓存目录
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAddInDemo",
                    "WebView2Data"
                );

                // 创建缓存文件夹
                if (!Directory.Exists(userDataFolder))
                {
                    Directory.CreateDirectory(userDataFolder);
                }

                // 创建 CoreWebView2 运行环境
                var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 确保 CoreWebView2 控件完全初始化完毕
                await _webView.EnsureCoreWebView2Async(environment);

                // 禁用右键菜单以防止用户意外呼出原生调试菜单
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;

                // 注册 Web 消息接收处理委托
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 检索 smart_input.html 物理文件路径
                string appDir = Tool.GetAppDirectory();
                string[] candidatePaths = new[]
                {
                    Path.Combine(appDir, "Resources", "smart_input.html"),
                    Path.Combine(appDir, "smart_input.html"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "smart_input.html")
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

                // 导航至前端界面资源
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    _webView.Source = new Uri(htmlPath);
                }
                else
                {
                    MessageBox.Show($"未找到智能输入配置界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                try
                {
                    messageJson = e.TryGetWebMessageAsString();
                }
                catch { }

                if (string.IsNullOrEmpty(messageJson))
                {
                    messageJson = e.WebMessageAsJson;
                }

                if (string.IsNullOrEmpty(messageJson)) return;

                // 解析 JSON 数据文档
                using var doc = JsonDocument.Parse(messageJson);
                var root = doc.RootElement;

                // 获取 action 指令
                string action = root.TryGetProperty("action", out var actProp) ? actProp.GetString() ?? "" : "";

                switch (action)
                {
                    // 1. 初始化获取配置与各表元器件缓存
                    case "getInitData":
                        var config = _controller.GetConfig();
                        var storage = _controller.GetStoredComponents();

                        // 构造回发数据包
                        var initData = new
                        {
                            action = "renderInitData",
                            config = config,
                            storage = storage
                        };

                        PostWebMessageSafe(JsonSerializer.Serialize(initData, JsonOptions));
                        break;

                    // 2. 刷新并从当前工作簿提取元器件
                    case "refreshAndExtract":
                        var refreshedStorage = _controller.RefreshAndExtract();
                        var currentConfig = _controller.GetConfig();

                        // 若之前未勾选任何表，则默认全选所有新提取的表
                        if (currentConfig.SelectedSheets == null || currentConfig.SelectedSheets.Count == 0)
                        {
                            currentConfig.SelectedSheets = refreshedStorage.Sheets.Select(s => s.SheetName).ToList();
                            _controller.SaveConfig(currentConfig);
                        }

                        var refreshRes = new
                        {
                            action = "renderRefreshResult",
                            success = true,
                            config = currentConfig,
                            storage = refreshedStorage,
                            message = $"刷新成功！共从 {refreshedStorage.Sheets.Count} 个表中提取去重元器件数据。"
                        };

                        PostWebMessageSafe(JsonSerializer.Serialize(refreshRes, JsonOptions));
                        break;

                    // 3. 保存配置
                    case "saveConfig":
                        if (root.TryGetProperty("config", out var cfgProp))
                        {
                            var cfg = JsonSerializer.Deserialize<SmartInputConfigModel>(cfgProp.GetRawText(), JsonOptions);
                            if (cfg != null)
                            {
                                bool ok = _controller.SaveConfig(cfg);
                                var saveRes = new
                                {
                                    action = "saveConfigResult",
                                    success = ok,
                                    message = ok ? "配置保存成功！" : "配置保存失败！"
                                };
                                PostWebMessageSafe(JsonSerializer.Serialize(saveRes, JsonOptions));
                            }
                        }
                        break;

                    // 4. 应用下拉列表至当前工作表 C 列
                    case "applyDropdown":
                        if (root.TryGetProperty("config", out var appCfgProp))
                        {
                            var cfg = JsonSerializer.Deserialize<SmartInputConfigModel>(appCfgProp.GetRawText(), JsonOptions);
                            if (cfg != null)
                            {
                                var (succ, msg) = _controller.ApplyDropdown(cfg);
                                var appRes = new
                                {
                                    action = "applyDropdownResult",
                                    success = succ,
                                    message = msg
                                };
                                PostWebMessageSafe(JsonSerializer.Serialize(appRes, JsonOptions));
                            }
                        }
                        break;

                    // 5. 将选中的物料回填至 Excel 活动单元格所在行
                    case "fillToActiveRow":
                        if (root.TryGetProperty("item", out var itemProp))
                        {
                            var compItem = JsonSerializer.Deserialize<SmartComponentItem>(itemProp.GetRawText(), JsonOptions);
                            SmartInputConfigModel? fillCfg = null;
                            if (root.TryGetProperty("config", out var fCfgProp))
                            {
                                fillCfg = JsonSerializer.Deserialize<SmartInputConfigModel>(fCfgProp.GetRawText(), JsonOptions);
                            }

                            if (compItem != null)
                            {
                                bool ok = _controller.FillToActiveRow(compItem, fillCfg ?? _controller.GetConfig());
                                var fillRes = new
                                {
                                    action = "fillResult",
                                    success = ok,
                                    message = ok ? $"已成功填入规格型号【{compItem.Model}】！" : "回填失败，请先选中 Excel 中的有效单元格！"
                                };
                                PostWebMessageSafe(JsonSerializer.Serialize(fillRes, JsonOptions));
                            }
                        }
                        break;

                    // 6. 窗口平滑位移拖拽 (基于非模态物理增量，彻底杜绝 Win32 模态循环死锁导致 Excel 崩溃)
                    case "moveWindow":
                        int deltaX = root.TryGetProperty("deltaX", out var dxProp) ? dxProp.GetInt32() : 0;
                        int deltaY = root.TryGetProperty("deltaY", out var dyProp) ? dyProp.GetInt32() : 0;
                        if (deltaX != 0 || deltaY != 0)
                        {
                            SafeInvoke(() =>
                            {
                                // 直接更新窗体屏幕坐标，微秒级响应且绝不挂起 STA 消息泵
                                this.Location = new Point(this.Left + deltaX, this.Top + deltaY);
                            });
                        }
                        break;

                    // 6.1 窗口顶部拖动响应 (旧版兼容兜底)
                    case "windowDrag":
                        if (this.WindowState == FormWindowState.Normal)
                        {
                            ReleaseCapture();
                            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        }
                        break;

                    // 7. 最小化窗口
                    case "minimize":
                        this.WindowState = FormWindowState.Minimized;
                        break;

                    // 8. 关闭窗口
                    case "close":
                        this.Close();
                        break;
                }
            }
            catch (Exception ex)
            {
                // 记录消息处理异常
                LogHelper.WriteLog($"SmartInputForm 消息处理异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 跨线程安全向前端 WebView2 发送 JSON 消息
        /// </summary>
        private void PostWebMessageSafe(string jsonMessage)
        {
            try
            {
                if (this.IsDisposed || this.Disposing) return;

                if (this.InvokeRequired)
                {
                    this.BeginInvoke(new Action(() =>
                    {
                        if (_webView != null && _webView.CoreWebView2 != null)
                        {
                            _webView.CoreWebView2.PostWebMessageAsString(jsonMessage);
                        }
                    }));
                }
                else
                {
                    if (_webView != null && _webView.CoreWebView2 != null)
                    {
                        _webView.CoreWebView2.PostWebMessageAsString(jsonMessage);
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 跨线程安全执行委托操作
        /// </summary>
        private void SafeInvoke(Action action)
        {
            // 检查窗体是否已经释放或句柄未创建
            if (this.IsDisposed || !this.IsHandleCreated) return;
            // 判断是否需要跨线程封送
            if (this.InvokeRequired)
            {
                // 异步派发到 UI 线程执行
                this.BeginInvoke(action);
            }
            else
            {
                // 当前即为 UI 线程直接执行
                action();
            }
        }
    }
}
