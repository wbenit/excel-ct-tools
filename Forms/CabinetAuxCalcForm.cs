using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“智能辅材与壳体计算”无边框置顶宿主窗口
    /// </summary>
    public class CabinetAuxCalcForm : Form
    {
        // WebView2 浏览器主控件
        private readonly WebView2 _webView;

        // 辅材壳体计算控制器
        private readonly CabinetAuxCalcController _controller;

        // 导入 Windows 原生 user32.dll 接口用于支持拖拽窗体
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 原生消息接口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // Win32 常量: 标题栏鼠标左键按下消息
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // JSON 序列化选项
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化控制器与 WebView2 控件属性
        /// </summary>
        public CabinetAuxCalcForm()
        {
            // 实例化控制器
            _controller = new CabinetAuxCalcController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 设置 Form 窗体尺寸与显示几何外观 (840x660 像素)
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸 (840x660 像素，居中无边框)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题
            this.Text = "智能辅材与壳体计算";

            // 依据设计布局设定尺寸为 960x720 像素，确保所有定额表格与参数配置从容舒展
            this.ClientSize = new Size(960, 720);

            // 设置屏幕中央弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 设为无边框样式
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用原生最大化
            this.MaximizeBox = false;

            // 禁用原生最小化
            this.MinimizeBox = false;

            // 设置窗口背景填充色
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件并注册加载回调
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 控件充满窗体
            _webView.Dock = DockStyle.Fill;

            // 挂载至 Controls 控件集
            this.Controls.Add(_webView);

            // 绑定 Form Load 异步加载监听
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载异步事件处理程序，初始化 WebView2 运行环境并导航页面
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 指定 WebView2 用户专属数据缓存目录，避免因程序目录无写权限抛出拒绝访问异常
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

                // 禁用默认右键菜单
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;

                // 禁用 DevTools 生产环境安全策略 (可根据需要开启)
                _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                // 注册 Web 消息接收处理函数
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 检索 HTML 页面资源路径
                string appDir = Tool.GetAppDirectory();
                string[] candidatePaths = new[]
                {
                    Path.Combine(appDir, "Resources", "cabinet_aux_calc.html"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "cabinet_aux_calc.html"),
                    Path.Combine(Application.StartupPath, "Resources", "cabinet_aux_calc.html")
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

                // 导航至前端页面
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    _webView.Source = new Uri(htmlPath);
                }
                else
                {
                    MessageBox.Show($"未找到辅材壳体计算界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                string jsonString = string.Empty;
                try { jsonString = e.TryGetWebMessageAsString(); } catch { }
                if (string.IsNullOrEmpty(jsonString))
                {
                    try { jsonString = e.WebMessageAsJson; } catch { }
                }
                if (string.IsNullOrEmpty(jsonString)) return;

                using var doc = JsonDocument.Parse(jsonString);
                var root = doc.RootElement;

                if (!root.TryGetProperty("action", out var actionElement)) return;
                string action = actionElement.GetString() ?? string.Empty;

                // 响应窗口平滑位移拖拽 (基于非模态物理增量，彻底杜绝 Win32 模态循环死锁导致 Excel 崩溃)
                if (action == "moveWindow")
                {
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
                }
                // 响应窗口拖拽 (旧版兼容兜底)
                else if (action == "dragWindow" || action == "dragMove")
                {
                    SafeInvoke(() =>
                    {
                        ReleaseCapture();
                        SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                    });
                }
                // 响应最小化
                else if (action == "minimize")
                {
                    SafeInvoke(() => this.WindowState = FormWindowState.Minimized);
                }
                // 响应关闭
                else if (action == "close")
                {
                    SafeInvoke(() => this.Close());
                }
                // 响应获取初始上下文数据
                else if (action == "getContext")
                {
                    var context = _controller.GetInitialContext();
                    string resJson = JsonSerializer.Serialize(new
                    {
                        action = "initContext",
                        data = context
                    }, JsonOptions);

                    SafeInvoke(() => _webView.CoreWebView2.PostWebMessageAsString(resJson));
                }
                // 响应切换工作表获取箱柜列表
                else if (action == "getCabinetsBySheet")
                {
                    string sheetName = root.TryGetProperty("sheetName", out var sn) ? sn.GetString() ?? string.Empty : string.Empty;
                    var cabList = _controller.GetCabinetsBySheet(sheetName);
                    string resJson = JsonSerializer.Serialize(new
                    {
                        action = "cabinetsList",
                        data = cabList
                    }, JsonOptions);

                    SafeInvoke(() => _webView.CoreWebView2.PostWebMessageAsString(resJson));
                }
                // 响应保存规则配置
                else if (action == "saveRules")
                {
                    if (root.TryGetProperty("rules", out var rulesElem))
                    {
                        var rules = JsonSerializer.Deserialize<QuotationRules>(rulesElem.GetRawText(), JsonOptions);
                        bool success = false;
                        if (rules != null)
                        {
                            success = _controller.SaveRules(rules);
                        }
                        string resJson = JsonSerializer.Serialize(new
                        {
                            action = "saveRulesResult",
                            success = success
                        }, JsonOptions);
                        SafeInvoke(() => _webView.CoreWebView2.PostWebMessageAsString(resJson));
                    }
                }
                // 响应单个箱柜推导分析
                else if (action == "analyzeCabinet")
                {
                    string sheetName = root.TryGetProperty("sheetName", out var sn) ? sn.GetString() ?? string.Empty : string.Empty;
                    string detName = root.TryGetProperty("detName", out var dn) ? dn.GetString() ?? string.Empty : string.Empty;
                    QuotationRules? rules = null;
                    if (root.TryGetProperty("rules", out var rulesElem))
                    {
                        rules = JsonSerializer.Deserialize<QuotationRules>(rulesElem.GetRawText(), JsonOptions);
                    }

                    var result = _controller.AnalyzeSingleCabinet(sheetName, detName, rules ?? new QuotationRules());
                    string resJson = JsonSerializer.Serialize(new
                    {
                        action = "analyzeResult",
                        success = result != null,
                        data = result
                    }, JsonOptions);
                    SafeInvoke(() => _webView.CoreWebView2.PostWebMessageAsString(resJson));
                }
                // 响应批量/单台应用回写
                else if (action == "applyCalculation")
                {
                    string sheetName = root.TryGetProperty("sheetName", out var sn) ? sn.GetString() ?? string.Empty : string.Empty;
                    List<string> detNames = new List<string>();
                    if (root.TryGetProperty("detNames", out var detArray))
                    {
                        foreach (var item in detArray.EnumerateArray())
                        {
                            string s = item.GetString() ?? string.Empty;
                            if (!string.IsNullOrWhiteSpace(s)) detNames.Add(s);
                        }
                    }
                    QuotationRules? rules = null;
                    if (root.TryGetProperty("rules", out var rulesElem))
                    {
                        rules = JsonSerializer.Deserialize<QuotationRules>(rulesElem.GetRawText(), JsonOptions);
                    }

                    int count = _controller.ApplyCalculation(sheetName, detNames, rules ?? new QuotationRules());
                    string resJson = JsonSerializer.Serialize(new
                    {
                        action = "applyResult",
                        success = count > 0,
                        count = count
                    }, JsonOptions);
                    SafeInvoke(() => _webView.CoreWebView2.PostWebMessageAsString(resJson));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"处理前端WebMessage发生异常: {ex.Message}");
            }
        }
    }
}
