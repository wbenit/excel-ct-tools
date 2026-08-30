using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的元器件物料匹配与品牌必含约束规则配置弹窗
    /// </summary>
    public class ComponentMatchForm : Form
    {
        // 声明 WebView2 浏览器控件实例
        private readonly WebView2 _webView;

        // 声明业务控制器对象
        private readonly ComponentMatchController _controller;

        // 静态单例引用，防止重复弹窗
        private static ComponentMatchForm? _currentInstance;

        // JSON 序列化通用配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        #region Windows API 窗口无边框拖拽支持

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        #endregion

        /// <summary>
        /// 静态入口：显示或激活物料匹配规则设置窗口
        /// </summary>
        public static void ShowDialogForm()
        {
            // 校验是否已存在活动窗口实例
            if (_currentInstance != null && !_currentInstance.IsDisposed)
            {
                if (_currentInstance.WindowState == FormWindowState.Minimized)
                {
                    _currentInstance.WindowState = FormWindowState.Normal;
                }
                _currentInstance.BringToFront();
                _currentInstance.Activate();
                return;
            }

            // 实例化新窗口并展示
            _currentInstance = new ComponentMatchForm();
            _currentInstance.Show();
        }

        /// <summary>
        /// 构造函数：初始化窗口外观尺寸与 WebView2 控件
        /// </summary>
        public ComponentMatchForm()
        {
            _controller = new ComponentMatchController();
            _webView = new WebView2();

            // 配置窗体基础属性 (480x620 像素无边框精致悬浮窗)
            this.Text = "元器件物料匹配与品牌规则设置";
            this.ClientSize = new Size(500, 620);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = true;
            this.TopMost = true; // 置顶显示方便对照 Excel
            this.BackColor = Color.FromArgb(248, 250, 252);

            // 配置 WebView2 填充布局
            _webView.Dock = DockStyle.Fill;
            this.Controls.Add(_webView);

            // 注册加载与关闭生命周期事件
            this.Load += OnFormLoadAsync;
            this.FormClosed += (s, e) => { _currentInstance = null; };
        }

        /// <summary>
        /// 异步加载 WebView2 环境并导航加载 HTML
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 获取专属用户缓存数据目录
                string userDataDir = Path.Combine(Tool.GetAppDataDirectory(), "WebView2_MatchProfile");
                var env = await CoreWebView2Environment.CreateAsync(null, userDataDir);
                await _webView.EnsureCoreWebView2Async(env);

                // 配置 WebView2 运行环境与右键上下文菜单禁用
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

                // 注册 Web 消息监听事件
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 定位前端 HTML 物理文件路径
                string appDir = Tool.GetAppDirectory();
                string htmlPath = Path.Combine(appDir, "Resources", "component_match_dialog.html");

                // 若发布目录不存在，回退查找本地项目源码目录
                if (!File.Exists(htmlPath))
                {
                    htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "component_match_dialog.html");
                }

                if (File.Exists(htmlPath))
                {
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
                }
                else
                {
                    _webView.CoreWebView2.NavigateToString("<h3 style='color:red;padding:20px;'>未找到前端页面模板 component_match_dialog.html</h3>");
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ComponentMatchForm 初始化异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 集中处理来自 Vue 3 前端界面的 postMessage 指令
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string rawJson = e.TryGetWebMessageAsString();
                if (string.IsNullOrWhiteSpace(rawJson)) return;

                using var doc = JsonDocument.Parse(rawJson);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;

                string action = actionProp.GetString() ?? string.Empty;

                switch (action)
                {
                    // 1. 前端加载就绪，下发品牌列表与当前保存的配置
                    case "ready":
                        var brands = _controller.GetBrandStats();
                        var config = _controller.LoadConfig();
                        PostMessageToWeb(new
                        {
                            action = "initData",
                            brands,
                            config
                        });
                        break;

                    // 2. 保存配置
                    case "saveConfig":
                        if (root.TryGetProperty("config", out var saveCfgProp))
                        {
                            var saveCfg = JsonSerializer.Deserialize<ComponentMatchFilterConfig>(saveCfgProp.GetRawText(), JsonOptions);
                            if (saveCfg != null)
                            {
                                _controller.SaveConfig(saveCfg);
                                PostMessageToWeb(new { action = "saveSuccess" });
                            }
                        }
                        break;

                    // 3. 运行单条模拟测试
                    case "testMatch":
                        string tName = root.TryGetProperty("name", out var np) ? np.GetString() ?? "" : "";
                        string tCur = root.TryGetProperty("current", out var cp) ? cp.GetString() ?? "" : "";
                        string tPole = root.TryGetProperty("pole", out var pp) ? pp.GetString() ?? "" : "";
                        string tTrip = root.TryGetProperty("tripMode", out var tp) ? tp.GetString() ?? "" : "";
                        string tBrand = root.TryGetProperty("brand", out var bp) ? bp.GetString() ?? "" : "";

                        List<MustContainRule>? rules = null;
                        if (root.TryGetProperty("rules", out var rp))
                        {
                            rules = JsonSerializer.Deserialize<List<MustContainRule>>(rp.GetRawText(), JsonOptions);
                        }

                        var sw = Stopwatch.StartNew();
                        var testItems = _controller.TestMatch(tName, tCur, tPole, tTrip, tBrand, rules);
                        sw.Stop();

                        PostMessageToWeb(new
                        {
                            action = "testResult",
                            items = testItems,
                            duration = sw.ElapsedMilliseconds
                        });
                        break;

                    // 4. 立即执行当前选区批量识别反查
                    case "executeBatch":
                        ComponentMatchFilterConfig? batchCfg = null;
                        if (root.TryGetProperty("config", out var bCfgProp))
                        {
                            batchCfg = JsonSerializer.Deserialize<ComponentMatchFilterConfig>(bCfgProp.GetRawText(), JsonOptions);
                        }
                        var res = _controller.ExecuteBatch(batchCfg);
                        PostMessageToWeb(new
                        {
                            action = "batchResult",
                            result = res
                        });
                        break;

                    // 5. 窗口拖拽
                    case "dragWindow":
                        SafeInvoke(() =>
                        {
                            ReleaseCapture();
                            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        });
                        break;

                    // 6. 关闭窗口
                    case "closeWindow":
                        SafeInvoke(this.Close);
                        break;

                    // 7. 切换折叠/展开状态并动态调整宿主窗口高度
                    case "setCollapseState":
                        bool isCol = root.TryGetProperty("collapsed", out var colProp) && colProp.GetBoolean();
                        SafeInvoke(() =>
                        {
                            this.ClientSize = new Size(500, isCol ? 52 : 620);
                        });
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ComponentMatchForm 消息分发异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 安全向 WebView2 前端推送 JSON 消息
        /// </summary>
        private void PostMessageToWeb(object data)
        {
            SafeInvoke(() =>
            {
                if (_webView?.CoreWebView2 != null)
                {
                    string json = JsonSerializer.Serialize(data, JsonOptions);
                    _webView.CoreWebView2.PostWebMessageAsString(json);
                }
            });
        }

        /// <summary>
        /// 跨线程安全调度至 UI 主线程
        /// </summary>
        private void SafeInvoke(Action action)
        {
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
