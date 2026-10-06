using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 元器件拆分改型微型交互工作台窗体
    /// 基于 WebView2 + Vue 3 + Element Plus，主色调 #009688
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，Excel 操作收敛于 ExcelServices
    /// </summary>
    public class ComponentSplitForm : Form
    {
        // 单例窗体实例引用
        private static ComponentSplitForm? _instance;

        // WebView2 浏览器控件
        private readonly WebView2 _webView;

        // 待拆分元器件候选信息上下文
        private ComponentSplitCandidateDto _candidate;

        // 标记前端页面是否加载就绪
        private bool _isWebReady = false;

        // JSON 序列化配置选项 (配置小驼峰命名与大小写兼容)
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };

        // Win32 API 辅助窗口无边框拖拽移动
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int WM_NCLBUTTONDOWN = 0xA1; // --硬编码: Win32消息代码--
        private const int HTCAPTION = 0x2; // --硬编码: 标题栏区域标识--

        /// <summary>
        /// 私有构造函数：初始化窗体与 WebView2 控件
        /// </summary>
        /// <param name="candidate">待拆分元器件上下文</param>
        private ComponentSplitForm(ComponentSplitCandidateDto candidate)
        {
            _candidate = candidate ?? new ComponentSplitCandidateDto();

            // 实例化并装配 WebView2 控件
            _webView = new WebView2();
            _webView.Dock = DockStyle.Fill;
            this.Controls.Add(_webView);

            // 配置窗体规格与外观 (宽 760px，高 560px，居中显示) --硬编码: 窗体规格--
            this.Text = "✂️ 元器件拆分与改型工作台 (Split & Modify)"; // --硬编码: 窗体标题--
            this.Size = new Size(760, 560);
            this.MinimumSize = new Size(680, 480);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.ShowInTaskbar = false;
            this.TopMost = true;

            // 绑定窗体生命周期事件
            this.Load += OnFormLoadAsync;
            this.FormClosed += (s, e) => { _instance = null; };
        }

        /// <summary>
        /// 异步初始化 WebView2 并加载前端 HTML
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 1. 获取系统本地 LocalApplicationData 专有缓存目录，避免污染用户业务配置或网盘同步
                string userDataDir = Tool.GetWebView2UserDataFolder("WebView2_ComponentSplit");
                // 2. 异步创建 CoreWebView2 运行环境
                var env = await CoreWebView2Environment.CreateAsync(null, userDataDir);
                // 3. 确保 WebView2 核心就绪
                await _webView.EnsureCoreWebView2Async(env);

                if (_webView.CoreWebView2 != null)
                {
                    // 禁用浏览器默认右键菜单与状态栏
                    _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                    _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
                    _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                    // 注册 Web 消息接收事件
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                    // 寻找前端 HTML 资源物理路径
                    string htmlPath = FindHtmlResourcePath("component_split.html");
                    if (File.Exists(htmlPath))
                    {
                        string resDir = Path.GetDirectoryName(htmlPath)!;
                        // 将本地目录映射为虚拟主机域名 https://appassets.local (规避沙箱拦截)
                        _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                            "appassets.local",
                            resDir,
                            CoreWebView2HostResourceAccessKind.Allow);

                        // 导航加载页面 (带动态时间戳避免缓存)
                        string freshUrl = $"https://appassets.local/component_split.html?v={DateTime.Now.Ticks}";
                        _webView.Source = new Uri(freshUrl);
                    }
                    else
                    {
                        MessageBox.Show($"未找到界面资源文件: {htmlPath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); // --硬编码: 错误提示--
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录初始化异常日志
                LogHelper.WriteLog($"[ComponentSplitForm] 初始化异常: {ex.Message}");
                MessageBox.Show($"初始化拆分窗口失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); // --硬编码: 错误提示--
            }
        }

        /// <summary>
        /// 集中处理来自 Vue 3 前端的 postMessage 异步指令 (双轨容错解析)
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
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

                string action = actionProp.GetString() ?? string.Empty;

                switch (action)
                {
                    // 1. 前端加载就绪，主动回推候选元器件初始数据
                    case "ready":
                        _isWebReady = true;
                        PushDataToWeb("initCandidate", _candidate);
                        break;

                    // 2. 用户在列表中点击单项 Handle：联动向 CAD 发送对焦选中指令
                    case "selectHandle":
                        if (root.TryGetProperty("handle", out var hProp))
                        {
                            string handleVal = hProp.GetString()?.Trim() ?? string.Empty;
                            if (!string.IsNullOrEmpty(handleVal))
                            {
                                // 带有 50ms 防抖与自动缩放对焦推送到 CAD
                                CadSyncClient.SendHandlesDebounced(new List<string> { handleVal }, true);
                            }
                        }
                        break;

                    // 3. 关键字模糊搜索：查询本地 SQLite 个人物料库以供联想输入
                    case "searchSuggestions":
                        string kw = string.Empty;
                        if (root.TryGetProperty("keyword", out var kwProp))
                        {
                            kw = kwProp.GetString()?.Trim() ?? string.Empty;
                        }

                        // 后台线程检索避免阻塞 UI
                        Task.Run(() =>
                        {
                            var list = PersonalComponentDbService.QueryComponentSuggestions(kw, 15);
                            SafeInvoke(() =>
                            {
                                PushDataToWeb("suggestionsResult", new
                                {
                                    keyword = kw,
                                    items = list
                                });
                            });
                        });
                        break;

                    // 4. 用户点击【确认拆分并生成新行】
                    case "submitSplit":
                        if (root.TryGetProperty("data", out var dataProp))
                        {
                            var req = JsonSerializer.Deserialize<ComponentSplitSubmitRequest>(dataProp.GetRawText(), JsonOptions);
                            if (req != null)
                            {
                                // 调度 Excel 业务公共服务执行拆分
                                var res = ExcelServices.ExecuteComponentSplit(req);
                                PushDataToWeb("splitResult", res);

                                // 若执行成功，延迟 1.2 秒后平滑关闭窗体
                                if (res.Success)
                                {
                                    Task.Delay(1200).ContinueWith(_ =>
                                    {
                                        SafeInvoke(() =>
                                        {
                                            this.Close();
                                        });
                                    });
                                }
                            }
                        }
                        break;

                    // 5. 关闭窗体
                    case "close":
                        this.Close();
                        break;

                    // 6. 窗口标题拖拽
                    case "dragWindow":
                        ReleaseCapture();
                        SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        break;
                }
            }
            catch (Exception ex)
            {
                // 记录消息处理异常日志
                LogHelper.WriteLog($"[ComponentSplitForm] 消息处理异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 线程安全地向前端 WebView2 派发业务事件与数据实体
        /// </summary>
        /// <param name="action">事件名称</param>
        /// <param name="data">数据载荷对象</param>
        private void PushDataToWeb(string action, object data)
        {
            SafeInvoke(() =>
            {
                if (!_isWebReady || _webView?.CoreWebView2 == null) return;
                try
                {
                    string json = JsonSerializer.Serialize(new
                    {
                        action = action,
                        data = data
                    }, JsonOptions);
                    _webView.CoreWebView2.PostWebMessageAsString(json);
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[ComponentSplitForm] PushDataToWeb 异常: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 线程安全调度辅助方法
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

        /// <summary>
        /// 多重备用检索 HTML 物理路径
        /// </summary>
        private string FindHtmlResourcePath(string fileName)
        {
            string appDir = Tool.GetAppDirectory();
            string p1 = Path.Combine(appDir, "Resources", fileName);
            if (File.Exists(p1)) return p1;

            string p2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", fileName);
            if (File.Exists(p2)) return p2;

            string p3 = Path.Combine(appDir, fileName);
            if (File.Exists(p3)) return p3;

            return p1;
        }

        /// <summary>
        /// 外部统一公开的唤起入口：在 Excel 主线程模态居中弹出工作台
        /// </summary>
        /// <param name="candidate">待处理的元器件上下文</param>
        public static void ShowForm(ComponentSplitCandidateDto candidate)
        {
            try
            {
                // 若窗体已存在且可见，直接激活置顶
                if (_instance != null && !_instance.IsDisposed && _instance.IsHandleCreated)
                {
                    _instance.Activate();
                    return;
                }

                // 实例化新工作台
                _instance = new ComponentSplitForm(candidate);

                // 获取当前正在运行的 Excel 视口句柄
                IntPtr excelHwnd = ExcelDnaUtil.WindowHandle;
                if (excelHwnd != IntPtr.Zero)
                {
                    // 以 Excel 主视口为 Owner 模态居中打开
                    _instance.ShowDialog(new ExcelServices.ExcelWin32Window(excelHwnd));
                }
                else
                {
                    // 降级以默认模态显示
                    _instance.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                // 记录唤起异常日志
                LogHelper.WriteLog($"[ComponentSplitForm] ShowForm 异常: {ex.Message}");
            }
        }
    }
}
