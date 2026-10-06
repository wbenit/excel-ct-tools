using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 添加元器件到本地个人库核对确认窗体 (基于 WebView2 + Vue 3 + Element Plus，主色调 #009688)
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，硬编码显式标明
    /// </summary>
    public class AddToPersonalDbForm : Form
    {
        // 全局单例句柄，防止多开窗口产生冲突
        private static AddToPersonalDbForm? _instance;

        // WebView2 控件句柄
        private readonly WebView2 _webView;

        // 缓存的待传递候选元器件列表
        private List<AddToPersonalDbCandidateItem> _candidates;

        // 前端是否已加载就绪
        private bool _isWebReady = false;

        // JSON 序列化配置 (小驼峰命名)
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// 私有构造函数：配置标准窗体与 WebView2
        /// </summary>
        private AddToPersonalDbForm(List<AddToPersonalDbCandidateItem> candidates)
        {
            _candidates = candidates ?? new List<AddToPersonalDbCandidateItem>();

            // 初始化 WebView2 控件
            _webView = new WebView2();
            _webView.DefaultBackgroundColor = Color.White;
            _webView.Dock = DockStyle.Fill;
            this.Controls.Add(_webView);

            // 配置窗体规格与外观 (宽 900px，高 580px，支持居中) --硬编码: 窗体规格--
            this.Text = "📥 添加元器件到本地个人库 (核对确认)"; // --硬编码: 窗体标题--
            this.Size = new Size(900, 580);
            this.MinimumSize = new Size(780, 480);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.ShowInTaskbar = false;
            this.TopMost = true;

            // 订阅窗体加载与销毁事件
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
                // 1. 配置独立的系统本地用户缓存目录，避免污染用户共享配置目录与网盘同步
                string userDataDir = Tool.GetWebView2UserDataFolder("WebView2_AddToPersonalDb");
                // 2. 异步创建 CoreWebView2 运行环境
                var env = await CoreWebView2Environment.CreateAsync(null, userDataDir);
                // 3. 确保 WebView2 核心就绪
                await _webView.EnsureCoreWebView2Async(env);

                if (_webView.CoreWebView2 != null)
                {
                    // 禁用默认右键菜单与状态栏
                    _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                    _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
                    _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                    // 注册 Web 消息接收事件
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                    // 寻找前端 HTML 资源物理路径
                    string htmlPath = FindHtmlResourcePath("add_to_personal_db.html");
                    if (File.Exists(htmlPath))
                    {
                        string resDir = Path.GetDirectoryName(htmlPath)!;
                        // 将本地目录映射为虚拟主机域名 https://appassets.local (规避沙箱拦截)
                        _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                            "appassets.local",
                            resDir,
                            CoreWebView2HostResourceAccessKind.Allow);

                        // 导航加载页面 (带动态时间戳避免缓存)
                        string freshUrl = $"https://appassets.local/add_to_personal_db.html?v={DateTime.Now.Ticks}";
                        _webView.Source = new Uri(freshUrl);
                    }
                    else
                    {
                        MessageBox.Show($"未找到界面资源文件: {htmlPath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); // --硬编码: 提示语--
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[AddToPersonalDbForm] 初始化异常: {ex.Message}");
                MessageBox.Show($"初始化核对窗口失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); // --硬编码: 提示语--
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
                    // 1. 前端加载就绪，主动回推候选数据
                    case "ready":
                        _isWebReady = true;
                        PushDataToWeb("initCandidates", new
                        {
                            candidates = _candidates,
                            totalCount = _candidates.Count
                        });
                        break;

                    // 2. 用户点击确认执行入库
                    case "submitSave":
                        // 提取提交数据与覆盖更新设置
                        bool overwritePrice = true;
                        if (root.TryGetProperty("overwritePrice", out var owProp))
                        {
                            overwritePrice = owProp.GetBoolean();
                        }

                        // 提取更新后的元器件列表
                        var submitItems = new List<AddToPersonalDbCandidateItem>();
                        if (root.TryGetProperty("items", out var itemsProp) && itemsProp.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var el in itemsProp.EnumerateArray())
                            {
                                var item = JsonSerializer.Deserialize<AddToPersonalDbCandidateItem>(el.GetRawText(), JsonOptions);
                                if (item != null)
                                {
                                    submitItems.Add(item);
                                }
                            }
                        }

                        // 在独立后台线程中执行 SQLite 事务入库 (规避阻塞 Chromium IPC)
                        System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                        {
                            var saveResult = PersonalComponentDbService.BatchSaveOrUpdatePersonalComponents(submitItems, overwritePrice);
                            // 线程安全切回主线程通知前端展示结果
                            SafeInvoke(() =>
                            {
                                PushDataToWeb("saveCompleted", saveResult);
                            });
                        });
                        break;

                    // 3. 用户点击取消或关闭窗口
                    case "closeForm":
                        SafeInvoke(this.Close);
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[AddToPersonalDbForm] 消息处理异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 安全向 Vue 3 前端推送 JSON 消息
        /// </summary>
        private void PushDataToWeb(string action, object data)
        {
            SafeInvoke(() =>
            {
                try
                {
                    if (_webView?.CoreWebView2 != null)
                    {
                        var packet = new Dictionary<string, object>
                        {
                            { "action", action },
                            { "data", data }
                        };
                        string json = JsonSerializer.Serialize(packet, JsonOptions);
                        _webView.CoreWebView2.PostWebMessageAsString(json);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[AddToPersonalDbForm] PushDataToWeb 异常: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 线程安全调用辅助
        /// </summary>
        private void SafeInvoke(Action action)
        {
            if (this.IsDisposed) return;
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
        /// 智能寻找 HTML 资源物理全路径 (双轨检索)
        /// </summary>
        private static string FindHtmlResourcePath(string fileName)
        {
            string appDir = Tool.GetAppDirectory();
            string path1 = Path.Combine(appDir, "Resources", fileName);
            if (File.Exists(path1)) return path1;

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string path2 = Path.Combine(baseDir, "Resources", fileName);
            if (File.Exists(path2)) return path2;

            return path1;
        }

        /// <summary>
        /// 公共静态入口：在 Excel 主窗口正中央弹出核对确认窗体
        /// </summary>
        public static void ShowForm(List<AddToPersonalDbCandidateItem> candidates)
        {
            try
            {
                // 若已打开旧窗体，先激活带至前台
                if (_instance != null && !_instance.IsDisposed)
                {
                    _instance._candidates = candidates;
                    _instance.BringToFront();
                    _instance.Activate();
                    if (_instance._isWebReady)
                    {
                        _instance.PushDataToWeb("initCandidates", new
                        {
                            candidates = candidates,
                            totalCount = candidates.Count
                        });
                    }
                    return;
                }

                // 实例化新窗体
                _instance = new AddToPersonalDbForm(candidates);

                // 获取 Excel 顶层 Win32 视口窗口句柄进行模态置顶绑定
                IntPtr excelHwnd = ExcelDnaUtil.WindowHandle;
                if (excelHwnd != IntPtr.Zero)
                {
                    _instance.Show(new ExcelServices.ExcelWin32Window(excelHwnd));
                }
                else
                {
                    _instance.Show();
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[AddToPersonalDbForm] ShowForm 异常: {ex.Message}");
            }
        }
    }
}
