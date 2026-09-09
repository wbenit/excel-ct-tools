using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Controllers;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“云方案中心”无边框置顶宿主窗体
    /// 支持 4 列响应式网格布局、多图纸预览画廊与 BOM 动态回路倍增
    /// </summary>
    public class CloudSolutionForm : Form
    {
        // 声明 WebView2 浏览器控件
        private readonly WebView2 _webView;

        // 声明云方案业务控制器
        private readonly CloudSolutionController _controller;

        // 导入 Windows 原生 user32.dll 接口以支持无边框窗体拖拽
        [System.Runtime.InteropServices.DllImport("user32.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Winapi)]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 消息接口
        [System.Runtime.InteropServices.DllImport("user32.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Winapi)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 导入 GetAsyncKeyState 检测鼠标物理按键状态，防止幽灵拖拽死锁
        [System.Runtime.InteropServices.DllImport("user32.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Winapi)]
        private static extern short GetAsyncKeyState(int vKey);

        // Win32 常量: 鼠标左键虚拟键码
        private const int VK_LBUTTON = 0x01;

        // Win32 常量: 标题栏拖拽消息标识
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 通用 JSON 序列化设置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        /// <summary>
        /// 构造函数: 初始化控制器与 WebView2 控件属性
        /// </summary>
        public CloudSolutionForm()
        {
            // 实例化云方案业务控制器
            _controller = new CloudSolutionController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 设置 Form 基本尺寸与几何样式 (1180x820 像素，从容容纳 4 列卡片与图3详情)
            InitializeFormProperties();

            // 初始化并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 初始化窗体外观、尺寸与窗口属性
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题
            this.Text = "鑫壬云方案中心";

            // 尺寸设置为 1180x820 像素，支持 4 列卡片舒展显示
            this.ClientSize = new Size(1180, 820);

            // 居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 无边框工业现代样式
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用系统原生最大/最小化按钮
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 背景色设置为白色
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件及核心事件回调
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 设置 WebView2 填满整个窗体
            _webView.Dock = DockStyle.Fill;

            // 添加到窗体控件树中
            this.Controls.Add(_webView);

            // 绑定窗口加载完成后初始化 CoreWebView2 事件
            this.Load += async (s, e) =>
            {
                try
                {
                    // 设置独立的本地用户缓存数据目录
                    string userDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "ExcelAddInDemo",
                        "WebView2_CloudSolution"
                    );

                    // 创建 CoreWebView2 运行环境
                    var env = await CoreWebView2Environment.CreateAsync(null, userDir);

                    // 异步初始化核心引擎
                    await _webView.EnsureCoreWebView2Async(env);

                    // 禁用内置右键菜单，保障界面整洁纯净
                    _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;

                    // 启用原生开发者工具（便于排查前端渲染）
                    _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                    // 绑定接收前端发来的 WebMessage 消息路由
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                    // 加载前端 HTML 页面
                    LoadHtmlPage();
                }
                catch (Exception ex)
                {
                    // 记录 WebView2 初始化异常
                    LogHelper.WriteLog($"[CloudSolutionForm] 初始化 WebView2 异常: {ex.Message}");
                    MessageBox.Show($"初始化方案中心界面失败: {ex.Message}", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
        }

        /// <summary>
        /// 寻址并导航加载 cloud_solution.html 前端页面
        /// 支持多重路径降级检测，避免在调试或独立部署路径下白屏
        /// </summary>
        private void LoadHtmlPage()
        {
            try
            {
                // 获取插件运行根目录
                string appDir = Tool.GetAppDirectory();
                // 获取当前应用域基准目录
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // 构造多级备选查找路径集合
                string[] candidatePaths = new[]
                {
                    Path.Combine(appDir, "Resources", "cloud_solution.html"),
                    Path.Combine(appDir, "publish", "Resources", "cloud_solution.html"),
                    Path.Combine(baseDir, "Resources", "cloud_solution.html"),
                    Path.Combine(baseDir, "publish", "Resources", "cloud_solution.html"),
                    Path.Combine(baseDir, "..", "..", "Resources", "cloud_solution.html"),
                    Path.Combine(baseDir, "..", "..", "..", "Resources", "cloud_solution.html"),
                    @"d:\code\excel-ct-tools\Resources\cloud_solution.html" // --硬编码: 本地开发源码目录兜底--
                };

                // 遍历寻找首个物理存在的 HTML 页面
                string targetHtml = string.Empty;
                foreach (var candidate in candidatePaths)
                {
                    if (File.Exists(candidate))
                    {
                        targetHtml = Path.GetFullPath(candidate);
                        break;
                    }
                }

                // 若找到有效 HTML 物理文件则执行导航
                if (!string.IsNullOrEmpty(targetHtml))
                {
                    // 导航加载本地 HTML 文件
                    _webView.CoreWebView2.Navigate(new Uri(targetHtml).AbsoluteUri);
                }
                else
                {
                    // 未找到物理文件时，在 WebView2 内部呈现友好错误页，彻底杜绝静默空白白板
                    string notFoundHtml = "<div style='font-family:Segoe UI,sans-serif;padding:30px;color:#dc2626;'>" +
                                          "<h2>⚠️ 未找到云方案页面文件 (cloud_solution.html)</h2>" +
                                          "<p>请确保 <code>Resources/cloud_solution.html</code> 存在并已复制至输出目录。</p>" +
                                          "<p>检索路径列表:<ul>" +
                                          string.Join("", Array.ConvertAll(candidatePaths, p => $"<li>{p}</li>")) +
                                          "</ul></p></div>";
                    _webView.CoreWebView2.NavigateToString(notFoundHtml);
                    // 记录未找到页面日志
                    LogHelper.WriteLog($"[CloudSolutionForm] 未找到 cloud_solution.html 页面文件，已呈现错误提示。");
                }
            }
            catch (Exception ex)
            {
                // 记录加载失败异常
                LogHelper.WriteLog($"[CloudSolutionForm] 加载 HTML 文件失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 前端 IPC 消息接收分发中心
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 读取收到的原始 JSON 文本
                string json = e.WebMessageAsJson;
                if (string.IsNullOrWhiteSpace(json)) return;

                // 解析为 JsonDocument 提取 action 字段
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;

                string action = actionProp.GetString() ?? "";

                switch (action)
                {
                    // 1. 无边框拖拽移动窗体 (带物理鼠标状态检测，杜绝幽灵捕获死锁)
                    case "dragWindow":
                        if (this.WindowState == FormWindowState.Normal)
                        {
                            // 检测物理按键是否仍在按下状态
                            if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0)
                            {
                                ReleaseCapture();
                                SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                            }
                        }
                        break;

                    // 2. 关闭窗体
                    case "close":
                        this.Invoke(new Action(() => this.Close()));
                        break;

                    // 3. 最小化窗体
                    case "minimize":
                        this.Invoke(new Action(() => this.WindowState = FormWindowState.Minimized));
                        break;

                    // 4. 最大化与还原切换
                    case "toggleMaximize":
                        this.Invoke(new Action(() =>
                        {
                            this.WindowState = (this.WindowState == FormWindowState.Maximized)
                                ? FormWindowState.Normal
                                : FormWindowState.Maximized;
                        }));
                        break;

                    // 5. 分页多维检索方案列表
                    case "querySchemes":
                        string queryPayload = root.TryGetProperty("data", out var qProp) ? qProp.GetRawText() : "{}";
                        var result = _controller.QuerySchemes(queryPayload);
                        PostMessageSafe("querySchemesResult", result);
                        break;

                    // 6. 获取方案详情 (多图纸与 BOM)
                    case "getSchemeDetail":
                        string schemeId = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
                        var detail = _controller.GetSchemeDetail(schemeId);
                        PostMessageSafe("getSchemeDetailResult", detail);
                        break;

                    // 7. 切换收藏状态
                    case "toggleFavorite":
                        string favId = root.TryGetProperty("id", out var fIdProp) ? fIdProp.GetString() ?? "" : "";
                        bool isFav = _controller.ToggleFavorite(favId);
                        PostMessageSafe("toggleFavoriteResult", new { id = favId, isFavorite = isFav });
                        break;

                    // 8. 保存企业方案
                    case "saveEnterpriseScheme":
                        string schemePayload = root.TryGetProperty("data", out var sProp) ? sProp.GetRawText() : "{}";
                        bool saveOk = _controller.SaveEnterpriseScheme(schemePayload);
                        PostMessageSafe("saveEnterpriseSchemeResult", new { success = saveOk });
                        break;

                    // 9. 删除企业方案
                    case "deleteEnterpriseScheme":
                        string delId = root.TryGetProperty("id", out var dIdProp) ? dIdProp.GetString() ?? "" : "";
                        bool delOk = _controller.DeleteEnterpriseScheme(delId);
                        PostMessageSafe("deleteEnterpriseSchemeResult", new { success = delOk, id = delId });
                        break;

                    // 10. 插入方案到 Excel 活动工作表
                    case "insertSchemeToExcel":
                        string insertPayload = root.TryGetProperty("data", out var iProp) ? iProp.GetRawText() : "{}";
                        var (insertOk, insertMsg) = _controller.InsertSchemeToExcel(insertPayload);
                        PostMessageSafe("insertSchemeToExcelResult", new { success = insertOk, message = insertMsg });
                        break;
                }
            }
            catch (Exception ex)
            {
                // 记录消息分发错误
                LogHelper.WriteLog($"[CloudSolutionForm] 消息处理异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 线程安全地向前端 WebView2 发送 JSON 格式的回调消息
        /// </summary>
        private void PostMessageSafe(string action, object? data)
        {
            if (this.IsDisposed || _webView.IsDisposed || _webView.CoreWebView2 == null) return;

            // 切回 UI 主线程发送消息
            this.BeginInvoke(new Action(() =>
            {
                try
                {
                    // 封装统一的回调信封
                    var envelope = new
                    {
                        action = action,
                        data = data
                    };

                    // 序列化
                    string json = JsonSerializer.Serialize(envelope, JsonOptions);
                    // 投递给 JavaScript 端
                    _webView.CoreWebView2.PostWebMessageAsJson(json);
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[CloudSolutionForm] 投递消息失败: {ex.Message}");
                }
            }));
        }

        /// <summary>
        /// 窗体关闭时解绑 WebMessage 事件并释放 WebView2 控件资源
        /// 杜绝后台 Chromium 子进程僵死残留
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 执行基类关闭生命周期逻辑
            base.OnFormClosing(e);

            try
            {
                // 安全解绑前后端消息接收事件处理器
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
                }

                // 显式释放并销毁 WebView2 控件
                _webView.Dispose();
            }
            catch (Exception ex)
            {
                // 记录释放异常日志
                LogHelper.WriteLog($"[CloudSolutionForm] 释放 WebView2 异常: {ex.Message}");
            }
        }
    }
}
