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
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 消息接口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

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

                    // 禁用原生开发者工具（生产环境防误触）
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
        /// </summary>
        private void LoadHtmlPage()
        {
            try
            {
                // 获取基准路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // 候选路径1: 运行根目录下的 Resources/cloud_solution.html
                string path1 = Path.Combine(baseDir, "Resources", "cloud_solution.html");
                // 候选路径2: 源码工程目录下的 Resources/cloud_solution.html
                string path2 = Path.Combine(baseDir, "..", "..", "Resources", "cloud_solution.html");

                // 优先检查候选路径1
                string targetHtml = File.Exists(path1) ? path1 : (File.Exists(path2) ? path2 : "");

                if (!string.IsNullOrEmpty(targetHtml))
                {
                    // 导航加载本地 HTML 文件
                    _webView.CoreWebView2.Navigate(new Uri(Path.GetFullPath(targetHtml)).AbsoluteUri);
                }
                else
                {
                    // 未找到物理文件提示
                    LogHelper.WriteLog($"[CloudSolutionForm] 未找到 cloud_solution.html 页面文件，查找路径: {path1}");
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
                    // 1. 无边框拖拽移动窗体
                    case "dragWindow":
                        if (this.WindowState == FormWindowState.Normal)
                        {
                            ReleaseCapture();
                            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
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
    }
}
