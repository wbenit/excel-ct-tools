using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 + Element Plus 的“人工清单分类选择”向导窗口
    /// 遵循主色调 #009688、弹性布局、无水平滚动条与安全拖拽解耦规范
    /// </summary>
    public class LaborListForm : Form
    {
        // 声明 WebView2 浏览器核心控件
        private readonly WebView2 _webView;

        // 声明人工清单业务控制器实例
        private readonly LaborListController _controller;

        // 导入 Windows 原生 user32.dll 接口用于支持原生无边框拖拽
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 原生消息接口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 导入 GetAsyncKeyState 接口检测物理按键状态 (防范幽灵鼠标捕获死锁)
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int VK_LBUTTON = 0x01;
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 通用 JSON 序列化配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// 构造函数: 初始化控制器、WebView2 控件及窗口外观
        /// </summary>
        public LaborListForm()
        {
            // 实例化人工清单控制器 (注入当前窗体引用)
            _controller = new LaborListController(this);

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 配置 Form 窗体尺寸与外观
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸 (480x430 像素，弹性自适应)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 窗体标题 --硬编码: 窗体标题--
            this.Text = "导出人工清单 - 选择分类表";
            // 设定适宜分类列表展示的高宽尺寸
            this.ClientSize = new Size(480, 430);
            // 屏幕居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;
            // 设为无边框现代扁平样式
            this.FormBorderStyle = FormBorderStyle.None;
            // 禁用最大化与最小化
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            // 背景填充纯白
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件布局与加载监听
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 满框铺满 Form
            _webView.Dock = DockStyle.Fill;
            // 挂载至 Controls 控件集
            this.Controls.Add(_webView);
            // 绑定 Form Load 异步加载监听
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载触发的异步初始化逻辑
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 获取用户自定义数据目录下的专有 WebView2 用户缓存目录
                string userDataDir = Tool.GetCustomDataDirectoryFromGlobalConfig();
                if (string.IsNullOrWhiteSpace(userDataDir))
                {
                    userDataDir = Tool.GetAppDataDirectory();
                }
                string webViewCacheDir = Path.Combine(userDataDir, "WebView2_LaborList");

                // 异步创建 WebView2 运行环境
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, webViewCacheDir);

                // 确保 CoreWebView2 核心对象就绪
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 注册前端 postMessage 通信监听
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 寻找目标 HTML 静态资源物理路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string appDir = Tool.GetAppDirectory();

                // 配置多级备选路径集以实现全域高容错加载 (开发目录、输出目录、发布目录与根目录)
                string[] candidatePaths = new string[]
                {
                    Path.Combine(appDir, "Resources", "labor_list.html"),
                    Path.Combine(appDir, "publish", "Resources", "labor_list.html"),
                    Path.Combine(appDir, "labor_list.html"),
                    Path.Combine(baseDir, "Resources", "labor_list.html"),
                    Path.Combine(baseDir, "labor_list.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "labor_list.html"),
                    // 向上探测源码根目录 (bin/Debug/net48 -> 项目源码目录)
                    Path.Combine(appDir, "..", "..", "..", "Resources", "labor_list.html"),
                    Path.Combine(appDir, "..", "..", "..", "publish", "Resources", "labor_list.html"),
                    // 工程物理路径绝对兜底 --硬编码: 工程物理兜底路径--
                    @"E:\Ace\ExcelAddInCTtools\Resources\labor_list.html",
                    @"E:\Ace\ExcelAddInCTtools\publish\Resources\labor_list.html"
                };

                // 循环检索首个真实存在的目标 HTML 文件
                string htmlPath = string.Empty;
                foreach (string candidate in candidatePaths)
                {
                    if (File.Exists(candidate))
                    {
                        htmlPath = candidate;
                        break;
                    }
                }

                // 导航至前端 Vue 3 页面
                if (!string.IsNullOrWhiteSpace(htmlPath))
                {
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
                }
                else
                {
                    MessageBox.Show("未找到人工清单向导界面资源文件: labor_list.html", "资源缺失", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"LaborListForm 初始化失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 接收前端 Vue 3 postMessage 交互指令的核心路由器
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 解析前端传递的 JSON 消息体
                string jsonString = e.WebMessageAsJson;
                using var jsonDoc = JsonDocument.Parse(jsonString);
                var root = jsonDoc.RootElement;

                // 读取前端 action 动作标识
                if (!root.TryGetProperty("action", out var actionProp)) return;
                string action = actionProp.GetString() ?? string.Empty;

                switch (action)
                {
                    // 1. 前端请求加载初始化数据 (分类列表、项目名称及已有表状态)
                    case "getInitData":
                        var initData = _controller.GetInitData();
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "initDataLoaded",
                            projectName = initData.ProjectName,
                            categories = initData.Categories,
                            hasExistingLaborList = initData.HasExistingLaborList
                        }, JsonOptions));
                        break;

                    // 2. 前端提交勾选的分类列表并开始导出人工清单
                    case "exportLaborList":
                        var selectedCategories = new System.Collections.Generic.List<string>();
                        if (root.TryGetProperty("selectedCategories", out var catsProp) && catsProp.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var item in catsProp.EnumerateArray())
                            {
                                string cName = item.GetString() ?? "";
                                if (!string.IsNullOrWhiteSpace(cName))
                                {
                                    selectedCategories.Add(cName);
                                }
                            }
                        }

                        // 调用控制器执行导出操作 (内部纯静默，无任何 Win32 阻塞)
                        var result = _controller.ExportLaborList(selectedCategories);

                        // 将导出结果纯 JSON 安全回传给前端 Vue 3
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "exportCompleted",
                            success = result.Success,
                            message = result.Message,
                            cabinetCount = result.CabinetCount,
                            totalAmount = result.TotalAmount
                        }, JsonOptions));
                        break;

                    // 3. 原生无边框拖拽指令 (防御幽灵鼠标捕获)
                    case "dragWindow":
                        BeginInvoke(new Action(() =>
                        {
                            try
                            {
                                // 只有当物理左键真实处于按下状态时，才触发系统标题栏拖拽
                                if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0)
                                {
                                    ReleaseCapture();
                                    SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                                }
                            }
                            catch (Exception exDrag)
                            {
                                LogHelper.WriteLog($"原生拖拽容错: {exDrag.Message}");
                            }
                        }));
                        break;

                    // 4. 关闭向导窗口
                    case "closeWindow":
                        BeginInvoke(new Action(() =>
                        {
                            this.Close();
                        }));
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"LaborListForm 处理前端消息异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 跨线程安全向前端 WebView2 投递 JSON 消息
        /// </summary>
        /// <param name="json">待发送的序列化 JSON 文本</param>
        public void PostWebMessageSafe(string json)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;

            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => PostWebMessageSafe(json)));
                return;
            }

            try
            {
                if (_webView?.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.PostWebMessageAsJson(json);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"向 WebView2 发送消息失败: {ex.Message}");
            }
        }
    }
}
