using System;
using System.Collections.Generic;
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
    /// 基于 WebView2 + Vue 3 + Element Plus 的“三箱型号添写向导”无边框窗体
    /// 遵循主色调 #009688、弹性响应式布局、无水平滚动条与安全拖拽解耦规范
    /// </summary>
    public class CabinetModelPipelineForm : Form
    {
        // 声明 WebView2 浏览器核心控件
        private readonly WebView2 _webView;

        // 声明三箱型号管道业务控制器实例
        private readonly CabinetModelPipelineController _controller;

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
        public CabinetModelPipelineForm()
        {
            // 实例化业务控制器 (注入当前窗体引用)
            _controller = new CabinetModelPipelineController(this);

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 配置 Form 窗体尺寸与外观
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸 (1020x660 像素，弹性自适应)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 窗体标题 --硬编码: 窗体标题--
            this.Text = "三箱型号添写 - 管道决策向导";
            // 设定适宜左右双栏展示的高宽尺寸 (1100x680 提供更充裕的弹性水平宽度)
            this.ClientSize = new Size(1100, 680);
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
                string webViewCacheDir = Path.Combine(userDataDir, "WebView2_CabinetPipeline");

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

                // 配置多级备选路径集以实现高容错加载
                string[] candidatePaths = new string[]
                {
                    Path.Combine(appDir, "Resources", "cabinet_model_pipeline.html"),
                    Path.Combine(appDir, "cabinet_model_pipeline.html"),
                    Path.Combine(baseDir, "Resources", "cabinet_model_pipeline.html"),
                    Path.Combine(baseDir, "cabinet_model_pipeline.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "cabinet_model_pipeline.html")
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
                    MessageBox.Show("未找到三箱型号向导界面资源文件: cabinet_model_pipeline.html", "资源缺失", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[三箱型号管道窗体] 初始化失败: {ex.Message}");
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
                    // 1. 请求加载初始数据 (包含分类表列表、管道规则配置及首表箱柜)
                    case "getInitData":
                        var initData = _controller.GetInitData();
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "initDataLoaded",
                            data = initData
                        }, JsonOptions));
                        break;

                    // 2. 切换当前查看的分类工作表
                    case "switchSheet":
                        string sheetName = root.TryGetProperty("sheetName", out var sProp) ? sProp.GetString() ?? "" : "";
                        var switchConfig = new CabinetPipelineConfig();
                        if (root.TryGetProperty("config", out var cfgProp))
                        {
                            switchConfig = JsonSerializer.Deserialize<CabinetPipelineConfig>(cfgProp.GetRawText(), JsonOptions) ?? switchConfig;
                        }
                        var switchedCabinets = _controller.SwitchCategorySheet(sheetName, switchConfig);
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "sheetSwitched",
                            data = switchedCabinets
                        }, JsonOptions));
                        break;

                    // 3. 保存用户修改后的管道规则配置
                    case "saveConfig":
                        if (root.TryGetProperty("config", out var saveCfgProp))
                        {
                            var cfgToSave = JsonSerializer.Deserialize<CabinetPipelineConfig>(saveCfgProp.GetRawText(), JsonOptions);
                            bool saved = _controller.SaveConfig(cfgToSave);
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "configSaved",
                                success = saved
                            }, JsonOptions));
                        }
                        break;

                    // 4. 重置为默认标准预设规则
                    case "resetDefaultRules":
                        var defaultCfg = ExcelServices.CreateDefaultPipelineConfig();
                        ExcelServices.SaveCabinetPipelineConfig(defaultCfg);
                        // 重新获取全套初始化数据
                        var refreshedInit = _controller.GetInitData();
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "initDataLoaded",
                            data = refreshedInit
                        }, JsonOptions));
                        break;

                    // 5. 提交批量回写至 Excel
                    case "batchApply":
                        var applyItems = new List<CabinetPipelineApplyItem>();
                        if (root.TryGetProperty("items", out var itemsProp) && itemsProp.ValueKind == JsonValueKind.Array)
                        {
                            applyItems = JsonSerializer.Deserialize<List<CabinetPipelineApplyItem>>(itemsProp.GetRawText(), JsonOptions) ?? applyItems;
                        }
                        bool writeSummaryN = root.TryGetProperty("writeSummaryN", out var wnProp) ? wnProp.GetBoolean() : true;
                        bool writeSummaryD = root.TryGetProperty("writeSummaryD", out var wsProp) ? wsProp.GetBoolean() : false;
                        bool writeDetailD = root.TryGetProperty("writeDetailD", out var wdProp) ? wdProp.GetBoolean() : false;

                        var applyResult = _controller.BatchApplyModels(applyItems, writeSummaryN, writeSummaryD, writeDetailD);
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "applyFinished",
                            success = applyResult.Success,
                            message = applyResult.Message,
                            updatedCount = applyResult.UpdatedCount
                        }, JsonOptions));
                        break;

                    // 6. 无边框拖拽移动窗体
                    case "dragWindow":
                        HandleWindowDrag();
                        break;

                    // 7. 关闭向导窗口
                    case "closeWindow":
                        SafeInvoke(this.Close);
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[三箱型号管道窗体] WebMessageReceived 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 安全处理无边框拖拽移动，防范幽灵鼠标死锁 (遵循启发式经验 16)
        /// </summary>
        private void HandleWindowDrag()
        {
            try
            {
                // 检测物理鼠标左键是否仍在按下状态，避免在快速点按或已抬起后启动移动模态导致全局鼠标捕获死锁
                if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0)
                {
                    ReleaseCapture();
                    SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[三箱型号管道窗体] 窗口拖拽异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 线程安全的跨线程回传消息至 WebView2 渲染进程 (使用 PostWebMessageAsJson 保证前端对象解析)
        /// </summary>
        /// <param name="json">JSON 格式消息字符串</param>
        private void PostWebMessageSafe(string json)
        {
            // 窗体句柄与释放状态防御性校验
            if (this.IsDisposed || !this.IsHandleCreated || _webView.IsDisposed) return;

            // 跨线程安全切入 UI 主线程执行投递
            SafeInvoke(() =>
            {
                // 确保 WebView2 核心非空且处于可用状态
                if (!_webView.IsDisposed && _webView.CoreWebView2 != null)
                {
                    // 采用 PostWebMessageAsJson 投递原生 JSON，避免前端作为纯字符串解析失败
                    _webView.CoreWebView2.PostWebMessageAsJson(json);
                }
            });
        }

        /// <summary>
        /// 控件句柄与释放状态安全的 Invoke 调用包装 (遵循启发式经验 10)
        /// </summary>
        /// <param name="action">待执行的委托</param>
        private void SafeInvoke(System.Action action)
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
    }
}
