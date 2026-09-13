using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Controllers;
using ExcelDna.Integration;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“分布调价”无边框模态/非模态宿主窗口
    /// </summary>
    public class DistributedAdjustPriceForm : Form
    {
        // 声明 WebView2 浏览器控件实例
        private readonly WebView2 _webView;

        // 声明分布调价 WebAPI 控制器实例
        private readonly DistributedAdjustPriceController _controller;

        // 通用 JSON 序列化配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化控制器与 WebView2 控件属性
        /// </summary>
        public DistributedAdjustPriceForm()
        {
            // 实例化分布调价控制器
            _controller = new DistributedAdjustPriceController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 设置 Form 窗体尺寸与显示几何外观
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸 (720x640 像素)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题文本
            this.Text = "分布调价";

            // 设定宽 720、高 640 像素
            this.ClientSize = new Size(720, 640);

            // 居中显示
            this.StartPosition = FormStartPosition.CenterScreen;

            // 无边框现代窗口设计
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用 WinForm 原生最大化
            this.MaximizeBox = false;

            // 禁用原生最小化
            this.MinimizeBox = false;

            // 设置纯白底色
            this.BackColor = Color.White;

            // 保证置顶以便在与 Excel 联动时清晰可见
            this.TopMost = true;
        }

        /// <summary>
        /// 初始化 WebView2 控件并注册加载回调
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 充满窗体客户区
            _webView.Dock = DockStyle.Fill;

            // 挂载至 Controls 控件集
            this.Controls.Add(_webView);

            // 绑定 Form Load 异步加载监听
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 异步加载并初始化 WebView2 核心环境
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 指定 WebView2 用户运行数据独立缓存目录，避免与其他插件冲突
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAddInCTtools",
                    "WebView2Data"
                );

                // 异步创建核心运行时环境
                var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 初始化 CoreWebView2 实例
                await _webView.EnsureCoreWebView2Async(environment);

                // 关闭浏览器自带右键上下文菜单
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;

                // 关闭浏览器开发人员快捷键 (F12)
                _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                // 注册 WebMessage 前后端消息通信接收监听器
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 获取工程 Resources 资源目录绝对路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string resDir = Path.Combine(baseDir, "Resources");

                // 注册虚拟主机名映射，杜绝跨域拦截
                _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "appassets.local",
                    resDir,
                    CoreWebView2HostResourceAccessKind.Allow
                );

                // 检查本地 HTML 资源文件是否存在
                string htmlPath = Path.Combine(resDir, "distributed_adjust_price.html");
                if (File.Exists(htmlPath))
                {
                    // 使用安全域名协议导航加载
                    _webView.Source = new Uri("https://appassets.local/distributed_adjust_price.html");
                }
                else
                {
                    // 弹出未找到资源警告
                    MessageBox.Show($"未找到分布调价界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                // 弹出核心初始化失败异常
                MessageBox.Show($"初始化 WebView2 内核失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogHelper.WriteLog($"[分布调价] WebView2初始化异常: {ex.Message}\r\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// 统一处理来自 Vue 3 前端的异步通信消息与指令调用
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 读取收到的原始 JSON 文本
                string jsonString = e.TryGetWebMessageAsString();
                if (string.IsNullOrWhiteSpace(jsonString)) return;

                // 解析顶层 JSON DOM 节点
                using var doc = JsonDocument.Parse(jsonString);
                var root = doc.RootElement;

                // 提取操作类型 action 标识
                string action = root.TryGetProperty("action", out var actionProp) ? actionProp.GetString() ?? "" : "";

                // 根据指令类型进行业务路由分发
                switch (action)
                {
                    // 1. 查询当前工作簿所有分类表及台数列表
                    case "getCategorySheets":
                        HandleGetCategorySheets();
                        break;

                    // 2. 根据用户配置生成【材料分布表】
                    case "generateDistributionSheet":
                        HandleGenerateDistributionSheet(root);
                        break;

                    // 3. 从【材料分布表】一键反向更新数据至各分类表
                    case "updateFromDistributionSheet":
                        HandleUpdateFromDistributionSheet(root);
                        break;

                    // 4. 检查材料分布表是否存在
                    case "checkDistributionSheetExists":
                        HandleCheckDistributionSheetExists();
                        break;

                    // 5. 最小化宿主窗口
                    case "minimizeWindow":
                        this.WindowState = FormWindowState.Minimized;
                        break;

                    // 6. 关闭宿主窗口
                    case "closeWindow":
                        this.Close();
                        break;

                    // 7. 标题栏平滑拖拽移动
                    case "moveWindow":
                        if (root.TryGetProperty("deltaX", out var dxProp) && root.TryGetProperty("deltaY", out var dyProp))
                        {
                            int dx = dxProp.GetInt32();
                            int dy = dyProp.GetInt32();
                            this.Location = new Point(this.Location.X + dx, this.Location.Y + dy);
                        }
                        break;

                    default:
                        LogHelper.WriteLog($"[分布调价] 收到未知指令: {action}");
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[分布调价] 处理前端消息异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 响应获取分类工作表与箱柜统计数据请求（在 Excel STA 线程同步获取，规避多线程 COM 限制）
        /// </summary>
        private void HandleGetCategorySheets()
        {
            try
            {
                // 记录调试日志
                LogHelper.WriteLog("[分布调价] 开始获取当前工作簿所有分类表...");
                // 在当前 STA 线程中同步调用控制器获取分类数据 JSON
                string jsonResult = _controller.GetCategorySheetsJson();
                // 记录获取结果日志
                LogHelper.WriteLog($"[分布调价] 获取分类数据完成: {jsonResult}");

                // 封装标准回传报文
                var payload = new
                {
                    // 报文类型标识
                    type = "getCategorySheetsResult",
                    // 解析响应 JSON DOM
                    response = JsonDocument.Parse(jsonResult).RootElement
                };

                // 线程安全回发前端 Vue 3
                PostWebMessageSafe(JsonSerializer.Serialize(payload, JsonOptions));
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[分布调价] HandleGetCategorySheets 执行异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 响应生成【材料分布表】请求
        /// </summary>
        private void HandleGenerateDistributionSheet(JsonElement root)
        {
            // 提取参数载荷 payload
            string payloadJson = root.TryGetProperty("payload", out var payloadProp) ? payloadProp.GetRawText() : "";

            // 在 Excel 线程宏队列中执行生成逻辑，规避 COM 并发冲突
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                // 调用控制器执行分布表生成
                string jsonResult = _controller.GenerateDistributionSheetJson(payloadJson);

                // 封装回传报文
                var payload = new
                {
                    type = "generateDistributionSheetResult",
                    response = JsonDocument.Parse(jsonResult).RootElement
                };

                // 线程安全回发前端
                PostWebMessageSafe(JsonSerializer.Serialize(payload, JsonOptions));
            });
        }

        /// <summary>
        /// 响应从【材料分布表】一键反向同步更新请求
        /// </summary>
        private void HandleUpdateFromDistributionSheet(JsonElement root)
        {
            // 提取更新参数选项
            string optionsJson = root.TryGetProperty("payload", out var payloadProp) ? payloadProp.GetRawText() : "";

            // 在 Excel 纯净宏上下文中调度反向回写
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                // 调用控制器执行同步回写
                string jsonResult = _controller.UpdateFromDistributionSheetJson(optionsJson);

                // 封装回传消息
                var payload = new
                {
                    type = "updateFromDistributionSheetResult",
                    response = JsonDocument.Parse(jsonResult).RootElement
                };

                // 线程安全回传前端
                PostWebMessageSafe(JsonSerializer.Serialize(payload, JsonOptions));
            });
        }

        /// <summary>
        /// 检查材料分布表是否存在
        /// </summary>
        private void HandleCheckDistributionSheetExists()
        {
            // 调用控制器检查状态
            string jsonResult = _controller.CheckDistributionSheetExistsJson();

            // 封装回传报文
            var payload = new
            {
                type = "checkDistributionSheetExistsResult",
                response = JsonDocument.Parse(jsonResult).RootElement
            };

            // 发送给前端
            PostWebMessageSafe(JsonSerializer.Serialize(payload, JsonOptions));
        }

        /// <summary>
        /// 线程安全向 WebView2 前端推送消息
        /// </summary>
        private void PostWebMessageSafe(string messageJson)
        {
            try
            {
                // 校验句柄与控件存活状态
                if (this.IsDisposed || !this.IsHandleCreated) return;

                // 若在非 UI 线程则调度至主线程执行
                if (this.InvokeRequired)
                {
                    this.BeginInvoke(new Action(() => PostWebMessageSafe(messageJson)));
                    return;
                }

                // 推送消息给 Vue 3 前端
                _webView.CoreWebView2?.PostWebMessageAsString(messageJson);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[分布调价] 推送消息给前端异常: {ex.Message}");
            }
        }
    }
}
