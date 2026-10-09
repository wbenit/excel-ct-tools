using System;
using System.IO;
using System.Drawing;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using ExcelAddInDemo.Controllers;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的系统设置宿主窗口
    /// </summary>
    public class AppSettingsForm : Form
    {
        // 声明 WebView2 浏览器控件实例
        private readonly WebView2 _webView;

        // 声明系统设置数据控制器
        private readonly AppSettingsController _controller;

        // 声明全局通用的 JSON 序列化规范，支持驼峰转换与大小写容错
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            // 开启驼峰转换
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            // 开启忽略大小写
            PropertyNameCaseInsensitive = true,
            // 缩进格式化
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数：初始化窗体与 WebView2 控件属性
        /// </summary>
        public AppSettingsForm()
        {
            // 实例化系统设置控制器
            _controller = new AppSettingsController();

            // 实例化 WebView2 嵌入式浏览器控件
            _webView = new WebView2();

            // 配置窗体外观与初始尺寸
            InitializeFormProperties();

            // 初始化 WebView2 布局并挂载窗体
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体标题、尺寸与居中位置
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题为“系统设置” --硬编码: 窗体标题--
            this.Text = "系统设置";

            // 设置窗体初始客户区尺寸为 660x420 像素 --硬编码: 标准初始尺寸 660x420--
            this.ClientSize = new Size(660, 420);

            // 设置窗体在屏幕居中显示
            this.StartPosition = FormStartPosition.CenterScreen;

            // 固定对话框边框样式防止形变
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            // 禁用最大化按钮
            this.MaximizeBox = false;

            // 启用最小化按钮
            this.MinimizeBox = true;

            // 设置统一的现代浅灰背景底色
            this.BackColor = Color.FromArgb(248, 250, 252);
        }

        /// <summary>
        /// 初始化 WebView2 控件布局与加载监听
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 设置 WebView2 占满窗体工作区
            _webView.Dock = DockStyle.Fill;

            // 将控件添加至窗体集合
            this.Controls.Add(_webView);

            // 注册窗体加载事件
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载异步事件：初始化 CoreWebView2 运行环境并导航页面
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 若窗体已销毁则立即退出
                if (this.IsDisposed || this.Disposing) return;

                // 统一获取 LocalApplicationData 隔离缓存目录，绝不污染用户业务目录
                string userDataFolder = Tool.GetWebView2UserDataFolder("WebView2_AppSettings");

                // 创建独立的 WebView2 运行环境
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 校验异步等待后窗体状态
                if (this.IsDisposed || this.Disposing) return;

                // 初始化内核
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 注册 WebMessageReceived 消息监听
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 获取基础运行目录
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // 探测 HTML 文件候选路径数组
                string[] candidatePaths = new string[]
                {
                    Path.Combine(baseDir, "Resources", "app_settings.html"),
                    Path.Combine(baseDir, "publish", "Resources", "app_settings.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "app_settings.html")
                };

                // 最终定位的 HTML 路径
                string htmlPath = string.Empty;

                // 循环查找首个存在的 HTML
                foreach (string candidate in candidatePaths)
                {
                    if (File.Exists(candidate))
                    {
                        htmlPath = candidate;
                        break;
                    }
                }

                // 校验路径有效性并导航
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    _webView.Source = new Uri(htmlPath);
                }
                else
                {
                    MessageBox.Show($"未找到设置页面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                // 捕获提示初始化异常
                MessageBox.Show($"初始化 WebView2 控件失败: {ex.Message}", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 响应前端发送的 WebMessage 交互消息
        /// </summary>
        private async void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 声明接收到的消息 JSON 文本
                string messageJson = string.Empty;

                // 独立尝试读取纯文本消息（规避参数类型不符引起的异常）
                try
                {
                    messageJson = e.TryGetWebMessageAsString();
                }
                catch (Exception ex)
                {
                    // 记录轻量容错日志
                    LogHelper.WriteLog($"[AppSettingsForm] TryGetWebMessageAsString 容错: {ex.Message}");
                }

                // 若文本为空，尝试降级读取 WebMessageAsJson（独立保护防没有注册类 COM 异常）
                if (string.IsNullOrWhiteSpace(messageJson))
                {
                    try
                    {
                        messageJson = e.WebMessageAsJson;
                    }
                    catch (Exception ex)
                    {
                        // 记录轻量容错日志
                        LogHelper.WriteLog($"[AppSettingsForm] WebMessageAsJson 降级容错: {ex.Message}");
                    }
                }

                // 若依然为空则静默退出，严禁抛错打扰用户
                if (string.IsNullOrWhiteSpace(messageJson)) return;

                // 解析最外层 JSON 节点
                using var doc = JsonDocument.Parse(messageJson);
                var root = doc.RootElement;

                // 防御性深层解包：若消息被二次打包为 JSON String，自动解包深层真实的 JSON Object
                JsonDocument? subDoc = null;
                JsonElement effectiveRoot = root;
                if (root.ValueKind == JsonValueKind.String)
                {
                    // 解析嵌套的内部 JSON 字符串
                    subDoc = JsonDocument.Parse(root.GetString() ?? "{}");
                    effectiveRoot = subDoc.RootElement;
                }

                try
                {
                    // 从有效 Object 根节点读取 action 指令
                    string action = effectiveRoot.TryGetProperty("action", out var actProp) ? actProp.GetString() ?? "" : "";

                    // 根据指令进行分发处理
                    switch (action)
                    {
                        // 获取配置数据进行反显
                        case "getSettings":
                            // 异步加载当前全局配置实体
                            AppSettingsData data = await _controller.LoadSettingsAsync();
                            // 封装反显消息
                            var renderMsg = new { action = "renderSettings", data = data };
                            // 推送回前端
                            PostWebMessageSafe(JsonSerializer.Serialize(renderMsg, JsonOptions));
                            break;

                        // 浏览选择共享数据目录
                        case "selectDataDirectory":
                            // 启动独立 STA 线程弹出文件夹选择器
                            SelectDirectoryFolder("选择 CAD 与 Excel 共享的数据与配置文件存储目录", "setDataDirectory");
                            break;

                        // 浏览选择新建项目默认保存目录
                        case "selectProjectDirectory":
                            // 启动独立 STA 线程弹出新建项目文件夹选择器
                            SelectDirectoryFolder("选择新建项目默认保存目录", "setProjectDirectory");
                            break;

                        // 快捷重置新建项目目录为桌面
                        case "resetProjectDirectoryToDesktop":
                            // 提取系统桌面物理路径
                            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                            // 构造回传重置消息
                            var resetMsg = new { action = "setProjectDirectory", path = desktopPath };
                            PostWebMessageSafe(JsonSerializer.Serialize(resetMsg, JsonOptions));
                            break;

                        // 提交保存设置
                        case "saveSettings":
                            // 检查是否存在 data 节点
                            if (effectiveRoot.TryGetProperty("data", out var dataProp))
                            {
                                // 反序列化设置模型
                                var model = JsonSerializer.Deserialize<AppSettingsData>(dataProp.GetRawText(), JsonOptions);
                                if (model != null)
                                {
                                    // 异步持久化配置
                                    bool success = await _controller.SaveSettingsAsync(model);
                                    SafeInvoke(() =>
                                    {
                                        // 组织反馈文本
                                        string feedbackMsg = success ? "全局设置已成功保存并立即生效！" : "保存设置失败，请检查文件写入权限。";
                                        var resMsg = new
                                        {
                                            action = "saveSettingsResult",
                                            success = success,
                                            message = feedbackMsg
                                        };
                                        // 回发保存结果
                                        PostWebMessageSafe(JsonSerializer.Serialize(resMsg));
                                    });
                                }
                            }
                            break;

                        // 动态自适应调整窗体高度消灭留白
                        case "resizeWindow":
                            // 提取 height 属性
                            if (effectiveRoot.TryGetProperty("height", out var hProp))
                            {
                                // 兼容浮点型与整型高度数值
                                double dHeight = 0;
                                if (hProp.ValueKind == JsonValueKind.Number && hProp.TryGetDouble(out dHeight))
                                {
                                    // 四舍五入取整
                                    int domHeight = (int)Math.Round(dHeight);
                                    // 设置安全高度范围约束
                                    if (domHeight >= 320 && domHeight <= 650)
                                    {
                                        SafeInvoke(() =>
                                        {
                                            // 防抖调整物理高度
                                            if (Math.Abs(this.ClientSize.Height - domHeight) > 4)
                                            {
                                                this.ClientSize = new Size(this.ClientSize.Width, domHeight);
                                            }
                                        });
                                    }
                                }
                            }
                            break;

                        // 取消动作：关闭窗口
                        case "cancel":
                            // 关闭当前窗体
                            SafeInvoke(() => this.Close());
                            break;
                    }
                }
                finally
                {
                    // 显式释放临时嵌套子文档对象
                    subDoc?.Dispose();
                }
            }
            catch (Exception ex)
            {
                // 静默记录异常日志，杜绝弹出阻断性弹窗打扰用户
                LogHelper.WriteLog($"[AppSettingsForm] 处理消息异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 在独立的后台 STA 线程中弹出文件夹选择器，彻底规避 Chromium IPC 死锁
        /// </summary>
        /// <param name="description">对话框描述文本</param>
        /// <param name="actionName">回送前端的动作标识</param>
        private void SelectDirectoryFolder(string description, string actionName)
        {
            // 启动独立的 STA 线程
            var dialogThread = new System.Threading.Thread(() =>
            {
                try
                {
                    // 实例化 FolderBrowserDialog 对象
                    using var dialog = new FolderBrowserDialog
                    {
                        Description = description,
                        ShowNewFolderButton = true,
                        RootFolder = Environment.SpecialFolder.Desktop
                    };

                    // 弹出模态选择框
                    if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                    {
                        string chosenPath = dialog.SelectedPath;
                        // 线程安全切回主线程通知前端
                        SafeInvoke(() =>
                        {
                            var msg = new { action = actionName, path = chosenPath };
                            PostWebMessageSafe(JsonSerializer.Serialize(msg, JsonOptions));
                        });
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[AppSettingsForm] 文件夹选择异常: {ex.Message}");
                }
            });
            // 标记单线程单元
            dialogThread.SetApartmentState(System.Threading.ApartmentState.STA);
            dialogThread.IsBackground = true;
            dialogThread.Start();
        }

        /// <summary>
        /// 跨线程安全向前端 WebView2 投递 JSON 消息
        /// </summary>
        private void PostWebMessageSafe(string json)
        {
            SafeInvoke(() =>
            {
                if (!this.IsDisposed && _webView?.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.PostWebMessageAsJson(json);
                }
            });
        }

        /// <summary>
        /// 安全调度 UI 线程操作
        /// </summary>
        private void SafeInvoke(Action action)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            if (this.InvokeRequired)
            {
                this.Invoke(action);
            }
            else
            {
                action();
            }
        }

        /// <summary>
        /// 窗体关闭时显式释放 WebView2 资源
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try
            {
                if (_webView?.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
                }
                _webView?.Dispose();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[AppSettingsForm] OnFormClosing 异常: {ex.Message}");
            }
            base.OnFormClosing(e);
        }
    }
}
