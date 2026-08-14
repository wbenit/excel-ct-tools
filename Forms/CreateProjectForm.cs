using System;
using System.IO;
using System.Drawing;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using ExcelAddInDemo.Controllers;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“新建项目”宿主窗口
    /// </summary>
    public class CreateProjectForm : Form
    {
        // 声明 WebView2 浏览器控件
        private readonly WebView2 _webView;

        // 导入 Windows 原生 user32.dll 内存接口，用于支持无边框窗体鼠标按住拖拽
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 原生函数发送系统标头拖拽消息
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 声明项目控制器
        private readonly ProjectController _projectController;

        // 声明企业设置控制器，用于读取全局企业名称与报价人
        private readonly EnterpriseSettingsController _settingsController;

        // 全局通用的 JSON 序列化规范，开启驼峰转换与忽略大小写
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数：初始化控制器与 WebView2 控件
        /// </summary>
        public CreateProjectForm()
        {
            // 实例化新建项目控制器
            _projectController = new ProjectController();

            // 实例化企业设置控制器
            _settingsController = new EnterpriseSettingsController();

            // 实例化 WebView2 浏览器控件
            _webView = new WebView2();

            // 配置新建项目窗口基本外观与几何属性
            InitializeFormProperties();

            // 初始化 WebView2 控件布局并挂载
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置新建项目窗体无边框/固定尺寸与屏幕居中
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题文本
            this.Text = "新建项目";

            // 默认收起状态的显示尺寸调整为 540x510 像素，确保所有表单元素完整展示
            this.ClientSize = new Size(540, 510);

            // 设置窗体在屏幕中央弹窗
            this.StartPosition = FormStartPosition.CenterScreen;

            // 设置无边框窗口样式，去除 WinForm 原生外边框与标题栏
            this.FormBorderStyle = FormBorderStyle.None;

            // 隐藏最大化按钮
            this.MaximizeBox = false;

            // 开启最小化按钮
            this.MinimizeBox = true;

            // 设置背景填充色
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件布局与加载监听
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 控件充满当前工作区
            _webView.Dock = DockStyle.Fill;

            // 添加至 Controls 集合
            this.Controls.Add(_webView);

            // 注册 Form Load 异步处理
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载异步事件：初始化 CoreWebView2 并导航至 create_project.html
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 若窗体已销毁则立即退出
                if (this.IsDisposed || this.Disposing) return;

                // 计算用户 AppData 本地运行数据缓存路径
                string userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAddInDemo", "WebView2Data");

                // 创建缓存文件夹
                Directory.CreateDirectory(userDataFolder);

                // 创建 WebView2 环境
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 判断窗体有效性
                if (this.IsDisposed || this.Disposing) return;

                // 初始化 CoreWebView2 核心句柄
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 挂载前端 postMessage 消息监听
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 拼接目标 HTML 页面物理路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // 计算多层备选寻找路径
                string[] candidatePaths = new string[]
                {
                    Path.Combine(baseDir, "Resources", "create_project.html"),
                    Path.Combine(baseDir, "..", "Resources", "create_project.html"),
                    Path.Combine(baseDir, "publish", "Resources", "create_project.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "create_project.html")
                };

                // 命中目标路径
                string htmlPath = string.Empty;

                // 循环查找首个存在的文件
                foreach (string candidate in candidatePaths)
                {
                    if (File.Exists(candidate))
                    {
                        htmlPath = candidate;
                        break;
                    }
                }

                // 校验文件有效性并导航
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    _webView.Source = new Uri(htmlPath);
                }
                else
                {
                    MessageBox.Show($"未找到新建项目界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化 WebView2 控件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 响应 Vue 3 前端发来的 JSON 交互消息
        /// </summary>
        private async void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 获取前端传递的消息字符串
                string messageJson = e.TryGetWebMessageAsString();

                // 解析 JSON 文档
                using var doc = JsonDocument.Parse(messageJson);

                // 获取 Root 节点
                var root = doc.RootElement;

                // 获取 action 指令
                string action = root.TryGetProperty("action", out var actProp) ? actProp.GetString() ?? "" : "";

                // 判断动作分支
                switch (action)
                {
                    // 获取窗体初始化所需数据
                    case "getInitData":
                        // 自动生成单号
                        string qNum = _projectController.GenerateQuoteNumber();

                        // 获取默认保存桌面路径
                        string sPath = _projectController.GetDefaultDesktopPath();

                        // 读取“企业设置”中保存的单位名称与报价人
                        EnterpriseSettingsData settingData = await _settingsController.LoadSettingsAsync();

                        // 构造回发给前端的数据
                        var initMsg = new
                        {
                            action = "renderInitData",
                            quoteNumber = qNum,
                            savePath = sPath,
                            companyName = settingData.CompanyName,
                            quoter = settingData.Quoter
                        };

                        // 回发 JSON 给 Vue 3 前端
                        _webView.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(initMsg, JsonOptions));
                        break;

                    // 重新生成报价单号
                    case "refreshQuoteNumber":
                        // 生成新单号
                        string newQNum = _projectController.GenerateQuoteNumber();

                        // 回发 setQuoteNumber
                        var qMsg = new { action = "setQuoteNumber", quoteNumber = newQNum };
                        _webView.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(qMsg, JsonOptions));
                        break;

                    // 动态调整窗口高度 (收起 vs 展开)
                    case "toggleExpand":
                        if (root.TryGetProperty("expanded", out var expProp))
                        {
                            bool isExp = expProp.GetBoolean();
                            SafeInvoke(() =>
                            {
                                // 展开图 2 状态设为 770 像素高度；收起图 1 状态设为 510 像素高度
                                this.ClientSize = isExp ? new Size(540, 770) : new Size(540, 510);
                            });
                        }
                        break;

                    // 选择保存文件夹
                    case "selectFolder":
                        SelectSaveFolder();
                        break;

                    // 提交开始报价：新建 Workbook，复制 CabinetTemplate.xlsx 及其工作表并填入数据
                    case "startQuotation":
                        if (root.TryGetProperty("data", out var dataProp))
                        {
                            var model = JsonSerializer.Deserialize<CreateProjectModel>(dataProp.GetRawText(), JsonOptions);
                            if (model != null)
                            {
                                // 异步执行创建项目工作簿、复制工作表与单元格数据填入
                                bool success = await _projectController.CreateProjectAsync(model);

                                if (success)
                                {
                                    // 显式关闭当前新建项目弹窗
                                    SafeInvoke(() => this.Close());
                                }
                            }
                        }
                        break;

                    // 取消动作
                    case "cancel":
                        SafeInvoke(() => this.Close());
                        break;

                    // 响应 HTML 顶栏拖拽指令，移动无边框窗体
                    case "dragWindow":
                        SafeInvoke(() =>
                        {
                            // 释放当前鼠标捕获句柄
                            ReleaseCapture();

                            // 发送 WM_NCLBUTTONDOWN (0xA1) 消息触发原生无边框窗口拖拽
                            SendMessage(this.Handle, 0xA1, (IntPtr)0x2, IntPtr.Zero);
                        });
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理消息发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 安全跨线程调度 UI 动作，防止在句柄未创建或窗体已被释放时调用 Invoke 抛出 InvalidOperationException
        /// </summary>
        private void SafeInvoke(Action action)
        {
            // 校验窗体句柄有效性与释放状态
            if (this.IsDisposed || !this.IsHandleCreated) return;

            // 根据是否跨线程选择 Invoke 或直接执行
            if (this.InvokeRequired)
            {
                // 跨线程安全调度
                this.Invoke(action);
            }
            else
            {
                // 主 UI 线程直接同步执行
                action();
            }
        }

        /// <summary>
        /// 弹出文件夹选择对话框更新保存路径
        /// </summary>
        private void SelectSaveFolder()
        {
            SafeInvoke(() =>
            {
                using var dialog = new FolderBrowserDialog();
                dialog.Description = "请选择项目新建与出报表保存目录";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var msg = new { action = "setSavePath", savePath = dialog.SelectedPath };
                    _webView.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(msg, JsonOptions));
                }
            });
        }
    }
}
