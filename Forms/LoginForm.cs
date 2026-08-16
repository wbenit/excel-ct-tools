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
    /// 基于 WebView2 + Vue 3 的用户登录与配置宿主窗口
    /// </summary>
    public class LoginForm : Form
    {
        // 声明 WebView2 浏览器控件实例对象
        private readonly WebView2 _webView;

        // 声明后端 AuthController 身份验证控制器
        private readonly AuthController _authController;

        /// <summary>
        /// 构造函数：初始化窗体与 WebView2 控件属性
        /// </summary>
        public LoginForm()
        {
            // 实例化 Backend 认证控制器对象
            _authController = new AuthController();

            // 实例化 WebView2 嵌入式浏览器控件
            _webView = new WebView2();

            // 设置登录配置窗体的基本属性
            InitializeFormProperties();

            // 初始化 WebView2 控件并加入窗体控件集合
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体几何参数与外观样式
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题文本
            this.Text = "用户身份认证与系统配置";

            // 设置窗体尺寸固定大小为 420x540 像素
            this.ClientSize = new Size(420, 540);

            // 设置窗体在屏幕中央弹出显示
            this.StartPosition = FormStartPosition.CenterScreen;

            // 禁止用户随意调整登录窗口大小
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            // 隐藏最大化按钮以保证最佳美观度
            this.MaximizeBox = false;

            // 允许显示最小化按钮
            this.MinimizeBox = true;

            // 设置窗体背景颜色为深色调
            this.BackColor = Color.FromArgb(15, 23, 42);
        }

        /// <summary>
        /// 初始化 WebView2 控件布局并设置事件回调
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 设置 WebView2 控件充满整个窗体容器
            _webView.Dock = DockStyle.Fill;

            // 将 WebView2 控件添加至当前窗体的控件列表
            this.Controls.Add(_webView);

            // 注册窗体加载事件以便完成异步 WebView2 环境初始化
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载异步事件：初始化 WebView2 Core 并加载前端 HTML
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 计算用户可写的本地 AppData 数据目录路径，防止默认写入 Program Files 导致的 E_ACCESSDENIED 权限拒绝异常
                string userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAddInDemo", "WebView2Data");

                // 创建包含完整读写权限的缓存文件夹目录
                Directory.CreateDirectory(userDataFolder);

                // 创建具有特定用户数据路径的 CoreWebView2Environment 运行环境
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 使用指定的安全环境对象初始化 WebView2 运行环境
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 注册监听 Vue 3 前端发出的消息事件回调
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 计算 Resources/login.html 文件的绝对路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // 组合 HTML 文件的物理存放路径
                string htmlPath = Path.Combine(baseDir, "Resources", "login.html");

                // 如果本地没有发现资源文件，寻找项目所在路径或写入临时文件
                if (!File.Exists(htmlPath))
                {
                    // 向上回溯查找工程目录中的 Resources/login.html
                    string projectHtml = Path.Combine(Directory.GetCurrentDirectory(), "Resources", "login.html");

                    // 若找到了工程内部的 HTML 则更新加载路径
                    if (File.Exists(projectHtml))
                    {
                        // 使用工程文件路径
                        htmlPath = projectHtml;
                    }
                }

                // 导航 WebView2 页面至本地 Vue 3 HTML 文件
                _webView.Source = new Uri(htmlPath);
            }
            catch (Exception ex)
            {
                // 弹出异常错误提示框
                MessageBox.Show($"初始化 WebView2 控件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 响应 Vue 3 前端发来的 JSON 消息事件
        /// </summary>
        private async void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 获取前端传递过来的 Web Message 字符串
                string messageJson = e.TryGetWebMessageAsString();

                // 使用 JsonDocument 解析 JSON 格式数据
                using var doc = JsonDocument.Parse(messageJson);

                // 获取 RootElement 数据节点
                var root = doc.RootElement;

                // 检查动作类型是否为 login 登录
                if (root.TryGetProperty("action", out var actionProp) && actionProp.GetString() == "login")
                {
                    // 解析获取输入的用户名参数
                    string username = root.GetProperty("username").GetString() ?? string.Empty;

                    // 解析获取输入的密码参数
                    string password = root.GetProperty("password").GetString() ?? string.Empty;

                    // 解析获取 RememberMe 复选框勾选状态
                    bool rememberMe = root.TryGetProperty("rememberMe", out var remProp) && remProp.GetBoolean();

                    // 封装构造登录请求数据结构
                    var request = new LoginRequest
                    {
                        Username = username,
                        Password = password,
                        RememberMe = rememberMe
                    };

                    // 调用 Backend WebAPI 控制器接口执行异步校验
                    LoginResponse response = await _authController.LoginAsync(request);

                    // 将包含状态信息的响应结果发回给 Vue 3 界面显示 (跨线程安全)
                    string responseJson = JsonSerializer.Serialize(response);

                    // 使用 PostWebMessageSafe 发送给 Vue 3 页面
                    PostWebMessageSafe(responseJson);

                    // 若登录校验成功，保存当前用户登录凭据与状态
                    if (response.Success)
                    {
                        // 将配置信息与 Token 记录回 ExcelServices 模块
                        ExcelServices.CurrentToken = response.Token;

                        // 将登录的用户名更新回 ExcelServices
                        ExcelServices.CurrentUserDisplayName = response.User?.DisplayName ?? username;

                        // 异步等待 800 毫秒展示成功的动画给用户
                        await System.Threading.Tasks.Task.Delay(800);

                        // 登录完成后在主线程关闭当前登录配置窗口
                        SafeInvoke(() => this.Close());
                    }
                }
            }
            catch (Exception ex)
            {
                // 异常时提示失败信息给前端 (跨线程安全)
                var errResponse = new LoginResponse
                {
                    Success = false,
                    Message = $"处理登录失败: {ex.Message}"
                };

                // 将异常响应序列化并回发
                PostWebMessageSafe(JsonSerializer.Serialize(errResponse));
            }
        }

        /// <summary>
        /// 跨线程安全向 WebView2 发送 JSON 消息
        /// </summary>
        private void PostWebMessageSafe(string json)
        {
            // 在 UI 线程中调度发送
            SafeInvoke(() =>
            {
                // 确保 WebView2 控件及其内核有效
                if (!this.IsDisposed && _webView?.CoreWebView2 != null)
                {
                    // 向前端页面发送 JSON 消息
                    _webView.CoreWebView2.PostWebMessageAsJson(json);
                }
            });
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
    }
}
