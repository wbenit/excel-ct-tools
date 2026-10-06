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

                // 检查动作类型
                if (!root.TryGetProperty("action", out var actionProp)) return;
                string action = actionProp.GetString() ?? string.Empty;

                // 动作 1：获取当前会话状态
                if (action == "getSession")
                {
                    var sessionInfo = _authController.GetCurrentSessionInfo();
                    PostWebMessageSafe(JsonSerializer.Serialize(new { action = "sessionInfo", data = sessionInfo }));
                    return;
                }

                // 动作 2：注销当前账号
                if (action == "logout")
                {
                    _authController.Logout();
                    PostWebMessageSafe(JsonSerializer.Serialize(new { action = "logoutResult", success = true, message = "已安全退出当前账号！" }));
                    return;
                }

                // 动作 3：执行登录认证
                if (action == "login")
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

                    // 构造回发数据
                    var resObj = new
                    {
                        action = "loginResult",
                        Success = response.Success,
                        Message = response.Message,
                        Token = response.Token,
                        UserId = response.UserId,
                        DisplayName = response.DisplayName,
                        NeedSelectGroup = response.NeedSelectGroup,
                        SelectedGroupId = response.SelectedGroupId,
                        SelectedGroupName = response.SelectedGroupName,
                        Groups = response.Groups
                    };

                    // 将包含状态信息的响应结果发回给 Vue 3 界面显示 (跨线程安全)
                    PostWebMessageSafe(JsonSerializer.Serialize(resObj));

                    // 若登录校验成功且无需选择工作组，等待 800ms 关闭窗口
                    if (response.Success && !response.NeedSelectGroup)
                    {
                        await System.Threading.Tasks.Task.Delay(800);
                        SafeInvoke(() => this.Close());
                    }
                    return;
                }

                // 动作 4：用户手动选择并确认工作组
                if (action == "selectGroup")
                {
                    string token = root.GetProperty("token").GetString() ?? string.Empty;
                    string userId = root.GetProperty("userId").GetString() ?? string.Empty;
                    string userName = root.GetProperty("userName").GetString() ?? string.Empty;
                    string displayName = root.GetProperty("displayName").GetString() ?? string.Empty;
                    int groupId = root.GetProperty("groupId").GetInt32();
                    string groupName = root.TryGetProperty("groupName", out var gnProp) ? (gnProp.GetString() ?? string.Empty) : string.Empty;
                    bool rememberMe = root.TryGetProperty("rememberMe", out var remProp) && remProp.GetBoolean();

                    var selectReq = new SelectGroupRequest
                    {
                        Token = token,
                        UserId = userId,
                        UserName = userName,
                        DisplayName = displayName,
                        GroupId = groupId,
                        GroupName = groupName,
                        RememberMe = rememberMe
                    };

                    LoginResponse selectRes = await _authController.SelectGroupAsync(selectReq);

                    var resObj = new
                    {
                        action = "selectGroupResult",
                        Success = selectRes.Success,
                        Message = selectRes.Message,
                        SelectedGroupId = selectRes.SelectedGroupId,
                        SelectedGroupName = selectRes.SelectedGroupName
                    };

                    PostWebMessageSafe(JsonSerializer.Serialize(resObj));

                    if (selectRes.Success)
                    {
                        await System.Threading.Tasks.Task.Delay(800);
                        SafeInvoke(() => this.Close());
                    }
                }
            }
            catch (Exception ex)
            {
                // 异常时提示失败信息给前端 (跨线程安全)
                var errResponse = new
                {
                    action = "error",
                    Success = false,
                    Message = $"处理异常: {ex.Message}"
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
