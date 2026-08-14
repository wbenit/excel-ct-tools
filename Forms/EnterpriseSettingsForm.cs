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
    /// 基于 WebView2 + Vue 3 的“我的企业设置”宿主窗口
    /// </summary>
    public class EnterpriseSettingsForm : Form
    {
        // 声明 WebView2 浏览器控件实例对象
        private readonly WebView2 _webView;

        // 声明后端企业设置数据控制器对象
        private readonly EnterpriseSettingsController _controller;

        // 定义全域通用的 JSON 序列化规范，支持驼峰命名转换与忽略大小写校验
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            // 开启驼峰转换
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            // 开启忽略大小写
            PropertyNameCaseInsensitive = true,
            // 格式化输出
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数：初始化窗体与 WebView2 控件属性
        /// </summary>
        public EnterpriseSettingsForm()
        {
            // 实例化 Backend 企业设置控制器
            _controller = new EnterpriseSettingsController();

            // 实例化 WebView2 嵌入式浏览器控件
            _webView = new WebView2();

            // 配置企业设置窗体外观与几何尺寸
            InitializeFormProperties();

            // 初始化 WebView2 控件布局并挂载窗体
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体标题、尺寸与居中位置
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题为“我的企业设置”
            this.Text = "我的企业设置";

            // 设置窗体显示尺寸为 880x680 像素
            this.ClientSize = new Size(880, 680);

            // 设置窗体在屏幕正中央居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 固定窗体边框尺寸防止随意形变
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            // 禁用最大化按钮保持界面美观
            this.MaximizeBox = false;

            // 开启最小化按钮支持
            this.MinimizeBox = true;

            // 设置窗体默认背景色
            this.BackColor = Color.FromArgb(243, 244, 246);
        }

        /// <summary>
        /// 初始化 WebView2 控件布局与加载监听
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 设置 WebView2 充满整个窗体工作区
            _webView.Dock = DockStyle.Fill;

            // 将 WebView2 控件添加至窗体控件集合
            this.Controls.Add(_webView);

            // 注册窗体加载事件以便完成 WebView2 环境初始化
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载异步事件：初始化 CoreWebView2 环境并导航页面
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 若窗体已销毁则立即退出
                if (this.IsDisposed || this.Disposing) return;

                // 计算 AppData 本地缓存运行目录
                string userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAddInDemo", "WebView2Data");

                // 自动创建缓存文件夹
                Directory.CreateDirectory(userDataFolder);

                // 创建 WebView2 专属运行环境
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 判断窗体在异步等待后是否依然有效
                if (this.IsDisposed || this.Disposing) return;

                // 初始化 CoreWebView2 核心句柄
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 注册前端 postMessage 消息监听事件回调
                if (_webView.CoreWebView2 != null)
                {
                    // 挂载消息监听
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 获取当前程序集 AppDomain 基础路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // 计算候选页面定位路径数组，按优先级进行回溯寻找
                string[] candidatePaths = new string[]
                {
                    // 候选路径 1: 当前 AppDomain 根路径/Resources/enterprise_settings.html
                    Path.Combine(baseDir, "Resources", "enterprise_settings.html"),
                    // 候选路径 2: 上级输出目录/Resources/enterprise_settings.html
                   //Path.Combine(baseDir, "..", "Resources", "enterprise_settings.html"),
                    // 候选路径 3: 部署同级 publish/Resources/enterprise_settings.html
                   // Path.Combine(baseDir, "publish", "Resources", "enterprise_settings.html"),
                    // 候选路径 4: 当前工作路径/Resources/enterprise_settings.html
                   // Path.Combine(Directory.GetCurrentDirectory(), "Resources", "enterprise_settings.html")
                };

                // 最终找到的 HTML 路径
                string htmlPath = string.Empty;

                // 遍历查找首个存在的 HTML 文件
                foreach (string candidate in candidatePaths)
                {
                    // 判断候选路径是否存在文件
                    if (File.Exists(candidate))
                    {
                        // 记录命中路径
                        htmlPath = candidate;
                        // 找到后立即中断循环
                        break;
                    }
                }

                // 判断最终找到的文件路径是否有效
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    // 使用标准的 file:/// 协议绝对路径赋值给 Source 导航
                    _webView.Source = new Uri(htmlPath);
                }
                else
                {
                    // 提示找不到页面资源
                    MessageBox.Show($"未找到界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                // 安全捕获提示 WebView2 环境初始化异常，绝不闪退
                MessageBox.Show($"初始化 WebView2 控件失败: {ex.Message}", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 响应 Vue 3 前端发来的 JSON 交互消息事件
        /// </summary>
        private async void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 尝试获取前端传入的消息文本
                string messageJson = e.TryGetWebMessageAsString();

                // 解析 JSON 数据节点
                using var doc = JsonDocument.Parse(messageJson);

                // 获取 Root JSON 节点对象
                var root = doc.RootElement;

                // 读取 action 动作识别指令
                string action = root.TryGetProperty("action", out var actProp) ? actProp.GetString() ?? "" : "";

                // 根据不同的 action 执行逻辑分支处理
                switch (action)
                {
                    // 数据反显：前端要求加载本地企业配置
                    case "loadSettings":
                        // 调用控制器从本地 JSON 读取持久化数据
                        EnterpriseSettingsData data = await _controller.LoadSettingsAsync();

                        // 构造回发给前端的数据反显对象
                        var renderMsg = new { action = "renderSettings", data = data };

                        // 序列化回发 JSON 文本 (传入 JsonOptions 实现驼峰转换)
                        string renderJson = JsonSerializer.Serialize(renderMsg, JsonOptions);

                        // 发送给 Vue 3 前端实现反显绑定
                        _webView.CoreWebView2.PostWebMessageAsJson(renderJson);
                        break;

                    // 选择本地 Logo 图片
                    case "selectLogo":
                        // 调用本地选择图片对话框
                        SelectLogoImage();
                        break;

                    // 保存配置：前端提交保存最新的企业设置
                    case "saveSettings":
                        // 解析获取 data 数据节点
                        if (root.TryGetProperty("data", out var dataProp))
                        {
                            // 将 data 节点反序列化为 EnterpriseSettingsData (传入 JsonOptions 支持小写转 PascalCase)
                            var saveModel = JsonSerializer.Deserialize<EnterpriseSettingsData>(dataProp.GetRawText(), JsonOptions);

                            // 判断解析数据有效性
                            if (saveModel != null)
                            {
                                // 异步保存至本地磁盘文件
                                bool success = await _controller.SaveSettingsAsync(saveModel);

                                // 判断保存结果
                                if (success)
                                {
                                    // 弹出成功保存提示消息框
                                    MessageBox.Show("企业设置数据已成功保存至本地！", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                    // 保存成功后关闭当前配置窗口
                                    SafeInvoke(() => this.Close());
                                }
                                else
                                {
                                    // 保存失败时提示消息
                                    MessageBox.Show("保存数据到本地失败，请检查文件写入权限。", "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                            }
                        }
                        break;

                    // 取消动作：关闭窗口
                    case "cancel":
                        // 关闭当前企业设置窗体
                        SafeInvoke(() => this.Close());
                        break;
                }
            }
            catch (Exception ex)
            {
                // 捕获前端消息处理中的异常并弹窗
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
        /// 弹出本地图片选择对话框并将 Logo 转为 Base64 发送至前端
        /// </summary>
        private void SelectLogoImage()
        {
            // 在 STA 主线程打开 OpenFileDialog
            SafeInvoke(() =>
            {
                // 实例化 OpenFileDialog 对话框对象
                using var dialog = new OpenFileDialog();

                // 限制支持的图片文件后缀格式
                dialog.Filter = "图片文件 (*.png;*.jpg;*.jpeg;*.gif;*.bmp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp";

                // 设置对话框标题
                dialog.Title = "选择企业 Logo 图片";

                // 如果用户确认选择了图片文件
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // 读取所选图片的字节数组
                        byte[] imageBytes = File.ReadAllBytes(dialog.FileName);

                        // 获取文件扩展名并转小写
                        string ext = Path.GetExtension(dialog.FileName).ToLower().TrimStart('.');

                        // 处理 jpg 命名格式
                        if (ext == "jpg") ext = "jpeg";

                        // 组装 Data URI Base64 格式串
                        string base64Str = $"data:image/{ext};base64,{Convert.ToBase64String(imageBytes)}";

                        // 构造回发给前端更新 Logo 的 JSON 消息
                        var msg = new { action = "setLogo", logoBase64 = base64Str };

                        // 发送 Base64 图片消息给 Vue 3 预览
                        _webView.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(msg));
                    }
                    catch (Exception ex)
                    {
                        // 读取转换图片失败提示
                        MessageBox.Show($"读取图片失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            });
        }
    }
}
