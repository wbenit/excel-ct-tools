using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的云端项目智能探测与绑定确认窗口
    /// 遵循 Convergence Cognitive Architecture 规范与每 3 行至少 1 行中文注释
    /// </summary>
    public class ProjectBindForm : Form
    {
        // 声明嵌入式 WebView2 浏览器控件
        private readonly WebView2 _webView;

        /// <summary>
        /// 构造函数：初始化窗口外观、事件与 WebView2 控件
        /// </summary>
        public ProjectBindForm()
        {
            _webView = new WebView2();
            InitializeFormProperties();
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体几何尺寸与基本外观样式
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗口标题文本
            this.Text = "绑定 DrawCode 云端工程项目";
            // 设置窗体初始尺寸
            this.ClientSize = new Size(540, 620);
            // 屏幕居中弹出展示
            this.StartPosition = FormStartPosition.CenterScreen;
            // 固定边框禁止拖动尺寸
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            // 隐藏最大化按钮
            this.MaximizeBox = false;
            // 允许最小化
            this.MinimizeBox = true;
            // 设置窗体背景色为暗夜蓝
            this.BackColor = Color.FromArgb(15, 23, 42);
        }

        /// <summary>
        /// 初始化 WebView2 控件布局并挂接 Load 事件
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 填满整个窗体视口
            _webView.Dock = DockStyle.Fill;
            this.Controls.Add(_webView);
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体异步加载：创建安全环境并加载 project_bind.html
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 计算用户可写的独立缓存目录，防止写入 Program Files 权限拒绝
                string dataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAddInDemo", "WebView2Data");
                Directory.CreateDirectory(dataFolder);

                // 创建 WebView2 运行时环境
                var env = await CoreWebView2Environment.CreateAsync(null, dataFolder);
                await _webView.EnsureCoreWebView2Async(env);

                // 注册 Web 消息回调
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 计算 Resources/project_bind.html 文件路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string htmlPath = Path.Combine(baseDir, "Resources", "project_bind.html");
                if (!File.Exists(htmlPath))
                {
                    string fallback = Path.Combine(Directory.GetCurrentDirectory(), "Resources", "project_bind.html");
                    if (File.Exists(fallback)) htmlPath = fallback;
                }

                // 加载页面
                _webView.Source = new Uri(htmlPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化 WebView2 失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 响应前端发来的 JSON 交互消息
        /// </summary>
        private async void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string json = e.TryGetWebMessageAsString();
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;
                string action = actionProp.GetString() ?? string.Empty;

                // 动作 1：获取初始化绑定上下文信息
                if (action == "getInitContext")
                {
                    // 提取当前活动工作簿
                    dynamic? activeWb = ExcelDnaSafeAccessor.GetApplication()?.ActiveWorkbook;
                    string wbName = activeWb != null ? Convert.ToString(activeWb.Name) : "未命名工作簿";
                    // 智能提取候选项目名
                    string candidateName = ExcelServices.DetectCandidateProjectName((object?)activeWb);
                    // 读取已绑定的项目信息
                    var bound = ExcelServices.GetBoundProject((object?)activeWb);
                    int boundId = bound.ProjectId;
                    string boundName = bound.ProjectName;
                    int boundGroupId = bound.GroupId;

                    var initData = new
                    {
                        action = "initContextResult",
                        workbookName = wbName,
                        candidateName = candidateName,
                        currentGroupId = ExcelServices.CurrentGroupId,
                        currentGroupName = ExcelServices.CurrentGroupName,
                        boundProjectId = boundId,
                        boundProjectName = boundName
                    };

                    PostWebMessageSafe(JsonSerializer.Serialize(initData));
                    return;
                }

                // 动作 2：查询云端项目列表
                if (action == "searchProjects")
                {
                    string keyword = root.TryGetProperty("keyword", out var kwProp) ? (kwProp.GetString() ?? "") : "";
                    var (success, msg, list) = await DrawCodeApiClient.GetProjectsAsync(keyword);

                    var res = new
                    {
                        action = "searchProjectsResult",
                        success = success,
                        message = msg,
                        projects = list
                    };
                    PostWebMessageSafe(JsonSerializer.Serialize(res));
                    return;
                }

                // 动作 3：确认绑定项目到当前工作簿
                if (action == "bindProject")
                {
                    int projectId = root.GetProperty("projectId").GetInt32();
                    string projectName = root.GetProperty("projectName").GetString() ?? "";
                    int groupId = ExcelServices.CurrentGroupId;

                    bool ok = ExcelServices.BindProjectToWorkbook(projectId, projectName, groupId);
                    var res = new
                    {
                        action = "bindProjectResult",
                        success = ok,
                        message = ok ? $"成功绑定到云端项目：{projectName}" : "绑定项目失败！"
                    };
                    PostWebMessageSafe(JsonSerializer.Serialize(res));

                    if (ok)
                    {
                        await System.Threading.Tasks.Task.Delay(800);
                        SafeInvoke(() => this.Close());
                    }
                    return;
                }
            }
            catch (Exception ex)
            {
                var err = new { action = "error", message = $"处理异常: {ex.Message}" };
                PostWebMessageSafe(JsonSerializer.Serialize(err));
            }
        }

        /// <summary>
        /// 跨线程安全向 WebView2 发送 JSON 消息
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
        /// 跨线程安全调度 UI 执行
        /// </summary>
        private void SafeInvoke(Action action)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            if (this.InvokeRequired) this.Invoke(action);
            else action();
        }
    }
}
