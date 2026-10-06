using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的箱柜直推云端系统图库工作台窗口
    /// 遵循 Convergence Cognitive Architecture 规范与每 3 行至少 1 行中文注释
    /// </summary>
    public class CloudEBoxSyncForm : Form
    {
        // 声明 WebView2 浏览器组件
        private readonly WebView2 _webView;

        /// <summary>
        /// 构造函数：初始化窗体与 WebView2
        /// </summary>
        public CloudEBoxSyncForm()
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
            // 设置窗口标题
            this.Text = "推送箱柜至 DrawCode 云端系统图库";
            // 设置窗体初始尺寸 (支持表格宽屏展示)
            this.ClientSize = new Size(720, 680);
            // 屏幕居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;
            // 允许自由调整大小以便浏览更多箱柜
            this.FormBorderStyle = FormBorderStyle.Sizable;
            // 限制最小尺寸
            this.MinimumSize = new Size(600, 500);
            // 窗体暗色背景
            this.BackColor = Color.FromArgb(15, 23, 42);
        }

        /// <summary>
        /// 初始化 WebView2 控件布局
        /// </summary>
        private void InitializeWebViewControl()
        {
            _webView.Dock = DockStyle.Fill;
            this.Controls.Add(_webView);
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载初始化并导航到 ebox_sync.html
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                string dataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAddInDemo", "WebView2Data");
                Directory.CreateDirectory(dataFolder);

                var env = await CoreWebView2Environment.CreateAsync(null, dataFolder);
                await _webView.EnsureCoreWebView2Async(env);

                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string htmlPath = Path.Combine(baseDir, "Resources", "ebox_sync.html");
                if (!File.Exists(htmlPath))
                {
                    string fallback = Path.Combine(Directory.GetCurrentDirectory(), "Resources", "ebox_sync.html");
                    if (File.Exists(fallback)) htmlPath = fallback;
                }

                _webView.Source = new Uri(htmlPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化 WebView2 失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 响应前端 WebMessage 通信
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

                // 动作 1：提取当前工作簿箱柜数据并汇总
                if (action == "loadCabinets")
                {
                    dynamic? activeWb = ExcelDnaSafeAccessor.GetApplication()?.ActiveWorkbook;
                    string wbName = activeWb != null ? Convert.ToString(activeWb.Name) : "未命名工作簿";

                    // 读取绑定的项目信息
                    var bound = ExcelServices.GetBoundProject((object?)activeWb);
                    int boundProjId = bound.ProjectId;
                    string boundProjName = bound.ProjectName;
                    int boundGroupId = bound.GroupId;

                    // 结构化提取全部箱柜（符合 2-1 序号规范）
                    var cabinets = ExcelServices.ExtractAllCabinetsForCloudSync((object?)activeWb);

                    var res = new
                    {
                        action = "loadCabinetsResult",
                        workbookName = wbName,
                        projectId = boundProjId,
                        projectName = boundProjName,
                        currentGroupId = ExcelServices.CurrentGroupId,
                        currentGroupName = ExcelServices.CurrentGroupName,
                        cabinets = cabinets
                    };
                    PostWebMessageSafe(JsonSerializer.Serialize(res));
                    return;
                }

                // 动作 2：唤起项目绑定窗口
                if (action == "openProjectBind")
                {
                    ExcelServices.ShowProjectBindDialog();
                    return;
                }

                // 动作 3：执行一键推送云端（包含允许确认后覆盖）
                if (action == "syncToCloud")
                {
                    int projectId = root.GetProperty("projectId").GetInt32();
                    bool isCover = root.GetProperty("isCover").GetBoolean();
                    int produceOrder = root.TryGetProperty("produceOrder", out var poProp) ? poProp.GetInt32() : 1;
                    if (produceOrder < 1) produceOrder = 1;

                    // 从当前工作簿重新提取箱柜数据
                    dynamic? activeWb = ExcelDnaSafeAccessor.GetApplication()?.ActiveWorkbook;
                    var cabinets = ExcelServices.ExtractAllCabinetsForCloudSync((object?)activeWb);

                    if (cabinets.Count == 0)
                    {
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "syncResult",
                            success = false,
                            message = "未在当前工作簿中检测到任何有效箱柜！"
                        }));
                        return;
                    }

                    // 注入统一项目 ID
                    foreach (var c in cabinets)
                    {
                        c.ProjectId = projectId;
                    }

                    // 构造导入包
                    var importDto = new CloudImportRequestDto
                    {
                        EBoxs = cabinets,
                        IsCover = isCover, // 遵循用户指令：允许确认后覆盖
                        ProduceOrder = produceOrder
                    };

                    // 调用 DrawCodeApiClient 推送
                    var syncResult = await DrawCodeApiClient.ImportExcelDataAsync(importDto);

                    var res = new
                    {
                        action = "syncResult",
                        success = syncResult.Success,
                        message = syncResult.Message,
                        eBoxCount = syncResult.EBoxCount,
                        eTwoCount = syncResult.ETwoCount
                    };
                    PostWebMessageSafe(JsonSerializer.Serialize(res));
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
