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
            this.ClientSize = new Size(1150, 680);
            // 屏幕居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;
            // 允许自由调整大小以便浏览更多箱柜
            this.FormBorderStyle = FormBorderStyle.Sizable;
            // 限制最小尺寸
            this.MinimumSize = new Size(800, 500);
            // 窗体背景色：与图二保持一致的清爽浅色背景 (#f8fafc)
            this.BackColor = Color.FromArgb(248, 250, 252);
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

        // 记录窗口折叠前的展开高度，便于展开时准确还原
        private int _expandedHeight = 680;
        // 当前窗口是否处于 100px 折叠收起状态
        private bool _isCollapsed = false;

        /// <summary>
        /// 响应前端 WebMessage 通信
        /// </summary>
        private async void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 双轨安全提取消息文本：优先作为原生 JSON 提取，若为空再尝试纯字符串提取
                string json = string.Empty;
                try { json = e.WebMessageAsJson; } catch { }
                // 若原生 JSON 为空，则降级尝试作为纯字符串提取
                if (string.IsNullOrWhiteSpace(json))
                {
                    try { json = e.TryGetWebMessageAsString(); } catch { }
                }
                // 校验消息体有效性，若皆为空则直接忽略退出
                if (string.IsNullOrWhiteSpace(json)) return;

                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;
                string action = actionProp.GetString() ?? string.Empty;

                // 动作 0：响应前端窗口折叠与展开切换（遵循用户要求可折叠至 100px 高）
                if (action == "toggleCollapse")
                {
                    bool collapse = root.TryGetProperty("collapsed", out var cProp) && cProp.GetBoolean();
                    SafeInvoke(() =>
                    {
                        if (collapse)
                        {
                            // 记录折叠前的展开高度
                            if (!_isCollapsed)
                            {
                                _expandedHeight = this.Height;
                            }
                            _isCollapsed = true;
                            // 临时放开最小尺寸约束
                            this.MinimumSize = new Size(480, 80);
                            // 折叠窗口至 100px 高
                            this.Height = 100;
                        }
                        else
                        {
                            // 展开恢复正常尺寸
                            _isCollapsed = false;
                            // 恢复常规最小尺寸约束
                            this.MinimumSize = new Size(600, 500);
                            // 恢复记忆的展开高度
                            this.Height = _expandedHeight > 120 ? _expandedHeight : 680;
                        }
                    });
                    return;
                }

                // 动作 1：提取当前工作簿箱柜数据并汇总 (支持根据用户选中的分类明细生成，并获取项目最大批次+1)
                if (action == "loadCabinets")
                {
                    dynamic? activeWb = ExcelDnaSafeAccessor.GetApplication()?.ActiveWorkbook;
                    string wbName = activeWb != null ? Convert.ToString(activeWb.Name) : "未命名工作簿";

                    // 读取绑定的项目信息
                    var bound = ExcelServices.GetBoundProject((object?)activeWb);
                    int boundProjId = bound.ProjectId;
                    string boundProjName = bound.ProjectName;
                    int boundGroupId = bound.GroupId;

                    // 1. 获取当前工作簿登记在册的所有分类明细工作表名称列表
                    var allCategories = ExcelServices.GetProjectCategorySheetNamesList((object?)activeWb);

                    // 2. 解析前端传入的用户勾选分类列表，若未传入则默认全选
                    List<string> selectedCategories = new List<string>();
                    if (root.TryGetProperty("selectedCategories", out var scElem) && scElem.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var item in scElem.EnumerateArray())
                        {
                            string? cat = item.GetString();
                            if (!string.IsNullOrEmpty(cat)) selectedCategories.Add(cat);
                        }
                    }
                    if (selectedCategories.Count == 0)
                    {
                        selectedCategories = new List<string>(allCategories);
                    }

                    // 3. 结构化提取被选中的分类明细箱柜（符合 2-1 序号规范）
                    var cabinets = ExcelServices.ExtractAllCabinetsForCloudSync((object?)activeWb, true, selectedCategories);

                    // 4. 异步从云端获取该项目的最大生产批次并计算最大批次 + 1
                    int maxProduceOrder = 0;
                    if (boundProjId > 0)
                    {
                        maxProduceOrder = await DrawCodeApiClient.GetProjectMaxProduceOrderAsync(boundProjId);
                    }
                    int recommendProduceOrder = maxProduceOrder > 0 ? (maxProduceOrder + 1) : 1;

                    var res = new
                    {
                        action = "loadCabinetsResult",
                        workbookName = wbName,
                        projectId = boundProjId,
                        projectName = boundProjName,
                        currentGroupId = ExcelServices.CurrentGroupId,
                        currentGroupName = ExcelServices.CurrentGroupName,
                        allCategories = allCategories,
                        selectedCategories = selectedCategories,
                        maxProduceOrder = maxProduceOrder,
                        recommendProduceOrder = recommendProduceOrder,
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

                // 动作 2.5：根据用户选择的分类明细，实时拉取云端动态字段并刷新【云端箱柜属性】工作表
                if (action == "syncCustomPropsSheet")
                {
                    // 获取当前活动的 Excel 工作簿对象
                    dynamic? activeWb = ExcelDnaSafeAccessor.GetApplication()?.ActiveWorkbook;
                    // 读取绑定的项目信息
                    var bound = ExcelServices.GetBoundProject((object?)activeWb);
                    int boundProjId = bound.ProjectId;

                    // 若未绑定项目则终止并提示
                    if (boundProjId <= 0)
                    {
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "syncCustomPropsSheetResult",
                            success = false,
                            message = "当前工作簿尚未绑定云端项目，请先绑定项目！"
                        }));
                        return;
                    }

                    // 解析前端选中的分类明细列表
                    List<string>? selectedCategories = null;
                    if (root.TryGetProperty("selectedCategories", out var scElem) && scElem.ValueKind == JsonValueKind.Array)
                    {
                        selectedCategories = new List<string>();
                        foreach (var item in scElem.EnumerateArray())
                        {
                            string? cat = item.GetString();
                            if (!string.IsNullOrEmpty(cat)) selectedCategories.Add(cat);
                        }
                    }

                    // 遵循用户指令：仅为选中的分类明细生成属性表
                    var (ok, msg) = await ExcelServices.SyncCloudCustomPropsSheetAsync(boundProjId, (object?)activeWb, selectedCategories);

                    // 重新提取属于选定分类明细的全部箱柜集合
                    var cabinets = ExcelServices.ExtractAllCabinetsForCloudSync((object?)activeWb, true, selectedCategories);

                    // 回传处理结果与最新箱柜数据列表给前端
                    PostWebMessageSafe(JsonSerializer.Serialize(new
                    {
                        action = "syncCustomPropsSheetResult",
                        success = ok,
                        message = msg,
                        cabinets = cabinets
                    }));
                    return;
                }

                // 动作 3：执行一键推送云端（包含允许确认后覆盖）
                if (action == "syncToCloud")
                {
                    int projectId = root.GetProperty("projectId").GetInt32();
                    bool isCover = root.GetProperty("isCover").GetBoolean();
                    int produceOrder = root.TryGetProperty("produceOrder", out var poProp) ? poProp.GetInt32() : 1;
                    if (produceOrder < 1) produceOrder = 1;

                    // 解析选中的分类明细列表
                    List<string>? selectedCategories = null;
                    if (root.TryGetProperty("selectedCategories", out var scElem) && scElem.ValueKind == JsonValueKind.Array)
                    {
                        selectedCategories = new List<string>();
                        foreach (var item in scElem.EnumerateArray())
                        {
                            string? cat = item.GetString();
                            if (!string.IsNullOrEmpty(cat)) selectedCategories.Add(cat);
                        }
                    }

                    // 从当前工作簿重新提取选中的箱柜数据
                    dynamic? activeWb = ExcelDnaSafeAccessor.GetApplication()?.ActiveWorkbook;
                    var cabinets = ExcelServices.ExtractAllCabinetsForCloudSync((object?)activeWb, true, selectedCategories);

                    if (cabinets.Count == 0)
                    {
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "syncResult",
                            success = false,
                            message = "未在当前工作簿中检测到任何属于选定分类的有效箱柜！"
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
                // 记录详细异常日志便于追溯
                LogHelper.WriteLog($"[CloudEBoxSyncForm] OnWebMessageReceived 异常: {ex}");
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
