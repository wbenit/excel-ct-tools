using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Models;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 当前行单元格参数上下文传递模型
    /// </summary>
    public class CellParamsContext
    {
        public string Name { get; set; } = string.Empty;
        public string Current { get; set; } = string.Empty;
        public string Pole { get; set; } = string.Empty;
        public string TripMode { get; set; } = string.Empty;
    }

    /// <summary>
    /// 基于 WebView2 + Vue 3 的“点击查询”单元格贴合智能联想下拉悬浮窗口
    /// </summary>
    public class ComponentMatchOverlayForm : Form
    {
        // 声明 WebView2 浏览器控件实例
        private readonly WebView2 _webView;

        // 当前绑定的 Excel 活动单元格 COM 对象
        private dynamic? _targetCell;

        // 当前生效的过滤管道配置 (品牌 + 必含字段规则)
        private ComponentMatchFilterConfig _filterConfig = new ComponentMatchFilterConfig();

        // 当前行的参数上下文
        private CellParamsContext _cellParams = new CellParamsContext();

        // 暂存的初始候选物料列表
        private List<ComponentApiDto> _pendingInitialItems = new List<ComponentApiDto>();

        // WebView2 是否已完成初始化
        private bool _isWebReady = false;

        // JSON 序列化配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        /// <summary>
        /// 构造函数: 初始化窗口几何属性与 WebView2 控件
        /// </summary>
        public ComponentMatchOverlayForm()
        {
            _webView = new WebView2();

            // 配置窗体外观与尺寸 (480x340 像素)
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.Size = new Size(480, 340);
            this.StartPosition = FormStartPosition.Manual;
            this.BackColor = Color.White;

            // 控件充满窗体
            _webView.Dock = DockStyle.Fill;
            this.Controls.Add(_webView);

            // 注册生命周期与失焦事件
            this.Load += OnFormLoadAsync;
            this.Deactivate += OnOverlayDeactivate;
        }

        /// <summary>
        /// 异步加载 WebView2 环境并导航至 component_match_overlay.html
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 获取专属用户缓存数据目录
                string userDataDir = Path.Combine(Tool.GetAppDataDirectory(), "WebView2_MatchOverlay");
                var env = await CoreWebView2Environment.CreateAsync(null, userDataDir);
                await _webView.EnsureCoreWebView2Async(env);

                // 配置环境参数
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

                // 注册 Web 消息监听
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 获取前端 HTML 路径
                string appDir = Tool.GetAppDirectory();
                string htmlPath = Path.Combine(appDir, "Resources", "component_match_overlay.html");
                if (!File.Exists(htmlPath))
                {
                    htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "component_match_overlay.html");
                }

                if (File.Exists(htmlPath))
                {
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ComponentMatchOverlayForm 初始化异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 在指定的 Excel 活动单元格下方精准定位并展示智能下拉框
        /// </summary>
        public void ShowAtCell(
            dynamic activeCell,
            List<ComponentApiDto> initialItems,
            CellParamsContext cellParams,
            ComponentMatchFilterConfig filterConfig)
        {
            if (activeCell == null) return;

            try
            {
                _targetCell = activeCell;
                _cellParams = cellParams ?? new CellParamsContext();
                _filterConfig = filterConfig ?? ExcelServices.LoadComponentMatchFilterConfig();
                _pendingInitialItems = initialItems ?? new List<ComponentApiDto>();

                // 计算单元格屏幕像素矩形区域
                Rectangle cellRect = CalculateCellScreenRect(activeCell);

                // 将悬浮窗定位在单元格正下方 (对齐左侧)
                int targetX = cellRect.Left;
                int targetY = cellRect.Bottom + 2;

                // 获取当前屏幕可用工作区域，防止超出屏幕边缘
                Screen currentScreen = Screen.FromPoint(new Point(targetX, targetY));
                Rectangle workingArea = currentScreen.WorkingArea;

                // 若下方空间不足，则向上弹出
                if (targetY + this.Height > workingArea.Bottom)
                {
                    targetY = Math.Max(workingArea.Top, cellRect.Top - this.Height - 2);
                }
                // 若右侧超出屏幕则向左靠拢
                if (targetX + this.Width > workingArea.Right)
                {
                    targetX = Math.Max(workingArea.Left, workingArea.Right - this.Width - 10);
                }

                this.Location = new Point(targetX, targetY);

                // 显示窗口并激活
                if (!this.Visible)
                {
                    this.Show();
                }
                this.BringToFront();

                // 若 WebView2 已经就绪，立即推送初始候选数据
                if (_isWebReady)
                {
                    PushInitialCandidates();
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ShowAtCell 计算定位异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 推送初始候选数据至前端界面
        /// </summary>
        private void PushInitialCandidates()
        {
            var activeMustRules = (_filterConfig.MustContainRules ?? new List<MustContainRule>())
                .Where(r => r.Enabled && !string.IsNullOrWhiteSpace(r.Keyword))
                .Select(r => r.Keyword.Trim())
                .ToList();

            PostMessageToWeb(new
            {
                action = "initCandidates",
                items = _pendingInitialItems,
                cellParams = _cellParams,
                filterBrand = _filterConfig.SelectedBrand ?? string.Empty,
                activeMustRules
            });
        }

        /// <summary>
        /// 集中处理前端 Vue 3 发来的操作指令
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string rawJson = e.TryGetWebMessageAsString();
                if (string.IsNullOrWhiteSpace(rawJson)) return;

                using var doc = JsonDocument.Parse(rawJson);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;

                string action = actionProp.GetString() ?? string.Empty;

                switch (action)
                {
                    // 1. 前端页面加载完成
                    case "overlayReady":
                        _isWebReady = true;
                        PushInitialCandidates();
                        break;

                    // 2. 即时模糊搜索
                    case "searchKeyword":
                        string kw = root.TryGetProperty("keyword", out var kwProp) ? kwProp.GetString() ?? "" : "";
                        var searchResults = ComponentApiClient.SearchComponents(
                            kw,
                            _cellParams.Name,
                            _cellParams.Current,
                            _cellParams.Pole,
                            _cellParams.TripMode,
                            _filterConfig.SelectedBrand,
                            _filterConfig.MustContainRules
                        );

                        PostMessageToWeb(new
                        {
                            action = "searchResult",
                            items = searchResults
                        });
                        break;

                    // 3. 用户确认选择某一条物料 -> 执行 Excel 回填并清除底色
                    case "selectComponent":
                        if (root.TryGetProperty("item", out var itemProp))
                        {
                            var selectedItem = JsonSerializer.Deserialize<ComponentApiDto>(itemProp.GetRawText(), JsonOptions);
                            if (selectedItem != null && _targetCell != null)
                            {
                                // 调用业务服务层执行单元格所在行多列回填
                                ExcelServices.FillSelectedComponentToActiveRow(selectedItem, _targetCell);
                            }
                        }
                        // 回填完成后隐藏悬浮窗
                        SafeInvoke(this.Hide);
                        break;

                    // 4. 关闭悬浮窗
                    case "closeOverlay":
                        SafeInvoke(this.Hide);
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ComponentMatchOverlayForm 消息分发异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 窗体失去焦点时自动隐藏
        /// </summary>
        private void OnOverlayDeactivate(object? sender, EventArgs e)
        {
            try
            {
                // 失去焦点时自动平滑隐藏，不干扰 Excel 操作
                this.Hide();
            }
            catch { }
        }

        /// <summary>
        /// 计算活动单元格在屏幕上的绝对像素坐标矩形
        /// </summary>
        private Rectangle CalculateCellScreenRect(dynamic targetCell)
        {
            try
            {
                dynamic app = targetCell.Application;
                dynamic win = app.ActiveWindow;

                double cellLeft = Convert.ToDouble(targetCell.Left);
                double cellTop = Convert.ToDouble(targetCell.Top);
                double cellWidth = Convert.ToDouble(targetCell.Width);
                double cellHeight = Convert.ToDouble(targetCell.Height);

                // 优先使用 Panes 坐标精准换算
                if (win.Panes != null && win.Panes.Count > 0)
                {
                    try
                    {
                        dynamic pane = win.ActivePane ?? win.Panes[1];
                        int px1 = pane.PointsToScreenPixelsX((int)cellLeft);
                        int py1 = pane.PointsToScreenPixelsY((int)cellTop);
                        int px2 = pane.PointsToScreenPixelsX((int)(cellLeft + cellWidth));
                        int py2 = pane.PointsToScreenPixelsY((int)(cellTop + cellHeight));

                        if (px1 > 0 && py1 > 0 && px2 > px1 && py2 > py1)
                        {
                            return new Rectangle(px1, py1, px2 - px1, py2 - py1);
                        }
                    }
                    catch { }
                }

                // 备选 Window 坐标换算
                int wx1 = win.PointsToScreenPixelsX((int)cellLeft);
                int wy1 = win.PointsToScreenPixelsY((int)cellTop);
                int wx2 = win.PointsToScreenPixelsX((int)(cellLeft + cellWidth));
                int wy2 = win.PointsToScreenPixelsY((int)(cellTop + cellHeight));

                return new Rectangle(wx1, wy1, Math.Max(wx2 - wx1, 60), Math.Max(wy2 - wy1, 24));
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"计算单元格屏幕坐标异常: {ex.Message}");
                return new Rectangle(Cursor.Position.X, Cursor.Position.Y + 20, 200, 30);
            }
        }

        /// <summary>
        /// 向 WebView2 前端推送消息
        /// </summary>
        private void PostMessageToWeb(object data)
        {
            SafeInvoke(() =>
            {
                if (_webView?.CoreWebView2 != null)
                {
                    string json = JsonSerializer.Serialize(data, JsonOptions);
                    _webView.CoreWebView2.PostWebMessageAsString(json);
                }
            });
        }

        /// <summary>
        /// 跨线程安全调度
        /// </summary>
        private void SafeInvoke(Action action)
        {
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
