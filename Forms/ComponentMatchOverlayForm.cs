using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;
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
        // 当前 D 列已有的型号内容
        public string CurrentModel { get; set; } = string.Empty;
        // 当前 G 列已有的单价或公式内容
        public string CurrentPrice { get; set; } = string.Empty;
        // 当前所属品牌
        public string Brand { get; set; } = string.Empty;
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

        // 搜索请求序号计数器 (防止快速输入时异步结果乱序覆盖)
        private long _searchReqCounter = 0;
        private long _latestSearchReqId = 0;

        /// <summary>
        /// 在指定的 Excel 活动单元格下方精准定位并展示智能下拉框
        /// </summary>
        public void ShowAtCell(
            dynamic activeCell,
            List<ComponentApiDto>? initialItems,
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

                // 若 WebView2 已经就绪，立即推送初始候选数据或触发后台异步加载
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
        /// 推送初始候选数据至前端界面 (若无初始数据则触发后台异步拉取，绝不阻塞 UI 线程)
        /// </summary>
        private void PushInitialCandidates()
        {
            var activeMustRules = (_filterConfig.MustContainRules ?? new List<MustContainRule>())
                .Where(r => r.Enabled && !string.IsNullOrWhiteSpace(r.Keyword))
                .Select(r => r.Keyword.Trim())
                .ToList();

            // 若已有初始数据直接推送到前端
            if (_pendingInitialItems != null && _pendingInitialItems.Count > 0)
            {
                PostMessageToWeb(new
                {
                    action = "initCandidates",
                    items = _pendingInitialItems,
                    cellParams = _cellParams,
                    filterBrand = _filterConfig.SelectedBrand ?? string.Empty,
                    activeMustRules,
                    dataSource = _filterConfig.DataSource ?? "cloud",
                    loading = false
                });
                return;
            }

            // 先推送上下文并让前端展示 loading 状态
            PostMessageToWeb(new
            {
                action = "initCandidates",
                items = new List<ComponentApiDto>(),
                cellParams = _cellParams,
                filterBrand = _filterConfig.SelectedBrand ?? string.Empty,
                activeMustRules,
                dataSource = _filterConfig.DataSource ?? "cloud",
                loading = true
            });

            // 在后台工作线程异步拉取初始候选数据
            long reqId = Interlocked.Increment(ref _searchReqCounter);
            _latestSearchReqId = reqId;

            var cp = _cellParams;
            var fc = _filterConfig;

            Task.Run(async () =>
            {
                try
                {
                    List<ComponentApiDto> items;
                    bool isPersonal = string.Equals(fc.DataSource, "personal", StringComparison.OrdinalIgnoreCase);
                    if (isPersonal)
                    {
                        // 从本地 SQLite 个人物料库高速检索
                        items = PersonalComponentDbService.SearchComponents(
                            null,
                            cp.Name,
                            cp.Current,
                            cp.Pole,
                            cp.TripMode,
                            fc.SelectedBrand,
                            fc.MustContainRules
                        );
                    }
                    else
                    {
                        // 异步调用云端商城 WebAPI 检索
                        items = await ComponentApiClient.SearchComponentsAsync(
                            null,
                            cp.Name,
                            cp.Current,
                            cp.Pole,
                            cp.TripMode,
                            fc.SelectedBrand,
                            fc.MustContainRules
                        ).ConfigureAwait(false);
                    }

                    // 校验请求版本，丢弃过期结果
                    if (reqId == Interlocked.Read(ref _latestSearchReqId))
                    {
                        SafeInvoke(() =>
                        {
                            PostMessageToWeb(new
                            {
                                action = "searchResult",
                                items,
                                reqId
                            });
                        });
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[ComponentMatchOverlayForm] 后台拉取初始候选数据异常: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 集中处理前端 Vue 3 发来的操作指令 (全异步非阻塞)
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

                    // 2. 即时模糊搜索 (全异步非阻塞 + 请求防竞态版本保护 + 支持个人库与云端分流 + 临时必含规则覆盖)
                    case "searchKeyword":
                        string kw = root.TryGetProperty("keyword", out var kwProp) ? kwProp.GetString() ?? "" : "";
                        // 自增请求计数器，防止快速打字异步响应乱序覆盖
                        long currentReqId = Interlocked.Increment(ref _searchReqCounter);
                        _latestSearchReqId = currentReqId;

                        var searchCp = _cellParams;
                        var searchFc = _filterConfig;

                        // 解析前端当前过滤管道传入的最新必含规则 (支持用户快捷删除/编辑后的临时覆盖)
                        List<MustContainRule> effectiveMustRules = searchFc.MustContainRules;
                        if (root.TryGetProperty("overrideMustRules", out var overrideProp) && overrideProp.ValueKind == JsonValueKind.Array)
                        {
                            // 构建临时生效的必含规则集合 (模式 B：临时过滤不直接篡改全局配置文件)
                            var customRules = new List<MustContainRule>();
                            foreach (var elem in overrideProp.EnumerateArray())
                            {
                                // 提取非空白关键字文本
                                string ruleKw = elem.GetString() ?? string.Empty;
                                if (!string.IsNullOrWhiteSpace(ruleKw))
                                {
                                    // 添加为已启用的临时必含约束项
                                    customRules.Add(new MustContainRule { Keyword = ruleKw.Trim(), Enabled = true });
                                }
                            }
                            effectiveMustRules = customRules;
                        }

                        Task.Run(async () =>
                        {
                            try
                            {
                                List<ComponentApiDto> searchResults;
                                bool isPersonal = string.Equals(searchFc.DataSource, "personal", StringComparison.OrdinalIgnoreCase);
                                if (isPersonal)
                                {
                                    // 路由到本地 SQLite 个人物料库执行模糊查询 (内置无匹配时自动降级放宽约束，支持动态必含规则)
                                    searchResults = PersonalComponentDbService.SearchComponents(
                                        kw,
                                        searchCp.Name,
                                        searchCp.Current,
                                        searchCp.Pole,
                                        searchCp.TripMode,
                                        searchFc.SelectedBrand,
                                        effectiveMustRules
                                    );
                                }
                                else
                                {
                                    // 异步调用云端商城 WebAPI 执行动态必含规则约束检索
                                    searchResults = await ComponentApiClient.SearchComponentsAsync(
                                        kw,
                                        searchCp.Name,
                                        searchCp.Current,
                                        searchCp.Pole,
                                        searchCp.TripMode,
                                        searchFc.SelectedBrand,
                                        effectiveMustRules
                                    ).ConfigureAwait(false);
                                }

                                // 仅当返回结果是最新一次搜索时才推送到前端，彻底消除数据跳变
                                if (currentReqId == Interlocked.Read(ref _latestSearchReqId))
                                {
                                    SafeInvoke(() =>
                                    {
                                        PostMessageToWeb(new
                                        {
                                            action = "searchResult",
                                            items = searchResults,
                                            reqId = currentReqId
                                        });
                                    });
                                }
                            }
                            catch (Exception ex)
                            {
                                LogHelper.WriteLog($"[ComponentMatchOverlayForm] 异步模糊搜索异常: {ex.Message}");
                            }
                        });
                        break;

                    // 2.1 用户在悬浮窗中点击“保存规则”将当前必含规则持久化为全局默认配置 (模式 B 双态分流)
                    case "saveMustRules":
                        if (root.TryGetProperty("rules", out var rulesProp) && rulesProp.ValueKind == JsonValueKind.Array)
                        {
                            // 构建需持久化保存的规则列表
                            var updatedRules = new List<MustContainRule>();
                            foreach (var elem in rulesProp.EnumerateArray())
                            {
                                // 过滤并提取关键字
                                string ruleKw = elem.GetString() ?? string.Empty;
                                if (!string.IsNullOrWhiteSpace(ruleKw))
                                {
                                    // 封装标准规则实体
                                    updatedRules.Add(new MustContainRule { Keyword = ruleKw.Trim(), Enabled = true });
                                }
                            }

                            // 更新当前运行时的过滤配置对象
                            _filterConfig.MustContainRules = updatedRules;

                            // 将最新配置持久化写入磁盘 JSON 配置文件
                            ExcelServices.SaveComponentMatchFilterConfig(_filterConfig);

                            // 回发确认消息通知前端更新快照与状态
                            PostMessageToWeb(new
                            {
                                action = "mustRulesSaved",
                                success = true
                            });
                        }
                        break;

                    // 3. 用户确认选择某一条物料 -> 先隐藏窗口，后台执行 Excel 回填
                    case "selectComponent":
                        SafeInvoke(this.Hide);
                        if (root.TryGetProperty("item", out var itemProp))
                        {
                            var selectedItem = JsonSerializer.Deserialize<ComponentApiDto>(itemProp.GetRawText(), JsonOptions);
                            if (selectedItem != null && _targetCell != null)
                            {
                                // 调用业务服务层执行单元格所在行多列回填
                                ExcelServices.FillSelectedComponentToActiveRow(selectedItem, _targetCell);
                            }
                        }
                        break;

                    // 3.1 用户请求加载当前物料的配套附件列表 (支持个人库与云端分流)
                    case "getAttachments":
                        string brandToQuery = _filterConfig.SelectedBrand ?? _cellParams.Brand ?? string.Empty;
                        string nameToQuery = _cellParams.Name ?? string.Empty;
                        string modelToQuery = _cellParams.CurrentModel ?? string.Empty;

                        Task.Run(async () =>
                        {
                            try
                            {
                                List<ComponentApiDto> attachmentList;
                                bool isPersonal = string.Equals(_filterConfig.DataSource, "personal", StringComparison.OrdinalIgnoreCase);
                                if (isPersonal)
                                {
                                    // 从本地 SQLite 查询配套附件
                                    attachmentList = PersonalComponentDbService.GetAttachments(brandToQuery, nameToQuery, modelToQuery);
                                }
                                else
                                {
                                    // 从云端商城 WebAPI 异步查询配套附件
                                    attachmentList = await ComponentApiClient.GetAttachmentsAsync(brandToQuery, nameToQuery, modelToQuery).ConfigureAwait(false);
                                }

                                SafeInvoke(() =>
                                {
                                    PostMessageToWeb(new
                                    {
                                        action = "attachmentsResult",
                                        items = attachmentList,
                                        currentModel = modelToQuery,
                                        brand = brandToQuery,
                                        name = nameToQuery
                                    });
                                });
                            }
                            catch (Exception ex)
                            {
                                LogHelper.WriteLog($"[ComponentMatchOverlayForm] 异步拉取配套附件异常: {ex.Message}");
                            }
                        });
                        break;

                    // 3.2 用户选定配套附件 -> 先隐藏窗口，后台拼接型号并累加价格公式
                    case "selectAttachment":
                        SafeInvoke(this.Hide);
                        if (root.TryGetProperty("item", out var attachItemProp))
                        {
                            var selectedAttach = JsonSerializer.Deserialize<ComponentApiDto>(attachItemProp.GetRawText(), JsonOptions);
                            int quantity = 1;
                            if (root.TryGetProperty("quantity", out var qtyProp) && qtyProp.TryGetInt32(out int q))
                            {
                                quantity = q > 0 ? q : 1;
                            }

                            if (selectedAttach != null && _targetCell != null)
                            {
                                ExcelServices.FillSelectedAttachmentToActiveRow(selectedAttach, _targetCell, quantity);
                            }
                        }
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
