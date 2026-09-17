using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
        // 当前型号内容 (汇总表为 D 列，分类明细表为 C 列)
        public string CurrentModel { get; set; } = string.Empty;
        // 当前单价或基准表价公式/数值
        public string CurrentPrice { get; set; } = string.Empty;
        // 扩展参数 1 (汇总表 X 列，分类明细表 AA 列)
        public string Param1 { get; set; } = string.Empty;
        // 扩展参数 2 (汇总表 Y 列，分类明细表 AB 列)
        public string Param2 { get; set; } = string.Empty;
        // 当前工作表是否为分类明细表 (false 为元件汇总表)
        public bool IsCategorySheet { get; set; } = false;
        // 多选品牌偏好列表
        public List<string> Brands { get; set; } = new List<string>();
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

        // 窗口是否处于“固定置顶”模式 (固定时失焦不关闭、切行不重新搜索)
        private bool _isPinned = false;

        // 是否待切入配套附件模式 (供右键一键选配附件使用)
        private bool _pendingAttachmentMode = false;

        /// <summary>
        /// 对外暴露当前窗口是否处于固定状态
        /// </summary>
        public bool IsPinned => _isPinned;

        /// <summary>
        /// 重写展示无焦点激活属性，确保弹窗时不争抢 Excel 键盘焦点，保障 Excel 原生自由就地编辑
        /// </summary>
        protected override bool ShowWithoutActivation => true;

        // Windows 原生拖拽 API 声明
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // Windows 窗口消息分发 API
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        // Windows 窗口 Z-order 与无激活显示 API
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        // 置顶窗口句柄标识
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        // 保留原尺寸
        private const uint SWP_NOSIZE = 0x0001;
        // 保留原位置
        private const uint SWP_NOMOVE = 0x0002;
        // 关键: 不激活窗口焦点
        private const uint SWP_NOACTIVATE = 0x0010;
        // 显示窗口
        private const uint SWP_SHOWWINDOW = 0x0040;

        // 标题栏按下常数标识
        private const int WM_NCLBUTTONDOWN = 0xA1;
        // 客户区命中标题栏常数
        private const int HT_CAPTION = 0x2;

        /// <summary>
        /// 开启系统级无抖动平滑窗口拖拽
        /// </summary>
        public void BeginDrag()
        {
            SafeInvoke(() =>
            {
                // 释放鼠标捕获
                ReleaseCapture();
                // 向窗体句柄发送标题栏按下消息触发拖拽
                SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            });
        }

        // JSON 序列化配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        // 静态记忆用户自定义的窗口高度 (默认 340 像素，范围 220~800)
        private static int _customHeight = 340;

        /// <summary>
        /// 构造函数: 初始化窗口几何属性与 WebView2 控件
        /// </summary>
        public ComponentMatchOverlayForm()
        {
            _webView = new WebView2();

            // 配置窗体外观与尺寸 (480 宽，高度采用用户自定义记忆值)
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.Size = new Size(480, _customHeight);
            this.StartPosition = FormStartPosition.Manual;
            this.BackColor = Color.White;

            // 控件充满窗体
            _webView.Dock = DockStyle.Fill;
            this.Controls.Add(_webView);

            // 注册生命周期与失焦事件
            this.Load += OnFormLoadAsync;
            this.Deactivate += OnOverlayDeactivate;
        }

        // 标记是否正在执行 WebView2 异步初始化，防止多次并发初始化
        private bool _isInitializing = false;

        /// <summary>
        /// 预热浮窗控件与 WebView2 运行时环境 (在后台静默就绪，消除首次点击冷启动延迟)
        /// </summary>
        public void WarmUp()
        {
            try
            {
                // 确保 WinForm 控件句柄已在 UI 线程创建
                if (!this.IsHandleCreated)
                {
                    // 强制触发底层窗口句柄创建
                    this.CreateControl();
                }

                // 若尚未就绪且未在初始化中，立即触发 WebView2 环境加载
                if (!_isWebReady && !_isInitializing)
                {
                    // 主动调用初始化函数
                    OnFormLoadAsync(this, EventArgs.Empty);
                }
            }
            catch (Exception ex)
            {
                // 记录预热异常日志
                LogHelper.WriteLog($"ComponentMatchOverlayForm 预热异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 异步加载 WebView2 环境并导航至 component_match_overlay.html
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            // 防重复初始化门控
            if (_isInitializing || _isWebReady) return;
            // 标记初始化状态
            _isInitializing = true;

            try
            {
                // 获取专属用户缓存数据目录
                string userDataDir = Path.Combine(Tool.GetAppDataDirectory(), "WebView2_MatchOverlay");
                // 异步创建环境实例
                var env = await CoreWebView2Environment.CreateAsync(null, userDataDir);
                // 确保 WebView2 控件初始化成功
                await _webView.EnsureCoreWebView2Async(env);

                // 配置环境参数
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                // 禁用底部状态栏
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

                // 注册 Web 消息监听
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 获取前端 HTML 路径
                string appDir = Tool.GetAppDirectory();
                // 拼接资源目录绝对路径
                string htmlPath = Path.Combine(appDir, "Resources", "component_match_overlay.html");
                if (!File.Exists(htmlPath))
                {
                    // 回退基准目录
                    htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "component_match_overlay.html");
                }

                // 若文件存在则导航加载页面
                if (File.Exists(htmlPath))
                {
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
                }
            }
            catch (Exception ex)
            {
                // 记录初始化异常日志
                LogHelper.WriteLog($"ComponentMatchOverlayForm 初始化异常: {ex.Message}");
            }
            finally
            {
                // 重置正在初始化标记
                _isInitializing = false;
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

            // 若当前处于固定置顶模式且窗口已在显示中，则锁定现有位置与搜索结果，绝不打扰用户连续回填
            if (_isPinned && this.Visible)
            {
                // 仅更新当前绑定的活动单元格引用
                _targetCell = activeCell;
                return;
            }

            try
            {
                // 应用用户自定义记忆高度
                this.Height = _customHeight;
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

                // 显示窗口但绝不强占 Excel 焦点
                if (!this.Visible)
                {
                    this.Show();
                }
                // 使用 SWP_NOACTIVATE 保持窗口位于最前端，但 100% 将输入焦点留在 Excel 单元格中
                SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);

                // 重置附件模式标记
                _pendingAttachmentMode = false;

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
        /// 在活动单元格下方弹出并直接切入配套附件选配模式 (供右键菜单一键选配配套附件使用)
        /// </summary>
        /// <param name="activeCell">当前选中的活动单元格 COM 句柄</param>
        /// <param name="cellParams">当前行元器件参数上下文</param>
        /// <param name="filterConfig">匹配过滤配置</param>
        public void ShowAttachmentsAtCell(
            dynamic activeCell,
            CellParamsContext cellParams,
            ComponentMatchFilterConfig filterConfig)
        {
            // 校验目标单元格有效性
            if (activeCell == null) return;

            try
            {
                // 应用用户自定义记忆高度
                this.Height = _customHeight;
                // 绑定当前活动单元格句柄
                _targetCell = activeCell;
                // 缓存参数上下文与过滤配置
                _cellParams = cellParams ?? new CellParamsContext();
                _filterConfig = filterConfig ?? ExcelServices.LoadComponentMatchFilterConfig();
                // 标记为附件模式
                _pendingAttachmentMode = true;
                _pendingInitialItems = new List<ComponentApiDto>();

                // 计算单元格屏幕像素矩形区域
                Rectangle cellRect = CalculateCellScreenRect(activeCell);

                // 将悬浮窗定位在单元格正下方
                int targetX = cellRect.Left;
                int targetY = cellRect.Bottom + 2;

                // 获取当前屏幕可用工作区域
                Screen currentScreen = Screen.FromPoint(new Point(targetX, targetY));
                Rectangle workingArea = currentScreen.WorkingArea;

                // 若下方空间不足则向上弹出
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

                // 显示窗口并置顶但不抢占输入焦点
                if (!this.Visible)
                {
                    this.Show();
                }
                SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);

                // 若前端已就绪，立即拉取配套附件并通知切入附件模式
                if (_isWebReady)
                {
                    TriggerLoadAttachments();
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ShowAttachmentsAtCell 计算定位异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 重新加载最新的全局物料匹配过滤配置并向前端推送同步 (在规则设置窗口修改并保存后自动联动)
        /// </summary>
        /// <param name="newConfig">最新的过滤配置对象 (若为 null 则自动从磁盘文件重载)</param>
        public void ReloadFilterConfig(ComponentMatchFilterConfig? newConfig = null)
        {
            SafeInvoke(() =>
            {
                // 重新读取或更新本地过滤配置
                _filterConfig = newConfig ?? ExcelServices.LoadComponentMatchFilterConfig();

                // 若前端已就绪且当前处于可见状态，立即触发最新候选拉取与数据源状态刷新
                if (_isWebReady && this.Visible)
                {
                    // 重新推送初始候选数据以刷新数据源与品牌规则
                    PushInitialCandidates();
                }
            });
        }

        /// <summary>
        /// 后台异步拉取当前行物料的配套附件并向前端推送切入附件模式指令
        /// </summary>
        private void TriggerLoadAttachments()
        {
            // 准备待查询的品牌、名称与型号
            string brandToQuery = _filterConfig.GetEffectiveBrands().FirstOrDefault() ?? _cellParams.Brands.FirstOrDefault() ?? string.Empty;
            string nameToQuery = _cellParams.Name ?? string.Empty;
            string modelToQuery = _cellParams.CurrentModel ?? string.Empty;

            // 先通知前端进入 loading 状态
            PostMessageToWeb(new
            {
                action = "autoEnterAttachmentMode",
                items = new List<ComponentApiDto>(),
                currentModel = modelToQuery,
                brand = brandToQuery,
                name = nameToQuery,
                loading = true
            });

            // 在工作线程中异步拉取附件数据
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

                    // 切回 UI 主线程推送附件模式数据
                    SafeInvoke(() =>
                    {
                        if (this.IsDisposed || !this.Visible) return;
                        PostMessageToWeb(new
                        {
                            action = "autoEnterAttachmentMode",
                            items = attachmentList ?? new List<ComponentApiDto>(),
                            currentModel = modelToQuery,
                            brand = brandToQuery,
                            name = nameToQuery,
                            loading = false
                        });
                    });
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[ComponentMatchOverlayForm] TriggerLoadAttachments 异常: {ex.Message}");
                }
            });
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
                    filterBrand = _filterConfig.GetEffectiveBrands().FirstOrDefault() ?? string.Empty,
                    filterBrands = _filterConfig.GetEffectiveBrands(),
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
                filterBrand = _filterConfig.GetEffectiveBrands().FirstOrDefault() ?? string.Empty,
                filterBrands = _filterConfig.GetEffectiveBrands(),
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
                    // 提取配置中生效的多选品牌列表
                    var effectiveBrands = fc.GetEffectiveBrands();
                    if (isPersonal)
                    {
                        // 从本地 SQLite 个人物料库高速检索 (支持多选品牌)
                        items = PersonalComponentDbService.SearchComponents(
                            null,
                            cp.Name,
                            cp.Current,
                            cp.Pole,
                            cp.TripMode,
                            effectiveBrands,
                            fc.MustContainRules
                        );
                    }
                    else
                    {
                        // 异步调用云端商城 WebAPI 检索 (支持多选品牌)
                        items = await ComponentApiClient.SearchComponentsAsync(
                            null,
                            cp.Name,
                            cp.Current,
                            cp.Pole,
                            cp.TripMode,
                            effectiveBrands,
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
                        // 若处于待切入附件模式，立即触发加载附件；否则拉取常规初始候选物料
                        if (_pendingAttachmentMode)
                        {
                            TriggerLoadAttachments();
                        }
                        else
                        {
                            PushInitialCandidates();
                        }
                        break;

                    // 2. 即时模糊搜索 (全异步非阻塞 + 请求防竞态版本保护 + 支持个人库与云端分流 + 临时必含规则覆盖)
                    case "searchKeyword":
                        string kw = root.TryGetProperty("keyword", out var kwProp) ? kwProp.GetString() ?? "" : "";
                        // 自增请求计数器，防止快速打字异步响应乱序覆盖
                        long currentReqId = Interlocked.Increment(ref _searchReqCounter);
                        _latestSearchReqId = currentReqId;

                        var searchCp = _cellParams;
                        var searchFc = _filterConfig;

                        // 提取动态生效的检索多品牌条件默认值
                        var effectiveBrands = searchFc.GetEffectiveBrands();
                        string effectiveBrand = effectiveBrands.Count > 0 ? effectiveBrands[0] : string.Empty;
                        // 初始元器件名称条件
                        string effectiveName = searchCp.Name;
                        // 初始额定电流条件
                        string effectiveCurrent = searchCp.Current;
                        // 初始极数条件
                        string effectivePole = searchCp.Pole;
                        // 初始脱扣方式条件
                        string effectiveTrip = searchCp.TripMode;

                        // 解析前端动态过滤参数对象 (支持用户在界面上点击 ✕ 移除某项后放宽查询)
                        if (root.TryGetProperty("filters", out var filtersProp) && filtersProp.ValueKind == JsonValueKind.Object)
                        {
                            // 动态覆盖品牌筛选 (前端关闭品牌后传入空字符串，实现不限品牌检索)
                            if (filtersProp.TryGetProperty("brand", out var bProp))
                            {
                                effectiveBrand = bProp.GetString() ?? string.Empty;
                            }

                            // 动态覆盖名称筛选 (前端关闭名称后传入空字符串，放宽名称约束)
                            if (filtersProp.TryGetProperty("name", out var nProp))
                            {
                                effectiveName = nProp.GetString() ?? string.Empty;
                            }

                            // 动态覆盖电流筛选 (前端关闭电流后传入空字符串，放宽电流阶梯限制)
                            if (filtersProp.TryGetProperty("current", out var cProp))
                            {
                                effectiveCurrent = cProp.GetString() ?? string.Empty;
                            }

                            // 动态覆盖极数筛选 (前端关闭极数后传入空字符串，放宽极数约束)
                            if (filtersProp.TryGetProperty("pole", out var pProp))
                            {
                                effectivePole = pProp.GetString() ?? string.Empty;
                            }

                            // 动态覆盖脱扣方式筛选 (前端关闭脱扣方式后传入空字符串)
                            if (filtersProp.TryGetProperty("tripMode", out var tProp))
                            {
                                effectiveTrip = tProp.GetString() ?? string.Empty;
                            }
                        }

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
                                    // 路由到本地 SQLite 个人物料库执行模糊查询 (支持动态放宽多维参数)
                                    searchResults = PersonalComponentDbService.SearchComponents(
                                        kw,
                                        effectiveName,
                                        effectiveCurrent,
                                        effectivePole,
                                        effectiveTrip,
                                        effectiveBrand,
                                        effectiveMustRules
                                    );
                                }
                                else
                                {
                                    // 异步调用云端商城 WebAPI 执行动态放宽参数与必含规则约束检索
                                    searchResults = await ComponentApiClient.SearchComponentsAsync(
                                        kw,
                                        effectiveName,
                                        effectiveCurrent,
                                        effectivePole,
                                        effectiveTrip,
                                        effectiveBrand,
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

                    // 2.2 用户切换“固定”置顶状态 (固定后切行不重搜、失焦不关闭)
                    case "togglePin":
                        if (root.TryGetProperty("pinned", out var pinProp))
                        {
                            _isPinned = pinProp.GetBoolean();
                            if (_isPinned)
                            {
                                this.TopMost = true;
                                this.BringToFront();
                            }
                        }
                        break;

                    // 2.3 响应前端请求启动窗口平滑拖拽移动
                    case "startDrag":
                        BeginDrag();
                        break;

                    // 2.4 响应前端拖拽调整窗口高度并实时记忆
                    case "resizeHeight":
                        if (root.TryGetProperty("height", out var hProp) && hProp.TryGetInt32(out int newH))
                        {
                            // 限制窗口高度在合理区间内 (220 ~ 800 像素)
                            int clampedH = Math.Max(220, Math.Min(800, newH));
                            _customHeight = clampedH;
                            SafeInvoke(() =>
                            {
                                this.Height = clampedH;
                            });
                        }
                        break;

                    // 3. 用户确认选择某一条物料 -> 回填至 Excel (固定模式下不关窗，支持跨行连续点击)
                    case "selectComponent":
                        if (root.TryGetProperty("item", out var itemProp))
                        {
                            var selectedItem = JsonSerializer.Deserialize<ComponentApiDto>(itemProp.GetRawText(), JsonOptions);
                            // 动态获取当前 Excel 的活动单元格 (若用户切行则优先回填至最新的 ActiveCell)
                            dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                            dynamic? curActiveCell = null;
                            try { curActiveCell = app?.ActiveCell; } catch { }
                            dynamic? target = curActiveCell ?? _targetCell;

                            if (selectedItem != null && target != null)
                            {
                                // 1. 立即回填主体元器件至当前目标单元格所在行
                                ExcelServices.FillSelectedComponentToActiveRow(selectedItem, target);

                                int targetRow = 0;
                                try { targetRow = Convert.ToInt32(target.Row); } catch { }

                                // 2. 若处于“固定”模式：保持窗口继续显示，不触发关闭，保留物料列表供连续回填
                                if (_isPinned)
                                {
                                    PostMessageToWeb(new
                                    {
                                        action = "fillSuccess",
                                        row = targetRow,
                                        model = selectedItem.Model ?? string.Empty
                                    });
                                    break;
                                }

                                // 3. 非固定模式：同步更新上下文参数中的主体型号、单价与名称
                                _cellParams.CurrentModel = selectedItem.Model ?? string.Empty;
                                _cellParams.CurrentPrice = selectedItem.Price > 0 ? selectedItem.Price.ToString("F2") : string.Empty;
                                if (!string.IsNullOrWhiteSpace(selectedItem.Name))
                                {
                                    _cellParams.Name = selectedItem.Name;
                                }

                                // 4. 提取用于查询配套附件的品牌、名称与主体型号
                                string hostBrand = !string.IsNullOrWhiteSpace(selectedItem.Brand)
                                    ? selectedItem.Brand
                                    : (_filterConfig.GetEffectiveBrands().FirstOrDefault() ?? _cellParams.Brands.FirstOrDefault() ?? string.Empty);
                                string hostName = selectedItem.Name ?? _cellParams.Name ?? string.Empty;
                                string hostModel = selectedItem.Model ?? string.Empty;

                                // 5. 在后台异步探测并拉取当前选定元器件的配套附件
                                Task.Run(async () =>
                                {
                                    try
                                    {
                                        List<ComponentApiDto> attachmentList;
                                        bool isPersonal = string.Equals(_filterConfig.DataSource, "personal", StringComparison.OrdinalIgnoreCase);
                                        if (isPersonal)
                                        {
                                            // 从本地个人库 SQLite 查询配套附件
                                            attachmentList = PersonalComponentDbService.GetAttachments(hostBrand, hostName, hostModel);
                                        }
                                        else
                                        {
                                            // 从云端商城 WebAPI 异步检索配套附件
                                            attachmentList = await ComponentApiClient.GetAttachmentsAsync(hostBrand, hostName, hostModel).ConfigureAwait(false);
                                        }

                                        // 切回 UI 主线程分流处理
                                        SafeInvoke(() =>
                                        {
                                            // 校验窗体可用状态，若已被用户关闭则不再打扰
                                            if (this.IsDisposed || !this.Visible) return;

                                            if (attachmentList != null && attachmentList.Count > 0)
                                            {
                                                // 分支 A: 存在配套附件 -> 保持悬浮窗显示，通知前端无缝切入附件选配模式
                                                PostMessageToWeb(new
                                                {
                                                    action = "autoEnterAttachmentMode",
                                                    items = attachmentList,
                                                    currentModel = hostModel,
                                                    brand = hostBrand,
                                                    name = hostName
                                                });
                                            }
                                            else
                                            {
                                                // 分支 B: 无配套附件 -> 顺畅隐藏关闭悬浮窗，完成主体回填闭环
                                                this.Hide();
                                            }
                                        });
                                    }
                                    catch (Exception ex)
                                    {
                                        LogHelper.WriteLog($"[ComponentMatchOverlayForm] 自动探查配套附件异常: {ex.Message}");
                                        SafeInvoke(this.Hide);
                                    }
                                });
                                break;
                            }
                        }
                        SafeInvoke(this.Hide);
                        break;

                    // 3.1 用户请求加载当前物料的配套附件列表 (支持个人库与云端分流)
                    case "getAttachments":
                        string brandToQuery = _filterConfig.GetEffectiveBrands().FirstOrDefault() ?? _cellParams.Brands.FirstOrDefault() ?? string.Empty;
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
                        _isPinned = false;
                        SafeInvoke(this.Hide);
                        break;

                    // 5. 点击“云端库/个人库”药丸徽标或右上角设置按钮 -> 弹出“元器件物料匹配与品牌规则设置”窗口 (图2)
                    case "openMatchSettingDialog":
                        LogHelper.WriteLog("[ComponentMatchOverlayForm] 收到 openMatchSettingDialog 请求，正在唤起规则设置窗口");
                        SafeInvoke(() =>
                        {
                            // 启动并弹出基于 WebView2 + Vue 3 的规则设置对话框
                            ExcelServices.ShowComponentMatchDialog();
                        });
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ComponentMatchOverlayForm 消息分发异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 窗体失去焦点时事件处理
        /// </summary>
        private void OnOverlayDeactivate(object? sender, EventArgs e)
        {
            try
            {
                // 若用户已开启“固定”置顶状态，保持前端显示，绝不随失焦隐藏
                if (_isPinned)
                {
                    return;
                }

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
