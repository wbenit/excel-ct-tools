using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelAddInDemo.Forms;
using ExcelAddInDemo.Models;
using ExcelDna.Integration;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 核心业务服务分部类：聚光灯 (Spotlight) 工业级高亮引擎
    /// 基于 Win32 GDI Region 穿透半透明浮窗与原生窗口消息钩子，实现 0 侵入、无损撤销栈的十字行列高亮
    /// </summary>
    public static partial class ExcelServices
    {
        #region Win32 API 声明与常量

        // 窗口矩形结构体
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
            public int Width => Right - Left;
            public int Height => Bottom - Top;
        }

        // 二维坐标点结构体
        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        // 子窗口枚举委托
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        // 枚举子窗口 API
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        // 获取窗口类名 API
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        // 获取窗口客户区矩形 API
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        // 获取窗口绝对屏幕矩形 API
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        // 将窗口客户区坐标换算为屏幕绝对物理坐标 API
        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        // 判断窗口是否处于可见状态 API
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        // 获取当前系统前台获得焦点的窗口句柄 API
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        // 获取窗口所属进程 PID API
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // 获取当前宿主进程自身 PID API
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentProcessId();

        // 设置窗口剪裁区域 (Win32 GDI Region) API
        [DllImport("user32.dll")]
        private static extern int SetWindowRgn(IntPtr hWnd, IntPtr hRgn, bool bRedraw);

        // 控制窗口显隐状态 API
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        // 创建矩形区域 API
        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect);

        // 组合区域 API
        [DllImport("gdi32.dll")]
        private static extern int CombineRgn(IntPtr hrgnDest, IntPtr hrgnSrc1, IntPtr hrgnSrc2, int fnCombineMode);

        // 删除 GDI 区域对象 API
        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        // 区域组合方式常量：并集 (OR)
        private const int RGN_OR = 2;
        // 区域组合方式常量：差集 (DIFF) (用于排除当前单元格)
        private const int RGN_DIFF = 4;
        // 窗口显示隐藏常量：隐藏
        private const int SW_HIDE = 0;
        // 窗口显示隐藏常量：无焦点激活展示
        private const int SW_SHOWNOACTIVATE = 4;

        // 设置窗口属性 64 位 API (用于绑定 Owner 窗口句柄)
        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        // 设置窗口属性 32 位 API
        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        // Win32 GWLP_HWNDPARENT 常量：设置所有者窗口 (Owner Window)
        private const int GWLP_HWNDPARENT = -8;

        #endregion

        #region 静态生命周期字段

        // 聚光灯无边框穿透浮窗实例
        private static SpotlightOverlayForm? _spotlightForm = null;

        // EXCEL7 视口窗口原生消息监听钩子 (仅监听视口自身，严禁挂钩 XLMAIN 顶级主窗口防范死锁)
        private static ExcelWindowHook? _hookExcel7 = null;

        // 当前缓存的 EXCEL7 视口窗口句柄
        private static IntPtr _currentExcel7Hwnd = IntPtr.Zero;

        // 当前聚光灯运行激活状态
        private static bool _isSpotlightEnabled = false;

        // 刷新位置防重入标志位，消除死循环瀑布
        private static bool _isUpdatingPosition = false;

        // 前台焦点守护检测定时器，检测用户是否切到了微信等其他软件 (120ms 极轻量轮询)
        private static System.Windows.Forms.Timer? _foregroundGuardTimer = null;

        // 线程安全互斥锁
        private static readonly object _spotlightSyncRoot = new object();

        #endregion

        #region 公开控制属性与方法

        /// <summary>
        /// 查询当前聚光灯是否处于开启状态
        /// </summary>
        public static bool IsSpotlightEnabled => _isSpotlightEnabled;

        /// <summary>
        /// 切换聚光灯开启/关闭状态 (提供给 Ribbon 按钮与快捷键调用)
        /// </summary>
        public static void ToggleSpotlight()
        {
            // 依据当前激活状态取反流转
            if (_isSpotlightEnabled)
            {
                // 若当前已开启则执行关闭
                DisableSpotlight();
            }
            else
            {
                // 若当前为关闭则执行开启
                EnableSpotlight();
            }
        }

        /// <summary>
        /// 开启聚光灯功能，初始化穿透浮窗与视口消息钩子
        /// </summary>
        public static void EnableSpotlight()
        {
            lock (_spotlightSyncRoot)
            {
                try
                {
                    // 标记当前聚光灯激活状态为 true
                    _isSpotlightEnabled = true;
                    // 同步并持久化配置对象到磁盘
                    SpotlightConfig.Current.IsEnabled = true;
                    SpotlightConfig.Current.SaveToDisk();

                    // 通知 Ribbon 刷新按钮选中状态
                    RibbonController.InvalidateRibbon();

                    // 检查 Excel 宿主与工作簿是否已就绪 (空启动或调试初始化阶段工作簿为0)
                    dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                    // 校验 Application 对象非空
                    if (app == null) return;

                    // 检测当前是否有任何已打开的工作簿
                    bool isReady = false;
                    try
                    {
                        // 读取 Workbooks 集合总数及活动窗口
                        isReady = (app.Workbooks != null && app.Workbooks.Count > 0 && app.ActiveWindow != null);
                    }
                    catch
                    {
                        // 启动初期 COM 消息泵未就绪时可能抛异常，静默拦截
                        isReady = false;
                    }

                    // 启动前台焦点守护定时器，时刻检测切出切回状态
                    StartForegroundGuard();

                    // 若当前暂无打开的工作簿 (如纯空启动调试)，保持激活状态退出，后续打开表格时自愈
                    if (!isReady)
                    {
                        // 记录未就绪等待自愈日志
                        LogHelper.WriteLog("[Spotlight] 暂无活动工作簿，保持激活状态待表格打开后自愈");
                        return;
                    }

                    // 1. 初始化穿透浮窗并确保句柄建立
                    EnsureSpotlightFormCreated();

                    // 2. 绑定当前活动 EXCEL7 视口并挂载原生消息监听
                    AttachExcelHooks();

                    // 3. 立即触发一次坐标与区域高亮刷新
                    UpdateSpotlightPosition(null);

                    // 记录开启成功日志
                    LogHelper.WriteLog("[Spotlight] 聚光灯功能已开启");
                }
                catch (Exception ex)
                {
                    // 即使首次刷新产生临时状态波动，保持开启状态以允许在后续单元格切换中自愈
                    LogHelper.WriteLog($"[Spotlight] 开启聚光灯过程异常 (非致命): {ex.Message}");
                    // 通知 Ribbon 依然展示最新状态
                    RibbonController.InvalidateRibbon();
                }
            }
        }

        /// <summary>
        /// 关闭聚光灯功能，安全隐匿浮窗并释放消息钩子
        /// </summary>
        public static void DisableSpotlight()
        {
            lock (_spotlightSyncRoot)
            {
                try
                {
                    // 标记关闭状态
                    _isSpotlightEnabled = false;
                    // 同步并持久化配置对象为禁用
                    SpotlightConfig.Current.IsEnabled = false;
                    SpotlightConfig.Current.SaveToDisk();

                    // 停止前台焦点守护定时器
                    StopForegroundGuard();

                    // 1. 解除原生消息钩子
                    DetachExcelHooks();

                    // 2. 隐匿并安全销毁穿透浮窗
                    if (_spotlightForm != null)
                    {
                        // 隐藏原生窗口
                        if (_spotlightForm.IsHandleCreated)
                        {
                            ShowWindow(_spotlightForm.Handle, SW_HIDE);
                        }
                        // 销毁窗体资源
                        _spotlightForm.Dispose();
                        _spotlightForm = null;
                    }

                    // 重置缓存视口句柄
                    _currentExcel7Hwnd = IntPtr.Zero;

                    // 通知 Ribbon 刷新按钮复选状态
                    RibbonController.InvalidateRibbon();

                    // 记录停用日志
                    LogHelper.WriteLog("[Spotlight] 聚光灯功能已关闭");
                }
                catch (Exception ex)
                {
                    // 记录关闭过程异常
                    LogHelper.WriteLog($"[Spotlight] 关闭聚光灯异常: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 响应活动单元格改变或视口变动，重新计算并刷新聚光灯高亮十字区域
        /// </summary>
        /// <param name="target">发生改变的单元格区域 (为 null 时自动提取 ActiveCell)</param>
        public static void UpdateSpotlightPosition(dynamic? target = null)
        {
            // 若未开启则直接退出
            if (!_isSpotlightEnabled) return;

            // 若当前系统前台获得焦点的窗口不属于当前 Excel 进程，保持隐匿并退出
            if (!IsExcelForeground())
            {
                // 隐匿已显示的浮窗
                if (_spotlightForm != null && _spotlightForm.Visible)
                {
                    _spotlightForm.Visible = false;
                    ShowWindow(_spotlightForm.Handle, SW_HIDE);
                }
                return;
            }

            // 防重入互斥防护：若当前正在执行刷新，忽略并发调用避免死锁风暴
            if (_isUpdatingPosition) return;
            _isUpdatingPosition = true;

            try
            {
                // 获取 Excel COM Application 实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return;

                // 校验是否存在已打开的工作簿与活动窗口
                dynamic? activeWin = null;
                try
                {
                    // 检查 Workbooks 集合有效且数量大于 0
                    if (app.Workbooks == null || app.Workbooks.Count == 0) return;
                    // 读取活动窗口实例
                    activeWin = app.ActiveWindow;
                }
                catch
                {
                    // 捕获 COM 未就绪异常并安全退出
                    return;
                }
                // 校验活动窗口实例非空
                if (activeWin == null) return;

                // 确保浮窗已就绪并完成 HWND 句柄分配
                EnsureSpotlightFormCreated();
                if (_spotlightForm == null || !_spotlightForm.IsHandleCreated) return;

                // 检查并寻找当前活动可见 EXCEL7 视口句柄
                IntPtr excel7Hwnd = FindExcel7Hwnd();
                if (excel7Hwnd == IntPtr.Zero)
                {
                    // 若当前无法获取有效视口，暂时隐匿浮窗
                    ShowWindow(_spotlightForm.Handle, SW_HIDE);
                    return;
                }

                // 若视口句柄发生改变 (例如用户切换了活动工作簿)，重新挂载消息钩子
                if (_currentExcel7Hwnd != excel7Hwnd)
                {
                    AttachExcelHooks();
                }

                // 获取 EXCEL7 视口客户区物理大小
                if (!GetClientRect(excel7Hwnd, out RECT clientRect) || clientRect.Width <= 10 || clientRect.Height <= 10)
                {
                    // 视口过小或处于最小化时隐匿浮窗
                    ShowWindow(_spotlightForm.Handle, SW_HIDE);
                    return;
                }

                // 获取 EXCEL7 在屏幕上的绝对物理像素原点
                POINT screenOrigin = new POINT { X = 0, Y = 0 };
                ClientToScreen(excel7Hwnd, ref screenOrigin);

                // 调整穿透浮窗位置与几何尺寸严丝合缝贴合 EXCEL7 视口
                if (_spotlightForm.Left != screenOrigin.X ||
                    _spotlightForm.Top != screenOrigin.Y ||
                    _spotlightForm.Width != clientRect.Width ||
                    _spotlightForm.Height != clientRect.Height)
                {
                    _spotlightForm.SetBounds(screenOrigin.X, screenOrigin.Y, clientRect.Width, clientRect.Height);
                }

                // 提取目标活动选区 (优先取传入 target，若为空则从当前 app.Selection 提取多选选区，最后回退至 ActiveCell)
                dynamic? targetRange = target;
                if (targetRange == null)
                {
                    try
                    {
                        // 从当前应用程序提取用户最新的选择对象
                        dynamic sel = app.Selection;
                        // 校验 sel 是否为具备 Areas 属性的有效 Range 选区
                        if (sel != null && sel.Areas != null && sel.Areas.Count > 0)
                        {
                            targetRange = sel;
                        }
                    }
                    catch { }
                }

                // 兜底回退至 ActiveCell (单格)
                if (targetRange == null)
                {
                    try
                    {
                        targetRange = app.ActiveCell;
                    }
                    catch { }
                }

                // 若无法获取任何有效单元格或选区则隐藏浮窗
                if (targetRange == null)
                {
                    ShowWindow(_spotlightForm.Handle, SW_HIDE);
                    return;
                }

                // 获取聚光灯全局配置
                var cfg = SpotlightConfig.Current;
                // 创建初始空的 Windows GDI 组合区域
                IntPtr hCombined = CreateRectRgn(0, 0, 0, 0);
                // 标记当前视口内是否生成了任何有效的高亮区域
                bool hasValidRegion = false;

                // 尝试获取选区中的 Areas 集合 (兼容多区域 Ctrl 组合选区与连续多行多列)
                dynamic? areas = null;
                int areaCount = 1;
                try
                {
                    // 读取 Areas 集合
                    areas = targetRange.Areas;
                    // 读取区域个数
                    areaCount = Convert.ToInt32(areas.Count);
                }
                catch
                {
                    // 若不支持 Areas 则视为单区域
                    areaCount = 1;
                }

                // 限制单次最大遍历区域数，防止超大数量多选引发遍历性能卡顿
                int safeAreaLimit = Math.Min(areaCount, 20); // --硬编码: 多选区遍历安全上限20个--

                // 遍历多选区中的每一个子区域
                for (int i = 1; i <= safeAreaLimit; i++)
                {
                    dynamic? currentArea = null;
                    try
                    {
                        // 提取第 i 个子选区
                        currentArea = (areas != null && areaCount > 1) ? areas.Item(i) : targetRange;
                    }
                    catch
                    {
                        currentArea = targetRange;
                    }

                    if (currentArea == null) continue;

                    // 计算该区域的绝对屏幕物理矩形 (双轨容错)
                    Rectangle cellScreenRect = CalculateCellScreenRect(activeWin, currentArea);
                    if (cellScreenRect.Width <= 0 || cellScreenRect.Height <= 0) continue;

                    // 转换为相对于 EXCEL7 视口 (即浮窗自身客户区) 的局部像素坐标
                    int relLeft = cellScreenRect.Left - screenOrigin.X;
                    int relTop = cellScreenRect.Top - screenOrigin.Y;
                    int relRight = cellScreenRect.Right - screenOrigin.X;
                    int relBottom = cellScreenRect.Bottom - screenOrigin.Y;

                    // 边界安全裁剪，防止滚出屏幕产生负数或超出视口宽高
                    int safeLeft = Math.Max(0, Math.Min(clientRect.Width, relLeft));
                    int safeRight = Math.Max(0, Math.Min(clientRect.Width, relRight));
                    int safeTop = Math.Max(0, Math.Min(clientRect.Height, relTop));
                    int safeBottom = Math.Max(0, Math.Min(clientRect.Height, relBottom));

                    // 依据用户选定模式构造当前 Area 的 GDI 区域
                    switch (cfg.Mode)
                    {
                        // 仅行高亮模式
                        case SpotlightMode.RowOnly:
                            if (safeBottom > safeTop)
                            {
                                // 构造横贯整个视口宽度的行高亮带
                                IntPtr hRow = CreateRectRgn(0, safeTop, clientRect.Width, safeBottom);
                                // 并入总高亮区域
                                CombineRgn(hCombined, hCombined, hRow, RGN_OR);
                                // 释放临时行矩形
                                DeleteObject(hRow);
                                hasValidRegion = true;
                            }
                            break;

                        // 仅列高亮模式
                        case SpotlightMode.ColumnOnly:
                            if (safeRight > safeLeft)
                            {
                                // 构造纵贯整个视口高度的列高亮带
                                IntPtr hCol = CreateRectRgn(safeLeft, 0, safeRight, clientRect.Height);
                                // 并入总高亮区域
                                CombineRgn(hCombined, hCombined, hCol, RGN_OR);
                                // 释放临时列矩形
                                DeleteObject(hCol);
                                hasValidRegion = true;
                            }
                            break;

                        // 十字交叉高亮模式 (默认)
                        case SpotlightMode.Crosshair:
                        default:
                            if (safeBottom > safeTop)
                            {
                                // 构造整行区域
                                IntPtr hRow = CreateRectRgn(0, safeTop, clientRect.Width, safeBottom);
                                CombineRgn(hCombined, hCombined, hRow, RGN_OR);
                                DeleteObject(hRow);
                                hasValidRegion = true;
                            }
                            if (safeRight > safeLeft)
                            {
                                // 构造整列区域
                                IntPtr hCol = CreateRectRgn(safeLeft, 0, safeRight, clientRect.Height);
                                CombineRgn(hCombined, hCombined, hCol, RGN_OR);
                                DeleteObject(hCol);
                                hasValidRegion = true;
                            }
                            break;
                    }

                    // 若配置了排除当前活动单元格 (镂空活动选区中心避免遮挡输入光标)
                    if (cfg.ExcludeActiveCell && safeRight > safeLeft && safeBottom > safeTop)
                    {
                        // 创建该选区自身所占矩形区域
                        IntPtr hCell = CreateRectRgn(safeLeft, safeTop, safeRight, safeBottom);
                        // 从组合高亮区域中扣除单元格矩形 (RGN_DIFF)
                        CombineRgn(hCombined, hCombined, hCell, RGN_DIFF);
                        // 释放临时单元格矩形
                        DeleteObject(hCell);
                    }
                }

                // 若选区在当前视口内完全不可见，隐匿浮窗并释放 GDI 对象
                if (!hasValidRegion)
                {
                    // 释放组合 GDI 区域
                    DeleteObject(hCombined);
                    // 同步托管隐藏
                    if (_spotlightForm.Visible)
                    {
                        _spotlightForm.Visible = false;
                    }
                    // Win32 底层隐藏
                    ShowWindow(_spotlightForm.Handle, SW_HIDE);
                    return;
                }

                // 设置裁剪几何区域给无边框浮窗 (Windows 操作系统接管 hCombined 生命周期)
                SetWindowRgn(_spotlightForm.Handle, hCombined, true);

                // 确保浮窗在托管层面处于显示态
                if (!_spotlightForm.Visible)
                {
                    _spotlightForm.Visible = true;
                }

                // 强制分层窗体刷新重绘，消除桌面合成器残留
                _spotlightForm.Invalidate();
                _spotlightForm.Update();

                // 显式无激活显示浮窗
                ShowWindow(_spotlightForm.Handle, SW_SHOWNOACTIVATE);
            }
            catch (Exception ex)
            {
                // 记录高亮渲染中的瞬时异常
                LogHelper.WriteLog($"[Spotlight] 刷新高亮位置异常: {ex.Message}");
            }
            finally
            {
                // 恢复防重入标志位
                _isUpdatingPosition = false;
            }
        }

        /// <summary>
        /// 采用双轨方案高精度换算单元格或选区在屏幕上的绝对物理矩形
        /// </summary>
        /// <param name="activeWin">Excel 活动窗口</param>
        /// <param name="cell">活动单元格或多选选区对象</param>
        /// <returns>单元格屏幕物理矩形</returns>
        private static Rectangle CalculateCellScreenRect(dynamic activeWin, dynamic cell)
        {
            try
            {
                // 读取单元格或选区几何尺寸 (Point 磅为单位)
                double cellLeftPt = Convert.ToDouble(cell.Left);
                double cellTopPt = Convert.ToDouble(cell.Top);
                // 对超大跨度选区进行安全上限截断，防范整行整列导致 PointsToScreenPixels 整数溢出
                double cellWidthPt = Math.Min(Convert.ToDouble(cell.Width), 8000); // --硬编码: 磅值安全上限8000磅--
                double cellHeightPt = Math.Min(Convert.ToDouble(cell.Height), 8000); // --硬编码: 磅值安全上限8000磅--

                // 轨 1: 优先尝试通过 ActivePane 进行视口换算
                try
                {
                    dynamic pane = activeWin.ActivePane;
                    if (pane != null)
                    {
                        int pLeft = pane.PointsToScreenPixelsX((int)cellLeftPt);
                        int pTop = pane.PointsToScreenPixelsY((int)cellTopPt);
                        int pRight = pane.PointsToScreenPixelsX((int)(cellLeftPt + cellWidthPt));
                        int pBottom = pane.PointsToScreenPixelsY((int)(cellTopPt + cellHeightPt));

                        if (pRight > pLeft && pBottom > pTop)
                        {
                            return new Rectangle(pLeft, pTop, pRight - pLeft, pBottom - pTop);
                        }
                    }
                }
                catch { }

                // 轨 2: 回退基于 Window.PointsToScreenPixels 换算
                try
                {
                    int wLeft = activeWin.PointsToScreenPixelsX((int)cellLeftPt);
                    int wTop = activeWin.PointsToScreenPixelsY((int)cellTopPt);
                    int wRight = activeWin.PointsToScreenPixelsX((int)(cellLeftPt + cellWidthPt));
                    int wBottom = activeWin.PointsToScreenPixelsY((int)(cellTopPt + cellHeightPt));

                    if (wRight > wLeft && wBottom > wTop)
                    {
                        return new Rectangle(wLeft, wTop, wRight - wLeft, wBottom - wTop);
                    }
                }
                catch { }

                // 轨 3: 极端异常兜底
                return Rectangle.Empty;
            }
            catch
            {
                return Rectangle.Empty;
            }
        }

        #endregion

        #region 内部辅助逻辑与钩子挂载

        /// <summary>
        /// 确保穿透浮窗已实例化并完成 HWND 原生句柄创建
        /// </summary>
        private static void EnsureSpotlightFormCreated()
        {
            // 检查浮窗实例是否未创建或已处于释放态
            if (_spotlightForm == null || _spotlightForm.IsDisposed)
            {
                // 实例化全新无边框半透明穿透浮窗
                _spotlightForm = new SpotlightOverlayForm();

                // 获取 Excel 宿主主窗口句柄
                IntPtr mainHwnd = ExcelDnaSafeAccessor.GetWindowHandle();
                if (mainHwnd != IntPtr.Zero)
                {
                    // 设置 Owner 依附关系，确保浮窗只从属于 Excel，切出时跟随 Excel 沉入后台
                    SetWindowOwner(_spotlightForm.Handle, mainHwnd);
                }

                // 显式调用 Show 确保窗体 Visible = true 并加入系统分层渲染
                // 因 SpotlightOverlayForm 已重写 ShowWithoutActivation => true，绝不抢占焦点
                _spotlightForm.Show();
            }
        }

        /// <summary>
        /// 全方位查找当前活动 Excel 工作表网格对应的 EXCEL7 视口窗口句柄 (兼容 SDI 与多工作簿模式)
        /// </summary>
        /// <returns>EXCEL7 原生 HWND 句柄，未找到返回 IntPtr.Zero</returns>
        private static IntPtr FindExcel7Hwnd()
        {
            try
            {
                // 获取 Excel COM Application 宿主实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return IntPtr.Zero;

                // 优先从活动工作簿窗口中提取 HWND
                IntPtr activeWinHwnd = IntPtr.Zero;
                try
                {
                    // 安全读取 ActiveWindow.Hwnd 并转换为 long 整数
                    object rawHwnd = app.ActiveWindow.Hwnd;
                    long lVal = Convert.ToInt64(rawHwnd);
                    // 构造 64位/32位通用原生窗口句柄
                    activeWinHwnd = new IntPtr(lVal);
                }
                catch (Exception ex)
                {
                    // 记录提取活动窗口句柄时的诊断信息
                    LogHelper.WriteLog($"[Spotlight] 获取活动窗口 Hwnd 异常: {ex.Message}");
                }

                IntPtr candidateHwnd = IntPtr.Zero;
                int maxArea = 0;

                // 1. 若活动窗口有效，优先枚举其子窗口
                if (activeWinHwnd != IntPtr.Zero)
                {
                    // 检查活动窗口自身是否直接为 EXCEL7
                    StringBuilder sbSelf = new StringBuilder(256);
                    GetClassName(activeWinHwnd, sbSelf, sbSelf.Capacity);
                    if (sbSelf.ToString().Equals("EXCEL7", StringComparison.OrdinalIgnoreCase) && IsWindowVisible(activeWinHwnd))
                    {
                        return activeWinHwnd;
                    }

                    // 枚举活动窗口的所有可见子层级
                    EnumChildWindows(activeWinHwnd, (childHwnd, lParam) =>
                    {
                        // 过滤不可见的隐藏窗口
                        if (!IsWindowVisible(childHwnd)) return true;

                        StringBuilder sb = new StringBuilder(256);
                        GetClassName(childHwnd, sb, sb.Capacity);
                        if (sb.ToString().Equals("EXCEL7", StringComparison.OrdinalIgnoreCase))
                        {
                            // 计算面积选取最大的可见视口
                            if (GetClientRect(childHwnd, out RECT rc) && rc.Width * rc.Height > maxArea)
                            {
                                maxArea = rc.Width * rc.Height;
                                candidateHwnd = childHwnd;
                            }
                        }
                        return true;
                    }, IntPtr.Zero);

                    if (candidateHwnd != IntPtr.Zero) return candidateHwnd;
                }

                // 2. 兜底方案：从 Excel 顶级宿主主窗口 (XLMAIN) 深度枚举查找面积最大的可见 EXCEL7
                IntPtr mainHwnd = ExcelDnaSafeAccessor.GetWindowHandle();
                if (mainHwnd != IntPtr.Zero)
                {
                    EnumChildWindows(mainHwnd, (childHwnd, lParam) =>
                    {
                        if (!IsWindowVisible(childHwnd)) return true;

                        StringBuilder sb = new StringBuilder(256);
                        GetClassName(childHwnd, sb, sb.Capacity);
                        if (sb.ToString().Equals("EXCEL7", StringComparison.OrdinalIgnoreCase))
                        {
                            if (GetClientRect(childHwnd, out RECT rc) && rc.Width * rc.Height > maxArea)
                            {
                                maxArea = rc.Width * rc.Height;
                                candidateHwnd = childHwnd;
                            }
                        }
                        return true;
                    }, IntPtr.Zero);
                }

                return candidateHwnd;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[Spotlight] 探测 EXCEL7 视口异常: {ex.Message}");
                return IntPtr.Zero;
            }
        }

        /// <summary>
        /// 挂载 EXCEL7 视口原生窗口消息监听钩子 (仅对视口监听，绝不触碰 XLMAIN 顶级窗口)
        /// </summary>
        private static void AttachExcelHooks()
        {
            try
            {
                // 先安全解绑旧的监听钩子
                DetachExcelHooks();

                // 寻找最新的活动 EXCEL7 视口句柄
                _currentExcel7Hwnd = FindExcel7Hwnd();
                if (_currentExcel7Hwnd != IntPtr.Zero)
                {
                    // 建立针对当前 EXCEL7 的原生消息监听
                    _hookExcel7 = new ExcelWindowHook(_currentExcel7Hwnd);
                    // 挂载尺寸或位置变动回调 (由 Hook 内部防抖定时器异步调用)
                    _hookExcel7.OnBoundsChanged = () => UpdateSpotlightPosition(null);
                    // 挂载平滑滚屏防抖重绘回调 (由 Hook 内部防抖定时器异步调用)
                    _hookExcel7.OnScrolled = () => UpdateSpotlightPosition(null);
                    // 挂载激活/失去焦点回调：当视口失焦且当前前台不是 Excel 时立即隐匿
                    _hookExcel7.OnActivationChanged = (isActive) =>
                    {
                        // 校验若失焦且当前前台非 Excel
                        if (!isActive && !IsExcelForeground())
                        {
                            // 立即隐藏浮窗
                            if (_spotlightForm != null && _spotlightForm.Visible)
                            {
                                _spotlightForm.Visible = false;
                                ShowWindow(_spotlightForm.Handle, SW_HIDE);
                            }
                        }
                    };
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[Spotlight] 挂载窗口消息钩子异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 解除并释放挂载的视口窗口消息钩子
        /// </summary>
        private static void DetachExcelHooks()
        {
            // 释放 EXCEL7 钩子实例
            if (_hookExcel7 != null)
            {
                // 销毁钩子对象
                _hookExcel7.Dispose();
                _hookExcel7 = null;
            }
        }

        /// <summary>
        /// 将浮窗设置为特定宿主窗口的所属窗口 (Owner)，确保 Z 序永远依附于宿主，绝不遮挡其他软件
        /// </summary>
        /// <param name="childHwnd">浮窗原生句柄</param>
        /// <param name="ownerHwnd">Excel 主窗口原生句柄</param>
        private static void SetWindowOwner(IntPtr childHwnd, IntPtr ownerHwnd)
        {
            try
            {
                // 校验传入句柄有效性
                if (childHwnd == IntPtr.Zero || ownerHwnd == IntPtr.Zero) return;
                // 依据系统位数分流设置 GWLP_HWNDPARENT
                if (IntPtr.Size == 8)
                {
                    // 64位系统调用 SetWindowLongPtr64
                    SetWindowLongPtr64(childHwnd, GWLP_HWNDPARENT, ownerHwnd);
                }
                else
                {
                    // 32位系统调用 SetWindowLong32
                    SetWindowLong32(childHwnd, GWLP_HWNDPARENT, ownerHwnd.ToInt32());
                }
            }
            catch { }
        }

        /// <summary>
        /// 极速判断当前系统前台获得焦点的窗口是否属于 Excel 宿主进程
        /// </summary>
        /// <returns>若前台窗口属于当前 Excel 进程返回 true，否则返回 false</returns>
        private static bool IsExcelForeground()
        {
            try
            {
                // 读取当前系统前台激活的窗口 HWND
                IntPtr fgHwnd = GetForegroundWindow();
                // 若前台窗口为空则返回 false
                if (fgHwnd == IntPtr.Zero) return false;

                // 提取前台窗口归属进程 PID
                GetWindowThreadProcessId(fgHwnd, out uint fgPid);
                // 对比当前 Excel 自身 PID
                return fgPid == GetCurrentProcessId();
            }
            catch
            {
                // 异常时默认放行
                return true;
            }
        }

        /// <summary>
        /// 启动前台焦点守护检测定时器
        /// </summary>
        private static void StartForegroundGuard()
        {
            // 若定时器尚未初始化
            if (_foregroundGuardTimer == null)
            {
                // 实例化 WinForms 定时器
                _foregroundGuardTimer = new System.Windows.Forms.Timer
                {
                    // 设定 120ms 极轻量轮询间隔
                    Interval = 120 // --硬编码: 前台焦点轮询间隔 120毫秒--
                };
                // 绑定 Tick 事件委托
                _foregroundGuardTimer.Tick += ForegroundGuardTimer_Tick;
            }
            // 启动定时器
            _foregroundGuardTimer.Start();
        }

        /// <summary>
        /// 停止并清理前台焦点守护定时器
        /// </summary>
        private static void StopForegroundGuard()
        {
            // 若定时器实例存在
            if (_foregroundGuardTimer != null)
            {
                // 停止定时器运行
                _foregroundGuardTimer.Stop();
                // 解绑 Tick 事件委托
                _foregroundGuardTimer.Tick -= ForegroundGuardTimer_Tick;
                // 销毁定时器资源
                _foregroundGuardTimer.Dispose();
                _foregroundGuardTimer = null;
            }
        }

        /// <summary>
        /// 前台焦点守护定时器回调：当用户切到其他软件时立即隐藏高亮，切回 Excel 时自动唤醒
        /// </summary>
        private static void ForegroundGuardTimer_Tick(object? sender, EventArgs e)
        {
            // 若未开启聚光灯或浮窗未创建则直接退出
            if (!_isSpotlightEnabled || _spotlightForm == null || _spotlightForm.IsDisposed) return;

            // 检测当前前台焦点窗口是否属于 Excel 宿主进程
            bool isExcelActive = IsExcelForeground();

            // 若用户切到了其他外部软件 (如微信、VS Code、浏览器)
            if (!isExcelActive)
            {
                // 若聚光灯浮窗正处于显示状态，立即隐匿！
                if (_spotlightForm.Visible)
                {
                    // 标记隐藏状态
                    _spotlightForm.Visible = false;
                    // 调用底层 API 隐藏原生窗口
                    ShowWindow(_spotlightForm.Handle, SW_HIDE);
                }
            }
            else
            {
                // 当前正处于 Excel 前台：若此前被隐藏了，立即自动唤醒重绘！
                if (!_spotlightForm.Visible)
                {
                    // 触发位置计算并自愈恢复高亮
                    UpdateSpotlightPosition(null);
                }
            }
        }

        #endregion

        #region Excel-DNA 快捷键宏命令注册

        /// <summary>
        /// Excel-DNA 快捷键宏命令，绑定 Ctrl + Alt + L 快速切换聚光灯
        /// </summary>
        [ExcelCommand(ShortCut = "^%L")]
        public static void ToggleSpotlightCommand()
        {
            // 调度聚光灯切换逻辑
            ToggleSpotlight();
        }

        #endregion
    }
}
