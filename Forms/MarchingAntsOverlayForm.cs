using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 橙色流动细虚线穿透浮窗 (Marching Ants Overlay Form)
    /// 核心特性：
    /// 1. 绝对纯净透明：内部零填充、零遮挡，Excel 单元格文字 100% 清晰可见
    /// 2. 纯正活力鲜橙色：1px 极细流动虚线，关闭抗锯齿杜绝杂色伪影
    /// 3. 稳健生命周期：随用随建，用完即毁，彻底根除取消后再复制无效果缺陷
    /// 4. 穿透无激活：鼠标操作与滚动 100% 穿透，不抢占 Excel 输入光标
    /// </summary>
    public class MarchingAntsOverlayForm : Form
    {
        #region Win32 常量与 API 声明

        // 工具栏窗口样式 (不在任务栏与 Alt+Tab 切换列表中出现)
        private const int WS_EX_TOOLWINDOW = 0x00000080;

        // 鼠标点击时不激活窗口，保持 Excel 焦点
        private const int WS_EX_NOACTIVATE = 0x08000000;

        // 鼠标事件完全穿透到底层 Excel
        private const int WS_EX_TRANSPARENT = 0x00000020;

        // 分层窗口支持透明色键
        private const int WS_EX_LAYERED = 0x00080000;

        #endregion

        #region 外观与动画配置常数 (--硬编码: 样式与动效参数--)

        // 窗体背景设为品红色，作为穿透透明色键 --硬编码: 纯品红透明色键--
        private static readonly Color TransparentKeyColor = Color.Magenta;

        // 纯正活力鲜橙色边框主色调 --硬编码: 鲜艳橙色 #FF7800--
        private static readonly Color VividOrangeBorderColor = Color.FromArgb(255, 120, 0);

        // 虚线边框线宽：1 像素极细线宽 --硬编码: 细虚线线宽 1.0px--
        private const float BorderLineWidth = 1.0f;

        // 精细虚线分段样式：3 像素实线，3 像素空白 --硬编码: 蚂蚁线分段比例 [3, 3]--
        private static readonly float[] DashPatternArray = new float[] { 3f, 3f };

        #endregion

        // 当前虚线流动偏移量 (随时间推移递增)
        private float _dashOffset = 0f;

        /// <summary>
        /// 构造函数: 初始化穿透无边框浮窗外观与双缓冲配置
        /// </summary>
        public MarchingAntsOverlayForm()
        {
            // 设置无边框模式
            this.FormBorderStyle = FormBorderStyle.None;
            // 设置手动绝对屏幕坐标定位
            this.StartPosition = FormStartPosition.Manual;
            // 不在任务栏展示图标
            this.ShowInTaskbar = false;
            // 不使用系统级 TopMost，避免遮挡外部应用
            this.TopMost = false;
            // 启用双缓冲，彻底消除虚线流动时的重绘闪烁
            this.DoubleBuffered = true;
            // 将背景底色与透明色键绑定为纯品红色
            this.BackColor = TransparentKeyColor;
            this.TransparencyKey = TransparentKeyColor;
        }

        /// <summary>
        /// 重写展示无焦点激活属性，确保显示时不夺取 Excel 主网格编辑焦点
        /// </summary>
        protected override bool ShowWithoutActivation => true;

        /// <summary>
        /// 重写底层窗口创建参数，注入穿透与无激活扩展样式
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                // 获取基础创建参数
                CreateParams cp = base.CreateParams;
                // 组合穿透、非激活、工具栏及分层属性
                cp.ExStyle |= (WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TRANSPARENT | WS_EX_LAYERED);
                return cp;
            }
        }

        /// <summary>
        /// 更新当前的虚线流动偏移量并触发界面局部快速重绘
        /// </summary>
        /// <param name="offset">新的流动偏移步长</param>
        public void UpdateDashOffset(float offset)
        {
            // 更新虚线偏移量
            _dashOffset = offset;
            // 触发客户区重绘
            this.Invalidate();
        }

        /// <summary>
        /// 自定义绘制流动橙色细虚线矩形外边框
        /// 严格遵循：内部 100% 绝对不填充任何颜色，保持纯透明穿透；关闭抗锯齿杜绝品红混色伪影
        /// </summary>
        /// <param name="e">绘图事件参数</param>
        protected override void OnPaint(PaintEventArgs e)
        {
            // 调用基类重绘
            base.OnPaint(e);
            // 宽度或高度不足时放弃绘制避免异常
            if (this.Width <= 2 || this.Height <= 2) return;

            Graphics g = e.Graphics;
            // 关闭抗锯齿平滑渲染，确保虚线像素纯正橙色，杜绝与品红背景混合产生杂色毛边
            g.SmoothingMode = SmoothingMode.None;

            // 绘制最外圈流动橙色 1px 细虚线 (Marching Ants 动态流动虚线)
            using (Pen pen = new Pen(VividOrangeBorderColor, BorderLineWidth))
            {
                // 设置精细虚线分段样式
                pen.DashPattern = DashPatternArray;
                // 关键动效：应用动态递增的虚线偏移量，产生丝滑顺时针流动的视觉效果
                pen.DashOffset = _dashOffset;

                // 仅在外圈绘制 1 像素细虚线框，内部绝对不涂抹任何像素，100% 透出 Excel 表格文字
                g.DrawRectangle(pen, 0, 0, this.Width - 1, this.Height - 1);
            }
        }
    }

    /// <summary>
    /// 全局静态橙色流动细虚线管理器 (MarchingAntsManager)
    /// 集中管理元器件复制/剪切动效的创建、销毁、画布跟踪与按ESC自动消退
    /// </summary>
    public static class MarchingAntsManager
    {
        #region Win32 API 辅助声明

        // 异步按键状态检测 API (用于捕获用户按下 ESC 键)
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        // 获取前台激活窗口句柄 API
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        // 获取窗口所属进程 PID API
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // 获取当前进程 PID
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentProcessId();

        // ESC 虚拟键码常量
        private const int VK_ESCAPE = 0x1B;

        #endregion

        // 单例穿透浮窗实例
        private static MarchingAntsOverlayForm? _overlayForm = null;

        // 动画定时器 (40ms 刷新率，约 25 帧/秒丝滑流动) --硬编码: 动画帧间隔 40ms--
        private static System.Windows.Forms.Timer? _animTimer = null;

        // 标记当前动效是否处于激活运行状态
        private static bool _isActive = false;

        // 虚线流动偏移累加量
        private static float _dashOffset = 0f;

        // 目标工作表名称
        private static string _targetSheetName = string.Empty;

        // 目标起始物理行号
        private static int _startRow = 0;

        // 目标终止物理行号
        private static int _endRow = 0;

        // 目标展示列宽 (覆盖前台可见业务列，如 A~T 列)
        private static int _endCol = 20;

        // 低频位置校准计数器 (避免每帧高频调用 COM 发生锁冲突)
        private static int _tickCounter = 0;

        // 互斥对象锁
        private static readonly object _syncRoot = new object();

        /// <summary>
        /// 查询当前橙色流动虚线动效是否正在激活运行
        /// </summary>
        public static bool IsRunning => _isActive;

        /// <summary>
        /// 在指定的元器件行区域启动并呈现纯正橙色流动细虚线动效
        /// 每次调用均彻底销毁旧实例并重建全新浮窗，确保任何取消操作后再次复制百分之百生效
        /// </summary>
        /// <param name="sheet">目标工作表句柄</param>
        /// <param name="startRow">起始物理行号</param>
        /// <param name="endRow">终止物理行号</param>
        /// <param name="endCol">结束物理列号 (默认 20 列，即 A~T 列)</param>
        public static void Show(dynamic sheet, int startRow, int endRow, int endCol = 20)
        {
            // 加锁保障线程安全
            lock (_syncRoot)
            {
                try
                {
                    // 1. 彻底关闭并释放旧实例，根除隐藏后再显示导致的 DWM 分层渲染失效缺陷
                    HideInternal();

                    // 2. 缓存当前选区关键参数
                    _targetSheetName = Convert.ToString(sheet.Name) ?? "";
                    _startRow = startRow;
                    _endRow = endRow;
                    _endCol = Math.Max(1, endCol);
                    _isActive = true;
                    _dashOffset = 0f;
                    _tickCounter = 0;

                    // 3. 获取 Excel Application 与活动窗口
                    dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                    if (app == null) return;

                    // 4. 定位目标区域并高精度计算屏幕物理像素坐标
                    string startColLetter = GetColumnLetter(1);
                    string endColLetter = GetColumnLetter(_endCol);
                    string rangeAddress = $"{startColLetter}{_startRow}:{endColLetter}{_endRow}";
                    dynamic targetRange = sheet.Range[rangeAddress];

                    Rectangle screenRect = CalculateRangeScreenRect(app.ActiveWindow, targetRange);
                    // 若几何尺寸不合法则暂不上屏
                    if (screenRect.Width <= 4 || screenRect.Height <= 4)
                    {
                        return;
                    }

                    // 外扩 1 像素包裹单元格边框外沿
                    screenRect.Inflate(1, 1);

                    // 5. 实例化全新穿透浮窗并定位
                    _overlayForm = new MarchingAntsOverlayForm();
                    _overlayForm.SetBounds(screenRect.X, screenRect.Y, screenRect.Width, screenRect.Height);

                    // 6. 绑定 Excel 主窗口为 Owner，使其自然随 Excel 最小化/还原，不浮在其他应用之上
                    IntPtr excelHwnd = IntPtr.Zero;
                    try
                    {
                        excelHwnd = new IntPtr((int)app.Hwnd);
                    }
                    catch { }

                    if (excelHwnd != IntPtr.Zero)
                    {
                        _overlayForm.Show(new Win32WindowWrapper(excelHwnd));
                    }
                    else
                    {
                        _overlayForm.Show();
                    }

                    // 7. 启动纯内存高速动画定时器
                    if (_animTimer == null)
                    {
                        _animTimer = new System.Windows.Forms.Timer();
                        // 40ms 刷新率，约 25 帧/秒 --硬编码: 动画帧间隔 40ms--
                        _animTimer.Interval = 40;
                        _animTimer.Tick += OnAnimTimerTick;
                    }
                    _animTimer.Start();
                }
                catch (Exception ex)
                {
                    // 记录异常日志
                    LogHelper.WriteLog($"[MarchingAntsManager] 启动橙色流动虚线异常: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 停止并销毁橙色流动细虚线动效 (插入粘贴后、删除元件或按ESC时调用)
        /// </summary>
        public static void Hide()
        {
            // 加锁安全停止
            lock (_syncRoot)
            {
                try
                {
                    HideInternal();
                }
                catch (Exception ex)
                {
                    // 记录异常日志
                    LogHelper.WriteLog($"[MarchingAntsManager] 隐藏橙色流动虚线异常: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 内部物理停止与彻底销毁释放浮窗实例
        /// </summary>
        private static void HideInternal()
        {
            // 标记置为未激活
            _isActive = false;

            // 停止动画定时器
            if (_animTimer != null)
            {
                _animTimer.Stop();
            }

            // 彻底关闭并释放浮窗实例，置空引用
            if (_overlayForm != null)
            {
                try
                {
                    _overlayForm.Hide();
                    _overlayForm.Close();
                    _overlayForm.Dispose();
                }
                catch { }
                _overlayForm = null;
            }

            // 重置几何参数状态
            _targetSheetName = string.Empty;
            _startRow = 0;
            _endRow = 0;
            _dashOffset = 0f;
            _tickCounter = 0;
        }

        /// <summary>
        /// 动画定时器滴答事件：高速纯内存推进流动偏移量，低频校准坐标
        /// </summary>
        private static void OnAnimTimerTick(object? sender, EventArgs e)
        {
            // 若未激活或浮窗已被释放直接退出
            if (!_isActive || _overlayForm == null || _overlayForm.IsDisposed)
            {
                Hide();
                return;
            }

            // 1. 监听用户是否按下了 ESC 键，若按下则自动平滑消退动效
            try
            {
                if ((GetAsyncKeyState(VK_ESCAPE) & 0x8000) != 0 && IsExcelForeground())
                {
                    Hide();
                    return;
                }
            }
            catch { }

            // 2. 高速推进虚线流动偏移量 (每次递增 0.8px，周期为 6.0px) --硬编码: 流动步长 0.8px，周期 6.0px--
            _dashOffset = (_dashOffset + 0.8f) % 6.0f;
            _overlayForm.UpdateDashOffset(_dashOffset);

            // 3. 低频安全校准位置 (每 10 帧/约 400ms 校准一次)，避免高频 COM 阻塞
            _tickCounter++;
            if (_tickCounter % 10 == 0)
            {
                SyncPositionLowFrequency();
            }
        }

        /// <summary>
        /// 低频安全校准浮窗绝对屏幕位置 (适应视口滚动与缩放)
        /// </summary>
        private static void SyncPositionLowFrequency()
        {
            try
            {
                // 校验基础状态
                if (_overlayForm == null || _overlayForm.IsDisposed || !_isActive) return;

                // 若 Excel 当前被最小化，隐藏浮窗
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return;

                // 获取活动工作表
                dynamic? activeSheet = app.ActiveSheet;
                if (activeSheet == null) return;

                // 若用户切换到了其他工作表，隐藏浮窗
                string curSheetName = Convert.ToString(activeSheet.Name) ?? "";
                if (!string.Equals(curSheetName, _targetSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    if (_overlayForm.Visible) _overlayForm.Visible = false;
                    return;
                }

                // 重新换算当前屏幕像素矩形
                string startColLetter = GetColumnLetter(1);
                string endColLetter = GetColumnLetter(_endCol);
                string rangeAddress = $"{startColLetter}{_startRow}:{endColLetter}{_endRow}";
                dynamic targetRange = activeSheet.Range[rangeAddress];

                Rectangle screenRect = CalculateRangeScreenRect(app.ActiveWindow, targetRange);
                if (screenRect.Width <= 4 || screenRect.Height <= 4)
                {
                    if (_overlayForm.Visible) _overlayForm.Visible = false;
                    return;
                }

                // 外扩 1 像素包裹边框外沿
                screenRect.Inflate(1, 1);

                // 更新位置
                if (_overlayForm.Bounds != screenRect)
                {
                    _overlayForm.SetBounds(screenRect.X, screenRect.Y, screenRect.Width, screenRect.Height);
                }

                // 确保浮窗可见
                if (!_overlayForm.Visible)
                {
                    _overlayForm.Visible = true;
                }
            }
            catch
            {
                // 忽略偶发 COM 繁忙异常
            }
        }

        /// <summary>
        /// 利用 Excel 视口 Panes 精准计算 Range 区域在显示器上的物理像素矩形
        /// </summary>
        private static Rectangle CalculateRangeScreenRect(dynamic activeWin, dynamic targetRange)
        {
            try
            {
                // 读取区域在 Excel 逻辑坐标系下的坐标与尺寸 (以 point 为单位)
                double rLeft = Convert.ToDouble(targetRange.Left);
                double rTop = Convert.ToDouble(targetRange.Top);
                double rWidth = Convert.ToDouble(targetRange.Width);
                double rHeight = Convert.ToDouble(targetRange.Height);

                // 优先使用 Panes 进行高精度自适应屏幕像素换算
                if (activeWin != null && activeWin.Panes != null && activeWin.Panes.Count > 0)
                {
                    try
                    {
                        // 获取当前活动视口窗格
                        dynamic pane = activeWin.ActivePane ?? activeWin.Panes[1];
                        // 换算左上角绝对屏幕物理像素
                        int px1 = pane.PointsToScreenPixelsX((int)rLeft);
                        int py1 = pane.PointsToScreenPixelsY((int)rTop);
                        // 换算右下角绝对屏幕物理像素
                        int px2 = pane.PointsToScreenPixelsX((int)(rLeft + rWidth));
                        int py2 = pane.PointsToScreenPixelsY((int)(rTop + rHeight));

                        // 校验换算坐标合法性
                        if (px1 > 0 && py1 > 0 && px2 > px1 && py2 > py1)
                        {
                            return new Rectangle(px1, py1, px2 - px1, py2 - py1);
                        }
                    }
                    catch { }
                }
            }
            catch { }

            // 失败返回空矩形
            return Rectangle.Empty;
        }

        /// <summary>
        /// 将 1-based 数字列索引转换为 Excel 列英文字母 (如 1 -> A, 20 -> T)
        /// </summary>
        private static string GetColumnLetter(int colIndex)
        {
            // 防御性校验
            if (colIndex <= 0) return "A";
            string colLetter = string.Empty;
            // 循环处理 26 进制字符映射
            while (colIndex > 0)
            {
                int modulo = (colIndex - 1) % 26;
                colLetter = Convert.ToChar(65 + modulo) + colLetter;
                colIndex = (colIndex - modulo) / 26;
            }
            return colLetter;
        }

        /// <summary>
        /// 判断当前系统前台激活窗口是否属于 Excel 主进程
        /// </summary>
        private static bool IsExcelForeground()
        {
            try
            {
                // 获取当前获得焦点的窗口句柄
                IntPtr fgHwnd = GetForegroundWindow();
                if (fgHwnd == IntPtr.Zero) return false;

                // 获取该窗口所属进程 PID
                GetWindowThreadProcessId(fgHwnd, out uint fgPid);
                // 与当前 Excel 进程 PID 比较
                return fgPid == GetCurrentProcessId();
            }
            catch
            {
                return true;
            }
        }
    }

    /// <summary>
    /// Win32 HWND 原生窗口包装器，用于将 Excel 主窗口安全指定为浮窗 Owner
    /// </summary>
    public class Win32WindowWrapper : IWin32Window
    {
        // 目标原生窗口句柄
        private readonly IntPtr _hwnd;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="handle">原生 HWND</param>
        public Win32WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }

        /// <summary>
        /// 实现 IWin32Window 接口的 Handle 属性
        /// </summary>
        public IntPtr Handle => _hwnd;
    }
}
