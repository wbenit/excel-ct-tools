using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using ExcelAddInDemo.Forms;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// AutoCAD 窗口嵌入式任务窗格统一管理器
    /// 负责 CustomTaskPane 的初始化、Win32 句柄挂载、生命周期管理与安全还原
    /// </summary>
    public static class CadEmbedManager
    {
        // Excel-DNA 自定义任务窗格引用
        private static CustomTaskPane? _taskPane;

        // 承载 AutoCAD 的 WinForms 宿主控件
        private static CadHostControl? _hostControl;

        // 当前正在嵌入中的 AutoCAD 主窗口句柄
        private static IntPtr _embeddedCadHwnd = IntPtr.Zero;

        // 嵌入前 AutoCAD 原始父句柄（桌面句柄）
        private static IntPtr _originalParent = IntPtr.Zero;

        // 嵌入前 AutoCAD 原始窗口样式
        private static IntPtr _originalStyle = IntPtr.Zero;

        // 嵌入前 AutoCAD 原始窗口在屏幕中的尺寸与坐标
        private static Win32Interop.RECT _originalRect;

        // 当前绑定的 AutoCAD 进程对象引用
        private static Process? _boundCadProcess;

        // 任务窗格默认初始宽度 --硬编码: 默认窗格宽度 680 像素--
        private const int DefaultTaskPaneWidth = 680;

        // AutoCAD 主程序顶层标题栏裁剪偏移高度（纯画布模式下向上移出视口以消除多余标题栏）--硬编码: 标题栏裁剪高度 32 像素--
        private const int CadTitleBarOffset = 32;

        // 是否开启纯画布模式（自动隐藏所有工具栏与 Ribbon 功能区，默认开启）
        public static bool IsCleanScreenEnabled { get; set; } = true;

        // 互斥操作锁
        private static readonly object _syncLock = new object();

        /// <summary>
        /// 切换 AutoCAD 协同任务窗格的显示与隐藏状态
        /// </summary>
        public static void ToggleTaskPane()
        {
            lock (_syncLock)
            {
                // 若任务窗格尚未创建，执行首次初始化与呈现
                if (_taskPane == null)
                {
                    // 实例化 WinForms 宿主控件
                    _hostControl = new CadHostControl();

                    // 通过 Excel-DNA 工厂构建右侧自定义任务窗格
                    _taskPane = CustomTaskPaneFactory.CreateCustomTaskPane(_hostControl, "AutoCAD 协同画图");

                    // 设定停靠在 Excel 主界面右侧
                    _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;

                    // 设置默认宽度
                    _taskPane.Width = DefaultTaskPaneWidth;

                    // 监听窗格可见性变更事件（如用户点击右上角叉号关闭窗格）
                    _taskPane.VisibleStateChange += OnTaskPaneVisibleStateChange;

                    // 呈现任务窗格
                    _taskPane.Visible = true;

                    // 窗格打开后，自动探测并尝试嵌入当前正在运行的 AutoCAD
                    EmbedActiveCad();
                }
                else
                {
                    // 切换窗格显隐状态
                    _taskPane.Visible = !_taskPane.Visible;

                    // 若重新显示且尚未嵌入 CAD，自动触发探测与嵌入
                    if (_taskPane.Visible && _embeddedCadHwnd == IntPtr.Zero)
                    {
                        EmbedActiveCad();
                    }
                }
            }
        }

        /// <summary>
        /// 显式呈现 AutoCAD 协同任务窗格
        /// </summary>
        public static void ShowTaskPane()
        {
            // 若未打开则开启
            if (_taskPane == null || !_taskPane.Visible)
            {
                ToggleTaskPane();
            }
        }

        /// <summary>
        /// 任务窗格可见性状态改变回调（处理用户手动关闭）
        /// </summary>
        private static void OnTaskPaneVisibleStateChange(CustomTaskPane customTaskPaneInst)
        {
            // 当用户在 Excel 中关闭了任务窗格时
            if (!customTaskPaneInst.Visible)
            {
                // 如果需要避免 CAD 在后台不可见，可在此处选择保留或解绑
                // 保留嵌入状态，用户下次点击 Ribbon 再次打开时无需重复绑定
            }
        }

        /// <summary>
        /// 探测当前系统中处于活动状态的单个 AutoCAD 进程与窗口
        /// 遵循“只绑定激活的单张图纸”原则
        /// </summary>
        /// <returns>包含窗口句柄、图纸标题、进程对象的元组</returns>
        private static (IntPtr Hwnd, string DocName, Process? Proc) FindActiveAutoCad()
        {
            // 策略 1: 优先尝试通过 COM 获取当前活跃的 AutoCAD 实例
            try
            {
                // 从运行对象表 ROT 获取 AutoCAD.Application
                dynamic acadApp = Marshal.GetActiveObject("AutoCAD.Application");
                if (acadApp != null)
                {
                    // 提取主窗口句柄
                    IntPtr hwnd = new IntPtr((long)acadApp.HWND);
                    if (hwnd != IntPtr.Zero && Win32Interop.IsWindow(hwnd))
                    {
                        string docName = string.Empty;
                        try { docName = (string)acadApp.ActiveDocument?.Name ?? string.Empty; } catch { }

                        int pid = 0;
                        try { pid = (int)acadApp.ProcessId; } catch { }

                        Process? proc = pid > 0 ? Process.GetProcessById(pid) : null;
                        return (hwnd, docName, proc);
                    }
                }
            }
            catch
            {
                // COM 未就绪时静默降级为进程枚举
            }

            // 策略 2: 降级通过 acad 进程列表扫描主窗口
            try
            {
                var acadProcesses = Process.GetProcessesByName("acad");
                foreach (var p in acadProcesses)
                {
                    // 必须具备有效窗口句柄且窗口在操作系统中存活
                    if (p.MainWindowHandle != IntPtr.Zero && Win32Interop.IsWindow(p.MainWindowHandle))
                    {
                        string title = p.MainWindowTitle;
                        return (p.MainWindowHandle, title, p);
                    }
                }
            }
            catch
            {
                // 忽略进程枚举异常
            }

            return (IntPtr.Zero, string.Empty, null);
        }

        /// <summary>
        /// 检测并嵌入当前活动的 AutoCAD 窗口到任务窗格
        /// </summary>
        public static bool EmbedActiveCad()
        {
            lock (_syncLock)
            {
                // 确保宿主控件已就绪
                if (_hostControl == null) return false;

                // 探测当前运行的 AutoCAD 实例
                var (cadHwnd, docName, proc) = FindActiveAutoCad();
                if (cadHwnd == IntPtr.Zero)
                {
                    // 未检测到运行中的 CAD，更新界面为空状态
                    _hostControl.UpdateConnectionState(false);
                    return false;
                }

                // 如果当前检测到的句柄与已嵌入的句柄相同，只需同步尺寸并赋予焦点
                if (_embeddedCadHwnd == cadHwnd)
                {
                    SyncCadSize();
                    FocusCad();
                    return true;
                }

                // 若之前已嵌入其他 CAD 窗口，先安全还原旧窗口
                if (_embeddedCadHwnd != IntPtr.Zero)
                {
                    DetachCad();
                }

                try
                {
                    // 1. 记录 AutoCAD 窗口原始属性以便后续完美还原
                    _originalParent = Win32Interop.GetParent(cadHwnd);
                    _originalStyle = Win32Interop.GetWindowLongPtr(cadHwnd, Win32Interop.GWL_STYLE);
                    Win32Interop.GetWindowRect(cadHwnd, out _originalRect);

                    // 2. 剥离标题栏、独立边框与最大/最小化按钮，赋予子窗口属性
                    int style = _originalStyle.ToInt32();
                    style &= ~Win32Interop.WS_POPUP;
                    style &= ~Win32Interop.WS_CAPTION;
                    style &= ~Win32Interop.WS_THICKFRAME;
                    style &= ~Win32Interop.WS_MINIMIZEBOX;
                    style &= ~Win32Interop.WS_MAXIMIZEBOX;
                    style &= ~Win32Interop.WS_SYSMENU;
                    style |= Win32Interop.WS_CHILD | Win32Interop.WS_VISIBLE | Win32Interop.WS_CLIPCHILDREN | Win32Interop.WS_CLIPSIBLINGS;

                    // 设置精简后的子窗口样式
                    Win32Interop.SetWindowLongPtr(cadHwnd, Win32Interop.GWL_STYLE, new IntPtr(style));

                    // 3. 将 AutoCAD 句柄的父容器重定向为 WinForms Panel
                    IntPtr containerHwnd = _hostControl.GetContainerHandle();
                    Win32Interop.SetParent(cadHwnd, containerHwnd);

                    // 4. 同步视口尺寸
                    Size containerSize = _hostControl.GetContainerSize();
                    Win32Interop.MoveWindow(cadHwnd, 0, 0, containerSize.Width, containerSize.Height, true);
                    Win32Interop.ShowWindow(cadHwnd, Win32Interop.SW_SHOW);

                    // 5. 注册进程退出监控：防止 CAD 意外退出时残留白屏
                    _boundCadProcess = proc;
                    if (_boundCadProcess != null)
                    {
                        try
                        {
                            _boundCadProcess.EnableRaisingEvents = true;
                            _boundCadProcess.Exited += OnCadProcessExited;
                        }
                        catch { }
                    }

                    // 6. 保存当前嵌入状态并刷新宿主 UI 顶栏
                    _embeddedCadHwnd = cadHwnd;
                    _hostControl.UpdateConnectionState(true, docName);

                    // 7. 若开启纯画布模式，自动隐藏 AutoCAD 所有工具栏与 Ribbon 功能区
                    if (IsCleanScreenEnabled)
                    {
                        SetCadCleanScreen(true);
                    }

                    // 赋予键盘焦点
                    FocusCad();
                    return true;
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[CadEmbedManager] 嵌入 AutoCAD 异常: {ex.Message}");
                    _hostControl.UpdateConnectionState(false);
                    return false;
                }
            }
        }

        /// <summary>
        /// 将 AutoCAD 窗口安全从任务窗格剥离并归还至桌面独立窗口
        /// </summary>
        public static void DetachCad()
        {
            lock (_syncLock)
            {
                // 若当前没有正在嵌入的 CAD 窗口，直接退出
                if (_embeddedCadHwnd == IntPtr.Zero) return;

                try
                {
                    // 0. 将 AutoCAD 工具栏与 Ribbon 功能区完整恢复
                    SetCadCleanScreen(false);

                    // 解绑进程退出监听
                    if (_boundCadProcess != null)
                    {
                        try { _boundCadProcess.Exited -= OnCadProcessExited; } catch { }
                        _boundCadProcess = null;
                    }

                    // 1. 重置父窗口为桌面（IntPtr.Zero）
                    Win32Interop.SetParent(_embeddedCadHwnd, _originalParent);

                    // 2. 恢复原生的窗口样式（带标题栏、控制按钮与边框）
                    if (_originalStyle != IntPtr.Zero)
                    {
                        Win32Interop.SetWindowLongPtr(_embeddedCadHwnd, Win32Interop.GWL_STYLE, _originalStyle);
                    }

                    // 3. 恢复原始位置与尺寸
                    if (_originalRect.Width > 100 && _originalRect.Height > 100)
                    {
                        Win32Interop.MoveWindow(
                            _embeddedCadHwnd,
                            _originalRect.Left,
                            _originalRect.Top,
                            _originalRect.Width,
                            _originalRect.Height,
                            true);
                    }

                    // 4. 显示并还原窗口状态
                    Win32Interop.ShowWindow(_embeddedCadHwnd, Win32Interop.SW_RESTORE);
                    Win32Interop.SetForegroundWindow(_embeddedCadHwnd);
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[CadEmbedManager] 归还 AutoCAD 异常: {ex.Message}");
                }
                finally
                {
                    // 重置状态
                    _embeddedCadHwnd = IntPtr.Zero;
                    _originalStyle = IntPtr.Zero;
                    _originalParent = IntPtr.Zero;
                    _hostControl?.UpdateConnectionState(false);
                }
            }
        }

        /// <summary>
        /// 当绑定的 AutoCAD 进程在外部被用户关闭或崩溃时的容灾清理
        /// </summary>
        private static void OnCadProcessExited(object? sender, EventArgs e)
        {
            lock (_syncLock)
            {
                // 重置句柄状态
                _embeddedCadHwnd = IntPtr.Zero;
                _boundCadProcess = null;

                // 通知宿主控件更新为空状态
                _hostControl?.UpdateConnectionState(false);
            }
        }

        /// <summary>
        /// 同步 AutoCAD 窗口尺寸以铺满宿主容器
        /// 当开启纯画布模式时，向上偏移裁剪 AutoCAD 原生标题栏
        /// </summary>
        public static void SyncCadSize()
        {
            // 校验当前嵌入句柄是否有效
            if (_embeddedCadHwnd != IntPtr.Zero && _hostControl != null)
            {
                Size size = _hostControl.GetContainerSize();
                if (size.Width > 0 && size.Height > 0)
                {
                    // 若开启纯画布模式，将 AutoCAD 窗口向上偏移以裁剪顶层标题栏，彻底消除多余外壳
                    int yOffset = IsCleanScreenEnabled ? -CadTitleBarOffset : 0;
                    // 同步补偿增加高度，确保底端画图区与状态栏完整充满
                    int extraHeight = IsCleanScreenEnabled ? CadTitleBarOffset : 0;

                    // 调整 CAD 窗口尺寸与物理偏移
                    Win32Interop.MoveWindow(_embeddedCadHwnd, 0, yOffset, size.Width, size.Height + extraHeight, true);
                }
            }
        }

        /// <summary>
        /// 将键盘输入焦点切换至 AutoCAD 窗口，确保绘图快捷键（LINE, ESC, 空格）直接生效
        /// </summary>
        public static void FocusCad()
        {
            if (_embeddedCadHwnd != IntPtr.Zero && Win32Interop.IsWindow(_embeddedCadHwnd))
            {
                // 赋予焦点
                Win32Interop.SetFocus(_embeddedCadHwnd);
            }
        }

        /// <summary>
        /// 设置 AutoCAD 是否开启极简纯净视口模式
        /// 开启时自动隐藏所有工具栏、Ribbon功能区、文件标签与视口辅助件，只保留纯绘图画布
        /// </summary>
        /// <param name="enableClean">是否启用纯画布模式</param>
        public static void SetCadCleanScreen(bool enableClean)
        {
            try
            {
                // 优先通过 COM 接口指令控制 AutoCAD 功能区与工具栏显隐
                dynamic acadApp = Marshal.GetActiveObject("AutoCAD.Application");
                if (acadApp != null && acadApp.ActiveDocument != null)
                {
                    if (enableClean)
                    {
                        // 1. 发送 (command) 退出当前可能处于活动中的命令（安全避开非法的 ASCII 27 转义字符）
                        // 2. CLEANSCREENON 隐藏所有停靠/浮动工具栏与系统菜单
                        // 3. RIBBONCLOSE 彻底隐藏顶部功能区
                        // 4. FILETABCLOSE 隐藏顶部文件标签栏（开始、图纸名标签）
                        // 5. NAVVCUBEDISPLAY 0 隐藏右上角 ViewCube 视角立方体
                        // 6. NAVBAR 0 隐藏右侧悬浮导航栏
                        // 7. (setvar "layouttab" 0) 隐藏底端布局标签栏
                        string cleanCmd = "(command)\nCLEANSCREENON\nRIBBONCLOSE\nFILETABCLOSE\nNAVVCUBEDISPLAY 0\nNAVBAR 0\n(setvar \"layouttab\" 0)\n";
                        acadApp.ActiveDocument.SendCommand(cleanCmd);
                    }
                    else
                    {
                        // 还原全功能模式：恢复功能区、工具栏、文件标签、ViewCube、导航栏与布局标签
                        string restoreCmd = "(command)\nCLEANSCREENOFF\nRIBBON\nFILETAB\nNAVVCUBEDISPLAY 3\nNAVBAR 1\n(setvar \"layouttab\" 1)\n";
                        acadApp.ActiveDocument.SendCommand(restoreCmd);
                    }
                }

                // 同步刷新物理视口剪裁位置
                SyncCadSize();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CadEmbedManager] 切换 CAD 纯净视口模式异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 插件卸载或 Excel 关闭时的全局安全清理方法
        /// 彻底杜绝 AutoCAD 随 Excel 关闭而异常崩溃或被宿主连带销毁
        /// </summary>
        public static void Cleanup()
        {
            // 强制将 CAD 还原回独立桌面状态
            DetachCad();

            // 隐藏任务窗格
            if (_taskPane != null)
            {
                try
                {
                    _taskPane.Visible = false;
                }
                catch { }
            }
        }
    }
}
