using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// AutoCAD 实例信息模型
    /// 封装活动 CAD 进程的窗口句柄、文档名称、文件全路径与进程标识
    /// </summary>
    public class CadInstanceInfo
    {
        // AutoCAD 主窗口物理句柄
        public IntPtr Hwnd { get; set; } = IntPtr.Zero;

        // 当前处于活动编辑状态的图纸名称
        public string DocName { get; set; } = string.Empty;

        // 当前图纸在磁盘上的绝对路径
        public string DocPath { get; set; } = string.Empty;

        // AutoCAD 进程 ID
        public int ProcessId { get; set; } = 0;

        // 当前 AutoCAD 实例是否有效且存活在操作系统中
        public bool IsConnected => Hwnd != IntPtr.Zero && Win32Interop.IsWindow(Hwnd);
    }

    /// <summary>
    /// AutoCAD 窗口与屏幕分屏协同管理器
    /// 彻底剔除 SetParent 嵌入与任务窗格面板，仅保留原生顶层窗口 5:5 均等分屏核心逻辑
    /// 保障 AutoCAD 100% 原生绘图性能、命令行输入、鼠标中键拖拽平移无损
    /// </summary>
    public static class CadEmbedManager
    {
        // 记录当前是否处于并排分屏状态
        private static bool _isSideBySideSplit = false;

        // 操作互斥锁
        private static readonly object _syncLock = new object();

        /// <summary>
        /// 极简核心交互：一键 5:5 均等分屏与全屏还原切换
        /// 点击后将 Excel 靠左 50%、AutoCAD 靠右 50% 均等排列；再次点击则还原 Excel 全屏最大化
        /// </summary>
        /// <param name="excelRatio">分屏比例，默认 0.5 即 5:5 均等分</param>
        public static void ToggleSideBySide(double excelRatio = 0.5)
        {
            lock (_syncLock)
            {
                // 若当前已处于并排分屏状态，则触发一键还原全屏最大化
                if (_isSideBySideSplit)
                {
                    // 还原 Excel 最大化全屏
                    MaximizeExcel();
                    // 复位分屏标志
                    _isSideBySideSplit = false;
                    return;
                }

                // 否则执行 5:5 均等并排分屏
                var (ok, msg) = SnapSideBySide(excelRatio);
                if (ok)
                {
                    // 标记分屏成功
                    _isSideBySideSplit = true;
                }
                else
                {
                    // 若未检测到 CAD 或分屏失败，弹出提示
                    MessageBox.Show(msg, "CAD 分屏提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    // 复位标志
                    _isSideBySideSplit = false;
                }
            }
        }

        /// <summary>
        /// 向后兼容方法：直接触发 5:5 均等分屏切换
        /// </summary>
        public static void ToggleTaskPane()
        {
            // 执行 5:5 均等分屏
            ToggleSideBySide(0.5);
        }

        /// <summary>
        /// 向后兼容方法：直接触发 5:5 均等分屏
        /// </summary>
        public static void ShowTaskPane()
        {
            // 执行 5:5 均等分屏
            ToggleSideBySide(0.5);
        }

        /// <summary>
        /// 执行智能一键左右并排分屏：Excel 靠左，AutoCAD 靠右
        /// 两者均为 Windows 原生顶级独立窗口，拥有极致的原生性能与画图体验
        /// </summary>
        /// <param name="excelRatio">Excel 占据当前屏幕宽度的比例 (默认 0.5 代表 5:5 平分)</param>
        /// <returns>操作结果元组 (是否成功, 提示消息)</returns>
        public static (bool Success, string Message) SnapSideBySide(double excelRatio = 0.5)
        {
            try
            {
                // 1. 获取 Excel 主窗口句柄
                IntPtr excelHwnd = GetExcelMainWindowHandle();
                if (excelHwnd == IntPtr.Zero || !Win32Interop.IsWindow(excelHwnd))
                {
                    return (false, "无法获取 Excel 主窗口句柄！");
                }

                // 2. 检测当前活跃的 AutoCAD 实例
                CadInstanceInfo cadInfo = GetActiveCadInfo();
                if (!cadInfo.IsConnected)
                {
                    return (false, "未检测到正在运行的 AutoCAD！请先启动 AutoCAD 并打开工程图纸。");
                }

                // 3. 读取当前 Excel 窗口所在的物理显示器屏幕
                Screen screen = Screen.FromHandle(excelHwnd);
                // 获取排除 Windows 任务栏后的实际可用工作区域
                Rectangle workArea = screen.WorkingArea;

                // 4. 计算两侧窗口的物理像素尺寸
                int excelWidth = (int)(workArea.Width * excelRatio);
                int cadWidth = workArea.Width - excelWidth;

                // 5. 调整 Excel 主窗口：退出最大化还原，定位至左半屏
                Win32Interop.ShowWindow(excelHwnd, Win32Interop.SW_RESTORE);
                Win32Interop.MoveWindow(excelHwnd, workArea.Left, workArea.Top, excelWidth, workArea.Height, true);

                // 6. 调整 AutoCAD 主窗口：退出最小化还原，定位至右半屏
                Win32Interop.ShowWindow(cadInfo.Hwnd, Win32Interop.SW_RESTORE);
                Win32Interop.MoveWindow(cadInfo.Hwnd, workArea.Left + excelWidth, workArea.Top, cadWidth, workArea.Height, true);

                // 7. 协同激活：依次激活 CAD 与 Excel，确保两边窗口并排在前台
                Win32Interop.SetForegroundWindow(cadInfo.Hwnd);
                Win32Interop.SetForegroundWindow(excelHwnd);

                return (true, $"已成功启动 5:5 并排分屏协同！");
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[CadEmbedManager] 智能分屏异常: {ex.Message}");
                return (false, $"分屏执行失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取当前 Excel 主窗口物理句柄
        /// </summary>
        public static IntPtr GetExcelMainWindowHandle()
        {
            try
            {
                // 优先通过 Excel-DNA 内置接口获取
                IntPtr hwnd = (IntPtr)ExcelDnaUtil.WindowHandle;
                if (hwnd != IntPtr.Zero && Win32Interop.IsWindow(hwnd))
                {
                    return hwnd;
                }
            }
            catch { }

            // 降级使用当前宿主进程 MainWindowHandle
            return Process.GetCurrentProcess().MainWindowHandle;
        }

        /// <summary>
        /// 将当前 Excel 窗口恢复全屏最大化显示
        /// </summary>
        public static void MaximizeExcel()
        {
            try
            {
                // 读取 Excel 窗口句柄
                IntPtr excelHwnd = GetExcelMainWindowHandle();
                if (excelHwnd != IntPtr.Zero && Win32Interop.IsWindow(excelHwnd))
                {
                    // 调用 Win32 API 还原全屏最大化
                    Win32Interop.ShowWindow(excelHwnd, Win32Interop.SW_MAXIMIZE);
                    Win32Interop.SetForegroundWindow(excelHwnd);
                }
            }
            catch { }
        }

        /// <summary>
        /// 将 AutoCAD 主窗口快速唤醒并置顶到用户桌面最前端
        /// </summary>
        public static void ActivateCad()
        {
            try
            {
                // 获取当前 CAD 实例
                CadInstanceInfo cadInfo = GetActiveCadInfo();
                if (cadInfo.IsConnected)
                {
                    // 先还原显示再置顶
                    Win32Interop.ShowWindow(cadInfo.Hwnd, Win32Interop.SW_RESTORE);
                    Win32Interop.SetForegroundWindow(cadInfo.Hwnd);
                }
            }
            catch { }
        }

        /// <summary>
        /// 探测当前系统中处于活动状态的单个 AutoCAD 进程、窗口与图纸信息
        /// </summary>
        /// <returns>AutoCAD 实例信息结构模型</returns>
        public static CadInstanceInfo GetActiveCadInfo()
        {
            var info = new CadInstanceInfo();

            // 策略 1: 优先尝试通过 COM 获取当前活跃的 AutoCAD 实例详细属性
            try
            {
                // 获取当前活动 COM 实例
                dynamic? acadApp = GetActiveAcadApp();
                if (acadApp != null)
                {
                    // 提取主窗口句柄
                    info.Hwnd = new IntPtr((long)acadApp.HWND);
                    // 提取活动文档名与全路径
                    try { info.DocName = (string)acadApp.ActiveDocument?.Name ?? string.Empty; } catch { }
                    try { info.DocPath = (string)acadApp.ActiveDocument?.FullName ?? string.Empty; } catch { }
                    // 提取进程 PID
                    try { info.ProcessId = (int)acadApp.ProcessId; } catch { }

                    // 若句柄有效且为有效窗口则直接返回
                    if (info.IsConnected)
                    {
                        return info;
                    }
                }
            }
            catch { }

            // 策略 2: 降级通过 acad 进程列表扫描主窗口
            try
            {
                // 获取所有名为 acad 的进程
                var acadProcesses = Process.GetProcessesByName("acad");
                foreach (var p in acadProcesses)
                {
                    // 校验是否拥有主窗口句柄
                    if (p.MainWindowHandle != IntPtr.Zero && Win32Interop.IsWindow(p.MainWindowHandle))
                    {
                        info.Hwnd = p.MainWindowHandle;
                        info.DocName = p.MainWindowTitle;
                        info.ProcessId = p.Id;
                        return info;
                    }
                }
            }
            catch { }

            return info;
        }

        /// <summary>
        /// 安全获取当前处于活动运行状态的 AutoCAD COM Application 实例
        /// 支持 AutoCAD.Application 以及多版本 ProgID (2016~2026) 自动降级探测
        /// </summary>
        /// <returns>成功获取返回 dynamic COM 实例，否则返回 null</returns>
        public static dynamic? GetActiveAcadApp()
        {
            // 1. 优先尝试标准无版本后缀 ProgID
            try
            {
                // 从系统运行对象表 ROT 读取活动实例
                dynamic app = Marshal.GetActiveObject("AutoCAD.Application");
                // 命中立即返回
                if (app != null) return app;
            }
            catch { }

            // 2. 降级遍历探测各版本 ProgID (支持 2016 ~ 2026)
            string[] progIds = new string[]
            {
                "AutoCAD.Application.25",   // AutoCAD 2025
                "AutoCAD.Application.24.3", // AutoCAD 2024
                "AutoCAD.Application.24.2", // AutoCAD 2023
                "AutoCAD.Application.24.1", // AutoCAD 2022
                "AutoCAD.Application.24",   // AutoCAD 2021
                "AutoCAD.Application.23.1", // AutoCAD 2020
                "AutoCAD.Application.23",   // AutoCAD 2019
                "AutoCAD.Application.22",   // AutoCAD 2018
                "AutoCAD.Application.21",   // AutoCAD 2017
                "AutoCAD.Application.20"    // AutoCAD 2016
            };

            // 循环遍历探测
            foreach (string pid in progIds)
            {
                try
                {
                    // 尝试以指定版本 ProgID 读取
                    dynamic app = Marshal.GetActiveObject(pid);
                    // 命中即刻返回
                    if (app != null) return app;
                }
                catch { }
            }

            // 全部未命中返回 null
            return null;
        }

        /// <summary>
        /// 插件卸载或退出时的清理方法
        /// </summary>
        public static void Cleanup()
        {
            // 重置分屏状态标记
            _isSideBySideSplit = false;
        }
    }
}
