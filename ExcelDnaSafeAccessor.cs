using System;
using System.Runtime.CompilerServices;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel-DNA COM/环境安全访问器
    /// 通过方法不内联 (NoInlining) 隔离 ExcelDna.Integration.dll 依赖
    /// 确保在 AutoCAD (TuFan) 或外部非 Excel 进程调用时，不会因 JIT 探测未加载的 ExcelDna.Integration 而抛出 FileNotFoundException
    /// </summary>
    internal static class ExcelDnaSafeAccessor
    {
        /// <summary>
        /// 安全获取 ExcelDna Application 实例（若非 ExcelDna 环境则安全返回 null）
        /// </summary>
        /// <returns>Excel Application COM 对象，无法获取时返回 null</returns>
        [MethodImpl(MethodImplOptions.NoInlining)]
        public static dynamic? GetApplication()
        {
            try
            {
                // 尝试从 Excel-DNA 集成接口获取宿主 Application
                return ExcelDna.Integration.ExcelDnaUtil.Application;
            }
            catch
            {
                // 捕获程序集未加载或类初始化异常并安全返回 null
                return null;
            }
        }

        /// <summary>
        /// 安全获取当前 XLL 文件物理路径（若非 ExcelDna 环境则安全返回 null）
        /// </summary>
        /// <returns>XLL 绝对路径，无法获取时返回 null</returns>
        [MethodImpl(MethodImplOptions.NoInlining)]
        public static string? GetXllPath()
        {
            try
            {
                // 获取 Excel-DNA 的 XLL 文件物理路径
                return ExcelDna.Integration.ExcelDnaUtil.XllPath;
            }
            catch
            {
                // 捕获异常并兜底返回 null
                return null;
            }
        }

        /// <summary>
        /// 安全获取 Excel 主窗口句柄（若非 ExcelDna 环境则返回 IntPtr.Zero）
        /// </summary>
        /// <returns>窗口句柄</returns>
        [MethodImpl(MethodImplOptions.NoInlining)]
        public static IntPtr GetWindowHandle()
        {
            try
            {
                // 获取 Excel 顶级窗口句柄
                return ExcelDna.Integration.ExcelDnaUtil.WindowHandle;
            }
            catch
            {
                // 捕获异常并返回空句柄
                return IntPtr.Zero;
            }
        }
    }
}
