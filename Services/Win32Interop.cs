using System;
using System.Runtime.InteropServices;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// Win32 操作系统底层互操作 API 封装类
    /// 用于 AutoCAD 窗口句柄宿主、样式重置、焦点切换与位置同步
    /// </summary>
    public static class Win32Interop
    {
        // 窗口样式索引起始位：标准样式（32位整型）--硬编码: Win32标准常量--
        public const int GWL_STYLE = -16;

        // 窗口扩展样式索引起始位（32位整型）--硬编码: Win32标准常量--
        public const int GWL_EXSTYLE = -20;

        // 窗口样式掩码：子窗口样式
        public const int WS_CHILD = 0x40000000;

        // 窗口样式掩码：弹出窗口样式
        public const int WS_POPUP = unchecked((int)0x80000000);

        // 窗口样式掩码：可见样式
        public const int WS_VISIBLE = 0x10000000;

        // 窗口样式掩码：裁切子窗口区域，防止重绘闪烁
        public const int WS_CLIPCHILDREN = 0x02000000;

        // 窗口样式掩码：裁切同级兄弟窗口区域
        public const int WS_CLIPSIBLINGS = 0x04000000;

        // 窗口样式掩码：标题栏组合样式（包含边框与标题文字）
        public const int WS_CAPTION = 0x00C00000;

        // 窗口样式掩码：可调整大小的厚边框
        public const int WS_THICKFRAME = 0x00040000;

        // 窗口样式掩码：最小化按钮
        public const int WS_MINIMIZEBOX = 0x00020000;

        // 窗口样式掩码：最大化按钮
        public const int WS_MAXIMIZEBOX = 0x00010000;

        // 窗口样式掩码：系统菜单
        public const int WS_SYSMENU = 0x00080000;

        // 显示窗口标志：以原大小和位置激活并显示窗口
        public const int SW_SHOW = 5;

        // 显示窗口标志：激活并以最小化状态显示
        public const int SW_MINIMIZE = 6;

        // 显示窗口标志：激活并最大化窗口
        public const int SW_MAXIMIZE = 3;

        // 显示窗口标志：还原窗口（退出最大化/最小化状态）
        public const int SW_RESTORE = 9;

        /// <summary>
        /// Win32 矩形结构体
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            // 左边界 X 坐标
            public int Left;
            // 顶边界 Y 坐标
            public int Top;
            // 右边界 X 坐标
            public int Right;
            // 底边界 Y 坐标
            public int Bottom;

            // 获取矩形宽度
            public int Width => Right - Left;
            // 获取矩形高度
            public int Height => Bottom - Top;
        }

        // 设置指定窗口的父窗口句柄
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        // 获取指定窗口的直接父窗口句柄
        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        public static extern IntPtr GetParent(IntPtr hWnd);

        // 32 位系统下读取窗口指定索引的属性长整型值
        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        // 64 位系统下读取窗口指定索引的属性指针长整型值
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        // 32 位系统下设置窗口指定索引的属性值
        [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        // 64 位系统下设置窗口指定索引的属性指针值
        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        /// <summary>
        /// 跨 32 位与 64 位架构自适应读取窗口样式的安全方法
        /// </summary>
        public static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex)
        {
            // 根据当前操作系统指针宽度决定调用 64 位或 32 位底层接口
            if (IntPtr.Size == 8)
            {
                // 64 位架构调用 GetWindowLongPtr
                return GetWindowLongPtr64(hWnd, nIndex);
            }
            // 32 位架构回退调用 GetWindowLong
            return new IntPtr(GetWindowLong32(hWnd, nIndex));
        }

        /// <summary>
        /// 跨 32 位与 64 位架构自适应设置窗口样式的安全方法
        /// </summary>
        public static IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
        {
            // 根据当前操作系统指针宽度决定调用 64 位或 32 位底层接口
            if (IntPtr.Size == 8)
            {
                // 64 位架构调用 SetWindowLongPtr
                return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
            }
            // 32 位架构回退调用 SetWindowLong
            return new IntPtr(SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32()));
        }

        // 移动窗口并调整其物理尺寸与重绘状态
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        // 控制窗口的显示状态（显示、隐藏、最大化、还原）
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        // 为指定窗口赋予键盘焦点，使键盘输入直接流入该窗口消息循环
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetFocus(IntPtr hWnd);

        // 将指定窗口拉至桌面顶层并激活（拉入前台）
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        // 获取指定窗口在屏幕坐标系下的包围矩形
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        // 判断指定句柄是否为当前操作系统中的有效活动窗口
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);
    }
}
