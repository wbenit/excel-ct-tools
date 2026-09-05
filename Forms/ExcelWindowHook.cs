using System;
using System.Windows.Forms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// Excel 视口 Win32 原生消息监听钩子 (对齐 ExWinner 消息防抖节流架构)
    /// 通过 NativeWindow 拦截窗口移动、尺寸变化与滚轮滚动，完全通过防抖定时器异步执行回调
    /// 严禁在 WndProc 内部同步调用 Excel COM 接口，彻底杜绝 UI 线程消息重入死锁
    /// </summary>
    public class ExcelWindowHook : NativeWindow, IDisposable
    {
        // 窗口销毁消息
        private const int WM_DESTROY = 0x0002;
        // 窗口移动消息
        private const int WM_MOVE = 0x0003;
        // 窗口尺寸变更消息
        private const int WM_SIZE = 0x0005;
        private const int WM_ACTIVATE = 0x0006;
        // 窗口失去焦点消息
        private const int WM_KILLFOCUS = 0x0008;
        // 窗口位置与层级综合变更消息
        private const int WM_WINDOWPOSCHANGED = 0x0047;
        // 水平滚动条滚动消息
        private const int WM_HSCROLL = 0x0114;
        // 垂直滚动条滚动消息
        private const int WM_VSCROLL = 0x0115;
        // 鼠标滚轮滑动消息
        private const int WM_MOUSEWHEEL = 0x020A;

        // 窗口边界或位置变动时的异步回调委托
        public Action? OnBoundsChanged { get; set; }

        // 视口内容滚动时的异步回调委托
        public Action? OnScrolled { get; set; }

        // 窗口激活/失活状态变更时的回调委托
        public Action<bool>? OnActivationChanged { get; set; }

        // 视口滚动防抖定时器
        private System.Windows.Forms.Timer? _scrollDebounceTimer;

        // 视口尺寸移动防抖定时器，彻底隔离 WndProc 与 COM 重入
        private System.Windows.Forms.Timer? _boundsDebounceTimer;

        // 释放标志位
        private bool _isDisposed = false;

        /// <summary>
        /// 构造原生窗口监听钩子，并挂载至目标 HWND 句柄
        /// </summary>
        /// <param name="targetHwnd">目标视口窗口句柄 (EXCEL7)</param>
        public ExcelWindowHook(IntPtr targetHwnd)
        {
            // 校验句柄非零且有效
            if (targetHwnd != IntPtr.Zero)
            {
                // 将当前 NativeWindow 实例绑定至操作系统窗口句柄
                this.AssignHandle(targetHwnd);
            }

            // 初始化滚动防抖定时器
            _scrollDebounceTimer = new System.Windows.Forms.Timer
            {
                // 设定防抖时间间隔为 30ms (约 33 FPS 平滑响应)
                Interval = 30 // --硬编码: 滚动防抖间隔 30毫秒--
            };
            _scrollDebounceTimer.Tick += ScrollDebounceTimer_Tick;

            // 初始化尺寸移动防抖定时器
            _boundsDebounceTimer = new System.Windows.Forms.Timer
            {
                // 设定防抖时间间隔为 30ms (合并高频移动缩放消息)
                Interval = 30 // --硬编码: 尺寸移动防抖间隔 30毫秒--
            };
            _boundsDebounceTimer.Tick += BoundsDebounceTimer_Tick;
        }

        /// <summary>
        /// 滚动防抖定时器触发时，安全异步调用视口重绘回调
        /// </summary>
        private void ScrollDebounceTimer_Tick(object? sender, EventArgs e)
        {
            _scrollDebounceTimer?.Stop();
            if (!_isDisposed && OnScrolled != null)
            {
                try
                {
                    OnScrolled.Invoke();
                }
                catch { }
            }
        }

        /// <summary>
        /// 尺寸移动防抖定时器触发时，安全异步调用位置更新回调
        /// </summary>
        private void BoundsDebounceTimer_Tick(object? sender, EventArgs e)
        {
            _boundsDebounceTimer?.Stop();
            if (!_isDisposed && OnBoundsChanged != null)
            {
                try
                {
                    OnBoundsChanged.Invoke();
                }
                catch { }
            }
        }

        /// <summary>
        /// 重写操作系统窗口过程消息处理方法
        /// </summary>
        /// <param name="m">Win32 消息体</param>
        protected override void WndProc(ref Message m)
        {
            // 优先让基类执行默认消息分发，绝不阻断系统原始消息管道
            base.WndProc(ref m);

            // 若对象已销毁则直接返回
            if (_isDisposed) return;

            // 依据消息类型进行分流响应 (仅触发定时器重置，0 耗时，严禁同步调用任何 COM)
            switch (m.Msg)
            {
                // 处理窗口移动或尺寸变化：启动防抖定时器异步处理，消除 COM 重入死锁
                case WM_WINDOWPOSCHANGED:
                case WM_MOVE:
                case WM_SIZE:
                    _boundsDebounceTimer?.Stop();
                    _boundsDebounceTimer?.Start();
                    break;

                // 处理滚轮滑动或滚动条拉动：启动防抖定时器进行节流合并
                case WM_MOUSEWHEEL:
                case WM_VSCROLL:
                case WM_HSCROLL:
                    _scrollDebounceTimer?.Stop();
                    _scrollDebounceTimer?.Start();
                    break;

                // 处理窗口激活或失活
                case WM_ACTIVATE:
                    bool isActive = (m.WParam.ToInt64() & 0xFFFF) != 0;
                    OnActivationChanged?.Invoke(isActive);
                    break;

                // 处理视口失去键盘焦点消息 (用户切出至其他窗口)
                case WM_KILLFOCUS:
                    OnActivationChanged?.Invoke(false);
                    break;

                // 目标窗口被销毁时自动解绑
                case WM_DESTROY:
                    this.ReleaseHandle();
                    break;
            }
        }

        /// <summary>
        /// 释放当前钩子对象，解除底层句柄挂载与定时器
        /// </summary>
        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            // 停止并清理滚动定时器
            if (_scrollDebounceTimer != null)
            {
                _scrollDebounceTimer.Stop();
                _scrollDebounceTimer.Tick -= ScrollDebounceTimer_Tick;
                _scrollDebounceTimer.Dispose();
                _scrollDebounceTimer = null;
            }

            // 停止并清理尺寸移动定时器
            if (_boundsDebounceTimer != null)
            {
                _boundsDebounceTimer.Stop();
                _boundsDebounceTimer.Tick -= BoundsDebounceTimer_Tick;
                _boundsDebounceTimer.Dispose();
                _boundsDebounceTimer = null;
            }

            // 安全解除原生句柄挂载
            if (this.Handle != IntPtr.Zero)
            {
                try
                {
                    this.ReleaseHandle();
                }
                catch { }
            }
        }
    }
}
