using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 通用迷你悬浮条窗口辅助器：
    /// 1. 隐藏任务栏独立图标，依附 Excel 主窗口
    /// 2. 最小化为 160x30px 悬浮条（默认停靠 Excel 左下角，支持自由拖动放置与相对位置记忆）
    /// 3. 自动随动 Excel 移动/缩放/最小化，支持一键还原与关闭
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public sealed class MiniWindowHelper : IDisposable
    {
        // Windows 原生 user32.dll 标题栏拖拽与状态检测接口
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        // 常量定义: 标题栏拖拽消息
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 迷你悬浮条尺寸与默认边距常量 --硬编码: 最小化悬浮条尺寸 160x30px 及左下角边距--
        private const int MiniWidth = 160;
        private const int MiniHeight = 30;
        private const int MiniMarginLeft = 10;
        private const int MiniMarginBottom = 28;
        private const int FollowTimerIntervalMs = 150;

        // 宿主 WinForm 与 WebView2 引用
        private readonly Form _form;
        private readonly WebView2 _webView;
        private readonly bool _useInjectedOverlay;

        // 窗口模式与原始状态缓存
        private bool _isMiniMode;
        private Rectangle _savedNormalBounds;
        private FormBorderStyle _savedBorderStyle;
        private Size _savedMinimumSize;
        private FormWindowState _savedWindowState;

        // 迷你条随动定时器与用户自定义拖放偏移记录
        private System.Windows.Forms.Timer? _excelFollowTimer;
        private bool _isDraggingMini;
        private bool _hasCustomMiniOffset;
        private Point _miniOffsetFromExcel;
        private bool _disposed;

        /// <summary>
        /// 当前是否处于迷你悬浮条模式
        /// </summary>
        public bool IsMiniMode => _isMiniMode;

        /// <summary>
        /// 初始化迷你悬浮条辅助器并绑定宿主窗体与 WebView2 事件
        /// </summary>
        /// <param name="form">宿主 WinForm 窗体</param>
        /// <param name="webView">宿主内的 WebView2 控件</param>
        /// <param name="useInjectedOverlay">是否由辅助器注入通用迷你条 UI（若前端已自绘可传 false）</param>
        public MiniWindowHelper(Form form, WebView2 webView, bool useInjectedOverlay = true)
        {
            _form = form ?? throw new ArgumentNullException(nameof(form));
            _webView = webView ?? throw new ArgumentNullException(nameof(webView));
            _useInjectedOverlay = useInjectedOverlay;

            // 全程隐藏任务栏独立图标并禁用原生最小化框
            _form.ShowInTaskbar = false;
            _form.MinimizeBox = false;
            // 窗体关闭时自动释放定时器等资源
            _form.FormClosed += (_, __) => Dispose();

            // 监听 WebView2 内核就绪以绑定通用迷你条交互消息
            _webView.CoreWebView2InitializationCompleted += OnCoreWebView2InitializationCompleted;
            if (_webView.CoreWebView2 != null)
            {
                _webView.CoreWebView2.WebMessageReceived += OnMiniWebMessageReceived;
            }
        }

        /// <summary>
        /// WebView2 初始化完成后挂载迷你条内部消息监听
        /// </summary>
        private void OnCoreWebView2InitializationCompleted(object? sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            if (e.IsSuccess && _webView.CoreWebView2 != null)
            {
                // 防止重复挂载，先解绑再绑定
                _webView.CoreWebView2.WebMessageReceived -= OnMiniWebMessageReceived;
                _webView.CoreWebView2.WebMessageReceived += OnMiniWebMessageReceived;
            }
        }

        /// <summary>
        /// 拦截并处理通用迷你条发出的拖拽、还原与关闭消息
        /// </summary>
        private void OnMiniWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string raw = e.TryGetWebMessageAsString();
                if (string.IsNullOrWhiteSpace(raw)) return;

                using var doc = JsonDocument.Parse(raw);
                if (!doc.RootElement.TryGetProperty("action", out var actionProp)) return;
                string? action = actionProp.GetString();

                // 仅处理迷你悬浮条专用内部指令，避免与各业务窗体已有指令冲突
                if (action == "__agyMiniDrag")
                {
                    HandleDragWindow();
                }
                else if (action == "__agyMiniRestore")
                {
                    ExitMiniMode();
                }
                else if (action == "__agyMiniClose")
                {
                    RunOnUiThread(() => _form.Close());
                }
            }
            catch
            {
                // 非 JSON 或无关消息交由业务窗体自身监听器处理
            }
        }

        /// <summary>
        /// 进入 160x30px 迷你悬浮条模式，默认停靠 Excel 左下角或记忆的拖放位置
        /// </summary>
        public void EnterMiniMode()
        {
            RunOnUiThread(() =>
            {
                // 已释放或已处于迷你模式时直接返回
                if (_disposed || _form.IsDisposed || _isMiniMode) return;

                // 缓存正常窗口矩形、边框样式、窗口状态与最小尺寸限制
                _savedWindowState = _form.WindowState;
                _savedNormalBounds = (_form.WindowState == FormWindowState.Normal) ? _form.Bounds : _form.RestoreBounds;
                _savedBorderStyle = _form.FormBorderStyle;
                _savedMinimumSize = _form.MinimumSize;
                _isMiniMode = true;

                // 若当前处于最大化状态，先切回 Normal 再缩放为 160x30px
                if (_form.WindowState != FormWindowState.Normal)
                {
                    _form.WindowState = FormWindowState.Normal;
                }
                // 切换为无边框并解除最小尺寸限制，确保可缩至 160x30px
                _form.FormBorderStyle = FormBorderStyle.None;
                _form.MinimumSize = Size.Empty;
                _form.Size = new Size(MiniWidth, MiniHeight);

                // 立即定位到 Excel 左下角或用户上次拖放的相对位置
                UpdateMiniWindowPosition();
                // 启动 Excel 随动定时器
                EnsureFollowTimerStarted();
                // 通知前端切换为迷你悬浮条界面
                ApplyFrontendMiniState(true);
            });
        }

        /// <summary>
        /// 退出迷你悬浮条模式，恢复最小化前的窗口尺寸、边框与位置
        /// </summary>
        public void ExitMiniMode()
        {
            RunOnUiThread(() =>
            {
                // 已释放或非迷你模式时直接返回
                if (_disposed || _form.IsDisposed || !_isMiniMode) return;

                // 停止随动定时器
                _excelFollowTimer?.Stop();
                _isMiniMode = false;
                if (!_form.Visible) _form.Visible = true;

                // 还原边框样式、最小尺寸约束与原窗口状态
                _form.FormBorderStyle = _savedBorderStyle;
                _form.MinimumSize = _savedMinimumSize;
                if (_savedNormalBounds.Width > MiniWidth && _savedNormalBounds.Height > MiniHeight)
                {
                    _form.Bounds = _savedNormalBounds;
                }
                if (_savedWindowState == FormWindowState.Maximized)
                {
                    _form.WindowState = FormWindowState.Maximized;
                }

                // 通知前端恢复完整业务界面
                ApplyFrontendMiniState(false);
            });
        }

        /// <summary>
        /// 处理窗口拖拽；若当前处于迷你模式则暂停归位并在松手后记忆相对 Excel 的偏移量
        /// </summary>
        public void HandleDragWindow()
        {
            RunOnUiThread(() =>
            {
                if (_disposed || _form.IsDisposed) return;

                // 迷你模式下拖拽需暂停定时器归位，松手后更新相对偏移量
                if (_isMiniMode)
                {
                    _isDraggingMini = true;
                    try
                    {
                        ReleaseCapture();
                        SendMessage(_form.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                    }
                    finally
                    {
                        // 记录松手后相对于 Excel 主窗口左上角的自定义偏移
                        IntPtr excelHwnd = ExcelDnaSafeAccessor.GetWindowHandle();
                        if (excelHwnd != IntPtr.Zero && Win32Interop.GetWindowRect(excelHwnd, out Win32Interop.RECT rect))
                        {
                            _miniOffsetFromExcel = new Point(_form.Location.X - rect.Left, _form.Location.Y - rect.Top);
                            _hasCustomMiniOffset = true;
                        }
                        _isDraggingMini = false;
                    }
                    return;
                }

                // 正常模式下直接执行原生拖拽
                ReleaseCapture();
                SendMessage(_form.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
            });
        }

        /// <summary>
        /// 启动或复用 Excel 主窗口随动定时器
        /// </summary>
        private void EnsureFollowTimerStarted()
        {
            if (_excelFollowTimer == null)
            {
                _excelFollowTimer = new System.Windows.Forms.Timer { Interval = FollowTimerIntervalMs };
                // 定时校验并同步迷你条与 Excel 主窗口的相对位置
                _excelFollowTimer.Tick += (_, __) => UpdateMiniWindowPosition();
            }
            _excelFollowTimer.Start();
        }

        /// <summary>
        /// 根据 Excel 主窗口矩形更新迷你悬浮条位置，并在 Excel 最小化时同步隐藏
        /// </summary>
        private void UpdateMiniWindowPosition()
        {
            if (_disposed || !_isMiniMode || _isDraggingMini || _form.IsDisposed || _form.Disposing) return;

            IntPtr excelHwnd = ExcelDnaSafeAccessor.GetWindowHandle();
            if (excelHwnd == IntPtr.Zero) return;

            // Excel 最小化时同步隐藏悬浮条，避免孤立悬浮在桌面
            if (IsIconic(excelHwnd))
            {
                if (_form.Visible) _form.Visible = false;
                return;
            }

            if (!_form.Visible) _form.Visible = true;

            if (Win32Interop.GetWindowRect(excelHwnd, out Win32Interop.RECT rect))
            {
                // 获取 Excel 所在屏幕的工作区边界，防止越出任务栏或屏幕边缘
                Rectangle workingArea = Screen.FromHandle(excelHwnd).WorkingArea;
                int targetX;
                int targetY;

                // 若用户已拖放过位置，则按记忆的相对偏移量随动并钳位在工作区内
                if (_hasCustomMiniOffset)
                {
                    targetX = Math.Max(workingArea.Left, Math.Min(workingArea.Right - _form.Width, rect.Left + _miniOffsetFromExcel.X));
                    targetY = Math.Max(workingArea.Top, Math.Min(workingArea.Bottom - _form.Height, rect.Top + _miniOffsetFromExcel.Y));
                }
                else
                {
                    // 默认停靠在 Excel 左下角并留出状态栏边距
                    targetX = Math.Max(workingArea.Left + MiniMarginLeft, rect.Left + MiniMarginLeft);
                    targetY = Math.Min(workingArea.Bottom - _form.Height - MiniMarginBottom, rect.Bottom - _form.Height - MiniMarginBottom);
                }

                // 仅坐标变化时才更新 Location，避免无谓重绘闪烁
                if (_form.Location.X != targetX || _form.Location.Y != targetY)
                {
                    _form.Location = new Point(targetX, targetY);
                }
            }
        }

        /// <summary>
        /// 切换前端迷你悬浮条视图（自绘页面发消息，通用页面直接注入悬浮条覆盖层）
        /// </summary>
        private void ApplyFrontendMiniState(bool isMinimized)
        {
            if (_webView?.CoreWebView2 == null) return;

            if (!_useInjectedOverlay)
            {
                // 自绘页面通过标准 windowModeChanged 消息同步状态
                string json = JsonSerializer.Serialize(new { action = "windowModeChanged", isMinimized });
                _webView.CoreWebView2.PostWebMessageAsString(json);
                return;
            }

            // 通用页面通过注入独立悬浮条实现 160x30px 无滚动条迷你视图，不破坏原页面 DOM 与状态
            string safeTitle = JsonSerializer.Serialize(string.IsNullOrWhiteSpace(_form.Text) ? "工具窗口" : _form.Text);
            string script = $@"
(function() {{
  var isMin = {(isMinimized ? "true" : "false")};
  var root = document.documentElement;
  if (isMin) {{
    root.classList.add('__agy-mini-active__');
    if (!document.getElementById('__agy_mini_style__')) {{
      var st = document.createElement('style');
      st.id = '__agy_mini_style__';
      st.textContent =
        'html.__agy-mini-active__, html.__agy-mini-active__ body {{ overflow: hidden !important; width: 100vw !important; height: 100vh !important; margin: 0 !important; padding: 0 !important; }}' +
        '#__agy_mini_bar__ {{ position: fixed; inset: 0; z-index: 2147483647; display: flex; align-items: center; justify-content: space-between; padding: 0 8px; background: linear-gradient(135deg, #009688 0%, #00796b 100%); color: #fff; font-family: -apple-system, BlinkMacSystemFont, ""Microsoft YaHei"", sans-serif; font-size: 12px; font-weight: 600; cursor: move; user-select: none; box-sizing: border-box; border: 1px solid rgba(0, 121, 107, 0.9); border-radius: 4px; }}' +
        '#__agy_mini_bar__ .agy-mini-title {{ flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; margin-right: 6px; }}' +
        '#__agy_mini_bar__ .agy-mini-actions {{ display: flex; align-items: center; gap: 4px; flex-shrink: 0; }}' +
        '#__agy_mini_bar__ .agy-mini-btn {{ width: 20px; height: 20px; border: none; background: transparent; color: #fff; opacity: 0.9; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; border-radius: 3px; padding: 0; font-size: 12px; }}' +
        '#__agy_mini_bar__ .agy-mini-btn:hover {{ opacity: 1; background: rgba(255,255,255,0.18); }}' +
        '#__agy_mini_bar__ .agy-mini-btn.close:hover {{ color: #ff5252; background: rgba(255,255,255,0.12); }}';
      document.head.appendChild(st);
    }}
    var bar = document.getElementById('__agy_mini_bar__');
    if (!bar) {{
      bar = document.createElement('div');
      bar.id = '__agy_mini_bar__';
      bar.innerHTML =
        '<span class=""agy-mini-title""></span>' +
        '<div class=""agy-mini-actions"">' +
        '  <button type=""button"" class=""agy-mini-btn restore"" title=""还原窗口""><svg width=""11"" height=""11"" viewBox=""0 0 12 12"" fill=""none"" stroke=""currentColor"" stroke-width=""1.4""><rect x=""1.5"" y=""3.5"" width=""7"" height=""7"" rx=""1""/><path d=""M3.5 3.5V2a.5.5 0 0 1 .5-.5h6a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-.5.5H8.5""/></svg></button>' +
        '  <button type=""button"" class=""agy-mini-btn close"" title=""关闭""><svg width=""11"" height=""11"" viewBox=""0 0 12 12"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""><path d=""M2.5 2.5l7 7M9.5 2.5l-7 7""/></svg></button>' +
        '</div>';
      bar.addEventListener('mousedown', function(ev) {{
        if (ev.target && ev.target.closest('.agy-mini-btn')) return;
        window.chrome.webview.postMessage(JSON.stringify({{ action: '__agyMiniDrag' }}));
      }});
      bar.addEventListener('dblclick', function(ev) {{
        if (ev.target && ev.target.closest('.agy-mini-btn')) return;
        window.chrome.webview.postMessage(JSON.stringify({{ action: '__agyMiniRestore' }}));
      }});
      bar.querySelector('.restore').addEventListener('click', function() {{
        window.chrome.webview.postMessage(JSON.stringify({{ action: '__agyMiniRestore' }}));
      }});
      bar.querySelector('.close').addEventListener('click', function() {{
        window.chrome.webview.postMessage(JSON.stringify({{ action: '__agyMiniClose' }}));
      }});
      document.body.appendChild(bar);
    }}
    bar.querySelector('.agy-mini-title').textContent = {safeTitle};
    bar.style.display = 'flex';
  }} else {{
    root.classList.remove('__agy-mini-active__');
    var bar = document.getElementById('__agy_mini_bar__');
    if (bar) bar.style.display = 'none';
  }}
}})();";
            _webView.CoreWebView2.ExecuteScriptAsync(script);
        }

        /// <summary>
        /// 统一派发到宿主 UI 线程执行
        /// </summary>
        private void RunOnUiThread(Action action)
        {
            if (_form.IsDisposed) return;
            if (_form.InvokeRequired)
            {
                _form.BeginInvoke(action);
            }
            else
            {
                action();
            }
        }

        /// <summary>
        /// 释放定时器并解绑事件
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _excelFollowTimer?.Stop();
            _excelFollowTimer?.Dispose();
            _excelFollowTimer = null;
        }
    }
}
