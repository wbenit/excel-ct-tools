using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 聚光灯半透明无边框穿透浮窗，用于承载 Windows GDI Region 裁剪后的高亮交叉带
    /// 100% 鼠标点击穿透至 Excel 网格，显示重绘时不夺取 Excel 编辑焦点
    /// </summary>
    public class SpotlightOverlayForm : Form
    {
        // 扩展窗口样式常量：不在任务栏和 Alt+Tab 列表展示
        private const int WS_EX_TOOLWINDOW = 0x00000080;

        // 扩展窗口样式常量：鼠标点击时不激活窗口，保持 Excel 当前光标闪烁
        private const int WS_EX_NOACTIVATE = 0x08000000;

        // 扩展窗口样式常量：鼠标事件完全穿透到底层 Excel 视口，实现 0 阻碍编辑
        private const int WS_EX_TRANSPARENT = 0x00000020;

        // 扩展窗口样式常量：分层窗口，用于支持透明度与区域裁剪
        private const int WS_EX_LAYERED = 0x00080000;

        // 分层窗口属性常量：使用 Alpha 混合度
        private const uint LWA_ALPHA = 0x00000002;

        // P/Invoke 设置分层窗口透明度与 Alpha 属性
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

        /// <summary>
        /// 构造聚光灯穿透浮窗并完成窗口属性初始化
        /// </summary>
        public SpotlightOverlayForm()
        {
            // 设置无边框外观样式
            this.FormBorderStyle = FormBorderStyle.None;
            // 设置手动绝对屏幕坐标定位
            this.StartPosition = FormStartPosition.Manual;
            // 窗体不在任务栏展示图标
            this.ShowInTaskbar = false;
            // 绝不使用系统级全局 TopMost，避免遮挡其他第三方软件，通过 Owner 依附于 Excel
            this.TopMost = false;
            // 开启双缓冲减少重绘闪烁
            this.DoubleBuffered = true;

            // 应用当前配置的主题色与半透明度
            ApplyStyleFromConfig();
        }

        /// <summary>
        /// 重写展示无焦点激活属性，确保打开与显示时不争抢 Excel 主焦点
        /// </summary>
        protected override bool ShowWithoutActivation => true;

        /// <summary>
        /// 重写底层窗口创建参数，注入关键的 Win32 穿透与非激活扩展样式
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                // 获取基础窗体创建参数
                CreateParams cp = base.CreateParams;
                // 组合穿透、非激活、工具栏及分层扩展属性
                cp.ExStyle |= (WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TRANSPARENT | WS_EX_LAYERED);
                // 返回配置后的创建参数
                return cp;
            }
        }

        /// <summary>
        /// 原生窗口句柄创建完成回调，显式调用 Win32 API 写入分层透明度
        /// </summary>
        /// <param name="e">事件参数</param>
        protected override void OnHandleCreated(EventArgs e)
        {
            // 调用基类句柄创建逻辑
            base.OnHandleCreated(e);
            // 显式写入 Alpha 混合度，解决 Win32 默认全透明隐形问题
            SyncLayeredAttributes();
        }

        /// <summary>
        /// 从全局配置同步最新的高亮底色与不透明度
        /// </summary>
        public void ApplyStyleFromConfig()
        {
            try
            {
                // 获取全局聚光灯配置
                var cfg = SpotlightConfig.Current;
                // 解析十六进制颜色字符串
                Color parsedColor = ColorTranslator.FromHtml(cfg.ColorHex);
                // 设置窗口背景底色
                this.BackColor = parsedColor;

                // 限制不透明度安全取值范围在 0.05 ~ 0.85 之间
                double safeOpacity = Math.Max(0.05, Math.Min(0.85, cfg.Opacity));
                // 赋予窗口不透明度属性
                this.Opacity = safeOpacity;

                // 同步分层窗口 Alpha 混合值
                SyncLayeredAttributes();
            }
            catch
            {
                // 若颜色解析失败，安全回退到默认主题色 #009688
                this.BackColor = Color.FromArgb(0, 150, 136); // --硬编码: 异常回退默认颜色--
                // 回退到默认不透明度 0.22
                this.Opacity = 0.22; // --硬编码: 异常回退默认不透明度--
                // 回退同步分层属性
                SyncLayeredAttributes();
            }
        }

        /// <summary>
        /// 显式调用 Win32 API 同步当前分层窗口的透明度与色彩键
        /// </summary>
        public void SyncLayeredAttributes()
        {
            // 检查句柄是否已经创建且未释放
            if (this.IsHandleCreated && !this.IsDisposed)
            {
                try
                {
                    // 计算 0-255 字节级别的 Alpha 混合值
                    byte alphaByte = (byte)(Math.Max(0.05, Math.Min(0.85, this.Opacity)) * 255);
                    // 调用系统底层 API 写入分层窗口 Alpha
                    SetLayeredWindowAttributes(this.Handle, 0, alphaByte, LWA_ALPHA);
                }
                catch { }
            }
        }

        /// <summary>
        /// 清理并安全释放窗体资源
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            // 判断是否处于托管资源释放阶段
            if (disposing)
            {
                // 确保隐藏并释放
                this.Hide();
            }
            // 调用基类销毁流程
            base.Dispose(disposing);
        }
    }
}
