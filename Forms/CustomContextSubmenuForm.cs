using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using ExcelDna.Integration;
using WinFormsTimer = System.Windows.Forms.Timer;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 右键菜单专属独立二级子菜单窗体 (搜索价格: 个人库 / 云库 / 关闭搜索)
    /// 彻底解决单视口拓宽导致右上方露出大背景窗体与最外层边框阴影的视觉缺陷
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class CustomContextSubmenuForm : Form
    {
        // 静态单例实例引用
        private static CustomContextSubmenuForm? _instance;

        // 关联的主菜单窗体引用
        private CustomContextMenuForm? _parentMenu;

        // 延时关闭防抖定时器
        private readonly WinFormsTimer _hideTimer;

        // 当前悬浮菜单项索引 (-1 表示未悬浮)
        private int _hoverIndex = -1;

        // 菜单项数据模型定义
        private class SubmenuItem
        {
            // 菜单项标题文本
            public string Title { get; set; } = string.Empty;
            // 右侧快捷提示文本
            public string Shortcut { get; set; } = string.Empty;
            // 触发的动作指令标识
            public string Action { get; set; } = string.Empty;
            // 是否在下方绘制分割线
            public bool HasDividerAfter { get; set; }
        }

        // 固定的 3 项二级菜单配置集合
        private readonly SubmenuItem[] _items = new[]
        {
            // 1. 个人库选项
            new SubmenuItem { Title = "个人库", Shortcut = "本地", Action = "searchPricePersonal" }, // --硬编码: 菜单项标题--
            // 2. 云库选项 (下方带分割线)
            new SubmenuItem { Title = "云库", Shortcut = "云端", Action = "searchPriceCloud", HasDividerAfter = true }, // --硬编码: 菜单项标题--
            // 3. 快速关闭搜索选项
            new SubmenuItem { Title = "关闭搜索", Shortcut = "停用", Action = "disablePriceSearch" } // --硬编码: 菜单项标题--
        };

        /// <summary>
        /// 确保子菜单显示时不抢占主菜单焦点
        /// </summary>
        protected override bool ShowWithoutActivation => true;

        /// <summary>
        /// 配置窗体样式，添加无激活扩展样式 WS_EX_NOACTIVATE
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                // WS_EX_NOACTIVATE 扩展样式常数 (0x08000000) --硬编码: 无激活扩展样式--
                const int WS_EX_NOACTIVATE = 0x08000000;
                // 获取基础创建参数
                CreateParams cp = base.CreateParams;
                // 附加无激活样式，避免弹出时夺取焦点
                cp.ExStyle |= WS_EX_NOACTIVATE;
                // 返回配置后的创建参数
                return cp;
            }
        }

        /// <summary>
        /// 私有构造函数：初始化无边框置顶窗体几何属性与交互事件
        /// </summary>
        public CustomContextSubmenuForm()
        {
            // 设置无边框模式
            this.FormBorderStyle = FormBorderStyle.None;
            // 不在任务栏中展示窗体图标
            this.ShowInTaskbar = false;
            // 窗体始终置顶显示
            this.TopMost = true;
            // 启用绝对手动定位
            this.StartPosition = FormStartPosition.Manual;
            // 启用双缓冲消除鼠标移动重绘闪烁
            this.DoubleBuffered = true;
            // 默认背景色纯白
            this.BackColor = Color.White;

            // 初始化防抖关闭定时器 (180ms)
            _hideTimer = new WinFormsTimer { Interval = 180 }; // --硬编码: 防抖时间 180ms--
            // 订阅定时器超时事件
            _hideTimer.Tick += OnHideTimerTick;

            // 订阅鼠标移动事件更新高亮项
            this.MouseMove += OnSubmenuMouseMove;
            // 订阅鼠标离开事件触发延迟隐藏
            this.MouseLeave += OnSubmenuMouseLeave;
            // 订阅鼠标点击事件触发业务指令
            this.MouseClick += OnSubmenuMouseClick;
        }

        /// <summary>
        /// 获取或创建二级子菜单单例
        /// </summary>
        public static CustomContextSubmenuForm Instance
        {
            get
            {
                // 若实例不存在或已释放则重新实例化
                if (_instance == null || _instance.IsDisposed)
                {
                    _instance = new CustomContextSubmenuForm();
                }
                return _instance;
            }
        }

        /// <summary>
        /// 获取当前屏幕或窗体的 DPI 缩放比率 (以标准 96 DPI 为 1.0 基准)
        /// </summary>
        private float GetDpiScale()
        {
            try
            {
                // 获取当前窗体 GDI 句柄以读取横向物理 DPI
                using (Graphics g = this.CreateGraphics())
                {
                    // 96 DPI 对应 100% 缩放
                    float dpi = g.DpiX;
                    // 返回安全计算比率 (最低保底 1.0)
                    return dpi > 0 ? (dpi / 96.0f) : 1.0f;
                }
            }
            catch
            {
                // 异常回退标准比率 1.0
                return 1.0f;
            }
        }

        /// <summary>
        /// 精准将二级子菜单定位展示在指定屏幕坐标贴合处
        /// </summary>
        /// <param name="parent">所属主菜单窗体实例</param>
        /// <param name="screenAnchor">主菜单项右侧屏幕锚点坐标</param>
        public void ShowSubmenu(CustomContextMenuForm parent, Point screenAnchor)
        {
            // 绑定主菜单父引用
            _parentMenu = parent;
            // 停止正在运行的隐藏定时器
            _hideTimer.Stop();

            // 获取 DPI 缩放比率
            float dpiScale = GetDpiScale();
            // 计算逻辑尺寸换算后的物理尺寸 (宽 146px，高 88px)
            int targetW = (int)Math.Ceiling(146 * dpiScale); // --硬编码: 子菜单逻辑宽度 146px--
            int targetH = (int)Math.Ceiling(88 * dpiScale);  // --硬编码: 子菜单逻辑高度 88px--
            this.Size = new Size(targetW, targetH);

            // 获取当前屏幕可用工作区域
            Screen currentScreen = Screen.FromPoint(screenAnchor);
            Rectangle workArea = currentScreen.WorkingArea;

            // 默认贴在主菜单右侧
            int x = screenAnchor.X;
            int y = screenAnchor.Y;

            // 若右侧超出工作区边界，翻转至主菜单左侧
            if (x + targetW > workArea.Right && parent != null)
            {
                x = Math.Max(workArea.Left, parent.Left - targetW + 1);
            }

            // 若底部超出工作区边界，向上贴合
            if (y + targetH > workArea.Bottom)
            {
                y = Math.Max(workArea.Top, workArea.Bottom - targetH - 2);
            }

            // 设置物理坐标
            this.Location = new Point(x, y);

            // 显示并置顶窗体
            if (!this.Visible)
            {
                this.Show();
            }
            // 置于前台
            this.BringToFront();
        }

        /// <summary>
        /// 启动防抖延迟隐藏 (供主菜单项离开时调用)
        /// </summary>
        public void ScheduleHide()
        {
            // 启动 180ms 定时器
            _hideTimer.Stop();
            _hideTimer.Start();
        }

        /// <summary>
        /// 取消延迟隐藏 (供鼠标移入子菜单时调用)
        /// </summary>
        public void CancelHide()
        {
            // 停止定时器
            _hideTimer.Stop();
        }

        /// <summary>
        /// 定时器触发执行隐藏
        /// </summary>
        private void OnHideTimerTick(object? sender, EventArgs e)
        {
            // 停止定时器
            _hideTimer.Stop();
            // 若主菜单已不存在或已关闭，立即强制隐藏子菜单，彻底杜绝子菜单独自残留
            if (_parentMenu == null || !_parentMenu.Visible)
            {
                this.Hide();
                _hoverIndex = -1;
                return;
            }

            // 校验鼠标当前是否仍在子菜单或主菜单内部
            Point cur = Cursor.Position;
            if (this.Bounds.Contains(cur)) return;
            if (_parentMenu.Bounds.Contains(cur)) return;

            // 彻底隐藏子菜单
            this.Hide();
            _hoverIndex = -1;
        }

        /// <summary>
        /// 鼠标移动更新悬浮项索引并重绘
        /// </summary>
        private void OnSubmenuMouseMove(object? sender, MouseEventArgs e)
        {
            // 取消延迟隐藏
            CancelHide();

            // 计算物理项目行高与分割线高度
            float dpiScale = GetDpiScale();
            int itemH = (int)Math.Ceiling(25 * dpiScale); // --硬编码: 项高度--
            int divH = (int)Math.Ceiling(6 * dpiScale);   // --硬编码: 分割线高度--
            int padY = (int)Math.Ceiling(3 * dpiScale);   // --硬编码: 上内边距--

            int newHover = -1;
            int curY = padY;

            // 遍历各个项检测命中区域
            for (int i = 0; i < _items.Length; i++)
            {
                if (e.Y >= curY && e.Y < curY + itemH)
                {
                    newHover = i;
                    break;
                }
                curY += itemH;
                if (_items[i].HasDividerAfter)
                {
                    curY += divH;
                }
            }

            // 悬浮项发生变化时触发重绘
            if (newHover != _hoverIndex)
            {
                _hoverIndex = newHover;
                this.Invalidate();
            }
        }

        /// <summary>
        /// 鼠标移出子菜单窗体触发延迟隐藏
        /// </summary>
        private void OnSubmenuMouseLeave(object? sender, EventArgs e)
        {
            // 重置悬浮索引
            _hoverIndex = -1;
            this.Invalidate();
            // 启动延迟隐藏
            ScheduleHide();
        }

        /// <summary>
        /// 鼠标点击项分发执行业务逻辑
        /// </summary>
        private void OnSubmenuMouseClick(object? sender, MouseEventArgs e)
        {
            if (_hoverIndex < 0 || _hoverIndex >= _items.Length) return;
            var item = _items[_hoverIndex];

            // 立即隐藏自身与主菜单
            this.Hide();
            if (_parentMenu != null && _parentMenu.Visible)
            {
                _parentMenu.Hide();
            }

            // 脱离 WinForms 线程，交由 Excel 纯净宏队列异步调度业务
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    if (item.Action == "searchPricePersonal")
                    {
                        // 调度搜索个人库
                        ExcelServices.ShowPriceSearchOverlay("personal");
                    }
                    else if (item.Action == "searchPriceCloud")
                    {
                        // 调度搜索云库
                        ExcelServices.ShowPriceSearchOverlay("cloud");
                    }
                    else if (item.Action == "disablePriceSearch")
                    {
                        // 调度快速关闭搜索功能
                        ExcelServices.DisablePriceSearch();
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"执行子菜单动作 [{item.Action}] 异常: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 精致自绘子菜单外观与图标 (高对比度原生 Office 风格)
        /// </summary>
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            // 启用高质量抗锯齿渲染
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // 背景填充纯白
            g.Clear(Color.White);

            float dpiScale = GetDpiScale();
            int itemH = (int)Math.Ceiling(25 * dpiScale);
            int divH = (int)Math.Ceiling(6 * dpiScale);
            int padY = (int)Math.Ceiling(3 * dpiScale);
            int padX = (int)Math.Ceiling(4 * dpiScale);
            int iconSize = (int)Math.Ceiling(14 * dpiScale);

            // 字体定义：与一级菜单 12px 和 11px 严格一致 (Microsoft YaHei) --硬编码: 主字号 12px, 辅助字号 11px--
            using var mainFont = new Font("Microsoft YaHei", 12f * dpiScale, FontStyle.Regular, GraphicsUnit.Pixel);
            // 辅助快捷键提示字体
            using var subFont = new Font("Microsoft YaHei", 11f * dpiScale, FontStyle.Regular, GraphicsUnit.Pixel);

            // 调色盘画刷与画笔
            using var borderPen = new Pen(Color.FromArgb(212, 212, 212), 1f); // #d4d4d4 --硬编码--
            using var dividerPen = new Pen(Color.FromArgb(229, 229, 229), 1f); // #e5e5e5 --硬编码--
            using var hoverBrush = new SolidBrush(Color.FromArgb(240, 240, 240)); // #f0f0f0 --硬编码--
            using var textMainBrush = new SolidBrush(Color.FromArgb(38, 38, 38)); // #262626 --硬编码--
            using var textMutedBrush = new SolidBrush(Color.FromArgb(115, 115, 115)); // #737373 --硬编码--
            using var iconPen = new Pen(Color.FromArgb(74, 74, 74), 1.2f); // #4a4a4a --硬编码--

            int curY = padY;

            // 循环绘制各项
            for (int i = 0; i < _items.Length; i++)
            {
                var it = _items[i];
                Rectangle itemRect = new Rectangle(padX, curY, this.Width - padX * 2, itemH);

                // 绘制悬浮高亮圆角背景
                if (i == _hoverIndex)
                {
                    g.FillRectangle(hoverBrush, itemRect);
                }

                // 绘制图标区域
                int iconX = padX + (int)(4 * dpiScale);
                int iconY = curY + (itemH - iconSize) / 2;

                // 针对不同项绘制极简精致矢量符号
                if (i == 0)
                {
                    // 个人库图标：微型桌面显示器
                    g.DrawRectangle(iconPen, iconX, iconY, iconSize - 2, iconSize - 5);
                    g.DrawLine(iconPen, iconX + 3, iconY + iconSize - 3, iconX + iconSize - 5, iconY + iconSize - 3);
                }
                else if (i == 1)
                {
                    // 云库图标：微型云朵
                    g.DrawArc(iconPen, iconX + 1, iconY + 3, iconSize / 2, iconSize / 2, 180, 180);
                    g.DrawArc(iconPen, iconX + iconSize / 3, iconY + 1, iconSize / 2 + 1, iconSize / 2 + 1, 200, 160);
                    g.DrawLine(iconPen, iconX + 2, iconY + iconSize - 4, iconX + iconSize - 3, iconY + iconSize - 4);
                }
                else
                {
                    // 关闭搜索图标：圆圈带斜线 (停用)
                    g.DrawEllipse(iconPen, iconX + 1, iconY + 1, iconSize - 3, iconSize - 3);
                    g.DrawLine(iconPen, iconX + 3, iconY + 3, iconX + iconSize - 5, iconY + iconSize - 5);
                }

                // 绘制标题文字
                int textX = iconX + iconSize + (int)(6 * dpiScale);
                int textY = curY + (itemH - (int)g.MeasureString(it.Title, mainFont).Height) / 2;
                g.DrawString(it.Title, mainFont, textMainBrush, textX, textY);

                // 绘制右侧快捷提示文字
                if (!string.IsNullOrEmpty(it.Shortcut))
                {
                    SizeF scSize = g.MeasureString(it.Shortcut, subFont);
                    int scX = this.Width - padX - (int)scSize.Width - (int)(4 * dpiScale);
                    int scY = curY + (itemH - (int)scSize.Height) / 2;
                    g.DrawString(it.Shortcut, subFont, textMutedBrush, scX, scY);
                }

                curY += itemH;

                // 绘制分割线
                if (it.HasDividerAfter)
                {
                    int lineY = curY + divH / 2;
                    g.DrawLine(dividerPen, padX + 4, lineY, this.Width - padX - 4, lineY);
                    curY += divH;
                }
            }

            // 绘制最外层 1px 细浅灰边框
            g.DrawRectangle(borderPen, 0, 0, this.Width - 1, this.Height - 1);
        }
    }
}
