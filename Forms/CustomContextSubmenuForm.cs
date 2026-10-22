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

        // 独立高频鼠标位置巡检定时器 (周期 50ms)
        private readonly WinFormsTimer _mouseTrackerTimer;

        // 鼠标离开子菜单与触发项热区的累计检测次数
        private int _outOfHoverTicks = 0;

        // 触发呼出当前子菜单的主菜单项屏幕物理热区 (含两窗体间微小接缝走廊)
        private Rectangle _triggerBounds = Rectangle.Empty;

        // 当前悬浮菜单项索引 (-1 表示未悬浮)
        private int _hoverIndex = -1;

        // 动态点击分发委托 (为空时走默认静态逻辑)
        private Action<string>? _customSelectHandler;

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
            // 图标渲染类型: "device", "cloud", "disable", "sheet", "info"
            public string IconKind { get; set; } = "default";
        }

        // 默认价格搜索二级菜单集合
        private static readonly SubmenuItem[] _defaultPriceItems = new[]
        {
            // 1. 个人库选项
            new SubmenuItem { Title = "个人库", Shortcut = "本地", Action = "searchPricePersonal", IconKind = "device" }, // --硬编码: 菜单项标题--
            // 2. 云库选项 (下方带分割线)
            new SubmenuItem { Title = "云库", Shortcut = "云端", Action = "searchPriceCloud", HasDividerAfter = true, IconKind = "cloud" }, // --硬编码: 菜单项标题--
            // 3. 快速关闭搜索选项
            new SubmenuItem { Title = "关闭搜索", Shortcut = "停用", Action = "disablePriceSearch", IconKind = "disable" } // --硬编码: 菜单项标题--
        };

        // 当前正在展示的二级菜单配置集合
        private SubmenuItem[] _items = _defaultPriceItems;

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

            // 初始化独立高频鼠标巡检定时器 (周期 50ms)
            _mouseTrackerTimer = new WinFormsTimer { Interval = 50 }; // --硬编码: 巡检定时器周期 50ms--
            // 订阅定时器巡检事件
            _mouseTrackerTimer.Tick += OnSubmenuTrackerTick;

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
        /// 精准将默认价格搜索二级子菜单定位展示在指定屏幕坐标贴合处
        /// </summary>
        /// <param name="parent">所属主菜单窗体实例</param>
        /// <param name="screenAnchor">主菜单项右侧屏幕锚点坐标</param>
        public void ShowSubmenu(CustomContextMenuForm parent, Point screenAnchor)
        {
            // 恢复为默认价格搜索菜单配置
            _items = _defaultPriceItems;
            _customSelectHandler = null;
            // 以标准 146px 逻辑宽度展示
            ShowSubmenuInternal(parent, screenAnchor, logicWidth: 146); // --硬编码: 默认逻辑宽度 146px--
        }

        /// <summary>
        /// 动态呼出【复制当前箱柜到...】的二级分类表列表子菜单
        /// 遵循规范：每 3 行代码至少包含 1 行中文注释，动态计算自适应宽度与高度
        /// </summary>
        /// <param name="parent">所属主菜单窗体实例</param>
        /// <param name="screenAnchor">主菜单项右侧屏幕锚点坐标</param>
        /// <param name="otherCategorySheets">当前工作簿其他有效分类表名称列表</param>
        public void ShowCabinetCopySubmenu(CustomContextMenuForm parent, Point screenAnchor, System.Collections.Generic.List<string> otherCategorySheets)
        {
            // 校验分类表集合是否为空
            if (otherCategorySheets == null || otherCategorySheets.Count == 0)
            {
                // 若无其他分类表，呈现空提示项
                _items = new[]
                {
                    new SubmenuItem
                    {
                        Title = "无其他分类表", // --硬编码: 空提示文本--
                        Shortcut = "提示",     // --硬编码: 提示标签--
                        Action = "none",
                        IconKind = "info"
                    }
                };
            }
            else
            {
                // 构造各分类表对应的子菜单项
                var list = new System.Collections.Generic.List<SubmenuItem>();
                for (int i = 0; i < otherCategorySheets.Count; i++)
                {
                    string sheetName = otherCategorySheets[i];
                    list.Add(new SubmenuItem
                    {
                        Title = sheetName,
                        Shortcut = "目标表", // --硬编码: 快捷提示标签--
                        Action = $"copyToSheet:{sheetName}",
                        IconKind = "sheet",
                        HasDividerAfter = false
                    });
                }
                _items = list.ToArray();
            }

            // 自定义点击分发处理
            _customSelectHandler = (action) =>
            {
                // 若命中有效分类表复制指令
                if (action.StartsWith("copyToSheet:"))
                {
                    string targetName = action.Substring("copyToSheet:".Length);
                    // 异步派发至 Excel 调度核心引擎
                    ExcelServices.CopyActiveCabinetToSheet(targetName);
                }
            };

            // 动态测量最长标题字数自适应计算逻辑宽度 (保底 160px，上限 260px)
            int maxChars = 6;
            foreach (var it in _items)
            {
                if (it.Title.Length > maxChars) maxChars = it.Title.Length;
            }
            int dynamicWidth = Math.Max(160, Math.Min(260, 110 + maxChars * 14)); // --硬编码: 自适应宽度系数--

            // 调用内部统一定位展示方法
            ShowSubmenuInternal(parent, screenAnchor, dynamicWidth);
        }

        /// <summary>
        /// 内部通用子菜单屏幕贴合定位与展示核心逻辑
        /// </summary>
        private void ShowSubmenuInternal(CustomContextMenuForm parent, Point screenAnchor, int logicWidth)
        {
            // 绑定主菜单父引用
            _parentMenu = parent;
            // 停止正在运行的隐藏防抖定时器
            _hideTimer.Stop();
            // 重置悬浮高亮项
            _hoverIndex = -1;

            // 获取 DPI 缩放比率
            float dpiScale = GetDpiScale();
            // 计算逻辑项高与分割线高度
            int itemH = (int)Math.Ceiling(25 * dpiScale); // --硬编码: 项高度--
            int divH = (int)Math.Ceiling(6 * dpiScale);   // --硬编码: 分割线高度--
            int padY = (int)Math.Ceiling(3 * dpiScale);   // --硬编码: 上内边距--

            // 统计分割线总数
            int dividerCount = 0;
            foreach (var it in _items)
            {
                if (it.HasDividerAfter) dividerCount++;
            }

            // 动态精确计算物理宽度与物理高度
            int targetW = (int)Math.Ceiling(logicWidth * dpiScale);
            int targetH = padY * 2 + _items.Length * itemH + dividerCount * divH;
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

            // 记录主菜单中触发项在屏幕上的热区走廊 (供巡检判定防误关)
            if (parent != null)
            {
                // 触发项高度预估 30px * dpiScale
                int triggerH = (int)Math.Ceiling(30 * dpiScale); // --硬编码: 触发项高度基准--
                // 覆盖主菜单触发项水平宽度，并在两窗体接缝处延伸 10px 走廊
                int left = Math.Min(parent.Left, x - 10);
                int width = Math.Max(parent.Right, x) - left + 10;
                _triggerBounds = new Rectangle(left, screenAnchor.Y - 4, width, triggerH + 8);
            }
            else
            {
                _triggerBounds = Rectangle.Empty;
            }

            // 重置离开计数并启动独立高频鼠标巡检定时器
            _outOfHoverTicks = 0;
            _mouseTrackerTimer.Stop();
            _mouseTrackerTimer.Start();

            // 显示并置顶窗体
            if (!this.Visible)
            {
                this.Show();
            }
            // 强制重绘
            this.Invalidate();
            // 置于前台
            this.BringToFront();
        }

        /// <summary>
        /// 窗体可见性变化时重置定时器与悬浮状态
        /// </summary>
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            // 窗体隐藏时停止巡检定时器并重置所有悬浮状态
            if (!this.Visible)
            {
                _mouseTrackerTimer.Stop();
                _hideTimer.Stop();
                _outOfHoverTicks = 0;
                _hoverIndex = -1;
            }
        }

        /// <summary>
        /// 独立高频巡检鼠标屏幕绝对位置：离开子菜单一定时间后自动隐藏，杜绝孤立残留
        /// </summary>
        private void OnSubmenuTrackerTick(object? sender, EventArgs e)
        {
            try
            {
                // 1. 若当前子菜单窗体已不可见，立即停用巡检定时器
                if (!this.Visible)
                {
                    _mouseTrackerTimer.Stop();
                    _outOfHoverTicks = 0;
                    return;
                }

                // 2. 若父级主菜单已关闭或不可见，立即强制隐藏子菜单并清除高亮
                if (_parentMenu == null || !_parentMenu.Visible)
                {
                    _mouseTrackerTimer.Stop();
                    _outOfHoverTicks = 0;
                    _hoverIndex = -1;
                    this.Hide();
                    return;
                }

                // 3. 读取当前鼠标屏幕绝对全局坐标
                Point cur = Cursor.Position;

                // 4. 判定鼠标是否在子菜单窗体边界内 (给予 3 像素微小贴边容差)
                Rectangle subRectWithMargin = new Rectangle(this.Left - 3, this.Top - 3, this.Width + 6, this.Height + 6);
                bool inSubmenu = subRectWithMargin.Contains(cur);

                // 5. 判定鼠标是否在主菜单对应的触发项热区走廊内
                bool inTrigger = !_triggerBounds.IsEmpty && _triggerBounds.Contains(cur);

                if (inSubmenu)
                {
                    // 鼠标正停留在子菜单上面，重置离开计数
                    _outOfHoverTicks = 0;

                    // 转换为窗体客户区物理坐标
                    Point clientPt = this.PointToClient(cur);
                    // 动态更新当前高亮悬浮项索引
                    UpdateHoverIndexByPoint(clientPt);
                }
                else
                {
                    // 鼠标已不在子菜单上面：立即清除子菜单项的悬停灰底高亮并重绘 (解决离开后项依然残留高亮缺陷)
                    if (_hoverIndex != -1)
                    {
                        _hoverIndex = -1;
                        this.Invalidate();
                    }

                    if (inTrigger)
                    {
                        // 鼠标停留在父菜单触发项上，保持显示并重置离开计数
                        _outOfHoverTicks = 0;
                    }
                    else
                    {
                        // 鼠标既不在子菜单上面，也不在触发项上：离开计数递增
                        _outOfHoverTicks++;
                        // 离开达到阈值 (按 50ms 周期累加 8 次，约 400ms 即 0.4 秒)，平滑自动隐藏
                        // --硬编码: 离开子菜单后自动隐藏延迟 400ms--
                        if (_outOfHoverTicks >= 8)
                        {
                            _mouseTrackerTimer.Stop();
                            _outOfHoverTicks = 0;
                            _hoverIndex = -1;
                            this.Hide();
                        }
                    }
                }
            }
            catch
            {
                // 异常安全保护
            }
        }

        /// <summary>
        /// 根据客户区坐标更新当前悬浮高亮项
        /// </summary>
        private void UpdateHoverIndexByPoint(Point clientPt)
        {
            float dpiScale = GetDpiScale();
            int itemH = (int)Math.Ceiling(25 * dpiScale); // --硬编码: 项高度--
            int divH = (int)Math.Ceiling(6 * dpiScale);   // --硬编码: 分割线高度--
            int padY = (int)Math.Ceiling(3 * dpiScale);   // --硬编码: 上内边距--

            int newHover = -1;
            int curY = padY;

            // 遍历各个项检测命中区域
            for (int i = 0; i < _items.Length; i++)
            {
                if (clientPt.Y >= curY && clientPt.Y < curY + itemH && clientPt.X >= 0 && clientPt.X <= this.Width)
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
            // 停止定时器并重置离开计数
            _hideTimer.Stop();
            _outOfHoverTicks = 0;
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

            // 校验鼠标当前是否仍在子菜单或触发项走廊内部
            Point cur = Cursor.Position;
            if (this.Bounds.Contains(cur)) return;
            if (!_triggerBounds.IsEmpty && _triggerBounds.Contains(cur)) return;

            // 彻底隐藏子菜单
            this.Hide();
            _hoverIndex = -1;
        }

        /// <summary>
        /// 鼠标移动更新悬浮项索引并重绘
        /// </summary>
        private void OnSubmenuMouseMove(object? sender, MouseEventArgs e)
        {
            // 取消延迟隐藏并重置离开计数
            CancelHide();
            _outOfHoverTicks = 0;

            // 更新悬浮高亮项
            UpdateHoverIndexByPoint(e.Location);
        }

        /// <summary>
        /// 鼠标移出子菜单窗体触发延迟隐藏
        /// </summary>
        private void OnSubmenuMouseLeave(object? sender, EventArgs e)
        {
            // 立即重置悬浮索引并重绘消除灰底
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
            if (item.Action == "none") return;

            // 立即隐藏自身与主菜单
            this.Hide();
            if (_parentMenu != null && _parentMenu.Visible)
            {
                _parentMenu.Hide();
            }

            // 优先检查是否有自定义点击委托
            if (_customSelectHandler != null)
            {
                // 脱离 WinForms 线程，派发至 Excel 纯净宏队列异步调度
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        _customSelectHandler(item.Action);
                    }
                    catch (Exception exHandler)
                    {
                        LogHelper.WriteLog($"执行自定义子菜单动作 [{item.Action}] 异常: {exHandler.Message}");
                    }
                });
                return;
            }

            // 兜底默认静态搜索价格分发逻辑
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
            using var brandPen = new Pen(Color.FromArgb(0, 150, 136), 1.2f); // #009688 翠青 --硬编码--

            int curY = padY;

            // 循环绘制各项
            for (int i = 0; i < _items.Length; i++)
            {
                var it = _items[i];
                Rectangle itemRect = new Rectangle(padX, curY, this.Width - padX * 2, itemH);

                // 绘制悬浮高亮圆角背景
                if (i == _hoverIndex && it.Action != "none")
                {
                    g.FillRectangle(hoverBrush, itemRect);
                }

                // 绘制图标区域
                int iconX = padX + (int)(4 * dpiScale);
                int iconY = curY + (itemH - iconSize) / 2;

                // 针对不同图标类型精致绘制轻量矢量图形
                if (it.IconKind == "sheet")
                {
                    // 分类表图标：轻量工作表网格 (带表头翠青点缀)
                    g.DrawRectangle(brandPen, iconX, iconY, iconSize - 2, iconSize - 3);
                    g.DrawLine(brandPen, iconX, iconY + 3, iconX + iconSize - 2, iconY + 3);
                    g.DrawLine(iconPen, iconX + 4, iconY + 3, iconX + 4, iconY + iconSize - 3);
                }
                else if (it.IconKind == "device")
                {
                    // 个人库图标：微型桌面显示器
                    g.DrawRectangle(iconPen, iconX, iconY, iconSize - 2, iconSize - 5);
                    g.DrawLine(iconPen, iconX + 3, iconY + iconSize - 3, iconX + iconSize - 5, iconY + iconSize - 3);
                }
                else if (it.IconKind == "cloud")
                {
                    // 云库图标：微型云朵
                    g.DrawArc(iconPen, iconX + 1, iconY + 3, iconSize / 2, iconSize / 2, 180, 180);
                    g.DrawArc(iconPen, iconX + iconSize / 3, iconY + 1, iconSize / 2 + 1, iconSize / 2 + 1, 200, 160);
                    g.DrawLine(iconPen, iconX + 2, iconY + iconSize - 4, iconX + iconSize - 3, iconY + iconSize - 4);
                }
                else if (it.IconKind == "disable")
                {
                    // 关闭搜索图标：圆圈带斜线 (停用)
                    g.DrawEllipse(iconPen, iconX + 1, iconY + 1, iconSize - 3, iconSize - 3);
                    g.DrawLine(iconPen, iconX + 3, iconY + 3, iconX + iconSize - 5, iconY + iconSize - 5);
                }
                else if (it.IconKind == "info")
                {
                    // 信息提示图标：微型小圆圈
                    g.DrawEllipse(iconPen, iconX + 1, iconY + 1, iconSize - 3, iconSize - 3);
                    g.DrawLine(iconPen, iconX + iconSize / 2 - 1, iconY + 4, iconX + iconSize / 2 - 1, iconY + iconSize - 4);
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
