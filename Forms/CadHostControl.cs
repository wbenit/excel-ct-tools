using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// AutoCAD 嵌入式任务窗格宿主控件（WinForms UserControl）
    /// 承载 AutoCAD 主窗口句柄，并提供顶栏状态监控、重新绑定与还原至桌面交互
    /// </summary>
    public class CadHostControl : UserControl
    {
        // 顶部工具栏容器
        private Panel _headerPanel = null!;

        // 标题与状态标签
        private Label _lblTitle = null!;

        // 重新绑定按钮
        private Button _btnRebind = null!;

        // 还原至桌面独立窗口按钮
        private Button _btnDetach = null!;

        // 夹点联动切换按钮
        private Button _btnToggleSync = null!;

        // 纯画布模式切换按钮（控制工具栏与 Ribbon 功能区显隐）
        private Button _btnToggleCleanScreen = null!;

        // 实际承载 AutoCAD 窗口的物理 Panel
        private Panel _cadContainerPanel = null!;

        // 空状态提示面板（未嵌入 CAD 时展示）
        private Panel _emptyPromptPanel = null!;

        // 空状态主提示文字
        private Label _lblEmptyPrompt = null!;

        // 空状态副提示文字
        private Label _lblEmptySubPrompt = null!;

        // 空状态一键绑定按钮
        private Button _btnEmptyBind = null!;

        // 主色调：绿蓝相间 --硬编码: 主色调 #009688--
        private readonly Color _primaryColor = ColorTranslator.FromHtml("#009688");

        // 顶栏深色渐变底色 --硬编码: 顶栏背景色--
        private readonly Color _headerBgColor = ColorTranslator.FromHtml("#00796B");

        // 画布空状态背景深色 --硬编码: CAD工作区深灰色背景--
        private readonly Color _canvasBgColor = ColorTranslator.FromHtml("#212121");

        /// <summary>
        /// 构造函数：初始化界面布局与子控件
        /// </summary>
        public CadHostControl()
        {
            // 开启双缓冲以减少重绘闪烁
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint, true);
            UpdateStyles();

            // 构建界面元素
            InitializeComponent();
        }

        /// <summary>
        /// 初始化界面元素与弹性自适应布局
        /// </summary>
        private void InitializeComponent()
        {
            // 设置主控件基础属性
            this.BackColor = _canvasBgColor;
            this.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            this.Size = new Size(680, 800);

            // 1. 创建顶部工具栏
            _headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = _headerBgColor,
                Padding = new Padding(8, 4, 8, 4)
            };

            // 标题与状态指示标签
            _lblTitle = new Label
            {
                Text = "⚡ AutoCAD 协同画图视口 [未连接]",
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };

            // 重新绑定按钮
            _btnRebind = CreateHeaderButton("⟳ 重新绑定", (s, e) =>
            {
                // 触发重新绑定当前活动的 AutoCAD
                CadEmbedManager.EmbedActiveCad();
            });

            // 还原至桌面按钮
            _btnDetach = CreateHeaderButton("↗ 还原至桌面", (s, e) =>
            {
                // 触发将 AutoCAD 归还至独立桌面窗口
                CadEmbedManager.DetachCad();
            });

            // 联动开关按钮
            _btnToggleSync = CreateHeaderButton("⚡ 联动: 开", (s, e) =>
            {
                // 切换夹点联动开关
                CadSyncClient.SyncToCadEnabled = !CadSyncClient.SyncToCadEnabled;
                UpdateSyncButtonState();
            });

            // 纯画布模式切换按钮（隐藏/显示工具栏与Ribbon）
            _btnToggleCleanScreen = CreateHeaderButton("✦ 纯画布: 开", (s, e) =>
            {
                // 切换纯画布模式并同步刷新视口裁剪
                CadEmbedManager.IsCleanScreenEnabled = !CadEmbedManager.IsCleanScreenEnabled;
                CadEmbedManager.SetCadCleanScreen(CadEmbedManager.IsCleanScreenEnabled);
                UpdateCleanScreenButtonState();
            });

            // 添加按钮到顶栏（Dock 为 Right，逆序添加）
            _headerPanel.Controls.Add(_lblTitle);
            _headerPanel.Controls.Add(_btnToggleSync);
            _headerPanel.Controls.Add(_btnToggleCleanScreen);
            _headerPanel.Controls.Add(_btnDetach);
            _headerPanel.Controls.Add(_btnRebind);

            // 2. 创建 AutoCAD 物理宿主容器 Panel
            _cadContainerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = _canvasBgColor
            };

            // 监听容器大小变更，实时调整嵌入的 AutoCAD 视口尺寸
            _cadContainerPanel.Resize += (s, e) =>
            {
                // 通知管理器按新尺寸重绘 AutoCAD 窗口
                CadEmbedManager.SyncCadSize();
            };

            // 监听鼠标移入事件，自动将键盘焦点赋给 AutoCAD，保障画图快捷键畅通
            _cadContainerPanel.MouseEnter += (s, e) =>
            {
                // 切换焦点至 CAD 句柄
                CadEmbedManager.FocusCad();
            };

            // 3. 创建未嵌入时的空状态引导界面
            _emptyPromptPanel = new Panel
            {
                Size = new Size(420, 240),
                BackColor = ColorTranslator.FromHtml("#2C2C2C"),
                BorderStyle = BorderStyle.None
            };

            // 居中放置空状态面板
            _cadContainerPanel.Controls.Add(_emptyPromptPanel);
            _cadContainerPanel.Resize += (s, e) => CenterEmptyPrompt();

            // 空状态主标题
            _lblEmptyPrompt = new Label
            {
                Text = "📐 暂未嵌入 AutoCAD 图纸",
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 12F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 45
            };

            // 空状态辅助说明文字
            _lblEmptySubPrompt = new Label
            {
                Text = "请在电脑上先启动并打开 AutoCAD 图纸，\n点击下方按钮即可将图纸窗口无缝嵌入本面板，\n在 Excel 中边看清单边画图，支持点击元件高亮定位。",
                ForeColor = Color.LightGray,
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 85
            };

            // 空状态主操作按钮
            _btnEmptyBind = new Button
            {
                Text = "🔍 立即检测并嵌入当前活动 AutoCAD",
                BackColor = _primaryColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
                Height = 42,
                Cursor = Cursors.Hand,
                Dock = DockStyle.Bottom
            };
            _btnEmptyBind.FlatAppearance.BorderSize = 0;
            _btnEmptyBind.Click += (s, e) =>
            {
                // 触发主动检测并嵌入
                CadEmbedManager.EmbedActiveCad();
            };

            _emptyPromptPanel.Controls.Add(_btnEmptyBind);
            _emptyPromptPanel.Controls.Add(_lblEmptySubPrompt);
            _emptyPromptPanel.Controls.Add(_lblEmptyPrompt);

            // 装配顶层控件结构
            this.Controls.Add(_cadContainerPanel);
            this.Controls.Add(_headerPanel);

            // 初始状态更新
            UpdateSyncButtonState();
            CenterEmptyPrompt();
        }

        /// <summary>
        /// 创建统一风格的顶栏操作按钮
        /// </summary>
        private Button CreateHeaderButton(string text, EventHandler onClick)
        {
            var btn = new Button
            {
                Text = text,
                Dock = DockStyle.Right,
                AutoSize = true,
                BackColor = ColorTranslator.FromHtml("#004D40"),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Regular),
                Margin = new Padding(4, 2, 4, 2),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += onClick;
            return btn;
        }

        /// <summary>
        /// 居中计算空状态面板坐标
        /// </summary>
        private void CenterEmptyPrompt()
        {
            if (_emptyPromptPanel != null && _cadContainerPanel != null)
            {
                // 计算水平与垂直居中偏移量
                int x = Math.Max(10, (_cadContainerPanel.Width - _emptyPromptPanel.Width) / 2);
                int y = Math.Max(20, (_cadContainerPanel.Height - _emptyPromptPanel.Height) / 2);
                _emptyPromptPanel.Location = new Point(x, y);
            }
        }

        /// <summary>
        /// 获取 AutoCAD 实际承载面板的 Windows 物理句柄
        /// </summary>
        public IntPtr GetContainerHandle()
        {
            // 确保控件句柄已创建
            if (!_cadContainerPanel.IsHandleCreated)
            {
                _cadContainerPanel.CreateControl();
            }
            return _cadContainerPanel.Handle;
        }

        /// <summary>
        /// 获取承载容器的实际物理尺寸
        /// </summary>
        public Size GetContainerSize()
        {
            return _cadContainerPanel.ClientSize;
        }

        /// <summary>
        /// 更新嵌入连接状态呈现（标题栏文本与空状态面板切换）
        /// </summary>
        /// <param name="isConnected">是否已成功嵌入</param>
        /// <param name="docName">当前绑定的图纸名称</param>
        public void UpdateConnectionState(bool isConnected, string? docName = null)
        {
            if (this.InvokeRequired)
            {
                // 切换至 UI 线程更新
                this.BeginInvoke(new Action(() => UpdateConnectionState(isConnected, docName)));
                return;
            }

            if (isConnected)
            {
                // 隐藏空状态提示，显示 CAD 视口
                _emptyPromptPanel.Visible = false;
                string displayTitle = !string.IsNullOrWhiteSpace(docName) ? docName : "活动图纸";
                _lblTitle.Text = $"⚡ AutoCAD 协同画图视口 [已连接: {displayTitle}]";
                _lblTitle.ForeColor = Color.LightGreen;
                _btnDetach.Enabled = true;
                _btnToggleCleanScreen.Enabled = true;
                UpdateCleanScreenButtonState();
            }
            else
            {
                // 切换回空状态引导界面
                _emptyPromptPanel.Visible = true;
                _lblTitle.Text = "⚡ AutoCAD 协同画图视口 [未连接]";
                _lblTitle.ForeColor = Color.White;
                _btnDetach.Enabled = false;
                _btnToggleCleanScreen.Enabled = false;
                CenterEmptyPrompt();
            }
        }

        /// <summary>
        /// 更新夹点联动按钮状态文案
        /// </summary>
        private void UpdateSyncButtonState()
        {
            if (CadSyncClient.SyncToCadEnabled)
            {
                _btnToggleSync.Text = "⚡ 联动: 开";
                _btnToggleSync.BackColor = _primaryColor;
            }
            else
            {
                _btnToggleSync.Text = "⏸ 联动: 关";
                _btnToggleSync.BackColor = ColorTranslator.FromHtml("#546E7A");
            }
        }

        /// <summary>
        /// 更新纯画布模式按钮状态文案与背景色
        /// </summary>
        private void UpdateCleanScreenButtonState()
        {
            if (CadEmbedManager.IsCleanScreenEnabled)
            {
                _btnToggleCleanScreen.Text = "✦ 纯画布: 开";
                _btnToggleCleanScreen.BackColor = _primaryColor;
            }
            else
            {
                _btnToggleCleanScreen.Text = "✦ 纯画布: 关";
                _btnToggleCleanScreen.BackColor = ColorTranslator.FromHtml("#546E7A");
            }
        }
    }
}
