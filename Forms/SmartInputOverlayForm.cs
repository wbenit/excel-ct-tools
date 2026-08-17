using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 方案 B: 单元格原生覆盖式智能输入窗体 (包含覆盖 TextBox 与下拉 ListBox)
    /// 100% 还原 ZhiNengEn.cs 的交互模型与属性联动带入逻辑 (规则 6 & 规则 7)
    /// </summary>
    public class SmartInputOverlayForm : Form
    {
        // 覆盖在单元格正上方的原生输入框控件
        private readonly TextBox _textbox;

        // 位于单元格正下方的原生候选列表框控件
        private readonly ListBox _listbox;

        // 当前正在编辑的 Excel 单元格 COM 引用
        private dynamic? _currentTargetCell;

        // 当前所有可供模糊检索的物料数据集
        private List<SmartComponentItem> _allComponents = new List<SmartComponentItem>();

        // 当前物料型号与完整属性映射字典 (用于快速联动带出属性)
        private Dictionary<string, SmartComponentItem> _componentDict = new Dictionary<string, SmartComponentItem>(StringComparer.OrdinalIgnoreCase);

        // 当前智能输入配置对象
        private SmartInputConfigModel _config = new SmartInputConfigModel();

        // 标记当前单元格屏幕物理高度
        private int _cellHeight = 22;

        // 标记当前单元格屏幕物理宽度
        private int _cellWidth = 120;

        // 标记当前单元格屏幕物理 X 坐标
        private int _screenX = 0;

        // 标记当前单元格屏幕物理 Y 坐标
        private int _screenY = 0;

        // 标记是否正在执行内部文本更新，避免递归触发 TextChanged
        private bool _isInternalUpdating = false;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        static SmartInputOverlayForm()
        {
            try
            {
                // 开启系统级 DPI 感知，防止在 125%/150% 缩放下坐标被虚拟化导致偏向左上角
                SetProcessDPIAware();
            }
            catch { }
        }

        /// <summary>
        /// 构造函数: 初始化无边框覆盖窗体及 TextBox 和 ListBox 控件
        /// </summary>
        public SmartInputOverlayForm()
        {
            // 初始化窗体基本外观属性
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.Manual;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.BackColor = Color.White;
            this.DoubleBuffered = true;

            // 1. 初始化覆盖 TextBox 控件
            _textbox = new TextBox
            {
                Multiline = true,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei", 9.5f, FontStyle.Regular),
                Margin = new Padding(0),
                BackColor = Color.White,
                ForeColor = Color.Black
            };

            // 绑定 TextBox 核心交互事件
            _textbox.TextChanged += Textbox_TextChanged;
            _textbox.KeyDown += Textbox_KeyDown;

            // 2. 初始化下拉 ListBox 控件
            _listbox = new ListBox
            {
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei", 9.5f, FontStyle.Regular),
                Margin = new Padding(0),
                Visible = false,
                ItemHeight = 22,
                DrawMode = DrawMode.OwnerDrawFixed,
                BackColor = Color.White
            };

            // 绑定 ListBox 绘制与交互事件
            _listbox.DrawItem += Listbox_DrawItem;
            _listbox.KeyDown += Listbox_KeyDown;
            _listbox.Click += Listbox_Click;
            _listbox.DoubleClick += Listbox_DoubleClick;

            // 将控件添加到窗体控件集合
            this.Controls.Add(_listbox);
            this.Controls.Add(_textbox);

            // 监听窗体失活事件，确保光标离开时自动提交并隐藏
            this.Deactivate += SmartInputOverlayForm_Deactivate;
        }

        /// <summary>
        /// 计算单元格在主屏幕上的真实绝对物理矩形 (兼顾 ActivePane 与 Windows 缩放)
        /// </summary>
        private Rectangle GetCellScreenRectangle(dynamic targetCell)
        {
            try
            {
                dynamic app = targetCell.Application;
                dynamic win = app.ActiveWindow;
                dynamic pane = win.ActivePane;

                double cellLeft = Convert.ToDouble(targetCell.Left);
                double cellTop = Convert.ToDouble(targetCell.Top);
                double cellWidth = Convert.ToDouble(targetCell.Width);
                double cellHeight = Convert.ToDouble(targetCell.Height);

                // 方案 1: 优先通过 ActivePane.PointsToScreenPixels 进行高精度换算
                if (pane != null)
                {
                    try
                    {
                        int px1 = pane.PointsToScreenPixelsX((int)cellLeft);
                        int py1 = pane.PointsToScreenPixelsY((int)cellTop);
                        int px2 = pane.PointsToScreenPixelsX((int)(cellLeft + cellWidth));
                        int py2 = pane.PointsToScreenPixelsY((int)(cellTop + cellHeight));

                        if (px1 > 0 && py1 > 0 && px2 > px1 && py2 > py1)
                        {
                            return new Rectangle(px1, py1, px2 - px1, py2 - py1);
                        }
                    }
                    catch { }
                }

                // 方案 2: 基于 Window.PointsToScreenPixels 兜底换算
                int wx1 = win.PointsToScreenPixelsX((int)cellLeft);
                int wy1 = win.PointsToScreenPixelsY((int)cellTop);
                int wx2 = win.PointsToScreenPixelsX((int)(cellLeft + cellWidth));
                int wy2 = win.PointsToScreenPixelsY((int)(cellTop + cellHeight));

                return new Rectangle(wx1, wy1, Math.Max(wx2 - wx1, 40), Math.Max(wy2 - wy1, 20));
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"计算单元格屏幕坐标异常: {ex.Message}");
                return new Rectangle(100, 100, 120, 24);
            }
        }

        /// <summary>
        /// 激活并覆盖在指定 C 列单元格正上方进行输入与联想
        /// </summary>
        /// <param name="targetCell">当前选中的 Excel C 列活动单元格</param>
        /// <param name="components">当前勾选数据源的所有有效物料集合</param>
        /// <param name="config">智能输入联动配置</param>
        public void ShuRu(dynamic targetCell, List<SmartComponentItem> components, SmartInputConfigModel config)
        {
            if (targetCell == null) return;

            try
            {
                // 保存上下文引用
                _currentTargetCell = targetCell;
                _allComponents = components ?? new List<SmartComponentItem>();
                _config = config ?? new SmartInputConfigModel();

                // 构建字典用于 Addfuzhu 快速匹配
                _componentDict.Clear();
                foreach (var c in _allComponents)
                {
                    if (!string.IsNullOrWhiteSpace(c.Model) && !_componentDict.ContainsKey(c.Model))
                    {
                        _componentDict[c.Model] = c;
                    }
                }

                // 高精度计算单元格在屏幕上的真实物理像素矩形
                Rectangle cellRect = GetCellScreenRectangle(targetCell);
                _screenX = cellRect.X;
                _screenY = cellRect.Y;
                _cellWidth = cellRect.Width;
                _cellHeight = cellRect.Height;

                // 读取单元格原有文本内容
                string initialText = Convert.ToString(targetCell.Value) ?? "";

                // 设置 TextBox 尺寸与初始文本
                _isInternalUpdating = true;
                _textbox.Text = initialText;
                _isInternalUpdating = false;

                // 设置 TextBox 几何位置 (严丝合缝填满单元格)
                _textbox.SetBounds(0, 0, _cellWidth, _cellHeight);

                // 初始收起 ListBox
                _listbox.Visible = false;
                _listbox.Items.Clear();

                // 初始调整窗体几何尺寸 (精准覆盖单元格自身)
                this.SetBounds(_screenX, _screenY, _cellWidth, _cellHeight);

                // 显示窗体并使 TextBox 获取输入焦点
                if (!this.Visible)
                {
                    this.Show();
                }

                _textbox.Focus();
                _textbox.Select(_textbox.Text.Length, 0);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"覆盖输入框 ShuRu 启动异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 响应 TextBox 文本变动事件，执行即时模糊过滤与 ListBox 自适应展开 (对应 ZhiNengEn.Textbox_TextChanged)
        /// </summary>
        private void Textbox_TextChanged(object? sender, EventArgs e)
        {
            if (_isInternalUpdating) return;

            try
            {
                string tex = _textbox.Text.Trim();

                // 清空已有候选项
                _listbox.Items.Clear();

                if (!string.IsNullOrEmpty(tex) && _allComponents.Count > 0)
                {
                    // 模糊匹配: 规格型号或元件名称包含输入字符
                    var matchedList = _allComponents
                        .Where(c => (!string.IsNullOrEmpty(c.Model) && c.Model.IndexOf(tex, StringComparison.OrdinalIgnoreCase) >= 0)
                                 || (!string.IsNullOrEmpty(c.Name) && c.Name.IndexOf(tex, StringComparison.OrdinalIgnoreCase) >= 0))
                        .Take(30)
                        .ToList();

                    foreach (var item in matchedList)
                    {
                        _listbox.Items.Add(item);
                    }
                }

                // 若有匹配项则展开 ListBox，否则收起
                if (_listbox.Items.Count > 0)
                {
                    // 宽度自适应 (至少保持单元格宽度，最高不超过 320)
                    int listWidth = Math.Max(_cellWidth, 220);

                    // 高度自适应 (最多展示 8 行)
                    int listHeight = Math.Min(_listbox.Items.Count * 22 + 4, 180);

                    // 保持 TextBox 始终对准单元格
                    _textbox.SetBounds(0, 0, _cellWidth, _cellHeight);

                    // 重新排布 ListBox 控件
                    _listbox.SetBounds(0, _cellHeight, listWidth, listHeight);
                    _listbox.Visible = true;

                    // 扩展外层窗体尺寸以容纳 ListBox
                    this.SetBounds(_screenX, _screenY, listWidth, _cellHeight + listHeight);
                }
                else
                {
                    // 无匹配项时收起 ListBox 并还原窗体尺寸
                    _listbox.Visible = false;
                    _textbox.SetBounds(0, 0, _cellWidth, _cellHeight);
                    this.SetBounds(_screenX, _screenY, _cellWidth, _cellHeight);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"模糊过滤匹配异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 重写顶级命令键处理，确保 Tab 键、Enter 键和 Escape 键不被 WinForms 内部焦点切换机制吞噬
        /// </summary>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            try
            {
                // 1. Tab 键或 Enter 键: 若 ListBox 有选中项则按选中项提交，否则按 TextBox 内容提交并隐藏
                if (keyData == Keys.Tab || keyData == Keys.Enter)
                {
                    if (_listbox.Visible && _listbox.SelectedIndex >= 0)
                    {
                        SelectCurrentListItemAndCommit();
                    }
                    else
                    {
                        CommitAndHide();
                        MoveActiveCell(0, 1);
                    }
                    return true;
                }

                // 2. Escape 键: 取消并直接隐藏
                if (keyData == Keys.Escape)
                {
                    this.Hide();
                    return true;
                }

                // 3. 向下箭头: 若 ListBox 展开则转入 ListBox，否则向下切单元格
                if (keyData == Keys.Down)
                {
                    if (_listbox.Visible && _listbox.Items.Count > 0)
                    {
                        if (!_listbox.Focused)
                        {
                            _listbox.SelectedIndex = 0;
                            _listbox.Focus();
                            return true;
                        }
                    }
                    else
                    {
                        CommitAndHide();
                        MoveActiveCell(1, 0);
                        return true;
                    }
                }

                // 4. 向上箭头: 若 ListBox 未展开则向上切单元格
                if (keyData == Keys.Up && !_listbox.Visible)
                {
                    CommitAndHide();
                    MoveActiveCell(-1, 0);
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ProcessCmdKey 拦截异常: {ex.Message}");
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        /// <summary>
        /// 响应 TextBox 键盘按键事件 (对应 ZhiNengEn.Textbox_KeyDown)
        /// </summary>
        private void Textbox_KeyDown(object? sender, KeyEventArgs e)
        {
            // 核心按键已在 ProcessCmdKey 统一优先处理
        }

        /// <summary>
        /// 响应 ListBox 键盘按键事件 (对应 ZhiNengEn.Listbox_KeyDown)
        /// </summary>
        private void Listbox_KeyDown(object? sender, KeyEventArgs e)
        {
            try
            {
                switch (e.KeyData)
                {
                    // 回车、Tab 或向右键: 选定当前物料项，立即提交回填并自动隐藏输入框
                    case Keys.Enter:
                    case Keys.Tab:
                    case Keys.Right:
                        SelectCurrentListItemAndCommit();
                        e.Handled = true;
                        break;

                    // Esc 键: 仅收起 ListBox 并将焦点交回 TextBox
                    case Keys.Escape:
                        _listbox.Visible = false;
                        this.SetBounds(_screenX, _screenY, _cellWidth, _cellHeight);
                        _textbox.Focus();
                        e.Handled = true;
                        break;
                }
            }
            catch { }
        }

        /// <summary>
        /// 鼠标单击 ListBox 候选项: 选定、自动回填同行属性并立即隐藏 _textbox
        /// </summary>
        private void Listbox_Click(object? sender, EventArgs e)
        {
            SelectCurrentListItemAndCommit();
        }

        /// <summary>
        /// 鼠标双击 ListBox 候选项: 选定、自动回填同行属性并立即隐藏 _textbox
        /// </summary>
        private void Listbox_DoubleClick(object? sender, EventArgs e)
        {
            SelectCurrentListItemAndCommit();
        }

        /// <summary>
        /// 选定 ListBox 当前高亮项，立即提交写入 Excel 并执行 Addfuzhu 联动回填同行属性，最后隐藏窗体
        /// </summary>
        private void SelectCurrentListItemAndCommit()
        {
            if (_listbox.SelectedItem is SmartComponentItem selectedItem)
            {
                _isInternalUpdating = true;
                _textbox.Text = selectedItem.Model;
                _isInternalUpdating = false;

                // 立即执行 C 列写入及 Addfuzhu 同行属性带入，并隐藏覆盖输入窗体
                CommitAndHide();

                // 自动将 Excel 焦点切换至下一列 (第 4 列: D列)
                MoveActiveCell(0, 1);
            }
        }

        /// <summary>
        /// 提交当前 TextBox 文本至 Excel，并自动执行 Addfuzhu 联动带入同行属性 (对应 ZhiNengEn.Addfuzhu)
        /// </summary>
        public void CommitAndHide()
        {
            try
            {
                if (_currentTargetCell == null)
                {
                    this.Hide();
                    return;
                }

                string finalModel = _textbox.Text.Trim();

                // 1. 将规格型号写入当前 C 列单元格
                _currentTargetCell.Value = finalModel;

                // 2. 执行 Addfuzhu: 根据规格型号自动从物料字典查找并联动回填同行属性
                if (!string.IsNullOrEmpty(finalModel) && _componentDict.TryGetValue(finalModel, out var matchedItem))
                {
                    Addfuzhu(matchedItem);
                }

                // 3. 隐藏输入窗体
                this.Hide();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"提交输入内容与属性联动异常: {ex.Message}");
                this.Hide();
            }
        }

        /// <summary>
        /// 自动执行同行多列属性辅助联动带入 (100% 还原 ZhiNengEn.Addfuzhu)
        /// </summary>
        private void Addfuzhu(SmartComponentItem item)
        {
            if (item == null || _currentTargetCell == null) return;

            try
            {
                dynamic sheet = _currentTargetCell.Worksheet;
                int row = Convert.ToInt32(_currentTargetCell.Row);
                dynamic app = _currentTargetCell.Application;

                // 暂停屏幕刷新加速写入
                app.ScreenUpdating = false;
                app.EnableEvents = false;

                try
                {
                    // 联动 B 列 (元件名称)
                    if (_config.FillName && !string.IsNullOrEmpty(item.Name))
                    {
                        sheet.Cells[row, 2].Value = item.Name;
                    }

                    // 联动 D 列 (生产厂家)
                    if (_config.FillManufacturer && !string.IsNullOrEmpty(item.Manufacturer))
                    {
                        sheet.Cells[row, 4].Value = item.Manufacturer;
                    }

                    // 联动 E 列 (计量单位)
                    if (_config.FillUnit && !string.IsNullOrEmpty(item.Unit))
                    {
                        sheet.Cells[row, 5].Value = item.Unit;
                    }

                    // 联动 G 列 (销售单价)
                    if (_config.FillUnitPrice && item.UnitPrice > 0)
                    {
                        sheet.Cells[row, 7].Value = item.UnitPrice;
                    }
                }
                finally
                {
                    // 恢复刷新与事件
                    app.ScreenUpdating = true;
                    app.EnableEvents = true;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"Addfuzhu 联动回填属性异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 偏移选定 Excel 活动单元格 (如回车向右一列，或下键向下一行)
        /// </summary>
        private void MoveActiveCell(int rowOffset, int colOffset)
        {
            try
            {
                if (_currentTargetCell != null)
                {
                    _currentTargetCell.Offset[rowOffset, colOffset].Select();
                }
            }
            catch { }
        }

        /// <summary>
        /// 自定义绘制 ListBox 项 (高仿 Excel 下拉，支持 #009688 主题高亮与型号/名称双列清晰排版)
        /// </summary>
        private void Listbox_DrawItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _listbox.Items.Count) return;

            if (_listbox.Items[e.Index] is SmartComponentItem item)
            {
                bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

                // 选中背景主色 #009688，未选中纯白
                Color bgColor = isSelected ? Color.FromArgb(0, 150, 136) : Color.White;
                Color modelColor = isSelected ? Color.White : Color.FromArgb(30, 41, 59);
                Color extraColor = isSelected ? Color.FromArgb(224, 242, 241) : Color.FromArgb(100, 116, 139);

                using (var brush = new SolidBrush(bgColor))
                {
                    e.Graphics.FillRectangle(brush, e.Bounds);
                }

                // 绘制左侧主规格型号
                using (var font = new Font("Microsoft YaHei", 9f, FontStyle.Bold))
                using (var brush = new SolidBrush(modelColor))
                {
                    string modelText = item.Model ?? "";
                    e.Graphics.DrawString(modelText, font, brush, e.Bounds.X + 4, e.Bounds.Y + 2);
                }

                // 绘制右侧辅助信息 (名称 / 价格)
                string extraText = $"{item.Name ?? ""}{(item.UnitPrice > 0 ? $" | ¥{item.UnitPrice}" : "")}";
                if (!string.IsNullOrEmpty(extraText))
                {
                    using (var subFont = new Font("Microsoft YaHei", 8.5f, FontStyle.Regular))
                    using (var subBrush = new SolidBrush(extraColor))
                    {
                        var size = e.Graphics.MeasureString(extraText, subFont);
                        float rightX = Math.Max(e.Bounds.Right - size.Width - 6, e.Bounds.X + 120);
                        e.Graphics.DrawString(extraText, subFont, subBrush, rightX, e.Bounds.Y + 3);
                    }
                }

                // 绘制底部分割细线
                if (!isSelected)
                {
                    using (var pen = new Pen(Color.FromArgb(241, 245, 249)))
                    {
                        e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
                    }
                }
            }
        }

        /// <summary>
        /// 当失去焦点且光标不在本窗体内部时自动提交并关闭
        /// </summary>
        private void SmartInputOverlayForm_Deactivate(object? sender, EventArgs e)
        {
            try
            {
                if (this.Visible)
                {
                    CommitAndHide();
                }
            }
            catch { }
        }

        /// <summary>
        /// 安全隐藏覆盖窗体
        /// </summary>
        public void SafeHide()
        {
            if (this.IsDisposed || !this.Visible) return;

            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(SafeHide));
                return;
            }

            try
            {
                this.Hide();
            }
            catch { }
        }
    }
}
