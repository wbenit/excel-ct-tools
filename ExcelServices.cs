using System;
using System.Drawing;
using ExcelDna.Integration;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 核心业务服务类，封装单元格数据读写与格式化逻辑
    /// </summary>
    public static class ExcelServices
    {
        // 记录用户是否开启了自动高亮格式的全局变量
        private static bool _isAutoHighlightEnabled = true;

        // 缓存编辑框中用户输入的自定义字符串
        private static string _customMessageText = "来自鑫壬成套服务的示例消息";

        // 保存用户认证通过后的 Token 授权密钥串
        private static string _currentToken = string.Empty;

        // 保存当前登录用户的显示名称
        private static string _currentUserDisplayName = "未登录";

        /// <summary>
        /// 获取或设置授权 Token 密钥串
        /// </summary>
        public static string CurrentToken
        {
            // 读取最新的 Token 凭据
            get => _currentToken;
            // 更新并保存 Token 凭据
            set => _currentToken = value ?? string.Empty;
        }

        /// <summary>
        /// 获取或设置当前登录用户显示的名称
        /// </summary>
        public static string CurrentUserDisplayName
        {
            // 读取用户显示的名称
            get => _currentUserDisplayName;
            // 更新用户显示名称
            set => _currentUserDisplayName = value ?? string.Empty;
        }

        /// <summary>
        /// Excel 主窗口 Win32 句柄包装类
        /// </summary>
        private class ExcelWin32Window : System.Windows.Forms.IWin32Window
        {
            // 存储 Win32 句柄
            public IntPtr Handle { get; }

            // 构造函数初始化句柄
            public ExcelWin32Window(IntPtr handle) => Handle = handle;
        }

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的登录配置窗口
        /// </summary>
        public static void ShowLoginDialog()
        {
            try
            {
                // 开启 Windows 窗体视觉样式支持
                System.Windows.Forms.Application.EnableVisualStyles();

                // 实例化基于 WebView2 的登录窗口容器
                using var form = new LoginForm();

                // 获取 Excel Application 主句柄
                IntPtr excelHwnd = ExcelDnaUtil.WindowHandle;

                // 依据句柄是否存在选择安全的弹出模式
                if (excelHwnd != IntPtr.Zero)
                {
                    // 模态附着至 Excel 主窗口弹出，防止独立线程闪退
                    form.ShowDialog(new ExcelWin32Window(excelHwnd));
                }
                else
                {
                    // 模态弹出
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                // 捕获弹窗异常防止 Excel 崩溃闪退
                System.Windows.Forms.MessageBox.Show($"弹出登录配置窗口发生异常: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 获取或设置自动高亮状态
        /// </summary>
        public static bool IsAutoHighlightEnabled
        {
            // 读取当前高亮开关状态
            get => _isAutoHighlightEnabled;
            // 写入并更新高亮开关状态
            set => _isAutoHighlightEnabled = value;
        }

        /// <summary>
        /// 获取或设置自定义消息文本
        /// </summary>
        public static string CustomMessageText
        {
            // 读取当前设置的文本内容
            get => _customMessageText;
            // 写入文本内容，防止为空
            set => _customMessageText = value ?? string.Empty;
        }

        /// <summary>
        /// 在 Excel 活动单元格写入时间戳和示例文本
        /// </summary>
        public static void InsertTimestampAndData()
        {
            // 获取 Excel 的 COM Application 对象
            dynamic app = ExcelDnaUtil.Application;
            // 获取当前选中的活动单元格
            dynamic activeCell = app.ActiveCell;

            // 校验单元格是否有效
            if (activeCell == null) return;

            // 写入包含当前系统精确时间的测试数据
            activeCell.Value2 = $"[测试数据] {DateTime.Now:yyyy-MM-dd HH:mm:ss}";

            // 判断是否启用了自动高亮选项
            if (_isAutoHighlightEnabled)
            {
                // 将单元格背景颜色设置为淡黄色
                activeCell.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 255, 204));
                // 将单元格文字字体加粗显示
                activeCell.Font.Bold = true;
            }
        }

        /// <summary>
        /// 清除选定区域的所有内容与格式
        /// </summary>
        public static void ClearActiveRange()
        {
            // 获取 Excel Application COM 引用
            dynamic app = ExcelDnaUtil.Application;
            // 获取当前用户框选的单元格区域
            dynamic selection = app.Selection;

            // 校验选中区域是否存在
            if (selection == null) return;

            // 调用 Excel 原生 API 清空数据与格式
            selection.Clear();
        }

        /// <summary>
        /// 将 EditBox 输入框的文本批量赋值给选中区域
        /// </summary>
        public static void ApplyCustomTextToSelection()
        {
            // 获取全局 Excel Application 实例
            dynamic app = ExcelDnaUtil.Application;
            // 获取选中的 Range 单元格集合
            dynamic selection = app.Selection;

            // 判断选中对象有效性
            if (selection == null) return;

            // 将用户输入的值批量赋值到选中的每个单元格
            selection.Value2 = _customMessageText;

            // 依据复选框状态决定是否应用突出格式
            if (_isAutoHighlightEnabled)
            {
                // 设置单元格填充背景为天蓝色
                selection.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(204, 229, 255));
                // 设置字体颜色为深蓝色以增强视效
                selection.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 51, 102));
            }
        }

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“我的企业设置”窗口
        /// </summary>
        public static void ShowEnterpriseSettingsDialog()
        {
            try
            {
                // 启用视觉样式效果
                System.Windows.Forms.Application.EnableVisualStyles();

                // 实例化企业设置 Form 窗体
                using var form = new EnterpriseSettingsForm();

                // 获取 Excel 主窗口的 HWND 句柄
                IntPtr excelHwnd = ExcelDnaUtil.WindowHandle;

                // 判断句柄有效性并选择 Safe 模式
                if (excelHwnd != IntPtr.Zero)
                {
                    // 将 WinForms 绑定为 Excel 的 Owner 模态显示，绝不闪退
                    form.ShowDialog(new ExcelWin32Window(excelHwnd));
                }
                else
                {
                    // 普通模态弹出
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                // 全局捕获异常防止程序闪退
                System.Windows.Forms.MessageBox.Show($"弹出企业设置窗口失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“新建项目”窗口
        /// </summary>
        public static void ShowCreateProjectDialog()
        {
            try
            {
                // 重置上一次创建的目标工作簿路径缓存
                Controllers.ProjectController.LastCreatedTargetFilePath = string.Empty;

                // 启用 Windows 窗体视觉样式效果
                System.Windows.Forms.Application.EnableVisualStyles();

                // 实例化新建项目 Form 窗体
                using var form = new CreateProjectForm();

                // 获取 Excel 主窗口 HWND 句柄
                IntPtr excelHwnd = ExcelDnaUtil.WindowHandle;

                // 依据句柄有效性安全弹出
                if (excelHwnd != IntPtr.Zero)
                {
                    // 模态附着至 Excel 主窗口，避免闪退与层级穿透
                    form.ShowDialog(new ExcelWin32Window(excelHwnd));
                }
                else
                {
                    // 普通模态显示
                    form.ShowDialog();
                }

                // 重点：当 ShowDialog 模态弹窗关闭后，Windows 消息队列会自动向父窗口发送焦点复位消息
                // 立即执行同步激活，并在 50 毫秒后通过 QueueAsMacro 再次进行 Win32 操作系统级置顶
                if (!string.IsNullOrEmpty(Controllers.ProjectController.LastCreatedTargetFilePath))
                {
                    string targetPath = Controllers.ProjectController.LastCreatedTargetFilePath;

                    // 1. 立即同步激活
                    ActivateCreatedWorkbook(targetPath);

                    // 2. 延迟 50ms 避开 OS 消息队列处置期，在 Excel 主线程宏回调中再次强力置顶
                    // System.Threading.Tasks.Task.Delay(50).ContinueWith(_ =>
                    // {
                    //     ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                    //     {
                    //         ActivateCreatedWorkbook(targetPath);
                    //     });
                    // });
                }
            }
            catch (Exception ex)
            {
                // 全局捕获异常防止程序闪退
                System.Windows.Forms.MessageBox.Show($"弹出新建项目窗口失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        // Win32 API 导入：将目标窗口强制置顶到 Desktop 最前台
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        // Win32 API 导入：还原与展示指定窗口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        /// <summary>
        /// 激活指定的项目工作簿及其视口窗口（包含 Win32 操作系统级硬置顶）
        /// </summary>
        public static void ActivateCreatedWorkbook(string targetFilePath)
        {
            try
            {
                // 获取 Excel Application COM 对象
                dynamic app = ExcelDnaUtil.Application;
                // 校验 app 句柄有效性
                if (app == null) return;

                // 获取目标物理文件名
                string targetFileName = System.IO.Path.GetFileName(targetFilePath);

                // 遍历当前运行的所有 Workbooks
                foreach (dynamic wb in app.Workbooks)
                {
                    // 安全转为 string 进行匹配
                    string wbName = Convert.ToString(wb.Name) ?? "";
                    string wbFullName = Convert.ToString(wb.FullName) ?? "";

                    // 精确与包含双重匹配
                    if (string.Equals(wbName, targetFileName, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(wbFullName, targetFilePath, StringComparison.OrdinalIgnoreCase) ||
                        (!string.IsNullOrEmpty(targetFileName) && wbName.IndexOf(targetFileName, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        // 1. 激活工作簿
                        wb.Activate();

                        // 2. 强力显化视口窗口
                        if (wb.Windows.Count > 0)
                        {
                            dynamic win = wb.Windows[1];
                            // 设为可见视口
                            win.Visible = true;
                            // 设为 xlMaximized 最大化 (-4137)
                            win.WindowState = -4137;
                            // 激活视口
                            win.Activate();

                            // 3. Win32 操作系统级置顶与最大化：获取该工作簿独立 Window 句柄 HWND 并硬性置顶最大化
                            try
                            {
                                long hwndVal = Convert.ToInt64(win.Hwnd);
                                IntPtr winHwnd = new IntPtr(hwndVal);
                                // nCmdShow = 3 即 SW_SHOWMAXIMIZED (SW_MAXIMIZE)，强力强制操作系统最大化展现实体窗口
                                ShowWindow(winHwnd, 3);
                                // 强制提拉至最前台
                                SetForegroundWindow(winHwnd);
                            }
                            catch { }
                        }

                        // 4. 选中并激活“项目信息”工作表
                        try
                        {
                            wb.Sheets["项目信息"].Activate();
                        }
                        catch { }

                        // 成功聚焦即退出
                        break;
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 检索物理磁盘上的 CabinetTemplate.xlsx 模板路径
        /// </summary>
        private static string GetCabinetTemplatePath()
        {
            // 获取应用程序运行基准路径
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            // 配置多重备选路径列表
            string[] candidates = new string[]
            {
                System.IO.Path.Combine(baseDir, "Resources", "CabinetTemplate.xlsx"),
                System.IO.Path.Combine(baseDir, "CabinetTemplate.xlsx"),
                System.IO.Path.Combine(Directory.GetCurrentDirectory(), "Resources", "CabinetTemplate.xlsx"),
                System.IO.Path.Combine(Directory.GetCurrentDirectory(), "CabinetTemplate.xlsx"),
                @"C:\Users\15262\Desktop\CabinetTemplate.xlsx"
            };

            // 循环遍历检索文件存在性
            foreach (string path in candidates)
            {
                // 判断物理路径是否存在
                if (System.IO.File.Exists(path))
                {
                    // 返回匹配成功的模板路径
                    return path;
                }
            }
            // 返回默认候选路径
            return candidates[0];
        }

        /// <summary>
        /// 点击“新建箱柜”业务逻辑：
        /// 1. 顶部汇总行：若有空位置则复用，否则从模板复制第7行在选中位置上方插入。
        /// 2. 序号列（A列）：保持原有的 =ROW()-ROW(A$6) 公式，不手动赋值。
        /// 3. 底部明细块：基于模板第41行特征在下半部分动态搜索，插入第41-72行（32行明细）。
        /// </summary>
        public static void CreateNewCabinetFromSelection()
        {
            try
            {
                // 获取 Excel Application COM 全局对象
                dynamic app = ExcelDnaUtil.Application;
                // 校验全局 app 句柄有效性
                if (app == null) return;

                // 获取当前激活的项目工作簿对象
                dynamic wb = app.ActiveWorkbook;
                // 校验工作簿对象有效性
                if (wb == null) return;

                // 获取当前正在编辑的活动工作表
                dynamic activeSheet = app.ActiveSheet;
                // 获取用户当前选中的焦点单元格
                dynamic activeCell = app.ActiveCell;

                // 暂停屏幕刷新以提升性能
                app.ScreenUpdating = false;
                // 屏蔽 Excel 系统操作对话框与警告弹窗
                app.DisplayAlerts = false;

                // 声明模板工作簿句柄
                dynamic? templateWb = null;
                // 声明插入/复用行号变量 insertRow
                int insertRow = 7;

                try
                {
                    // 获取 CabinetTemplate.xlsx 的物理路径
                    string templatePath = GetCabinetTemplatePath();
                    // 只读模式打开模板工作簿
                    if (System.IO.File.Exists(templatePath))
                    {
                        // 打开模板工作簿句柄
                        templateWb = app.Workbooks.Open(templatePath, ReadOnly: true);
                    }

                    // 【配置文件替代硬编码列举】
                    // 1. 原硬编码: "分类1" (模板默认引用的工作表名称)
                    // 2. 替代配置项: ConfigManager.Instance.Current.Excel.DefaultTemplateSheet
                    string templateSheetName = ConfigManager.Instance.Current.Excel.DefaultTemplateSheet;
                    // 获取模板工作簿中的目标工作表
                    dynamic templateSheet = (templateWb != null) ? templateWb.Sheets[templateSheetName] : activeSheet;

                    // 【配置文件替代硬编码列举】
                    // 1. 原硬编码: 行号 7 (成套产品汇总行默认模板行号及起始有效行号)
                    // 2. 替代配置项: ConfigManager.Instance.Current.Excel.TemplateSumRowIndex
                    int headerRowIndex = ConfigManager.Instance.Current.Excel.TemplateSumRowIndex;

                    // 1. 确定顶部汇总行的插入/复用位置 insertRow
                    // 若焦点单元格有效且行号大于等于模板汇总行号
                    if (activeCell != null)
                    {
                        // 安全转为整型行号
                        int cRow = Convert.ToInt32(activeCell.Row);
                        // 若有效行号大于等于模板汇总行号
                        if (cRow >= headerRowIndex) insertRow = cRow;
                    }

                    // 判断当前 insertRow 的 B 列（箱柜名称）是否为空位置
                    string nameInCell = Convert.ToString(activeSheet.Cells[insertRow, 2].Text) ?? "";
                    // 标记当前位置是否为空白行
                    bool isBlankRow = string.IsNullOrWhiteSpace(nameInCell);

                    // 若当前位置已有箱柜数据，则需在当前行上方插入一行
                    if (!isBlankRow)
                    {
                        // 复制模板工作表汇总行整行
                        if (templateSheet != null)
                        {
                            // 复制模板指定汇总行号
                            templateSheet.Rows[headerRowIndex].Copy();
                        }
                        else
                        {
                            // 复制当前工作表指定汇总行号
                            activeSheet.Rows[headerRowIndex].Copy();
                        }
                        // 在当前 insertRow 上方插入整行 (xlShiftDown = -4121)
                        activeSheet.Rows[insertRow].Insert(-4121);
                    }

                    // 2. 计算新箱柜在清单中的顺序位置 cabinetK (即第几个箱柜)
                    int cabinetK = 0;
                    // 从模板汇总起始行开始向下扫描至 insertRow，统计有效的箱柜数量
                    for (int r = headerRowIndex; r <= insertRow; r++)
                    {
                        // 读取该行 B 列的文本
                        string bText = Convert.ToString(activeSheet.Cells[r, 2].Text) ?? "";
                        // 若不为空，说明为一个有效箱柜节点
                        if (!string.IsNullOrWhiteSpace(bText) || r == insertRow)
                        {
                            // 箱柜计数加 1
                            cabinetK++;
                        }
                    }
                    // 兜底保障 cabinetK 至少为 1
                    if (cabinetK < 1) cabinetK = 1;

                    // 生成新箱柜名称
                    string cabinetName = $"箱柜{cabinetK}";

                    // 3. 填入顶部汇总行的数据 (注意：绝对不写 A 列，保留 =ROW()-ROW(A$6) 公式)
                    // 在 B 列写入箱柜名称
                    activeSheet.Cells[insertRow, 2].Value = cabinetName;
                    // 在 E 列写入单位“台”
                    activeSheet.Cells[insertRow, 5].Value = "台";
                    // 在 F 列写入默认数量 1
                    activeSheet.Cells[insertRow, 6].Value = 1;

                    // 【配置文件替代硬编码列举】
                    // 1. 原硬编码: 起始行 41，终止行 72 (模板明细块复制范围 "41:72")
                    // 2. 原硬编码列号: 1 (A列) (模板用于特征识别匹配的列号)
                    // 3. 替代配置项: ConfigManager.Instance.Current.Excel.TemplateDetailStartRowIndex / TemplateDetailEndRowIndex / FeatureColumnIndex
                    int featureRow = ConfigManager.Instance.Current.Excel.TemplateDetailStartRowIndex;
                    int detailEndRow = ConfigManager.Instance.Current.Excel.TemplateDetailEndRowIndex;
                    int featureCol = ConfigManager.Instance.Current.Excel.FeatureColumnIndex;

                    // 动态计算单个明细块所包含的总行数 (例如 72 - 41 + 1 = 32 行)
                    int detailRowCount = detailEndRow - featureRow + 1;
                    // 构造明细块复制的全局行区域字符串 (例如 "41:72")
                    string copyRangeText = $"{featureRow}:{detailEndRow}";

                    string signature = string.Empty;
                    // 若模板有效则读取模板指定特征单元格的文本作为匹配特征值
                    if (templateSheet != null)
                    {
                        // 动态读取模板特征单元格文本内容
                        signature = Convert.ToString(templateSheet.Cells[featureRow, featureCol].Text) ?? "";
                    }
                    // 若从模板未读取到特征文本，则兜底尝试读取当前工作表对应特征位置内容
                    if (string.IsNullOrWhiteSpace(signature))
                    {
                        // 读取当前工作表对应特征位置内容作为特征值
                        signature = Convert.ToString(activeSheet.Cells[featureRow, featureCol].Text) ?? "";
                    }

                    // 存储搜索到的明细块起始行号列表
                    List<int> detailStartRows = new List<int>();
                    // 动态计算获取当前工作表已使用的最大行号 (UsedRange)，消除硬编码 300
                    int usedRowCount = activeSheet.UsedRange.Row + activeSheet.UsedRange.Rows.Count - 1;
                    // 若当前工作表已使用最大行数小于基准行，则兜底设为基准行前一行
                    if (usedRowCount < featureRow - 1) usedRowCount = featureRow - 1;
                    // 从明细起始前一行开始向下扫描 A 列查找匹配标志
                    for (int r = featureRow - 1; r <= usedRowCount; r++)
                    {
                        // 读取 A 列文本
                        string cellText = Convert.ToString(activeSheet.Cells[r, 1].Text) ?? "";
                        // 若包含标志特征文本
                        if (!string.IsNullOrWhiteSpace(cellText) && cellText.Contains(signature))
                        {
                            // 记录匹配的起始行号
                            detailStartRows.Add(r);
                        }
                    }

                    // 确定目标明细块的插入行号 targetDetailRow
                    int targetDetailRow = featureRow;
                    // 若找到的特征块数量大于等于所需位置 cabinetK
                    if (detailStartRows.Count >= cabinetK)
                    {
                        // 在第 cabinetK 个特征行上方插入
                        targetDetailRow = detailStartRows[cabinetK - 1];
                    }
                    else if (detailStartRows.Count > 0)
                    {
                        // 若已有的特征块少于所需位置，放置在最后一个特征块下方明细总行数处
                        targetDetailRow = detailStartRows[detailStartRows.Count - 1] + detailRowCount;
                    }
                    else
                    {
                        // 若一个特征块都没找到，默认从明细起始行开始
                        targetDetailRow = featureRow;
                    }

                    // 5. 从模板复制指定明细行区域并在 targetDetailRow 位置插入
                    if (templateSheet != null)
                    {
                        // 复制模板动态计算的明细行区域
                        templateSheet.Range[copyRangeText].Copy();
                    }
                    else
                    {
                        // 复制当前表动态计算的明细行区域
                        activeSheet.Range[copyRangeText].Copy();
                    }
                    // 在目标行位置插入明细块 (xlShiftDown = -4121)
                    activeSheet.Rows[targetDetailRow].Insert(-4121);

                    // 6. 同步明细表头箱柜名称与顶底公式关联
                    // 在新明细块的表头 (targetDetailRow + 3 行) B 列名称单元格
                    dynamic detHeaderCell = activeSheet.Cells[targetDetailRow + 3, 2];
                    // 顶部汇总行 B 列名称单元格
                    dynamic sumHeaderCell = activeSheet.Cells[insertRow, 2];

                    // 顶部汇总行 A 列跳转锚点单元格
                    dynamic sumAnchorCell = activeSheet.Cells[insertRow, 1];
                    // 底部明细表头 A 列跳转锚点单元格
                    dynamic detAnchorCell = activeSheet.Cells[targetDetailRow + 3, 1];

                    // 写入顶部汇总行箱柜名称初始值 (使用纯文本，不依赖公式)
                    sumHeaderCell.Value = cabinetName;
                    // 写入底部明细表头箱柜名称初始值 (使用纯文本，不依赖公式)
                    detHeaderCell.Value = cabinetName;

                    // 设置顶部汇总行 G 列 (单价) 的关联公式：指向明细块中单台合计 (targetDetailRow + 27 行) H 列
                    int subtotalRow = targetDetailRow + 27;
                    // 写入 G 列单价联动公式
                    activeSheet.Cells[insertRow, 7].Formula = $"=H{subtotalRow}";
                    // 写入 H 列总价联动公式 (=F{insertRow}*G{insertRow})
                    activeSheet.Cells[insertRow, 8].Formula = $"=F{insertRow}*G{insertRow}";

                    // 绑定定义名称 (Defined Names) 实现无偏移超链接与双向修改
                    // 【配置文件替代硬编码列举】
                    // 1. 原硬编码: "Cab_Sum_" (顶部汇总行定义名称前缀)
                    // 2. 原硬编码: "Cab_Det_" (底部明细块定义名称前缀)
                    // 3. 替代配置项: ConfigManager.Instance.Current.Excel.SumNamePrefix / DetNamePrefix
                    string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix;
                    string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix;
                    string sumNameTag = $"{sumPrefix}{cabinetK}";
                    string detNameTag = $"{detPrefix}{cabinetK}";

                    try
                    {
                        // 在工作簿中定义顶部 A列单元格名称
                        wb.Names.Add(Name: sumNameTag, RefersTo: $"='{activeSheet.Name}'!$A${insertRow}");
                        // 在工作簿中定义底部 A列单元格名称
                        wb.Names.Add(Name: detNameTag, RefersTo: $"='{activeSheet.Name}'!$A${targetDetailRow + 3}");

                        // 获取顶部及底部 A 列原有的文本，防止覆盖序号或标识
                        string sumAnchorText = Convert.ToString(sumAnchorCell.Text) ?? "";
                        string detAnchorText = Convert.ToString(detAnchorCell.Text) ?? "";
                        // 若 A列文本为空，则使用默认名称
                        if (string.IsNullOrWhiteSpace(sumAnchorText)) sumAnchorText = cabinetName;
                        if (string.IsNullOrWhiteSpace(detAnchorText)) detAnchorText = cabinetName;

                        // 为顶部 A列单元格添加指向底部定义名称的超链接
                        activeSheet.Hyperlinks.Add(Anchor: sumAnchorCell, Address: "", SubAddress: $"'{activeSheet.Name}'!{detNameTag}", TextToDisplay: sumAnchorText);
                        // 为底部 A列单元格添加指向顶部定义名称的超链接
                        activeSheet.Hyperlinks.Add(Anchor: detAnchorCell, Address: "", SubAddress: $"'{activeSheet.Name}'!{sumNameTag}", TextToDisplay: detAnchorText);
                    }
                    catch { }

                    // 自动注册工作表修改双向同步事件
                    RegisterSheetChangeEvent();

                    // 7. 焦点定位：激活原工作表并安全选中新插入的单元格
                    try
                    {
                        // 确保激活目标工作簿
                        wb.Activate();
                        // 确保激活目标工作表
                        activeSheet.Activate();
                        // 安全选中 B 列单元格
                        activeSheet.Cells[insertRow, 2].Select();
                    }
                    catch { }
                }
                finally
                {
                    // 若打开了模板工作簿，进行安全关闭
                    if (templateWb != null)
                    {
                        // 关闭模板工作簿，不保存修改
                        try { templateWb.Close(false); } catch { }
                    }

                    // 恢复 Excel 屏幕实时刷新功能
                    app.ScreenUpdating = true;
                    // 恢复 Excel 系统操作对话框与警告弹窗
                    app.DisplayAlerts = true;

                    // 兜底再次尝试激活工作表与选中单元格
                    try
                    {
                        // 激活目标工作表
                        activeSheet.Activate();
                        // 选中单元格
                        activeSheet.Cells[insertRow, 2].Select();
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                // 弹出捕获的异常信息提示
                System.Windows.Forms.MessageBox.Show($"新建箱柜异常: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        // 保持对 Excel Application 实例的静态强引用，防止 GC 回收 COM 事件下沉节点
        private static Microsoft.Office.Interop.Excel.Application? _excelApp = null;

        /// <summary>
        /// 注册 Excel 全局 SheetChange 事件，拦截箱柜名称双向修改
        /// </summary>
        public static void RegisterSheetChangeEvent()
        {
            try
            {
                // 获取并保持 Excel Application 静态引用，避免被 GC 回收
                _excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;

                // 校验 _excelApp 对象有效性
                if (_excelApp == null) return;

                // 强行开启 Excel 系统级 EnableEvents 选项
                _excelApp.EnableEvents = true;

                // 先解除已有委托绑定，避免叠加；再重新绑定 SheetChange 事件处理函数 --硬编码--
                _excelApp.SheetChange -= OnSheetChange;
                _excelApp.SheetChange += OnSheetChange;
            }
            catch (Exception ex)
            {
                // 弹出注册异常提示帮助诊断
                System.Windows.Forms.MessageBox.Show($"注册 SheetChange 事件失败: {ex.Message}", "系统提示");
            }
        }

        /// <summary>
        /// 响应单元格修改事件，实现箱柜名称在顶部与底部的纯文本双向同步更新
        /// </summary>
        private static void OnSheetChange(object shObj, Microsoft.Office.Interop.Excel.Range target)
        {
            try
            {
                // 校验 app 与 target 句柄有效性
                if (_excelApp == null || target == null) return;

                // 转换目标 Sheet 工作表对象
                Microsoft.Office.Interop.Excel.Worksheet? sh = shObj as Microsoft.Office.Interop.Excel.Worksheet;
                // 校验 sh 对象有效性
                if (sh == null) return;

                // 安全获取改动单元格所在工作簿对象
                Microsoft.Office.Interop.Excel.Workbook? wb = sh.Parent as Microsoft.Office.Interop.Excel.Workbook;
                // 校验 wb 句柄有效性
                if (wb == null) return;

                // 限制仅处理第 2 列 (B列) 的修改
                if (target.Column != 2) return;

                // 读取改动单元格的新纯文本内容
                string newName = Convert.ToString(target.Value) ?? "";
                // 若新文本为空白则直接退出
                if (string.IsNullOrWhiteSpace(newName)) return;

                // 获取改动单元格所在行第 1 列 (A列) 单元格
                Microsoft.Office.Interop.Excel.Range aCell = (Microsoft.Office.Interop.Excel.Range)sh.Cells[target.Row, 1];

                // 校验 A列单元格是否包含超链接锚点
                if (aCell != null && aCell.Hyperlinks != null && aCell.Hyperlinks.Count > 0)
                {
                    // 获取超链接跳转目标子地址 (SubAddress)
                    string subAddr = "";
                    try { subAddr = aCell.Hyperlinks[1].SubAddress ?? ""; } catch { }

                    // 【配置文件替代硬编码列举】
                    // 1. 原硬编码前缀校验: "Cab_Sum_" (汇总标签检测) 和 "Cab_Det_" (明细标签检测)
                    // 2. 替代配置项: ConfigManager.Instance.Current.Excel.SumNamePrefix / DetNamePrefix
                    string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix;
                    string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix;

                    // 若超链接指向明细前缀 (表明当前修改的是顶部汇总行的 B列名称)
                    if (subAddr.Contains(detPrefix))
                    {
                        // 提取对应的目标明细定义名称标签 (例如 Cab_Det_1)
                        string targetTag = extractTag(subAddr, detPrefix);
                        if (!string.IsNullOrEmpty(targetTag))
                        {
                            // 查找对应的底部明细 A列单元格 Range
                            Microsoft.Office.Interop.Excel.Range? detAnchorA = FindRangeByTag(wb, sh, targetTag);
                            if (detAnchorA != null)
                            {
                                // 获取底部明细 B列名称单元格 (右移1列)
                                Microsoft.Office.Interop.Excel.Range detNameB = (Microsoft.Office.Interop.Excel.Range)detAnchorA.Offset[0, 1];

                                // 暂停事件触发防止无限递归
                                _excelApp.EnableEvents = false;
                                try
                                {
                                    // 1. 同步更新底部明细表头 B列纯文本数值
                                    detNameB.Value = newName;
                                }
                                catch { }
                                finally
                                {
                                    // 恢复事件触发机制
                                    _excelApp.EnableEvents = true;
                                }
                                return;
                            }
                        }
                    }
                    // 若超链接指向汇总前缀 (表明当前修改的是底部明细行的 B列名称)
                    else if (subAddr.Contains(sumPrefix))
                    {
                        // 提取对应的目标汇总定义名称标签 (例如 Cab_Sum_1)
                        string targetTag = extractTag(subAddr, sumPrefix);
                        if (!string.IsNullOrEmpty(targetTag))
                        {
                            // 查找对应的顶部汇总 A列单元格 Range
                            Microsoft.Office.Interop.Excel.Range? sumAnchorA = FindRangeByTag(wb, sh, targetTag);
                            if (sumAnchorA != null)
                            {
                                // 获取顶部汇总 B列名称单元格 (右移1列)
                                Microsoft.Office.Interop.Excel.Range sumNameB = (Microsoft.Office.Interop.Excel.Range)sumAnchorA.Offset[0, 1];

                                // 暂停事件触发防止无限递归
                                _excelApp.EnableEvents = false;
                                try
                                {
                                    // 1. 反向同步更新顶部汇总行 B列纯文本数值
                                    sumNameB.Value = newName;
                                }
                                catch { }
                                finally
                                {
                                    // 恢复事件触发机制
                                    _excelApp.EnableEvents = true;
                                }
                                return;
                            }
                        }
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 从超链接子地址中截取精准的定义名称标签字符串
        /// </summary>
        private static string extractTag(string subAddr, string prefix)
        {
            try
            {
                // 定位前缀关键字位置
                int idx = subAddr.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                {
                    // 截取前缀及其之后的所有字符串
                    string tag = subAddr.Substring(idx);
                    // 清理单引号分隔符
                    int endIdx = tag.IndexOf('\'');
                    if (endIdx > 0) tag = tag.Substring(0, endIdx);
                    // 返回修剪后的纯净标签名
                    return tag.Trim();
                }
            }
            catch { }
            return "";
        }

        /// <summary>
        /// 依据定义名称标签精准获取 Range 对象，带多层 Safe-Lookup 兜底
        /// </summary>
        private static Microsoft.Office.Interop.Excel.Range? FindRangeByTag(Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.Worksheet sh, string tagName)
        {
            try
            {
                // 方式 1：直接通过 wb.Names.Item 按标签名提取
                try
                {
                    Microsoft.Office.Interop.Excel.Name n1 = wb.Names.Item(tagName);
                    if (n1 != null && n1.RefersToRange != null) return n1.RefersToRange;
                }
                catch { }

                // 方式 2：通过带有工作表前缀的 Full Name 提取
                try
                {
                    string fullName = $"'{sh.Name}'!{tagName}";
                    Microsoft.Office.Interop.Excel.Name n2 = wb.Names.Item(fullName);
                    if (n2 != null && n2.RefersToRange != null) return n2.RefersToRange;
                }
                catch { }

                // 方式 3：遍历工作簿中的定义名称进行后缀匹配
                foreach (Microsoft.Office.Interop.Excel.Name n in wb.Names)
                {
                    string nStr = n.Name ?? "";
                    if (nStr.EndsWith(tagName, StringComparison.OrdinalIgnoreCase))
                    {
                        return n.RefersToRange;
                    }
                }
            }
            catch { }
            return null;
        }
    }
}
