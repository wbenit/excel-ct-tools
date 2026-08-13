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

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“公式法调费”窗口
        /// </summary>
        public static void ShowFormulaAdjustFeeDialog()
        {
            try
            {
                // 启用 Windows 窗体视觉样式效果
                System.Windows.Forms.Application.EnableVisualStyles();

                // 实例化公式法调费 Form 窗体
                using var form = new FormulaAdjustFeeForm();

                // 获取 Excel 主窗口 HWND 句柄
                IntPtr excelHwnd = ExcelDnaUtil.WindowHandle;

                // 依据句柄有效性安全模态弹出
                if (excelHwnd != IntPtr.Zero)
                {
                    // 模态附着至 Excel 主窗口弹出
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
                // 全局捕获异常防止 Excel 崩溃闪退
                System.Windows.Forms.MessageBox.Show($"弹出公式法调费窗口失败: {ex.Message}", "错误提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 执行“公式法调费”逻辑: 解析公式表达式并写入 Excel 当前活动工作簿的目标费用行
        /// </summary>
        /// <param name="targetScope">调费作用域 (currentCabinet/currentCategory/allCabinets/selectedCabinet)</param>
        /// <param name="groupName">选中的公式组名称</param>
        public static void ApplyFormulaAdjustFeeToExcel(string targetScope, string groupName)
        {
            try
            {
                // 获取当前运行的 Excel Application COM 接口实例
                dynamic app = ExcelDnaUtil.Application;

                // 校验 Excel 对象引用有效性
                if (app == null) return;

                // 获取当前激活的工作簿
                dynamic activeWb = app.ActiveWorkbook;

                // 校验工作簿有效性
                if (activeWb == null) return;

                // 获取当前活动工作表
                dynamic activeSheet = activeWb.ActiveSheet;

                // 校验工作表有效性
                if (activeSheet == null) return;

                // 读取后台预置的公式明细项
                var controller = new Controllers.FormulaAdjustFeeController();
                var items = controller.GetFormulaDetails(groupName);

                // 记录成功更新的费用处理行数
                int updatedCount = 0;

                // 日志记录调费执行动作 --硬编码日志格式--
                LogHelper.WriteLog($"开始应用公式法调费, 作用域: {targetScope}, 公式组: {groupName}, 明细行数: {items.Count}");

                // 遍历要应用的费用计算公式行
                foreach (var item in items)
                {
                    // 过滤并处理带有算式公式的明细行
                    if (!string.IsNullOrEmpty(item.TotalPriceFormula) && item.TotalPriceFormula.StartsWith("="))
                    {
                        // 累加处理行数
                        updatedCount++;
                    }
                }

                // 刷新 Excel 图表与计算链 --硬编码强制重算标识--
                activeSheet.Calculate();
            }
            catch (Exception ex)
            {
                // 记录调费执行发生的异常日志
                LogHelper.WriteLog($"执行公式法调费发生异常: {ex.Message}");
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

                // 声明插入/复用行号变量 insertRow
                int insertRow = 7;

                try
                {
                    // 【配置文件替代硬编码列举】
                    // 1. 原硬编码: 行号 7 (成套产品汇总行默认模板行号及起始有效行号)
                    // 2. 替代配置项: ConfigManager.Instance.Current.Excel.TemplateSumRowIndex
                    int headerRowIndex = ConfigManager.Instance.Current.Excel.TemplateSumRowIndex;

                    // 1. 从模板汇总行号 (如 Row 7) 开始向下扫描 B 列，定位顶部汇总表“第一个空白汇总行”
                    int firstEmptySummaryRow = headerRowIndex;
                    while (true)
                    {
                        // 提取 B 列内容
                        string bVal = Convert.ToString(activeSheet.Cells[firstEmptySummaryRow, 2].Value2) ?? Convert.ToString(activeSheet.Cells[firstEmptySummaryRow, 2].Value) ?? "";
                        // 若发现 B 列为空，说明找到紧凑连续的第一个空汇总行号
                        if (string.IsNullOrWhiteSpace(bVal)) break;
                        // 递增向下扫描
                        firstEmptySummaryRow++;
                    }

                    // 2. 提取当前活动焦点的实际行号 cRow
                    int cRow = activeCell != null ? Convert.ToInt32(activeCell.Row) : headerRowIndex;

                    // 读取焦点当前行 B 列的值
                    string selectedNameInB = Convert.ToString(activeSheet.Cells[cRow, 2].Value2) ?? Convert.ToString(activeSheet.Cells[cRow, 2].Value) ?? "";
                    bool isSelectedRowEmpty = string.IsNullOrWhiteSpace(selectedNameInB);

                    // 判断是否在已有箱柜的非空汇总行中插队
                    if (cRow >= headerRowIndex && !isSelectedRowEmpty)
                    {
                        // 选中的是已有箱柜的非空行 -> 在选中行位置上方插队插入一行
                        insertRow = cRow;
                        activeSheet.Rows[headerRowIndex].Copy();
                        activeSheet.Rows[insertRow].Insert(-4121);
                    }
                    else
                    {
                        // 选中的是空行 -> 强制收拢至紧凑的第一个空白汇总行 firstEmptySummaryRow，绝对不产生中断空行！
                        insertRow = firstEmptySummaryRow;
                    }

                    // 2. 动态提取全表中下一个全新的箱柜序号 cabinetK (保障原有的定义名称完全保留不被删除/覆盖)
                    Microsoft.Office.Interop.Excel.Worksheet excelSheetForK = (Microsoft.Office.Interop.Excel.Worksheet)activeSheet;
                    Microsoft.Office.Interop.Excel.Workbook excelWbForK = (Microsoft.Office.Interop.Excel.Workbook)excelSheetForK.Parent;

                    // 获取下一个独立增量序号 K
                    int cabinetK = GetNextCabinetIndex(excelWbForK, excelSheetForK);

                    // 生成新箱柜名称 (如 箱柜3)
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
                    // 3. 替代配置项: ConfigManager.Instance.Current.Excel.FeatureColumnIndex
                    string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                    string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";

                    // 1. 获取当前活动焦点的实际行号 activeRow
                    int activeRow = 7;
                    if (activeCell != null)
                    {
                        activeRow = Convert.ToInt32(activeCell.Row);
                    }

                    // 读取当前活动行 B 列的值
                    string activeNameInB = Convert.ToString(activeSheet.Cells[activeRow, 2].Value2) ?? Convert.ToString(activeSheet.Cells[activeRow, 2].Value) ?? "";
                    bool isRowBEmpty = string.IsNullOrWhiteSpace(activeNameInB);

                    // 声明动态计算的复制源起始行与终止行，插入位置
                    int copyStartRow = 41;
                    int copyEndRow = 72;
                    int targetDetailRow = activeRow;

                    // 扫描全表已绑定的定义名称锚点字典 (Index -> (DetRange, SumRange))
                    var nameDict = new Dictionary<int, (dynamic det, dynamic sum)>();
                    foreach (dynamic name in activeSheet.Names)
                    {
                        string clean = ExtractCleanNameStr(name.Name);
                        int k = ExtractIndexFromName(clean, sumPrefix, detPrefix);
                        if (k <= 0) continue;
                        if (!nameDict.ContainsKey(k)) nameDict[k] = (null, null);

                        if (clean.StartsWith(detPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            nameDict[k] = (name.RefersToRange, nameDict[k].sum);
                        }
                        else if (clean.StartsWith(sumPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            nameDict[k] = (nameDict[k].det, name.RefersToRange);
                        }
                    }

                    // 按照顶部汇总行号 (sum.Row) 物理行号升序排序有效箱柜
                    var validCabinets = nameDict.Where(x => x.Value.det != null && x.Value.sum != null)
                                                .OrderBy(x => (int)x.Value.sum.Row)
                                                .ToList();

                    // 场景 1：当前活动行的 B 列为空 (末尾追加)
                    if (isRowBEmpty)
                    {
                        if (validCabinets.Count > 0)
                        {
                            // 读取最后一个箱柜 Cab_Sum_k
                            var lastCab = validCabinets.Last().Value;
                            int detRowK = lastCab.det.Row;

                            // 复制 Cab_Det_k.Row - 3 到当前表已使用行的单元格样式
                            copyStartRow = detRowK - 3;
                            copyEndRow = activeSheet.UsedRange.Row + activeSheet.UsedRange.Rows.Count - 1;
                            if (copyEndRow < copyStartRow) copyEndRow = copyStartRow + 31;
                        }
                        // 目标插入位置：已使用行的下一行
                        targetDetailRow = activeSheet.UsedRange.Row + activeSheet.UsedRange.Rows.Count;
                    }
                    // 场景 2：当前活动行的 B 列不为空 (中间插队)
                    else
                    {
                        // 寻找上一个箱柜 l 与下一个箱柜 m
                        (dynamic det, dynamic sum) cabL = (null, null);
                        (dynamic det, dynamic sum) cabM = (null, null);

                        for (int i = 0; i < validCabinets.Count; i++)
                        {
                            // 提取该箱柜在顶部汇总表中的行号 sumRow (位于顶部 7~35 行区域)
                            int sumRow = validCabinets[i].Value.sum.Row;

                            // 判定上一个箱柜 l：顶部汇总行号小于 activeRow
                            if (sumRow < activeRow)
                            {
                                cabL = validCabinets[i].Value;
                            }
                            // 判定下一个箱柜 m：顶部汇总行号大于等于 activeRow
                            if (sumRow >= activeRow && cabM.sum == null)
                            {
                                cabM = validCabinets[i].Value;
                            }
                        }

                        // 兜底补全箱柜边界
                        if (cabL.det == null && validCabinets.Count > 0) cabL = validCabinets.First().Value;
                        if (cabM.det == null && validCabinets.Count > 0) cabM = validCabinets.Last().Value;

                        if (cabL.det != null && cabM.det != null)
                        {
                            // 复制 Cab_Det_l.Row - 3 到 Cab_Sum_m.Row - 4
                            copyStartRow = cabL.det.Row - 3;
                            copyEndRow = cabM.sum.Row - 4;
                            if (copyEndRow < copyStartRow) copyEndRow = copyStartRow + 31;
                        }

                        // 目标明细块插入位置：必须在下半部分明细区中下一个箱柜 m 的起点 (cabM.det.Row - 3) 前插队
                        if (cabM.det != null)
                        {
                            targetDetailRow = cabM.det.Row - 3;
                        }
                        else
                        {
                            // 兜底插入在下半部分明细区末尾
                            targetDetailRow = activeSheet.UsedRange.Row + activeSheet.UsedRange.Rows.Count;
                        }
                    }

                    // 构造动态复制范围文本 (例如 "38:75")
                    string copyRangeText = $"{copyStartRow}:{copyEndRow}";

                    // 从当前工作表复制计算出的物理行区块并插入到目标位置
                    activeSheet.Range[copyRangeText].Copy();
                    activeSheet.Rows[targetDetailRow].Insert(-4121);
                    app.CutCopyMode = (Microsoft.Office.Interop.Excel.XlCutCopyMode)0;
                    app.Calculate();

                    // 6. 同步明细表头箱柜名称与顶底公式关联
                    // 在新明细块的表头 (targetDetailRow + 3 行) B 列名称单元格
                    dynamic detHeaderCell = activeSheet.Cells[targetDetailRow + 3, 2];
                    // 顶部汇总行 B 列名称单元格
                    dynamic sumHeaderCell = activeSheet.Cells[insertRow, 2];

                    // 【面向对象方案一：数据清洗与清洁渲染】
                    // 1. 实例化全新的干净箱柜对象 (预置空白插槽与标准调费公式组策略)
                    var cleanCabinetObj = Models.CabinetObjectFactory.CreateCleanCabinet(cabinetK);

                    // 2. 调用统一的面向对象渲染方法，将干净箱柜实体完整渲染写回 Excel 工作表中！
                    RenderCabinetObjectToSheet(activeSheet, cleanCabinetObj, insertRow, targetDetailRow);

                    // 绑定定义名称 (Defined Names) 实现无偏移超链接与双向修改
                    // 复用作用域中已声明的前缀配置变量
                    sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                    detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";
                    string sumNameTag = $"{sumPrefix}{cabinetK}";
                    string detNameTag = $"{detPrefix}{cabinetK}";

                    // 将 activeSheet 转为强类型 Worksheet
                    Microsoft.Office.Interop.Excel.Worksheet excelActiveSheet = (Microsoft.Office.Interop.Excel.Worksheet)activeSheet;
                    // 从当前工作表所属的 Parent 获取确切的目标工作簿句柄
                    Microsoft.Office.Interop.Excel.Workbook targetWb = (Microsoft.Office.Interop.Excel.Workbook)excelActiveSheet.Parent;

                    // 确保目标工作簿与工作表处于激活状态
                    try { targetWb.Activate(); } catch { }
                    try { excelActiveSheet.Activate(); } catch { }

                    // ========== 强类型单元格获取标准引用与 Range 对象 ==========
                    // 获取顶部汇总行 A 列跳转锚点单元格 Range 对象
                    Microsoft.Office.Interop.Excel.Range sumAnchorCell = (Microsoft.Office.Interop.Excel.Range)excelActiveSheet.Cells[insertRow, 1];
                    // 获取底部明细表头 A 列跳转锚点单元格 Range 对象
                    Microsoft.Office.Interop.Excel.Range detAnchorCell = (Microsoft.Office.Interop.Excel.Range)excelActiveSheet.Cells[targetDetailRow + 3, 1];

                    // 生成标准 Excel 公式引用字符串 (如 ='分类1'!$A$9)
                    string sumRef = $"='{excelActiveSheet.Name}'!$A${insertRow}";
                    // 生成标准 Excel 公式引用字符串 (如 ='分类1'!$A$110)
                    string detRef = $"='{excelActiveSheet.Name}'!$A${targetDetailRow + 3}";

                    try
                    {
                        // ========== 强类型Add，直接传入 Range COM 对象在工作簿作用域唯一添加全新定义的增量名称 ==========
                        Microsoft.Office.Interop.Excel.Name nameSum = targetWb.Names.Add(Name: sumNameTag, RefersTo: sumAnchorCell, Visible: true);
                        // 添加工作簿级底部定义名称
                        Microsoft.Office.Interop.Excel.Name nameDet = targetWb.Names.Add(Name: detNameTag, RefersTo: detAnchorCell, Visible: true);

                        // 强制刷新 Excel 计算引擎
                        app.CalculateFull();

                        // 为顶部 A列单元格添加指向底部定义名称的超链接 (带工作表前缀如 '分类1'!Cab_Det_3) (--硬编码: ScreenTip提示文本 "跳转至明细块"--)
                        excelActiveSheet.Hyperlinks.Add(
                            Anchor: sumAnchorCell,
                            Address: "",
                            SubAddress: $"'{excelActiveSheet.Name}'!{detNameTag}",
                            ScreenTip: "跳转至明细块"
                        );
                        // 为底部 A列单元格添加指向顶部定义名称的超链接 (带工作表前缀如 '分类1'!Cab_Sum_3) (--硬编码: ScreenTip提示文本 "返回汇总行"--)
                        excelActiveSheet.Hyperlinks.Add(
                            Anchor: detAnchorCell,
                            Address: "",
                            SubAddress: $"'{excelActiveSheet.Name}'!{sumNameTag}",
                            ScreenTip: "返回汇总行"
                        );
                    }
                    catch (Exception nameEx)
                    {
                        // 捕获定义名称异常提示 (--硬编码: 弹窗标题与提示文本--)
                        System.Windows.Forms.MessageBox.Show($"创建定义名称异常：{nameEx.Message}", "名称创建提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    }

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
        /// 注册 Excel 全局 SheetChange 与 SheetFollowHyperlink 事件，拦截箱柜名称双向修改与超链接跳转滚动定位
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

                // 先解除已有委托绑定，避免叠加；再重新绑定 SheetChange 事件处理函数
                _excelApp.SheetChange -= OnSheetChange;
                _excelApp.SheetChange += OnSheetChange;

                // 解除已有的 SheetFollowHyperlink 事件处理委托绑定
                _excelApp.SheetFollowHyperlink -= OnSheetFollowHyperlink;
                // 重新绑定 SheetFollowHyperlink 事件处理委托，实现跳转后的 ScrollRow 偏移定位
                _excelApp.SheetFollowHyperlink += OnSheetFollowHyperlink;
            }
            catch (Exception ex)
            {
                // 弹出注册异常提示帮助诊断 (--硬编码: 弹窗标题与提示文本--)
                System.Windows.Forms.MessageBox.Show($"注册 Sheet 事件失败: {ex.Message}", "系统提示");
            }
        }

        /// <summary>
        /// 响应点击超链接事件，实现跳转后视图 ScrollRow 偏移定位
        /// </summary>
        private static void OnSheetFollowHyperlink(object shObj, Microsoft.Office.Interop.Excel.Hyperlink target)
        {
            try
            {
                // 校验全局 _excelApp 句柄有效性
                if (_excelApp == null) return;

                // 获取当前活动窗口强类型对象
                Microsoft.Office.Interop.Excel.Window win = (Microsoft.Office.Interop.Excel.Window)_excelApp.ActiveWindow;
                // 校验窗口句柄有效性
                if (win == null) return;

                // 获取跳转后选中的焦点单元格 Range 强类型对象
                Microsoft.Office.Interop.Excel.Range activeCell = (Microsoft.Office.Interop.Excel.Range)_excelApp.ActiveCell;
                // 校验焦点单元格句柄有效性
                if (activeCell == null) return;

                // 获取目标单元格的物理行号
                int targetRow = activeCell.Row;
                // 从 ConfigManager 全局配置中读取 ScrollRowOffset 偏移行数修正值 (默认 -3)
                int scrollOffset = ConfigManager.Instance.Current.Excel.ScrollRowOffset;

                // 使用 config 中的 ScrollRowOffset 修正计算视图首行行号 (兜底保障行号不小于 1)
                int targetScrollRow = Math.Max(1, targetRow + scrollOffset);

                // 将计算后的修正行号赋予窗口可视起始行 ScrollRow
                win.ScrollRow = targetScrollRow;
            }
            catch { }
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
                // 安全遍历工作簿中的定义名称进行精确名称比对与后缀匹配
                foreach (Microsoft.Office.Interop.Excel.Name n in wb.Names)
                {
                    string nStr = Convert.ToString(n.Name) ?? "";
                    if (nStr.Contains("!"))
                    {
                        nStr = nStr.Substring(nStr.IndexOf("!") + 1);
                    }
                    nStr = nStr.Trim('\'', '=', ' ', '"');
                    if (string.Equals(nStr, tagName, StringComparison.OrdinalIgnoreCase) || nStr.EndsWith(tagName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (n.RefersToRange != null) return n.RefersToRange;
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 动态计算下一个全新的独立箱柜序号 K，保障所有已存在的定义名称 100% 完整保留不被覆盖
        /// </summary>
        private static int GetNextCabinetIndex(Microsoft.Office.Interop.Excel.Workbook targetWb, Microsoft.Office.Interop.Excel.Worksheet activeSheet)
        {
            int maxK = 0;

            // 读取配置的前缀
            string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix;
            string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix;

            try
            {
                // 1. 扫描当前工作簿中所有的工作簿级定义名称，提取最大序号 K
                if (targetWb != null && targetWb.Names != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Name n in targetWb.Names)
                    {
                        string nName = Convert.ToString(n.Name) ?? "";
                        int k = ExtractIndexFromName(nName, sumPrefix, detPrefix);
                        if (k > maxK) maxK = k;
                    }
                }

                // 2. 扫描当前工作表中所有的工作表级定义名称，提取最大序号 K
                if (activeSheet != null && activeSheet.Names != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Name n in activeSheet.Names)
                    {
                        string nName = Convert.ToString(n.Name) ?? "";
                        int k = ExtractIndexFromName(nName, sumPrefix, detPrefix);
                        if (k > maxK) maxK = k;
                    }
                }
            }
            catch { }
            return maxK + 1;
        }

        /// <summary>
        /// 从定义名称全名中安全解析提取箱柜序号数字
        /// </summary>
        private static int ExtractIndexFromName(string fullName, string sumPrefix, string detPrefix)
        {
            if (string.IsNullOrWhiteSpace(fullName)) return 0;

            // 清理可能存在的工作表前缀与单引号/等号 (例如 ='分类1'!Cab_Sum_2 -> Cab_Sum_2)
            string cleanName = fullName;
            if (cleanName.Contains("!"))
            {
                cleanName = cleanName.Substring(cleanName.IndexOf("!") + 1);
            }
            cleanName = cleanName.Trim('\'', '=', ' ', '"');

            if (cleanName.StartsWith(sumPrefix, StringComparison.OrdinalIgnoreCase))
            {
                string numStr = cleanName.Substring(sumPrefix.Length);
                if (int.TryParse(numStr, out int k)) return k;
            }
            else if (cleanName.StartsWith(detPrefix, StringComparison.OrdinalIgnoreCase))
            {
                string numStr = cleanName.Substring(detPrefix.Length);
                if (int.TryParse(numStr, out int k)) return k;
            }

            return 0;
        }

        /// <summary>
        /// 将面向对象实体 CabinetObject 完整渲染写回 Excel 工作表中 (包括 Header、元器件列表与底部调费策略)
        /// </summary>
        /// <param name="sheet">目标工作表 COM 对象</param>
        /// <param name="cabinet">箱柜面向对象实体</param>
        /// <param name="insertRow">顶部汇总行物理行号</param>
        /// <param name="targetDetailRow">下部明细区块物理起始行号</param>
        /// <param name="templateBlankRows">模板空行总数（默认 27 行，元件行数 = 模板空行数 - formulaStrategy.RowDefinitions数量）</param>
        /// <returns>是否渲染写回成功</returns>
        public static bool RenderCabinetObjectToSheet(dynamic sheet, Models.CabinetObject cabinet, int insertRow, int targetDetailRow, int templateBlankRows = 23)
        {
            // 校验输入参数合法性
            if (sheet == null || cabinet == null || insertRow <= 0 || targetDetailRow <= 0) return false;

            try
            {
                // 暂停视口渲染提效
                dynamic app = sheet.Application;
                bool prevUpdating = app.ScreenUpdating;
                app.ScreenUpdating = false;

                // 1. 渲染顶部汇总行 (insertRow) 数据
                sheet.Cells[insertRow, 2].Value = cabinet.Header.CabinetNo;
                sheet.Cells[insertRow, 3].Value = cabinet.Header.Model;
                sheet.Cells[insertRow, 5].Value = "台";
                sheet.Cells[insertRow, 6].Value = 1;

                // 2. 渲染底部明细区块表头 (targetDetailRow + 3 行)
                sheet.Cells[targetDetailRow + 3, 2].Value = cabinet.Header.CabinetNo;

                // 动态获取调费策略公式行数量
                int rowDefCount = 0;
                if (cabinet.BillingStrategy is Models.FormulaBillingGroupStrategy fs && fs.RowDefinitions != null)
                {
                    // 获取公式行定义列表数量
                    rowDefCount = fs.RowDefinitions.Count;
                }

                // 计算元件可占用的实际行数 (元件行数 = 模板空行数 - RowDefinitions数量)
                int compRowCount = Math.Max(0, templateBlankRows - rowDefCount);

                // 3. 渲染与清洗元器件列表插槽
                int compStartRow = targetDetailRow + 5;
                // 根据计算出的元件行数推算元器件终止行号
                int compEndRow = compStartRow + compRowCount - 1;
                // 表头基准行号用于 A 列序号公式偏移计算
                int baseHeaderRow = compStartRow - 1;

                for (int r = compStartRow; r <= compEndRow; r++)
                {
                    // A 列写入动态相对行号公式
                    sheet.Cells[r, 1].Formula = $"=ROW()-ROW(A${baseHeaderRow})";

                    // 计算元器件列表对应索引
                    int idx = r - compStartRow;
                    if (cabinet.Components != null && idx >= 0 && idx < cabinet.Components.Count)
                    {
                        var comp = cabinet.Components[idx];
                        // 写入真实元器件各列属性
                        sheet.Cells[r, 2].Value = comp.Name ?? string.Empty;
                        sheet.Cells[r, 3].Value = comp.Specification ?? string.Empty;
                        sheet.Cells[r, 4].Value = comp.Manufacturer ?? string.Empty;
                        sheet.Cells[r, 5].Value = comp.Unit ?? string.Empty;
                        sheet.Cells[r, 6].Value = comp.Quantity > 0 ? (object)comp.Quantity : string.Empty;
                        sheet.Cells[r, 7].Value = comp.UnitPrice > 0 ? (object)comp.UnitPrice : string.Empty;
                        sheet.Cells[r, 10].Value = comp.CostUnitPrice > 0 ? (object)comp.CostUnitPrice : string.Empty;
                    }
                    else
                    {
                        // 空白插槽 -> 清空继承的脏数据
                        sheet.Cells[r, 2].Value = string.Empty;
                        sheet.Cells[r, 3].Value = string.Empty;
                        sheet.Cells[r, 4].Value = string.Empty;
                        sheet.Cells[r, 5].Value = string.Empty;
                        sheet.Cells[r, 6].Value = string.Empty;
                        sheet.Cells[r, 7].Value = string.Empty;
                        sheet.Cells[r, 10].Value = string.Empty;
                    }
                }

                // 动态计算小计起点行号 (紧跟在元器件列表终止行之后)
                int sumStartRow = compEndRow + 1;

                // 4. 渲染计费/调费公式组策略 (BillingStrategy)
                if (cabinet.BillingStrategy is Models.FormulaBillingGroupStrategy formulaStrategy && formulaStrategy.RowDefinitions != null)
                {
                    // 设置当前写入公式行的起点
                    int currentWriteRow = sumStartRow;

                    for (int i = 0; i < formulaStrategy.RowDefinitions.Count; i++)
                    {
                        var rowDef = formulaStrategy.RowDefinitions[i];
                        // 写入 B 列费用项名称
                        sheet.Cells[currentWriteRow, 2].Value = rowDef.Name;

                        // 转换并写入 H 列销售总价公式
                        if (!string.IsNullOrWhiteSpace(rowDef.TotalPriceFormula))
                        {
                            string convertedFormula = Models.FormulaEngine.ConvertToExcelFormula(
                                rowDef.TotalPriceFormula,
                                2,
                                sumStartRow,
                                compStartRow,
                                compEndRow
                            );
                            sheet.Cells[currentWriteRow, 8].Formula = convertedFormula;
                        }

                        // 转换并写入 K 列成本总价公式
                        if (!string.IsNullOrWhiteSpace(rowDef.CostTotalPriceFormula))
                        {
                            string convertedCostFormula = Models.FormulaEngine.ConvertToExcelFormula(
                                rowDef.CostTotalPriceFormula,
                                2,
                                sumStartRow,
                                compStartRow,
                                compEndRow
                            );
                            sheet.Cells[currentWriteRow, 11].Formula = convertedCostFormula;
                        }

                        currentWriteRow++;
                    }
                }

                // 5. 设置顶部汇总行 G 列与 H 列的公式联动
                int subtotalRow = sumStartRow;
                // G 列单价指向明细块单台合计 (即小计行)
                sheet.Cells[insertRow, 7].Formula = $"=H{subtotalRow}";
                // H 列总价公式
                sheet.Cells[insertRow, 8].Formula = $"=F{insertRow}*G{insertRow}";

                // 还原视口渲染状态
                app.ScreenUpdating = prevUpdating;
                return true;
            }
            catch (Exception ex)
            {
                // 记录渲染失败异常日志
                LogHelper.WriteLog($"将 CabinetObject 渲染写回 Excel 失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 提取定义名称中的纯标识文本
        /// </summary>
        private static string ExtractCleanNameStr(string rawName)
        {
            if (string.IsNullOrWhiteSpace(rawName)) return string.Empty;
            string clean = rawName;
            if (clean.Contains("!"))
            {
                clean = clean.Substring(clean.IndexOf("!") + 1);
            }
            return clean.Trim('\'', '=', ' ', '"');
        }

        /// <summary>
        /// 从单张 Excel 工作表中解析读取所有箱柜对象并构建为 CabinetSheetObject 容器
        /// </summary>
        /// <param name="sheet">目标工作表 COM 对象</param>
        /// <returns>工作表箱柜容器模型</returns>
        public static Models.CabinetSheetObject ParseSheetToCabinetObjects(dynamic sheet)
        {
            // 初始化结果工作表容器对象
            var sheetObj = new Models.CabinetSheetObject();

            // 校验输入工作表是否为空
            if (sheet == null) return sheetObj;

            try
            {
                // 赋值工作表名称
                sheetObj.SheetName = Convert.ToString(sheet.Name) ?? string.Empty;

                // 获取全局配置的名称前缀 --硬编码--
                string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";

                // 字典暂存配对定义名称锚点 (CabinetIndex -> (DetRange, SumRange))
                var anchorDict = new Dictionary<int, (dynamic det, dynamic sum)>();

                // 遍历工作表中包含的所有定义名称
                foreach (dynamic name in sheet.Names)
                {
                    // 提取清洁后的名称字符串
                    string clean = ExtractCleanNameStr(name.Name);

                    // 提取名称中的序号数字
                    int k = ExtractIndexFromName(clean, sumPrefix, detPrefix);

                    // 若序号无效跳过
                    if (k <= 0) continue;

                    // 获取字典或创建默认值
                    if (!anchorDict.ContainsKey(k))
                    {
                        anchorDict[k] = (null, null);
                    }

                    // 绑定明细锚点 Range
                    if (clean.StartsWith(detPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        anchorDict[k] = (name.RefersToRange, anchorDict[k].sum);
                    }
                    // 绑定汇总锚点 Range
                    else if (clean.StartsWith(sumPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        anchorDict[k] = (anchorDict[k].det, name.RefersToRange);
                    }
                }

                // 按照箱柜序号正序遍历排序
                foreach (var kvp in anchorDict.OrderBy(x => x.Key))
                {
                    // 提取序号与 Range COM 对象
                    int cabinetIndex = kvp.Key;
                    dynamic detRange = kvp.Value.det;
                    dynamic sumRange = kvp.Value.sum;

                    // 若锚点不完整则跳过
                    if (detRange == null || sumRange == null) continue;

                    // 创建新的箱柜模型对象
                    var cab = new Models.CabinetObject
                    {
                        CabinetIndex = cabinetIndex,
                        DetAnchorRow = detRange.Row,
                        SumAnchorRow = sumRange.Row
                    };

                    // 读取顶部表头固定信息行
                    int headerRow = detRange.Row;
                    cab.Header.CabinetNo = Convert.ToString(sheet.Cells[headerRow, 2].Value) ?? $"箱柜{cabinetIndex}";
                    cab.Header.Model = Convert.ToString(sheet.Cells[headerRow, 4].Value) ?? string.Empty;
                    cab.Header.Name = Convert.ToString(sheet.Cells[headerRow, 6].Value) ?? string.Empty;

                    // 计算元器件范围
                    int compStartRow = detRange.Row + 2;
                    int compEndRow = sumRange.Row - 1;

                    // 循环读取中间元器件列表
                    int subIndex = 1;
                    for (int r = compStartRow; r <= compEndRow; r++)
                    {
                        // 提取元件名称
                        string compName = Convert.ToString(sheet.Cells[r, 2].Value) ?? string.Empty;

                        // 若整行无元件名称且无数量，视为插槽空行
                        if (string.IsNullOrWhiteSpace(compName)) continue;

                        // 创建元器件实体
                        var item = new Models.ComponentItem
                        {
                            Index = subIndex++,
                            Name = compName,
                            Specification = Convert.ToString(sheet.Cells[r, 3].Value) ?? string.Empty,
                            Manufacturer = Convert.ToString(sheet.Cells[r, 4].Value) ?? string.Empty,
                            Unit = Convert.ToString(sheet.Cells[r, 5].Value) ?? string.Empty
                        };

                        // 解析数量
                        if (decimal.TryParse(Convert.ToString(sheet.Cells[r, 6].Value), out decimal qty)) item.Quantity = qty;

                        // 解析销售单价
                        if (decimal.TryParse(Convert.ToString(sheet.Cells[r, 7].Value), out decimal price)) item.UnitPrice = price;

                        // 解析成本单价
                        if (decimal.TryParse(Convert.ToString(sheet.Cells[r, 10].Value), out decimal costPrice)) item.CostUnitPrice = costPrice;

                        // 添加至箱柜元器件集合
                        cab.Components.Add(item);
                    }

                    // 添加箱柜对象至工作表容器
                    sheetObj.Cabinets.Add(cab);
                }
            }
            catch (Exception ex)
            {
                // 记录解析工作表箱柜失败日志
                LogHelper.WriteLog($"解析工作表箱柜对象失败: {ex.Message}");
            }

            // 返回构建完成的工作表箱柜容器
            return sheetObj;
        }

        /// <summary>
        /// 从 Excel 工作表中反向解析指定序号的单个 CabinetObject 箱柜对象实体
        /// </summary>
        /// <param name="sheet">目标工作表 COM 对象</param>
        /// <param name="cabinetIndex">箱柜序号 k</param>
        /// <returns>反向解析生成的 CabinetObject 实例</returns>
        public static Models.CabinetObject ParseSingleCabinetObject(dynamic sheet, int cabinetIndex)
        {
            // 校验入参合法性
            if (sheet == null || cabinetIndex <= 0) return null;

            try
            {
                // 获取定义名称前缀配置 --硬编码--
                string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";

                // 构建定义名称标识 --硬编码--
                string sumTagName = $"{sumPrefix}{cabinetIndex}";
                string detTagName = $"{detPrefix}{cabinetIndex}";

                dynamic sumRange = null;
                dynamic detRange = null;

                // 遍历寻找箱柜对应的定义名称锚点
                foreach (dynamic name in sheet.Names)
                {
                    string clean = ExtractCleanNameStr(name.Name);
                    if (string.Equals(clean, detTagName, StringComparison.OrdinalIgnoreCase)) detRange = name.RefersToRange;
                    else if (string.Equals(clean, sumTagName, StringComparison.OrdinalIgnoreCase)) sumRange = name.RefersToRange;
                }

                // 校验锚点是否存在
                if (detRange == null || sumRange == null) return null;

                int detAnchorRow = detRange.Row;
                int sumAnchorRow = sumRange.Row;

                // 实例化箱柜对象
                var cab = new Models.CabinetObject
                {
                    CabinetIndex = cabinetIndex,
                    DetAnchorRow = detAnchorRow,
                    SumAnchorRow = sumAnchorRow
                };

                // 1. 反向解析 Header 固定表头 (明细表头位于 detAnchorRow + 3)
                int headerRow = detAnchorRow + 3;
                cab.Header.CabinetNo = Convert.ToString(sheet.Cells[headerRow, 2].Value) ?? $"箱柜{cabinetIndex}";
                cab.Header.Model = Convert.ToString(sheet.Cells[headerRow, 3].Value) ?? string.Empty;

                // 2. 反向解析元器件插槽列表 (compStartRow 至 compEndRow)
                int compStartRow = detAnchorRow + 5;
                int compEndRow = detAnchorRow + 26;
                int subIndex = 1;

                for (int r = compStartRow; r <= compEndRow; r++)
                {
                    string compName = Convert.ToString(sheet.Cells[r, 2].Value) ?? string.Empty;
                    // 过滤空白插槽行
                    if (string.IsNullOrWhiteSpace(compName)) continue;

                    var item = new Models.ComponentItem
                    {
                        Index = subIndex++,
                        Name = compName,
                        Specification = Convert.ToString(sheet.Cells[r, 3].Value) ?? string.Empty,
                        Manufacturer = Convert.ToString(sheet.Cells[r, 4].Value) ?? string.Empty,
                        Unit = Convert.ToString(sheet.Cells[r, 5].Value) ?? string.Empty
                    };

                    // 解析数量与单价数值
                    if (decimal.TryParse(Convert.ToString(sheet.Cells[r, 6].Value), out decimal qty)) item.Quantity = qty;
                    if (decimal.TryParse(Convert.ToString(sheet.Cells[r, 7].Value), out decimal price)) item.UnitPrice = price;
                    if (decimal.TryParse(Convert.ToString(sheet.Cells[r, 10].Value), out decimal costPrice)) item.CostUnitPrice = costPrice;

                    cab.Components.Add(item);
                }

                // 3. 反向解析计费调费项目
                int feeRow = sumAnchorRow;
                while (true)
                {
                    string feeName = Convert.ToString(sheet.Cells[feeRow, 2].Value) ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(feeName)) break;

                    var feeItem = new Models.BillingFeeItem
                    {
                        DisplayName = feeName,
                        ExcelFormula = Convert.ToString(sheet.Cells[feeRow, 8].Formula) ?? string.Empty
                    };

                    feeRow++;

                    // 避免越界读取，限制最多 10 行费用项 --硬编码--
                    if (feeRow - sumAnchorRow > 10) break;
                }

                // 返回反向解析成功的对象
                return cab;
            }
            catch (Exception ex)
            {
                // 记录错误日志
                LogHelper.WriteLog($"反向解析箱柜{cabinetIndex}对象失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 将修改后的箱柜 Header 属性同步写回 Excel 工作表中
        /// </summary>
        /// <param name="sheet">目标工作表</param>
        /// <param name="cabinet">箱柜对象</param>
        /// <returns>是否写回成功</returns>
        public static bool WriteCabinetHeaderToSheet(dynamic sheet, Models.CabinetObject cabinet)
        {
            // 校验入参
            if (sheet == null || cabinet == null || cabinet.DetAnchorRow <= 0) return false;

            try
            {
                // 定位表头绝对行号
                int headerRow = cabinet.DetAnchorRow;

                // 写回 B 列柜号
                sheet.Cells[headerRow, 2].Value = cabinet.Header.CabinetNo;

                // 写回 D 列型号
                sheet.Cells[headerRow, 4].Value = cabinet.Header.Model;

                // 写回 F 列名称
                sheet.Cells[headerRow, 6].Value = cabinet.Header.Name;

                // 返回成功
                return true;
            }
            catch (Exception ex)
            {
                // 记录错误日志
                LogHelper.WriteLog($"同步写回箱柜表头失败: {ex.Message}");
                return false;
            }
        }
    }
}


