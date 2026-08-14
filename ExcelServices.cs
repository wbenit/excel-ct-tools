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

                }
            }
            catch (Exception ex)
            {
                // 全局捕获异常防止程序闪退
                System.Windows.Forms.MessageBox.Show($"弹出新建项目窗口失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 新建项目初始化工作簿：完整回填【项目信息】与【分类1】工作表的数据、公式联动与定义名称锚点
        /// </summary>
        /// <param name="newWb">新创建的目标工作簿 COM 对象</param>
        /// <param name="model">前端提交的新建项目表单数据模型</param>
        public static void InitializeCreatedProjectWorkbook(dynamic newWb, Controllers.CreateProjectModel model)
        {
            if (newWb == null || model == null) return;

            try
            {
                // 1. 读取配置文件中的基准行号与前缀定义
                // 汇总行行号配置
                int cabSumRow = ConfigManager.Instance.Current.Excel.CabSumRowIndex;
                // 明细信息行行号配置
                int cabDetRow = ConfigManager.Instance.Current.Excel.CabDetRowIndex;
                // 总计行行号配置
                int cabTolsumRow = ConfigManager.Instance.Current.Excel.CabTolsumRowIndex;

                // 读取 4 种名称前缀配置
                string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";
                string subsumPrefix = ConfigManager.Instance.Current.Excel.SubsumNamePrefix ?? "Cab_Subsum_";
                string tolsumPrefix = ConfigManager.Instance.Current.Excel.TolsumNamePrefix ?? "Cab_Tolsum_";
                string defaultSheetName = ConfigManager.Instance.Current.Excel.DefaultTemplateSheet ?? "分类1";

                // 2. 填写【项目信息】工作表
                try
                {
                    dynamic infoSheet = newWb.Sheets["项目信息"];
                    if (infoSheet != null)
                    {
                        // 顶栏单位名称 (Row 1)
                        infoSheet.Range["B1"].Value = model.CompanyName;

                        // 【工程信息】区域填值 (Row 5 - Row 12)
                        infoSheet.Range["B5"].Value = model.ProjectName;      // 项目名称 (Cell B5)
                        infoSheet.Range["B6"].Value = model.ProjectRemark;    // 描述 (Cell B6)
                        infoSheet.Range["B7"].Value = model.QuoteNumber;      // 报价单号 (Cell B7)
                        infoSheet.Range["B8"].Value = model.Quoter;           // 报价人 (Cell B8)
                        infoSheet.Range["B9"].Value = model.ProjectDate;      // 创建日期 (Cell B9)
                        infoSheet.Range["B12"].Value = model.ProjectRemark;   // 项目备注 (Cell B12)

                        // 【客户信息】区域填值 (Row 14 - Row 17)
                        infoSheet.Range["B14"].Value = model.CustomerName;    // 客户名称 (Cell B14)
                        infoSheet.Range["B15"].Value = model.CustomerContact; // 联系人 (Cell B15)
                        infoSheet.Range["B16"].Value = model.CustomerPhone;   // 联系电话 (Cell B16)
                        infoSheet.Range["B17"].Value = model.CustomerAddress; // 客户地址 (Cell B17)

                        // 【本企业信息】区域填值 (Row 22 - Row 25)
                        infoSheet.Range["B22"].Value = model.CompanyName;    // 单位名称 (Cell B22)
                        infoSheet.Range["B23"].Value = model.EnglishName;     // 英文名称 (Cell B23)
                        infoSheet.Range["B24"].Value = model.CompanyContact;  // 联系人 (Cell B24)
                        infoSheet.Range["B25"].Value = model.CompanyPhone;    // 联系电话 (Cell B25)
                    }
                }
                catch (Exception exInfo)
                {
                    // 记录填写项目信息表的异常
                    LogHelper.WriteLog($"填写项目信息表异常: {exInfo.Message}");
                }

                // 3. 填写与初始化【分类1】工作表
                try
                {
                    dynamic catSheet = null;
                    try { catSheet = newWb.Sheets[defaultSheetName]; } catch { }
                    if (catSheet == null)
                    {
                        try { catSheet = newWb.Sheets["分类1"]; } catch { }
                    }

                    if (catSheet != null)
                    {
                        string sheetName = catSheet.Name;
                        int subsumRow = cabDetRow + 22; // 兜底小计行物理行号
                        int compStartRow = cabDetRow + 2; // 元器件起始行物理行号
                        int compEndRow = cabTolsumRow - 1; // 兜底元器件终止行物理行号

                        try
                        {
                            // 实例化公式调费控制器以获取当前默认公式组配置
                            var feeController = new Controllers.FormulaAdjustFeeController();
                            // 获取当前激活的默认公式组模型
                            var defaultGroup = feeController.GetDefaultGroup();
                            // 提取该公式组中的明细项目集合
                            var items = defaultGroup?.Details ?? new List<Controllers.FormulaItemModel>();

                            // 当明细项集合有效非空时，执行向上对齐覆盖
                            if (items.Count > 0)
                            {
                                int N = items.Count;
                                // 向上对齐总计行计算小计行物理行号: 小计行 = 总计行 - N + 1
                                subsumRow = cabTolsumRow - N + 1;
                                // 元器件区域终止物理行号: 元器件终止行 = 小计行 - 1
                                compEndRow = subsumRow - 1;

                                // 调用 Tool 公共工具方法构建 N 行 17 列的计费二维矩阵 (覆盖 A 列至 Q 列)
                                object[,] feeMatrix = Tool.BuildFeeMatrix(items, cabDetRow, subsumRow, compStartRow, compEndRow, 17);

                                // 将构建完成的计费二维矩阵一次性批量覆盖写入 Excel 计费区域 (规则 7)
                                dynamic feeRange = catSheet.Range[$"A{subsumRow}:Q{cabTolsumRow}"];
                                feeRange.Formula = feeMatrix;

                                // 为元器件区域 (compStartRow 到 compEndRow) 批量写入 A 列动态序号公式
                                int compRowCount = compEndRow - compStartRow + 1;
                                if (compRowCount > 0)
                                {
                                    object[,] compNoMatrix = new object[compRowCount, 1];
                                    for (int r = 0; r < compRowCount; r++)
                                    {
                                        compNoMatrix[r, 0] = $"=ROW()-ROW(A${cabDetRow + 1})";
                                    }
                                    catSheet.Range[$"A{compStartRow}:A{compEndRow}"].Formula = compNoMatrix;
                                }
                            }

                            // 绑定 4 个标准定义名称锚点 (规则 6)
                            newWb.Names.Add($"{sumPrefix}1", $"='{sheetName}'!$A${cabSumRow}");
                            newWb.Names.Add($"{detPrefix}1", $"='{sheetName}'!$A${cabDetRow}");
                            newWb.Names.Add($"{subsumPrefix}1", $"='{sheetName}'!$A${subsumRow}");
                            newWb.Names.Add($"{tolsumPrefix}1", $"='{sheetName}'!$A${cabTolsumRow}");
                        }
                        catch (Exception exNames)
                        {
                            LogHelper.WriteLog($"绑定分类1定义名称与写入计费矩阵异常: {exNames.Message}");
                        }

                        // 3.3 顶部汇总行 (cabSumRow) 公式与超链接联动
                        try
                        {
                            // A 列超链接跳转至明细信息行定义名称 (Cab_Det_1)，以便 OnSheetChange 识别标签触发双向同步
                            catSheet.Hyperlinks.Add(
                                Anchor: catSheet.Range[$"A{cabSumRow}"],
                                Address: "",
                                SubAddress: $"'{sheetName}'!{detPrefix}1",
                                TextToDisplay: "1"
                            );

                            // G 列单价公式指向明细总计行的销售总价 (H 列)
                            catSheet.Cells[cabSumRow, 7].Formula = $"=H{cabTolsumRow}";
                            // H 列总价公式 = 数量(F列) * 单价(G列)
                            catSheet.Cells[cabSumRow, 8].Formula = $"=F{cabSumRow}*G{cabSumRow}";
                            // J 列成本总价公式指向明细总计行的成本总价 (K 列)
                            catSheet.Cells[cabSumRow, 10].Formula = $"=K{cabTolsumRow}";
                            // K 列毛利公式 = 总价 - 成本总价
                            catSheet.Cells[cabSumRow, 11].Formula = $"=H{cabSumRow}-J{cabSumRow}";
                            // L 列毛利率公式
                            catSheet.Cells[cabSumRow, 12].Formula = $"=IF(H{cabSumRow}=0,0,K{cabSumRow}/H{cabSumRow})";
                        }
                        catch { }

                        // 3.4 底部明细信息行 (cabDetRow) 联动与超链接
                        try
                        {
                            // A 列超链接跳转回顶部汇总行定义名称 (Cab_Sum_1)，以便 OnSheetChange 识别标签触发双向同步
                            catSheet.Hyperlinks.Add(
                                Anchor: catSheet.Range[$"A{cabDetRow}"],
                                Address: "",
                                SubAddress: $"'{sheetName}'!{sumPrefix}1",
                                TextToDisplay: "柜号:"
                            );

                            // B 列初始箱柜名称填入默认值 (修改任意一处即通过 OnSheetChange 自动双向同步)
                            catSheet.Cells[cabDetRow, 2].Value = "箱柜1";
                            catSheet.Cells[cabSumRow, 2].Value = "箱柜1";

                        }
                        catch { }

                        // 3.7 激活分类1工作表为当前主视口
                        try
                        {
                            catSheet.Activate();
                            catSheet.Range["A1"].Select();
                        }
                        catch { }
                    }
                }
                catch (Exception exCat)
                {
                    // 记录填写分类表的异常日志
                    LogHelper.WriteLog($"初始化分类1工作表异常: {exCat.Message}");
                }
            }
            catch (Exception ex)
            {
                // 记录全局初始化异常
                LogHelper.WriteLog($"初始化新项目工作簿异常: {ex.Message}");
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
        /// 执行“公式法调费”逻辑: 解析公式表达式并精准更新写回 Excel 目标箱柜的费用行
        /// </summary>
        /// <param name="targetScope">调费作用域 (currentCabinet/currentCategory/allCabinets/selectedCabinet)</param>
        /// <param name="groupName">选中的公式组名称</param>
        /// <param name="items">前端编辑传递的公式明细项</param>
        public static void ApplyFormulaAdjustFeeToExcel(string targetScope, string groupName, System.Collections.Generic.List<Controllers.FormulaItemModel>? items = null)
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

                // 若前端未显式传递 items，则从控制器读取预置公式明细
                if (items == null || items.Count == 0)
                {
                    var controller = new Controllers.FormulaAdjustFeeController();
                    items = controller.GetFormulaDetails(groupName);
                }

                // 校验公式明细项有效性
                if (items == null || items.Count == 0) return;

                // 读取前缀配置项
                string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";

                // 扫描全表已绑定的箱柜定义名称
                var cabinetMap = new System.Collections.Generic.Dictionary<int, (dynamic det, dynamic sum)>();
                foreach (dynamic name in activeSheet.Names)
                {
                    string nName = Convert.ToString(name.Name) ?? "";
                    int k = ExtractIndexFromName(nName, sumPrefix, detPrefix);
                    if (k <= 0) continue;
                    if (!cabinetMap.ContainsKey(k)) cabinetMap[k] = (null, null);

                    string clean = nName.Contains("!") ? nName.Substring(nName.IndexOf("!") + 1) : nName;
                    clean = clean.Trim('\'', '=', ' ', '"');

                    if (clean.StartsWith(detPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        cabinetMap[k] = (name.RefersToRange, cabinetMap[k].sum);
                    }
                    else if (clean.StartsWith(sumPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        cabinetMap[k] = (cabinetMap[k].det, name.RefersToRange);
                    }
                }

                // 过滤出拥有完整 det 和 sum 的有效箱柜集合
                var validCabinets = cabinetMap.Where(x => x.Value.det != null && x.Value.sum != null)
                                            .OrderBy(x => (int)x.Value.sum.Row)
                                            .ToList();

                // 读取当前活动光标所在的物理行号
                dynamic activeCell = app.ActiveCell;
                int activeRow = activeCell != null ? Convert.ToInt32(activeCell.Row) : 0;

                // 筛选要执行更新的目标箱柜列表
                var targetCabinets = new System.Collections.Generic.List<(dynamic det, dynamic sum)>();

                // 若未识别到箱柜定义名称，启动智能降级兜底方案 (按标准模板结构自动构造物理行号)
                if (validCabinets.Count == 0)
                {
                    // 默认模板中顶部汇总行物理行号为 7
                    int fallbackSumRow = 7;

                    // 默认模板中下部明细区块起始物理行号为 41
                    int fallbackDetRow = 41;

                    // 自动尝试为当前工作表补充绑定 Cab_Sum_1 与 Cab_Det_1 定义名称锚点
                    try
                    {
                        // 获取当前工作表名称
                        string sheetName = activeSheet.Name;

                        // 绑定 Cab_Sum_1 锚点
                        activeWb.Names.Add($"{sumPrefix}1", $"='{sheetName}'!$A${fallbackSumRow}");

                        // 绑定 Cab_Det_1 锚点
                        activeWb.Names.Add($"{detPrefix}1", $"='{sheetName}'!$A${fallbackDetRow}");
                    }
                    catch { }

                    // 将智能构造的降级锚点加入更新目标列表
                    targetCabinets.Add((activeSheet.Range[$"A{fallbackDetRow}"], activeSheet.Range[$"A{fallbackSumRow}"]));
                }
                else if (targetScope == "currentCabinet")
                {
                    // 寻觅光标落点所在的箱柜
                    var matched = validCabinets.FirstOrDefault(c =>
                    {
                        int sumRow = c.Value.sum.Row;
                        int detRow = c.Value.det.Row;
                        return (activeRow == sumRow) || (activeRow >= detRow && activeRow <= detRow + 35);
                    });

                    if (matched.Key > 0)
                    {
                        targetCabinets.Add(matched.Value);
                    }
                    else
                    {
                        // 若光标不在箱柜内部，默认使用第一个有效箱柜
                        targetCabinets.Add(validCabinets.First().Value);
                    }
                }
                else
                {
                    // "allCabinets", "currentCategory", "selectedCabinet" 均覆盖更新当前全表所有箱柜
                    foreach (var kvp in validCabinets)
                    {
                        targetCabinets.Add(kvp.Value);
                    }
                }

                // 暂停屏刷提效
                bool prevUpdating = app.ScreenUpdating;
                app.ScreenUpdating = false;

                try
                {
                    // 遍历每一个目标箱柜，逐一渲染写回公式明细
                    foreach (var cab in targetCabinets)
                    {
                        int sumRow = cab.sum.Row;
                        int detRow = cab.det.Row;

                        // 找到小计行 (即明细块中元件部分之后的小计行物理行号)
                        int subtotalRow = detRow + 27;

                        // 写入明细块公式调费行
                        for (int i = 0; i < items.Count; i++)
                        {
                            var item = items[i];
                            int currentRow = subtotalRow + i;

                            // A 列写入序号 (若为 [序号] 占位符则转换为动态序号公式 =ROW()-ROW(A${detRow+1}))
                            if (item.No == "[序号]")
                            {
                                // 写入动态序号公式
                                activeSheet.Cells[currentRow, 1].Formula = $"=ROW()-ROW(A${detRow + 1})";
                            }
                            else if (!string.IsNullOrEmpty(item.No) && item.No.StartsWith("="))
                            {
                                // 写入自定义公式
                                activeSheet.Cells[currentRow, 1].Formula = item.No;
                            }
                            else
                            {
                                // 写入文本或数字值 (如 总计)
                                activeSheet.Cells[currentRow, 1].Value = item.No;
                            }

                            // B 列写入元件/费用名称
                            activeSheet.Cells[currentRow, 2].Value = item.Name;

                            // C 列 (只有用户填写了非空值才覆写项目内容)
                            if (!string.IsNullOrWhiteSpace(item.Model))
                            {
                                activeSheet.Cells[currentRow, 3].Value = item.Model;
                            }

                            // D 列 (只有用户填写了非空值才覆写项目内容)
                            if (!string.IsNullOrWhiteSpace(item.Manufacturer))
                            {
                                activeSheet.Cells[currentRow, 4].Value = item.Manufacturer;
                            }

                            // E 列单位
                            activeSheet.Cells[currentRow, 5].Value = item.Unit;

                            // F 列数量 (若以 = 开头写公式，否则写值)
                            if (!string.IsNullOrEmpty(item.Quantity))
                            {
                                if (item.Quantity.StartsWith("="))
                                    activeSheet.Cells[currentRow, 6].Formula = Tool.TransformFormulaRowOffset(item.Quantity, subtotalRow);
                                else
                                    activeSheet.Cells[currentRow, 6].Value = item.Quantity;
                            }

                            // G 列单价
                            if (!string.IsNullOrEmpty(item.Price))
                            {
                                if (item.Price.StartsWith("="))
                                    activeSheet.Cells[currentRow, 7].Formula = Tool.TransformFormulaRowOffset(item.Price, subtotalRow);
                                else
                                    activeSheet.Cells[currentRow, 7].Value = item.Price;
                            }

                            // H 列总价公式 (转换模板相对行号为当前箱柜实际物理行号)
                            if (!string.IsNullOrEmpty(item.TotalPriceFormula))
                            {
                                if (item.TotalPriceFormula.StartsWith("="))
                                {
                                    activeSheet.Cells[currentRow, 8].Formula = Tool.TransformFormulaRowOffset(item.TotalPriceFormula, subtotalRow);
                                }
                                else
                                {
                                    activeSheet.Cells[currentRow, 8].Value = item.TotalPriceFormula;
                                }
                            }

                            // J 列成本单价
                            if (!string.IsNullOrEmpty(item.CostPrice))
                            {
                                activeSheet.Cells[currentRow, 10].Value = item.CostPrice;
                            }

                            // K 列成本总价公式
                            if (!string.IsNullOrEmpty(item.CostTotalPriceFormula))
                            {
                                if (item.CostTotalPriceFormula.StartsWith("="))
                                {
                                    activeSheet.Cells[currentRow, 11].Formula = Tool.TransformFormulaRowOffset(item.CostTotalPriceFormula, subtotalRow);
                                }
                                else
                                {
                                    activeSheet.Cells[currentRow, 11].Value = item.CostTotalPriceFormula;
                                }
                            }

                            // Q 列类别 (列号 17)
                            if (!string.IsNullOrEmpty(item.Category))
                            {
                                activeSheet.Cells[currentRow, 17].Value = item.Category;
                            }
                        }

                        // 重新设置顶部汇总行 G 列与 H 列公式联动
                        activeSheet.Cells[sumRow, 7].Formula = $"=H{subtotalRow}";
                        activeSheet.Cells[sumRow, 8].Formula = $"=F{sumRow}*G{sumRow}";
                    }

                    // 刷新计算全表
                    activeSheet.Calculate();
                }
                finally
                {
                    // 恢复屏刷
                    app.ScreenUpdating = prevUpdating;
                }

                // 日志记录调费完成
                LogHelper.WriteLog($"成功完成公式法调费应用, 目标箱柜数: {targetCabinets.Count}, 作用域: {targetScope}");
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
        /// <summary>
        /// 点击“新建箱柜”业务逻辑：
        /// 1. 根据配置文件获取模板明细的标准总行数 A（TemplateDetailTotalRows）。
        /// 2. 顶部汇总行：若有空位置则复用，否则从模板行在选中位置上方插队插入。
        /// 3. 底部明细块：根据位置复制本表中的箱柜，根据 A 在元器件区域补齐或删除多余行。
        /// 4. 元器件区域智能清洗：采用内存二维数组读取，保留以 '=' 开头的计算公式，常量内容清洗为空。
        /// 5. 按规则 6 架构绑定 4 个标准定义名称及双向跳转超链接，遵循规则 7 内存批量读写。
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
                    // 1. 【配置文件获取包含总计行的模板明细总行数 A = CabTolsumRowIndex - CabDetRowIndex + 1】
                    // 读取箱柜信息行基准行号配置
                    int cabDetConfigRow = ConfigManager.Instance.Current.Excel.CabDetRowIndex;
                    // 读取总计行基准行号配置
                    int cabTolsumConfigRow = ConfigManager.Instance.Current.Excel.CabTolsumRowIndex;
                    // 动态计算得出包含总计行的模板明细总行数 A (如 68 - 44 + 1 = 25 行)
                    int templateTotalRowsA = cabTolsumConfigRow >= cabDetConfigRow ? (cabTolsumConfigRow - cabDetConfigRow + 1) : 25;

                    // 读取顶部汇总行基准行号配置 (CabSumRowIndex)
                    int headerRowIndex = ConfigManager.Instance.Current.Excel.CabSumRowIndex;

                    // 规则 7：采用数组一次性读到内存扫描 B 列，定位顶部汇总表“第一个空白汇总行”
                    int maxScanRows = 500; // --硬编码: 预估单表最多 100 个箱柜汇总行--
                    // 获取 B 列扫描区域 Range
                    dynamic bColRange = activeSheet.Range[$"B{headerRowIndex}:B{headerRowIndex + maxScanRows - 1}"];
                    // 批量读取 B 列内容为二维数组
                    object[,] bValues = bColRange.Value2 as object[,];
                    // 记录首个空汇总行行号
                    int firstEmptySummaryRow = headerRowIndex;
                    if (bValues != null)
                    {
                        // 遍历读取的二维内存数组
                        int rowCount = bValues.GetLength(0);
                        for (int i = 1; i <= rowCount; i++)
                        {
                            // 提取 B 列内容
                            string bVal = Convert.ToString(bValues[i, 1]) ?? "";
                            // 若发现 B 列为空，说明找到紧凑连续的第一个空汇总行号
                            if (string.IsNullOrWhiteSpace(bVal))
                            {
                                firstEmptySummaryRow = headerRowIndex + i - 1;
                                break;
                            }
                        }
                    }

                    // 2. 提取当前活动焦点的实际行号 cRow
                    int cRow = activeCell != null ? Convert.ToInt32(activeCell.Row) : headerRowIndex;

                    // 读取焦点当前行 B 列的值
                    string selectedNameInB = Convert.ToString(activeSheet.Cells[cRow, 2].Value2) ?? Convert.ToString(activeSheet.Cells[cRow, 2].Value) ?? "";
                    // 判断当前行是否为空行
                    bool isSelectedRowEmpty = string.IsNullOrWhiteSpace(selectedNameInB);

                    // 判断是否在已有箱柜的非空汇总行中插队
                    if (cRow >= headerRowIndex && !isSelectedRowEmpty)
                    {
                        // 选中的是已有箱柜的非空行 -> 在选中行位置上方插队插入一行
                        insertRow = cRow;
                        // 复制基准汇总行格式与公式
                        activeSheet.Rows[headerRowIndex].Copy();
                        // 插入新汇总行
                        activeSheet.Rows[insertRow].Insert(-4121);
                    }
                    else
                    {
                        // 选中的是空行 -> 强制收拢至紧凑的第一个空白汇总行 firstEmptySummaryRow，绝对不产生中断空行！
                        insertRow = firstEmptySummaryRow;
                    }

                    // 3. 动态提取全表中下一个全新的箱柜序号 cabinetK (保障原有的定义名称完全保留不被删除/覆盖)
                    Microsoft.Office.Interop.Excel.Worksheet excelSheetForK = (Microsoft.Office.Interop.Excel.Worksheet)activeSheet;
                    // 获取所属 Workbook 对象
                    Microsoft.Office.Interop.Excel.Workbook excelWbForK = (Microsoft.Office.Interop.Excel.Workbook)excelSheetForK.Parent;

                    // 获取下一个独立增量序号 K
                    int cabinetK = GetNextCabinetIndex(excelWbForK, excelSheetForK);

                    // 生成新箱柜名称 (如 箱柜3)
                    string cabinetName = $"箱柜{cabinetK}";

                    // 4. 填入顶部汇总行的数据 (保留 =ROW()-ROW(A$6) 序号公式)
                    // 在 B 列写入箱柜名称
                    activeSheet.Cells[insertRow, 2].Value = cabinetName;
                    // 在 E 列写入单位“台”
                    activeSheet.Cells[insertRow, 5].Value = "台";
                    // 在 F 列写入默认数量 1
                    activeSheet.Cells[insertRow, 6].Value = 1;

                    // 读取 4 种定义名称前缀配置
                    string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                    string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";
                    string subsumPrefix = ConfigManager.Instance.Current.Excel.SubsumNamePrefix ?? "Cab_Subsum_";
                    string tolsumPrefix = ConfigManager.Instance.Current.Excel.TolsumNamePrefix ?? "Cab_Tolsum_";

                    // 获取当前活动焦点的实际行号 activeRow
                    int activeRow = 7;
                    if (activeCell != null)
                    {
                        activeRow = Convert.ToInt32(activeCell.Row);
                    }

                    // 读取当前活动行 B 列的值
                    string activeNameInB = Convert.ToString(activeSheet.Cells[activeRow, 2].Value2) ?? Convert.ToString(activeSheet.Cells[activeRow, 2].Value) ?? "";
                    // 判断活动行 B 列是否为空
                    bool isRowBEmpty = string.IsNullOrWhiteSpace(activeNameInB);

                    // 声明动态计算的复制源起始行与终止行，目标插入位置
                    // 兜底默认值由配置动态计算：起始行 = CabDetRowIndex - 3，终止行 = CabTolsumRowIndex
                    // 保证在没有找到任何已有箱柜时，默认复制模板箱柜的完整行块（含总计行共 A 行）
                    int copyStartRow = cabDetConfigRow - 3;
                    int copyEndRow = cabTolsumConfigRow;
                    int targetDetailRow = activeRow;

                    // 扫描全表已绑定的定义名称锚点字典 (Index -> (DetRange, SumRange, SubsumRange, TolsumRange))
                    var nameDict = new Dictionary<int, (dynamic det, dynamic sum, dynamic subsum, dynamic tolsum)>();
                    // 遍历工作表中所有定义名称
                    foreach (dynamic name in activeSheet.Names)
                    {
                        // 清洗提取定义名称字符串
                        string clean = ExtractCleanNameStr(name.Name);
                        // 提取箱柜数字序号
                        int k = ExtractIndexFromName(clean, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                        if (k <= 0) continue;
                        // 初始化字典元组
                        if (!nameDict.ContainsKey(k)) nameDict[k] = (null, null, null, null);

                        // 匹配并存储 Det 锚点
                        if (clean.StartsWith(detPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            nameDict[k] = (name.RefersToRange, nameDict[k].sum, nameDict[k].subsum, nameDict[k].tolsum);
                        }
                        // 匹配并存储 Sum 锚点
                        else if (clean.StartsWith(sumPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            nameDict[k] = (nameDict[k].det, name.RefersToRange, nameDict[k].subsum, nameDict[k].tolsum);
                        }
                        // 匹配并存储 Subsum 锚点
                        else if (clean.StartsWith(subsumPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            nameDict[k] = (nameDict[k].det, nameDict[k].sum, name.RefersToRange, nameDict[k].tolsum);
                        }
                        // 匹配并存储 Tolsum 锚点
                        else if (clean.StartsWith(tolsumPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            nameDict[k] = (nameDict[k].det, nameDict[k].sum, nameDict[k].subsum, name.RefersToRange);
                        }
                    }

                    // 按照顶部汇总行号 (sum.Row) 物理行号升序排序有效箱柜
                    var validCabinets = nameDict.Where(x => x.Value.det != null && x.Value.sum != null)
                                                .OrderBy(x => (int)x.Value.sum.Row)
                                                .ToList();

                    // 声明选中的复制源箱柜结构
                    (dynamic det, dynamic sum, dynamic subsum, dynamic tolsum) srcCabinet = (null, null, null, null);

                    // 5. 【根据位置复制本表中的箱柜】
                    // 场景 1：当前活动行的 B 列为空 (末尾追加)
                    if (isRowBEmpty)
                    {
                        if (validCabinets.Count > 0)
                        {
                            // 复制源选择本表中最后一个箱柜
                            srcCabinet = validCabinets.Last().Value;
                            int detRowK = srcCabinet.det.Row;

                            // 复制源起始行：箱柜信息行上方 3 行
                            copyStartRow = detRowK - 3;
                            // 复制源终止行：取总计行 Cab_Tolsum
                            int tolsumRowK = srcCabinet.tolsum != null ? Convert.ToInt32(srcCabinet.tolsum.Row) : (detRowK + 24);
                            copyEndRow = tolsumRowK;
                            // 安全范围校验
                            if (copyEndRow < copyStartRow) copyEndRow = copyStartRow + templateTotalRowsA + 2;
                        }
                        // 目标插入位置：已使用行的下一行
                        targetDetailRow = activeSheet.UsedRange.Row + activeSheet.UsedRange.Rows.Count;
                    }
                    // 场景 2：当前活动行的 B 列不为空 (中间插队)
                    else
                    {
                        // 寻找插队位置的上一个箱柜 l 与下一个箱柜 m
                        (dynamic det, dynamic sum, dynamic subsum, dynamic tolsum) cabL = (null, null, null, null);
                        (dynamic det, dynamic sum, dynamic subsum, dynamic tolsum) cabM = (null, null, null, null);

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

                        // 复制源优先选取上一个箱柜 cabL，其次选用 cabM
                        srcCabinet = cabL.det != null ? cabL : cabM;

                        if (srcCabinet.det != null)
                        {
                            int srcDet = srcCabinet.det.Row;
                            // 复制源起始行：箱柜信息行上方 3 行
                            copyStartRow = srcDet - 3;
                            // 复制源终止行：取总计行 Cab_Tolsum
                            int srcTol = srcCabinet.tolsum != null ? Convert.ToInt32(srcCabinet.tolsum.Row) : (srcDet + 24);
                            copyEndRow = srcTol;
                            // 安全范围校验
                            if (copyEndRow < copyStartRow) copyEndRow = copyStartRow + templateTotalRowsA + 2;
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

                    // 构造动态复制范围文本 (例如 "41:68")
                    string copyRangeText = $"{copyStartRow}:{copyEndRow}";

                    // 从当前工作表复制计算出的物理行区块并整块插入到目标位置
                    activeSheet.Range[copyRangeText].Copy();
                    activeSheet.Rows[targetDetailRow].Insert(-4121);
                    // 清空剪贴板选中高亮
                    app.CutCopyMode = (Microsoft.Office.Interop.Excel.XlCutCopyMode)0;

                    // 6. 【根据 A 补齐或删除多余行，使元器件总行与计费行之和与 A 一致】
                    // 插入后新箱柜的明细块起始行即为 targetDetailRow
                    int newStartRow = targetDetailRow;
                    // 新箱柜信息行 Cab_Det (明细起始行 + 3)
                    int newDetRow = newStartRow + 3;
                    // 规则 6：Cab_Det_k.Row + 2 为元器件起始行
                    int newCompStartRow = newDetRow + 2;

                    // 获取源箱柜从 Cab_Det 到 Cab_Tolsum 的总行数 B
                    int srcDetRow = srcCabinet.det != null ? Convert.ToInt32(srcCabinet.det.Row) : (copyStartRow + 3);
                    int srcTolsumRow = srcCabinet.tolsum != null ? Convert.ToInt32(srcCabinet.tolsum.Row) : copyEndRow;
                    int srcTotalRowsInDetBlockB = srcTolsumRow - srcDetRow + 1;

                    // 获取源箱柜的计费区域行数 (小计行到总计行)
                    int feeRowCount = 4; // 默认 4 行计费项
                    if (srcCabinet.tolsum != null && srcCabinet.subsum != null)
                    {
                        feeRowCount = Convert.ToInt32(srcCabinet.tolsum.Row) - Convert.ToInt32(srcCabinet.subsum.Row) + 1;
                    }
                    // 限制计费行数有效范围
                    if (feeRowCount < 1) feeRowCount = 4;

                    // 计算复制插入后新箱柜在元器件区域中的当前行数
                    int srcSubsumRow = srcCabinet.subsum != null ? Convert.ToInt32(srcCabinet.subsum.Row) : (srcTolsumRow - feeRowCount + 1);
                    int currentCompRowCount = srcSubsumRow - (srcDetRow + 2);
                    if (currentCompRowCount < 1) currentCompRowCount = 1;

                    // 计算为对齐总行数 A 所需的目标元器件行数 (A - 2 - feeRowCount，因为 Det 占 1 行，表头占 1 行)
                    int targetCompRowCount = templateTotalRowsA - 2 - feeRowCount;
                    if (targetCompRowCount < 1) targetCompRowCount = 1;

                    // 比较源箱柜包含总计行的总行数 B 与模板标准总行数 A
                    if (srcTotalRowsInDetBlockB > templateTotalRowsA)
                    {
                        // B > A: 源箱柜元器件行数过多，需在元器件区域尾部删除多余行 (B - A)
                        int diffRows = srcTotalRowsInDetBlockB - templateTotalRowsA;
                        // 删除起始行：元器件目标行数之后第一行
                        int deleteStartRow = newCompStartRow + targetCompRowCount;
                        int deleteEndRow = deleteStartRow + diffRows - 1;

                        // ★ 安全边界保护：计费区（Subsum 行到 Tolsum 行）紧跟在元器件区之后，
                        //   绝对不能删到计费区！计费区在新箱柜中的起始行等于
                        //   newDetRow + (B - feeRowCount)，即元器件区（含表头）之后。
                        //   删除终止行必须 <= 该起始行 - 1，否则会导致计费区文字丢失。
                        int feeAreaNewStartRow = newDetRow + srcTotalRowsInDetBlockB - feeRowCount;
                        if (deleteEndRow >= feeAreaNewStartRow)
                        {
                            // 修正终止行，严格保护计费区不被删除
                            deleteEndRow = feeAreaNewStartRow - 1;
                        }

                        // 确认删除范围有效后执行
                        if (deleteEndRow >= deleteStartRow)
                        {
                            activeSheet.Rows[$"{deleteStartRow}:{deleteEndRow}"].Delete();
                        }
                    }
                    else if (srcTotalRowsInDetBlockB < templateTotalRowsA)
                    {
                        // B < A: 源箱柜元器件行数不足，在元器件区域尾部插入补齐行数 (A - B)
                        int diffRows = templateTotalRowsA - srcTotalRowsInDetBlockB;
                        // 插入位置：当前元器件区末尾（紧贴计费区上方），保证计费区向下平移
                        int insertStartRow = newCompStartRow + currentCompRowCount;
                        int insertEndRow = insertStartRow + diffRows - 1;
                        // 插入对应数量的空行（计费区自动随之下移，内容不丢失）
                        activeSheet.Rows[$"{insertStartRow}:{insertEndRow}"].Insert(-4121);
                        // 从元器件样本行复制格式到新插入的行
                        activeSheet.Rows[newCompStartRow].Copy();
                        activeSheet.Rows[$"{insertStartRow}:{insertEndRow}"].PasteSpecial(-4122); // xlPasteFormats
                        app.CutCopyMode = (Microsoft.Office.Interop.Excel.XlCutCopyMode)0;
                    }

                    // 7. 计算行数对齐后的新箱柜各锚点物理行号 (包含总计行从 newDetRow 到 newTolsumRow 精确为 A 行)
                    // 新箱柜总计行 Cab_Tolsum_k.Row
                    int newTolsumRow = newDetRow + templateTotalRowsA - 1;
                    // 新箱柜小计行 Cab_Subsum_k.Row
                    int newSubsumRow = newTolsumRow - feeRowCount + 1;
                    // 规则 6：Cab_Subsum_k.Row - 1 为元器件终止行
                    int newCompEndRow = newSubsumRow - 1;

                    // 8. 【将元器件部分的内容清洗，公式不清洗 (规则 7 内存批量读写)】
                    // 定位元器件区域 Range (覆盖 A 列至 Q 列)
                    dynamic compRange = activeSheet.Range[$"A{newCompStartRow}:Q{newCompEndRow}"];
                    // 批量读取元器件区域所有公式矩阵
                    object[,] compFormulas = compRange.Formula as object[,];

                    if (compFormulas != null)
                    {
                        int compRCount = compFormulas.GetLength(0);
                        int compCCount = compFormulas.GetLength(1);
                        // 构建清洗后的二维数据矩阵
                        object[,] cleanedMatrix = new object[compRCount, compCCount];
                        // 动态序号公式基准行号
                        int baseHeaderRow = newDetRow + 1;

                        for (int r = 1; r <= compRCount; r++)
                        {
                            for (int c = 1; c <= compCCount; c++)
                            {
                                // A 列 (c == 1)：序号列，统一重置为动态相对序号公式
                                if (c == 1)
                                {
                                    cleanedMatrix[r - 1, c - 1] = $"=ROW()-ROW(A${baseHeaderRow})";
                                    continue;
                                }

                                // 提取该单元格的公式字符串
                                string cellFormula = Convert.ToString(compFormulas[r, c]) ?? string.Empty;

                                // 判断是否为以 '=' 开头的计算公式 (例如 H列销售总价 =F*G, K列成本总价 =F*J 等)
                                if (cellFormula.StartsWith("="))
                                {
                                    // 属于计算公式：原样保留，绝不清洗！
                                    cleanedMatrix[r - 1, c - 1] = cellFormula;
                                }
                                else
                                {
                                    // 属于常量数据（元件名称、规格型号、厂家、数量、单价等）：清洗置空！
                                    cleanedMatrix[r - 1, c - 1] = string.Empty;
                                }
                            }
                        }

                        // 将清洗后的矩阵一次性批量写回 Excel 元器件区域 (规则 7)
                        compRange.Formula = cleanedMatrix;
                    }

                    // 9. 更新箱柜信息行与顶部汇总行数据
                    // 底部明细箱柜信息行 (B 列) 写入新箱柜名称
                    activeSheet.Cells[newDetRow, 2].Value2 = cabinetName;

                    // 顶部汇总行设置正确的公式与联动
                    // G 列单价公式指向明细总计行的销售总价 (H 列)
                    activeSheet.Cells[insertRow, 7].Formula = $"=H{newTolsumRow}";
                    // H 列总价公式 = 数量(F列) * 单价(G列)
                    activeSheet.Cells[insertRow, 8].Formula = $"=F{insertRow}*G{insertRow}";
                    // J 列成本总价公式指向明细总计行的成本总价 (K 列)
                    activeSheet.Cells[insertRow, 10].Formula = $"=K{newTolsumRow}";
                    // K 列毛利公式 = 总价 - 成本总价
                    activeSheet.Cells[insertRow, 11].Formula = $"=H{insertRow}-J{insertRow}";
                    // L 列毛利率公式
                    activeSheet.Cells[insertRow, 12].Formula = $"=IF(H{insertRow}=0,0,K{insertRow}/H{insertRow})";

                    // 10. 绑定规则 6 要求的 4 个定义名称 (Cab_Sum_, Cab_Det_, Cab_Subsum_, Cab_Tolsum_)
                    string sumNameTag = $"{sumPrefix}{cabinetK}";
                    string detNameTag = $"{detPrefix}{cabinetK}";
                    string subsumNameTag = $"{subsumPrefix}{cabinetK}";
                    string tolsumNameTag = $"{tolsumPrefix}{cabinetK}";

                    // 将 activeSheet 转为强类型 Worksheet
                    Microsoft.Office.Interop.Excel.Worksheet excelActiveSheet = (Microsoft.Office.Interop.Excel.Worksheet)activeSheet;
                    // 从当前工作表所属的 Parent 获取确切的目标工作簿句柄
                    Microsoft.Office.Interop.Excel.Workbook targetWb = (Microsoft.Office.Interop.Excel.Workbook)excelActiveSheet.Parent;

                    // 确保目标工作簿与工作表处于激活状态
                    try { targetWb.Activate(); } catch { }
                    try { excelActiveSheet.Activate(); } catch { }

                    // 获取 4 个锚点单元格 Range 对象
                    Microsoft.Office.Interop.Excel.Range sumAnchorCell = (Microsoft.Office.Interop.Excel.Range)excelActiveSheet.Cells[insertRow, 1];
                    Microsoft.Office.Interop.Excel.Range detAnchorCell = (Microsoft.Office.Interop.Excel.Range)excelActiveSheet.Cells[newDetRow, 1];
                    Microsoft.Office.Interop.Excel.Range subsumAnchorCell = (Microsoft.Office.Interop.Excel.Range)excelActiveSheet.Cells[newSubsumRow, 1];
                    Microsoft.Office.Interop.Excel.Range tolsumAnchorCell = (Microsoft.Office.Interop.Excel.Range)excelActiveSheet.Cells[newTolsumRow, 1];

                    try
                    {
                        // 强类型 Add，注册规则 6 要求的 4 个定义名称
                        targetWb.Names.Add(Name: sumNameTag, RefersTo: sumAnchorCell, Visible: true);
                        targetWb.Names.Add(Name: detNameTag, RefersTo: detAnchorCell, Visible: true);
                        targetWb.Names.Add(Name: subsumNameTag, RefersTo: subsumAnchorCell, Visible: true);
                        targetWb.Names.Add(Name: tolsumNameTag, RefersTo: tolsumAnchorCell, Visible: true);

                        // 强制刷新 Excel 计算引擎
                        app.CalculateFull();

                        // 为顶部 A列单元格添加指向底部定义名称的超链接
                        excelActiveSheet.Hyperlinks.Add(
                            Anchor: sumAnchorCell,
                            Address: "",
                            SubAddress: $"'{excelActiveSheet.Name}'!{detNameTag}",
                            ScreenTip: "跳转至明细块"
                        );
                        // 为底部 A列单元格添加指向顶部定义名称的超链接
                        excelActiveSheet.Hyperlinks.Add(
                            Anchor: detAnchorCell,
                            Address: "",
                            SubAddress: $"'{excelActiveSheet.Name}'!{sumNameTag}",
                            ScreenTip: "返回汇总行"
                        );
                    }
                    catch (Exception nameEx)
                    {
                        // 记录名称创建异常
                        LogHelper.WriteLog($"创建定义名称异常: {nameEx.Message}");
                    }

                    // 自动注册工作表修改双向同步事件
                    RegisterSheetChangeEvent();

                    // 11. 焦点定位：激活原工作表并安全选中新插入的单元格
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

            // 【配置文件替代硬编码列举】
            // 1. 原硬编码前缀: "Cab_Sum_", "Cab_Det_", "Cab_Subsum_", "Cab_Tolsum_"
            // 2. 替代配置项: ConfigManager.Instance.Current.Excel.SumNamePrefix / DetNamePrefix / SubsumNamePrefix / TolsumNamePrefix
            string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
            string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";
            string subsumPrefix = ConfigManager.Instance.Current.Excel.SubsumNamePrefix ?? "Cab_Subsum_";
            string tolsumPrefix = ConfigManager.Instance.Current.Excel.TolsumNamePrefix ?? "Cab_Tolsum_";

            try
            {
                // 1. 扫描当前工作簿中所有的工作簿级定义名称，提取最大序号 K
                if (targetWb != null && targetWb.Names != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Name n in targetWb.Names)
                    {
                        string nName = Convert.ToString(n.Name) ?? "";
                        int k = ExtractIndexFromName(nName, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                        if (k > maxK) maxK = k;
                    }
                }

                // 2. 扫描当前工作表中所有的工作表级定义名称，提取最大序号 K
                if (activeSheet != null && activeSheet.Names != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Name n in activeSheet.Names)
                    {
                        string nName = Convert.ToString(n.Name) ?? "";
                        int k = ExtractIndexFromName(nName, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                        if (k > maxK) maxK = k;
                    }
                }
            }
            catch { }
            return maxK + 1;
        }

        /// <summary>
        /// 从定义名称全名中安全解析提取箱柜序号数字 (支持 Cab_Sum_ / Cab_Det_ / Cab_Subsum_ / Cab_Tolsum_)
        /// </summary>
        private static int ExtractIndexFromName(string fullName, string sumPrefix = null, string detPrefix = null, string subsumPrefix = null, string tolsumPrefix = null)
        {
            if (string.IsNullOrWhiteSpace(fullName)) return 0;

            // 提取或回退默认前缀
            sumPrefix = sumPrefix ?? ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
            detPrefix = detPrefix ?? ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";
            subsumPrefix = subsumPrefix ?? ConfigManager.Instance.Current.Excel.SubsumNamePrefix ?? "Cab_Subsum_";
            tolsumPrefix = tolsumPrefix ?? ConfigManager.Instance.Current.Excel.TolsumNamePrefix ?? "Cab_Tolsum_";

            // 清理可能存在的工作表前缀与单引号/等号 (例如 ='分类1'!Cab_Sum_2 -> Cab_Sum_2)
            string cleanName = fullName;
            if (cleanName.Contains("!"))
            {
                cleanName = cleanName.Substring(cleanName.IndexOf("!") + 1);
            }
            cleanName = cleanName.Trim('\'', '=', ' ', '"');

            // 遍历 4 个前缀进行匹配提取序号
            string[] prefixes = new[] { sumPrefix, detPrefix, subsumPrefix, tolsumPrefix };
            foreach (var prefix in prefixes)
            {
                if (!string.IsNullOrEmpty(prefix) && cleanName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    string numStr = cleanName.Substring(prefix.Length);
                    if (int.TryParse(numStr, out int k)) return k;
                }
            }

            return 0;
        }

        /// <summary>
        /// 将面向对象实体 CabinetObject 完整渲染写回 Excel 工作表中 (包括 Header、元器件列表与底部调费策略)
        /// 遵循规则 6（行号结构与空行/插入行规则）及规则 7（内存二维数组批量读写）
        /// </summary>
        /// <param name="sheet">目标工作表 COM 对象</param>
        /// <param name="cabinet">箱柜面向对象实体</param>
        /// <param name="insertRow">顶部汇总行物理行号</param>
        /// <param name="targetDetailRow">下部明细区块物理起始行号</param>
        /// <param name="templateBlankRows">模板空行总数（默认 23 行）</param>
        /// <returns>是否渲染写回成功</returns>
        public static bool RenderCabinetObjectToSheet(dynamic sheet, Models.CabinetObject cabinet, int insertRow, int targetDetailRow, int templateBlankRows = 23)
        {
            // 校验输入参数合法性
            if (sheet == null || cabinet == null || insertRow <= 0 || targetDetailRow <= 0) return false;

            try
            {
                // 获取 Excel 全局 Application 实例
                dynamic app = sheet.Application;
                bool prevUpdating = app.ScreenUpdating;
                app.ScreenUpdating = false;

                // 1. 定位箱柜信息行 Cab_Det_k.Row
                int detRow = targetDetailRow + 3;
                cabinet.DetAnchorRow = detRow;
                cabinet.SumAnchorRow = insertRow;

                // 2. 渲染底部明细表头箱柜名称 (B 列)
                sheet.Cells[detRow, 2].Value2 = cabinet.Header.CabinetNo;

                // 3. 动态获取计费策略公式行定义列表
                var rowDefs = new List<Models.FormulaFeeRowDefinition>();
                if (cabinet.BillingStrategy is Models.FormulaBillingGroupStrategy fs && fs.RowDefinitions != null)
                {
                    rowDefs = fs.RowDefinitions;
                }
                int feeRowCount = rowDefs.Count;

                // 4. 根据规则 6：Cab_Det_k.Row + 2 为元器件起始行
                int compStartRow = detRow + 2;

                // 计算模板默认预留的元器件行数
                int defaultCompRowCount = Math.Max(1, templateBlankRows - feeRowCount);

                // 判定实际元器件列表数量
                int compCount = cabinet.Components != null ? cabinet.Components.Count : 0;
                int compRowCount = Math.Max(defaultCompRowCount, compCount);

                // 规则 6：“如果元器件数量多于区域行数，先要插入行”
                if (compCount > defaultCompRowCount)
                {
                    // 计算需要插入的差额行数
                    int insertLineCount = compCount - defaultCompRowCount;
                    int insertStartRow = compStartRow + defaultCompRowCount;
                    sheet.Rows[$"{insertStartRow}:{insertStartRow + insertLineCount - 1}"].Insert(-4121);
                }

                // 规则 6：Cab_Subsum_k.Row - 1 为元器件终止行
                int compEndRow = compStartRow + compRowCount - 1;

                // 规则 6：Cab_Subsum_k.Row 为小计行
                int subsumRow = compEndRow + 1;
                cabinet.SubsumAnchorRow = subsumRow;

                // 规则 6：Cab_Subsum_k.Row 到 Cab_Tolsum_k.Row - 1 为计费区域，Cab_Tolsum_k.Row 为总计行
                int tolsumRow = feeRowCount > 0 ? subsumRow + feeRowCount - 1 : subsumRow;
                cabinet.TolsumAnchorRow = tolsumRow;

                // 5. 规则 7：元器件区域采用二维数组一次性批量写入内存与 Excel
                int totalCompCols = 11; // A 到 K 共 11 列
                object[,] compArray = new object[compRowCount, totalCompCols];
                int baseHeaderRow = compStartRow - 1;

                for (int i = 0; i < compRowCount; i++)
                {
                    // A 列 (索引 0)：写入动态相对序号公式
                    compArray[i, 0] = $"=ROW()-ROW(A${baseHeaderRow})";

                    if (cabinet.Components != null && i < cabinet.Components.Count)
                    {
                        var comp = cabinet.Components[i];
                        // B 列 (索引 1)：元件名称
                        compArray[i, 1] = comp.Name ?? string.Empty;
                        // C 列 (索引 2)：规格型号
                        compArray[i, 2] = comp.Specification ?? string.Empty;
                        // D 列 (索引 3)：生产厂家
                        compArray[i, 3] = comp.Manufacturer ?? string.Empty;
                        // E 列 (索引 4)：单位
                        compArray[i, 4] = comp.Unit ?? string.Empty;
                        // F 列 (索引 5)：数量
                        compArray[i, 5] = comp.Quantity > 0 ? (object)comp.Quantity : string.Empty;
                        // G 列 (索引 6)：销售单价
                        compArray[i, 6] = comp.UnitPrice > 0 ? (object)comp.UnitPrice : string.Empty;
                        // J 列 (索引 9)：成本单价
                        compArray[i, 9] = comp.CostUnitPrice > 0 ? (object)comp.CostUnitPrice : string.Empty;
                    }
                    else
                    {
                        // 预留空白插槽清洗
                        compArray[i, 1] = string.Empty;
                        compArray[i, 2] = string.Empty;
                        compArray[i, 3] = string.Empty;
                        compArray[i, 4] = string.Empty;
                        compArray[i, 5] = string.Empty;
                        compArray[i, 6] = string.Empty;
                        compArray[i, 9] = string.Empty;
                    }
                }

                // 将构建完成的元器件二维数组一次性批量赋值写入 Range (规则 7)
                dynamic compRange = sheet.Range[$"A{compStartRow}:K{compEndRow}"];
                compRange.Formula = compArray;

                // 6. 规则 7：计费区域（从 Cab_Subsum_k.Row 至 Cab_Tolsum_k.Row）批量写入
                if (feeRowCount > 0)
                {
                    object[,] feeArray = new object[feeRowCount, totalCompCols];

                    for (int j = 0; j < feeRowCount; j++)
                    {
                        var rowDef = rowDefs[j];
                        // A 列 (索引 0)：[序号] 转换为 =ROW()-ROW(A${baseHeaderRow})，总计行直接写入“总计”
                        if (rowDef.Name == "总计" || rowDef.IndexTag == "总计")
                        {
                            // 总计行标记为“总计”
                            feeArray[j, 0] = "总计";
                        }
                        else
                        {
                            // 计费项序号公式：=ROW()-ROW(A${baseHeaderRow})，其中 baseHeaderRow 为 Cab_Det.row + 1
                            feeArray[j, 0] = $"=ROW()-ROW(A${baseHeaderRow})";
                        }

                        // B 列 (索引 1)：费用项名称
                        feeArray[j, 1] = rowDef.Name ?? string.Empty;

                        // H 列 (索引 7)：销售总价计算公式转换
                        if (!string.IsNullOrWhiteSpace(rowDef.TotalPriceFormula))
                        {
                            feeArray[j, 7] = Models.FormulaEngine.ConvertToExcelFormula(
                                rowDef.TotalPriceFormula,
                                1,
                                subsumRow,
                                compStartRow,
                                compEndRow
                            );
                        }

                        // K 列 (索引 10)：成本总价计算公式转换
                        if (!string.IsNullOrWhiteSpace(rowDef.CostTotalPriceFormula))
                        {
                            feeArray[j, 10] = Models.FormulaEngine.ConvertToExcelFormula(
                                rowDef.CostTotalPriceFormula,
                                1,
                                subsumRow,
                                compStartRow,
                                compEndRow
                            );
                        }
                    }

                    // 将计费公式数组一次性批量写入计费区域 (规则 7)
                    dynamic feeRange = sheet.Range[$"A{subsumRow}:K{tolsumRow}"];
                    feeRange.Formula = feeArray;
                }

                // 7. 渲染顶部汇总行 (insertRow) 数据与公式联动
                sheet.Cells[insertRow, 2].Value2 = cabinet.Header.CabinetNo;
                sheet.Cells[insertRow, 3].Value2 = cabinet.Header.Model;
                sheet.Cells[insertRow, 5].Value2 = "台";
                sheet.Cells[insertRow, 6].Value2 = 1;
                // G 列单价公式指向明细总计行 (tolsumRow) 的 H 列销售总价
                sheet.Cells[insertRow, 7].Formula = $"=H{tolsumRow}";
                // H 列总价公式 = 数量 * 单价
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


