using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 公共服务分部类: 二次元件组规则管道匹配与 Excel 批量插行生成
    /// </summary>
    public static partial class ExcelServices
    {
        // 缓存当前已打开的二次元件组规则管道配置窗体单例引用
        private static ComponentGroupBuilderForm? _componentGroupBuilderForm = null;

        /// <summary>
        /// 弹出基于 WebView2 + Vue 3 的“二次元件组规则管道构建器”窗口 (非模态，可编辑 Excel)
        /// </summary>
        public static void ShowComponentGroupBuilderDialog()
        {
            try
            {
                // 若窗体已存在且未被销毁，则直接还原并推至最前
                if (_componentGroupBuilderForm != null && !_componentGroupBuilderForm.IsDisposed)
                {
                    // 还原最小化状态
                    if (_componentGroupBuilderForm.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    {
                        _componentGroupBuilderForm.WindowState = System.Windows.Forms.FormWindowState.Normal;
                    }
                    _componentGroupBuilderForm.BringToFront();
                    _componentGroupBuilderForm.Activate();
                    return;
                }

                // 实例化全新窗体
                _componentGroupBuilderForm = new ComponentGroupBuilderForm();
                // 绑定销毁事件清空单例引用
                _componentGroupBuilderForm.FormClosed += (s, e) => _componentGroupBuilderForm = null;

                // 获取 Excel 主窗口 HWND 句柄以非模态方式依附弹出
                IntPtr excelHwnd = ExcelDnaSafeAccessor.GetWindowHandle();
                if (excelHwnd != IntPtr.Zero)
                {
                    _componentGroupBuilderForm.Show(new ExcelWin32Window(excelHwnd));
                }
                else
                {
                    _componentGroupBuilderForm.Show();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"打开二次元件组规则管道窗口失败: {ex.Message}", "错误提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 从当前 Excel 活动工作表及光标选中的箱柜中读取一次元件数据 (供前端沙盒测试)
        /// </summary>
        /// <param name="config">列映射及全局配置对象</param>
        /// <returns>元件 DTO 列表</returns>
        public static List<EleComponentDto> GetActiveCabinetComponentsFromExcel(ComponentGroupConfig config)
        {
            var resultList = new List<EleComponentDto>();
            try
            {
                // 获取当前活动 Excel 环境上下文
                var context = Tool.GetActiveExcelContext();
                if (context == null || context.App == null || context.Sheet == null)
                {
                    // 记录未检测到 Excel 上下文的日志
                    LogHelper.WriteLog("[抓取元件] 未检测到有效的活动 Excel 上下文 (ActiveSheet 为空)。");
                    return resultList;
                }

                dynamic sheet = context.Sheet;
                dynamic wb = context.Wb;
                dynamic app = context.App;
                string sheetName = sheet.Name?.ToString() ?? "未知工作表";

                // 获取当前工作表中所有有效箱柜
                var validCabinets = Tool.GetSheetValidCabinets(sheet, wb);
                LogHelper.WriteLog($"[抓取元件] 当前工作表 [{sheetName}] 扫描到的有效箱柜数量: {validCabinets?.Count ?? 0}");

                if (validCabinets == null || validCabinets.Count == 0)
                {
                    // 记录未找到有效箱柜的诊断日志
                    LogHelper.WriteLog($"[抓取元件] 工作表 [{sheetName}] 未找到任何箱柜定义名称 (如 Cab_Det_1/Cab_Subsum_1)，抓取终止。");
                    return resultList;
                }

                // 智能获取光标命中的活动箱柜实体 (显式强类型声明与阻断 DLR 动态调用传播)
                KeyValuePair<int, Models.CabinetAnchorModel>? activeCab = Tool.GetActiveCabinet((object)app, validCabinets, fallbackSingle: true);
                if (!activeCab.HasValue)
                {
                    // 记录未能定位箱柜实体的日志
                    LogHelper.WriteLog("[抓取元件] 未能定位到任何有效箱柜实体，抓取终止。");
                    return resultList;
                }

                int k = activeCab.Value.Key;
                // 阻断 DLR 动态传播，强类型解构获取该箱柜的标准行索引
                var (sumRow, detRow, subsumRow, tolsumRow) = Tool.FindStandardCategoryRowIndexes((object)sheet, k);

                // 输出箱柜行号探测日志
                LogHelper.WriteLog($"[抓取元件] 命中箱柜序号 K={k}, Det行={detRow}, Subsum小计行={subsumRow}, Tolsum总计行={tolsumRow}");

                // 元件有效行区域: detRow + 2 至 subsumRow - 1
                int compStartRow = detRow + 2;
                int compEndRow = subsumRow - 1;
                LogHelper.WriteLog($"[抓取元件] 计算得出元器件扫描区域: 行 {compStartRow} 至 行 {compEndRow} (共计 {compEndRow - compStartRow + 1} 行)");

                if (compStartRow > compEndRow)
                {
                    // 记录元器件行区间异常日志
                    LogHelper.WriteLog($"[抓取元件] 元器件起始行 {compStartRow} > 终止行 {compEndRow}，该箱柜元器件区域为空。");
                    return resultList;
                }

                // 提取列映射配置 (B=2, C=3, F=6, V=22, W=23, X=24)
                var map = config?.ColumnMapping ?? new ComponentGroupColumnMapping();

                // 使用二维数组一次性读取 B 列到 X 列 (从 Col 2 到 Col 24)
                int colStart = 2; // B 列
                int colEnd = 24;  // X 列
                int totalRows = compEndRow - compStartRow + 1;
                int totalCols = colEnd - colStart + 1;

                dynamic range = sheet.Range[sheet.Cells[compStartRow, colStart], sheet.Cells[compEndRow, colEnd]];
                object[,] data = ConvertTo2DArray(range.Value2, totalRows, totalCols);

                // 遍历解析每一行元件
                for (int r = 1; r <= totalRows; r++)
                {
                    int realRow = compStartRow + r - 1;

                    // 计算在 2D 数组内的相对列偏移 (基准从 B列=1 开始)
                    int nameRelCol = map.NameCol - colStart + 1;       // B 列相对索引 = 1
                    int normsRelCol = map.NormsCol - colStart + 1;     // C 列相对索引 = 2
                    int qtyRelCol = map.QuantityCol - colStart + 1;    // F 列相对索引 = 5
                    int curRelCol = map.CurrentCol - colStart + 1;     // V 列相对索引 = 21
                    int poleRelCol = map.PolesCol - colStart + 1;      // W 列相对索引 = 22
                    int appRelCol = map.AppendixCol - colStart + 1;    // X 列相对索引 = 23

                    string eleName = data[r, nameRelCol]?.ToString()?.Trim() ?? "";
                    string eleNorms = data[r, normsRelCol]?.ToString()?.Trim() ?? "";
                    string rawQty = data[r, qtyRelCol]?.ToString()?.Trim() ?? "0";
                    string eleCurrent = data[r, curRelCol]?.ToString()?.Trim() ?? "";
                    string elePoles = data[r, poleRelCol]?.ToString()?.Trim() ?? "";
                    string eleAppendix = data[r, appRelCol]?.ToString()?.Trim() ?? "";

                    // 跳过空白行
                    if (string.IsNullOrEmpty(eleName) && string.IsNullOrEmpty(eleNorms))
                    {
                        continue;
                    }

                    int.TryParse(rawQty, out int eleNums);
                    if (eleNums <= 0) eleNums = 1;

                    resultList.Add(new EleComponentDto
                    {
                        RowIndex = realRow,
                        EleName = eleName,
                        EleNorms = eleNorms,
                        EleNums = eleNums,
                        EleCurrent = eleCurrent,
                        ElePoles = elePoles,
                        EleAppendix = eleAppendix
                    });
                }

                // 记录成功抓取的元件数量
                LogHelper.WriteLog($"[抓取元件] 成功抓取箱柜 [{k}] 的有效元件数据共 {resultList.Count} 条。");
            }
            catch (Exception ex)
            {
                // 记录读取活动箱柜元件数据异常日志
                LogHelper.WriteLog($"读取活动箱柜元件数据发生异常: {ex.Message}");
            }

            return resultList;
        }

        /// <summary>
        /// 执行沙盒规则管道测试 (支持箱柜元件动态资源池扣减机制)
        /// </summary>
        public static PipelineTestResultDto RunSandboxPipelineTest(ComponentGroupConfig config, List<EleComponentDto> components)
        {
            var testResult = new PipelineTestResultDto
            {
                TotalComponents = components?.Count ?? 0
            };

            if (config == null || config.Rules == null || components == null || components.Count == 0)
            {
                // 输入参数为空时的友好提示
                testResult.Logs.Add("未输入有效的测试元件列表或规则库为空。");
                return testResult;
            }

            // 按优先级排序筛选所有已启用的规则管道
            var activeRules = config.Rules
                .Where(r => r.Enabled)
                .OrderBy(r => r.Priority)
                .ToList();

            // 执行带动态资源池扣减的完整规则流评估
            var matchedResults = PipelineEvaluator.EvaluateRulesWithResourcePool(activeRules, components, out var poolLogs, out var remainingComps);

            testResult.MatchedRules = matchedResults;
            testResult.Logs = poolLogs;
            testResult.RemainingComponents = remainingComps;
            testResult.Logs.Add($"沙盒评估完成: 成功命中 {testResult.MatchedRules.Count} 项二次元件组生成规则。");

            return testResult;
        }

        /// <summary>
        /// 执行 Excel 批量规则管道匹配并安全插行回填
        /// </summary>
        /// <param name="config">规则与列映射配置</param>
        /// <param name="activeCabinetOnly">是否仅处理当前选中的单个箱柜 (false 表示处理当前工作表全部箱柜)</param>
        public static BatchGroupResultDto ExecuteBatchComponentGroup(ComponentGroupConfig config, bool activeCabinetOnly = true)
        {
            var result = new BatchGroupResultDto();
            var sw = Stopwatch.StartNew();

            if (config == null || config.Rules == null || config.Rules.Count == 0)
            {
                result.Success = false;
                result.Message = "规则管道库为空，请先添加并配置规则！";
                return result;
            }

            dynamic? app = null;
            dynamic? originalCalc = null;
            bool originalScreenUpdating = true;

            try
            {
                // 获取活动 Excel 环境上下文
                var context = Tool.GetActiveExcelContext();
                if (context == null || context.App == null || context.Sheet == null)
                {
                    result.Success = false;
                    result.Message = "未检测到活动的 Excel 工作表，请先打开工程报价表！";
                    return result;
                }

                app = context.App;
                dynamic sheet = context.Sheet;
                dynamic wb = context.Wb;

                // 备份并优化 Excel 性能参数 (关闭屏幕刷新，设为手动重算)
                originalScreenUpdating = app.ScreenUpdating;
                originalCalc = app.Calculation;
                app.ScreenUpdating = false;
                app.Calculation = -4135; // xlCalculationManual

                // 获取工作表中所有有效箱柜
                var validCabinets = Tool.GetSheetValidCabinets(sheet, wb);
                if (validCabinets == null || validCabinets.Count == 0)
                {
                    result.Success = false;
                    result.Message = "当前工作表中未识别到符合规范的箱柜定义名称 (Cab_Det/Cab_Subsum)！";
                    return result;
                }

                // 确定待处理的箱柜清单
                var targetCabinets = new List<KeyValuePair<int, Models.CabinetAnchorModel>>();
                if (activeCabinetOnly)
                {
                    // 智能获取光标命中的活动箱柜实体 (显式强类型声明与阻断 DLR 动态调用传播)
                    KeyValuePair<int, Models.CabinetAnchorModel>? activeCab = Tool.GetActiveCabinet((object)app, validCabinets, fallbackSingle: true);
                    if (activeCab.HasValue)
                    {
                        // 将命中的箱柜实体加入处理列表
                        targetCabinets.Add(activeCab.Value);
                    }
                }
                else
                {
                    targetCabinets.AddRange(validCabinets);
                }

                if (targetCabinets.Count == 0)
                {
                    result.Success = false;
                    result.Message = "未能定位到目标箱柜！";
                    return result;
                }

                // 获取已启用的规则管道
                var activeRules = config.Rules
                    .Where(r => r.Enabled)
                    .OrderBy(r => r.Priority)
                    .ToList();

                var map = config.ColumnMapping ?? new ComponentGroupColumnMapping();
                int colStart = 2; // B 列
                int colEnd = 24;  // X 列
                int totalCols = colEnd - colStart + 1;

                // 倒序遍历箱柜 (自底向上处理箱柜，防止上方箱柜插行影响下方箱柜行号)
                for (int cabIdx = targetCabinets.Count - 1; cabIdx >= 0; cabIdx--)
                {
                    int k = targetCabinets[cabIdx].Key;
                    // 阻断 DLR 动态传播，强类型解构获取该箱柜的标准行索引
                    var (sumRow, detRow, subsumRow, tolsumRow) = Tool.FindStandardCategoryRowIndexes((object)sheet, k);

                    int compStartRow = detRow + 2;
                    int compEndRow = subsumRow - 1;

                    if (compStartRow > compEndRow) continue;

                    int totalRows = compEndRow - compStartRow + 1;
                    dynamic range = sheet.Range[sheet.Cells[compStartRow, colStart], sheet.Cells[compEndRow, colEnd]];
                    object[,] data = ConvertTo2DArray(range.Value2, totalRows, totalCols);

                    // 1. 内存中提取该箱柜的一次元件列表
                    var components = new List<EleComponentDto>();
                    var existingGroupNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    for (int r = 1; r <= totalRows; r++)
                    {
                        int nameRel = map.NameCol - colStart + 1;
                        int normsRel = map.NormsCol - colStart + 1;
                        int qtyRel = map.QuantityCol - colStart + 1;
                        int curRel = map.CurrentCol - colStart + 1;
                        int poleRel = map.PolesCol - colStart + 1;
                        int appRel = map.AppendixCol - colStart + 1;

                        string eleName = data[r, nameRel]?.ToString()?.Trim() ?? "";
                        string eleNorms = data[r, normsRel]?.ToString()?.Trim() ?? "";
                        string rawQty = data[r, qtyRel]?.ToString()?.Trim() ?? "0";
                        string eleCurrent = data[r, curRel]?.ToString()?.Trim() ?? "";
                        string elePoles = data[r, poleRel]?.ToString()?.Trim() ?? "";
                        string eleAppendix = data[r, appRel]?.ToString()?.Trim() ?? "";

                        if (!string.IsNullOrEmpty(eleNorms))
                        {
                            // 收集已有元件组名称以备去重
                            existingGroupNames.Add(eleNorms);
                        }

                        if (string.IsNullOrEmpty(eleName) && string.IsNullOrEmpty(eleNorms)) continue;

                        int.TryParse(rawQty, out int eleNums);
                        if (eleNums <= 0) eleNums = 1;

                        components.Add(new EleComponentDto
                        {
                            RowIndex = compStartRow + r - 1,
                            EleName = eleName,
                            EleNorms = eleNorms,
                            EleNums = eleNums,
                            EleCurrent = eleCurrent,
                            ElePoles = elePoles,
                            EleAppendix = eleAppendix
                        });
                    }

                    // 2. 评估规则管道 (基于箱柜元件动态资源池扣减机制)
                    var poolResults = PipelineEvaluator.EvaluateRulesWithResourcePool(activeRules, components, out var poolLogs, out _);
                    var matchedList = new List<RuleMatchResult>();

                    // 遍历命中的规则进行智能去重过滤
                    foreach (var match in poolResults)
                    {
                        // 检查去重
                        if (config.EnableDeduplication && existingGroupNames.Contains(match.TargetGroup))
                        {
                            result.SkippedDuplicateCount++;
                            result.Details.Add($"箱柜 [{k}] 已存在元件组 [{match.TargetGroup}]，已自动跳过重复插入。");
                            continue;
                        }
                        matchedList.Add(match);
                    }

                    if (matchedList.Count == 0)
                    {
                        result.Details.Add($"箱柜 [{k}] 未匹配到任何新的二次元件组规则。");
                        result.ProcessedCabinets++;
                        continue;
                    }

                    // 3. 在小计行前 (subsumRow 处) 批量插入行并写入二次元件组
                    // 提取句柄 (AD/AE 列 = 30/31)
                    string handleA = sheet.Cells[detRow, 30].Value?.ToString() ?? "";
                    string handleB = sheet.Cells[detRow, 31].Value?.ToString() ?? "";

                    int insertPoint = subsumRow;

                    foreach (var match in matchedList)
                    {
                        // 物理向下插入一行
                        sheet.Rows[insertPoint].Insert(-4121); // xlShiftDown

                        // B 列写入 "元件组" (类别)
                        sheet.Cells[insertPoint, map.CategoryCol].Value = config.DefaultCategoryText;

                        // C 列写入二次元件组名称 (如 *ATS+MXOF)
                        sheet.Cells[insertPoint, map.NormsCol].Value = match.TargetGroup;

                        // E 列写入计量单位 "套"
                        sheet.Cells[insertPoint, map.UnitCol].Value = config.DefaultUnitText;

                        // F 列写入计算得到的套数 (按用户指示: F 列写入计算套数)
                        sheet.Cells[insertPoint, map.QuantityCol].Value = match.Quantity;

                        // 继承 CAD 图元句柄 (AD / AE 列)
                        if (!string.IsNullOrEmpty(handleA)) sheet.Cells[insertPoint, 30].Value = handleA;
                        if (!string.IsNullOrEmpty(handleB)) sheet.Cells[insertPoint, 31].Value = handleB;

                        result.InsertedGroupsCount++;
                        result.Details.Add($"箱柜 [{k}] 成功在行 {insertPoint} 插入二次元件组: [{match.TargetGroup}]，套数: {match.Quantity} 套");

                        // 插入点向下递增
                        insertPoint++;
                    }

                    result.ProcessedCabinets++;
                }

                // 全表定义名称与公式自适应刷新
                Tool.FixAndFillCabinetNamesForSheet(sheet);

                sw.Stop();
                result.ElapsedMilliseconds = sw.ElapsedMilliseconds;
                result.Success = true;
                result.Message = $"批量生成成功! 共处理 {result.ProcessedCabinets} 台箱柜，生成二次元件组 {result.InsertedGroupsCount} 行，跳过重复项 {result.SkippedDuplicateCount} 项，耗时 {result.ElapsedMilliseconds} ms。";
            }
            catch (Exception ex)
            {
                sw.Stop();
                result.Success = false;
                result.Message = $"执行批量生成二次元件组时发生异常: {ex.Message}";
                result.Details.Add($"异常堆栈: {ex.StackTrace}");
            }
            finally
            {
                // 恢复 Excel 运行状态
                if (app != null)
                {
                    try
                    {
                        app.Calculate();
                        if (originalCalc != null) app.Calculation = originalCalc;
                        app.ScreenUpdating = originalScreenUpdating;
                    }
                    catch { }
                }
            }

            return result;
        }

        /// <summary>
        /// 将 COM 读取的二维数组对象转换为标准 object[,]
        /// </summary>
        private static object[,] ConvertTo2DArray(object? value, int expectedRows, int expectedCols)
        {
            if (value is object[,] arr2D)
            {
                return arr2D;
            }

            object[,] result = new object[expectedRows + 1, expectedCols + 1];
            if (value != null && expectedRows >= 1 && expectedCols >= 1)
            {
                result[1, 1] = value;
            }
            return result;
        }
    }
}
