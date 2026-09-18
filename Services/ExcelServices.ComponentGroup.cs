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

                // 规则 8: 在操作/读取 Excel 表格前，先调用 FixAndFillCabinetNamesForSheet 确保规则 6 正确性
                Tool.FixAndFillCabinetNamesForSheet(sheet);

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

                // 提取列映射配置 (B=2, C=3, F=6, W=23, X=24, Z=26)
                var map = config?.ColumnMapping ?? new ComponentGroupColumnMapping();

                // 使用二维数组一次性读取 B 列到 Z 列 (从 Col 2 到至少 Col 26)
                int colStart = 2; // B 列
                // 动态计算最大所需列，确保至少完整覆盖到 Z 列 (Col 26)
                int colEnd = Math.Max(26, Math.Max(map.CurrentCol, Math.Max(map.PolesCol, map.AppendixCol)));
                // 计算行总数
                int totalRows = compEndRow - compStartRow + 1;
                // 计算列总数
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
                    int curRelCol = map.CurrentCol - colStart + 1;     // W 列相对索引 = 22
                    int poleRelCol = map.PolesCol - colStart + 1;      // X 列相对索引 = 23
                    int appRelCol = map.AppendixCol - colStart + 1;    // Z 列相对索引 = 25

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
        /// <param name="progressCallback">执行进度与状态提示回调委托 (百分比 0~100, 状态说明)</param>
        public static BatchGroupResultDto ExecuteBatchComponentGroup(
            ComponentGroupConfig config, 
            bool activeCabinetOnly = true,
            Action<int, string>? progressCallback = null)
        {
            // 初始化返回结果 DTO
            var result = new BatchGroupResultDto();
            // 启动高精度耗时统计计时器
            var sw = Stopwatch.StartNew();

            // 规则库非空校验
            if (config == null || config.Rules == null || config.Rules.Count == 0)
            {
                // 标记失败
                result.Success = false;
                // 给出友好错误提示
                result.Message = "规则管道库为空，请先添加并配置规则！";
                // 提前返回
                return result;
            }

            // 声明 Excel 宿主对象
            dynamic? app = null;
            // 声明原始计算模式
            dynamic? originalCalc = null;
            // 声明原始屏幕重绘状态
            bool originalScreenUpdating = true;
            // 声明原始事件监听状态
            bool originalEnableEvents = true;

            try
            {
                // 推送初始化进度
                progressCallback?.Invoke(3, "正在获取活动 Excel 工作簿与工作表环境...");

                // 获取活动 Excel 环境上下文
                var context = Tool.GetActiveExcelContext();
                // 校验 Excel 上下文有效性
                if (context == null || context.App == null || context.Sheet == null)
                {
                    // 标记失败
                    result.Success = false;
                    // 返回无活动表提示
                    result.Message = "未检测到活动的 Excel 工作表，请先打开工程报价表！";
                    // 提前返回
                    return result;
                }

                // 提取 Excel 句柄引用
                app = context.App;
                // 提取当前活动工作表
                dynamic sheet = context.Sheet;
                // 提取当前所属工作簿
                dynamic wb = context.Wb;

                // 1. 彻底冻结 Excel 重绘、公式重算与全局事件（核心性能护城河）
                originalScreenUpdating = app.ScreenUpdating;
                // 备份原始计算模式
                originalCalc = app.Calculation;
                // 备份原始事件监听触发开关
                try { originalEnableEvents = app.EnableEvents; } catch { }

                // 关闭屏幕重绘提升执行效率
                app.ScreenUpdating = false;
                // 暂停自动重算设为手动模式 (-4135: xlCalculationManual)
                app.Calculation = -4135; // --硬编码: Excel xlCalculationManual 枚举值--
                // 关键性能优化：彻底关闭事件触发，阻断 OnSheetChange 与智能学习级联风暴
                app.EnableEvents = false;

                // 推送工作表自愈校验进度
                progressCallback?.Invoke(6, "正在校验并确保箱柜定义名称架构规范 (规则 8)...");

                // 规则 8: 在操作/插行修改 Excel 表格前，先调用 FixAndFillCabinetNamesForSheet 确保规则 6 正确性
                Tool.FixAndFillCabinetNamesForSheet(sheet);

                // 获取工作表中所有有效箱柜锚点集合
                var validCabinets = Tool.GetSheetValidCabinets(sheet, wb);
                // 校验是否识别到有效箱柜
                if (validCabinets == null || validCabinets.Count == 0)
                {
                    // 标记失败
                    result.Success = false;
                    // 返回未识别到箱柜提示
                    result.Message = "当前工作表中未识别到符合规范的箱柜定义名称 (Cab_Det/Cab_Subsum)！";
                    // 提前退出
                    return result;
                }

                // 确定待处理的箱柜清单
                var targetCabinets = new List<KeyValuePair<int, Models.CabinetAnchorModel>>();
                // 根据作用域选择单箱柜还是全表箱柜
                if (activeCabinetOnly)
                {
                    // 智能获取光标命中的活动箱柜实体 (显式强类型声明与阻断 DLR 动态调用传播)
                    KeyValuePair<int, Models.CabinetAnchorModel>? activeCab = Tool.GetActiveCabinet((object)app, validCabinets, fallbackSingle: true);
                    // 校验是否成功定位活动箱柜
                    if (activeCab.HasValue)
                    {
                        // 将命中的箱柜实体加入处理列表
                        targetCabinets.Add(activeCab.Value);
                    }
                }
                else
                {
                    // 全表模式：将所有已识别箱柜全部加入待处理队列
                    targetCabinets.AddRange(validCabinets);
                }

                // 校验目标箱柜列表是否非空
                if (targetCabinets.Count == 0)
                {
                    // 标记失败
                    result.Success = false;
                    // 返回未定位到箱柜提示
                    result.Message = "未能定位到目标箱柜！";
                    // 提前返回
                    return result;
                }

                // 获取已启用的规则管道列表
                var activeRules = config.Rules
                    .Where(r => r.Enabled)
                    .OrderBy(r => r.Priority)
                    .ToList();

                // 提取列映射配置
                var map = config.ColumnMapping ?? new ComponentGroupColumnMapping();
                // B 列为数据起始列 (列索引 2)
                int colStart = 2; // --硬编码: B 列起始索引--
                // 动态计算最大所需列，确保至少完整覆盖到 Z 列 (Col 26) 及 CAD 句柄列 AE (Col 31)
                int colEnd = Math.Max(31, Math.Max(map.CurrentCol, Math.Max(map.PolesCol, map.AppendixCol)));
                // 计算元器件读取总列数
                int totalCols = colEnd - colStart + 1;

                // 记录所有箱柜累计物理插入的新行数
                int totalInsertedRowsAcrossCabinets = 0;
                // 记录总待处理箱柜数
                int totalCabinetsToProcess = targetCabinets.Count;

                // 倒序遍历箱柜 (自底向上处理箱柜，下方箱柜插行绝对不影响上方箱柜的物理行号)
                for (int cabIdx = totalCabinetsToProcess - 1; cabIdx >= 0; cabIdx--)
                {
                    // 提取当前箱柜序号 K
                    int k = targetCabinets[cabIdx].Key;
                    // 提取当前箱柜内存锚点模型
                    var anchor = targetCabinets[cabIdx].Value;

                    // 计算平滑进度百分比 (10% ~ 90%)
                    int processedIndex = totalCabinetsToProcess - cabIdx;
                    // 映射百分比到 10%~90% 区间
                    int percent = 10 + (int)(processedIndex * 80.0 / Math.Max(1, totalCabinetsToProcess));
                    // 实时通知当前处理进度
                    progressCallback?.Invoke(percent, $"正在处理箱柜 [{k}] ({processedIndex}/{totalCabinetsToProcess})...");

                    // 关键性能优化：直接复用已有锚点对象的行号，彻底消除循环内 O(N^2) 全簿定义名称重复扫描！
                    int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;
                    // 提取小计行行号
                    int subsumRow = anchor.Subsum != null ? Convert.ToInt32(anchor.Subsum.Row) : 0;

                    // 若行号异常则跳过
                    if (detRow <= 0 || subsumRow <= 0) continue;

                    // 计算元器件起始行 (Det 行下方第 2 行)
                    int compStartRow = detRow + 2;
                    // 计算元器件终止行 (Subsum 行上方第 1 行)
                    int compEndRow = subsumRow - 1;

                    // 若元器件区域非法则跳过
                    if (compStartRow > compEndRow) continue;

                    // 计算元器件总行数
                    int totalRows = compEndRow - compStartRow + 1;
                    // 一次性批量读取元器件区域数据矩阵
                    dynamic range = sheet.Range[sheet.Cells[compStartRow, colStart], sheet.Cells[compEndRow, colEnd]];
                    // 转换为二维数组
                    object[,] data = ConvertTo2DArray(range.Value2, totalRows, totalCols);

                    // 1. 内存中提取该箱柜的一次元件列表
                    var components = new List<EleComponentDto>();
                    // 声明已存在的二次元件组名称哈希表以备去重
                    var existingGroupNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    // 记录最后一个非空有效元器件行的相对索引 (1-based，对应 1 到 totalRows)
                    int lastUsedIndex = 0;

                    // 遍历元器件区域数据矩阵行
                    for (int r = 1; r <= totalRows; r++)
                    {
                        // 计算各属性列在数组中的相对索引
                        int nameRel = map.NameCol - colStart + 1;
                        // 型号规格相对列
                        int normsRel = map.NormsCol - colStart + 1;
                        // 数量相对列
                        int qtyRel = map.QuantityCol - colStart + 1;
                        // 电流相对列
                        int curRel = map.CurrentCol - colStart + 1;
                        // 极数相对列
                        int poleRel = map.PolesCol - colStart + 1;
                        // 附件相对列
                        int appRel = map.AppendixCol - colStart + 1;

                        // 提取名称与规格型号
                        string eleName = data[r, nameRel]?.ToString()?.Trim() ?? "";
                        // 提取规格型号
                        string eleNorms = data[r, normsRel]?.ToString()?.Trim() ?? "";
                        // 提取原始数量文本
                        string rawQty = data[r, qtyRel]?.ToString()?.Trim() ?? "0";
                        // 提取电流与极数
                        string eleCurrent = data[r, curRel]?.ToString()?.Trim() ?? "";
                        // 提取极数
                        string elePoles = data[r, poleRel]?.ToString()?.Trim() ?? "";
                        // 提取附件
                        string eleAppendix = data[r, appRel]?.ToString()?.Trim() ?? "";

                        // 收集已有型号用于二次元件组查重
                        if (!string.IsNullOrEmpty(eleNorms))
                        {
                            // 加入哈希集合
                            existingGroupNames.Add(eleNorms);
                        }

                        // 跳过空白行
                        if (string.IsNullOrEmpty(eleName) && string.IsNullOrEmpty(eleNorms)) continue;

                        // 记录最后一个有效非空元器件行的相对位置
                        lastUsedIndex = r;

                        // 解析数量数值
                        int.TryParse(rawQty, out int eleNums);
                        // 兜底最小数量为 1
                        if (eleNums <= 0) eleNums = 1;

                        // 添加至一次元器件清单
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
                    // 存储去重后最终待生成的二次元件组集合
                    var matchedList = new List<RuleMatchResult>();

                    // 遍历命中的规则进行智能去重过滤
                    foreach (var match in poolResults)
                    {
                        // 检查去重开关与名称重复
                        if (config.EnableDeduplication && existingGroupNames.Contains(match.TargetGroup))
                        {
                            // 累加跳过计数
                            result.SkippedDuplicateCount++;
                            // 记录去重日志
                            result.Details.Add($"箱柜 [{k}] 已存在元件组 [{match.TargetGroup}]，已自动跳过重复插入。");
                            // 跳过本条
                            continue;
                        }
                        // 加入有效生成列表
                        matchedList.Add(match);
                    }

                    // 若未命中任何新规则
                    if (matchedList.Count == 0)
                    {
                        // 记录日志并继续下一箱柜
                        result.Details.Add($"箱柜 [{k}] 未匹配到任何新的二次元件组规则。");
                        // 处理计数递增
                        result.ProcessedCabinets++;
                        // 进入下一个箱柜循环
                        continue;
                    }

                    // 3. 用户指令: 元件组需要放在元器件的最下面，有空行就不要插入 (遵循规则 6)
                    // 计算最后一个非空元器件所在绝对物理行号
                    int lastUsedRow = lastUsedIndex > 0 ? (compStartRow + lastUsedIndex - 1) : (compStartRow - 1);
                    // 计算元器件区域内部剩余可直接复用的空行总数
                    int availableEmptyRows = Math.Max(0, compEndRow - lastUsedRow);
                    // 待写入的二次元件组总数
                    int reqCount = matchedList.Count;

                    // 规则 6: 如果元器件数量多于区域行数，先要插入行；若空行足够，则绝不插入新行
                    if (reqCount > availableEmptyRows)
                    {
                        // 仅当空行不足时，按差额计算需要插入的行数
                        int rowsToInsert = reqCount - availableEmptyRows;
                        // 在小计行上方一次性批量插入差额空行 (避免单行循环插入)
                        dynamic insertRange = sheet.Range[$"A{subsumRow}:A{subsumRow + rowsToInsert - 1}"];
                        // 执行下移插行 (-4121: xlShiftDown)
                        insertRange.EntireRow.Insert(-4121); // --硬编码: xlShiftDown--
                        // 插入后小计行行号相应下移
                        subsumRow += rowsToInsert;
                        // 同步更新元器件终止行号
                        compEndRow = subsumRow - 1;
                        // 累计全表插行总数
                        totalInsertedRowsAcrossCabinets += rowsToInsert;
                    }

                    // 4. 关键性能优化（落实规则 7）：二维数组矩阵一次性批量写入二次元件组，消除几百次细碎 COM 往返！
                    // 确定二次元件组起始写入物理行号
                    int currentWriteRow = lastUsedRow + 1;
                    // 提取箱柜表头 CAD 图元句柄 (AD/AE 列 = 30/31)
                    string handleA = "";
                    // 提取次要图元句柄
                    string handleB = "";
                    try { handleA = sheet.Cells[detRow, 30].Value?.ToString() ?? ""; } catch { }
                    try { handleB = sheet.Cells[detRow, 31].Value?.ToString() ?? ""; } catch { }

                    // 计算待写入二维矩阵的总列数 (从 B 列到 AE 列)
                    int batchWriteCols = colEnd - colStart + 1;
                    // 构造数据二维数组 (1-based, 行数 reqCount, 列数 batchWriteCols)
                    object[,] batchValues = new object[reqCount + 1, batchWriteCols + 1];
                    // 构造 A 列动态公式二维数组 (1-based, 行数 reqCount, 列数 1)
                    object[,] formulaA = new object[reqCount + 1, 2];

                    // 在内存中高速组装全部数据矩阵
                    for (int mIdx = 0; mIdx < reqCount; mIdx++)
                    {
                        // 提取当前项
                        var match = matchedList[mIdx];
                        // 数组行号 (1-based)
                        int r = mIdx + 1;

                        // A 列自适应动态序号公式 =ROW()-ROW(A${detRow+1}) (detRow+1 为表头行)
                        formulaA[r, 1] = $"=ROW()-ROW(A${detRow + 1})"; // --硬编码: 动态序号公式表达式--

                        // B 列写入类别 (默认 "元件组")
                        batchValues[r, map.CategoryCol - colStart + 1] = config.DefaultCategoryText;

                        // C 列写入二次元件组名称 (如 *多功能表)
                        batchValues[r, map.NormsCol - colStart + 1] = match.TargetGroup;

                        // E 列写入计量单位 "套"
                        batchValues[r, map.UnitCol - colStart + 1] = config.DefaultUnitText;

                        // F 列写入计算得到的套数
                        batchValues[r, map.QuantityCol - colStart + 1] = match.Quantity;

                        // 继承 CAD 图元句柄 (AD 列 = 30, AE 列 = 31)
                        if (!string.IsNullOrEmpty(handleA)) batchValues[r, 30 - colStart + 1] = handleA;
                        // 赋值次要句柄
                        if (!string.IsNullOrEmpty(handleB)) batchValues[r, 31 - colStart + 1] = handleB;

                        // 统计生成行数
                        result.InsertedGroupsCount++;
                        // 记录执行明细日志
                        result.Details.Add($"箱柜 [{k}] 成功在行 {currentWriteRow + mIdx} 写入二次元件组: [{match.TargetGroup}]，套数: {match.Quantity} 套");
                    }

                    // 一次性批量写入 A 列序号公式
                    dynamic aRange = sheet.Range[sheet.Cells[currentWriteRow, 1], sheet.Cells[currentWriteRow + reqCount - 1, 1]];
                    // 赋值公式向量
                    aRange.Formula = formulaA;

                    // 一次性批量写入 B 列至 AE 列数据矩阵 (单次跨进程 COM 搞定整箱柜全部二次元件组)
                    dynamic dataRange = sheet.Range[sheet.Cells[currentWriteRow, colStart], sheet.Cells[currentWriteRow + reqCount - 1, colEnd]];
                    // 赋值数值矩阵
                    dataRange.Value2 = batchValues;

                    // 箱柜处理完成计数递增
                    result.ProcessedCabinets++;

                    // 触发该箱柜元器件自愈与计费联动公式刷新，确保单价等核心公式不丢失 (规则 8)
                    int tolsumRow = anchor.Tolsum != null ? Convert.ToInt32(anchor.Tolsum.Row) : (subsumRow + 5);
                    RefreshCabinetFeeAreaFormulas(sheet, detRow, compStartRow, subsumRow, tolsumRow);
                }

                // 5. 规则 8: 仅当实际产生过物理插行时，才需要执行定义名称自愈刷新；若纯复用空行则 0ms 瞬间跳过！
                if (totalInsertedRowsAcrossCabinets > 0)
                {
                    // 推送自愈进度
                    progressCallback?.Invoke(95, "正在校准定义名称与公式自愈...");
                    // 触发全表定义名称与计费区间自愈校准
                    Tool.FixAndFillCabinetNamesForSheet(sheet);
                }

                // 标记处理完成 100% 进度
                progressCallback?.Invoke(100, "二次元件组批量生成完毕！");

                // 停止计时器
                sw.Stop();
                // 记录执行毫秒数
                result.ElapsedMilliseconds = sw.ElapsedMilliseconds;
                // 标记成功
                result.Success = true;
                // 组织成功信息
                result.Message = $"批量生成成功! 共处理 {result.ProcessedCabinets} 台箱柜，生成二次元件组 {result.InsertedGroupsCount} 行，跳过重复项 {result.SkippedDuplicateCount} 项，耗时 {result.ElapsedMilliseconds} ms。";
            }
            catch (Exception ex)
            {
                // 异常停止计时
                sw.Stop();
                // 标记失败
                result.Success = false;
                // 返回异常描述
                result.Message = $"执行批量生成二次元件组时发生异常: {ex.Message}";
                result.Details.Add($"异常堆栈: {ex.StackTrace}");
            }
            finally
            {
                // 恢复 Excel 运行状态环境
                if (app != null)
                {
                    try
                    {
                        // 统一触发一次公式重算
                        app.Calculate();
                        // 恢复原始计算模式
                        if (originalCalc != null) app.Calculation = originalCalc;
                        // 恢复屏幕重绘刷新
                        app.ScreenUpdating = originalScreenUpdating;
                        // 恢复 Excel 全局事件触发监听
                        app.EnableEvents = originalEnableEvents;
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
