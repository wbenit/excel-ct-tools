using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// 可撤销重做的命令抽象接口
    /// </summary>
    public interface IUndoableCommand
    {
        // 操作显示描述（例如：“撤销: 批量匹配物料 (32行)”）
        string ActionName { get; }

        // 命令发生的时间戳
        DateTime Timestamp { get; }

        // 估算的内存占用大小（字节），用于防爆监控
        long EstimatedMemoryBytes { get; }

        // 执行撤销操作，将工作表数据还原到操作前状态
        void Undo();

        // 执行重做/还原操作，将数据重新应用为操作后状态
        void Redo();
    }

    /// <summary>
    /// 单元格单区域差量切片数据包
    /// </summary>
    public class RangeDeltaSlice
    {
        // 目标工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 目标区域的标准 A1 地址 (如 "C10:C25", "D12")
        public string RangeAddress { get; set; } = string.Empty;

        // 修改前的单元格值二维数组 (使用二维 object[,] 存储，遵循规则7)
        public object[,]? OldValues { get; set; }

        // 修改后的单元格值二维数组
        public object[,]? NewValues { get; set; }

        // 修改前单元格的公式二维数组 (可选，若涉及公式则备份)
        public object[,]? OldFormulas { get; set; }

        // 修改后单元格的公式二维数组 (可选)
        public object[,]? NewFormulas { get; set; }

        // 修改前的单元格底色索引快照
        public int[,]? OldColorIndexes { get; set; }

        // 修改后的单元格底色索引快照
        public int[,]? NewColorIndexes { get; set; }

        // 修改前的单元格真实 RGB 底色快照
        public int[,]? OldColors { get; set; }

        // 修改后的单元格真实 RGB 底色快照
        public int[,]? NewColors { get; set; }
    }

    /// <summary>
    /// 单元格区域多切片差量命令 (支持多列、多选区、矩阵批量数据)
    /// </summary>
    public class RangeDeltaCommand : IUndoableCommand
    {
        // 操作描述文本
        public string ActionName { get; }

        // 命令记录时戳
        public DateTime Timestamp { get; }

        // 包含的所有单元格区域切片列表
        private readonly List<RangeDeltaSlice> _slices;

        // 估算的内存大小
        public long EstimatedMemoryBytes { get; }

        /// <summary>
        /// 构造单元格区域差量命令
        /// </summary>
        /// <param name="actionName">命令显示名称</param>
        /// <param name="slices">涉及的区域差量切片列表</param>
        public RangeDeltaCommand(string actionName, List<RangeDeltaSlice> slices)
        {
            // 记录显示名称
            ActionName = actionName;
            // 记录生成时间
            Timestamp = DateTime.Now;
            // 缓存所有区域切片
            _slices = slices ?? new List<RangeDeltaSlice>();

            // 粗略估算内存大小，按切片数组元素数计算
            long totalBytes = 128;
            foreach (var slice in _slices)
            {
                if (slice.OldValues != null)
                {
                    totalBytes += slice.OldValues.Length * 16;
                }
                if (slice.NewValues != null)
                {
                    totalBytes += slice.NewValues.Length * 16;
                }
            }
            EstimatedMemoryBytes = totalBytes;
        }

        /// <summary>
        /// 执行撤销：将 OldValues 与 OldColors 整块写回 Excel
        /// </summary>
        public void Undo()
        {
            ApplySlices(isUndo: true);
        }

        /// <summary>
        /// 执行重做：将 NewValues 与 NewColors 整块写回 Excel
        /// </summary>
        public void Redo()
        {
            ApplySlices(isUndo: false);
        }

        /// <summary>
        /// 批量应用切片数据，内部统一挂起屏幕重绘与自动重算
        /// </summary>
        private void ApplySlices(bool isUndo)
        {
            if (_slices == null || _slices.Count == 0) return;

            // 获取 Excel Application 实例句柄
            dynamic? app = ExcelDnaUtil.Application;
            if (app == null) return;

            // 记录原始重绘与计算状态
            bool oldScreenUpdating = true;
            int oldCalculation = -4105; // xlCalculationAutomatic --硬编码: Excel原生自动计算常数--
            bool oldEnableEvents = true;

            try
            {
                // 读取原宿主配置
                oldScreenUpdating = app.ScreenUpdating;
                oldCalculation = app.Calculation;
                oldEnableEvents = app.EnableEvents;

                // 挂起屏幕重绘，防止界面剧烈闪烁
                app.ScreenUpdating = false;
                // 暂时切换为手动计算，阻断级联公式重算风暴
                app.Calculation = -4135; // xlCalculationManual --硬编码: Excel原生手动计算常数--
                // 暂停事件通知，避免二次触发 SheetChange
                app.EnableEvents = false;

                // 获取活动工作簿
                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null) return;

                // 遍历每个切片并执行写回
                foreach (var slice in _slices)
                {
                    if (string.IsNullOrWhiteSpace(slice.SheetName) || string.IsNullOrWhiteSpace(slice.RangeAddress))
                    {
                        continue;
                    }

                    // 检索目标工作表
                    dynamic? targetSheet = null;
                    try
                    {
                        targetSheet = activeWb.Worksheets[slice.SheetName];
                    }
                    catch { }

                    if (targetSheet == null) continue;

                    // 获取目标区域
                    dynamic targetRange = targetSheet.Range[slice.RangeAddress];
                    if (targetRange == null) continue;

                    // 1. 恢复单元格数值 (使用二维数组一次性赋值，遵循规则7)
                    object[,]? valuesToApply = isUndo ? slice.OldValues : slice.NewValues;
                    if (valuesToApply != null)
                    {
                        targetRange.Value2 = valuesToApply;
                    }

                    // 2. 若有公式，恢复单元格公式
                    object[,]? formulasToApply = isUndo ? slice.OldFormulas : slice.NewFormulas;
                    if (formulasToApply != null)
                    {
                        targetRange.Formula = formulasToApply;
                    }

                    // 3. 恢复底色样式
                    int[,]? colorIndexesToApply = isUndo ? slice.OldColorIndexes : slice.NewColorIndexes;
                    int[,]? colorsToApply = isUndo ? slice.OldColors : slice.NewColors;

                    if (colorIndexesToApply != null && colorsToApply != null)
                    {
                        ApplyColorsToRange(targetRange, colorIndexesToApply, colorsToApply);
                    }
                }

                // 统一触发一次工作簿重算，保证公式联动刷新
                try { app.Calculate(); } catch { }
            }
            catch (Exception ex)
            {
                // 记录写回异常日志
                LogHelper.WriteLog($"RangeDeltaCommand ApplySlices 异常: {ex.Message}");
            }
            finally
            {
                // 可靠恢复原计算环境与事件通知
                try
                {
                    app.ScreenUpdating = oldScreenUpdating;
                    app.Calculation = oldCalculation;
                    app.EnableEvents = oldEnableEvents;
                }
                catch { }
            }
        }

        /// <summary>
        /// 批量恢复指定 Range 的单元格底色
        /// </summary>
        private static void ApplyColorsToRange(dynamic targetRange, int[,] colorIndexes, int[,] colors)
        {
            try
            {
                int rowCount = colorIndexes.GetLength(0);
                int colCount = colorIndexes.GetLength(1);

                // 单单元格快速直出分支
                if (rowCount == 1 && colCount == 1)
                {
                    int cIdx = colorIndexes[0, 0];
                    if (cIdx == -4142) // xlNone --硬编码: 无底色索引--
                    {
                        targetRange.Interior.ColorIndex = -4142;
                    }
                    else
                    {
                        targetRange.Interior.Color = colors[0, 0];
                    }
                    return;
                }

                // 多单元格逐格还原底色
                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        int cIdx = colorIndexes[r, c];
                        // Excel COM Cells 索引为 1-based
                        dynamic cell = targetRange.Cells[r + 1, c + 1];
                        if (cIdx == -4142) // xlNone
                        {
                            cell.Interior.ColorIndex = -4142;
                        }
                        else
                        {
                            cell.Interior.Color = colors[r, c];
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ApplyColorsToRange 底色恢复异常: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 复合操作命令包装器：将多个原子子命令打包为一个可撤销事务
    /// </summary>
    public class CompositeUndoableCommand : IUndoableCommand
    {
        // 复合事务描述
        public string ActionName { get; }

        // 时戳
        public DateTime Timestamp { get; }

        // 内部子命令列表
        private readonly List<IUndoableCommand> _commands;

        // 内存开销估算
        public long EstimatedMemoryBytes => _commands.Sum(c => c.EstimatedMemoryBytes);

        /// <summary>
        /// 构造复合命令
        /// </summary>
        public CompositeUndoableCommand(string actionName, List<IUndoableCommand> commands)
        {
            ActionName = actionName;
            Timestamp = DateTime.Now;
            _commands = commands ?? new List<IUndoableCommand>();
        }

        /// <summary>
        /// 撤销：倒序执行内部所有子命令
        /// </summary>
        public void Undo()
        {
            // 按倒序依次撤销每个子命令
            for (int i = _commands.Count - 1; i >= 0; i--)
            {
                try
                {
                    _commands[i].Undo();
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"CompositeUndoableCommand 子命令撤销异常: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 重做：顺向执行内部所有子命令
        /// </summary>
        public void Redo()
        {
            // 按正序依次重做每个子命令
            for (int i = 0; i < _commands.Count; i++)
            {
                try
                {
                    _commands[i].Redo();
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"CompositeUndoableCommand 子命令重做异常: {ex.Message}");
                }
            }
        }
    }

    /// <summary>
    /// 基于委托轻量执行的命令包装器 (适用于成对出现的业务操作如筛选/清除筛选等)
    /// </summary>
    public class ActionUndoableCommand : IUndoableCommand
    {
        // 命令操作显示名称
        public string ActionName { get; }

        // 命令发生时间戳
        public DateTime Timestamp { get; }

        // 估算的内存大小
        public long EstimatedMemoryBytes => 128;

        // 撤销委托回调
        private readonly Action _undoAction;

        // 重做委托回调
        private readonly Action _redoAction;

        /// <summary>
        /// 构造委托型可撤销命令
        /// </summary>
        public ActionUndoableCommand(string actionName, Action undoAction, Action redoAction)
        {
            ActionName = actionName;
            Timestamp = DateTime.Now;
            _undoAction = undoAction;
            _redoAction = redoAction;
        }

        /// <summary>
        /// 执行撤销委托
        /// </summary>
        public void Undo()
        {
            _undoAction?.Invoke();
        }

        /// <summary>
        /// 执行重做委托
        /// </summary>
        public void Redo()
        {
            _redoAction?.Invoke();
        }
    }

    /// <summary>
    /// 分布调价单个箱柜差量切片数据包 (支持物理行插入与 30 列公式矩阵无损快照)
    /// </summary>
    public class DistributionCabinetSlice
    {
        // 目标工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 箱柜柜号
        public string CabinetNo { get; set; } = string.Empty;

        // 元器件起始行 (Det.Row + 2)
        public int CompStartRow { get; set; }

        // 修改前的元器件终止行 (原 Subsum.Row - 1)
        public int OrigCompEndRow { get; set; }

        // 修改前原始小计行 (Subsum.Row)
        public int OrigSubsumRow { get; set; }

        // 本次针对该箱柜物理插入的新行数 (若空行足够未插行则为 0)
        public int InsertedRowCount { get; set; }

        // 插入行的起始物理行号 (即原 Subsum.Row)
        public int InsertedRowStart { get; set; }

        // 修改后的元器件终止行 (OrigCompEndRow + InsertedRowCount)
        public int NewCompEndRow { get; set; }

        // 修改后的小计行号 (OrigSubsumRow + InsertedRowCount)
        public int NewSubsumRow { get; set; }

        // 修改前元器件区域完整的 30 列公式/数值矩阵快照 (覆盖 A 列至 AD 列 CAD 句柄，规则 7)
        public object[,]? OldFormulas { get; set; }

        // 修改后元器件区域完整的 30 列公式/数值矩阵快照 (覆盖 A 列至 AD 列 CAD 句柄，规则 7)
        public object[,]? NewFormulas { get; set; }
    }

    /// <summary>
    /// 分布调价可逆命令 (双模自愈：无插行 30 列矩阵秒级还原，有插行倒序物理行删除与规则 8 闭环自愈)
    /// </summary>
    public class DistributionAdjustPriceCommand : IUndoableCommand
    {
        // 命令操作显示名称
        public string ActionName { get; }

        // 命令发生时间戳
        public DateTime Timestamp { get; }

        // 包含的所有箱柜差量切片列表
        private readonly List<DistributionCabinetSlice> _slices;

        // 估算的内存大小
        public long EstimatedMemoryBytes { get; }

        /// <summary>
        /// 构造分布调价可逆命令
        /// </summary>
        public DistributionAdjustPriceCommand(string actionName, List<DistributionCabinetSlice> slices)
        {
            // 记录显示名称
            ActionName = actionName;
            // 记录生成时间
            Timestamp = DateTime.Now;
            // 缓存所有箱柜切片
            _slices = slices ?? new List<DistributionCabinetSlice>();

            // 粗略估算内存大小
            long totalBytes = 256;
            foreach (var slice in _slices)
            {
                if (slice.OldFormulas != null) totalBytes += slice.OldFormulas.Length * 16;
                if (slice.NewFormulas != null) totalBytes += slice.NewFormulas.Length * 16;
            }
            EstimatedMemoryBytes = totalBytes;
        }

        /// <summary>
        /// 执行撤销：自下而上倒序物理删除插入行并无损还原原始 30 列公式矩阵
        /// </summary>
        public void Undo()
        {
            ApplyDistributionSlices(isUndo: true);
        }

        /// <summary>
        /// 执行重做：自上而下正序重新插入行并应用更新后 30 列公式矩阵
        /// </summary>
        public void Redo()
        {
            ApplyDistributionSlices(isUndo: false);
        }

        /// <summary>
        /// 应用分布调价切片 (支持撤销与重做双模调度)
        /// </summary>
        private void ApplyDistributionSlices(bool isUndo)
        {
            if (_slices == null || _slices.Count == 0) return;

            // 获取 Excel Application 实例句柄
            dynamic? app = ExcelDnaUtil.Application;
            if (app == null) return;

            // 记录原始重绘与计算状态
            bool oldScreenUpdating = true;
            int oldCalculation = -4105; // xlCalculationAutomatic --硬编码: Excel原生自动计算常数--
            bool oldEnableEvents = true;

            try
            {
                // 读取原宿主配置
                oldScreenUpdating = app.ScreenUpdating;
                oldCalculation = app.Calculation;
                oldEnableEvents = app.EnableEvents;

                // 挂起屏幕重绘，防止界面闪烁
                app.ScreenUpdating = false;
                // 暂时切换为手动计算，阻断公式级联重算
                app.Calculation = -4135; // xlCalculationManual --硬编码: Excel原生手动计算常数--
                // 暂停系统事件
                app.EnableEvents = false;

                // 获取活动工作簿句柄
                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null) return;

                // 收集发生变动的工作表名称集合 (用于后续规则 8 自愈)
                var affectedSheetNames = new HashSet<string>(
                    _slices.Select(s => s.SheetName).Where(n => !string.IsNullOrWhiteSpace(n))
                );

                // 按工作表分组分别处理，确保每个工作表内的行号操作严格有序
                var sheetsGroup = _slices.GroupBy(s => s.SheetName);

                foreach (var group in sheetsGroup)
                {
                    string sheetName = group.Key;
                    if (string.IsNullOrWhiteSpace(sheetName)) continue;

                    // 获取目标工作表
                    dynamic? targetSheet = null;
                    try { targetSheet = activeWb.Worksheets[sheetName]; } catch { }
                    if (targetSheet == null) continue;

                    // 核心：若为撤销，必须在单表内自下而上倒序执行（按 CompStartRow 从大到小），确保删除插入行时上方物理行号绝不偏移！
                    // 若为重做，则自上而下正序执行（按 CompStartRow 从小到大）
                    var orderedCabinetSlices = isUndo
                        ? group.OrderByDescending(s => s.CompStartRow).ToList()
                        : group.OrderBy(s => s.CompStartRow).ToList();

                    foreach (var slice in orderedCabinetSlices)
                    {
                        if (isUndo)
                        {
                            // 1. 若该箱柜曾物理插入过行，先精准倒序删除插入的多余行
                            if (slice.InsertedRowCount > 0 && slice.InsertedRowStart > 0)
                            {
                                try
                                {
                                    int delStart = slice.InsertedRowStart;
                                    int delEnd = slice.InsertedRowStart + slice.InsertedRowCount - 1;
                                    // 物理整行删除插入行
                                    targetSheet.Range[$"{delStart}:{delEnd}"].EntireRow.Delete();
                                }
                                catch (Exception exDel)
                                {
                                    LogHelper.WriteLog($"[分布调价撤销] 箱柜 {slice.CabinetNo} 删除插入行异常: {exDel.Message}");
                                }
                            }

                            // 2. 将原始 30 列公式与值矩阵一次性赋回原元器件区域 (遵循规则 7)
                            if (slice.OldFormulas != null)
                            {
                                dynamic origCompRange = targetSheet.Range[$"A{slice.CompStartRow}:AD{slice.OrigCompEndRow}"];
                                origCompRange.Formula = slice.OldFormulas;
                            }

                            // 3. 联动恢复小计行求和公式 (规则 6: Cab_Subsum_k)
                            if (slice.OrigSubsumRow > 0)
                            {
                                try
                                {
                                    targetSheet.Cells[slice.OrigSubsumRow, 8].Formula = $"=SUM(H{slice.CompStartRow}:H{slice.OrigCompEndRow})";
                                }
                                catch { }
                            }
                        }
                        else
                        {
                            // 1. 重做逻辑：若该箱柜需要插行，重新在小计行前插入行
                            if (slice.InsertedRowCount > 0 && slice.InsertedRowStart > 0)
                            {
                                try
                                {
                                    int insStart = slice.InsertedRowStart;
                                    int insEnd = slice.InsertedRowStart + slice.InsertedRowCount - 1;
                                    // 批量插入整行 --硬编码: xlDown -4121--
                                    targetSheet.Range[$"{insStart}:{insEnd}"].Insert(-4121);
                                }
                                catch (Exception exIns)
                                {
                                    LogHelper.WriteLog($"[分布调价重做] 箱柜 {slice.CabinetNo} 重新插入行异常: {exIns.Message}");
                                }
                            }

                            // 2. 将修改后的 30 列新公式矩阵一次性赋回 (遵循规则 7)
                            if (slice.NewFormulas != null)
                            {
                                dynamic newCompRange = targetSheet.Range[$"A{slice.CompStartRow}:AD{slice.NewCompEndRow}"];
                                newCompRange.Formula = slice.NewFormulas;
                            }

                            // 3. 联动刷新修改后的小计行求和公式
                            if (slice.NewSubsumRow > 0)
                            {
                                try
                                {
                                    targetSheet.Cells[slice.NewSubsumRow, 8].Formula = $"=SUM(H{slice.CompStartRow}:H{slice.NewCompEndRow})";
                                }
                                catch { }
                            }
                        }
                    }

                    // 针对当前分类表，严格执行规则 8：闭环自愈校准定义名称与超链接
                    try
                    {
                        Tool.FixAndFillCabinetNamesForSheet(targetSheet);
                    }
                    catch (Exception exFix)
                    {
                        LogHelper.WriteLog($"[分布调价] 工作表 {sheetName} 自愈校准异常: {exFix.Message}");
                    }
                }

                // 统一触发一次全工作簿公式重新计算，刷新所有小计与合价联动
                try { app.Calculate(); } catch { }
            }
            catch (Exception ex)
            {
                // 记录写回异常日志
                LogHelper.WriteLog($"DistributionAdjustPriceCommand ApplyDistributionSlices 异常: {ex.Message}");
            }
            finally
            {
                // 可靠恢复原计算环境与事件通知
                try
                {
                    app.ScreenUpdating = oldScreenUpdating;
                    app.Calculation = oldCalculation;
                    app.EnableEvents = oldEnableEvents;
                }
                catch { }
            }
        }
    }

    /// <summary>
    /// 全局撤销/重做生命周期与双栈调度管理中心 (单例)
    /// </summary>
    public sealed class UndoRedoManager
    {
        // 单例懒加载实例
        private static readonly Lazy<UndoRedoManager> _instance = new Lazy<UndoRedoManager>(() => new UndoRedoManager());

        // 获取全局唯一调度器单例
        public static UndoRedoManager Instance => _instance.Value;

        // 默认历史记录最大步数上限 (超出自动丢弃最老记录，防止内存泄漏)
        private const int MaxHistoryLimit = 30; // --硬编码: 历史记录最大容量--

        // 可撤销历史栈 (先进先出截断)
        private readonly LinkedList<IUndoableCommand> _undoStack = new LinkedList<IUndoableCommand>();

        // 可重做历史栈
        private readonly LinkedList<IUndoableCommand> _redoStack = new LinkedList<IUndoableCommand>();

        // 互斥防重入锁标志：防止在执行 Undo/Redo 过程中产生新的日志污染或事件死循环
        private bool _isExecuting = false;

        // 线程同步锁对象
        private readonly object _syncLock = new object();

        // 私有构造函数
        private UndoRedoManager() { }

        /// <summary>
        /// 是否处于撤销/还原内部执行态
        /// </summary>
        public bool IsExecuting => _isExecuting;

        /// <summary>
        /// 是否有可撤销的历史操作
        /// </summary>
        public bool CanUndo
        {
            get
            {
                lock (_syncLock)
                {
                    return _undoStack.Count > 0;
                }
            }
        }

        /// <summary>
        /// 是否有可还原的重做操作
        /// </summary>
        public bool CanRedo
        {
            get
            {
                lock (_syncLock)
                {
                    return _redoStack.Count > 0;
                }
            }
        }

        /// <summary>
        /// 获取下一个可撤销操作的显示名称
        /// </summary>
        public string? CurrentUndoName
        {
            get
            {
                lock (_syncLock)
                {
                    return _undoStack.Last?.Value.ActionName;
                }
            }
        }

        /// <summary>
        /// 获取下一个可还原重做操作的显示名称
        /// </summary>
        public string? CurrentRedoName
        {
            get
            {
                lock (_syncLock)
                {
                    return _redoStack.Last?.Value.ActionName;
                }
            }
        }

        /// <summary>
        /// 推入一个新的可撤销业务命令
        /// </summary>
        /// <param name="command">要入栈的命令对象</param>
        public void PushCommand(IUndoableCommand? command)
        {
            if (command == null) return;
            // 若当前正在执行撤销/重做，坚决不推入新命令，避免形成死循环
            if (_isExecuting) return;

            lock (_syncLock)
            {
                // 新操作发生，清空重做栈
                _redoStack.Clear();

                // 压入撤销栈顶
                _undoStack.AddLast(command);

                // 若超出历史容量上限，从头部移除最老的一个记录
                while (_undoStack.Count > MaxHistoryLimit)
                {
                    _undoStack.RemoveFirst();
                }
            }

            // 挂接与同步 Excel 原生 Application.OnUndo
            RegisterExcelOnUndo(command.ActionName);
        }

        /// <summary>
        /// 执行一步撤销操作
        /// </summary>
        /// <returns>是否成功执行了撤销</returns>
        public bool Undo()
        {
            IUndoableCommand? cmd = null;
            lock (_syncLock)
            {
                if (_undoStack.Count == 0) return false;
                // 从撤销栈弹出最新的一个命令
                cmd = _undoStack.Last?.Value;
                if (cmd == null) return false;
                _undoStack.RemoveLast();
            }

            // 开启防重入互斥标志
            _isExecuting = true;
            try
            {
                // 调用具体命令的 Undo 执行回滚
                cmd.Undo();

                lock (_syncLock)
                {
                    // 将撤销成功的命令推入重做栈
                    _redoStack.AddLast(cmd);
                }

                // 在 Excel 状态栏提示撤销成功
                SetExcelStatusBar($"[成套工具] 已撤销: {cmd.ActionName}");

                // 若撤销栈还有上一步，更新 Excel 原生 OnUndo；否则重置
                if (CanUndo)
                {
                    RegisterExcelOnUndo(CurrentUndoName ?? "撤销上一步操作");
                }

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"UndoRedoManager Undo 执行失败: {ex.Message}");
                return false;
            }
            finally
            {
                // 解除防重入互斥标志
                _isExecuting = false;
            }
        }

        /// <summary>
        /// 执行一步重做/还原操作
        /// </summary>
        /// <returns>是否成功执行了重做</returns>
        public bool Redo()
        {
            IUndoableCommand? cmd = null;
            lock (_syncLock)
            {
                if (_redoStack.Count == 0) return false;
                // 从重做栈弹出最新的一个命令
                cmd = _redoStack.Last?.Value;
                if (cmd == null) return false;
                _redoStack.RemoveLast();
            }

            // 开启防重入互斥标志
            _isExecuting = true;
            try
            {
                // 调用具体命令的 Redo 重新应用
                cmd.Redo();

                lock (_syncLock)
                {
                    // 重新推回撤销栈
                    _undoStack.AddLast(cmd);
                }

                // 在 Excel 状态栏给出成功提示
                SetExcelStatusBar($"[成套工具] 已还原: {cmd.ActionName}");

                // 同步更新原生 OnUndo
                RegisterExcelOnUndo(cmd.ActionName);

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"UndoRedoManager Redo 执行失败: {ex.Message}");
                return false;
            }
            finally
            {
                // 解除防重入互斥标志
                _isExecuting = false;
            }
        }

        /// <summary>
        /// 清空所有撤销与重做历史记录
        /// </summary>
        public void Clear()
        {
            lock (_syncLock)
            {
                _undoStack.Clear();
                _redoStack.Clear();
            }
        }

        /// <summary>
        /// 挂接 Excel 原生 Application.OnUndo 入口
        /// </summary>
        private void RegisterExcelOnUndo(string actionName)
        {
            try
            {
                // 将任务投递到 Excel 消息队列安全执行
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        dynamic? app = ExcelDnaUtil.Application;
                        if (app != null)
                        {
                            // 挂接自定义宏过程名，当用户点击 Excel 原生撤销按钮或按系统快捷键时调用
                            app.OnUndo($"成套工具: {actionName}", "MacroUndoAction");
                            // 挂接重做宏过程
                            if (CanRedo)
                            {
                                app.OnRepeat($"成套工具: {CurrentRedoName}", "MacroRedoAction");
                            }
                        }
                    }
                    catch { }
                });
            }
            catch { }
        }

        /// <summary>
        /// 设置 Excel 底部状态栏提示文本
        /// </summary>
        private static void SetExcelStatusBar(string message)
        {
            try
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        dynamic? app = ExcelDnaUtil.Application;
                        if (app != null)
                        {
                            app.StatusBar = message;
                        }
                    }
                    catch { }
                });
            }
            catch { }
        }
    }
}
