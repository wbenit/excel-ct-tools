using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Forms;

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

                // 标记撤回内容：撤销回滚完成后，自动在工作表中定位并选中被恢复的单元格区域
                if (isUndo && _slices != null && _slices.Count > 0)
                {
                    try
                    {
                        // 提取首个切片信息
                        var firstSlice = _slices[0];
                        // 校验工作表与地址有效性
                        if (!string.IsNullOrWhiteSpace(firstSlice.SheetName) && !string.IsNullOrWhiteSpace(firstSlice.RangeAddress))
                        {
                            // 激活目标工作表
                            dynamic targetSheet = app.ActiveWorkbook.Worksheets[firstSlice.SheetName];
                            targetSheet.Activate();
                            // 选中该单元格区域，高亮标记出刚被撤回的内容
                            dynamic focusRange = targetSheet.Range[firstSlice.RangeAddress];
                            focusRange.Select();
                        }
                    }
                    catch { }
                }
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
    /// 成套元器件整行删除可逆命令 (支持连续多行、带公式、保留首行空行骨架保护及撤销选区高亮)
    /// </summary>
    public class ComponentRowDeleteCommand : IUndoableCommand
    {
        // 操作描述文本 (如：“撤销: 删除元件 (行 12: 断路器)”)
        public string ActionName { get; }

        // 命令记录时戳
        public DateTime Timestamp { get; }

        // 内存占用估算
        public long EstimatedMemoryBytes { get; }

        // 目标工作表名称
        public string SheetName { get; }

        // 删除起始物理行号
        public int StartRow { get; }

        // 删除终止物理行号
        public int EndRow { get; }

        // 删除的总行数
        public int RowCount { get; }

        // 备份的有效列数
        public int ColCount { get; }

        // 是否触发了“保留首行空行”骨架安全防御分支
        public bool IsPreservedFirstBlankRow { get; }

        // 备份的完整二维数据矩阵 (规则 7 一次性读写)
        public object[,] OldValues { get; }

        // 备份的公式集合 (列号 -> 公式字符串)
        public Dictionary<int, string>? OldFormulas { get; }

        /// <summary>
        /// 构造元器件删除可逆命令
        /// </summary>
        public ComponentRowDeleteCommand(
            string actionName,
            string sheetName,
            int startRow,
            int endRow,
            int rowCount,
            int colCount,
            bool isPreservedFirstBlankRow,
            object[,] oldValues,
            Dictionary<int, string>? oldFormulas)
        {
            // 记录显示文本
            ActionName = actionName;
            // 记录发生时间戳
            Timestamp = DateTime.Now;
            // 记录目标工作表名称
            SheetName = sheetName;
            // 记录起始行号
            StartRow = startRow;
            // 记录终止行号
            EndRow = endRow;
            // 记录总行数
            RowCount = rowCount;
            // 记录列数
            ColCount = colCount;
            // 记录骨架防御标志
            IsPreservedFirstBlankRow = isPreservedFirstBlankRow;
            // 缓存旧数据二维数组
            OldValues = oldValues;
            // 缓存旧公式字典
            OldFormulas = oldFormulas;

            // 粗略估算内存大小
            EstimatedMemoryBytes = 128 + (oldValues != null ? oldValues.Length * 16 : 0);
        }

        /// <summary>
        /// 执行撤销：将删除的行重新插回并精准还原二维数据矩阵与计算公式
        /// </summary>
        public void Undo()
        {
            // 获取 Excel Application 实例
            dynamic? app = ExcelDnaSafeAccessor.GetApplication();
            if (app == null) return;

            // 获取活动工作簿
            dynamic? activeWb = app.ActiveWorkbook;
            if (activeWb == null) return;

            // 记录原环境状态
            bool prevUpdating = true;
            int prevCalc = -4105; // xlCalculationAutomatic --硬编码: 原生自动计算--
            bool prevEvents = true;
            bool prevAlerts = true;

            try
            {
                // 备份环境状态
                prevUpdating = Convert.ToBoolean(app.ScreenUpdating);
                prevCalc = Convert.ToInt32(app.Calculation);
                prevEvents = Convert.ToBoolean(app.EnableEvents);
                prevAlerts = Convert.ToBoolean(app.DisplayAlerts);

                // 挂起屏幕重绘与自动重算
                app.ScreenUpdating = false;
                app.Calculation = -4135; // xlCalculationManual --硬编码: 原生手动计算--
                app.EnableEvents = false;
                app.DisplayAlerts = false;

                // 提取目标工作表
                dynamic targetSheet = activeWb.Worksheets[SheetName];
                if (targetSheet == null) return;

                // 分支处理：若当时触发了骨架防御（首行清空保留，仅删除了第 2 行至末尾行）
                if (IsPreservedFirstBlankRow)
                {
                    // 若删除行数大于 1，在 StartRow + 1 处向下插回删掉的行
                    if (RowCount > 1)
                    {
                        // 在首行下方整块插入物理行
                        targetSheet.Range[targetSheet.Rows[StartRow + 1], targetSheet.Rows[EndRow]].Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                    }
                    // 一次性写回完整二维数据矩阵
                    targetSheet.Range[targetSheet.Cells[StartRow, 1], targetSheet.Cells[EndRow, ColCount]].Value2 = OldValues;
                }
                else
                {
                    // 普通整行删除分支：整块向下平移插回 RowCount 行
                    targetSheet.Range[targetSheet.Rows[StartRow], targetSheet.Rows[EndRow]].Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                    // 一次性批量写回完整二维数据矩阵 (规则 7)
                    targetSheet.Range[targetSheet.Cells[StartRow, 1], targetSheet.Cells[EndRow, ColCount]].Value2 = OldValues;
                }

                // 恢复自适应标准计算公式 (H 列总价 = F*G, K 列成本总价 = F*J)
                for (int i = 0; i < RowCount; i++)
                {
                    // 计算当前物理行号
                    int curR = StartRow + i;
                    // 恢复销售总价公式
                    targetSheet.Cells[curR, 8].Formula = $"=F{curR}*G{curR}";
                    // 恢复成本总价公式
                    targetSheet.Cells[curR, 11].Formula = $"=F{curR}*J{curR}";
                }

                // 触发工作簿公式联动重算
                try { app.Calculate(); } catch { }

                // 自动激活工作表并高亮选中恢复回来的元器件行
                try
                {
                    // 激活工作表
                    targetSheet.Activate();
                    // 圈选恢复的前台业务数据列 (第 1 列至第 20 列)
                    int focusCol = Math.Min(ColCount, 20);
                    // 选中并高亮显示恢复的单元格区域
                    targetSheet.Range[targetSheet.Cells[StartRow, 1], targetSheet.Cells[EndRow, focusCol]].Select();
                }
                catch { }
            }
            catch (Exception ex)
            {
                // 记录撤销失败异常
                LogHelper.WriteLog($"ComponentRowDeleteCommand Undo 异常: {ex.Message}");
            }
            finally
            {
                // 安全成对恢复原宿主环境
                try
                {
                    app.DisplayAlerts = prevAlerts;
                    app.EnableEvents = prevEvents;
                    app.Calculation = prevCalc;
                    app.ScreenUpdating = prevUpdating;
                }
                catch { }
            }
        }

        /// <summary>
        /// 执行重做：再次将指定元器件整行物理删除
        /// </summary>
        public void Redo()
        {
            // 获取 Excel Application 实例
            dynamic? app = ExcelDnaSafeAccessor.GetApplication();
            if (app == null) return;

            // 获取活动工作簿
            dynamic? activeWb = app.ActiveWorkbook;
            if (activeWb == null) return;

            // 记录原环境状态
            bool prevUpdating = true;
            int prevCalc = -4105; // xlCalculationAutomatic --硬编码: 原生自动计算--
            bool prevEvents = true;
            bool prevAlerts = true;

            try
            {
                // 备份环境状态
                prevUpdating = Convert.ToBoolean(app.ScreenUpdating);
                prevCalc = Convert.ToInt32(app.Calculation);
                prevEvents = Convert.ToBoolean(app.EnableEvents);
                prevAlerts = Convert.ToBoolean(app.DisplayAlerts);

                // 挂起屏幕重绘与自动重算
                app.ScreenUpdating = false;
                app.Calculation = -4135; // xlCalculationManual --硬编码: 原生手动计算--
                app.EnableEvents = false;
                app.DisplayAlerts = false;

                // 提取目标工作表
                dynamic targetSheet = activeWb.Worksheets[SheetName];
                if (targetSheet == null) return;

                // 分支处理：骨架防御模式下清空首行并删除后续行
                if (IsPreservedFirstBlankRow)
                {
                    // 若大于 1 行则物理删除后续行
                    if (RowCount > 1)
                    {
                        targetSheet.Range[targetSheet.Rows[StartRow + 1], targetSheet.Rows[EndRow]].Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                    }
                    // 清空首行元器件数据
                    object[,] blankRow = new object[1, ColCount];
                    targetSheet.Range[targetSheet.Cells[StartRow, 1], targetSheet.Cells[StartRow, ColCount]].Value2 = blankRow;
                }
                else
                {
                    // 普通模式下直接物理删除整块行
                    targetSheet.Range[targetSheet.Rows[StartRow], targetSheet.Rows[EndRow]].Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                }

                // 触发工作簿公式联动重算
                try { app.Calculate(); } catch { }
            }
            catch (Exception ex)
            {
                // 记录重做失败异常
                LogHelper.WriteLog($"ComponentRowDeleteCommand Redo 异常: {ex.Message}");
            }
            finally
            {
                // 安全成对恢复原宿主环境
                try
                {
                    app.DisplayAlerts = prevAlerts;
                    app.EnableEvents = prevEvents;
                    app.Calculation = prevCalc;
                    app.ScreenUpdating = prevUpdating;
                }
                catch { }
            }
        }
    }

    /// <summary>
    /// 成套元器件整行插入可逆命令 (支持复制插入回滚、剪切插入双端原子回滚与剪贴板动效自愈)
    /// </summary>
    public class ComponentRowInsertCommand : IUndoableCommand
    {
        // 操作描述文本 (如：“撤销: 插入元件 (行 20: 接触器)”)
        public string ActionName { get; }

        // 命令发生时间戳
        public DateTime Timestamp { get; }

        // 内存占用估算
        public long EstimatedMemoryBytes { get; }

        // 目标工作表名称
        public string TargetSheetName { get; }

        // 插入起始物理行号
        public int TargetRow { get; }

        // 插入的连续行数
        public int RowCount { get; }

        // 数据列数
        public int ColCount { get; }

        // 目标插入行写入的数据矩阵 (规则 7 一次性读写)
        public object[,] InsertedValues { get; }

        // 是否为剪切模式插入 (若为 true 则需同时回滚源工作表被删除的行)
        public bool IsCutMode { get; }

        // 发生剪切时的原始数据交换实体快照
        public ComponentRowExchangeDto? SourceClipSnapshot { get; }

        /// <summary>
        /// 构造元器件插入可逆命令
        /// </summary>
        public ComponentRowInsertCommand(
            string actionName,
            string targetSheetName,
            int targetRow,
            int rowCount,
            int colCount,
            object[,] insertedValues,
            bool isCutMode,
            ComponentRowExchangeDto? sourceClipSnapshot)
        {
            // 记录描述文本
            ActionName = actionName;
            // 记录发生时间戳
            Timestamp = DateTime.Now;
            // 记录目标工作表
            TargetSheetName = targetSheetName;
            // 记录目标插入行号
            TargetRow = targetRow;
            // 记录总行数
            RowCount = rowCount;
            // 记录列数
            ColCount = colCount;
            // 缓存插入数据二维矩阵
            InsertedValues = insertedValues;
            // 记录是否为剪切模式
            IsCutMode = isCutMode;
            // 缓存源剪切数据快照
            SourceClipSnapshot = sourceClipSnapshot;

            // 粗略估算内存大小
            EstimatedMemoryBytes = 128 + (insertedValues != null ? insertedValues.Length * 16 : 0);
        }

        /// <summary>
        /// 执行撤销：删除目标位置插入的行；若为剪切模式，则将源行完整插回源工作表并还原剪切板与流动虚线
        /// </summary>
        public void Undo()
        {
            // 获取 Excel Application 实例
            dynamic? app = ExcelDnaSafeAccessor.GetApplication();
            if (app == null) return;

            // 获取活动工作簿
            dynamic? activeWb = app.ActiveWorkbook;
            if (activeWb == null) return;

            // 记录原环境状态
            bool prevUpdating = true;
            int prevCalc = -4105; // xlCalculationAutomatic --硬编码: 原生自动计算--
            bool prevEvents = true;
            bool prevAlerts = true;

            try
            {
                // 备份环境状态
                prevUpdating = Convert.ToBoolean(app.ScreenUpdating);
                prevCalc = Convert.ToInt32(app.Calculation);
                prevEvents = Convert.ToBoolean(app.EnableEvents);
                prevAlerts = Convert.ToBoolean(app.DisplayAlerts);

                // 挂起屏幕重绘与自动重算
                app.ScreenUpdating = false;
                app.Calculation = -4135; // xlCalculationManual --硬编码: 原生手动计算--
                app.EnableEvents = false;
                app.DisplayAlerts = false;

                // 1. 目标位置回滚：物理整块删除插入的 RowCount 行
                dynamic targetSheet = activeWb.Worksheets[TargetSheetName];
                if (targetSheet != null)
                {
                    // 物理删除目标行 (Shift Up)
                    targetSheet.Range[targetSheet.Rows[TargetRow], targetSheet.Rows[TargetRow + RowCount - 1]].Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                }

                // 2. 源位置回滚：若为剪切模式，恢复原工作表被删除的源行
                if (IsCutMode && SourceClipSnapshot != null)
                {
                    // 获取源工作表对象
                    dynamic srcSheet = activeWb.Worksheets[SourceClipSnapshot.SourceSheetName];
                    if (srcSheet != null)
                    {
                        // 源行物理位置 (因目标插入行已被上一步删除，行号已精确回弹至原始 SourceRowIndex)
                        int srcRow = SourceClipSnapshot.SourceRowIndex;
                        int srcCols = SourceClipSnapshot.ColumnCount;

                        // 物理向下平移插回原始源行
                        srcSheet.Range[srcSheet.Rows[srcRow], srcSheet.Rows[srcRow + RowCount - 1]].Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);

                        // 一次性写回源行的完整原始数据矩阵 (含原始 CadHandle 与业务列)
                        srcSheet.Range[srcSheet.Cells[srcRow, 1], srcSheet.Cells[srcRow + RowCount - 1, srcCols]].Value2 = SourceClipSnapshot.FullRowValues;

                        // 恢复源行的自适应计算公式
                        for (int i = 0; i < RowCount; i++)
                        {
                            int curR = srcRow + i;
                            srcSheet.Cells[curR, 8].Formula = $"=F{curR}*G{curR}";
                            srcSheet.Cells[curR, 11].Formula = $"=F{curR}*J{curR}";
                        }

                        // 重新还原内存剪贴板与剪切流动虚线
                        ComponentClipboardManager.Set(SourceClipSnapshot);
                        try
                        {
                            // 重新启动流动细虚线动效
                            int visualCol = Math.Min(srcCols, 20);
                            MarchingAntsManager.Show(srcSheet, srcRow, srcRow + RowCount - 1, visualCol);
                            // 标记剪切状态
                            app.CutCopyMode = (Microsoft.Office.Interop.Excel.XlCutCopyMode)1; // xlCut --硬编码: 原生剪切状态--
                        }
                        catch { }
                    }
                }

                // 触发工作簿公式重算
                try { app.Calculate(); } catch { }
            }
            catch (Exception ex)
            {
                // 记录撤销失败异常
                LogHelper.WriteLog($"ComponentRowInsertCommand Undo 异常: {ex.Message}");
            }
            finally
            {
                // 安全成对恢复原宿主环境
                try
                {
                    app.DisplayAlerts = prevAlerts;
                    app.EnableEvents = prevEvents;
                    app.Calculation = prevCalc;
                    app.ScreenUpdating = prevUpdating;
                }
                catch { }
            }
        }

        /// <summary>
        /// 执行重做：再次在目标位置插入行并写入数据；若为剪切模式则再次安全删除源行
        /// </summary>
        public void Redo()
        {
            // 获取 Excel Application 实例
            dynamic? app = ExcelDnaSafeAccessor.GetApplication();
            if (app == null) return;

            // 获取活动工作簿
            dynamic? activeWb = app.ActiveWorkbook;
            if (activeWb == null) return;

            // 记录原环境状态
            bool prevUpdating = true;
            int prevCalc = -4105; // xlCalculationAutomatic --硬编码: 原生自动计算--
            bool prevEvents = true;
            bool prevAlerts = true;

            try
            {
                // 备份环境状态
                prevUpdating = Convert.ToBoolean(app.ScreenUpdating);
                prevCalc = Convert.ToInt32(app.Calculation);
                prevEvents = Convert.ToBoolean(app.EnableEvents);
                prevAlerts = Convert.ToBoolean(app.DisplayAlerts);

                // 挂起屏幕重绘与自动重算
                app.ScreenUpdating = false;
                app.Calculation = -4135; // xlCalculationManual --硬编码: 原生手动计算--
                app.EnableEvents = false;
                app.DisplayAlerts = false;

                // 1. 目标位置执行向下平移插入 RowCount 行
                dynamic targetSheet = activeWb.Worksheets[TargetSheetName];
                if (targetSheet != null)
                {
                    // 物理插入整行
                    targetSheet.Range[targetSheet.Rows[TargetRow], targetSheet.Rows[TargetRow + RowCount - 1]].Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);

                    // 一次性批量写回插入数据矩阵 (规则 7)
                    int writeCols = InsertedValues.GetLength(1);
                    targetSheet.Range[targetSheet.Cells[TargetRow, 1], targetSheet.Cells[TargetRow + RowCount - 1, writeCols]].Value2 = InsertedValues;

                    // 恢复自适应标准计算公式 (F*G, F*J)
                    for (int i = 0; i < RowCount; i++)
                    {
                        int curR = TargetRow + i;
                        targetSheet.Cells[curR, 8].Formula = $"=F{curR}*G{curR}";
                        targetSheet.Cells[curR, 11].Formula = $"=F{curR}*J{curR}";
                    }
                }

                // 2. 若为剪切模式，再次物理删除源行
                if (IsCutMode && SourceClipSnapshot != null)
                {
                    // 获取源工作表
                    dynamic srcSheet = activeWb.Worksheets[SourceClipSnapshot.SourceSheetName];
                    if (srcSheet != null)
                    {
                        // 计算源物理行号
                        int actualSrcRow = SourceClipSnapshot.SourceRowIndex;
                        // 若同表且源行在插入点之后，因插入行增加了偏移
                        if (string.Equals(SourceClipSnapshot.SourceSheetName, TargetSheetName, StringComparison.OrdinalIgnoreCase)
                            && actualSrcRow >= TargetRow)
                        {
                            actualSrcRow += RowCount;
                        }

                        // 物理整块删除源剪切行
                        srcSheet.Range[srcSheet.Rows[actualSrcRow], srcSheet.Rows[actualSrcRow + RowCount - 1]].Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);

                        // 清空剪贴板并隐藏流动虚线
                        ComponentClipboardManager.Clear();
                        try
                        {
                            MarchingAntsManager.Hide();
                            app.CutCopyMode = (Microsoft.Office.Interop.Excel.XlCutCopyMode)0;
                        }
                        catch { }
                    }
                }

                // 触发工作簿公式重算
                try { app.Calculate(); } catch { }

                // 聚焦并选中目标插入行
                try
                {
                    targetSheet.Activate();
                    int focusCol = Math.Min(ColCount, 20);
                    targetSheet.Range[targetSheet.Cells[TargetRow, 1], targetSheet.Cells[TargetRow + RowCount - 1, focusCol]].Select();
                }
                catch { }
            }
            catch (Exception ex)
            {
                // 记录重做失败异常
                LogHelper.WriteLog($"ComponentRowInsertCommand Redo 异常: {ex.Message}");
            }
            finally
            {
                // 安全成对恢复原宿主环境
                try
                {
                    app.DisplayAlerts = prevAlerts;
                    app.EnableEvents = prevEvents;
                    app.Calculation = prevCalc;
                    app.ScreenUpdating = prevUpdating;
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

                // 状态变更：无论是否还有可撤销项，均通知 Ribbon 刷新控件显示与禁用状态
                RegisterExcelOnUndo(CurrentUndoName ?? string.Empty);

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

                // 同步刷新 Ribbon 控件的显示与状态
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
            // 历史栈清空后即时刷新 Ribbon 撤销与还原按钮为禁用/无记录状态
            RegisterExcelOnUndo(string.Empty);
        }

        /// <summary>
        /// 撤销/重做状态变更通知 (贯彻“撤回不使用快捷键”，刷新功能区 Ribbon 标签显示)
        /// </summary>
        private void RegisterExcelOnUndo(string actionName)
        {
            try
            {
                // 将刷新任务投递到 Excel 宏消息队列安全执行
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        // 动态刷新 Ribbon 功能区上的撤销、还原与清空历史按钮
                        RibbonController.InvalidateRibbon();
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
