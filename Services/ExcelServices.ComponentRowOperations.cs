using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 公共服务分部类: 成套元器件整行生命周期操作 (剪切/复制/插入/删除)
    /// 遵循规范：所有 Excel COM 操作收敛于公共服务中，采用二维数组一次性读写，自动自愈定义名称与重算公式
    /// </summary>
    public static partial class ExcelServices
    {
        // 元器件数据读取与回写的默认列数 (覆盖 A~T 前台列与 U~AF 隐藏列) --硬编码: 读取列宽常数--
        private const int DefaultComponentColumnCount = 35;

        // CadHandle (图元句柄) 所在的物理列号 (第 30 列，AD 列) --硬编码: CadHandle列号--
        private const int CadHandleColIndex = 30;

        // 辅助 Handle 所在的物理列号 (第 31 列，AE 列) --硬编码: HandleB列号--
        private const int CadHandleBColIndex = 31;

        /// <summary>
        /// 复制当前选中的元器件整行 (快捷键 Ctrl+Shift+C 或右键菜单【复制元件】触发)
        /// 提取包含前台可见列与后台隐藏参数在内的完整业务对象至内存剪贴板
        /// </summary>
        public static void CopyComponentRow()
        {
            // 调度通用提取逻辑，模式标记为复制
            ExtractActiveComponentRowToClipboard(isCutMode: false);
        }

        /// <summary>
        /// 剪切当前选中的元器件整行 (快捷键 Ctrl+Shift+X 或右键菜单【剪切元件】触发)
        /// 暂存元器件完整数据，并在后续插入成功后物理安全删除原行
        /// </summary>
        public static void CutComponentRow()
        {
            // 调度通用提取逻辑，模式标记为剪切
            ExtractActiveComponentRowToClipboard(isCutMode: true);
        }

        /// <summary>
        /// 提取活动单元格所在元器件行至剪贴板
        /// </summary>
        /// <param name="isCutMode">true 为剪切模式，false 为复制模式</param>
        private static void ExtractActiveComponentRowToClipboard(bool isCutMode)
        {
            try
            {
                // 获取当前正在运行的 Excel Application
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return;

                // 优先获取当前选区，若无选区则回退至当前活动单元格
                dynamic? selection = app.Selection;
                dynamic? activeCell = app.ActiveCell;
                if (selection == null && activeCell == null) return;

                // 获取所属工作表对象
                dynamic sheet = (selection != null) ? selection.Worksheet : activeCell!.Worksheet;
                string sheetName = Convert.ToString(sheet.Name) ?? "";

                // 解析选区连续起始行号与结束行号 (支持单选单元格与连续多选行)
                int startRow;
                int endRow;
                if (selection != null)
                {
                    // 检测是否存在按住 Ctrl 产生的离散多块选区
                    try
                    {
                        if (selection.Areas != null && selection.Areas.Count > 1)
                        {
                            System.Windows.Forms.MessageBox.Show(
                                "成套元器件复制/剪切暂不支持离散多选区，请拖拽选择单行或连续多行！",
                                "提示",
                                System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Information);
                            return;
                        }
                    }
                    catch { }

                    // 读取连续区域的起始行与连续行数
                    startRow = Convert.ToInt32(selection.Row);
                    int selRowsCount = Convert.ToInt32(selection.Rows.Count);
                    endRow = startRow + selRowsCount - 1;
                }
                else
                {
                    // 单个单元格回退模式
                    startRow = Convert.ToInt32(activeCell!.Row);
                    endRow = startRow;
                }

                // 计算本次待提取的元器件连续总行数
                int rowCount = endRow - startRow + 1;

                // 规则 8: 操作前强制执行 FixAndFillCabinetNamesForSheet 校准定义名称
                Tool.FixAndFillCabinetNamesForSheet(sheet);

                // 解析定位选区首行与末行所属的箱柜信息 (使用强类型引用模型)
                CabinetRowContext? startCab = FindCabinetByRow(sheet, startRow);
                CabinetRowContext? endCab = (rowCount == 1) ? startCab : FindCabinetByRow(sheet, endRow);

                // 校验：选区必须属于同一箱柜且所有行必须完整落入有效元器件区间内
                if (startCab == null || endCab == null || startCab.CabinetK != endCab.CabinetK
                    || startRow < startCab.CompStartRow || endRow > startCab.CompEndRow)
                {
                    // 若选区跨箱柜或触碰到了表头、小计行或计费区域则弹出警示
                    System.Windows.Forms.MessageBox.Show(
                        "请在同一箱柜的有效元器件数据行中选择行！\n(严禁跨箱柜选择，且不能包含箱柜信息行、表头行、小计行或计费区域)",
                        "提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // 计算当前表实际读取的列数
                int colCount = Math.Max(DefaultComponentColumnCount, GetSheetColumnCount(sheet));

                // 规则 7: 采用二维数组一次性读取所有选区行数据至内存矩阵 (rowCount 行 × colCount 列)
                dynamic readRange = sheet.Range[sheet.Cells[startRow, 1], sheet.Cells[endRow, colCount]];
                object[,] fullValues = (object[,])readRange.Value2;

                // 提取包含计算公式的单元格集合 (如总价公式、成本总价公式)
                var formulas = new Dictionary<int, string>();
                try
                {
                    // 检查首行 H 列 (总价) 是否包含公式
                    string formulaH = Convert.ToString(sheet.Cells[startRow, 8].Formula) ?? "";
                    if (formulaH.StartsWith("=")) formulas[8] = formulaH;

                    // 检查首行 K 列 (成本总价) 是否包含公式
                    string formulaK = Convert.ToString(sheet.Cells[startRow, 11].Formula) ?? "";
                    if (formulaK.StartsWith("=")) formulas[11] = formulaK;
                }
                catch { }

                // 提取首行 CadHandle 与 HandleB 供参考
                string handleVal = "";
                string handleBVal = "";
                if (colCount >= CadHandleColIndex)
                {
                    handleVal = Convert.ToString(fullValues[1, CadHandleColIndex])?.Trim() ?? "";
                }
                if (colCount >= CadHandleBColIndex)
                {
                    handleBVal = Convert.ToString(fullValues[1, CadHandleBColIndex])?.Trim() ?? "";
                }

                // 提取首行元器件名称与规格型号文本供界面提示展示
                string nameVal = Convert.ToString(fullValues[1, 2])?.Trim() ?? "";
                string modelVal = Convert.ToString(fullValues[1, 3])?.Trim() ?? "";

                // 获取工作簿名称
                string wbName = Convert.ToString(sheet.Parent?.Name) ?? "";

                // 构造多行数据交换实体
                var exchangeDto = new ComponentRowExchangeDto
                {
                    IsCutMode = isCutMode,
                    SourceWorkbookName = wbName,
                    SourceSheetName = sheetName,
                    SourceCabinetK = startCab.CabinetK,
                    SourceRowIndex = startRow,
                    RowCount = rowCount,
                    ColumnCount = colCount,
                    FullRowValues = fullValues,
                    CellFormulas = formulas,
                    CadHandle = handleVal,
                    HandleB = handleBVal,
                    ComponentName = nameVal,
                    ComponentModel = modelVal
                };

                // 压入全局内存剪贴板管理器
                ComponentClipboardManager.Set(exchangeDto);

                // 在 Excel 状态栏展示轻量反馈提示
                string actionText = isCutMode ? "已剪切" : "已复制";
                string summaryText = rowCount > 1
                    ? $"{rowCount} 行元件 (行 {startRow}~{endRow})"
                    : $"行 {startRow}: {nameVal} {modelVal}";
                app.StatusBar = $"[{actionText}元件] {summaryText}";

                // 呈现动态效果：启动高质感橙色流动细虚线穿透浮层动效 (Marching Ants 动态流动虚线)
                try
                {
                    // 圈定当前元器件行前台核心业务数据列 (A~T 列，第 1 列至第 20 列) --硬编码: 可见业务数据前20列--
                    int visualEndCol = Math.Min(colCount, 20);
                    // 启动橙色流动细虚线动效
                    MarchingAntsManager.Show(sheet, startRow, endRow, visualEndCol);
                }
                catch { }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[ComponentRowOperations] 提取元件数据异常: {ex.Message}");
                // 提取失败时弹出友好提示，杜绝静默失败导致后续插入扑空
                string actionName = isCutMode ? "剪切元件" : "复制元件";
                // 弹出异常信息提示对话框
                System.Windows.Forms.MessageBox.Show(
                    $"{actionName}失败: {ex.Message}",
                    "系统提示",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 插入复制或剪切的元器件 (快捷键 Ctrl+Shift+V 或右键菜单【插入复制/剪切的元件】触发)
        /// 核心特性：跨箱柜时自动清空 CadHandle；维护小计 SUM 公式；重排序号；自愈定义名称
        /// </summary>
        public static void InsertCopiedOrCutComponentRow()
        {
            try
            {
                // 校验剪贴板中是否存在有效数据
                if (!ComponentClipboardManager.HasData)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "剪贴板中暂无元器件数据，请先使用【复制元件】或【剪切元件】！",
                        "系统提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // 获取当前正在运行的 Excel Application
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return;

                // 获取活动单元格句柄
                dynamic? activeCell = app.ActiveCell;
                if (activeCell == null) return;

                // 获取目标工作表与行号
                dynamic targetSheet = activeCell.Worksheet;
                int targetRow = Convert.ToInt32(activeCell.Row);
                string targetSheetName = Convert.ToString(targetSheet.Name) ?? "";

                // 规则 8: 操作前强制修复并校准定义名称
                Tool.FixAndFillCabinetNamesForSheet(targetSheet);

                // 定位当前活动行所属的目标箱柜 (强类型接收避免 DLR 异常)
                CabinetRowContext? targetCab = FindCabinetByRow(targetSheet, targetRow);
                // 校验目标箱柜对象有效性
                if (targetCab == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "未检测到当前选区所属的箱柜信息，请在箱柜明细表内操作！",
                        "提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // 规范目标插入行位置：若选中小计行，则插入在小计行上方（作为箱柜最后一行元件）
                if (targetRow >= targetCab.SubsumRow)
                {
                    // 插入在小计行上方
                    targetRow = targetCab.SubsumRow;
                }
                else if (targetRow < targetCab.CompStartRow)
                {
                    // 若选中表头行或信息行，则插入为该箱柜的第一行元件
                    targetRow = targetCab.CompStartRow;
                }

                // 获取剪贴板数据实体
                // 获取剪贴板数据实体
                var clipData = ComponentClipboardManager.Get();
                if (clipData == null) return;

                // 读取待插入的元器件行数 (默认兼容单行为 1)
                int rowCount = clipData.RowCount > 0 ? clipData.RowCount : 1;

                // 判断是否为跨箱柜操作 (工作表不同，或箱柜序号 K 不同)
                bool isCrossCabinet = !string.Equals(clipData.SourceSheetName, targetSheetName, StringComparison.OrdinalIgnoreCase)
                                      || clipData.SourceCabinetK != targetCab.CabinetK;

                // 深度克隆待写入的二维数组矩阵 (规则 7 一次性写出)
                object[,] valuesToWrite = (object[,])clipData.FullRowValues.Clone();

                // 【核心指示落地】：复制剪切到不同的箱柜的时候 CadHandle 不要复制
                if (isCrossCabinet)
                {
                    int hIdx = CadHandleColIndex; // 1-based 下标
                    int hbIdx = CadHandleBColIndex;
                    int maxCols = valuesToWrite.GetLength(1);
                    // 批量清空全部待写入行的 CadHandle 与 HandleB
                    for (int r = 1; r <= rowCount; r++)
                    {
                        if (hIdx <= maxCols) valuesToWrite[r, hIdx] = null;
                        if (hbIdx <= maxCols) valuesToWrite[r, hbIdx] = null;
                    }
                }

                // 挂起屏幕刷新与自动计算以提速 COM 执行
                app.ScreenUpdating = false;
                app.Calculation = XlCalculation.xlCalculationManual;

                try
                {
                    // 在目标位置执行物理批量整块下推插入 (Shift Down，一次性插入 rowCount 行)
                    dynamic insertRows = targetSheet.Range[targetSheet.Rows[targetRow], targetSheet.Rows[targetRow + rowCount - 1]];
                    insertRows.Insert(XlInsertShiftDirection.xlShiftDown);

                    // 写入整行数据矩阵 (单次 COM 批量写入，规则 7)
                    int writeCols = valuesToWrite.GetLength(1);
                    dynamic insertRange = targetSheet.Range[targetSheet.Cells[targetRow, 1], targetSheet.Cells[targetRow + rowCount - 1, writeCols]];
                    insertRange.Value2 = valuesToWrite;

                    // 批量恢复自适应标准计算公式 (F 列数量 * G 列单价，F 列数量 * J 列成本单价)
                    for (int i = 0; i < rowCount; i++)
                    {
                        int curR = targetRow + i;
                        targetSheet.Cells[curR, 8].Formula = $"=F{curR}*G{curR}";
                        targetSheet.Cells[curR, 11].Formula = $"=F{curR}*J{curR}";
                    }

                    // 若原操作为【剪切】模式，安全删除来源行 (批量删除 rowCount 行)
                    if (clipData.IsCutMode)
                    {
                        // 来源与目标在同一张工作表，计算行号下移偏移量
                        if (string.Equals(clipData.SourceSheetName, targetSheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            int actualSourceRow = clipData.SourceRowIndex;
                            // 若来源行在插入点之后，因插入了 rowCount 行导致来源行号增加了 rowCount
                            if (actualSourceRow >= targetRow)
                            {
                                actualSourceRow += rowCount;
                            }

                            // 物理整块上移删除原剪切行 (一次性删除全部 rowCount 行)
                            targetSheet.Range[targetSheet.Rows[actualSourceRow], targetSheet.Rows[actualSourceRow + rowCount - 1]].Delete(XlDeleteShiftDirection.xlShiftUp);
                        }
                        else
                        {
                            // 跨工作表剪切，定位来源工作表并物理删除原连续行
                            try
                            {
                                dynamic srcSheet = targetSheet.Parent.Worksheets[clipData.SourceSheetName];
                                if (srcSheet != null)
                                {
                                    srcSheet.Range[srcSheet.Rows[clipData.SourceRowIndex], srcSheet.Rows[clipData.SourceRowIndex + rowCount - 1]].Delete(XlDeleteShiftDirection.xlShiftUp);
                                    // 刷新源工作表的序号与小计公式
                                    RefreshCabinetNumbersAndSubsumFormula(srcSheet);
                                    Tool.FixAndFillCabinetNamesForSheet(srcSheet, forceRebuild: true);
                                }
                            }
                            catch { }
                        }

                        // 剪切完成后彻底释放清空剪贴板
                        ComponentClipboardManager.Clear();
                    }

                    // 重新排布目标工作表所有箱柜的 A 列连续序号，并自适应刷新小计行 SUM 公式
                    RefreshCabinetNumbersAndSubsumFormula(targetSheet);

                    // 规则 8: 插入/删除行后强制调用 FixAndFillCabinetNamesForSheet 全量更新定义名称锚点
                    Tool.FixAndFillCabinetNamesForSheet(targetSheet, forceRebuild: true);

                    // 插入粘贴完成后，动态效果消失：清除橙色流动细虚线边框状态
                    try
                    {
                        // 停止并隐藏橙色流动细虚线动画
                        MarchingAntsManager.Hide();
                        // 同步将 CutCopyMode 状态重置为 0
                        app.CutCopyMode = (XlCutCopyMode)0;
                    }
                    catch { }

                    // 在状态栏反馈插入成功
                    string resultText = rowCount > 1
                        ? $"{rowCount} 行元件 (行 {targetRow}~{targetRow + rowCount - 1})"
                        : $"行 {targetRow}: {clipData.ComponentName}";
                    app.StatusBar = $"[插入元件成功] {resultText} (跨柜清空Handle: {isCrossCabinet})";
                }
                finally
                {
                    // 恢复自动计算与屏幕刷新
                    app.Calculation = XlCalculation.xlCalculationAutomatic;
                    app.ScreenUpdating = true;
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[ComponentRowOperations] 插入元件异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"插入元件失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 删除当前选中的元器件整行 (快捷键 Ctrl+Shift+D 或右键菜单【删除元件】触发)
        /// 核心保障：杜绝小计公式报 #REF!；自动重排序号；最后一行元件保留空行
        /// </summary>
        public static void DeleteComponentRow()
        {
            try
            {
                // 获取当前正在运行的 Excel Application
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return;

                // 优先获取选区对象，兼容单选单元格与连续多选行
                dynamic? selection = app.Selection;
                dynamic? activeCell = app.ActiveCell;
                if (selection == null && activeCell == null) return;

                // 获取所属工作表与待删除的起止行号
                dynamic sheet = (selection != null) ? selection.Worksheet : activeCell!.Worksheet;
                int startRow;
                int endRow;
                if (selection != null)
                {
                    startRow = Convert.ToInt32(selection.Row);
                    int selRowsCount = Convert.ToInt32(selection.Rows.Count);
                    endRow = startRow + selRowsCount - 1;
                }
                else
                {
                    startRow = Convert.ToInt32(activeCell!.Row);
                    endRow = startRow;
                }
                int deleteRowCount = endRow - startRow + 1;

                // 规则 8: 操作前校准定义名称
                Tool.FixAndFillCabinetNamesForSheet(sheet);

                // 校验并定位选区起止行所属箱柜
                CabinetRowContext? startCab = FindCabinetByRow(sheet, startRow);
                CabinetRowContext? endCab = (deleteRowCount == 1) ? startCab : FindCabinetByRow(sheet, endRow);

                // 校验是否在同一箱柜的有效元器件行内
                if (startCab == null || endCab == null || startCab.CabinetK != endCab.CabinetK
                    || startRow < startCab.CompStartRow || endRow > startCab.CompEndRow)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "请在同一箱柜的有效元器件数据行中执行【删除元件】！\n(严禁跨箱柜删除，且不能包含箱柜信息行、表头行、小计行或计费区域)",
                        "提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // 挂起刷新提升效率
                app.ScreenUpdating = false;
                app.Calculation = XlCalculation.xlCalculationManual;

                try
                {
                    // 检测当前箱柜内现有元器件的总行数
                    int totalCompCount = startCab.CompEndRow - startCab.CompStartRow + 1;

                    // 若选中的待删行数包含了该箱柜全部元器件
                    if (deleteRowCount >= totalCompCount)
                    {
                        // 规则 6: 保持箱柜基本骨架，不能将箱柜全部元件行物理删光导致小计紧挨表头
                        // 先将第 2 行至第 deleteRowCount 行删除，保留第 1 行并清空数据与公式
                        if (deleteRowCount > 1)
                        {
                            sheet.Range[sheet.Rows[startRow + 1], sheet.Rows[endRow]].Delete(XlDeleteShiftDirection.xlShiftUp);
                        }
                        int colCount = Math.Max(DefaultComponentColumnCount, GetSheetColumnCount(sheet));
                        object[,] blankRow = new object[1, colCount];
                        sheet.Range[sheet.Cells[startRow, 1], sheet.Cells[startRow, colCount]].Value2 = blankRow;

                        // 保留 A 列序号为 1
                        sheet.Cells[startRow, 1].Value2 = 1;
                    }
                    else
                    {
                        // 正常批量物理整块上移删除 (Shift Up)
                        sheet.Range[sheet.Rows[startRow], sheet.Rows[endRow]].Delete(XlDeleteShiftDirection.xlShiftUp);
                    }

                    // 重新排布箱柜元器件 A 列序号 1, 2, 3... 并自适应刷新小计行 SUM 公式
                    RefreshCabinetNumbersAndSubsumFormula(sheet);

                    // 规则 8: 删行后强制自愈定义名称
                    Tool.FixAndFillCabinetNamesForSheet(sheet, forceRebuild: true);

                    // 若当前处于复制/剪切动效状态，删行后同步清除橙色流动细虚线
                    try
                    {
                        // 停止并隐藏橙色流动细虚线
                        MarchingAntsManager.Hide();
                        // 重置 CutCopyMode 状态
                        app.CutCopyMode = (XlCutCopyMode)0;
                    }
                    catch { }

                    // 状态栏提示成功
                    string delDesc = deleteRowCount > 1
                        ? $"已批量移除 {deleteRowCount} 行元件 (行 {startRow}~{endRow})"
                        : $"已移除行 {startRow}";
                    app.StatusBar = $"[删除元件成功] {delDesc}";
                }
                finally
                {
                    // 恢复刷新与计算
                    app.Calculation = XlCalculation.xlCalculationAutomatic;
                    app.ScreenUpdating = true;
                }
            }
            catch (Exception ex)
            {
                // 记录删除异常日志
                LogHelper.WriteLog($"[ComponentRowOperations] 删除元件异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"删除元件失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 箱柜行属性分析辅助模型 (改用 class 引用类型，根治 dynamic 与 struct 比较时抛出的 RuntimeBinderException)
        /// </summary>
        private class CabinetRowContext
        {
            // 箱柜序号 K 编号
            public int CabinetK { get; set; }
            // 箱柜信息行 Cab_Det 行号
            public int DetRow { get; set; }
            // 元器件起始行号 (Cab_Det + 2)
            public int CompStartRow { get; set; }
            // 元器件终止行号 (Cab_Subsum - 1)
            public int CompEndRow { get; set; }
            // 小计行 Cab_Subsum 行号
            public int SubsumRow { get; set; }
            // 总计行 Cab_Tolsum 行号
            public int TolsumRow { get; set; }
            // 标记待查行是否落于合法元器件区间内
            public bool IsComponentRow { get; set; }
        }

        /// <summary>
        /// 根据绝对物理行号逆向定位所属的箱柜与边界范围
        /// </summary>
        private static CabinetRowContext? FindCabinetByRow(dynamic sheet, int row)
        {
            try
            {
                // 读取 4 种定义名称前缀
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;
                string sheetName = Convert.ToString(sheet.Name) ?? "";

                // 收集并构建当前表的所有箱柜锚点映射
                var allNames = Tool.CollectAllDefinedNames(sheet.Parent, sheet, autoRebuildIfEmpty: false);
                var cabinetMap = Tool.BuildCabinetMap(allNames, sheetName, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);

                if (cabinetMap == null || cabinetMap.Count == 0) return null;

                // 遍历每一个箱柜进行区间命中检测
                for (int i = 0; i < cabinetMap.Count; i++)
                {
                    var cab = cabinetMap[i];
                    if (cab.Value.Det == null || cab.Value.Subsum == null) continue;

                    int detRow = Convert.ToInt32(cab.Value.Det.Row);
                    int compStart = detRow + 2;
                    int subsumRow = Convert.ToInt32(cab.Value.Subsum.Row);
                    int compEnd = subsumRow - 1;
                    int tolsumRow = cab.Value.Tolsum != null ? Convert.ToInt32(cab.Value.Tolsum.Row) : subsumRow + 10;

                    // 计算该箱柜的起止行范围 (从 Det 行至 Tolsum 行或下一个箱柜 Det 行前)
                    int nextDetRow = (i + 1 < cabinetMap.Count && cabinetMap[i + 1].Value.Det != null)
                        ? Convert.ToInt32(cabinetMap[i + 1].Value.Det.Row)
                        : tolsumRow + 5;

                    // 检查行是否属于当前箱柜
                    if (row >= detRow && row < nextDetRow)
                    {
                        return new CabinetRowContext
                        {
                            CabinetK = cab.Key,
                            DetRow = detRow,
                            CompStartRow = compStart,
                            CompEndRow = compEnd,
                            SubsumRow = subsumRow,
                            TolsumRow = tolsumRow,
                            IsComponentRow = (row >= compStart && row <= compEnd)
                        };
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 重新自适应刷新当前工作表所有箱柜的 A 列连续序号，并重写小计行的 SUM 求和公式
        /// 彻底杜绝 #REF! 错误，并确保新插入的元器件 100% 纳入小计总额
        /// </summary>
        /// <param name="sheet">目标工作表</param>
        private static void RefreshCabinetNumbersAndSubsumFormula(dynamic sheet)
        {
            try
            {
                // 读取 4 种定义名称前缀
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;
                string sheetName = Convert.ToString(sheet.Name) ?? "";

                // 收集现有定义名称
                var allNames = Tool.CollectAllDefinedNames(sheet.Parent, sheet, autoRebuildIfEmpty: false);
                var cabinetMap = Tool.BuildCabinetMap(allNames, sheetName, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);

                if (cabinetMap == null) return;

                // 遍历各箱柜进行顺号与公式校准
                foreach (var cab in cabinetMap)
                {
                    if (cab.Value.Det == null || cab.Value.Subsum == null) continue;

                    int detRow = Convert.ToInt32(cab.Value.Det.Row);
                    int compStartRow = detRow + 2;
                    int subsumRow = Convert.ToInt32(cab.Value.Subsum.Row);
                    int compEndRow = subsumRow - 1;

                    int rowCount = compEndRow - compStartRow + 1;
                    if (rowCount > 0)
                    {
                        // 1. 一次性重排 A 列连续序号 1, 2, 3... (规则 7 二维数组写入)
                        object[,] seqValues = new object[rowCount, 1];
                        for (int idx = 0; idx < rowCount; idx++)
                        {
                            seqValues[idx, 0] = idx + 1;
                        }
                        sheet.Range[$"A{compStartRow}:A{compEndRow}"].Value2 = seqValues;

                        // 2. 重写真实小计公式，小计求和范围精准绑定 [compStartRow, compEndRow]
                        // H 列总价小计
                        sheet.Cells[subsumRow, 8].Formula = $"=SUM(H{compStartRow}:H{compEndRow})";
                        // K 列成本总价小计
                        sheet.Cells[subsumRow, 11].Formula = $"=SUM(K{compStartRow}:K{compEndRow})";
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录自愈刷新日志
                LogHelper.WriteLog($"[ComponentRowOperations] 刷新序号与小计公式异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 安全获取当前工作表的总有效列数
        /// </summary>
        private static int GetSheetColumnCount(dynamic sheet)
        {
            try
            {
                // 读取已用区域的列数
                return Convert.ToInt32(sheet.UsedRange.Columns.Count);
            }
            catch
            {
                // 容错返回默认列数
                return DefaultComponentColumnCount;
            }
        }
    }
}
