using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Forms;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 公共服务分部类：元器件拆分改型业务调度与物理表格处理
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，Excel 操作收敛于公共服务
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 供右键菜单或 Ribbon 调用的“拆分改型元件”入口方法
        /// </summary>
        public static void OpenComponentSplitDialog()
        {
            try
            {
                // 获取当前正在运行的 Excel Application 实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                // 校验工作簿是否已打开
                if (app == null || app.ActiveWorkbook == null)
                {
                    // 若未打开工作簿则弹出友好提示
                    MessageBox.Show("请先打开报价项目工作簿！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 获取当前活动单元格
                dynamic? activeCell = app.ActiveCell;
                if (activeCell == null) return;
                // 获取当前活动工作表句柄
                dynamic sheet = activeCell.Worksheet;
                // 获取当前活动物理行号 (1-based)
                int activeRow = Convert.ToInt32(activeCell.Row);

                // 规则 8 强约束: 在操作/提取分类明细表前，先调用 FixAndFillCabinetNamesForSheet 确保规则 6 定义名称与架构正确
                Tool.FixAndFillCabinetNamesForSheet(sheet, forceRebuild: false);

                // 根据当前行号逆向定位所属箱柜上下文模型
                CabinetRowContext? cab = FindCabinetByRow(sheet, activeRow);
                // 校验当前行是否落在合法元器件区间内 (规则 6: Cab_Det+2 至 Cab_Subsum-1)
                if (cab == null || !cab.IsComponentRow)
                {
                    // 若选中的是非元器件行，弹出警示对话框并安全退出
                    MessageBox.Show("请在箱柜明细表的有效元器件行中操作！\n(严禁在箱柜信息行、表头行、小计行或计费区域拆分)", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 提取箱柜名称 (从 DetRow 的 B 列读取)
                string cabName = Convert.ToString(sheet.Cells[cab.DetRow, 2]?.Value2)?.Trim() ?? $"箱柜{cab.CabinetK}";

                // 规则 7: 采用二维数组一次性读取当前行的 A~AF 列数据 (共 35 列)
                dynamic readRange = sheet.Range[sheet.Cells[activeRow, 1], sheet.Cells[activeRow, DefaultComponentColumnCount]];
                object[,] fullValues = (object[,])readRange.Value2;

                // 提取前台核心业务属性
                string indexVal = Convert.ToString(fullValues[1, 1])?.Trim() ?? string.Empty; // A 列序号
                string nameVal = Convert.ToString(fullValues[1, 2])?.Trim() ?? string.Empty; // B 列名称
                string modelVal = Convert.ToString(fullValues[1, 3])?.Trim() ?? string.Empty; // C 列规格型号
                string brandVal = Convert.ToString(fullValues[1, 4])?.Trim() ?? string.Empty; // D 列品牌
                string unitVal = Convert.ToString(fullValues[1, 5])?.Trim() ?? "只"; // E 列单位 --硬编码: 默认单位--
                if (string.IsNullOrWhiteSpace(unitVal)) unitVal = "只"; // 缺省保护

                // 提取 F 列数量 (第 6 列)
                double qty = 0;
                if (fullValues[1, 6] != null)
                {
                    // 优先使用 Convert.ToDouble，失败回退 TryParse
                    try
                    {
                        qty = Convert.ToDouble(fullValues[1, 6]);
                    }
                    catch
                    {
                        double.TryParse(Convert.ToString(fullValues[1, 6]), out qty);
                    }
                }

                // 提取 G 列单价 (第 7 列)
                decimal price = 0;
                if (fullValues[1, 7] != null)
                {
                    try
                    {
                        price = Convert.ToDecimal(fullValues[1, 7]);
                    }
                    catch
                    {
                        decimal.TryParse(Convert.ToString(fullValues[1, 7]), out price);
                    }
                }

                // 提取 J 列成本单价 (第 10 列)
                decimal costPrice = 0;
                if (fullValues[1, 10] != null)
                {
                    try
                    {
                        costPrice = Convert.ToDecimal(fullValues[1, 10]);
                    }
                    catch
                    {
                        decimal.TryParse(Convert.ToString(fullValues[1, 10]), out costPrice);
                    }
                }

                // 校验元器件有效性：若名称与型号均为空则判定为空行
                if (string.IsNullOrWhiteSpace(modelVal) && string.IsNullOrWhiteSpace(nameVal))
                {
                    MessageBox.Show("当前选中的是空行，无法执行元器件拆分！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 校验元器件数量：数量必须大于 0
                if (qty <= 0)
                {
                    MessageBox.Show("当前元器件数量必须大于 0 才能执行拆分！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 提取 AD 列 (第 30 列) CAD 句柄扩展列
                string rawHandles = string.Empty;
                if (fullValues.GetLength(1) >= CadHandleColIndex)
                {
                    // 读取 AD 列句柄文本
                    rawHandles = Convert.ToString(fullValues[1, CadHandleColIndex])?.Trim() ?? string.Empty;
                }

                // 解析提取出独立的 Handle 列表
                var handleList = new List<string>();
                if (!string.IsNullOrEmpty(rawHandles))
                {
                    // 兼容两级分隔符 (逗号、分号与换行) 进行拆分
                    string[] parts = rawHandles.Split(new[] { ',', ';', ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var p in parts)
                    {
                        // 去除多余空格
                        string cleanP = p.Trim();
                        if (!string.IsNullOrEmpty(cleanP) && !handleList.Contains(cleanP))
                        {
                            // 记录唯一句柄
                            handleList.Add(cleanP);
                        }
                    }
                }

                // 记录详细提取上下文日志
                LogHelper.WriteLog($"[OpenComponentSplitDialog] 行:{activeRow}, 箱柜:{cabName}, 名称:{nameVal}, 型号:{modelVal}, 数量:{qty}, Handles:{handleList.Count}");

                // 构造前端数据传输上下文 DTO
                var candidate = new ComponentSplitCandidateDto
                {
                    SheetName = Convert.ToString(sheet.Name) ?? string.Empty,
                    RowIndex = activeRow,
                    CabinetK = cab.CabinetK,
                    CabinetName = cabName,
                    CompStartRow = cab.CompStartRow,
                    CompEndRow = cab.CompEndRow,
                    SubsumRow = cab.SubsumRow,
                    Index = indexVal,
                    Name = nameVal,
                    Model = modelVal,
                    Brand = brandVal,
                    Unit = unitVal,
                    Quantity = qty,
                    Price = price,
                    CostPrice = costPrice,
                    RawHandles = rawHandles,
                    Handles = handleList
                };

                // 打开拆分改型微型核对工作台
                ComponentSplitForm.ShowForm(candidate);
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[ExcelServices] OpenComponentSplitDialog 异常: {ex.Message}");
                // 弹出异常提示对话框
                MessageBox.Show($"打开元件拆分窗口失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 执行元器件拆分改型提交：原行扣减数量和句柄 + 新行生成与价格匹配 + 规则 6 动态插行与公式重算
        /// </summary>
        /// <param name="req">拆分执行请求数据包</param>
        /// <returns>执行结果模型</returns>
        public static ComponentSplitResult ExecuteComponentSplit(ComponentSplitSubmitRequest req)
        {
            var result = new ComponentSplitResult();
            // 参数合法性校验
            if (req == null)
            {
                result.Success = false;
                result.Message = "请求参数为空！";
                return result;
            }

            try
            {
                // 获取当前正在运行的 Excel Application
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveWorkbook == null)
                {
                    result.Success = false;
                    result.Message = "未检测到活动工作簿！";
                    return result;
                }

                // 获取目标工作表
                dynamic sheet = app.ActiveWorkbook.Worksheets[req.SheetName];
                if (sheet == null)
                {
                    result.Success = false;
                    result.Message = $"未找到工作表: {req.SheetName}";
                    return result;
                }

                // 规则 8: 操作前强制预检自愈
                Tool.FixAndFillCabinetNamesForSheet(sheet, forceRebuild: false);

                // 重新检索校验当前行所属箱柜
                int origRow = req.RowIndex;
                CabinetRowContext? cab = FindCabinetByRow(sheet, origRow);
                if (cab == null || !cab.IsComponentRow)
                {
                    result.Success = false;
                    result.Message = "目标行已失效或不再属于合法元器件区间，请重试！";
                    return result;
                }

                // 检查该箱柜有效元器件区间内是否有可复用的空行 (规则 6: 元器件区域可以有空行，计费区域不能有空行)
                int compStart = cab.CompStartRow;
                int compEnd = cab.CompEndRow;
                int lastEmptyRow = -1;

                // 规则 7: 一次性读取该箱柜所有元器件行的 B 列(名称)与 C 列(型号)
                dynamic checkRange = sheet.Range[sheet.Cells[compStart, 2], sheet.Cells[compEnd, 3]];
                object[,] nameModelMatrix = (object[,])checkRange.Value2;
                int matrixRows = nameModelMatrix.GetLength(0);

                // 从后往前逆向寻找未使用的空行
                for (int r = matrixRows; r >= 1; r--)
                {
                    string nm = Convert.ToString(nameModelMatrix[r, 1])?.Trim() ?? string.Empty;
                    string md = Convert.ToString(nameModelMatrix[r, 2])?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(nm) && string.IsNullOrEmpty(md))
                    {
                        // 记录空行的物理行号
                        lastEmptyRow = compStart + r - 1;
                        break;
                    }
                }

                // 备份原始状态以便退出时安全恢复
                bool prevUpdating = false;
                XlCalculation prevCalc = XlCalculation.xlCalculationAutomatic;
                bool prevEvents = true;
                bool prevAlerts = true;
                try
                {
                    prevUpdating = Convert.ToBoolean(app.ScreenUpdating);
                    prevCalc = (XlCalculation)app.Calculation;
                    prevEvents = Convert.ToBoolean(app.EnableEvents);
                    prevAlerts = Convert.ToBoolean(app.DisplayAlerts);
                }
                catch { }

                // 挂起 Excel 刷新加速并防止屏幕闪烁
                app.ScreenUpdating = false;
                app.Calculation = XlCalculation.xlCalculationManual;
                app.EnableEvents = false;
                app.DisplayAlerts = false;

                // 规定新行插入位置：就在当前原行的紧邻正下方
                int newRow = origRow + 1;

                try
                {
                    // 在原行正下方执行物理插行 (Shift Down 向下平移一行)
                    dynamic insertRange = sheet.Rows[newRow];
                    insertRange.Insert(XlInsertShiftDirection.xlShiftDown);

                    // 若原箱柜元器件区间内存在未使用的预留空行，为了维持箱柜小计行及下方计费区物理位置稳定，删除末尾的一个空行
                    if (lastEmptyRow > 0)
                    {
                        // 修正空行行号：若空行在插入点之后，因插入了 1 行导致其行号顺移了 1 行
                        int actualEmptyRow = lastEmptyRow >= newRow ? lastEmptyRow + 1 : lastEmptyRow;
                        // 物理删除多余的空行
                        dynamic emptyRowRange = sheet.Rows[actualEmptyRow];
                        emptyRowRange.Delete(XlDeleteShiftDirection.xlShiftUp);
                    }

                    // 规则 7: 更新原行与新行的数据矩阵
                    // 1. 更新原行数据 (保留项)
                    sheet.Cells[origRow, 6].Value2 = req.RemainQuantity; // F 列数量
                    string remainHandleStr = (req.RemainHandles != null && req.RemainHandles.Count > 0)
                        ? string.Join(",", req.RemainHandles)
                        : string.Empty;
                    sheet.Cells[origRow, CadHandleColIndex].Value2 = remainHandleStr; // AD 列
                    sheet.Cells[origRow, 8].Formula = $"=F{origRow}*G{origRow}"; // H 列总价公式
                    sheet.Cells[origRow, 11].Formula = $"=F{origRow}*J{origRow}"; // K 列成本总价公式

                    // 2. 组装新行整行数据矩阵 (A~AF 列)
                    object[,] newRowMatrix = new object[1, DefaultComponentColumnCount];
                    newRowMatrix[0, 0] = string.Empty; // A 列序号 (后续统一刷新)
                    newRowMatrix[0, 1] = req.NewName; // B 列名称
                    newRowMatrix[0, 2] = req.NewModel; // C 列规格型号
                    newRowMatrix[0, 3] = req.NewBrand; // D 列品牌
                    newRowMatrix[0, 4] = req.NewUnit; // E 列单位
                    newRowMatrix[0, 5] = req.SplitQuantity; // F 列数量
                    newRowMatrix[0, 6] = req.NewPrice; // G 列单价
                    newRowMatrix[0, 8] = 1.0; // I 列折扣率 (默认 1.0)
                    newRowMatrix[0, 9] = req.NewCostPrice; // J 列成本单价

                    // 写入 AD 列 CAD 句柄
                    string splitHandleStr = (req.SplitHandles != null && req.SplitHandles.Count > 0)
                        ? string.Join(",", req.SplitHandles)
                        : string.Empty;
                    if (DefaultComponentColumnCount >= CadHandleColIndex)
                    {
                        newRowMatrix[0, CadHandleColIndex - 1] = splitHandleStr; // AD 列
                    }

                    // 一次性批量写入新行数据
                    dynamic writeRange = sheet.Range[sheet.Cells[newRow, 1], sheet.Cells[newRow, DefaultComponentColumnCount]];
                    writeRange.Value2 = newRowMatrix;

                    // 设置新行自适应联动公式
                    sheet.Cells[newRow, 8].Formula = $"=F{newRow}*G{newRow}"; // H 列总价公式
                    sheet.Cells[newRow, 11].Formula = $"=F{newRow}*J{newRow}"; // K 列成本总价公式

                    // 刷新该箱柜的连续序号与小计 SUM 求和公式
                    RefreshCabinetNumbersAndSubsumFormula(sheet);

                    // 再次自愈定义名称，确保规则 6 架构健全
                    Tool.FixAndFillCabinetNamesForSheet(sheet, forceRebuild: false);

                    // 若用户勾选了同步更新 CAD，异步向 CAD 管道推送更新元器件规格型号指令
                    if (req.SyncToCad && req.SplitHandles != null && req.SplitHandles.Count > 0)
                    {
                        CadSyncClient.SendUpdateComponentSpec(req.SplitHandles, req.NewModel);
                    }

                    // 组装成功结果
                    result.Success = true;
                    result.Message = $"拆分成功：原规格保留 {req.RemainQuantity} 只，新规格【{req.NewModel}】{req.SplitQuantity} 只！";
                    result.InsertedRowIndex = newRow;
                }
                finally
                {
                    // 恢复原始警告提示状态
                    try { app.DisplayAlerts = prevAlerts; } catch { }
                    // 恢复原始系统事件响应状态
                    try { app.EnableEvents = prevEvents; } catch { }
                    // 恢复原始公式计算模式
                    try { app.Calculation = prevCalc; } catch { }
                    // 恢复屏幕刷新
                    try { app.ScreenUpdating = prevUpdating; } catch { }
                }
            }
            catch (Exception ex)
            {
                // 记录执行异常日志
                LogHelper.WriteLog($"[ExcelServices] ExecuteComponentSplit 异常: {ex.Message}");
                result.Success = false;
                result.Message = $"拆分执行失败: {ex.Message}";
            }

            return result;
        }
    }
}
