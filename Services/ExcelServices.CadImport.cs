using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 公共服务分部类: CAD 扒图元器件直接录入/反选同步服务
    /// 遵循成套电气分类表三大黄金法则：
    /// 规则 6: 四位一体定义名称与小计行前物理插行
    /// 规则 7: 二维数组一次性批量读写
    /// 规则 8: 操作前调用 Tool.FixAndFillCabinetNamesForSheet 进行结构自愈
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 将 CAD 提取的元器件直接导入/追加到 Excel 当前光标所在的活动箱柜中 (或处理复选反选扣减)
        /// 遵循用户决策: 1B(名称保持原样) + 2B(未选中箱柜严格拦截) + 3A(纯净轻量提取，不自动查价)
        /// </summary>
        /// <param name="payload">CAD 发送的数据交互载荷</param>
        /// <returns>操作执行结果对象</returns>
        public static CadImportResult ImportComponentsToActiveCabinet(CadImportComponentPayload payload)
        {
            // 校验入参有效性
            if (payload == null)
            {
                // 参数为空返回失败
                return CadImportResult.Fail("INVALID_PARAM", "传入的元器件导入载荷为空！");
            }

            // 1. 获取运行中的 Excel Application 实例
            dynamic? app = ExcelDnaSafeAccessor.GetApplication();
            if (app == null)
            {
                try
                {
                    // 尝试从 Windows ROT 运行对象表中获取正在运行的 Excel
                    app = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch { }
            }

            // 若仍无法获取 Excel Application 接口则返回异常提示
            if (app == null)
            {
                // 提示未检测到 Excel 实例
                return CadImportResult.Fail("NO_EXCEL", "未检测到运行中的 Excel 实例，请先打开报价工作簿！");
            }

            // 2. 获取活动工作表
            dynamic? activeSheet = null;
            try
            {
                // 读取当前活动的 Worksheet 对象
                activeSheet = app.ActiveSheet;
            }
            catch (Exception exSheet)
            {
                // 记录工作表获取异常
                LogHelper.WriteLog($"[CadImport] 获取活动工作表异常: {exSheet.Message}");
            }

            if (activeSheet == null)
            {
                // 提示无活动工作表
                return CadImportResult.Fail("NO_ACTIVE_SHEET", "当前 Excel 中没有活动的分类工作表！");
            }

            // 3. 执行规则 8 前置自愈守门：检查并修复定义名称与超链接
            try
            {
                // 确保规则 6 架构完整无损
                Tool.FixAndFillCabinetNamesForSheet(activeSheet, forceRebuild: false);
            }
            catch (Exception exFix)
            {
                // 记录自愈日志
                LogHelper.WriteLog($"[CadImport] 规则 8 自愈检查警告: {exFix.Message}");
            }

            // 4. 获取当前选区/活动单元格所在绝对物理行号
            int activeRow = 0;
            try
            {
                // 优先从 ActiveCell 读取行号
                if (app.ActiveCell != null)
                {
                    activeRow = Convert.ToInt32(app.ActiveCell.Row);
                }
                // 其次从 Selection 读取行号
                else if (app.Selection != null)
                {
                    activeRow = Convert.ToInt32(app.Selection.Row);
                }
            }
            catch { }

            // 5. 逆向定位当前选区所在的箱柜上下文 (FindCabinetByRow)
            CabinetRowContext? cab = (activeRow > 0) ? FindCabinetByRow(activeSheet, activeRow) : null;

            // 严格遵循用户决议 2B：若光标不在任何有效箱柜内，严格拦截提醒，绝不盲目追加
            if (cab == null)
            {
                // 记录拦截日志
                LogHelper.WriteLog($"[CadImport] 当前行 {activeRow} 不属于任何箱柜，执行 2B 拦截策略");
                // 返回明确的状态码供 CAD 命令行打印醒目提示
                return CadImportResult.Fail("NOT_IN_CABINET", "未在 Excel 中检测到选中的箱柜，请先点击 Excel 目标箱柜行！");
            }

            // 提取目标箱柜的关键行号与物理边界
            int detRow = cab.DetRow;
            int compStartRow = cab.CompStartRow;
            int subsumRow = cab.SubsumRow;
            int compEndRow = cab.CompEndRow;
            int tolsumRow = cab.TolsumRow;

            // 读取箱柜名称 (柜号)
            string cabName = string.Empty;
            try
            {
                // B 列第 2 列为箱柜名称
                cabName = Convert.ToString(activeSheet.Cells[detRow, 2].Value)?.Trim() ?? string.Empty;
            }
            catch { }
            if (string.IsNullOrWhiteSpace(cabName))
            {
                // 兜底箱柜名称
                cabName = $"箱柜{cab.CabinetK}";
            }

            // 6. 环境压栈保护：静默刷新以保障极致写入性能
            bool prevUpdating = true;
            bool prevAlerts = true;
            bool prevEvents = true;
            int prevCalculation = -4105; // 默认 xlCalculationAutomatic (-4105)
            try { prevUpdating = app.ScreenUpdating; } catch { }
            try { prevAlerts = app.DisplayAlerts; } catch { }
            try { prevEvents = app.EnableEvents; } catch { }
            try { prevCalculation = Convert.ToInt32(app.Calculation); } catch { }

            try
            {
                // 关闭屏幕刷新与事件驱动
                try { app.ScreenUpdating = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }
                try { app.EnableEvents = false; } catch { }
                try { app.Calculation = -4135; /* xlCalculationManual */ } catch { }

                // =========================================================================
                // 分支 A：CAD 端复选反选取消选择 (isDelete == true)
                // =========================================================================
                if (payload.IsDelete)
                {
                    // 提取所有待扣减图元的 Handle 集合
                    var targetHandles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    if (payload.Components != null)
                    {
                        foreach (var c in payload.Components)
                        {
                            if (string.IsNullOrWhiteSpace(c?.Handle)) continue;
                            // 兼容多个句柄以连字符或逗号拼接的情况
                            var parts = c.Handle.Split(new char[] { '-', ',' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var p in parts) targetHandles.Add(p.Trim());
                        }
                    }

                    int affectedDeleteCount = 0;
                    // 规则 7：二维数组一次性读取当前箱柜的元器件区域 (A 列至 AD 列共 30 列)
                    int capacity = compEndRow - compStartRow + 1;
                    if (capacity > 0 && targetHandles.Count > 0)
                    {
                        dynamic compRange = activeSheet.Range[activeSheet.Cells[compStartRow, 1], activeSheet.Cells[compEndRow, 30]];
                        object[,] matrix = (object[,])compRange.Value2;

                        // 逆向自底向上遍历行，防止删行导致物理行号漂移
                        for (int r = capacity; r >= 1; r--)
                        {
                            int currentPhysRow = compStartRow + r - 1;
                            // AD 列为第 30 列 (CadHandle 列号) --硬编码: AD列为第30列--
                            string rowHandleStr = Convert.ToString(matrix[r, 30])?.Trim() ?? string.Empty;
                            if (string.IsNullOrWhiteSpace(rowHandleStr)) continue;

                            // 检查当前行是否包含待取消图元的句柄
                            var rowHandles = rowHandleStr.Split(new char[] { '-', ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                         .Select(h => h.Trim())
                                                         .ToList();
                            bool hit = rowHandles.Any(h => targetHandles.Contains(h));
                            if (!hit) continue;

                            // 提取当前行数量 (F 列第 6 列)
                            decimal currentQty = 1;
                            try
                            {
                                currentQty = Convert.ToDecimal(matrix[r, 6]);
                            }
                            catch { }

                            // 若当前数量大于 1：执行数量减 1 并剔除已取消的 Handle
                            if (currentQty > 1)
                            {
                                // F 列数量扣减 1
                                activeSheet.Cells[currentPhysRow, 6].Value2 = currentQty - 1;
                                // 移除命中的句柄
                                rowHandles.RemoveAll(h => targetHandles.Contains(h));
                                // 回写更新后的 AD 列句柄
                                activeSheet.Cells[currentPhysRow, 30].Value2 = string.Join("-", rowHandles);
                                affectedDeleteCount++;
                            }
                            // 若数量等于 1 或更小：安全物理删除该行
                            else
                            {
                                // 物理删除该行并向上推移
                                activeSheet.Rows[currentPhysRow].Delete(-4119 /* xlShiftUp */);
                                // 更新后续小计行号与总计行号
                                subsumRow--;
                                compEndRow--;
                                tolsumRow--;
                                affectedDeleteCount++;
                            }
                        }

                        // 重新更新当前箱柜的小计求和公式
                        if (compEndRow >= compStartRow)
                        {
                            // H 列总价求和
                            activeSheet.Cells[subsumRow, 8].Formula = $"=ROUND(SUM(H{compStartRow}:H{compEndRow}), 2)";
                            // K 列成本求和
                            activeSheet.Cells[subsumRow, 11].Formula = $"=ROUND(SUM(K{compStartRow}:K{compEndRow}), 2)";
                        }
                    }

                    // 返回反选成功结果
                    return CadImportResult.Ok($"已从箱柜【{cabName}】成功反选扣减/移除 {affectedDeleteCount} 个元器件！", affectedDeleteCount, cabName);
                }

                // =========================================================================
                // 分支 B：追加新增元器件 (isDelete == false)
                // =========================================================================
                int newCompCount = payload.Components?.Count ?? 0;
                if (newCompCount == 0)
                {
                    // 无待写入元件直接返回
                    return CadImportResult.Fail("EMPTY_LIST", "待导入的元器件列表为空！");
                }

                // 规则 7：单次读取当前箱柜 A 列至 AD 列，探测有效占用行数与剩余空行
                int totalRows = compEndRow - compStartRow + 1;
                dynamic curRange = activeSheet.Range[activeSheet.Cells[compStartRow, 1], activeSheet.Cells[compEndRow, 30]];
                object[,] curMatrix = (object[,])curRange.Value2;

                int usedCount = 0;
                for (int r = 1; r <= totalRows; r++)
                {
                    // B 列 (名称) 或 C 列 (型号) 非空判定为有效元器件行
                    string bVal = Convert.ToString(curMatrix[r, 2])?.Trim() ?? string.Empty;
                    string cVal = Convert.ToString(curMatrix[r, 3])?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(bVal) || !string.IsNullOrWhiteSpace(cVal))
                    {
                        usedCount = r;
                    }
                }

                // 计算当前剩余可用空余行数
                int remainingEmpty = totalRows - usedCount;

                // 规则 6 容量保障：若新元件数量多于空余行数，必须在小计行前物理插入差额行
                if (newCompCount > remainingEmpty)
                {
                    // 计算需要向下物理插行的差额行数
                    int insertCount = newCompCount - remainingEmpty;
                    // 在小计行 subsumRow 处向下插行
                    activeSheet.Rows[$"{subsumRow}:{subsumRow + insertCount - 1}"].Insert(-4121 /* xlShiftDown */);

                    // 同步修正小计行号与总计行号
                    subsumRow += insertCount;
                    compEndRow += insertCount;
                    tolsumRow += insertCount;
                }

                // 计算本次待写入的物理起止行
                int writeStartRow = compStartRow + usedCount;
                int writeEndRow = writeStartRow + newCompCount - 1;

                // 规则 7 构造元器件二维写入矩阵 (A~U 列共 21 列，外加 AD 列句柄)
                object[,] compDataMatrix = new object[newCompCount, 21];
                object[,] handleMatrix = new object[newCompCount, 1];

                for (int i = 0; i < newCompCount; i++)
                {
                    var item = payload.Components![i];
                    int currPhysicalRow = writeStartRow + i;

                    // A 列: 序号动态公式
                    compDataMatrix[i, 0] = $"=ROW()-ROW(A${detRow + 1})";
                    // B 列: 元件名称 (严格遵循 1B: 保持 TuFan 传入原样，不自作主张推导)
                    compDataMatrix[i, 1] = item.Name ?? string.Empty;
                    // C 列: 规格型号
                    compDataMatrix[i, 2] = item.Specification ?? string.Empty;
                    // D 列: 生产厂家 (严格遵循 3A: 纯净轻量提取，保持空)
                    compDataMatrix[i, 3] = string.Empty;
                    // E 列: 计量单位 (默认写入"只" --硬编码: 默认单位--)
                    compDataMatrix[i, 4] = string.IsNullOrWhiteSpace(item.Unit) ? "只" : item.Unit;
                    // F 列: 数量 (最小为 1)
                    compDataMatrix[i, 5] = item.Quantity > 0 ? item.Quantity : 1;
                    // G 列: 销售单价 (严格遵循 3A: 保持空，由用户在 Excel 自行查价)
                    compDataMatrix[i, 6] = string.Empty;
                    // H 列: 销售合价公式
                    compDataMatrix[i, 7] = $"=ROUND(F{currPhysicalRow}*G{currPhysicalRow}, 2)";
                    // I 列: 备注
                    compDataMatrix[i, 8] = item.Remark ?? string.Empty;
                    // J 列: 成本单价 (保持空)
                    compDataMatrix[i, 9] = string.Empty;
                    // K 列: 成本合价公式
                    compDataMatrix[i, 10] = $"=ROUND(F{currPhysicalRow}*J{currPhysicalRow}, 2)";

                    // AD 列: CAD 图元 Handles
                    handleMatrix[i, 0] = item.Handle ?? string.Empty;
                }

                // 规则 7 单次 Range 批量回写 A~U 列公式与文本
                activeSheet.Range[$"A{writeStartRow}:U{writeEndRow}"].Formula = compDataMatrix;
                // 单次 Range 批量写入 AD 列 CAD 图元句柄 (原 AA 列扩展) --硬编码: AD列--
                activeSheet.Range[$"AD{writeStartRow}:AD{writeEndRow}"].Value2 = handleMatrix;

                // 重新校准当前箱柜小计行求和公式
                activeSheet.Cells[subsumRow, 8].Formula = $"=ROUND(SUM(H{compStartRow}:H{compEndRow}), 2)";
                activeSheet.Cells[subsumRow, 11].Formula = $"=ROUND(SUM(K{compStartRow}:K{compEndRow}), 2)";

                // 选中新写入的首行元器件，便于用户在 Excel 视口中直观看到
                try
                {
                    activeSheet.Cells[writeStartRow, 3].Select();
                }
                catch { }

                // 返回追加成功结果
                return CadImportResult.Ok($"成功将 {newCompCount} 个元器件追加至箱柜【{cabName}】！", newCompCount, cabName);
            }
            catch (Exception ex)
            {
                // 记录执行异常日志
                LogHelper.WriteLog($"[CadImport] 写入当前箱柜发生异常: {ex.Message}");
                return CadImportResult.Fail("ERROR", $"写入当前箱柜异常: {ex.Message}");
            }
            finally
            {
                // 环境状态安全恢复
                try { app.Calculate(); } catch { }
                try { app.Calculation = prevCalculation; } catch { }
                try { app.ScreenUpdating = prevUpdating; } catch { }
                try { app.DisplayAlerts = prevAlerts; } catch { }
                try { app.EnableEvents = prevEvents; } catch { }
            }
        }

        /// <summary>
        /// 从 CAD 提取的整张系统图数据直接在 Excel 当前分类表中新建一台箱柜入表
        /// 对应右键【抓系统图】(zxtn) 动作
        /// </summary>
        /// <param name="payload">系统图完整载荷</param>
        /// <returns>新建箱柜执行结果</returns>
        public static CadImportResult CreateNewCabinetFromCad(CadImportComponentPayload payload)
        {
            if (payload == null)
            {
                return CadImportResult.Fail("INVALID_PARAM", "传入的系统图载荷为空！");
            }

            // 获取运行中的 Excel Application 实例
            dynamic? app = ExcelDnaSafeAccessor.GetApplication();
            if (app == null)
            {
                try
                {
                    app = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch { }
            }

            if (app == null)
            {
                return CadImportResult.Fail("NO_EXCEL", "未检测到运行中的 Excel 实例，请先打开报价工作簿！");
            }

            // 转换数据为标准 CabinetObject 实体
            var cabObj = new CabinetObject
            {
                Header = new CabinetHeader
                {
                    Name = !string.IsNullOrWhiteSpace(payload.CabinetName) ? payload.CabinetName.Trim() : "新箱柜",
                    CabinetNo = !string.IsNullOrWhiteSpace(payload.CabinetName) ? payload.CabinetName.Trim() : "新箱柜",
                    Model = payload.CabinetName ?? string.Empty,
                    InstallMode = payload.InstallMode ?? string.Empty,
                    Remark = payload.InstallMode ?? string.Empty,
                    Category = payload.BuildPosition ?? string.Empty,
                    Quantity = payload.CabinetQuantity > 0 ? payload.CabinetQuantity : 1,
                    MinMaxPoints = payload.CabinetMinMaxPoints ?? new List<string>()
                }
            };

            // 填充元器件明细项
            if (payload.Components != null)
            {
                int subIdx = 1;
                foreach (var c in payload.Components)
                {
                    cabObj.Components.Add(new ComponentItem
                    {
                        Index = subIdx++,
                        Name = c.Name ?? string.Empty,
                        Specification = c.Specification ?? string.Empty,
                        // 计量单位默认为 "只" --硬编码: 元件默认单位--
                        Unit = string.IsNullOrWhiteSpace(c.Unit) ? "只" : c.Unit,
                        Quantity = c.Quantity > 0 ? c.Quantity : 1,
                        Handle = c.Handle ?? string.Empty,
                        Remark = c.Remark ?? string.Empty
                    });
                }
            }

            // 构造批量集合并调用已有的成熟高效批量导出服务
            var cabList = new List<CabinetObject> { cabObj };
            int count = BatchExportCabinets(app, cabList, null);

            if (count > 0)
            {
                return CadImportResult.Ok($"成功在 Excel 中新建箱柜【{cabObj.Header.Name}】并录入 {cabObj.Components.Count} 个元器件！", cabObj.Components.Count, cabObj.Header.Name);
            }
            else
            {
                return CadImportResult.Fail("CREATE_FAILED", $"在 Excel 中新建箱柜【{cabObj.Header.Name}】失败！");
            }
        }
    }
}
