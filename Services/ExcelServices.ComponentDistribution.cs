using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelAddInDemo.Forms;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：元器件分布穿透查询与定位
    /// 遵循规则 3: 所有 Excel 操作全部封装于 ExcelServices，不写在窗口类中
    /// </summary>
    public static partial class ExcelServices
    {
        // 元器件分布穿透窗口单例句柄 (非模态)
        private static ComponentDistributionForm? _componentDistributionForm;

        /// <summary>
        /// 从当前活动工作表(元件汇总表)指定行中提取元器件特征查询入参
        /// 遵循规则 7: 单行多列一次性以二维数组读取至内存
        /// </summary>
        public static ComponentDistributionQuery? ExtractComponentQueryFromSummarySheet(int summaryRow)
        {
            try
            {
                // 获取 Excel 顶层 Application 实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return null;

                // 获取当前活动工作簿与工作表
                dynamic? activeSheet = app.ActiveSheet;
                if (activeSheet == null) return null;

                // 校验工作表名称是否为元件汇总表或包含“元件汇总”
                string sheetName = Convert.ToString(activeSheet.Name)?.Trim() ?? string.Empty;
                if (!string.Equals(sheetName, ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase) &&
                    !sheetName.Contains("元件汇总"))
                {
                    // 非元件汇总表直接返回空
                    return null;
                }

                // 行号必须有效
                if (summaryRow <= 0) return null;

                // 规则 7: 一次性读取该行 A~U 列(共 21 列)至内存二维数组
                dynamic rowRange = activeSheet.Range[$"A{summaryRow}:U{summaryRow}"];
                object[,]? rowMatrix = rowRange?.Value2 as object[,];
                if (rowMatrix == null) return null;

                // 提取 B 列元件名称 (索引 2)
                string compName = Convert.ToString(rowMatrix[1, 2])?.Trim() ?? string.Empty;

                // 提取 C 列原型号规格 (索引 3)
                string origModel = rowMatrix.GetLength(1) >= 3 ? (Convert.ToString(rowMatrix[1, 3])?.Trim() ?? string.Empty) : string.Empty;

                // 提取 D 列新型号规格 (索引 4)
                string newModel = rowMatrix.GetLength(1) >= 4 ? (Convert.ToString(rowMatrix[1, 4])?.Trim() ?? string.Empty) : string.Empty;

                // 遇到表头关键字、合计、总计行则拦截 --硬编码: 汇总截断关键字--
                if (compName.Contains("序号") || compName.Contains("元件名称") || compName.Contains("设备名称") ||
                    origModel.Contains("型号规格") || newModel.Contains("型号规格") ||
                    compName.Contains("合计") || compName.Contains("总计") || compName.Contains("小计"))
                {
                    return null;
                }

                // 若名称与两个型号均为空，判定为空行
                if (string.IsNullOrWhiteSpace(compName) && string.IsNullOrWhiteSpace(origModel) && string.IsNullOrWhiteSpace(newModel))
                {
                    return null;
                }

                // 提取 E 列单位 (索引 5)
                string unit = Convert.ToString(rowMatrix[1, 5])?.Trim() ?? string.Empty;

                // 提取 F 列汇总数量 (索引 6)
                decimal sumQty = 0;
                object qVal = rowMatrix[1, 6];
                if (qVal != null && decimal.TryParse(Convert.ToString(qVal), out decimal parsedQty))
                {
                    sumQty = parsedQty;
                }

                // 提取 I 列生产厂家 (索引 9)
                string mfg = Convert.ToString(rowMatrix[1, 9])?.Trim() ?? string.Empty;

                // 提取 R 列原始型号 (索引 18，对应分类明细表 U 列原始串)
                string rawModelU = string.Empty;
                if (rowMatrix.GetLength(1) >= 18)
                {
                    rawModelU = Convert.ToString(rowMatrix[1, 18])?.Trim() ?? string.Empty;
                }

                // 规则 7: 从隐藏单元格 AB1 读取持久化的源分类明细信息 (SelectedSheets 列表)
                var sourceSheets = new List<string>();
                try
                {
                    // 读取 AB1 单元格内容
                    string rawConfig = Convert.ToString(activeSheet.Range["AB1"].Value2)?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(rawConfig))
                    {
                        // 优先尝试 JSON 反序列化提取源分类表列表
                        try
                        {
                            using var doc = System.Text.Json.JsonDocument.Parse(rawConfig);
                            // 1. 若根节点本身即为表名数组
                            if (doc.RootElement.ValueKind == System.Text.Json.JsonValueKind.Array)
                            {
                                foreach (var item in doc.RootElement.EnumerateArray())
                                {
                                    // 提取工作表名
                                    string sName = item.GetString()?.Trim() ?? string.Empty;
                                    if (!string.IsNullOrWhiteSpace(sName) && !sourceSheets.Contains(sName))
                                    {
                                        sourceSheets.Add(sName);
                                    }
                                }
                            }
                            // 2. 若根节点为配置对象，检索包含 sheet 关键字的数组属性
                            else if (doc.RootElement.ValueKind == System.Text.Json.JsonValueKind.Object)
                            {
                                foreach (var prop in doc.RootElement.EnumerateObject())
                                {
                                    // 属性名包含 sheet 且值为数组 (如 SelectedSheets, selectedSheets, sourceSheets 等)
                                    if (prop.Name.IndexOf("sheet", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                        prop.Value.ValueKind == System.Text.Json.JsonValueKind.Array)
                                    {
                                        foreach (var item in prop.Value.EnumerateArray())
                                        {
                                            // 提取工作表名
                                            string sName = item.GetString()?.Trim() ?? string.Empty;
                                            if (!string.IsNullOrWhiteSpace(sName) && !sourceSheets.Contains(sName))
                                            {
                                                sourceSheets.Add(sName);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // 若非有效 JSON，采用常用分隔符切分表名 (逗号、分号、换行)
                            var parts = rawConfig.Split(new[] { ',', ';', '，', '；', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var p in parts)
                            {
                                // 清洗前后多余引号或中括号
                                string sName = p.Trim().Trim('[', ']', '"', '\'', ' ');
                                if (!string.IsNullOrWhiteSpace(sName) && !sourceSheets.Contains(sName))
                                {
                                    sourceSheets.Add(sName);
                                }
                            }
                        }
                    }
                }
                catch { }

                // 组装查询模型并返回
                return new ComponentDistributionQuery
                {
                    ComponentName = compName,
                    Model = !string.IsNullOrWhiteSpace(newModel) ? newModel : origModel,
                    OriginalModel = origModel,
                    Manufacturer = mfg,
                    Unit = unit,
                    SummaryQuantity = sumQty,
                    RawModelFromU = rawModelU,
                    SummaryRowIndex = summaryRow,
                    SourceSheets = sourceSheets
                };
            }
            catch (Exception ex)
            {
                // 记录查询入参解析异常日志
                LogHelper.WriteLog($"[查看分布] ExtractComponentQueryFromSummarySheet 异常: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 扫描全工作簿所有分类表，检索指定元器件在各分类明细表及箱柜中的分布
        /// 严格遵循规则 6、7、8、9 进行结构校验与内存二维数组扫描
        /// </summary>
        public static ComponentDistributionResultDto ScanComponentDistribution(ComponentDistributionQuery query)
        {
            // 初始化秒表记录执行耗时
            var stopwatch = Stopwatch.StartNew();

            // 初始化返回结果对象
            var result = new ComponentDistributionResultDto
            {
                ComponentInfo = query
            };

            // 基础入参有效性校验
            if (query == null || (string.IsNullOrWhiteSpace(query.ComponentName) && string.IsNullOrWhiteSpace(query.Model)))
            {
                result.Success = false;
                result.Message = "元器件查询条件不完整，缺少名称或型号";
                return result;
            }

            try
            {
                // 获取 Excel 应用程序上下文
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    result.Success = false;
                    result.Message = "无法连接至 Excel 应用程序上下文";
                    return result;
                }

                // 获取当前活动工作簿
                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null)
                {
                    result.Success = false;
                    result.Message = "当前未打开任何有效 Excel 工作簿";
                    return result;
                }

                // 收集全簿有效明细项列表
                var allMatchedRows = new List<ComponentDistributionRowDto>();
                int globalIndex = 1;

                // 目标匹配元器件特征清洗
                string targetName = query.ComponentName.Trim();
                string targetModel = query.Model.Trim();
                string targetOrigModel = query.OriginalModel.Trim();
                string targetRawU = query.RawModelFromU.Trim();

                // 用户明确要求：请求的时候需要携带源分类表名称，只扫描源分类表
                if (query.SourceSheets == null || query.SourceSheets.Count == 0)
                {
                    // 规则 9: 不符合要求直接报错提醒，不进行模糊容错
                    result.Success = false;
                    result.Message = "元件汇总表 AB1 单元格未配置源分类表信息，无法确定需扫描的分类表";
                    return result;
                }

                // 遍历指定的源分类表清单，严格只扫描源分类表
                foreach (string rawSheetName in query.SourceSheets)
                {
                    // 清洗工作表名称
                    string sheetName = rawSheetName?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(sheetName)) continue;

                    // 获取目标分类工作表
                    dynamic sheet;
                    try
                    {
                        sheet = activeWb.Worksheets[sheetName];
                    }
                    catch (Exception ex)
                    {
                        // 规则 9: 源分类表丢失直接报错提醒
                        result.Success = false;
                        result.Message = $"源分类表【{sheetName}】不存在于当前工作簿: {ex.Message}";
                        return result;
                    }

                    if (sheet == null)
                    {
                        // 无法获取工作表对象报错
                        result.Success = false;
                        result.Message = $"无法获取源分类表【{sheetName}】对象";
                        return result;
                    }

                    // 规则 8: 检查当前表格规则 6 的正确性 (通过 GetSheetValidCabinets 检索定义名称)
                    List<KeyValuePair<int, CabinetAnchorModel>> validCabinets;
                    try
                    {
                        validCabinets = Tool.GetSheetValidCabinets((object)sheet, (object)activeWb);
                    }
                    catch (Exception ex)
                    {
                        // 规则 9: 不符合要求直接报错提醒，不进行模糊容错
                        result.Success = false;
                        result.Message = $"工作表【{sheetName}】箱柜定义名称损坏或不符合规则6: {ex.Message}";
                        return result;
                    }

                    // 若无有效箱柜定义，判定为普通非分类表，跳过
                    if (validCabinets == null || validCabinets.Count == 0)
                    {
                        continue;
                    }

                    // 获取工作表最大已用行数，避免过度分配
                    int usedRows = 0;
                    try { usedRows = sheet.UsedRange.Rows.Count; } catch { }
                    if (usedRows < 5) continue;

                    // 扩展安全冗余读取边界
                    int maxReadRow = Math.Max(usedRows + 20, 100);

                    // 规则 7: 一次性将该分类表整表数据读入内存二维数组 (覆盖 A~Z 列共 26 列)
                    dynamic sheetRange = sheet.Range[$"A1:Z{maxReadRow}"];
                    object[,]? sheetMatrix = sheetRange?.Value2 as object[,];
                    if (sheetMatrix == null) continue;

                    int maxArrRow = sheetMatrix.GetLength(0);
                    int maxArrCol = sheetMatrix.GetLength(1);

                    // 本地高速安全取值委托 (0 次 COM 调用)
                    object? GetVal(int r, int c)
                    {
                        if (r >= 1 && r <= maxArrRow && c >= 1 && c <= maxArrCol) return sheetMatrix[r, c];
                        return null;
                    }

                    // 遍历该分类表下的每一个有效箱柜
                    foreach (var cabPair in validCabinets)
                    {
                        var cab = cabPair.Value;
                        if (cab == null || cab.Det == null) continue;

                        // 提取箱柜锚点行号 (纯内存整数)
                        int detRow = cab.DetRow > 0 ? cab.DetRow : Convert.ToInt32(cab.Det.Row);
                        int subsumRow = cab.SubsumRow > 0 ? cab.SubsumRow : (cab.Subsum != null ? Convert.ToInt32(cab.Subsum.Row) : 0);
                        int sumRow = cab.SumRow > 0 ? cab.SumRow : (cab.Sum != null ? Convert.ToInt32(cab.Sum.Row) : 0);

                        // 规则 6 校验: 必须存在 Cab_Subsum 行
                        if (subsumRow <= detRow)
                        {
                            // 规则 9: 结构不符合要求直接报错提醒
                            result.Success = false;
                            result.Message = $"工作表【{sheetName}】箱柜(Det行:{detRow})缺少有效的 Cab_Subsum 定义名称，违反规则6";
                            return result;
                        }

                        // 从汇总行(sumRow)读取箱柜柜号、设备名称与箱柜台数
                        string cabNo = string.Empty;
                        string cabName = string.Empty;
                        int cabQty = 1;

                        if (sumRow > 0)
                        {
                            // 汇总行 B 列柜号 (第 2 列)
                            cabNo = Convert.ToString(GetVal(sumRow, 2))?.Trim() ?? string.Empty;
                            // 汇总行 C 列设备名称 (第 3 列)
                            cabName = Convert.ToString(GetVal(sumRow, 3))?.Trim() ?? string.Empty;
                            // 汇总行 F 列箱柜数量/台数 (第 6 列)
                            object? qObj = GetVal(sumRow, 6) ?? GetVal(sumRow, 5);
                            if (qObj != null && int.TryParse(Convert.ToString(qObj), out int parsedCabCount) && parsedCabCount > 0)
                            {
                                cabQty = parsedCabCount;
                            }
                        }

                        // 兜底: 若汇总行未读出柜号，从明细上一行尝试补齐
                        if (string.IsNullOrWhiteSpace(cabNo) && detRow > 1)
                        {
                            cabNo = Convert.ToString(GetVal(detRow - 1, 2))?.Trim() ?? string.Empty;
                        }

                        // 规则 6 明确规定: 元器件起始行 = Cab_Det.row + 2，终止行 = Cab_Subsum.row - 1
                        int compStartRow = detRow + 2;
                        int compEndRow = subsumRow - 1;

                        // 若该箱柜内无元器件行则跳过
                        if (compEndRow < compStartRow) continue;

                        // 循环扫描元器件区间内的所有行
                        for (int r = compStartRow; r <= compEndRow; r++)
                        {
                            // B 列元件名称 (索引 2)
                            string rowCompName = Convert.ToString(GetVal(r, 2))?.Trim() ?? string.Empty;
                            // C 列型号规格 (索引 3)
                            string rowModel = Convert.ToString(GetVal(r, 3))?.Trim() ?? string.Empty;

                            // 排除名称与型号均为空的空白行
                            if (string.IsNullOrWhiteSpace(rowCompName) && string.IsNullOrWhiteSpace(rowModel))
                            {
                                continue;
                            }

                            // D 列生产厂家 (索引 4)
                            string rowMfg = Convert.ToString(GetVal(r, 4))?.Trim() ?? string.Empty;

                            // F 列单柜元器件数量 (索引 6)
                            decimal rowQty = 0;
                            object? rQ = GetVal(r, 6);
                            if (rQ != null && decimal.TryParse(Convert.ToString(rQ), out decimal q))
                            {
                                rowQty = q;
                            }

                            // I 列回路编号或备注 (索引 9)
                            string rowRemark = Convert.ToString(GetVal(r, 9))?.Trim() ?? string.Empty;

                            // U 列原始型号 (索引 21)
                            string rowRawU = Convert.ToString(GetVal(r, 21))?.Trim() ?? string.Empty;

                            // 判定当前行是否为目标元器件
                            bool isMatched = false;

                            // 1. 严格名称匹配
                            if (string.Equals(rowCompName, targetName, StringComparison.OrdinalIgnoreCase))
                            {
                                // 2. 型号匹配 (与新型号、原型号或原始型号之一吻合)
                                if (string.Equals(rowModel, targetModel, StringComparison.OrdinalIgnoreCase) ||
                                    (!string.IsNullOrEmpty(targetOrigModel) && string.Equals(rowModel, targetOrigModel, StringComparison.OrdinalIgnoreCase)) ||
                                    (!string.IsNullOrEmpty(targetRawU) && string.Equals(rowRawU, targetRawU, StringComparison.OrdinalIgnoreCase)))
                                {
                                    isMatched = true;
                                }
                            }

                            // 命中分布记录
                            if (isMatched)
                            {
                                allMatchedRows.Add(new ComponentDistributionRowDto
                                {
                                    Index = globalIndex++,
                                    SheetName = sheetName,
                                    CabinetNo = cabNo,
                                    CabinetName = cabName,
                                    QuantityPerCabinet = rowQty,
                                    CabinetCount = cabQty,
                                    Remark = rowRemark,
                                    RowIndex = r,
                                    ColumnIndex = 2, // 默认定位在 B 列名称 --硬编码--
                                    DetRow = detRow // 记录所属箱柜 Cab_Det 行号供视口顶行定位
                                });
                            }
                        }
                    }
                }

                // 统计扁平明细与多维分组
                result.FlatRows = allMatchedRows;
                result.TotalQuantityPerCabinet = allMatchedRows.Sum(x => x.QuantityPerCabinet);

                // 按分类工作表进行分组组织
                var grouped = allMatchedRows
                    .GroupBy(x => x.SheetName)
                    .Select(g => new ComponentDistributionCategoryGroupDto
                    {
                        SheetName = g.Key,
                        CabinetCount = g.Select(r => r.CabinetNo).Distinct().Count(),
                        SubtotalQuantityPerCabinet = g.Sum(r => r.QuantityPerCabinet),
                        Rows = g.ToList()
                    })
                    .ToList();

                result.Groups = grouped;
                result.TotalSheetCount = grouped.Count;
                result.TotalCabinetCount = allMatchedRows.Select(x => $"{x.SheetName}#{x.CabinetNo}").Distinct().Count();
                result.Success = true;
                result.Message = allMatchedRows.Count > 0 ? "扫描完成" : "未在任何分类表中找到该元器件的明细分布记录";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"扫描元器件分布时发生未预期异常: {ex.Message}";
                LogHelper.WriteLog($"[查看分布] ScanComponentDistribution 异常: {ex.Message}");
            }
            finally
            {
                stopwatch.Stop();
                result.ElapsedMilliseconds = stopwatch.ElapsedMilliseconds;
            }

            return result;
        }

        /// <summary>
        /// 在 Excel 中穿透跳转定位并高亮目标单元格，将系统前台焦点无缝归还给 Excel
        /// 用户明确要求：定位的时候需要把 det 行放在第一行
        /// 遵循规则 3: Excel 操作收拢于 ExcelServices
        /// </summary>
        public static bool LocateComponentInSheet(string sheetName, int row, int col = 2, int detRow = 0)
        {
            try
            {
                // 获取 Excel Application 上下文
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return false;

                // 获取活动工作簿
                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null) return false;

                // 获取指定工作表
                dynamic targetSheet = activeWb.Worksheets[sheetName];
                if (targetSheet == null) return false;

                // 激活工作表
                targetSheet.Activate();

                // 提取目标行对应单元格
                dynamic targetCell = targetSheet.Cells[row, col];

                // 用户明确要求：定位的时候需要把 det 行放在第一行
                try
                {
                    // 若 detRow 传入无效或为 0，动态匹配当前元件行归属箱柜的 Cab_Det 行
                    if (detRow <= 0)
                    {
                        var validCabinets = Tool.GetSheetValidCabinets((object)targetSheet, (object)activeWb);
                        if (validCabinets != null)
                        {
                            foreach (var cabPair in validCabinets)
                            {
                                var cab = cabPair.Value;
                                // 提取 det 行与 subsum 行
                                int cDetRow = cab.DetRow > 0 ? cab.DetRow : (cab.Det != null ? Convert.ToInt32(cab.Det.Row) : 0);
                                int cSubsumRow = cab.SubsumRow > 0 ? cab.SubsumRow : (cab.Subsum != null ? Convert.ToInt32(cab.Subsum.Row) : 0);
                                // 命中当前元件所在行区间
                                if (cDetRow > 0 && row >= cDetRow && (cSubsumRow == 0 || row <= cSubsumRow))
                                {
                                    detRow = cDetRow;
                                    break;
                                }
                            }
                        }
                    }

                    if (detRow > 0)
                    {
                        // 用户明确要求：定位的时候需要把 det 行放在第一行 (视口垂直滚动置顶)
                        app.ActiveWindow.ScrollRow = Math.Max(1, detRow);
                        // 水平视口滚动回第一列 A 列
                        app.ActiveWindow.ScrollColumn = 1;
                    }
                    else
                    {
                        // 若未匹配到有效 detRow，使用 Goto 定位居中展示目标单元格
                        app.Goto(targetCell, true);
                    }
                }
                catch { }

                // 视口滚动就绪后，显式选中目标单元格，确保光标与选区准确落在目标单元格上
                try
                {
                    targetCell.Select();
                }
                catch { }

                // 若 Excel 处于最小化状态，恢复为正常窗口
                try
                {
                    if (app.WindowState == (int)Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized)
                    {
                        app.WindowState = (int)Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
                    }
                }
                catch { }

                // 将 Windows 系统输入焦点稳稳切回 Excel 主窗口 (EXCEL7)
                try
                {
                    IntPtr excelHwnd = new IntPtr((int)app.Hwnd);
                    Win32Interop.SetForegroundWindow(excelHwnd);
                }
                catch { }

                // 核心修复: 显式触发聚光灯 (Spotlight) 重新计算视口坐标并聚焦当前目标单元格
                try
                {
                    if (ExcelServices.IsSpotlightEnabled)
                    {
                        // 1. 同步立即更新聚光灯位置
                        ExcelServices.UpdateSpotlightPosition(targetCell);

                        // 2. 异步 50ms 再次微调补正，确保视口滚动重绘彻底完成后严丝合缝对齐
                        System.Threading.Tasks.Task.Delay(50).ContinueWith(_ =>
                        {
                            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                            {
                                ExcelServices.UpdateSpotlightPosition(targetCell);
                            });
                        });
                    }
                }
                catch { }

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[查看分布] LocateComponentInSheet 异常: {ex.Message}");
                return false;
            }
        }


        /// <summary>
        /// 弹出基于 WebView2 + Vue 3 的“元器件分布穿透”看板窗口 (非模态)
        /// </summary>
        public static void ShowComponentDistributionDialog(int summaryRow = 0)
        {
            try
            {
                // 若未指定行号，尝试从当前活动单元格中提取行号
                if (summaryRow <= 0)
                {
                    dynamic? dynApp = ExcelDnaSafeAccessor.GetApplication();
                    if (dynApp != null && dynApp.ActiveCell != null)
                    {
                        summaryRow = Convert.ToInt32(dynApp.ActiveCell.Row);
                    }
                }

                // 提取右键所在行的元器件信息
                var query = ExtractComponentQueryFromSummarySheet(summaryRow);
                if (query == null)
                {
                    MessageBox.Show(
                        "请在“元件汇总表”中右键点击有效的元器件数据行（B列名称或型号不能为空，且不能为表头或合计行）。",
                        "提示",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                // 以非模态方式创建或激活窗体
                ShowModelessForm(ref _componentDistributionForm, () => new ComponentDistributionForm(query));

                // 若窗口已经处于打开态，刷新其查询条件并重新加载
                if (_componentDistributionForm != null && !_componentDistributionForm.IsDisposed)
                {
                    _componentDistributionForm.UpdateQueryAndRefresh(query);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"弹出元器件分布看板窗口失败: {ex.Message}", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogHelper.WriteLog($"[查看分布] ShowComponentDistributionDialog 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 方案 A（智能差量计算，数量多柜时自动分流）：从分类明细表反向同步更新至元件汇总表 C 列（原型号规格）
        /// 遵循规则 3: 所有 Excel 操作全部封装于 ExcelServices，不写在窗口类中
        /// 遵循规则 7: 内存二维数组批读批写
        /// 遵循规则 8、9: 强校验定义名称与数据格式，不符合要求直接报错
        /// </summary>
        public static UpdateSummaryResultDto UpdateSummaryComponentFromDetailRow(string sheetName, int detailRow, int summaryRow)
        {
            // 初始化返回结果 DTO 模型
            var result = new UpdateSummaryResultDto();
            try
            {
                // 获取 Excel 顶层 Application 实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    // 无法连接 Excel 报错
                    result.Success = false;
                    result.Message = "无法连接至 Excel 应用程序上下文";
                    return result;
                }

                // 获取活动工作簿对象
                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null)
                {
                    // 未打开工作簿报错
                    result.Success = false;
                    result.Message = "当前未打开任何有效工作簿";
                    return result;
                }

                // 1. 获取源分类明细工作表
                dynamic detailSheet;
                try
                {
                    // 根据表名检索源分类工作表
                    detailSheet = activeWb.Worksheets[sheetName];
                }
                catch (Exception ex)
                {
                    // 规则 9: 源分类表丢失直接报错提醒
                    result.Success = false;
                    result.Message = $"源分类表【{sheetName}】不存在: {ex.Message}";
                    return result;
                }

                // 2. 获取元件汇总表工作表
                dynamic? summarySheet = null;
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    // 获取工作表名称
                    string wsName = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                    // 比对工作表名是否为“元件汇总表”或包含“元件汇总”
                    if (string.Equals(wsName, ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase) ||
                        wsName.Contains("元件汇总"))
                    {
                        summarySheet = ws;
                        break;
                    }
                }

                if (summarySheet == null)
                {
                    // 规则 9: 汇总表不存在直接报错
                    result.Success = false;
                    result.Message = "当前工作簿中未找到【元件汇总表】工作表，无法执行更新";
                    return result;
                }

                // 3. 校验分类表规则 6 正确性并找到归属箱柜 (规则 8)
                var validCabinets = Tool.GetSheetValidCabinets((object)detailSheet, (object)activeWb);
                if (validCabinets == null || validCabinets.Count == 0)
                {
                    // 规则 9: 箱柜定义损坏直接报错
                    result.Success = false;
                    result.Message = $"工作表【{sheetName}】箱柜定义名称不完整或不存在，违反规则6";
                    return result;
                }

                // 查找归属箱柜
                CabinetAnchorModel? targetCab = null;
                foreach (var cabPair in validCabinets)
                {
                    var cab = cabPair.Value;
                    // 读取 det 行与 subsum 行号
                    int detR = cab.DetRow > 0 ? cab.DetRow : Convert.ToInt32(cab.Det.Row);
                    int subsumR = cab.SubsumRow > 0 ? cab.SubsumRow : Convert.ToInt32(cab.Subsum.Row);
                    // 规则 6: 元器件起始行 = Cab_Det.row + 2，终止行 = Cab_Subsum.row - 1
                    if (detailRow >= detR + 2 && detailRow <= subsumR - 1)
                    {
                        targetCab = cab;
                        break;
                    }
                }

                if (targetCab == null)
                {
                    // 未命中有效箱柜元器件区间报错
                    result.Success = false;
                    result.Message = $"行 {detailRow} 不在工作表【{sheetName}】的任何箱柜元器件有效区间内，无法更新";
                    return result;
                }

                // 4. 规则 7: 一次性读取分类表明细行与归属箱柜汇总行数据
                int sumR = targetCab.SumRow > 0 ? targetCab.SumRow : Convert.ToInt32(targetCab.Sum.Row);
                // 读取明细行 A~I 列 (第 1 到 9 列)
                dynamic detailRange = detailSheet.Range[$"A{detailRow}:I{detailRow}"];
                object[,]? detailMatrix = detailRange?.Value2 as object[,];
                if (detailMatrix == null)
                {
                    // 读取失败报错
                    result.Success = false;
                    result.Message = "读取分类表当前明细行数据失败";
                    return result;
                }

                // B 列元件名称 (索引 2)
                string detailCompName = Convert.ToString(detailMatrix[1, 2])?.Trim() ?? string.Empty;
                // C 列型号规格 (用户在分类表中修改的新型号)
                string detailNewModel = Convert.ToString(detailMatrix[1, 3])?.Trim() ?? string.Empty;
                // D 列生产厂家 (索引 4)
                string detailMfg = Convert.ToString(detailMatrix[1, 4])?.Trim() ?? string.Empty;
                // E 列单位 (索引 5)
                string detailUnit = Convert.ToString(detailMatrix[1, 5])?.Trim() ?? string.Empty;
                // F 列单柜数量 (索引 6)
                decimal detailSingleQty = 0;
                object? qVal = detailMatrix[1, 6];
                if (qVal != null && decimal.TryParse(Convert.ToString(qVal), out decimal parsedSingleQty))
                {
                    detailSingleQty = parsedSingleQty;
                }

                if (string.IsNullOrWhiteSpace(detailNewModel))
                {
                    // 型号为空拦截
                    result.Success = false;
                    result.Message = "分类明细表该行的型号规格为空，无法同步更新";
                    return result;
                }

                // 读取箱柜台数 cabQty (从 sumRow 的 F 列或 E 列读取)
                int cabQty = 1;
                if (sumR > 0)
                {
                    // 规则 7: 读取汇总行台数
                    dynamic sumRowRange = detailSheet.Range[$"E{sumR}:F{sumR}"];
                    object[,]? sumMatrix = sumRowRange?.Value2 as object[,];
                    if (sumMatrix != null)
                    {
                        // 优先 F 列数量，后备 E 列
                        object? cabCountVal = sumMatrix[1, 2] ?? sumMatrix[1, 1];
                        if (cabCountVal != null && int.TryParse(Convert.ToString(cabCountVal), out int parsedCabCount) && parsedCabCount > 0)
                        {
                            cabQty = parsedCabCount;
                        }
                    }
                }

                // 计算该箱柜本行元器件在工程中的总差量用量
                decimal deltaQty = detailSingleQty * cabQty;
                if (deltaQty <= 0)
                {
                    // 用量小于等于 0 提醒
                    result.Success = false;
                    result.Message = $"该元器件在此箱柜中的总用量为 {deltaQty}，无需更新";
                    return result;
                }

                // 5. 规则 7: 一次性读取元件汇总表现有行 (summaryRow) 的数据 (A~R 列)
                if (summaryRow < 6)
                {
                    // 数据行从第 6 行开始
                    result.Success = false;
                    result.Message = $"目标汇总表行号 {summaryRow} 无效 (数据从第 6 行开始)";
                    return result;
                }

                // 一次性读取汇总行 A~R 列
                dynamic sumCompRange = summarySheet.Range[$"A{summaryRow}:R{summaryRow}"];
                object[,]? sumCompMatrix = sumCompRange?.Value2 as object[,];
                if (sumCompMatrix == null)
                {
                    // 无法读取汇总行数据报错
                    result.Success = false;
                    result.Message = "读取元件汇总表目标行数据失败";
                    return result;
                }

                // B 列元件名称
                string sumCompName = Convert.ToString(sumCompMatrix[1, 2])?.Trim() ?? string.Empty;
                // C 列原型号规格 (用户指定更新到这一列)
                string sumOrigModel = Convert.ToString(sumCompMatrix[1, 3])?.Trim() ?? string.Empty;
                // F 列汇总数量
                decimal currSumQty = 0;
                object? currQtyObj = sumCompMatrix[1, 6];
                if (currQtyObj != null && decimal.TryParse(Convert.ToString(currQtyObj), out decimal parsedSumQty))
                {
                    currSumQty = parsedSumQty;
                }

                // 校验：若分类明细表中的型号与汇总表现有原型号完全一致，提示无需更新
                if (string.Equals(detailNewModel, sumOrigModel, StringComparison.OrdinalIgnoreCase))
                {
                    result.Success = false;
                    result.Message = $"分类明细表的型号【{detailNewModel}】与元件汇总表现有原型号一致，未做修改，无需更新";
                    return result;
                }

                // 6. 执行方案 A 智能差量分流算法
                // 颜色定义: 浅蓝色背景色 (软蓝底色 220, 235, 252) --硬编码: 基准列底色--
                int softBlueColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(220, 235, 252));

                // 分支 1: deltaQty >= currSumQty (单柜独占或用量占满汇总行) -> 直接就地改写
                if (deltaQty >= currSumQty)
                {
                    // 用户指令：更新到 C 列（原型号规格）
                    summarySheet.Cells[summaryRow, 3].Value = detailNewModel;
                    // D 列调价型号保持清空
                    summarySheet.Cells[summaryRow, 4].Value = string.Empty;

                    // 若厂家有填写且不同，同步更新 I 列
                    if (!string.IsNullOrWhiteSpace(detailMfg))
                    {
                        summarySheet.Cells[summaryRow, 9].Value = detailMfg;
                    }
                    // 若单位有填写且不同，同步更新 E 列
                    if (!string.IsNullOrWhiteSpace(detailUnit))
                    {
                        summarySheet.Cells[summaryRow, 5].Value = detailUnit;
                    }

                    // 数量与用量校准为最新 deltaQty
                    summarySheet.Cells[summaryRow, 6].Value = (double)deltaQty;

                    // 填充成功返回值
                    result.Success = true;
                    result.IsSplit = false;
                    result.TargetSummaryRow = summaryRow;
                    result.NewModel = detailNewModel;
                    result.DeltaQuantity = deltaQty;
                    result.Message = $"已将元件汇总表第 {summaryRow} 行原型号规格更新为【{detailNewModel}】（用量: {deltaQty}）";
                    return result;
                }

                // 分支 2: deltaQty < currSumQty (多柜共用，智能差量分流)
                // 步骤 2.1: 从原汇总行扣减用量
                decimal remainingQty = currSumQty - deltaQty;
                summarySheet.Cells[summaryRow, 6].Value = (double)remainingQty;

                // 步骤 2.2: 扫描元件汇总表所有数据行，寻找是否已存在相同名称且 C 列为 detailNewModel 的行
                int totalRow = FindSummaryTotalRow(summarySheet);
                int endDataRow = totalRow - 1;
                int existingTargetRow = -1;

                if (endDataRow >= 6)
                {
                    // 规则 7: 一次性读取整个汇总表数据区 B~C 列至二维数组进行比对
                    dynamic allDataRange = summarySheet.Range[$"B6:C{endDataRow}"];
                    object[,]? allDataMatrix = allDataRange?.Value2 as object[,];
                    if (allDataMatrix != null)
                    {
                        int rowCount = allDataMatrix.GetLength(0);
                        // 循环比对已有元器件
                        for (int i = 1; i <= rowCount; i++)
                        {
                            string rName = Convert.ToString(allDataMatrix[i, 1])?.Trim() ?? string.Empty;
                            string rModel = Convert.ToString(allDataMatrix[i, 2])?.Trim() ?? string.Empty;
                            // 匹配相同元件名称与相同新型号
                            if (string.Equals(rName, detailCompName, StringComparison.OrdinalIgnoreCase) &&
                                string.Equals(rModel, detailNewModel, StringComparison.OrdinalIgnoreCase))
                            {
                                existingTargetRow = 6 + i - 1;
                                break;
                            }
                        }
                    }
                }

                // 情况 2.2.A: 汇总表中已存在同名同新型号行，直接累加数量
                if (existingTargetRow > 0)
                {
                    // 读取目标行现有数量
                    decimal targetCurQty = 0;
                    object? tQtyObj = summarySheet.Cells[existingTargetRow, 6].Value;
                    if (tQtyObj != null && decimal.TryParse(Convert.ToString(tQtyObj), out decimal pTQty))
                    {
                        targetCurQty = pTQty;
                    }

                    // 累加差量用量
                    decimal newTotalQty = targetCurQty + deltaQty;
                    summarySheet.Cells[existingTargetRow, 6].Value = (double)newTotalQty;

                    result.Success = true;
                    result.IsSplit = true;
                    result.TargetSummaryRow = existingTargetRow;
                    result.NewModel = detailNewModel;
                    result.DeltaQuantity = deltaQty;
                    result.Message = $"智能分流：原汇总行扣减 {deltaQty}（剩余 {remainingQty}）；新型号已累加至第 {existingTargetRow} 行（最新数量: {newTotalQty}）";
                    return result;
                }

                // 情况 2.2.B: 汇总表中无此型号，在当前 summaryRow 下方插入新行并更新到 C 列
                int insertRowIndex = summaryRow + 1;
                // 插入物理行并向下推移 (xlShiftDown = -4121) --硬编码: COM 常量--
                summarySheet.Rows[insertRowIndex].Insert(-4121);

                // 彻底杜绝整行 16384 列 Copy/PasteSpecial 造成的 0x8007000E 内存溢出！
                // 仅针对有效数据列 A~Y 列 (25列) 进行定向格式复制，不走全表剪贴板
                try
                {
                    // 仅复制有效数据区域 A~Y 列格式 (规则 4 最小变动法则)
                    dynamic srcRange = summarySheet.Range[$"A{summaryRow}:Y{summaryRow}"];
                    dynamic destRange = summarySheet.Range[$"A{insertRowIndex}:Y{insertRowIndex}"];
                    srcRange.Copy();
                    destRange.PasteSpecial(-4122); // xlPasteFormats = -4122 --硬编码--
                    app.CutCopyMode = 0; // 立即清空剪贴板跑马灯
                }
                catch (Exception cpEx)
                {
                    // 即使剪贴板因系统外部程序冲突受限，记录日志并不阻断后续核心数据写入
                    LogHelper.WriteLog($"[查看分布] 复制行格式容错捕获: {cpEx.Message}");
                }

                // 填充新行纯数据矩阵 (B~R 列共 17 列一次性构建写入，规则 7)
                // 重点: 纯文本与型号必须使用 Value2 写入，杜绝 put_Formula 误将包含 - + / 的型号当数学公式编译导致 0x8007000E 内存溢出
                object[,] newRowData = new object[1, 17];
                newRowData[0, 0] = detailCompName; // B: 元件名称
                newRowData[0, 1] = detailNewModel; // C: 原型号规格 (用户指定更新到 C 列！)
                newRowData[0, 2] = string.Empty; // D: 型号规格 (调价列留空)
                newRowData[0, 3] = !string.IsNullOrWhiteSpace(detailUnit) ? detailUnit : sumCompMatrix[1, 5]; // E: 单位
                newRowData[0, 4] = (double)deltaQty; // F: 数量
                newRowData[0, 5] = 0; // G: 单价 (占位数值，稍后由单独 Formula 赋公式)
                newRowData[0, 6] = 0; // H: 总价 (占位数值，稍后由单独 Formula 赋公式)
                newRowData[0, 7] = !string.IsNullOrWhiteSpace(detailMfg) ? detailMfg : sumCompMatrix[1, 9]; // I: 生产厂家
                newRowData[0, 8] = 0; // J: 成本单价 (占位数值，稍后由单独 Formula 赋公式)
                newRowData[0, 9] = sumCompMatrix[1, 11] ?? 1.0; // K: 报出系数
                newRowData[0, 10] = sumCompMatrix[1, 12] ?? string.Empty; // L: 本体表价
                newRowData[0, 11] = sumCompMatrix[1, 13] ?? 1.0; // M: 本体折扣
                newRowData[0, 12] = sumCompMatrix[1, 14] ?? string.Empty; // N: 附件表价
                newRowData[0, 13] = sumCompMatrix[1, 15] ?? string.Empty; // O: 附件折扣
                newRowData[0, 14] = sumCompMatrix[1, 16] ?? string.Empty; // P: 备注
                newRowData[0, 15] = sumCompMatrix[1, 17] ?? string.Empty; // Q: 类别
                newRowData[0, 16] = detailNewModel; // R: 原始型号

                // 规则 7: 一次性将新行数据通过 Value2 写入，确保带加减号型号当纯文本存储
                summarySheet.Range[$"B{insertRowIndex}:R{insertRowIndex}"].Value2 = newRowData;

                // 单独为真正需要动态计算的 3 列单元格赋标准 Excel 公式 (G 列单价、H 列总价、J 列成本单价)
                try
                {
                    // G 列单价公式
                    summarySheet.Cells[insertRowIndex, 7].Formula = $"=ROUND(K{insertRowIndex}*(L{insertRowIndex}*M{insertRowIndex}+N{insertRowIndex}*O{insertRowIndex}),2)";
                    // H 列总价公式
                    summarySheet.Cells[insertRowIndex, 8].Formula = $"=ROUND(F{insertRowIndex}*G{insertRowIndex},2)";
                    // J 列成本单价公式
                    summarySheet.Cells[insertRowIndex, 10].Formula = $"=ROUND(L{insertRowIndex}*M{insertRowIndex}+N{insertRowIndex}*O{insertRowIndex}),2)";
                }
                catch (Exception formulaEx)
                {
                    // 记录公式写入容错日志
                    LogHelper.WriteLog($"[查看分布] 新行公式设置容错捕获: {formulaEx.Message}");
                }


                // 恢复浅蓝底色与标准行高 --硬编码--
                try
                {
                    summarySheet.Range[$"A{insertRowIndex}"].Interior.Color = softBlueColor;
                    summarySheet.Range[$"C{insertRowIndex}"].Interior.Color = softBlueColor;
                    summarySheet.Range[$"F{insertRowIndex}"].Interior.Color = softBlueColor;
                    summarySheet.Range[$"J{insertRowIndex}"].Interior.Color = softBlueColor;
                    summarySheet.Range[$"{insertRowIndex}:{insertRowIndex}"].RowHeight = 20;
                }
                catch { }

                // 辅助刷新整表序号与底部合计行公式 (采用独立 try-catch 防御，避免次生异常阻断主流程成功)
                try
                {
                    // 重新定位合计行
                    int newTotalRow = FindSummaryTotalRow(summarySheet);
                    int newDataEndRow = newTotalRow - 1;
                    int dataRowCount = newDataEndRow - 6 + 1;

                    if (dataRowCount > 0 && newDataEndRow >= 6)
                    {
                        // 规则 7: 内存构建连续序号数组并一次性刷入 A 列
                        object[,] seqMatrix = new object[dataRowCount, 1];
                        for (int s = 0; s < dataRowCount; s++)
                        {
                            seqMatrix[s, 0] = s + 1;
                        }
                        // 一次性写入序号
                        summarySheet.Range[$"A6:A{newDataEndRow}"].Value2 = seqMatrix;

                        // 刷新合计行公式
                        summarySheet.Cells[newTotalRow, 6].Formula = $"=SUM(F6:F{newDataEndRow})";
                        summarySheet.Cells[newTotalRow, 8].Formula = $"=SUM(H6:H{newDataEndRow})";
                    }
                }
                catch (Exception postEx)
                {
                    LogHelper.WriteLog($"[查看分布] 序号或合计公式重排容错捕获: {postEx.Message}");
                }

                // 组装成功结果并返回
                result.Success = true;
                result.IsSplit = true;
                result.TargetSummaryRow = insertRowIndex;
                result.NewModel = detailNewModel;
                result.DeltaQuantity = deltaQty;
                result.Message = $"智能分流：原汇总行扣减 {deltaQty}（剩余 {remainingQty}）；已在第 {insertRowIndex} 行新增型号记录【{detailNewModel}】（用量: {deltaQty}）";
                return result;
            }
            catch (Exception ex)
            {
                // 异常日志完整记录 (包含堆栈)
                result.Success = false;
                result.Message = $"更新元件汇总表发生异常: {ex.Message}";
                LogHelper.WriteLog($"[查看分布] UpdateSummaryComponentFromDetailRow 异常详情: {ex}");
                return result;
            }
        }

        /// <summary>
        /// 规则 7: 一次性读取内存二维数组，高速定位元件汇总表的“合计”或“总计”行物理行号
        /// 杜绝逐单元格 COM 循环调用引发的 COM 句柄泄漏与 0x8007000E 内存溢出
        /// </summary>
        private static int FindSummaryTotalRow(dynamic summarySheet)
        {
            try
            {
                // 获取工作表已用最大行号
                int usedRowCount = 100;
                try
                {
                    usedRowCount = summarySheet.UsedRange.Rows.Count;
                }
                catch { }

                // 设定安全扫描上限
                int scanEndRow = Math.Max(usedRowCount + 10, 50);
                if (scanEndRow > 3000) scanEndRow = 3000; // --硬编码: 安全扫描上限--

                if (scanEndRow >= 6)
                {
                    // 规则 7: 一次性将 A6:C{scanEndRow} 读入内存二维数组 (仅 1 次 COM 调用)
                    dynamic scanRange = summarySheet.Range[$"A6:C{scanEndRow}"];
                    object[,]? scanMatrix = scanRange?.Value2 as object[,];
                    if (scanMatrix != null)
                    {
                        int rowCount = scanMatrix.GetLength(0);
                        int lastDataRowOffset = 0;

                        // 循环内存二维数组寻找合计
                        for (int i = 1; i <= rowCount; i++)
                        {
                            string aVal = Convert.ToString(scanMatrix[i, 1])?.Trim() ?? string.Empty;
                            string bVal = Convert.ToString(scanMatrix[i, 2])?.Trim() ?? string.Empty;
                            string cVal = Convert.ToString(scanMatrix[i, 3])?.Trim() ?? string.Empty;

                            // 优先检测 A 列或 B 列包含“合计”或“总计”
                            if (aVal.Contains("合计") || aVal.Contains("总计") ||
                                bVal.Contains("合计") || bVal.Contains("总计"))
                            {
                                return 6 + i - 1;
                            }

                            // 记录最后有效数据行
                            if (!string.IsNullOrWhiteSpace(bVal) || !string.IsNullOrWhiteSpace(cVal))
                            {
                                lastDataRowOffset = i;
                            }
                        }

                        // 若未显式找到合计字样，返回最后数据行的下一行
                        if (lastDataRowOffset > 0)
                        {
                            return 6 + lastDataRowOffset;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[查看分布] FindSummaryTotalRow 异常: {ex.Message}");
            }
            // 兜底返回第 7 行
            return 7;
        }
    }
}

