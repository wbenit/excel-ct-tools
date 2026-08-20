using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：汇总调价与元件汇总表生成 (对应 summary_adjust_price.html)
    /// </summary>
    public static partial class ExcelServices
    {
        // 汇总调价窗口静态单例引用 (可空)
        private static SummaryAdjustPriceForm? _summaryAdjustPriceForm;

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“汇总调价”窗口 (非模态，可编辑 Excel)
        /// </summary>
        public static void ShowSummaryAdjustPriceDialog()
        {
            try
            {
                // 以非模态方式展示汇总调价窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _summaryAdjustPriceForm, () => new SummaryAdjustPriceForm());
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"弹出汇总调价窗口异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"弹出汇总调价窗口失败: {ex.Message}", "错误提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 遍历当前工作簿，读取所有分类工作表（Sheet）及对应箱柜台数统计
        /// </summary>
        public static List<Controllers.SummaryCategoryDto> GetCategorySheetsWithCabinetCount()
        {
            var result = new List<Controllers.SummaryCategoryDto>();

            try
            {
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return result;

                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null) return result;

                // 读取箱柜定义名称前缀值对象 (零堆分配)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

                var allNames = new List<dynamic>();
                if (activeWb.Names != null)
                {
                    try { foreach (dynamic n in activeWb.Names) allNames.Add(n); } catch { }
                }

                // 遍历工作簿中的所有工作表
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    string sheetName = Convert.ToString(sheet.Name) ?? string.Empty;
                    string trimmedName = sheetName.Trim();

                    // 精准排除明确的系统辅助表 --硬编码--
                    if (string.Equals(trimmedName, "项目信息", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var currentSheetNames = new List<dynamic>(allNames);
                    if (sheet.Names != null)
                    {
                        try { foreach (dynamic n in sheet.Names) currentSheetNames.Add(n); } catch { }
                    }

                    // 1. 优先通过定义名称构建箱柜锚点列表 (规则 6)
                    var validCabinets = Tool.BuildCabinetMap(
                        currentSheetNames,
                        sheetName,
                        sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);

                    int totalCount = 0;
                    if (validCabinets.Count > 0)
                    {
                        foreach (var cab in validCabinets)
                        {
                            int cabQty = 1;
                            try
                            {
                                if (cab.Value.Sum != null)
                                {
                                    int sumRow = Convert.ToInt32(cab.Value.Sum.Row);
                                    object qVal = sheet.Cells[sumRow, 5].Value ?? sheet.Cells[sumRow, 6].Value ?? sheet.Cells[sumRow, 4].Value;
                                    if (qVal != null && int.TryParse(Convert.ToString(qVal), out int parsedQty) && parsedQty > 0)
                                    {
                                        cabQty = parsedQty;
                                    }
                                }
                            }
                            catch { }
                            totalCount += cabQty;
                        }
                    }

                    // 2. 双轨兜底：若定义名称未统计出台数，直接智能扫描顶部箱柜汇总表区域
                    if (totalCount == 0)
                    {
                        try
                        {
                            dynamic topRange = sheet.Range["A1:H40"];
                            object[,] topMatrix = (object[,])topRange.Value2;

                            int headerRow = 0;
                            for (int r = 1; r <= 15; r++)
                            {
                                string c1 = Convert.ToString(topMatrix[r, 1]) ?? "";
                                string c2 = Convert.ToString(topMatrix[r, 2]) ?? "";
                                string c3 = Convert.ToString(topMatrix[r, 3]) ?? "";
                                if (c1.Contains("序号") || c2.Contains("序号") || c2.Contains("柜号") || c2.Contains("设备") || c3.Contains("型号"))
                                {
                                    headerRow = r;
                                    break;
                                }
                            }

                            if (headerRow > 0)
                            {
                                for (int r = headerRow + 1; r <= 38; r++)
                                {
                                    string noStr = Convert.ToString(topMatrix[r, 1]) ?? "";
                                    string nameStr = Convert.ToString(topMatrix[r, 2]) ?? "";
                                    string modelStr = Convert.ToString(topMatrix[r, 3]) ?? "";

                                    if (string.IsNullOrWhiteSpace(noStr) && string.IsNullOrWhiteSpace(nameStr) && string.IsNullOrWhiteSpace(modelStr))
                                    {
                                        continue;
                                    }
                                    if (nameStr.Contains("明细") || nameStr.Contains("元件") || nameStr.Contains("小计"))
                                    {
                                        break;
                                    }

                                    int rowQty = 1;
                                    object q1 = topMatrix[r, 5] ?? topMatrix[r, 6] ?? topMatrix[r, 4];
                                    if (q1 != null && int.TryParse(Convert.ToString(q1), out int pQty) && pQty > 0)
                                    {
                                        rowQty = pQty;
                                    }
                                    totalCount += rowQty;
                                }
                            }
                        }
                        catch { }
                    }

                    // 3. 最终兜底
                    if (totalCount == 0)
                    {
                        try
                        {
                            dynamic usedRange = sheet.UsedRange;
                            if (usedRange != null && usedRange.Rows.Count > 5)
                            {
                                totalCount = 1;
                            }
                        }
                        catch { }
                    }

                    result.Add(new Controllers.SummaryCategoryDto
                    {
                        SheetName = sheetName,
                        CabinetCount = totalCount,
                        IsSelected = true
                    });
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"获取分类工作表与箱柜数量异常: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 内存汇总聚合模型辅助类
        /// </summary>
        private class AggregatedComponent
        {
            // 元件名称 (对应 Excel B 列)
            public string Name { get; set; } = string.Empty;

            // 型号规格 (对应 Excel C 列)
            public string Model { get; set; } = string.Empty;

            // 生产厂家 / 品牌 (对应 Excel D 列)
            public string Manufacturer { get; set; } = string.Empty;

            // 单位 (对应 Excel E 列)
            public string Unit { get; set; } = string.Empty;

            // 数量 (包含箱柜台数乘积计算后的汇总总数量)
            public decimal Quantity { get; set; } = 0;

            // 元件单价 (对应 Excel G 列)
            public decimal UnitPrice { get; set; } = 0;

            // 总价 / 合计金额 (对应 Excel H 列)
            public decimal TotalPrice { get; set; } = 0;

            // 备注说明 (对应 Excel I 列)
            public string Remark { get; set; } = string.Empty;

            // 原始型号规格 (预留扩展)
            public string OriginalModel { get; set; } = string.Empty;

            // 位号 / 安装位置 (预留扩展)
            public string Position { get; set; } = string.Empty;

            // 所属分类 / 工作表 Sheet 名称
            public string Category { get; set; } = string.Empty;
        }

        /// <summary>
        /// 执行元件数据提取、内存合并聚合并在当前工作簿生成“元件汇总表”
        /// </summary>
        public static bool GenerateComponentSummarySheet(Controllers.GenerateSummaryRequest request)
        {
            if (request == null || request.SelectedSheets == null || request.SelectedSheets.Count == 0) return false;

            try
            {
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return false;

                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null) return false;

                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                // 读取箱柜定义名称前缀值对象 (零堆分配)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

                var allWbNames = new List<dynamic>();
                if (activeWb.Names != null)
                {
                    try { foreach (dynamic n in activeWb.Names) allWbNames.Add(n); } catch { }
                }

                var rawComponents = new List<AggregatedComponent>();

                // 遍历用户选中的所有分类工作表
                foreach (string sheetName in request.SelectedSheets)
                {
                    dynamic? sheet = null;
                    try
                    {
                        sheet = activeWb.Worksheets[sheetName];
                    }
                    catch { continue; }
                    if (sheet == null) continue;

                    var sheetNames = new List<dynamic>(allWbNames);
                    if (sheet.Names != null)
                    {
                        try { foreach (dynamic n in sheet.Names) sheetNames.Add(n); } catch { }
                    }

                    var validCabinets = Tool.BuildCabinetMap(
                        sheetNames,
                        sheetName,
                        sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);

                    foreach (var cab in validCabinets)
                    {
                        if (cab.Value.Det == null) continue;

                        int detRow = Convert.ToInt32(cab.Value.Det.Row);
                        int subsumRow = cab.Value.Subsum != null ? Convert.ToInt32(cab.Value.Subsum.Row) : detRow + 27;
                        int sumRow = cab.Value.Sum != null ? Convert.ToInt32(cab.Value.Sum.Row) : 0;
                        int cabQty = 1;
                        // 获取柜体数量（规则 8）
                        if (sumRow > 0)
                        {
                            try
                            {
                                object qVal = sheet.Cells[sumRow, 6].Value;
                                if (qVal != null && int.TryParse(Convert.ToString(qVal), out int pQty) && pQty > 0)
                                {
                                    cabQty = pQty;
                                }
                            }
                            catch { }
                        }

                        int compStartRow = detRow + 2;
                        int compEndRow = subsumRow - 1;
                        if (compEndRow < compStartRow) continue;

                        int rowCount = compEndRow - compStartRow + 1;

                        // 采用 2D 数组一次性批量读入内存 (规则 7)
                        dynamic compRange = sheet.Range[$"A{compStartRow}:K{compEndRow}"];
                        object[,] valMatrix = (object[,])compRange.Value2;

                        for (int r = 1; r <= rowCount; r++)
                        {
                            string compName = Convert.ToString(valMatrix[r, 2]) ?? string.Empty;

                            string model = Convert.ToString(valMatrix[r, 3]) ?? string.Empty;
                            string mfg = Convert.ToString(valMatrix[r, 4]) ?? string.Empty;
                            string unit = Convert.ToString(valMatrix[r, 5]) ?? string.Empty;

                            decimal qty = 0;
                            if (decimal.TryParse(Convert.ToString(valMatrix[r, 6]), out decimal q)) qty = q;

                            decimal unitPrice = 0;
                            if (decimal.TryParse(Convert.ToString(valMatrix[r, 7]), out decimal p)) unitPrice = p;

                            decimal totalP = 0;
                            if (decimal.TryParse(Convert.ToString(valMatrix[r, 8]), out decimal tp)) totalP = tp;
                            else totalP = qty * unitPrice;

                            string remark = Convert.ToString(valMatrix[r, 9]) ?? string.Empty;

                            decimal totalQty = qty * cabQty;
                            decimal finalTotalPrice = totalP * cabQty;

                            rawComponents.Add(new AggregatedComponent
                            {
                                Name = compName.Trim(),
                                Model = model.Trim(),
                                Manufacturer = mfg.Trim(),
                                Unit = unit.Trim(),
                                Quantity = totalQty,
                                UnitPrice = unitPrice,
                                TotalPrice = finalTotalPrice,
                                Remark = remark.Trim(),
                                Category = sheetName
                            });
                        }
                    }
                }

                // 内存合并聚合
                var mergeConditions = request.MergeConditions;
                var grouped = rawComponents.GroupBy(c =>
                {
                    string key = $"{c.Name}||{c.Model}";
                    if (mergeConditions.ByManufacturer)
                    {
                        key += $"||{(mergeConditions.IncludeNoManufacturer && string.IsNullOrWhiteSpace(c.Manufacturer) ? "" : c.Manufacturer)}";
                    }
                    if (mergeConditions.ByPrice)
                    {
                        key += $"||{c.UnitPrice}";
                    }
                    if (mergeConditions.ByRemark)
                    {
                        key += $"||{c.Remark}";
                    }
                    return key;
                }).Select(g =>
                {
                    var first = g.First();
                    decimal sumQty = g.Sum(x => x.Quantity);
                    decimal sumTotal = g.Sum(x => x.TotalPrice);
                    decimal avgPrice = sumQty > 0 ? (sumTotal / sumQty) : first.UnitPrice;

                    return new AggregatedComponent
                    {
                        Name = first.Name,
                        Model = first.Model,
                        Manufacturer = first.Manufacturer,
                        Unit = first.Unit,
                        Quantity = sumQty,
                        UnitPrice = avgPrice,
                        TotalPrice = sumTotal,
                        Remark = string.Join(";", g.Select(x => x.Remark).Where(r => !string.IsNullOrWhiteSpace(r)).Distinct()),
                        Category = first.Category
                    };
                }).ToList();

                // 排序
                if (request.SortSettings.SortType == "mfg_name_model")
                {
                    grouped = grouped.OrderBy(x => x.Manufacturer).ThenBy(x => x.Name).ThenBy(x => x.Model).ToList();
                }
                else
                {
                    grouped = grouped.OrderBy(x => x.Category).ThenBy(x => x.Name).ThenBy(x => x.Manufacturer).ThenBy(x => x.Model).ToList();
                }

                // 创建或清空“元件汇总表”
                string summarySheetName = "元件汇总表"; // --硬编码--
                dynamic summarySheet = null;
                try
                {
                    summarySheet = activeWb.Worksheets[summarySheetName];
                    summarySheet.Cells.Clear();
                }
                catch
                {
                    summarySheet = activeWb.Worksheets.Add(After: activeWb.Worksheets[activeWb.Worksheets.Count]);
                    summarySheet.Name = summarySheetName;
                }

                // 写入大标题与表头
                summarySheet.Cells[1, 1].Value = "元件汇总调价清单";
                dynamic titleRange = summarySheet.Range["A1:I1"];
                titleRange.Merge();
                titleRange.Font.Size = 14;
                titleRange.Font.Bold = true;
                titleRange.HorizontalAlignment = -4108; // 居中 xlCenter --硬编码--

                string[] headers = new string[] { "序号", "元件名称", "型号规格", "生产厂家", "单位", "总数量", "单价", "合价", "备注" };
                for (int col = 0; col < headers.Length; col++)
                {
                    dynamic headerCell = summarySheet.Cells[3, col + 1];
                    headerCell.Value = headers[col];
                    headerCell.Font.Bold = true;
                    headerCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#009688"));
                    headerCell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    headerCell.HorizontalAlignment = -4108;
                }

                // 批量回写数据
                int dataRowCount = grouped.Count;
                if (dataRowCount > 0)
                {
                    object[,] outMatrix = new object[dataRowCount, headers.Length];
                    for (int i = 0; i < dataRowCount; i++)
                    {
                        var comp = grouped[i];
                        outMatrix[i, 0] = i + 1;
                        outMatrix[i, 1] = comp.Name;
                        outMatrix[i, 2] = comp.Model;
                        outMatrix[i, 3] = comp.Manufacturer;
                        outMatrix[i, 4] = comp.Unit;
                        outMatrix[i, 5] = (double)comp.Quantity;
                        outMatrix[i, 6] = (double)comp.UnitPrice;
                        outMatrix[i, 7] = (double)comp.TotalPrice;
                        outMatrix[i, 8] = comp.Remark;
                    }

                    int startDataRow = 4;
                    int endDataRow = startDataRow + dataRowCount - 1;
                    summarySheet.Range[$"A{startDataRow}:I{endDataRow}"].Value2 = outMatrix;

                    int totalRow = endDataRow + 1;
                    summarySheet.Cells[totalRow, 2].Value = "合计";
                    summarySheet.Cells[totalRow, 2].Font.Bold = true;
                    summarySheet.Cells[totalRow, 8].Formula = $"=SUM(H{startDataRow}:H{endDataRow})";
                    summarySheet.Cells[totalRow, 8].Font.Bold = true;

                    // 设置表格全区域实线边框
                    dynamic tableRange = summarySheet.Range[$"A3:I{totalRow}"];
                    tableRange.Borders.LineStyle = 1; // 实线 --硬编码--

                    // 设置单价列与合价列的标准数值格式
                    summarySheet.Range[$"G{startDataRow}:H{totalRow}"].NumberFormat = "#,##0.00"; // 两位小数货币格式 --硬编码--
                    // 设置总数量列数值格式
                    summarySheet.Range[$"F{startDataRow}:F{endDataRow}"].NumberFormat = "0.##"; // 数量格式 --硬编码--
                }

                // 自动对齐列宽
                summarySheet.Columns["A:I"].AutoFit();
                // 激活并高亮显示元件汇总表
                summarySheet.Activate();
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"生成元件汇总表异常: {ex.Message}");
                return false;
            }
            finally
            {
                try
                {
                    dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                    if (app != null)
                    {
                        app.ScreenUpdating = true;
                        app.DisplayAlerts = true;
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// 一键更新：从“元件汇总表”将修改后的白色列（名称、型号、厂家、单价、备注等）反向同步更新至各分类表箱柜明细块（预留）
        /// </summary>
        public static bool UpdateFromComponentSummarySheet()
        {
            try
            {
                // 获取当前活动 Excel Application 实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return false;

                // 获取当前活动工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null) return false;

                // 目标汇总表名称 --硬编码--
                string summarySheetName = "元件汇总表";

                // 探测工作簿中是否存在“元件汇总表”
                dynamic? summarySheet = null;
                try
                {
                    summarySheet = activeWb.Worksheets[summarySheetName];
                }
                catch
                {
                    // 未找到汇总表直接返回
                    LogHelper.WriteLog("未找到【元件汇总表】工作表，无法执行一键更新");
                    return false;
                }

                if (summarySheet == null) return false;

                // 记录日志：一键更新预留通路已打通
                LogHelper.WriteLog("执行一键更新预留处理流程：已校验元件汇总表存在，等待后续细化批量回填算法");

                // 当前阶段返回成功，提示前端已就绪
                return true;
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"UpdateFromComponentSummarySheet 异常: {ex.Message}");
                return false;
            }
        }
    }
}
