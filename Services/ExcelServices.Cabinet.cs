using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：新建箱柜与对象模型渲染
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 供 Ribbon 菜单及右键快捷菜单调用的新建箱柜入口
        /// </summary>
        public static void CreateNewCabinetFromSelection()
        {
            // 调度核心新建箱柜业务逻辑
            CreateNewCabinet();
        }

        /// <summary>
        /// 核心方法：在当前分类表中新建箱柜
        /// 遵循规则 6（顶部汇总行、底部明细块、4 个定义名称及超链接）与规则 7（内存二维数组批量操作）
        /// </summary>
        public static void CreateNewCabinet()
        {
            try
            {
                // 获取当前运行的 Excel Application COM 接口实例
                dynamic app = ExcelDnaUtil.Application;
                if (app == null) return;

                // 获取当前激活的工作簿
                dynamic wb = app.ActiveWorkbook;
                if (wb == null) return;

                // 获取当前激活的工作表
                dynamic activeSheet = wb.ActiveSheet;
                if (activeSheet == null) return;

                // 读取全局配置参数
                var cfg = ConfigManager.Instance.Current.Excel;
                string sumPrefix = cfg.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = cfg.DetNamePrefix ?? "Cab_Det_";
                string subsumPrefix = cfg.SubsumNamePrefix ?? "Cab_Subsum_";
                string tolsumPrefix = cfg.TolsumNamePrefix ?? "Cab_Tolsum_";

                // 关闭屏幕刷新与系统弹窗以提升执行性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                int insertRow = 0;
                try
                {
                    // 1. 扫描当前工作表所有定义名称并构建有效箱柜映射 (规则 6)
                    var allNames = new List<dynamic>();
                    if (wb.Names != null)
                    {
                        try { foreach (dynamic n in wb.Names) allNames.Add(n); } catch { }
                    }
                    if (activeSheet.Names != null)
                    {
                        try { foreach (dynamic n in activeSheet.Names) allNames.Add(n); } catch { }
                    }

                    var validCabinets = Tool.BuildCabinetMap(
                        allNames,
                        Convert.ToString(activeSheet.Name) ?? "",
                        sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);

                    // 2. 动态计算下一个全新的独立箱柜序号 K
                    Microsoft.Office.Interop.Excel.Workbook excelWb = (Microsoft.Office.Interop.Excel.Workbook)wb;
                    Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)activeSheet;
                    int cabinetK = GetNextCabinetIndex(excelWb, excelSheet);

                    // 3. 定位顶部汇总行插入位置 (空行复用或插入新行)
                    int maxExistingSumRow = 0;
                    if (validCabinets.Count > 0)
                    {
                        foreach (var c in validCabinets)
                        {
                            if (c.Value.Sum != null)
                            {
                                int r = Convert.ToInt32(c.Value.Sum.Row);
                                if (r > maxExistingSumRow) maxExistingSumRow = r;
                            }
                        }
                    }

                    if (maxExistingSumRow > 0)
                    {
                        insertRow = maxExistingSumRow + 1;
                        // 判断是否需要物理插入行
                        string checkCellVal = Convert.ToString(activeSheet.Cells[insertRow, 2].Value) ?? "";
                        if (!string.IsNullOrWhiteSpace(checkCellVal))
                        {
                            activeSheet.Rows[$"{insertRow}:{insertRow}"].Insert(-4121);
                        }
                    }
                    else
                    {
                        insertRow = cfg.CabSumRowIndex;
                    }

                    // 4. 定位底部明细块插入位置 (搜索明细大标题并整块复制)
                    int lastDetBlockEnd = 0;
                    if (validCabinets.Count > 0)
                    {
                        foreach (var c in validCabinets)
                        {
                            if (c.Value.Tolsum != null)
                            {
                                int r = Convert.ToInt32(c.Value.Tolsum.Row);
                                if (r > lastDetBlockEnd) lastDetBlockEnd = r;
                            }
                        }
                    }

                    // 4. 定位底部明细块插入位置 (结合配置的 CabDetRowIndex 与 CabTolsumRowIndex 动态计算明细块总行数)
                    int templateStartRow = cfg.CabDetRowIndex - 3; // 模板明细块大标题物理行 (如 44 - 3 = 41)
                    // 结合 CabTolsumRowIndex 与 CabDetRowIndex 计算明细整块行数 (如 71 - 41 + 1 = 31 行)
                    int templateRowCount = cfg.TemplateDetailBlockTotalRows;

                    int newDetailStartRow = lastDetBlockEnd > 0 ? lastDetBlockEnd + 2 : templateStartRow + templateRowCount;

                    // 5. 复制模板明细区块直接写入目标新位置 (避免剪贴板模式冲突)
                    dynamic copyRange = activeSheet.Rows[$"{templateStartRow}:{templateStartRow + templateRowCount - 1}"];
                    dynamic targetRange = activeSheet.Rows[$"{newDetailStartRow}:{newDetailStartRow + templateRowCount - 1}"];
                    copyRange.Copy(targetRange);

                    // 6. 结合 CabDetRowIndex 与 CabTolsumRowIndex 计算新箱柜的 4 个关键行号
                    int newDetRow = newDetailStartRow + 3; // 箱柜信息行
                    int newTolsumRow = newDetailStartRow + templateRowCount - 1; // 总计行

                    // 动态获取计费区域行数 (优先读取当前表已有箱柜或公式调费组)
                    int feeSpan = 6;
                    if (validCabinets.Count > 0)
                    {
                        var firstCab = validCabinets[0].Value;
                        if (firstCab.Subsum != null && firstCab.Tolsum != null)
                        {
                            int fSub = Convert.ToInt32(firstCab.Subsum.Row);
                            int fTol = Convert.ToInt32(firstCab.Tolsum.Row);
                            if (fTol >= fSub) feeSpan = fTol - fSub + 1;
                        }
                    }

                    // 结合 CabTolsumRowIndex 与计费项数向上对齐计算小计行
                    int newSubsumRow = newTolsumRow - feeSpan + 1;
                    // 元器件起始行 (规则 6: Cab_Det + 2)
                    int newCompStartRow = newDetRow + 2;
                    // 元器件终止行 (规则 6: Cab_Subsum - 1)
                    int newCompEndRow = newSubsumRow - 1;

                    // 7. 规则 7: 调用公共方法一次性批量重写新箱柜的元器件区域 (覆盖 A~Q 列自适应空行公式)
                    if (newCompEndRow >= newCompStartRow)
                    {
                        object[,] compMatrix = Tool.BuildComponentRowsMatrix(newCompStartRow, newCompEndRow, newDetRow, 17);
                        activeSheet.Range[$"A{newCompStartRow}:Q{newCompEndRow}"].Formula = compMatrix;
                    }

                    // 8. 刷新计费区域小计行公式为自适应求和公式 (H 列与 K 列)
                    try
                    {
                        // 销售总价小计自适应求和公式
                        activeSheet.Cells[newSubsumRow, 8].Formula = $"=ROUND(SUM(H{newCompStartRow}:INDEX(H:H,ROW()-1)),2)";
                        // 成本总价小计自适应求和公式
                        activeSheet.Cells[newSubsumRow, 11].Formula = $"=ROUND(SUM(K{newCompStartRow}:INDEX(K:K,ROW()-1)),2)";
                    }
                    catch { }

                    // 9. 注册规则 6 要求的 4 个定义名称
                    string sumNameTag = $"{sumPrefix}{cabinetK}";
                    string detNameTag = $"{detPrefix}{cabinetK}";
                    string subsumNameTag = $"{subsumPrefix}{cabinetK}";
                    string tolsumNameTag = $"{tolsumPrefix}{cabinetK}";

                    Microsoft.Office.Interop.Excel.Range sumAnchorCell = (Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[insertRow, 1];
                    Microsoft.Office.Interop.Excel.Range detAnchorCell = (Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[newDetRow, 1];
                    Microsoft.Office.Interop.Excel.Range subsumAnchorCell = (Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[newSubsumRow, 1];
                    Microsoft.Office.Interop.Excel.Range tolsumAnchorCell = (Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[newTolsumRow, 1];

                    // 注册工作表级别的 4 个定义名称锚点 (规则 6)
                    // 设置箱柜汇总行定义名称
                    Tool.SafeSetSheetName(excelSheet, excelSheet.Name, sumNameTag, insertRow);
                    // 设置箱柜信息行定义名称
                    Tool.SafeSetSheetName(excelSheet, excelSheet.Name, detNameTag, newDetRow);
                    // 设置箱柜小计行定义名称
                    Tool.SafeSetSheetName(excelSheet, excelSheet.Name, subsumNameTag, newSubsumRow);
                    // 设置箱柜总计行定义名称
                    Tool.SafeSetSheetName(excelSheet, excelSheet.Name, tolsumNameTag, newTolsumRow);

                    // 10. 建立双向超链接绑定 (规则 6)
                    // 汇总行 A 列超链接跳转至明细行并显示箱柜序号
                    excelSheet.Hyperlinks.Add(
                        Anchor: sumAnchorCell,
                        Address: "",
                        SubAddress: $"'{excelSheet.Name}'!{detNameTag}",
                        TextToDisplay: Convert.ToString(cabinetK)
                    );

                    // 明细行 A 列超链接返回顶部汇总行
                    excelSheet.Hyperlinks.Add(
                        Anchor: detAnchorCell,
                        Address: "",
                        SubAddress: $"'{excelSheet.Name}'!{sumNameTag}",
                        ScreenTip: "返回汇总行"
                    );

                    // 9. 写入初始箱柜名称并同步公式
                    // 设置初始箱柜名称
                    string cabDisplayName = $"箱柜{cabinetK}";
                    activeSheet.Cells[insertRow, 2].Value = cabDisplayName;
                    activeSheet.Cells[newDetRow, 2].Value = cabDisplayName;

                    // 汇总行公式绑定至明细总计行
                    // G 列单价公式指向明细总计行的销售总价 (H 列)
                    activeSheet.Cells[insertRow, 7].Formula = $"=H{newTolsumRow - 1}";
                    // H 列总价公式 = 数量(F列) * 单价(G列)
                    activeSheet.Cells[insertRow, 8].Formula = $"=F{insertRow}*G{insertRow}";
                    // J 列成本总价公式指向明细总计行的成本总价 (K 列)
                    activeSheet.Cells[insertRow, 10].Formula = $"=K{newTolsumRow}";
                    // K 列毛利公式 = 总价 - 成本总价
                    activeSheet.Cells[insertRow, 11].Formula = $"=H{insertRow}-J{insertRow}";
                    // L 列毛利率公式
                    activeSheet.Cells[insertRow, 12].Formula = $"=IF(H{insertRow}=0,0,K{insertRow}/H{insertRow})";

                    // 10. 激活原工作表并选中新插入的汇总行
                    activeSheet.Activate();
                    activeSheet.Cells[insertRow, 2].Select();
                }
                finally
                {
                    // 恢复屏幕刷新与系统事件响应
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }
            }
            catch (Exception ex)
            {
                // 弹出异常提示
                System.Windows.Forms.MessageBox.Show($"新建箱柜异常: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 动态计算下一个全新的独立箱柜序号 K，保障所有已存在的定义名称 100% 完整保留不被覆盖
        /// </summary>
        private static int GetNextCabinetIndex(Microsoft.Office.Interop.Excel.Workbook targetWb, Microsoft.Office.Interop.Excel.Worksheet activeSheet)
        {
            int maxK = 0;

            var cfg = ConfigManager.Instance.Current.Excel;
            string sumPrefix = cfg.SumNamePrefix ?? "Cab_Sum_";
            string detPrefix = cfg.DetNamePrefix ?? "Cab_Det_";
            string subsumPrefix = cfg.SubsumNamePrefix ?? "Cab_Subsum_";
            string tolsumPrefix = cfg.TolsumNamePrefix ?? "Cab_Tolsum_";

            try
            {
                // 1. 优先扫描当前工作表中所有的工作表级定义名称，提取当前表最大序号 K
                if (activeSheet != null && activeSheet.Names != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Name n in activeSheet.Names)
                    {
                        // 提取定义名称字符串
                        string nName = Convert.ToString(n.Name) ?? "";
                        // 解析其中的数字序号
                        int k = ExtractIndexFromName(nName, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                        // 更新当前表最大序号
                        if (k > maxK) maxK = k;
                    }
                }

                // 2. 兼容扫描历史残留的工作簿级定义名称（仅提取指向当前工作表的名称）
                if (targetWb != null && targetWb.Names != null)
                {
                    string currentSheetName = Convert.ToString(activeSheet?.Name) ?? "";
                    foreach (Microsoft.Office.Interop.Excel.Name n in targetWb.Names)
                    {
                        try
                        {
                            // 校验定义名称是否属于当前工作表
                            if (n.RefersToRange != null && n.RefersToRange.Worksheet != null &&
                                string.Equals(Convert.ToString(n.RefersToRange.Worksheet.Name), currentSheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                // 提取定义名称文本
                                string nName = Convert.ToString(n.Name) ?? "";
                                // 解析序号
                                int k = ExtractIndexFromName(nName, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                                // 更新最大序号
                                if (k > maxK) maxK = k;
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return maxK + 1;
        }

        /// <summary>
        /// 将面向对象实体 CabinetObject 完整渲染写回 Excel 工作表中
        /// 遵循规则 6（行号结构与空行/插入行规则）及规则 7（内存二维数组批量读写）
        /// </summary>
        public static bool RenderCabinetObjectToSheet(dynamic sheet, Models.CabinetObject cabinet, int insertRow, int targetDetailRow, int templateBlankRows = 23)
        {
            if (sheet == null || cabinet == null || insertRow <= 0 || targetDetailRow <= 0) return false;

            try
            {
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
                int defaultCompRowCount = Math.Max(1, templateBlankRows - feeRowCount);

                // 判定实际元器件列表数量
                int compCount = cabinet.Components != null ? cabinet.Components.Count : 0;
                int compRowCount = Math.Max(defaultCompRowCount, compCount);

                // 规则 6：“如果元器件数量多于区域行数，先要插入行”
                if (compCount > defaultCompRowCount)
                {
                    int insertLineCount = compCount - defaultCompRowCount;
                    int insertStartRow = compStartRow + defaultCompRowCount;
                    sheet.Rows[$"{insertStartRow}:{insertStartRow + insertLineCount - 1}"].Insert(-4121);
                }

                // 规则 6：Cab_Subsum_k.Row - 1 为元器件终止行
                int compEndRow = compStartRow + compRowCount - 1;
                int subsumRow = compEndRow + 1;
                cabinet.SubsumAnchorRow = subsumRow;

                // 规则 6：Cab_Tolsum_k.Row 为总计行
                int tolsumRow = feeRowCount > 0 ? subsumRow + feeRowCount - 1 : subsumRow;
                cabinet.TolsumAnchorRow = tolsumRow;

                // 5. 规则 7：元器件区域采用二维数组一次性批量写入内存与 Excel (覆盖 A 列至 Q 列)
                int baseHeaderRow = compStartRow - 1;
                int cabDetRow = baseHeaderRow - 1;
                // 调用公共工具方法构建包含 F/G/H/J/K/L/N/Q 自适应公式与已有元件属性的 17 列矩阵
                object[,] compArray = Tool.BuildComponentRowsMatrix(compStartRow, compEndRow, cabDetRow, 17, cabinet.Components);

                // 批量一次性回写元器件二维数组至 A~Q 列
                dynamic compRange = sheet.Range[$"A{compStartRow}:Q{compEndRow}"];
                compRange.Formula = compArray;

                // 6. 规则 7：计费区域（从 Cab_Subsum_k.Row 至 Cab_Tolsum_k.Row）批量写入
                if (feeRowCount > 0)
                {
                    // 计费区域二维数组 (覆盖 A 列至 Q 列共 17 列)
                    object[,] feeArray = new object[feeRowCount, 17];

                    for (int j = 0; j < feeRowCount; j++)
                    {
                        var rowDef = rowDefs[j];
                        if (rowDef.Name == "总计" || rowDef.IndexTag == "总计")
                        {
                            feeArray[j, 0] = "总计";
                        }
                        else
                        {
                            feeArray[j, 0] = $"=ROW()-ROW(A${baseHeaderRow})";
                        }

                        feeArray[j, 1] = rowDef.Name ?? string.Empty;

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

                    // 覆盖写入 A 列至 Q 列完整计费二维矩阵 (规则 7)
                    dynamic feeRange = sheet.Range[$"A{subsumRow}:Q{tolsumRow}"];
                    feeRange.Formula = feeArray;
                }

                app.ScreenUpdating = prevUpdating;
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"RenderCabinetObjectToSheet 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 从 Excel 工作表中反向解析指定箱柜的 CabinetObject 实体数据模型
        /// </summary>
        public static Models.CabinetObject? ParseCabinetObjectFromSheet(dynamic sheet, int cabinetIndex)
        {
            if (sheet == null || cabinetIndex <= 0) return null;

            try
            {
                var cfg = ConfigManager.Instance.Current.Excel;
                string sumPrefix = cfg.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = cfg.DetNamePrefix ?? "Cab_Det_";

                string detTagName = $"{detPrefix}{cabinetIndex}";
                string sumTagName = $"{sumPrefix}{cabinetIndex}";

                dynamic sumRange = null;
                dynamic detRange = null;

                // 遍历寻找箱柜对应的定义名称锚点
                foreach (dynamic name in sheet.Names)
                {
                    string clean = ExtractCleanNameStr(name.Name);
                    if (string.Equals(clean, detTagName, StringComparison.OrdinalIgnoreCase)) detRange = name.RefersToRange;
                    else if (string.Equals(clean, sumTagName, StringComparison.OrdinalIgnoreCase)) sumRange = name.RefersToRange;
                }

                if (detRange == null || sumRange == null) return null;

                int detAnchorRow = detRange.Row;
                int sumAnchorRow = sumRange.Row;

                var cab = new Models.CabinetObject
                {
                    CabinetIndex = cabinetIndex,
                    DetAnchorRow = detAnchorRow,
                    SumAnchorRow = sumAnchorRow
                };

                // 反向解析 Header 表头
                int headerRow = detAnchorRow + 3;
                cab.Header.CabinetNo = Convert.ToString(sheet.Cells[headerRow, 2].Value) ?? $"箱柜{cabinetIndex}";
                cab.Header.Model = Convert.ToString(sheet.Cells[headerRow, 3].Value) ?? string.Empty;

                // 反向解析元器件列表
                int compStartRow = detAnchorRow + 5;
                int compEndRow = detAnchorRow + 26;
                int subIndex = 1;

                for (int r = compStartRow; r <= compEndRow; r++)
                {
                    string compName = Convert.ToString(sheet.Cells[r, 2].Value) ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(compName)) continue;

                    var item = new Models.ComponentItem
                    {
                        Index = subIndex++,
                        Name = compName,
                        Specification = Convert.ToString(sheet.Cells[r, 3].Value) ?? string.Empty,
                        Manufacturer = Convert.ToString(sheet.Cells[r, 4].Value) ?? string.Empty,
                        Unit = Convert.ToString(sheet.Cells[r, 5].Value) ?? string.Empty
                    };

                    if (decimal.TryParse(Convert.ToString(sheet.Cells[r, 6].Value), out decimal qty)) item.Quantity = qty;
                    if (decimal.TryParse(Convert.ToString(sheet.Cells[r, 7].Value), out decimal price)) item.UnitPrice = price;
                    if (decimal.TryParse(Convert.ToString(sheet.Cells[r, 10].Value), out decimal costPrice)) item.CostUnitPrice = costPrice;

                    cab.Components.Add(item);
                }

                return cab;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"反向解析箱柜{cabinetIndex}对象失败: {ex.Message}");
                return null;
            }
        }
    }
}
