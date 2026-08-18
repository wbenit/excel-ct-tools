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
        /// 支持显式传入外部 COM Application 与 Worksheet（方便 AutoCAD / TuFan 等外部模块直接调用）
        /// </summary>
        /// <param name="explicitApp">可选显式传入的 Excel COM Application 实例</param>
        /// <param name="explicitSheet">可选显式传入的 Worksheet 实例</param>
        /// <param name="initialCabName">可选显式传入的箱柜名称</param>
        /// <returns>新建箱柜的核心行号及序号元组，失败返回 null</returns>
        public static (int cabinetK, int sumRow, int detRow, int subsumRow, int tolsumRow)? CreateNewCabinet(
            dynamic? explicitApp = null,
            dynamic? explicitSheet = null,
            string? initialCabName = null)
        {
            try
            {
                // 获取当前运行的 Excel Application COM 接口实例 (优先使用传入实例，回退 ExcelDnaSafeAccessor)
                dynamic? app = explicitApp ?? ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return null;

                // 获取当前激活的工作簿
                dynamic wb = app.ActiveWorkbook;
                if (wb == null) return null;

                // 获取目标工作表 (优先使用传入工作表，回退活动工作表)
                dynamic activeSheet = explicitSheet ?? wb.ActiveSheet;
                if (activeSheet == null) return null;

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

                    // 探测工作表基准行号分布
                    var baseIndexes = Tool.FindStandardCategoryRowIndexes((object)activeSheet);

                    // 2. 动态计算下一个全新的独立箱柜序号 K
                    int cabinetK = GetNextCabinetIndex(wb, activeSheet);

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
                    }
                    else
                    {
                        insertRow = baseIndexes.cabSumRow + 1;
                    }

                    // 判断是否需要物理插入行
                    string checkCellVal = Convert.ToString(activeSheet.Cells[insertRow, 2].Value) ?? "";
                    if (!string.IsNullOrWhiteSpace(checkCellVal))
                    {
                        activeSheet.Rows[$"{insertRow}:{insertRow}"].Insert(-4121);
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

                    if (lastDetBlockEnd == 0)
                    {
                        lastDetBlockEnd = baseIndexes.cabTolsumRow;
                    }

                    // 动态计算模板明细块大标题物理行与整块行数
                    int templateStartRow = baseIndexes.cabDetRow - 3;
                    int templateRowCount = baseIndexes.cabTolsumRow - templateStartRow + 1;
                    if (templateRowCount <= 0) templateRowCount = cfg.TemplateDetailBlockTotalRows;

                    int newDetailStartRow = lastDetBlockEnd + 2;

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
                    else
                    {
                        feeSpan = baseIndexes.cabTolsumRow - baseIndexes.cabSubsumRow + 1;
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
                    string curSheetName = Convert.ToString(activeSheet.Name) ?? "";

                    dynamic sumAnchorCell = activeSheet.Cells[insertRow, 1];
                    dynamic detAnchorCell = activeSheet.Cells[newDetRow, 1];
                    dynamic subsumAnchorCell = activeSheet.Cells[newSubsumRow, 1];
                    dynamic tolsumAnchorCell = activeSheet.Cells[newTolsumRow, 1];

                    // 注册工作表级别的 4 个定义名称锚点 (规则 6)
                    // 设置箱柜汇总行定义名称
                    Tool.SafeSetSheetName(activeSheet, curSheetName, sumNameTag, insertRow);
                    // 设置箱柜信息行定义名称
                    Tool.SafeSetSheetName(activeSheet, curSheetName, detNameTag, newDetRow);
                    // 设置箱柜小计行定义名称
                    Tool.SafeSetSheetName(activeSheet, curSheetName, subsumNameTag, newSubsumRow);
                    // 设置箱柜总计行定义名称
                    Tool.SafeSetSheetName(activeSheet, curSheetName, tolsumNameTag, newTolsumRow);

                    // 10. 建立双向超链接绑定 (规则 6)
                    try
                    {
                        // 汇总行 A 列超链接跳转至明细行并显示箱柜序号
                        activeSheet.Hyperlinks.Add(
                            Anchor: sumAnchorCell,
                            Address: "",
                            SubAddress: $"'{curSheetName}'!{detNameTag}",
                            TextToDisplay: Convert.ToString(cabinetK)
                        );

                        // 明细行 A 列超链接返回顶部汇总行
                        activeSheet.Hyperlinks.Add(
                            Anchor: detAnchorCell,
                            Address: "",
                            SubAddress: $"'{curSheetName}'!{sumNameTag}",
                            ScreenTip: "返回汇总行"
                        );
                    }
                    catch { }

                    // 9. 写入初始箱柜名称并同步公式
                    // 设置初始箱柜名称 (优先使用传入的名称)
                    string cabDisplayName = string.IsNullOrWhiteSpace(initialCabName) ? $"箱柜{cabinetK}" : initialCabName.Trim();
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

                    // 返回新建箱柜的关键信息元组
                    return (cabinetK, insertRow, newDetRow, newSubsumRow, newTolsumRow);
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
                return null;
            }
        }

        /// <summary>
        /// 动态计算下一个全新的独立箱柜序号 K，保障所有已存在的定义名称 100% 完整保留不被覆盖
        /// </summary>
        private static int GetNextCabinetIndex(dynamic? targetWb, dynamic? activeSheet)
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
                    foreach (dynamic n in activeSheet.Names)
                    {
                        try
                        {
                            // 提取定义名称字符串
                            string nName = Convert.ToString(n.Name) ?? "";
                            // 解析其中的数字序号
                            int k = ExtractIndexFromName(nName, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                            // 更新当前表最大序号
                            if (k > maxK) maxK = k;
                        }
                        catch { }
                    }
                }

                // 2. 兼容扫描历史残留的工作簿级定义名称（仅提取指向当前工作表的名称）
                if (targetWb != null && targetWb.Names != null)
                {
                    string currentSheetName = Convert.ToString(activeSheet?.Name) ?? "";
                    foreach (dynamic n in targetWb.Names)
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

        /// <summary>
        /// 核心公共方法：批量创建/导出箱柜列表并写入活动 Excel 工作簿
        /// 支持跨分类表自动路由（基于 Category 新建或切换分类表）、箱柜插入、17列公式矩阵批量回写及定义名称管理
        /// 严格遵循规则 6（4个定义名称及超链接）与规则 7（内存二维数组批量操作）
        /// </summary>
        /// <param name="explicitApp">Excel Application COM 接口实例（可选，为空时回退 ExcelDnaUtil）</param>
        /// <param name="cabinets">待导出的箱柜对象集合</param>
        /// <param name="progressCallback">进度回调 (progressPercentage, cabinetName)</param>
        /// <returns>成功导出的箱柜总数</returns>
        public static int BatchExportCabinets(
            dynamic? explicitApp,
            List<Models.CabinetObject> cabinets,
            Action<int, string>? progressCallback = null)
        {
            // 校验待导出集合是否有效
            if (cabinets == null || cabinets.Count == 0) return 0;

            // 获取 Excel COM Application 接口 (优先支持外部传入，回退 ExcelDnaSafeAccessor)
            dynamic? app = explicitApp ?? ExcelDnaSafeAccessor.GetApplication();
            if (app == null) return 0;

            // 获取活动工作簿
            dynamic? activeWb = app.ActiveWorkbook;
            if (activeWb == null) return 0;

            // 临时保存 Excel 环境状态以支持还原
            bool prevUpdating = app.ScreenUpdating;
            bool prevAlerts = app.DisplayAlerts;
            bool prevEvents = app.EnableEvents;

            int successCount = 0;

            try
            {
                // 关闭屏幕刷新与事件调度以大幅提升批量写入性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                // 记录上一个箱柜所绑定的分类工作表名称
                string lastSheetName = string.Empty;
                // 当前操作的目标工作表引用
                dynamic? currentSheet = null;

                // 循环处理所有待导出的箱柜实体
                for (int i = 0; i < cabinets.Count; i++)
                {
                    var cab = cabinets[i];
                    // 计算进度百分比并触发回调
                    int progress = (int)((i + 1) * 100.0 / cabinets.Count);
                    progressCallback?.Invoke(progress, cab.Header.Name);

                    // 提取箱柜归属的分类表名称 (Category)
                    string targetCategory = string.IsNullOrWhiteSpace(cab.Header.Category)
                        ? string.Empty
                        : cab.Header.Category.Trim();

                    // 1. 分类表路由与新建：智能获取已有表或新建分类工作表
                    currentSheet = ResolveCategorySheet(app, activeWb, targetCategory, lastSheetName, ref lastSheetName);
                    if (currentSheet == null) continue;

                    // 2. 在目标分类工作表中渲染并写入单个箱柜
                    bool ok = ExportSingleCabinetObject(app, activeWb, currentSheet, cab);
                    if (ok) successCount++;
                }

                // 激活最终操作的分类工作表
                if (currentSheet != null)
                {
                    try { currentSheet.Activate(); } catch { }
                }
            }
            catch (Exception ex)
            {
                // 记录批量导出异常日志
                LogHelper.WriteLog($"BatchExportCabinets 异常: {ex.Message}");
            }
            finally
            {
                // 恢复 Excel 原有环境配置
                app.ScreenUpdating = prevUpdating;
                app.DisplayAlerts = prevAlerts;
                app.EnableEvents = prevEvents;
            }

            // 返回成功导出的箱柜数量
            return successCount;
        }

        /// <summary>
        /// 根据目标分类名称获取已有分类工作表或新建标准分类表
        /// </summary>
        private static dynamic? ResolveCategorySheet(
            dynamic app,
            dynamic activeWb,
            string targetCategory,
            string lastSheetName,
            ref string updatedSheetName)
        {
            // 若未填写分类名称，优先沿用上一个工作表或当前活动表
            if (string.IsNullOrWhiteSpace(targetCategory))
            {
                // 若上一个表名有效且存在
                if (!string.IsNullOrWhiteSpace(lastSheetName))
                {
                    try { return activeWb.Worksheets[lastSheetName]; } catch { }
                }

                // 否则获取当前活动工作表
                dynamic? activeSheet = activeWb.ActiveSheet;
                if (activeSheet != null)
                {
                    // 记录表名并返回
                    updatedSheetName = Convert.ToString(activeSheet.Name) ?? "";
                    return activeSheet;
                }

                // 若均无则返回第一个工作表
                dynamic firstWs = activeWb.Worksheets[1];
                updatedSheetName = Convert.ToString(firstWs.Name) ?? "";
                return firstWs;
            }

            // 若与上一个箱柜分类名称相同，直接复用上一个工作表
            if (string.Equals(targetCategory, lastSheetName, StringComparison.OrdinalIgnoreCase))
            {
                try { return activeWb.Worksheets[lastSheetName]; } catch { }
            }

            // 检查工作簿中是否已经存在同名工作表
            foreach (dynamic ws in activeWb.Worksheets)
            {
                // 忽略大小写比对表名
                if (string.Equals(Convert.ToString(ws.Name)?.Trim(), targetCategory, StringComparison.OrdinalIgnoreCase))
                {
                    // 找到了同名分类表，直接激活并复用
                    ws.Activate();
                    updatedSheetName = Convert.ToString(ws.Name) ?? "";
                    return ws;
                }
            }

            // 工作簿中不存在该分类表，调用统一创建分类表接口
            var createReq = new Models.CreateCategoryRequest
            {
                CategoryName = targetCategory,
                InitialCabinetName = "箱柜1"
            };

            // 执行新建分类表与联动项目信息表
            var res = CreateNewCategory(createReq, app);
            if (res != null && res.Success)
            {
                try
                {
                    // 获取新建的分类工作表
                    dynamic newWs = activeWb.Worksheets[targetCategory];
                    updatedSheetName = Convert.ToString(newWs.Name) ?? "";
                    return newWs;
                }
                catch { }
            }

            // 若新建失败则回退当前活动工作表
            dynamic? fallback = activeWb.ActiveSheet;
            if (fallback != null) updatedSheetName = Convert.ToString(fallback.Name) ?? "";
            return fallback;
        }

        /// <summary>
        /// 在指定的分类工作表中渲染并写入单个箱柜对象
        /// </summary>
        private static bool ExportSingleCabinetObject(
            dynamic app,
            dynamic targetWb,
            dynamic sheet,
            Models.CabinetObject cab)
        {
            if (sheet == null || cab == null) return false;

            try
            {
                string sheetName = Convert.ToString(sheet.Name) ?? "";
                string safeBoxName = string.IsNullOrWhiteSpace(cab.Header.Name)
                    ? (string.IsNullOrWhiteSpace(cab.Header.CabinetNo) ? "箱柜" : cab.Header.CabinetNo)
                    : cab.Header.Name.Trim();

                // 提取安装方式 (优先读取 InstallMode，回退 Remark)
                string installMode = string.IsNullOrWhiteSpace(cab.Header.InstallMode)
                    ? (cab.Header.Remark ?? string.Empty)
                    : cab.Header.InstallMode.Trim();

                // 1. 智能探测当前分类表的基准行号分布
                var baseIndexes = Tool.FindStandardCategoryRowIndexes((object)sheet);
                int baseSumRow = baseIndexes.cabSumRow;
                int baseDetRow = baseIndexes.cabDetRow;
                int baseSubsumRow = baseIndexes.cabSubsumRow;
                int baseTolsumRow = baseIndexes.cabTolsumRow;

                // 读取首台箱柜汇总行与明细行名称及首行元件，精准判定首个槽位是否为未使用的空白预留位置
                string sumBText = Convert.ToString(sheet.Cells[baseSumRow, 2].Value)?.Trim() ?? "";
                string detBText = Convert.ToString(sheet.Cells[baseDetRow, 2].Value)?.Trim() ?? "";
                string comp1BText = Convert.ToString(sheet.Cells[baseDetRow + 2, 2].Value)?.Trim() ?? "";
                string comp1CText = Convert.ToString(sheet.Cells[baseDetRow + 2, 3].Value)?.Trim() ?? "";

                // 首个槽位已被占用的判定条件：元器件首行有数据，或者箱柜名已非空且不为纯模板占位
                bool isFirstSlotOccupied = (!string.IsNullOrWhiteSpace(comp1BText) || !string.IsNullOrWhiteSpace(comp1CText))
                    || (!string.IsNullOrWhiteSpace(sumBText) && sumBText != "箱柜1")
                    || (!string.IsNullOrWhiteSpace(detBText) && detBText != "箱柜1");

                bool isFirstSlotEmpty = !isFirstSlotOccupied;

                int cabinetK = 1;
                int sumRow = baseSumRow;
                int detRow = baseDetRow;
                int subsumRow = baseSubsumRow;
                int tolsumRow = baseTolsumRow;

                if (!isFirstSlotEmpty)
                {
                    // 首个槽位已有数据，调用 CreateNewCabinet 创建新箱柜结构
                    var createdInfo = CreateNewCabinet(app, sheet, safeBoxName);
                    if (createdInfo.HasValue)
                    {
                        // 提取新建箱柜的关键行号与序号
                        cabinetK = createdInfo.Value.cabinetK;
                        sumRow = createdInfo.Value.sumRow;
                        detRow = createdInfo.Value.detRow;
                        subsumRow = createdInfo.Value.subsumRow;
                        tolsumRow = createdInfo.Value.tolsumRow;
                    }
                    else
                    {
                        // 记录失败日志并中断，防止覆盖首台箱柜
                        LogHelper.WriteLog($"为箱柜【{safeBoxName}】新建箱柜结构失败，跳过写入以保护已有箱柜！");
                        return false;
                    }
                }

                // 2. 计算元器件起始行与默认容量，并按规则 6 在小计行前插入行
                int compStartRow = detRow + 2;
                int defaultCompCount = subsumRow - compStartRow;
                int actualCompCount = cab.Components != null ? cab.Components.Count : 0;

                if (actualCompCount > defaultCompCount)
                {
                    // 计算需要物理插入的差额行数
                    int insertCount = actualCompCount - defaultCompCount;
                    int insertRowStart = compStartRow + defaultCompCount;

                    // 执行批量物理插行
                    sheet.Rows[$"{insertRowStart}:{insertRowStart + insertCount - 1}"].Insert(-4121);

                    // 更新行号偏移
                    subsumRow += insertCount;
                    tolsumRow += insertCount;

                    // 重新校准小计行与总计行定义名称
                    Tool.SafeSetSheetName(sheet, sheetName, $"Cab_Subsum_{cabinetK}", subsumRow);
                    Tool.SafeSetSheetName(sheet, sheetName, $"Cab_Tolsum_{cabinetK}", tolsumRow);
                }

                int compEndRow = subsumRow - 1;

                // 3. 规则 7：调用 Tool.BuildComponentRowsMatrix 一次性批量生成并写入 17 列矩阵
                if (compEndRow >= compStartRow)
                {
                    // 生成 17 列包含自适应公式与元件属性的二维矩阵 (detRow 对应箱柜信息行，使得起始序号为 1)
                    object[,] compMatrix = Tool.BuildComponentRowsMatrix(compStartRow, compEndRow, detRow, 17, cab.Components);
                    // 批量写入元器件完整区域
                    sheet.Range[$"A{compStartRow}:Q{compEndRow}"].Formula = compMatrix;

                    // 批量写入 AA 列 CAD 句柄扩展列 (用于反查联动)
                    int rowCount = compEndRow - compStartRow + 1;
                    object[,] handleMatrix = new object[rowCount, 1];
                    for (int r = 0; r < rowCount; r++)
                    {
                        if (cab.Components != null && r < cab.Components.Count && !string.IsNullOrWhiteSpace(cab.Components[r].Handle))
                        {
                            handleMatrix[r, 0] = cab.Components[r].Handle;
                        }
                        else
                        {
                            handleMatrix[r, 0] = string.Empty;
                        }
                    }
                    sheet.Range[$"AA{compStartRow}:AA{compEndRow}"].Value2 = handleMatrix;
                }

                // 4. 补充与更新箱柜基础信息（明细头行及汇总行）
                // 明细信息行 B 列写入箱柜名称
                sheet.Cells[detRow, 2].Value2 = safeBoxName;
                // 规则落实：BoxInstallMode（安装方式）写入明细信息行 I 备注列（第 9 列）
                sheet.Cells[detRow, 9].Value2 = installMode;

                if (cab.Header.MinMaxPoints != null && cab.Header.MinMaxPoints.Count > 0)
                {
                    // 明细头 AA 列写入 CAD 图元句柄/坐标范围
                    sheet.Cells[detRow, 27].Value2 = string.Join("-", cab.Header.MinMaxPoints);
                }

                // 汇总行更新箱柜属性
                sheet.Cells[sumRow, 2].Value2 = safeBoxName;
                sheet.Cells[sumRow, 3].Value2 = string.IsNullOrWhiteSpace(cab.Header.Model) ? safeBoxName : cab.Header.Model;
                sheet.Cells[sumRow, 6].Value2 = cab.Header.Quantity > 0 ? cab.Header.Quantity : 1;
                // 汇总行 M 列写入安装方式备注
                sheet.Cells[sumRow, 13].Value2 = installMode;

                // 5. 注册 4 个定义名称 (规则 6)
                string sumNameTag = $"Cab_Sum_{cabinetK}";
                string detNameTag = $"Cab_Det_{cabinetK}";
                string subsumNameTag = $"Cab_Subsum_{cabinetK}";
                string tolsumNameTag = $"Cab_Tolsum_{cabinetK}";

                Tool.SafeSetSheetName(sheet, sheetName, sumNameTag, sumRow);
                Tool.SafeSetSheetName(sheet, sheetName, detNameTag, detRow);
                Tool.SafeSetSheetName(sheet, sheetName, subsumNameTag, subsumRow);
                Tool.SafeSetSheetName(sheet, sheetName, tolsumNameTag, tolsumRow);

                // 6. 建立双向超链接绑定 (规则 6)
                try
                {
                    dynamic sumAnchorCell = sheet.Cells[sumRow, 1];
                    dynamic detAnchorCell = sheet.Cells[detRow, 1];
                    sheet.Hyperlinks.Add(
                        Anchor: sumAnchorCell,
                        Address: "",
                        SubAddress: $"'{sheetName}'!{detNameTag}",
                        TextToDisplay: Convert.ToString(cabinetK)
                    );
                    sheet.Hyperlinks.Add(
                        Anchor: detAnchorCell,
                        Address: "",
                        SubAddress: $"'{sheetName}'!{sumNameTag}",
                        ScreenTip: "返回汇总行"
                    );
                }
                catch { }

                // 7. 绑定汇总行公式指向明细总计行
                try
                {
                    sheet.Cells[sumRow, 7].Formula = $"=H{tolsumRow - 1}";
                    sheet.Cells[sumRow, 8].Formula = $"=F{sumRow}*G{sumRow}";
                    sheet.Cells[sumRow, 10].Formula = $"=K{tolsumRow}";
                    sheet.Cells[sumRow, 11].Formula = $"=H{sumRow}-J{sumRow}";
                    sheet.Cells[sumRow, 12].Formula = $"=IF(H{sumRow}=0,0,K{sumRow}/H{sumRow})";
                }
                catch { }

                // 8. 刷新小计行求和公式（H 列与 K 列）
                try
                {
                    // 销售总价小计自适应求和公式
                    sheet.Cells[subsumRow, 8].Formula = $"=ROUND(SUM(H{compStartRow}:INDEX(H:H,ROW()-1)),2)";
                    // 成本总价小计自适应求和公式
                    sheet.Cells[subsumRow, 11].Formula = $"=ROUND(SUM(K{compStartRow}:INDEX(K:K,ROW()-1)),2)";
                }
                catch { }

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ExportSingleCabinetObject 异常: {ex.Message}");
                return false;
            }
        }
    }
}
