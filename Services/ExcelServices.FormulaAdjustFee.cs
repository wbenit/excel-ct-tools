using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：公式法调费 (对应 formula_adjust_fee.html)
    /// </summary>
    public static partial class ExcelServices
    {
        // 公式法调费窗口静态单例引用 (可空)
        private static FormulaAdjustFeeForm? _formulaAdjustFeeForm;

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“公式法调费”窗口 (非模态，可编辑 Excel)
        /// </summary>
        public static void ShowFormulaAdjustFeeDialog()
        {
            try
            {
                // 以非模态方式展示公式法调费窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _formulaAdjustFeeForm, () => new FormulaAdjustFeeForm());
            }
            catch (Exception ex)
            {
                // 全局捕获异常防止 Excel 崩溃闪退
                System.Windows.Forms.MessageBox.Show($"弹出公式法调费窗口失败: {ex.Message}", "错误提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 执行“公式法调费”逻辑: 解析公式表达式并精准更新写回 Excel 目标箱柜的费用行
        /// </summary>
        /// <param name="targetScope">调费作用域 (currentCabinet/currentCategory/allCabinets/selectedCabinet)</param>
        /// <param name="groupName">选中的公式组名称</param>
        /// <param name="targetScope">目标调费作用域 (currentCabinet / currentCategory / allCabinets)</param>
        /// <param name="groupName">公式组名称</param>
        /// <param name="items">前端编辑传递的公式明细项</param>
        /// <returns>返回包含是否成功、更新工作表数、更新箱柜数、是否有警告和提示文本的执行结果元组</returns>
        public static (bool Success, int UpdatedSheets, int UpdatedCabinets, bool HasWarning, string Message) ApplyFormulaAdjustFeeToExcel(
            string targetScope,
            string groupName,
            System.Collections.Generic.List<Controllers.FormulaItemModel>? items = null)
        {
            try
            {
                // 清空上一次的箱柜警告缓存
                Tool.ClearCabinetWarnings();
                // 收集本次调费过程中跳过或未识别的箱柜
                var skippedCabinets = new List<string>();

                // 获取当前运行的 Excel Application COM 接口实例 (安全调用)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return (false, 0, 0, false, "无法连接到 Excel 应用程序。");

                // 获取当前激活的工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null) return (false, 0, 0, false, "当前没有打开的 Excel 工作簿。");

                // 若前端未显式传递 items，则从控制器读取预置公式明细
                if (items == null || items.Count == 0)
                {
                    // 实例化公式控制器
                    var controller = new Controllers.FormulaAdjustFeeController();
                    // 读取对应组名的公式明细
                    items = controller.GetFormulaDetails(groupName);
                }

                // 校验公式项集合有效性
                if (items == null || items.Count == 0)
                {
                    return (false, 0, 0, false, $"未获取到公式组【{groupName}】的明细项，请检查配置。");
                }

                // 读取 4 种定义名称前缀配置项 (零堆分配元组解构)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

                // 临时关闭刷新与提示以提升批量计算与写入性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                try
                {
                    // 1. 若作用域为“更新所有箱柜” (allCabinets): 遍历工作簿中所有的分类工作表
                    if (targetScope == "allCabinets")
                    {
                        // 记录原始活动工作表以便最后安全切回
                        dynamic? origActiveSheet = null;
                        try { origActiveSheet = activeWb.ActiveSheet; } catch { }

                        int totalSheets = 0;
                        int totalCabinets = 0;

                        // 遍历当前活动工作簿下的每一个 Worksheet
                        foreach (dynamic ws in activeWb.Worksheets)
                        {
                            try
                            {
                                // 提取工作表纯文本名称
                                string wsName = Convert.ToString(ws.Name) ?? "";
                                string trimmed = wsName.Trim();

                                // 排除明确的系统非分类辅助表 (如 项目信息、元件汇总表) --硬编码--
                                if (string.Equals(trimmed, "项目信息", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(trimmed, Models.ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }

                                // 探测当前工作表的有效箱柜集合 (显式传入所属工作簿)
                                List<KeyValuePair<int, Models.CabinetAnchorModel>> sheetCabinets = Tool.GetSheetValidCabinets((object)ws, activeWb);

                                // 若未探测到箱柜，尝试自动触发一次定义名称识别补齐
                                if (sheetCabinets.Count == 0)
                                {
                                    // 触发定义名称补齐
                                    Tool.FixAndFillCabinetNamesForSheet(ws);
                                    sheetCabinets = Tool.GetSheetValidCabinets((object)ws, activeWb);
                                }

                                // 若确认该表为分类表且包含有效箱柜，执行整表箱柜自底向上倒序调费更新
                                if (sheetCabinets.Count > 0)
                                {
                                    // 临时激活目标工作表，规避非活动表跨表执行行操作或公式写入时的 COM 异常
                                    try { ws.Activate(); } catch { }

                                    // 执行该表箱柜计费区倒序原子替换
                                    int updatedCount = UpdateCabinetsForSheet(ws, app, activeWb, sheetCabinets, items, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix, skippedCabinets);
                                    if (updatedCount > 0)
                                    {
                                        // 累计分类表数
                                        totalSheets++;
                                        // 累计更新箱柜数
                                        totalCabinets += updatedCount;
                                    }
                                }
                            }
                            catch (Exception exSheet)
                            {
                                // 单个工作表异常记录日志并继续遍历后续工作表
                                LogHelper.WriteLog($"遍历更新工作表调费异常: {exSheet.Message}");
                            }
                        }

                        // 安全切回原本激活的工作表
                        try { origActiveSheet?.Activate(); } catch { }

                        // 汇集跳过的箱柜与底层未匹配方案的全部警告
                        var allWarnings = new List<string>();
                        if (skippedCabinets.Count > 0) allWarnings.AddRange(skippedCabinets);
                        var toolWarnings = Tool.GetCabinetWarnings();
                        if (toolWarnings.Count > 0) allWarnings.AddRange(toolWarnings);
                        allWarnings = allWarnings.Distinct().ToList();
                        bool hasWarning = allWarnings.Count > 0;

                        // 统计结果并返回
                        if (totalCabinets > 0)
                        {
                            string msg = hasWarning
                                ? $"成功更新 {totalSheets} 个分类表，共 {totalCabinets} 个箱柜。\n⚠️ 提示（部分箱柜已跳过）：\n" + string.Join("\n", allWarnings)
                                : $"成功更新 {totalSheets} 个分类表，共 {totalCabinets} 个箱柜！";
                            return (true, totalSheets, totalCabinets, hasWarning, msg);
                        }
                        else
                        {
                            string msg = hasWarning
                                ? $"未在当前工作簿中成功更新任何箱柜计费区。\n⚠️ 原因：\n" + string.Join("\n", allWarnings)
                                : "未在当前工作簿中识别到包含有效箱柜的分类表。";
                            return (false, 0, 0, hasWarning, msg);
                        }
                    }
                    // 2. 否则为“当前箱柜” (currentCabinet) 或“当前分类” (currentCategory): 仅在当前活动表执行
                    else
                    {
                        // 获取当前活动工作表
                        dynamic activeSheet = activeWb.ActiveSheet;
                        if (activeSheet == null) return (false, 0, 0, false, "当前没有打开或激活的 Excel 工作表。");

                        // 构建当前工作表有效箱柜映射
                        List<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets = Tool.GetSheetValidCabinets((object)activeSheet, activeWb);
                        // 若有效箱柜为空自动触发单表识别补齐
                        if (validCabinets.Count == 0)
                        {
                            // 补齐定义名称
                            Tool.FixAndFillCabinetNamesForSheet(activeSheet);
                            validCabinets = Tool.GetSheetValidCabinets((object)activeSheet, activeWb);
                        }

                        // 校验是否识别到有效箱柜
                        if (validCabinets.Count == 0)
                        {
                            return (false, 0, 0, false, "当前工作表未识别到有效箱柜，请确认是否为标准分类表。");
                        }

                        // 初始化目标箱柜集合
                        var targetCabinets = new List<KeyValuePair<int, Models.CabinetAnchorModel>>();

                        // 依据目标作用域筛选
                        if (targetScope == "currentCabinet")
                        {
                            // 智能匹配当前光标所属的单个箱柜
                            KeyValuePair<int, Models.CabinetAnchorModel>? matched = Tool.GetActiveCabinet(app, validCabinets, fallbackSingle: true);
                            if (matched.HasValue)
                            {
                                targetCabinets.Add(matched.Value);
                            }
                            else
                            {
                                return (false, 0, 0, false, "未识别到当前光标所在的箱柜。");
                            }
                        }
                        else
                        {
                            // 包含当前分类表中具备底表明细的所有箱柜 (安全过滤排除纯汇总箱柜)
                            targetCabinets.AddRange(validCabinets.Where(c => c.Value?.Det != null));
                        }

                        // 执行当前工作表目标箱柜调费更新
                        int updated = UpdateCabinetsForSheet(activeSheet, app, activeWb, targetCabinets, items, sumPrefix, detPrefix, subsumPrefix, tolsumPrefix, skippedCabinets);

                        // 汇集跳过的箱柜与底层未匹配方案的全部警告
                        var allWarnings = new List<string>();
                        if (skippedCabinets.Count > 0) allWarnings.AddRange(skippedCabinets);
                        var toolWarnings = Tool.GetCabinetWarnings();
                        if (toolWarnings.Count > 0) allWarnings.AddRange(toolWarnings);
                        allWarnings = allWarnings.Distinct().ToList();
                        bool hasWarning = allWarnings.Count > 0;

                        if (updated > 0)
                        {
                            // 组织反馈描述文本
                            string desc = targetScope == "currentCabinet" ? $"箱柜(序号 {targetCabinets[0].Key})" : $"当前分类共 {updated} 个箱柜";
                            string msg = hasWarning
                                ? $"成功更新{desc}的公式计费区间！\n⚠️ 提示：\n" + string.Join("\n", allWarnings)
                                : $"成功更新{desc}的公式计费区间！";
                            return (true, 1, updated, hasWarning, msg);
                        }
                        else
                        {
                            string msg = hasWarning
                                ? $"未能成功更新箱柜计费区间。\n⚠️ 原因：\n" + string.Join("\n", allWarnings)
                                : "未能成功更新任何箱柜的计费区间。";
                            return (false, 0, 0, hasWarning, msg);
                        }
                    }
                }
                finally
                {
                    // 安全恢复 Excel 屏幕刷新与系统提示
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                }
            }
            catch (Exception ex)
            {
                // 异常日志记录 (包含详细调用堆栈，便于快速排查)
                LogHelper.WriteLog($"执行公式法调费异常: {ex.Message}\n堆栈跟踪: {ex.StackTrace}");
                return (false, 0, 0, false, $"执行公式法调费失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 私有辅助方法: 在指定的工作表上按行号倒序批量更新目标箱柜的计费区域 (原子覆盖替换)
        /// 遵循规则 6 与规则 7
        /// </summary>
        /// <returns>成功更新的箱柜数量</returns>
        private static int UpdateCabinetsForSheet(
            dynamic sheet,
            dynamic app,
            dynamic activeWb,
            List<KeyValuePair<int, Models.CabinetAnchorModel>> targetCabinets,
            List<Controllers.FormulaItemModel> items,
            string sumPrefix,
            string detPrefix,
            string subsumPrefix,
            string tolsumPrefix,
            List<string>? skippedWarnings = null)
        {
            // 校验工作表与目标箱柜列表有效性
            if (sheet == null || targetCabinets == null || targetCabinets.Count == 0) return 0;

            // 获取新公式组总项数 N
            int N = items.Count;
            // 统计成功更新的箱柜数
            int updatedCount = 0;

            // 路线 1 增强：调费前先执行一次定义名称与计费起止行自愈校准，彻底纠偏历史可能存在的行号漂移
            Tool.FixAndFillCabinetNamesForSheet(sheet);
            // 显式强类型接收，彻底切断 dynamic 传染以支持 LINQ 静态编译
            List<KeyValuePair<int, Models.CabinetAnchorModel>> latestCabinets = Tool.GetSheetValidCabinets((object)sheet, (object?)activeWb);
            var targetDict = new HashSet<int>(targetCabinets.Select(c => c.Key));

            // 规则：多箱柜批量调费必须自底向上 (按箱柜物理行号降序) 遍历
            // 确保下方箱柜的增删行完全不会破坏上方箱柜在 Excel 中的物理行号
            // 安全防护：过滤排除无底表明细行 Det 的纯汇总箱柜，杜绝 dynamic 为 null 时抛出运行时绑定异常
            List<KeyValuePair<int, Models.CabinetAnchorModel>> sortedCabinets = latestCabinets
                .Where(c => targetDict.Contains(c.Key) && c.Value?.Det != null)
                .OrderByDescending(c => Convert.ToInt32(c.Value.Det.Row))
                .ToList();

            // 若自愈刷新后未匹配到有效列表，回退原传入列表 (同样安全过滤 Det != null)
            if (sortedCabinets.Count == 0)
            {
                sortedCabinets = targetCabinets
                    .Where(c => c.Value?.Det != null)
                    .OrderByDescending(c => Convert.ToInt32(c.Value.Det.Row))
                    .ToList();
            }

            // 遍历当前工作表中的所有目标箱柜
            foreach (var cab in sortedCabinets)
            {
                // 提取箱柜数字序号
                int k = cab.Key;
                // 防御性校验：若箱柜无底表明细信息行 Det，跳过计费区更新
                if (cab.Value?.Det == null) continue;
                // 提取箱柜信息行物理行号
                int cabDetRow = Convert.ToInt32(cab.Value.Det.Row);
                // 安全提取小计行物理行号 (基于 SUM+INDEX 双关键词判定)
                int oldSubsumRow = cab.Value.Subsum != null ? Convert.ToInt32(cab.Value.Subsum.Row) : 0;
                // 安全提取总计行物理行号 (基于 G 列公式引用其他行判定)
                int oldTolsumRow = cab.Value.Tolsum != null ? Convert.ToInt32(cab.Value.Tolsum.Row) : 0;

                // 若定义名称缺失或行号倒挂，自动调用双锚点识别重新校准
                if (oldSubsumRow <= 0 || oldTolsumRow <= oldSubsumRow)
                {
                    // 触发当前工作表的定义名称校准
                    Tool.FixAndFillCabinetNamesForSheet(sheet);
                    // 重新提取当前工作表有效箱柜
                    List<KeyValuePair<int, Models.CabinetAnchorModel>> refreshedCabinets = Tool.GetSheetValidCabinets((object)sheet, activeWb);
                    // 遍历提取当前序号对应的新锚点
                    foreach (var refreshedCab in refreshedCabinets)
                    {
                        if (refreshedCab.Key == k && refreshedCab.Value?.Subsum != null && refreshedCab.Value?.Tolsum != null)
                        {
                            oldSubsumRow = Convert.ToInt32(refreshedCab.Value.Subsum.Row);
                            oldTolsumRow = Convert.ToInt32(refreshedCab.Value.Tolsum.Row);
                            break;
                        }
                    }
                }

                // 校验校准后行号有效性，无效则记录警告并跳过防止破坏表格
                if (oldSubsumRow <= 0 || oldTolsumRow <= oldSubsumRow)
                {
                    // 提取当前工作表名称
                    string curWsName = Convert.ToString(sheet.Name) ?? "";
                    // 收集跳过箱柜的警告提示
                    skippedWarnings?.Add($"【{curWsName}】箱柜 [{k}]：未能定位有效小计行或总计行，已跳过");
                    continue;
                }

                // 计算元器件起始行 (依据规则 6: Cab_Det + 2)
                int compStartRow = cabDetRow + 2;
                // 计算原旧计费区间的总行数
                int oldM = oldTolsumRow - oldSubsumRow + 1;

                // 安全防线 1：校验旧计费行数合理性 (正常计费项 4~12 项，若 > 15 或侵入元器件区则拦截)
                if (oldM > 15 || oldSubsumRow < compStartRow)
                {
                    // 提取当前工作表纯文本名称
                    string curWsName = Convert.ToString(sheet.Name) ?? "";
                    // 收集跳过箱柜的警告提示
                    skippedWarnings?.Add($"【{curWsName}】箱柜 [{k}]：旧计费区域行数异常({oldM}行)，疑似侵入元器件区，已安全跳过更新以防止误删明细");
                    // 记录安全拦截日志
                    LogHelper.WriteLog($"[调费安全拦截] 工作表 [{curWsName}] 箱柜 [{k}] oldM={oldM}, oldSubsumRow={oldSubsumRow}, compStartRow={compStartRow}，超出安全阈值，终止更新！");
                    continue;
                }

                // 计算新旧计费行数差额 (delta > 0 需插行，delta < 0 需删行)
                int delta = N - oldM;

                // 差额插入或删除行以对齐计费行数 (规则 6: 计费区域可以替换，不能有空行)
                if (delta > 0)
                {
                    // 在旧总计行处向下插入差额空白行以对齐空间 (总计行自然下移，保留底边框)
                    sheet.Rows[$"{oldTolsumRow}:{oldTolsumRow + delta - 1}"].Insert(-4121);
                }
                else if (delta < 0)
                {
                    // 差额删除多余行: 在总计行上方删除，绝不删除总计行本身，确保总计行底边框完好上浮
                    int deleteCount = -delta;
                    int delStart = oldTolsumRow - deleteCount;
                    int delEnd = oldTolsumRow - 1;

                    // 安全防线 2：绝对安全红线拦截，严禁删行侵入元器件区域 (delStart 必须大于等于 oldSubsumRow)
                    if (delStart < oldSubsumRow)
                    {
                        // 提取当前工作表名称
                        string curWsName = Convert.ToString(sheet.Name) ?? "";
                        // 收集删行越界警告提示
                        skippedWarnings?.Add($"【{curWsName}】箱柜 [{k}]：删行边界异常(起始行{delStart}小于计费首行{oldSubsumRow})，已安全拦截跳过");
                        // 记录越界拦截日志
                        LogHelper.WriteLog($"[调费安全拦截] 工作表 [{curWsName}] 箱柜 [{k}] delStart={delStart} < oldSubsumRow={oldSubsumRow}，已阻止删行！");
                        continue;
                    }

                    // 安全执行计费区域内部多余行物理删除
                    sheet.Rows[$"{delStart}:{delEnd}"].Delete(-4121);
                }

                // 新小计起始物理行号保持与旧计费起点对齐
                int newSubsumRow = oldSubsumRow;
                // 新总计行物理行号
                int newTolsumRow = newSubsumRow + N - 1;
                // 元器件终止行 (依据规则 6: Cab_Subsum - 1)
                int compEndRow = newSubsumRow - 1;

                // 构建 17 列完整二维计费公式矩阵 (规则 7: 内存一次性生成)
                object[,] feeMatrix = Tool.BuildFeeMatrix(items, cabDetRow, newSubsumRow, compStartRow, compEndRow, 17);

                // 批量一次性覆盖写入 Excel 计费区域 (彻底替换旧计费区域)
                dynamic feeRange = sheet.Range[$"A{newSubsumRow}:Q{newTolsumRow}"];
                feeRange.Formula = feeMatrix;

                // 确保总计行底边框实线完好 (xlEdgeBottom = -4107, xlContinuous = 1, xlThin = 2) --硬编码--
                try
                {
                    // 获取总计行 Range 区域
                    dynamic tolsumRange = sheet.Range[$"A{newTolsumRow}:Q{newTolsumRow}"];
                    // 设置底边框为连续实线
                    tolsumRange.Borders[-4107].LineStyle = 1;
                    // 设置底边框线宽为细线
                    tolsumRange.Borders[-4107].Weight = 2;
                }
                catch { }

                // 获取当前工作表纯文本名称
                string sheetName = Convert.ToString(sheet.Name) ?? "";
                // 覆盖更新当前箱柜小计行定义名称 (规则 6)
                Tool.SafeSetSheetName(sheet, sheetName, $"{subsumPrefix}{k}", newSubsumRow);
                // 覆盖更新当前箱柜总计行定义名称 (规则 6)
                Tool.SafeSetSheetName(sheet, sheetName, $"{tolsumPrefix}{k}", newTolsumRow);

                // 同步更新顶部汇总行公式联动
                if (cab.Value.Sum != null)
                {
                    // 读取顶部汇总行行号
                    int sumRow = Convert.ToInt32(cab.Value.Sum.Row);
                    // G 列销售单价公式指向明细总计行的销售总价 H 列
                    sheet.Cells[sumRow, 7].Formula = $"=H{newTolsumRow}";
                    // J 列成本单价公式指向明细总计行的成本总价 K 列
                    sheet.Cells[sumRow, 10].Formula = $"=K{newTolsumRow}";
                }

                // 成功计数累加
                updatedCount++;
            }

            // 返回成功更新的箱柜总数
            return updatedCount;
        }

        /// <summary>
        /// 公式法调费设为默认：将新的默认公式组计费行写回标准模板 CabinetTemplate.xlsx 中的【分类1】工作表
        /// 遵循规则：先删除模板中旧的计费行，采用汇总行对齐的方式，写入模板新的计费行，并更新定义名称与汇总行公式
        /// </summary>
        /// <param name="items">公式明细项集合</param>
        /// <param name="groupName">公式组名称</param>
        /// <param name="explicitApp">可选传入的 Excel Application COM 实例</param>
        /// <returns>操作是否成功</returns>
        public static bool UpdateCabinetTemplateDefaultFee(
            System.Collections.Generic.List<Controllers.FormulaItemModel>? items = null,
            string? groupName = null,
            dynamic? explicitApp = null)
        {
            // 声明模板工作簿句柄
            dynamic? templateWb = null;
            // 声明 Excel Application COM 句柄
            dynamic? app = null;

            try
            {
                // 1. 获取 Excel Application COM 接口实例 (安全调用)
                app = explicitApp ?? ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return false;

                // 2. 若入参 items 为空，从控制器读取当前默认公式组明细
                if (items == null || items.Count == 0)
                {
                    // 实例化公式控制器
                    var controller = new Controllers.FormulaAdjustFeeController();
                    // 提取指定组或默认组明细
                    items = !string.IsNullOrWhiteSpace(groupName)
                        ? controller.GetFormulaDetails(groupName)
                        : controller.GetFormulaDetails("多费用公式");
                }

                // 校验明细集合有效性
                if (items == null || items.Count == 0)
                {
                    LogHelper.WriteLog("更新模板计费行失败: 公式明细集合为空");
                    return false;
                }

                // 3. 获取 CabinetTemplate.xlsx 模板物理文件路径
                string templatePath = Controllers.ProjectController.EnsureCabinetTemplate(app);
                if (!System.IO.File.Exists(templatePath))
                {
                    LogHelper.WriteLog($"未找到模板物理文件: {templatePath}");
                    return false;
                }

                // 确保模板文件不是只读属性
                try
                {
                    var fileInfo = new System.IO.FileInfo(templatePath);
                    if (fileInfo.IsReadOnly) fileInfo.IsReadOnly = false;
                }
                catch { }

                // 4. 读取配置与定义名称前缀
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;
                string defaultTemplateSheet = ConfigManager.Instance.Current.Excel.DefaultTemplateSheet ?? "分类1";

                // 5. 以可写方式打开模板工作簿
                templateWb = app.Workbooks.Open(templatePath, UpdateLinks: 0, ReadOnly: false);
                if (templateWb == null) return false;

                // 获取目标模板工作表 (优先分类1，若无取第2个或第1个)
                dynamic? catSheet = null;
                try { catSheet = templateWb.Sheets[defaultTemplateSheet]; } catch { }
                if (catSheet == null)
                {
                    try { catSheet = templateWb.Sheets.Count >= 2 ? templateWb.Sheets[2] : templateWb.Sheets[1]; } catch { }
                }
                if (catSheet == null)
                {
                    templateWb.Close(false);
                    return false;
                }

                string sheetName = Convert.ToString(catSheet.Name) ?? defaultTemplateSheet;

                // 6. 探测模板工作表中箱柜 1 的标准行号分布 (复用 Tool 公共方法)
                var rowIndexes = Tool.FindStandardCategoryRowIndexes((object)catSheet, 1);
                int cabSumRow = rowIndexes.cabSumRow;
                int cabDetRow = rowIndexes.cabDetRow;
                int oldSubsumRow = rowIndexes.cabSubsumRow;
                int oldTolsumRow = rowIndexes.cabTolsumRow;

                // 若有任意一个标准行号无效 (<= 0)，调用现成方法修复定义名称并重新获取所有标准行
                if (cabSumRow <= 0 || cabDetRow <= 0 || oldSubsumRow <= 0 || oldTolsumRow <= 0)
                {
                    // 调用现成方法自动补齐与修复当前工作表定义名称
                    Tool.FixAndFillCabinetNamesForSheet(catSheet);
                    // 重新获取所有标准行分布
                    rowIndexes = Tool.FindStandardCategoryRowIndexes((object)catSheet, 1);
                    // 重新赋值汇总行号
                    cabSumRow = rowIndexes.cabSumRow;
                    // 重新赋值明细信息行号
                    cabDetRow = rowIndexes.cabDetRow;
                    // 重新赋值小计行号
                    oldSubsumRow = rowIndexes.cabSubsumRow;
                    // 重新赋值总计行号
                    oldTolsumRow = rowIndexes.cabTolsumRow;
                }

                int compStartRow = cabDetRow + 2;
                int newFeeRows = items.Count;

                // 7. 直接清空旧计费行的 B~Q 列数据 (A 列序号与定义名称保留不删)
                if (oldTolsumRow >= oldSubsumRow && oldSubsumRow > 0)
                {
                    // 清空旧计费区域 B 列至 Q 列数据
                    catSheet.Range[$"B{oldSubsumRow}:Q{oldTolsumRow}"].ClearContents();
                }

                // 8. 采用汇总行对齐方式计算新计费行物理区间: 以总计行向上对齐计算小计行
                int newSubsumRow = oldTolsumRow - newFeeRows + 1;
                // 新总计行保持与模板汇总行对齐
                int newTolsumRow = oldTolsumRow;
                // 计算元器件终止行
                int compEndRow = newSubsumRow - 1;

                // 9. 构建计费矩阵并批量一次性写入模板计费区域 (覆盖 A 列至 Q 列)
                object[,] feeMatrix = Tool.BuildFeeMatrix(items, cabDetRow, newSubsumRow, compStartRow, compEndRow, 17);
                // 覆盖写入新计费区域
                dynamic feeRange = catSheet.Range[$"A{newSubsumRow}:Q{newTolsumRow}"];
                feeRange.Formula = feeMatrix;

                // 10. 为元器件区域 (compStartRow 到 compEndRow) 重新刷入自适应公式矩阵 (保证 A~Q 列纯净)
                int compRowCount = compEndRow - compStartRow + 1;
                if (compRowCount > 0)
                {
                    // 生成自适应公式矩阵
                    object[,] compMatrix = Tool.BuildComponentRowsMatrix(compStartRow, compEndRow, cabDetRow, 17);
                    // 批量覆盖写回元器件区域
                    catSheet.Range[$"A{compStartRow}:Q{compEndRow}"].Formula = compMatrix;
                }

                // 11. 重新注册并更新模板工作表 4 个定义名称锚点
                Tool.SafeSetSheetName(catSheet, sheetName, $"{sumPrefix}1", cabSumRow);
                Tool.SafeSetSheetName(catSheet, sheetName, $"{detPrefix}1", cabDetRow);
                Tool.SafeSetSheetName(catSheet, sheetName, $"{subsumPrefix}1", newSubsumRow);
                Tool.SafeSetSheetName(catSheet, sheetName, $"{tolsumPrefix}1", newTolsumRow);

                // 12. 采用汇总行对齐的方式: 重新对齐绑定顶部汇总行 (cabSumRow, Row 7) 引用公式与超链接
                catSheet.Cells[cabSumRow, 7].Formula = $"=H{newTolsumRow}";
                catSheet.Cells[cabSumRow, 8].Formula = $"=F{cabSumRow}*G{cabSumRow}";
                catSheet.Cells[cabSumRow, 10].Formula = $"=K{newTolsumRow}";
                catSheet.Cells[cabSumRow, 11].Formula = $"=H{cabSumRow}-J{cabSumRow}";
                catSheet.Cells[cabSumRow, 12].Formula = $"=IF(H{cabSumRow}=0,0,K{cabSumRow}/H{cabSumRow})";

                // 双向超链接绑定 (严格挂载于 A 列，B 列箱柜名称严禁添加超链接)
                try
                {
                    // 汇总行 A 列超链接指向明细行
                    dynamic sumAnchor = catSheet.Cells[cabSumRow, 1];
                    catSheet.Hyperlinks.Add(
                        Anchor: sumAnchor,
                        Address: "",
                        SubAddress: $"{sheetName}!A{cabDetRow}",
                        ScreenTip: "点击进入本箱柜明细表" // --硬编码: 屏幕提示文本--
                    );
                    // 汇总行 A 列序号自适应动态公式
                    sumAnchor.Formula = "=ROW()-ROW(A$6)"; // --硬编码: 公式表达式--

                    // 明细行 A 列超链接返回顶部汇总行
                    catSheet.Hyperlinks.Add(
                        Anchor: catSheet.Cells[cabDetRow, 1],
                        Address: "",
                        SubAddress: $"{sheetName}!A{cabSumRow}",
                        ScreenTip: "返回汇总行" // --硬编码: 屏幕提示文本--
                    );

                    // 确保汇总行与明细行 B 列从源头杜绝任何超链接
                    try { catSheet.Cells[cabSumRow, 2].Hyperlinks.Delete(); } catch { }
                    try { catSheet.Cells[cabDetRow, 2].Hyperlinks.Delete(); } catch { }
                }
                catch { }

                // 13. 保存模板工作簿
                templateWb.Save();
                LogHelper.WriteLog($"成功将默认公式组计费行写回模板: {templatePath} (总计行对齐至 Row {newTolsumRow})");

                // 14. 若开发目录存在 CabinetTemplate.xlsx 副本，同步更新覆盖
                try
                {
                    string currentDevPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "CabinetTemplate.xlsx");
                    if (!string.Equals(templatePath, currentDevPath, StringComparison.OrdinalIgnoreCase) && System.IO.File.Exists(templatePath))
                    {
                        System.IO.File.Copy(templatePath, currentDevPath, true);
                    }
                }
                catch { }

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"更新模板默认计费行异常: {ex.Message}");
                return false;
            }
            finally
            {
                // 确保关闭模板工作簿句柄
                if (templateWb != null)
                {
                    try { templateWb.Close(false); } catch { }
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(templateWb); } catch { }
                }
            }
        }
    }
}
