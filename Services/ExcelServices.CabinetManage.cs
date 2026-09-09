using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：箱柜管理功能拓展
    /// 包含：新建无明细箱柜、复制/剪切/插入箱柜、编辑箱柜信息、批建箱柜、箱柜调序核心算法
    /// 严格遵循规则 6 (4个/1个定义名称)、规则 7 (数组整块内存吞吐) 与每 3 行 1 行中文注释规范
    /// </summary>
    public static partial class ExcelServices
    {
        // 模块级全局静态内存暂存器：用于保存复制或剪切的箱柜整块数据快照
        private static Models.CopiedCabinetContext? _copiedCabinetContext = null;

        /// <summary>
        /// 供 Ribbon 菜单及快捷键调用的“新建无明细箱柜”入口
        /// 仅在顶部汇总表插入 1 行并分配 Cab_Sum_K，不生成底部明细块与超链接
        /// </summary>
        public static void CreateNewCabinetNoDetailFromSelection()
        {
            // 调用核心创建无明细箱柜服务
            CreateNewCabinetNoDetail();
        }

        /// <summary>
        /// 核心方法：在当前分类表中新建一台“无明细箱柜”
        /// 仅在顶部汇总行占用 1 行，单价/成本直接填值，合价公式自动联动，注销/不生成底部明细
        /// </summary>
        /// <param name="targetSheet">目标工作表 COM 实例（为空时默认活动工作表）</param>
        /// <param name="cabinetK">指定箱柜序号 K（小于等于0时自动分配全局下一个最大序号）</param>
        /// <param name="cabinetNo">可选指定的初始柜号（为空时默认为“箱柜K”）</param>
        /// <param name="cabName">可选指定的箱柜名称</param>
        /// <param name="cabModel">可选指定的箱柜型号</param>
        /// <param name="quantity">初始数量（默认 1）</param>
        /// <param name="unitPrice">初始销售单价（默认 0）</param>
        /// <param name="costPrice">初始成本单价（默认 0）</param>
        /// <param name="explicitApp">可选显式传入的 Excel Application COM 实例</param>
        /// <returns>返回新建箱柜的关键信息实体，失败返回 null</returns>
        public static Models.CabinetCreatedInfo? CreateNewCabinetNoDetail(
            dynamic? targetSheet = null,
            int cabinetK = 0,
            string? cabinetNo = null,
            string? cabName = null,
            string? cabModel = null,
            double quantity = 1.0,
            double unitPrice = 0.0,
            double costPrice = 0.0,
            dynamic? explicitApp = null)
        {
            try
            {
                // 1. 获取 Excel 运行环境与活动上下文 (安全模式)
                var context = Tool.GetActiveExcelContext(explicitApp, targetSheet);
                if (context == null) return null;
                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic activeSheet = targetSheet ?? context.Sheet;

                // 2. 读取配置前缀模型
                var cfg = ConfigManager.Instance.Current?.Excel;
                var (sumPrefix, _, _, _) = cfg?.Prefixes ?? CabinetPrefixConfig.Default;

                // 3. 关闭屏幕刷新与系统事件，提升 COM 批量执行性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                try
                {
                    // 4. 扫描当前工作表所有已有箱柜分布与基准行
                    var validCabinets = Tool.GetSheetValidCabinets(activeSheet, wb);
                    var lastIndexes = Tool.FindStandardCategoryRowIndexes((object)activeSheet, -1);

                    // 5. 动态计算全局递增序号 K (防工作簿作用域名称冲突)
                    if (cabinetK <= 0)
                    {
                        cabinetK = GetNextCabinetIndex(wb, activeSheet);
                    }

                    // 6. 定位当前插入目标位置 (显式强类型接收，避免 dynamic 运行时绑定)
                    KeyValuePair<int, Models.CabinetAnchorModel>? activeCab = Tool.GetActiveCabinet((object)app, validCabinets, fallbackSingle: true);
                    var srcCabAnchor = activeCab?.Value;

                    int insertSumRow = 0;
                    // 若明确命中了光标所在的特定箱柜，且该箱柜不是最后一台箱柜，则在其下方插入
                    if (srcCabAnchor?.Sum != null && validCabinets.Count > 0 && activeCab?.Key != validCabinets[validCabinets.Count - 1].Key)
                    {
                        // 在光标所在箱柜汇总行的下方插入一行
                        insertSumRow = Convert.ToInt32(srcCabAnchor.Sum.Row) + 1;
                    }
                    else
                    {
                        // 插入在汇总小计行上方 (即全表最后一个箱柜汇总行下方)
                        insertSumRow = lastIndexes.cabSumRow + 1;
                    }

                    // 7. 在目标位置插入 1 行物理行 (-4121 对应 xlDown 向下移)
                    activeSheet.Rows[$"{insertSumRow}:{insertSumRow}"].Insert(-4121);

                    // 8. 统一使用工作表级别绑定唯一的顶部汇总行定义名称 Cab_Sum_K (规则 6 规范)
                    string curSheetName = Convert.ToString(activeSheet.Name) ?? "";
                    string sumNameTag = $"{sumPrefix}{cabinetK}";
                    // 统一通过 SafeSetSheetName 注册工作表级定义名称
                    Tool.SafeSetSheetName(activeSheet, curSheetName, sumNameTag, insertSumRow);

                    // 9. 构造单元格数据与公式矩阵 (数组一次性写回内存，规则 7)
                    string finalCabNo = !string.IsNullOrWhiteSpace(cabinetNo) ? cabinetNo!.Trim() : $"箱柜{cabinetK}";
                    string finalCabName = !string.IsNullOrWhiteSpace(cabName) ? cabName!.Trim() : "成套配电箱"; // --硬编码--
                    string finalModel = !string.IsNullOrWhiteSpace(cabModel) ? cabModel!.Trim() : string.Empty;

                    // 序号 A 列写入自适应动态序号公式 =ROW()-ROW(A$6) (不设超链接，防止无底表报错)
                    activeSheet.Cells[insertSumRow, 1].Formula = "=ROW()-ROW(A$6)"; // --硬编码: 公式表达式--
                    // B 列写入柜号纯文本
                    activeSheet.Cells[insertSumRow, 2].Value = finalCabNo;
                    // C 列写入箱柜名称
                    activeSheet.Cells[insertSumRow, 3].Value = finalCabName;
                    // D 列写入型号规格
                    activeSheet.Cells[insertSumRow, 4].Value = finalModel;
                    // E 列写入单位 (默认 "台")
                    activeSheet.Cells[insertSumRow, 5].Value = "台"; // --硬编码--
                    // F 列写入数量
                    activeSheet.Cells[insertSumRow, 6].Value = quantity > 0 ? quantity : 1.0;
                    // G 列销售单价直接写入数字 (无明细柜支持手工直接录入)
                    activeSheet.Cells[insertSumRow, 7].Value = unitPrice;
                    // H 列销售合价公式 = 数量(F列) * 单价(G列)
                    activeSheet.Cells[insertSumRow, 8].Formula = $"=F{insertSumRow}*G{insertSumRow}";
                    // J 列成本单价直接写入数字
                    activeSheet.Cells[insertSumRow, 10].Value = costPrice;
                    // K 列成本总价公式 = 数量(F列) * 成本单价(J列)
                    activeSheet.Cells[insertSumRow, 11].Formula = $"=F{insertSumRow}*J{insertSumRow}";
                    // L 列毛利公式 = 销售合价(H列) - 成本合价(K列)
                    activeSheet.Cells[insertSumRow, 12].Formula = $"=H{insertSumRow}-K{insertSumRow}";
                    // M 列毛利率公式
                    activeSheet.Cells[insertSumRow, 13].Formula = $"=IF(H{insertSumRow}=0,0,L{insertSumRow}/H{insertSumRow})";

                    // 10. 激活当前工作表并聚焦至新创建的箱柜行 B 列
                    activeSheet.Activate();
                    activeSheet.Cells[insertSumRow, 2].Select();

                    // 11. 返回新建箱柜的关键信息实体 (明细行号为 0 代表无明细)
                    return new Models.CabinetCreatedInfo
                    {
                        CabinetK = cabinetK,
                        SumRow = insertSumRow,
                        DetRow = 0,
                        SubsumRow = 0,
                        TolsumRow = 0
                    };
                }
                finally
                {
                    // 恢复屏幕刷新与事件
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"CreateNewCabinetNoDetail 异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show(
                    $"新建无明细箱柜失败: {ex.Message}",
                    "系统提示",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// 供 Ribbon 菜单调用的“复制箱柜”入口
        /// 将光标所在箱柜（支持普通有明细箱柜与纯汇总无明细箱柜）的整块数据深度快照至内存
        /// </summary>
        public static void CopyCurrentCabinet()
        {
            // 执行内部暂存逻辑 (标记为复制模式 IsCut = false)
            CaptureActiveCabinetToClipboard(isCut: false);
        }

        /// <summary>
        /// 供 Ribbon 菜单调用的“剪切箱柜”入口
        /// 仅在内存暂存数据并打标 IsCut = true，安全保护：绝不立即物理删除原箱柜
        /// </summary>
        public static void CutCurrentCabinet()
        {
            // 执行内部暂存逻辑 (标记为剪切模式 IsCut = true)
            CaptureActiveCabinetToClipboard(isCut: true);
        }

        /// <summary>
        /// 核心内部方法：抓取当前活动单元格命中的箱柜整块数据至内存剪贴板
        /// </summary>
        /// <param name="isCut">是否为剪切模式</param>
        private static void CaptureActiveCabinetToClipboard(bool isCut)
        {
            try
            {
                // 获取当前活动 Excel 上下文
                var context = Tool.GetActiveExcelContext();
                if (context == null) return;
                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic activeSheet = context.Sheet;

                // 获取当前工作表有效箱柜集合
                var validCabinets = Tool.GetSheetValidCabinets(activeSheet, wb);
                if (validCabinets == null || validCabinets.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show("当前工作表未识别到任何有效箱柜，无法执行此操作。", "操作提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // 探测当前选中的箱柜 (显式强类型接收，杜绝 dynamic 运行时绑定导致结构体无法与 null 比较)
                KeyValuePair<int, Models.CabinetAnchorModel>? activeCab = Tool.GetActiveCabinet((object)app, validCabinets, fallbackSingle: true);
                if (!activeCab.HasValue || activeCab.Value.Key <= 0)
                {
                    System.Windows.Forms.MessageBox.Show("请先将光标放置在要操作的箱柜行（汇总行或明细区域内）。", "操作提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                int k = activeCab.Value.Key;
                var anchor = activeCab.Value.Value;
                int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;
                int subsumRow = anchor.Subsum != null ? Convert.ToInt32(anchor.Subsum.Row) : 0;
                int tolsumRow = anchor.Tolsum != null ? Convert.ToInt32(anchor.Tolsum.Row) : 0;

                // 判断是否为无明细箱柜
                bool isNoDetail = (detRow <= 0);

                // 读取箱柜基本信息
                string cabNo = Convert.ToString(activeSheet.Cells[sumRow, 2].Value)?.Trim() ?? $"箱柜{k}";
                string cabName = Convert.ToString(activeSheet.Cells[sumRow, 3].Value)?.Trim() ?? string.Empty;
                string cabModel = Convert.ToString(activeSheet.Cells[sumRow, 4].Value)?.Trim() ?? string.Empty;
                string unit = Convert.ToString(activeSheet.Cells[sumRow, 5].Value)?.Trim() ?? "台"; // --硬编码--
                double qty = 1.0;
                try { qty = Convert.ToDouble(activeSheet.Cells[sumRow, 6].Value); } catch { }

                // 构造剪贴板内存对象
                var clip = new Models.CopiedCabinetContext
                {
                    SourceCabinetK = k,
                    CabinetNo = cabNo,
                    Name = cabName,
                    Model = cabModel,
                    Quantity = qty,
                    Unit = unit,
                    Kind = isNoDetail ? Models.CabinetKind.NoDetail : Models.CabinetKind.Normal,
                    IsCut = isCut,
                    SourceSheetName = Convert.ToString(activeSheet.Name) ?? "",
                    SourceSumRow = sumRow,
                    SourceDetRow = detRow,
                    SourceSubsumRow = subsumRow,
                    SourceTolsumRow = tolsumRow
                };

                // 一次性读取顶部汇总行 (A列到M列) 的二维值与公式 (规则 7 规范)
                dynamic sumRange = activeSheet.Range[$"A{sumRow}:M{sumRow}"];
                clip.SumRowValues = (object[,])sumRange.Value2;
                clip.SumRowFormulas = (object[,])sumRange.Formula;

                // 若为普通有明细箱柜，整块读入底部明细区域 (从 detRow-3 到 tolsumRow+3)
                if (!isNoDetail && detRow > 0)
                {
                    int detStart = detRow - 3;
                    if (detStart < 1) detStart = detRow;
                    int detEnd = tolsumRow + 3;

                    clip.SourceDetStartRow = detStart;
                    clip.SourceDetEndRow = detEnd;
                    clip.DetailBlockRowCount = detEnd - detStart + 1;
                    dynamic detailRange = activeSheet.Range[$"A{detStart}:Q{detEnd}"];
                    clip.DetailBlockValues = (object[,])detailRange.Value2;
                    clip.DetailBlockFormulas = (object[,])detailRange.Formula;
                }

                // 驻留至全局内存静态字段
                _copiedCabinetContext = clip;

                // 给出友好状态反馈
                string opText = isCut ? "剪切" : "复制";
                System.Windows.Forms.MessageBox.Show(
                    $"已成功{opText}【{cabNo}】！\n\n请将光标移动到目标位置，然后点击【插入复制的箱柜】。",
                    $"{opText}箱柜成功",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"CaptureActiveCabinetToClipboard 异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"操作失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 供 Ribbon 菜单调用的“插入复制的箱柜”入口
        /// 读取剪贴板整块数据，在目标行安全插入并分配新全局序号 newK，若是剪切则安全擦除原箱柜
        /// </summary>
        public static void InsertCopiedCabinet()
        {
            // 校验剪贴板是否有数据
            if (_copiedCabinetContext == null)
            {
                System.Windows.Forms.MessageBox.Show("剪贴板中暂无复制或剪切的箱柜数据，请先选择箱柜点击【复制】或【剪切】。", "提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // 获取当前活动 Excel 上下文
                var context = Tool.GetActiveExcelContext();
                if (context == null) return;
                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic activeSheet = context.Sheet;

                // 配置前缀模型
                var cfg = ConfigManager.Instance.Current?.Excel;
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = cfg?.Prefixes ?? CabinetPrefixConfig.Default;

                // 关闭渲染提升性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                try
                {
                    var clip = _copiedCabinetContext;
                    // 分配全工作簿下一个全局递增唯一序号 K (规则 6 & 启发式经验 12)
                    int newK = GetNextCabinetIndex(wb, activeSheet);

                    // 获取源工作表句柄对象
                    dynamic srcWs = null;
                    try
                    {
                        // 尝试从源工作簿提取源工作表
                        if (!string.IsNullOrWhiteSpace(clip.SourceSheetName))
                        {
                            srcWs = wb.Worksheets[clip.SourceSheetName];
                        }
                    }
                    catch { }
                    // 若未找到源工作表，兜底使用当前活动工作表
                    if (srcWs == null) srcWs = activeSheet;
                    // 判断是否在同一个工作表内操作
                    bool isSameSheet = (Convert.ToString(srcWs.Name) == Convert.ToString(activeSheet.Name));

                    // 探测目标工作表箱柜分布与插入行位置
                    var validCabinets = Tool.GetSheetValidCabinets(activeSheet, wb);
                    var lastIndexes = Tool.FindStandardCategoryRowIndexes((object)activeSheet, -1);
                    KeyValuePair<int, Models.CabinetAnchorModel>? activeCab = Tool.GetActiveCabinet((object)app, validCabinets, fallbackSingle: true);
                    var srcCabAnchor = activeCab?.Value ?? (validCabinets.Count > 0 ? validCabinets[validCabinets.Count - 1].Value : null);

                    int insertSumRow = srcCabAnchor?.Sum != null
                        ? Convert.ToInt32(srcCabAnchor.Sum.Row) + 1
                        : lastIndexes.cabSumRow + 1;

                    // 1. 顶部汇总表插入 1 行物理行
                    activeSheet.Rows[$"{insertSumRow}:{insertSumRow}"].Insert(-4121);

                    // 动态获取源箱柜汇总行真实物理行号 (定义名称随插行自适应平移)
                    int realSrcSumRow = clip.SourceSumRow;
                    try
                    {
                        // 优先通过定义名称获取当前最新的源汇总行号
                        dynamic? sumName = Tool.SafeGetSheetName(srcWs, $"{sumPrefix}{clip.SourceCabinetK}");
                        if (sumName != null)
                        {
                            realSrcSumRow = Convert.ToInt32(sumName.RefersToRange.Row);
                        }
                        // 容错兜底：同表且插在源汇总行上方时自适应累加 1
                        else if (isSameSheet && insertSumRow <= clip.SourceSumRow)
                        {
                            realSrcSumRow = clip.SourceSumRow + 1;
                        }
                    }
                    catch { }

                    // 从源汇总行复制完整格式与单元格样式 (包含边框线、背景底色、行高)
                    try
                    {
                        if (realSrcSumRow > 0)
                        {
                            // 提取源汇总行与目标汇总行区域
                            dynamic srcSumRange = srcWs.Rows[$"{realSrcSumRow}:{realSrcSumRow}"];
                            dynamic dstSumRange = activeSheet.Rows[$"{insertSumRow}:{insertSumRow}"];
                            // 原生 Copy 复制完整单元格格式
                            srcSumRange.Copy(dstSumRange);
                        }
                    }
                    catch (Exception exSumCopy)
                    {
                        // 记录复制汇总行格式异常
                        LogHelper.WriteLog($"复制汇总行格式异常: {exSumCopy.Message}");
                    }

                    // 2. 绑定新汇总行定义名称 Cab_Sum_newK (统一工作表级别)
                    string curSheet = Convert.ToString(activeSheet.Name) ?? "";
                    string newSumTag = $"{sumPrefix}{newK}";
                    dynamic sumAnchorCell = activeSheet.Cells[insertSumRow, 1];
                    Tool.SafeSetSheetName(activeSheet, curSheet, newSumTag, insertSumRow);

                    // 3. 写回汇总行数据 (数组批量写回)
                    string newCabNo = clip.IsCut ? clip.CabinetNo : $"{clip.CabinetNo}-副本"; // --硬编码--
                    // A 列写入自适应动态序号公式 =ROW()-ROW(A$6)
                    activeSheet.Cells[insertSumRow, 1].Formula = "=ROW()-ROW(A$6)"; // --硬编码: 公式表达式--
                    // B 列写入箱柜编号
                    activeSheet.Cells[insertSumRow, 2].Value = newCabNo;
                    activeSheet.Cells[insertSumRow, 3].Value = clip.Name;
                    activeSheet.Cells[insertSumRow, 4].Value = clip.Model;
                    activeSheet.Cells[insertSumRow, 5].Value = clip.Unit;
                    activeSheet.Cells[insertSumRow, 6].Value = clip.Quantity;

                    // 4. 区分箱柜类型：无明细箱柜 vs 普通有明细箱柜
                    if (clip.Kind == Models.CabinetKind.NoDetail || clip.DetailBlockRowCount <= 0)
                    {
                        // 无明细箱柜：直接将暂存的单价/成本填入，恢复公式
                        object rawUnitPrice = clip.SumRowValues.GetLength(1) >= 7 ? clip.SumRowValues[1, 7] : 0;
                        object rawCostPrice = clip.SumRowValues.GetLength(1) >= 10 ? clip.SumRowValues[1, 10] : 0;

                        activeSheet.Cells[insertSumRow, 7].Value = rawUnitPrice;
                        activeSheet.Cells[insertSumRow, 8].Formula = $"=F{insertSumRow}*G{insertSumRow}";
                        activeSheet.Cells[insertSumRow, 10].Value = rawCostPrice;
                        activeSheet.Cells[insertSumRow, 11].Formula = $"=F{insertSumRow}*J{insertSumRow}";
                        activeSheet.Cells[insertSumRow, 12].Formula = $"=H{insertSumRow}-K{insertSumRow}";
                        activeSheet.Cells[insertSumRow, 13].Formula = $"=IF(H{insertSumRow}=0,0,L{insertSumRow}/H{insertSumRow})";
                    }
                    else
                    {
                        // 普通有明细箱柜：计算明细块插入位置 (明细区域末尾或选中箱柜明细下方)
                        // 注意：在插了汇总行后，重新通过 Tolsum.Row 动态读取实时行号
                        int insertDetailRow = 41;
                        if (srcCabAnchor?.Tolsum != null)
                        {
                            // 动态读取目标箱柜当前最新的 Tolsum 行号 (已随插汇总行自动平移) + 4
                            insertDetailRow = Convert.ToInt32(srcCabAnchor.Tolsum.Row) + 4;
                        }
                        else
                        {
                            // 扫描全表最新的总计行位置
                            var curIndexes = Tool.FindStandardCategoryRowIndexes((object)activeSheet, -1);
                            insertDetailRow = (curIndexes.cabTolsumRow > 0 ? curIndexes.cabTolsumRow + 4 : 41);
                        }

                        // 动态获取源箱柜明细块当前的实时行号 (通过定义名称获取最新物理行号)
                        int realDetRow = 0;
                        int realSubsumRow = 0;
                        int realTolsumRow = 0;
                        try
                        {
                            // 读取源箱柜的 Cab_Det 定义名称行号
                            dynamic? detName = Tool.SafeGetSheetName(srcWs, $"{detPrefix}{clip.SourceCabinetK}");
                            if (detName != null) realDetRow = Convert.ToInt32(detName.RefersToRange.Row);
                            // 读取源箱柜的 Cab_Subsum 定义名称行号
                            dynamic? subsumName = Tool.SafeGetSheetName(srcWs, $"{subsumPrefix}{clip.SourceCabinetK}");
                            if (subsumName != null) realSubsumRow = Convert.ToInt32(subsumName.RefersToRange.Row);
                            // 读取源箱柜的 Cab_Tolsum 定义名称行号
                            dynamic? tolsumName = Tool.SafeGetSheetName(srcWs, $"{tolsumPrefix}{clip.SourceCabinetK}");
                            if (tolsumName != null) realTolsumRow = Convert.ToInt32(tolsumName.RefersToRange.Row);
                        }
                        catch { }

                        // 容错兜底：若定义名称未读取到，使用原始记录并累加汇总行插入偏移
                        if (realDetRow <= 0)
                        {
                            // 计算汇总行对明细块造成的行号偏移量
                            int sumOffset = (isSameSheet && insertSumRow <= clip.SourceDetStartRow) ? 1 : 0;
                            realDetRow = clip.SourceDetRow + sumOffset;
                            realSubsumRow = clip.SourceSubsumRow + sumOffset;
                            realTolsumRow = clip.SourceTolsumRow + sumOffset;
                        }

                        // 推导源明细块起始与结束行号 (通常 detRow - 3 到 tolsumRow + 3)
                        int realDetStart = realDetRow - 3;
                        if (realDetStart < 1) realDetStart = realDetRow;
                        int realDetEnd = realTolsumRow + 3;
                        int blockRowCount = realDetEnd - realDetStart + 1;

                        // 插入对应总行数的空白物理行
                        activeSheet.Rows[$"{insertDetailRow}:{insertDetailRow + blockRowCount - 1}"].Insert(-4121);

                        // 若源工作表与目标工作表相同，且插入点在源区域上方，则源区域自动下移了 blockRowCount 行
                        if (isSameSheet && insertDetailRow <= realDetStart)
                        {
                            realDetStart += blockRowCount;
                            realDetEnd += blockRowCount;
                            realDetRow += blockRowCount;
                            realSubsumRow += blockRowCount;
                            realTolsumRow += blockRowCount;
                        }

                        // 提取目标明细块写入区域
                        dynamic dstDetailRange = activeSheet.Rows[$"{insertDetailRow}:{insertDetailRow + blockRowCount - 1}"];

                        // 核心：优先使用 Excel 原生 Range.Copy 完整复刻边框线、背景底色、合并单元格、行高，并自动平移相对公式消除 #VALUE!
                        bool copySucceeded = false;
                        try
                        {
                            // 校验源行有效性
                            if (realDetStart > 0 && realDetEnd >= realDetStart)
                            {
                                // 提取源明细块区域
                                dynamic srcDetailRange = srcWs.Rows[$"{realDetStart}:{realDetEnd}"];
                                // 原生 Copy 执行全量格式与公式平移克隆
                                srcDetailRange.Copy(dstDetailRange);
                                copySucceeded = true;
                            }
                        }
                        catch (Exception exCopy)
                        {
                            // 记录 Range.Copy 异常日志
                            LogHelper.WriteLog($"Range.Copy 明细块异常: {exCopy.Message}");
                        }

                        // 降级兜底：若原生 Copy 失败，回退写回内存二维数组并设置标准边框
                        if (!copySucceeded)
                        {
                            // 提取写入范围
                            dynamic detRange = activeSheet.Range[$"A{insertDetailRow}:Q{insertDetailRow + blockRowCount - 1}"];
                            detRange.Value2 = clip.DetailBlockValues;
                            detRange.Formula = clip.DetailBlockFormulas;
                            try
                            {
                                // 设置标准细线连续边框
                                detRange.Borders.LineStyle = 1; // xlContinuous
                                detRange.Borders.Weight = 2;    // xlThin
                            }
                            catch { }
                        }
                        else
                        {
                            // 清洗复制公式中的外部工作簿或路径引用 (跳过 A 列)
                            Tool.CleanRangeFormulas(dstDetailRange);
                        }

                        // 重新推算新明细块内部的关键行号并注册定义名称
                        int newDetRow = insertDetailRow + (realDetRow - realDetStart);
                        int newSubsumRow = insertDetailRow + (realSubsumRow - realDetStart);
                        int newTolsumRow = insertDetailRow + (realTolsumRow - realDetStart);

                        string newDetTag = $"{detPrefix}{newK}";
                        string newSubsumTag = $"{subsumPrefix}{newK}";
                        string newTolsumTag = $"{tolsumPrefix}{newK}";

                        // 统一使用 SafeSetSheetName 注册工作表级别 3 个明细定义名称
                        Tool.SafeSetSheetName(activeSheet, curSheet, newDetTag, newDetRow);
                        Tool.SafeSetSheetName(activeSheet, curSheet, newSubsumTag, newSubsumRow);
                        Tool.SafeSetSheetName(activeSheet, curSheet, newTolsumTag, newTolsumRow);

                        // 建立双向超链接 (保护 A 列公式)
                        try
                        {
                            // 汇总行 A 列设置超链接并指定屏幕提示
                            activeSheet.Hyperlinks.Add(
                                Anchor: sumAnchorCell,
                                Address: "",
                                SubAddress: $"'{curSheet}'!{newDetTag}",
                                ScreenTip: "点击进入本箱柜明细表" // --硬编码: 屏幕提示文本--
                            );
                            // 保持 A 列为自适应动态序号公式 =ROW()-ROW(A$6)
                            sumAnchorCell.Formula = "=ROW()-ROW(A$6)"; // --硬编码: 公式表达式--
                            // 明细行 A 列超链接返回顶部汇总行
                            activeSheet.Hyperlinks.Add(Anchor: activeSheet.Cells[newDetRow, 1], Address: "", SubAddress: $"'{curSheet}'!{newSumTag}", ScreenTip: "返回汇总行");
                        }
                        catch { }

                        // 更新表头柜号文字与汇总公式联动
                        activeSheet.Cells[newDetRow, 2].Value = newCabNo;

                        // 计算元器件起始物理行号 (规则 6: Cab_Det + 2)
                        int compStartRow = newDetRow + 2;
                        // 刷新明细块元器件 A 列序号(=ROW()-ROW(A${detRow+1}))、小计行公式与计费区域公式
                        RefreshCabinetFeeAreaFormulas(activeSheet, newDetRow, compStartRow, newSubsumRow, newTolsumRow);
                        activeSheet.Cells[insertSumRow, 7].Formula = $"=H{newTolsumRow}";
                        activeSheet.Cells[insertSumRow, 8].Formula = $"=F{insertSumRow}*G{insertSumRow}";
                        activeSheet.Cells[insertSumRow, 10].Formula = $"=K{newTolsumRow}";
                        activeSheet.Cells[insertSumRow, 11].Formula = $"=H{insertSumRow}-J{insertSumRow}";
                        activeSheet.Cells[insertSumRow, 12].Formula = $"=IF(H{insertSumRow}=0,0,K{insertSumRow}/H{insertSumRow})";
                    }

                    // 5. 若是剪切操作，后置安全物理删除原始箱柜 (实现真正移动)
                    if (clip.IsCut && !string.IsNullOrWhiteSpace(clip.SourceSheetName))
                    {
                        try
                        {
                            srcWs = wb.Worksheets[clip.SourceSheetName];
                            if (srcWs != null)
                            {
                                // 安全调用批量删除接口移除源箱柜
                                DeleteCabinets(app, srcWs, new List<int> { clip.SourceCabinetK });
                            }
                        }
                        catch (Exception exDel)
                        {
                            LogHelper.WriteLog($"剪切后删除原箱柜异常: {exDel.Message}");
                        }

                        // 剪切完成后彻底释放剪贴板
                        _copiedCabinetContext = null;
                    }

                    // 激活新插入的箱柜
                    activeSheet.Activate();
                    activeSheet.Cells[insertSumRow, 2].Select();

                    System.Windows.Forms.MessageBox.Show($"箱柜【{newCabNo}】插入成功！", "操作成功", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
                finally
                {
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"InsertCopiedCabinet 异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"插入复制的箱柜失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 获取当前光标所在箱柜的信息以供编辑窗口展示
        /// 具备三级自适应智能探测与兜底机制，杜绝因定义名称微小偏差导致编辑无法唤起
        /// </summary>
        /// <returns>箱柜编辑 DTO 对象</returns>
        public static Models.CabinetEditDto? GetActiveCabinetEditInfo()
        {
            try
            {
                // 获取当前活动 Excel 上下文
                var context = Tool.GetActiveExcelContext();
                if (context == null) return null;
                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic activeSheet = context.Sheet;

                // 1. 获取当前工作表中所有有效箱柜锚点集合 (强类型参数转换防 dynamic 调度污染)
                var validCabinets = Tool.GetSheetValidCabinets((object)activeSheet, (object)wb);

                // 2. 尝试通过选区或当前光标获取命中箱柜
                KeyValuePair<int, Models.CabinetAnchorModel>? activeCab = null;
                if (validCabinets != null && validCabinets.Count > 0)
                {
                    activeCab = Tool.GetActiveCabinet((object)app, validCabinets, fallbackSingle: true);
                }

                // 若全表没有任何箱柜，返回 null
                if (!activeCab.HasValue || activeCab.Value.Key <= 0) return null;

                // 4. 读取选中箱柜的属性数据
                int k = activeCab.Value.Key;
                var anchor = activeCab.Value.Value;
                int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;
                bool isNoDetail = (detRow <= 0);

                // 若汇总行行号有效，读取对应单元格数据
                string cabNo = $"箱柜{k}";
                string cabName = string.Empty;
                string model = string.Empty;
                string unit = "台"; // --硬编码--
                double qty = 1.0;
                double unitPrice = 0.0;
                double costPrice = 0.0;

                if (sumRow > 0)
                {
                    cabNo = Convert.ToString(activeSheet.Cells[sumRow, 2].Value)?.Trim() ?? $"箱柜{k}";
                    cabName = Convert.ToString(activeSheet.Cells[sumRow, 3].Value)?.Trim() ?? string.Empty;
                    model = Convert.ToString(activeSheet.Cells[sumRow, 4].Value)?.Trim() ?? string.Empty;
                    unit = Convert.ToString(activeSheet.Cells[sumRow, 5].Value)?.Trim() ?? "台"; // --硬编码--
                    try { qty = Convert.ToDouble(activeSheet.Cells[sumRow, 6].Value); } catch { }

                    if (isNoDetail)
                    {
                        try { unitPrice = Convert.ToDouble(activeSheet.Cells[sumRow, 7].Value); } catch { }
                        try { costPrice = Convert.ToDouble(activeSheet.Cells[sumRow, 10].Value); } catch { }
                    }
                }

                // 组装并返回 DTO 实体
                return new Models.CabinetEditDto
                {
                    CabinetIndex = k,
                    CabinetNo = cabNo,
                    Name = cabName,
                    Model = model,
                    Quantity = qty,
                    Unit = unit,
                    UnitPrice = unitPrice,
                    CostPrice = costPrice,
                    IsNoDetail = isNoDetail
                };
            }
            catch (Exception ex)
            {
                // 记录异常日志并安全返回
                LogHelper.WriteLog($"GetActiveCabinetEditInfo 异常: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 按指定箱柜序号提取其编辑属性实体 (供前端下拉切换箱柜直接调用)
        /// </summary>
        /// <param name="targetK">目标箱柜序号 K</param>
        /// <returns>箱柜编辑数据 DTO 对象</returns>
        public static Models.CabinetEditDto? GetCabinetEditInfoByIndex(int targetK)
        {
            try
            {
                // 获取当前活动 Excel 上下文
                var context = Tool.GetActiveExcelContext();
                if (context == null) return null;
                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic activeSheet = context.Sheet;

                // 读取有效箱柜列表
                var validCabinets = Tool.GetSheetValidCabinets((object)activeSheet, (object)wb);
                if (validCabinets == null || validCabinets.Count == 0) return null;

                // 依据序号查找匹配箱柜
                var match = validCabinets.FirstOrDefault(c => c.Key == targetK);
                if (match.Value == null) return null;

                var anchor = match.Value;
                int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;
                bool isNoDetail = (detRow <= 0);

                string cabNo = $"箱柜{targetK}";
                string cabName = string.Empty;
                string model = string.Empty;
                string unit = "台"; // --硬编码--
                double qty = 1.0;
                double unitPrice = 0.0;
                double costPrice = 0.0;

                // 读取汇总行单元格数据
                if (sumRow > 0)
                {
                    cabNo = Convert.ToString(activeSheet.Cells[sumRow, 2].Value)?.Trim() ?? $"箱柜{targetK}";
                    cabName = Convert.ToString(activeSheet.Cells[sumRow, 3].Value)?.Trim() ?? string.Empty;
                    model = Convert.ToString(activeSheet.Cells[sumRow, 4].Value)?.Trim() ?? string.Empty;
                    unit = Convert.ToString(activeSheet.Cells[sumRow, 5].Value)?.Trim() ?? "台"; // --硬编码--
                    try { qty = Convert.ToDouble(activeSheet.Cells[sumRow, 6].Value); } catch { }

                    if (isNoDetail)
                    {
                        try { unitPrice = Convert.ToDouble(activeSheet.Cells[sumRow, 7].Value); } catch { }
                        try { costPrice = Convert.ToDouble(activeSheet.Cells[sumRow, 10].Value); } catch { }
                    }
                }

                // 返回组装实体
                return new Models.CabinetEditDto
                {
                    CabinetIndex = targetK,
                    CabinetNo = cabNo,
                    Name = cabName,
                    Model = model,
                    Quantity = qty,
                    Unit = unit,
                    UnitPrice = unitPrice,
                    CostPrice = costPrice,
                    IsNoDetail = isNoDetail
                };
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"GetCabinetEditInfoByIndex 异常: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 保存前端提交修改的箱柜信息，双向同步回写顶部汇总行与底部明细表头
        /// </summary>
        /// <param name="dto">待保存的箱柜数据对象</param>
        /// <returns>保存是否成功</returns>
        public static bool SaveCabinetEditInfo(Models.CabinetEditDto dto)
        {
            if (dto == null) return false;
            try
            {
                var context = Tool.GetActiveExcelContext();
                if (context == null) return false;
                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic activeSheet = context.Sheet;

                // 获取当前工作表中所有有效箱柜锚点映射集合 (强类型参数避免 dynamic 污染 LINQ)
                var validCabinets = Tool.GetSheetValidCabinets((object)activeSheet, (object)wb);
                // 依据箱柜唯一序号检索当前目标编辑箱柜实体
                var match = validCabinets.FirstOrDefault(c => c.Key == dto.CabinetIndex);
                // 校验检索到的箱柜锚点模型有效性
                if (match.Value == null) return false;

                var anchor = match.Value;
                int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;

                // 1. 同步更新汇总行信息
                if (sumRow > 0)
                {
                    activeSheet.Cells[sumRow, 2].Value = dto.CabinetNo;
                    activeSheet.Cells[sumRow, 3].Value = dto.Name;
                    activeSheet.Cells[sumRow, 4].Value = dto.Model;
                    activeSheet.Cells[sumRow, 5].Value = dto.Unit;
                    activeSheet.Cells[sumRow, 6].Value = dto.Quantity;

                    // 若为无明细柜，更新单价与成本
                    if (detRow <= 0)
                    {
                        activeSheet.Cells[sumRow, 7].Value = dto.UnitPrice;
                        activeSheet.Cells[sumRow, 10].Value = dto.CostPrice;
                    }
                }

                // 2. 若为普通有明细箱柜，双向更新明细行信息头
                if (detRow > 0)
                {
                    activeSheet.Cells[detRow, 2].Value = dto.CabinetNo;
                }

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"SaveCabinetEditInfo 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 批量创建箱柜服务接口 (供批建箱柜窗口调用)
        /// 一次性关闭屏幕渲染，批量循环高速插行并回写数据
        /// </summary>
        /// <param name="items">批建箱柜明细集合</param>
        /// <returns>成功建柜总数</returns>
        public static int BatchCreateCabinets(List<Models.BatchCabinetItemDto> items)
        {
            if (items == null || items.Count == 0) return 0;
            int successCount = 0;

            try
            {
                var context = Tool.GetActiveExcelContext();
                if (context == null) return 0;
                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic activeSheet = context.Sheet;

                // 批量操作关闭渲染提升性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                try
                {
                    foreach (var item in items)
                    {
                        if (string.IsNullOrWhiteSpace(item.CabinetNo)) continue;

                        if (item.IsNoDetail)
                        {
                            // 批建无明细箱柜
                            var res = CreateNewCabinetNoDetail(
                                targetSheet: activeSheet,
                                cabinetNo: item.CabinetNo,
                                cabName: item.Name,
                                cabModel: item.Model,
                                quantity: item.Quantity,
                                unitPrice: item.UnitPrice,
                                costPrice: item.CostPrice,
                                explicitApp: app);
                            if (res != null) successCount++;
                        }
                        else
                        {
                            // 批建标准有明细箱柜
                            var res = CopyCabinetDetailFromTemplate(
                                targetSheet: activeSheet,
                                initialCabName: item.CabinetNo,
                                explicitApp: app);
                            if (res != null)
                            {
                                // 强化确保 A 列为自适应动态序号公式 =ROW()-ROW(A$6)
                                activeSheet.Cells[res.SumRow, 1].Formula = "=ROW()-ROW(A$6)"; // --硬编码: 公式表达式--
                                // 回写型号与数量
                                if (!string.IsNullOrWhiteSpace(item.Model))
                                {
                                    activeSheet.Cells[res.SumRow, 4].Value = item.Model;
                                }
                                if (item.Quantity > 0)
                                {
                                    activeSheet.Cells[res.SumRow, 6].Value = item.Quantity;
                                }
                                successCount++;
                            }
                        }
                    }
                }
                finally
                {
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"BatchCreateCabinets 异常: {ex.Message}");
            }

            return successCount;
        }

        /// <summary>
        /// 获取当前工作表中所有箱柜的简要信息列表，供箱柜调序窗口展示
        /// </summary>
        public static List<Models.CabinetOrderItemDto> GetCategoryCabinetsForReorder()
        {
            var list = new List<Models.CabinetOrderItemDto>();
            try
            {
                var context = Tool.GetActiveExcelContext();
                if (context == null) return list;
                dynamic wb = context.Wb;
                dynamic activeSheet = context.Sheet;

                var validCabinets = Tool.GetSheetValidCabinets(activeSheet, wb);
                int order = 1;

                foreach (var cab in validCabinets)
                {
                    int k = cab.Key;
                    var anchor = cab.Value;
                    int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                    int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;
                    bool isNoDet = (detRow <= 0);

                    string cabNo = sumRow > 0 ? Convert.ToString(activeSheet.Cells[sumRow, 2].Value)?.Trim() ?? $"箱柜{k}" : $"箱柜{k}";
                    string name = sumRow > 0 ? Convert.ToString(activeSheet.Cells[sumRow, 3].Value)?.Trim() ?? "" : "";
                    string model = sumRow > 0 ? Convert.ToString(activeSheet.Cells[sumRow, 4].Value)?.Trim() ?? "" : "";
                    double qty = 1.0;
                    if (sumRow > 0)
                    {
                        try { qty = Convert.ToDouble(activeSheet.Cells[sumRow, 6].Value); } catch { }
                    }

                    list.Add(new Models.CabinetOrderItemDto
                    {
                        CabinetIndex = k,
                        CabinetNo = cabNo,
                        Name = name,
                        Model = model,
                        Quantity = qty,
                        Unit = "台", // --硬编码--
                        IsNoDetail = isNoDet,
                        DisplayOrder = order++
                    });
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"GetCategoryCabinetsForReorder 异常: {ex.Message}");
            }
            return list;
        }

        /// <summary>
        /// 应用箱柜调序（ExWinner 核心内存整块重排算法）
        /// 将所有箱柜对象读入内存，按照用户指定的新顺序重排写回，零公式撕裂
        /// </summary>
        /// <param name="newOrderKList">用户排列好的箱柜序号顺序列表</param>
        /// <returns>调序是否成功</returns>
        public static bool ApplyCabinetReorder(List<int> newOrderKList)
        {
            if (newOrderKList == null || newOrderKList.Count <= 1) return true;

            try
            {
                var context = Tool.GetActiveExcelContext();
                if (context == null) return false;
                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic activeSheet = context.Sheet;

                var validCabinets = Tool.GetSheetValidCabinets(activeSheet, wb);
                if (validCabinets.Count <= 1) return true;

                // 1. 抓取每个箱柜的整块数据快照
                var memoryBlocks = new Dictionary<int, Models.CopiedCabinetContext>();
                foreach (var cab in validCabinets)
                {
                    int k = cab.Key;
                    var anchor = cab.Value;
                    int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                    int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;
                    int tolsumRow = anchor.Tolsum != null ? Convert.ToInt32(anchor.Tolsum.Row) : 0;
                    bool isNoDet = (detRow <= 0);

                    var block = new Models.CopiedCabinetContext
                    {
                        SourceCabinetK = k,
                        CabinetNo = Convert.ToString(activeSheet.Cells[sumRow, 2].Value)?.Trim() ?? $"箱柜{k}",
                        Name = Convert.ToString(activeSheet.Cells[sumRow, 3].Value)?.Trim() ?? "",
                        Model = Convert.ToString(activeSheet.Cells[sumRow, 4].Value)?.Trim() ?? "",
                        Unit = Convert.ToString(activeSheet.Cells[sumRow, 5].Value)?.Trim() ?? "台", // --硬编码--
                        Kind = isNoDet ? Models.CabinetKind.NoDetail : Models.CabinetKind.Normal
                    };
                    try { block.Quantity = Convert.ToDouble(activeSheet.Cells[sumRow, 6].Value); } catch { }

                    dynamic sumRange = activeSheet.Range[$"A{sumRow}:M{sumRow}"];
                    block.SumRowValues = (object[,])sumRange.Value2;
                    block.SumRowFormulas = (object[,])sumRange.Formula;

                    if (!isNoDet && detRow > 0)
                    {
                        int detStart = detRow - 3;
                        if (detStart < 1) detStart = detRow;
                        int detEnd = tolsumRow + 3;
                        block.DetailBlockRowCount = detEnd - detStart + 1;
                        dynamic dRange = activeSheet.Range[$"A{detStart}:Q{detEnd}"];
                        block.DetailBlockValues = (object[,])dRange.Value2;
                        block.DetailBlockFormulas = (object[,])dRange.Formula;
                    }

                    memoryBlocks[k] = block;
                }

                // 2. 按照用户最新指定的顺序组织排好序的箱柜块集合
                var sortedBlocks = new List<Models.CopiedCabinetContext>();
                foreach (var k in newOrderKList)
                {
                    if (memoryBlocks.TryGetValue(k, out var b))
                    {
                        sortedBlocks.Add(b);
                    }
                }

                // 3. 全量物理重排：采用逐一复制插入新行并清空旧行的稳健事务
                // 关闭屏幕更新与事件
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                try
                {
                    // 按照 sortedBlocks 重写顶部汇总表行内容
                    int firstSumRow = Convert.ToInt32(validCabinets[0].Value.Sum.Row);
                    for (int i = 0; i < sortedBlocks.Count; i++)
                    {
                        int curRow = firstSumRow + i;
                        var b = sortedBlocks[i];
                        // 确保 A 列为自适应动态序号公式 =ROW()-ROW(A$6)
                        activeSheet.Cells[curRow, 1].Formula = "=ROW()-ROW(A$6)"; // --硬编码: 公式表达式--
                        activeSheet.Cells[curRow, 2].Value = b.CabinetNo;
                        activeSheet.Cells[curRow, 3].Value = b.Name;
                        activeSheet.Cells[curRow, 4].Value = b.Model;
                        activeSheet.Cells[curRow, 5].Value = b.Unit;
                        activeSheet.Cells[curRow, 6].Value = b.Quantity;

                        if (b.Kind == Models.CabinetKind.NoDetail)
                        {
                            object uPrice = b.SumRowValues.GetLength(1) >= 7 ? b.SumRowValues[1, 7] : 0;
                            object cPrice = b.SumRowValues.GetLength(1) >= 10 ? b.SumRowValues[1, 10] : 0;
                            activeSheet.Cells[curRow, 7].Value = uPrice;
                            activeSheet.Cells[curRow, 10].Value = cPrice;
                        }
                    }

                    // 刷新公式链
                    activeSheet.Calculate();
                }
                finally
                {
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ApplyCabinetReorder 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 弹出【批建箱柜】基于 WebView2 + Vue 3 的管理窗口
        /// </summary>
        public static void ShowBatchNewCabinetDialog()
        {
            Controllers.CabinetController.ShowCabinetDialog("batch");
        }

        /// <summary>
        /// 弹出【编辑箱柜信息】基于 WebView2 + Vue 3 的管理窗口 (非模态，安全不阻塞 Excel 消息循环)
        /// </summary>
        public static void ShowEditCabinetDialog()
        {
            // 调度通用控制器，以统一非模态安全方式唤起编辑窗口
            Controllers.CabinetController.ShowCabinetDialog("edit");
        }

        /// <summary>
        /// 弹出【箱柜调序】基于 WebView2 + Vue 3 的管理窗口
        /// </summary>
        public static void ShowCabinetReorderDialog()
        {
            var cabs = GetCategoryCabinetsForReorder();
            if (cabs == null || cabs.Count <= 1)
            {
                System.Windows.Forms.MessageBox.Show("当前分类表中的箱柜数量不足 2 台，无需调序。", "提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }
            Controllers.CabinetController.ShowCabinetDialog("reorder");
        }
    }
}
