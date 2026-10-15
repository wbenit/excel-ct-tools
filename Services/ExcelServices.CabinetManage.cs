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

                    // 6. 定位当前插入目标位置：统一调用 DetermineTopSummaryInsertRow 实现光标行上方插入且绝不越界
                    int insertSumRow = DetermineTopSummaryInsertRow(app, activeSheet, validCabinets);

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
        /// <summary>
        /// 核心内部方法：抓取指定或当前活动单元格命中的箱柜整块数据至内存上下文模型
        /// 遵循规范：每 3 行代码至少包含 1 行中文注释，遵循规则 7 数组批量读取
        /// </summary>
        /// <param name="app">Excel Application 实例</param>
        /// <param name="activeSheet">当前工作表 COM 对象</param>
        /// <param name="wb">当前工作簿 COM 对象</param>
        /// <param name="specificK">指定箱柜序号 (<=0 时自动根据选区智能探测)</param>
        /// <param name="isCut">是否为剪切模式</param>
        /// <returns>抓取到的箱柜快照实体，未识别到返回 null</returns>
        public static Models.CopiedCabinetContext? CaptureCabinetContext(
            dynamic app,
            dynamic activeSheet,
            dynamic wb,
            int specificK = 0,
            bool isCut = false)
        {
            try
            {
                // 1. 获取当前工作表有效箱柜集合
                var validCabinets = Tool.GetSheetValidCabinets((object)activeSheet, (object)wb);
                // 校验工作表是否存在有效箱柜
                if (validCabinets == null || validCabinets.Count == 0)
                {
                    // 弹出友好提示
                    System.Windows.Forms.MessageBox.Show("当前工作表未识别到任何有效箱柜，无法执行此操作。", "操作提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return null;
                }

                // 2. 探测目标箱柜实体
                KeyValuePair<int, Models.CabinetAnchorModel>? targetCab = null;
                // 若显式指定了箱柜序号 K，直接按序号查找
                if (specificK > 0)
                {
                    // 按序号匹配
                    targetCab = validCabinets.FirstOrDefault(c => c.Key == specificK);
                }
                else
                {
                    // 否则通过光标选区智能识别命中的箱柜
                    targetCab = Tool.GetActiveCabinet((object)app, validCabinets, fallbackSingle: true);
                }

                // 校验命中箱柜的有效性
                if (!targetCab.HasValue || targetCab.Value.Key <= 0)
                {
                    // 提示用户定位光标
                    System.Windows.Forms.MessageBox.Show("请先将光标放置在要操作的箱柜行（汇总行或明细区域内）。", "操作提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return null;
                }

                // 3. 提取箱柜锚点与物理行号 (优先使用缓存的纯物理行号，兼顾 COM 属性读取)
                int k = targetCab.Value.Key;
                var anchor = targetCab.Value.Value;
                int sumRow = anchor.SumRow > 0 ? anchor.SumRow : (anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0);
                int detRow = anchor.DetRow > 0 ? anchor.DetRow : (anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0);
                int subsumRow = anchor.SubsumRow > 0 ? anchor.SubsumRow : (anchor.Subsum != null ? Convert.ToInt32(anchor.Subsum.Row) : 0);
                int tolsumRow = anchor.TolsumRow > 0 ? anchor.TolsumRow : (anchor.Tolsum != null ? Convert.ToInt32(anchor.Tolsum.Row) : 0);

                // 4. 读取箱柜顶部汇总基本信息 (优先读取以便支持后续物理表头匹配)
                string cabNo = Convert.ToString(activeSheet.Cells[sumRow, 2].Value)?.Trim() ?? $"箱柜{k}";
                string cabName = Convert.ToString(activeSheet.Cells[sumRow, 3].Value)?.Trim() ?? string.Empty;
                string cabModel = Convert.ToString(activeSheet.Cells[sumRow, 4].Value)?.Trim() ?? string.Empty;
                string unit = Convert.ToString(activeSheet.Cells[sumRow, 5].Value)?.Trim() ?? "台"; // --硬编码: 默认箱柜单位--
                double qty = 1.0;
                try { qty = Convert.ToDouble(activeSheet.Cells[sumRow, 6].Value); } catch { }

                // 健壮性保障 1：若 detRow 缺失，尝试通过定义名称精确提取
                if (detRow <= 0)
                {
                    // 尝试通过定义名称直接查找
                    var cfgPre = ConfigManager.Instance.Current?.Excel?.Prefixes ?? CabinetPrefixConfig.Default;
                    dynamic? detNameObj = Tool.SafeGetSheetName(activeSheet, $"{cfgPre.DetPrefix}{k}");
                    if (detNameObj != null)
                    {
                        try { detRow = Convert.ToInt32(detNameObj.RefersToRange.Row); } catch { }
                    }
                }

                // 健壮性保障 2：若定义名称未查到，尝试从汇总行 A 列超链接解析底表明细行
                if (detRow <= 0 && sumRow > 0)
                {
                    try
                    {
                        dynamic sumCell = activeSheet.Cells[sumRow, 1];
                        if (sumCell.Hyperlinks != null && sumCell.Hyperlinks.Count > 0)
                        {
                            string subAddr = Convert.ToString(sumCell.Hyperlinks.Item(1).SubAddress) ?? "";
                            if (!string.IsNullOrEmpty(subAddr))
                            {
                                // 清洗提取纯名称标识
                                string cleanTarget = subAddr.Replace("'", "").Replace($"{Convert.ToString(activeSheet.Name)}!", "");
                                dynamic? linkedName = Tool.SafeGetSheetName(activeSheet, cleanTarget);
                                if (linkedName != null)
                                {
                                    try { detRow = Convert.ToInt32(linkedName.RefersToRange.Row); } catch { }
                                }
                                // 若超链接直接包含绝对单元格行号引用 (如 $A$3153)
                                if (detRow <= 0)
                                {
                                    var m = System.Text.RegularExpressions.Regex.Match(subAddr, @"\$[A-Za-z]+\$(\d+)");
                                    if (m.Success && int.TryParse(m.Groups[1].Value, out int pr))
                                    {
                                        detRow = pr;
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                }

                // 健壮性保障 3：若仍未获取，以柜号文本在底表物理检索匹配配电箱明细表头
                if (detRow <= 0 && !string.IsNullOrWhiteSpace(cabNo))
                {
                    try
                    {
                        // 提取工作表已用区域下边界
                        dynamic usedRange = activeSheet.UsedRange;
                        int usedEndRow = Convert.ToInt32(usedRange.Row) + Convert.ToInt32(usedRange.Rows.Count) - 1;
                        int scanStart = Math.Max(sumRow + 2, 35); // --硬编码: 底表最小起始扫描行--

                        // 仅当底表区域存在可用行时执行快速检索
                        if (usedEndRow >= scanStart)
                        {
                            dynamic searchArea = activeSheet.Range[$"A{scanStart}:C{usedEndRow}"];
                            // 优先全字精确匹配柜号 (xlWhole = 1, xlValues = -4163)
                            dynamic found = searchArea.Find(
                                What: cabNo,
                                LookIn: -4163,
                                LookAt: 1,
                                SearchOrder: 1,
                                SearchDirection: 1,
                                MatchCase: false);

                            // 容错部分匹配
                            if (found == null)
                            {
                                found = searchArea.Find(
                                    What: cabNo,
                                    LookIn: -4163,
                                    LookAt: 2, // xlPart = 2
                                    SearchOrder: 1,
                                    SearchDirection: 1,
                                    MatchCase: false);
                            }

                            if (found != null)
                            {
                                int foundRow = Convert.ToInt32(found.Row);
                                // 校验下一行是否包含元器件明细表头特征
                                string nA = Convert.ToString(activeSheet.Cells[foundRow + 1, 1].Value2)?.Trim() ?? "";
                                string nB = Convert.ToString(activeSheet.Cells[foundRow + 1, 2].Value2)?.Trim() ?? "";
                                string nC = Convert.ToString(activeSheet.Cells[foundRow + 1, 3].Value2)?.Trim() ?? "";
                                if (nA.Contains("序号") || nA.Contains("项次") || nB.Contains("元件") || nB.Contains("名称") || nC.Contains("型号") || nC.Contains("规格") || nA.Contains("NO"))
                                {
                                    // 成功通过物理表头锁定明细行
                                    detRow = foundRow;
                                }
                            }
                        }
                    }
                    catch { }
                }

                // 健壮性保障 4：若 detRow 有效但 tolsumRow 缺失，向下安全探测总计行
                if (detRow > 0 && tolsumRow <= detRow)
                {
                    // 向下探测总计行 (通常位于 detRow + 5 至 detRow + 80 范围)
                    for (int r = detRow + 5; r <= detRow + 80; r++)
                    {
                        string cA = Convert.ToString(activeSheet.Cells[r, 1].Value2)?.Trim() ?? "";
                        string cB = Convert.ToString(activeSheet.Cells[r, 2].Value2)?.Trim() ?? "";
                        if (cA.Contains("总计") || cB.Contains("总计") || cA.Contains("合计") || cB.Contains("合计"))
                        {
                            tolsumRow = r;
                            break;
                        }
                    }
                    if (tolsumRow <= detRow) tolsumRow = detRow + 30; // --硬编码: 默认总计偏移兜底30行--
                }

                // 健壮性保障 5：若 detRow 有效但 subsumRow 缺失，向上安全探测小计行
                if (detRow > 0 && subsumRow <= detRow && tolsumRow > detRow)
                {
                    // 在 [detRow + 2, tolsumRow - 1] 范围内倒序探测小计行
                    for (int r = tolsumRow - 1; r >= detRow + 2; r--)
                    {
                        string cA = Convert.ToString(activeSheet.Cells[r, 1].Value2)?.Trim() ?? "";
                        string cB = Convert.ToString(activeSheet.Cells[r, 2].Value2)?.Trim() ?? "";
                        if (cA.Contains("小计") || cB.Contains("小计"))
                        {
                            subsumRow = r;
                            break;
                        }
                    }
                    if (subsumRow <= detRow) subsumRow = tolsumRow - 7; // --硬编码: 默认小计偏移兜底7行--
                }

                // 判断是否为无明细箱柜 (规则 6: detRow <= 0 为无明细)
                bool isNoDetail = (detRow <= 0);

                // 5. 构造整柜复制上下文模型实体
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

                // 6. 规则 7: 一次性读取顶部汇总行 (A列到M列) 的二维数组值与公式
                dynamic sumRange = activeSheet.Range[$"A{sumRow}:M{sumRow}"];
                clip.SumRowValues = (object[,])sumRange.Value2;
                clip.SumRowFormulas = (object[,])sumRange.Formula;

                // 7. 若为普通有明细箱柜，整块读入底部明细区域 (从 detRow-3 到 tolsumRow+3)
                if (!isNoDetail && detRow > 0)
                {
                    // 明细块起始物理行
                    int detStart = detRow - 3;
                    if (detStart < 1) detStart = detRow;
                    // 明细块结束物理行
                    int detEnd = tolsumRow + 3;

                    // 填充明细区间与行数
                    clip.SourceDetStartRow = detStart;
                    clip.SourceDetEndRow = detEnd;
                    clip.DetailBlockRowCount = detEnd - detStart + 1;
                    // 规则 7: 一次性将明细块大矩形加载至内存二维数组
                    dynamic detailRange = activeSheet.Range[$"A{detStart}:Q{detEnd}"];
                    clip.DetailBlockValues = (object[,])detailRange.Value2;
                    clip.DetailBlockFormulas = (object[,])detailRange.Formula;
                }

                // 返回构建就绪的整柜快照
                return clip;
            }
            catch (Exception ex)
            {
                // 记录快照提取异常日志
                LogHelper.WriteLog($"[CaptureCabinetContext] 提取箱柜快照异常: {ex.Message}");
                return null;
            }
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

                // 调用通用提取方法获取箱柜上下文快照
                var clip = CaptureCabinetContext(app, activeSheet, wb, specificK: 0, isCut: isCut);
                if (clip == null) return;

                // 驻留至全局内存静态字段
                _copiedCabinetContext = clip;

                // 给出友好状态反馈
                string opText = isCut ? "剪切" : "复制";
                System.Windows.Forms.MessageBox.Show(
                    $"已成功{opText}【{clip.CabinetNo}】！\n\n请将光标移动到目标位置，然后点击【插入复制的箱柜】。",
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
        /// 跨分类表完整复制当前选中的箱柜（右键菜单级联直发）
        /// 满足用户核心指令：跨分类表复制、名称不变、末尾追加、策略A计费区克隆、保留CAD句柄
        /// </summary>
        /// <param name="targetSheetName">目标分类表名称</param>
        /// <returns>操作是否成功</returns>
        public static bool CopyActiveCabinetToSheet(string targetSheetName)
        {
            try
            {
                // 1. 校验目标分类表名称非空
                if (string.IsNullOrWhiteSpace(targetSheetName)) return false;

                // 2. 获取当前活动 Excel 上下文
                var context = Tool.GetActiveExcelContext();
                if (context == null) return false;
                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic srcSheet = context.Sheet;

                // 3. 校验目标表不能与当前工作表重名
                string srcSheetName = Convert.ToString(srcSheet.Name) ?? "";
                if (string.Equals(srcSheetName, targetSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    System.Windows.Forms.MessageBox.Show("目标分类表不能是当前所在的分类表！", "提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }

                // 4. 定位目标分类表 COM 实例
                dynamic? targetSheet = null;
                try
                {
                    targetSheet = wb.Worksheets[targetSheetName];
                }
                catch { }

                // 校验目标工作表是否存在
                if (targetSheet == null)
                {
                    System.Windows.Forms.MessageBox.Show($"未找到目标分类表【{targetSheetName}】！", "提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    return false;
                }

                // 5. 抓取当前活动箱柜的完整快照 (模式为复制 isCut = false)
                var clip = CaptureCabinetContext(app, srcSheet, wb, specificK: 0, isCut: false);
                if (clip == null) return false;

                // 6. 调用统一克隆执行引擎 (跨表复制强制 appendToEnd = true，在目标表最后一个箱柜后末尾追加，彻底脱离光标依赖)
                var result = ExecuteCabinetInsertOrClone(app, wb, clip, targetSheet, keepName: true, appendToEnd: true);
                if (result == null) return false;

                // 7. 自动切换到目标分类表并将视口平滑定位至新箱柜明细
                try
                {
                    targetSheet.Activate();
                    if (result.DetRow > 0)
                    {
                        targetSheet.Cells[result.DetRow, 2].Select();
                    }
                    else if (result.SumRow > 0)
                    {
                        targetSheet.Cells[result.SumRow, 2].Select();
                    }
                }
                catch { }

                // 8. 状态栏与轻量弹窗反馈
                string tipMsg = $"箱柜【{clip.CabinetNo}】已成功复制到分类表【{targetSheetName}】！";
                try { app.StatusBar = $"✓ {tipMsg}"; } catch { }
                System.Windows.Forms.MessageBox.Show(tipMsg, "跨表复制成功", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return true;
            }
            catch (Exception ex)
            {
                // 记录跨表复制异常日志
                LogHelper.WriteLog($"[CopyActiveCabinetToSheet] 异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"复制箱柜失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// 核心安全算法：定位目标分类表中顶部汇总行的物理插入行号
        /// 严格遵循用户最新指令：需要在顶部汇总的光标行插入上方，且严格限制在顶部汇总区内部，绝不越界至底表明细区
        /// </summary>
        /// <param name="app">Excel Application 实例</param>
        /// <param name="targetSheet">目标工作表 COM 实例</param>
        /// <param name="validCabinets">目标工作表现存已识别箱柜集合</param>
        /// <param name="preferredCursorRow">显式优先指定的光标行号 (如右键上下文行，<=0 则从 ActiveCell 获取)</param>
        /// <returns>计算出的安全物理插入行号 (在该行执行 xlShiftDown，实现光标行上方插入)</returns>
        public static int DetermineTopSummaryInsertRow(
            dynamic app,
            dynamic targetSheet,
            List<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets,
            int preferredCursorRow = 0,
            bool appendToEnd = false)
        {
            // 0. 读取配置中的基准汇总起始物理行号 (默认 7)
            int cabSumStartRow = ConfigManager.Instance.Current?.Excel?.CabSumRowIndex ?? 7;
            if (cabSumStartRow <= 0) cabSumStartRow = 7; // --硬编码: 首个箱柜汇总基准行 7--

            // 1. 动态探测底表明细块的最小起始行 (分界线)
            int minDetRow = int.MaxValue;
            if (validCabinets != null && validCabinets.Count > 0)
            {
                // 遍历所有箱柜锚点检索最小明细行
                foreach (var pair in validCabinets)
                {
                    if (pair.Value?.Det != null)
                    {
                        try
                        {
                            // 提取明细行物理行号
                            int dR = Convert.ToInt32(pair.Value.Det.Row);
                            // 动态维护最小明细起始行号
                            if (dR > 0 && dR < minDetRow) minDetRow = dR;
                        }
                        catch { }
                    }
                }
            }

            // 2. 真实物理单元格扫描：探测顶部汇总区最后一个具有实质箱柜内容 (柜号或名称非空) 的物理行
            // 规则 7: 使用数组一次性读取顶部区域，避免逐行 COM 往返调用
            int lastPhysicalCabinetRow = 0;
            try
            {
                // 扫描上界：在明细行上方截断，支持数百台大型工程汇总表全量扫描
                int scanLimitRow = minDetRow != int.MaxValue ? (minDetRow - 1) : 300;
                if (scanLimitRow >= cabSumStartRow)
                {
                    // 一次性读取顶部汇总区 A 列至 C 列单元格数据到二维数组
                    dynamic scanRange = targetSheet.Range[targetSheet.Cells[cabSumStartRow, 1], targetSheet.Cells[scanLimitRow, 3]];
                    object[,]? arr = scanRange.Value2 as object[,];
                    if (arr != null)
                    {
                        int rowsCount = arr.GetLength(0);
                        for (int rIdx = 1; rIdx <= rowsCount; rIdx++)
                        {
                            int actualRow = cabSumStartRow + rIdx - 1;
                            string colA = Convert.ToString(arr[rIdx, 1])?.Trim() ?? "";
                            string colB = Convert.ToString(arr[rIdx, 2])?.Trim() ?? "";
                            string colC = Convert.ToString(arr[rIdx, 3])?.Trim() ?? "";

                            // 命中汇总区底部落款关键字立即截断，绝不穿透至小计、总计或明细大标题
                            if (colA.Contains("小计") || colB.Contains("小计") || colA.Contains("合计") || colB.Contains("合计") ||
                                colA.Contains("总计") || colB.Contains("总计") || colB.Contains("包装") || colB.Contains("运输") ||
                                colA.Contains("明细表") || colB.Contains("明细表") || colA.Contains("说明") || colB.Contains("说明"))
                            {
                                break;
                            }

                            // 排除重复出现的表头文字
                            if (colA.Contains("序号") || colB.Contains("柜号")) continue;

                            // 核心判定：B 列(柜号)或 C 列(箱柜名称)必须具有实质文字内容
                            if (!string.IsNullOrWhiteSpace(colB) || !string.IsNullOrWhiteSpace(colC))
                            {
                                // 记录具有真实箱柜内容的物理行号
                                lastPhysicalCabinetRow = Math.Max(lastPhysicalCabinetRow, actualRow);
                            }
                        }
                    }
                }
            }
            catch { }

            // 3. 统计定义名称中真正具备实质内容或对应明细块的最大汇总行 (严密剔除死名称与残留空行)
            int maxValidSumRow = 0;
            if (validCabinets != null && validCabinets.Count > 0)
            {
                foreach (var pair in validCabinets)
                {
                    if (pair.Value?.Sum != null)
                    {
                        try
                        {
                            int sR = Convert.ToInt32(pair.Value.Sum.Row);
                            if (sR >= 6 && sR < minDetRow)
                            {
                                // 必须小于等于实际内容行，或者单元格有内容，或者拥有真实明细块
                                bool hasRealData = (sR <= lastPhysicalCabinetRow);
                                if (!hasRealData)
                                {
                                    // 检查单元格实质内容
                                    string b = Convert.ToString(targetSheet.Cells[sR, 2].Value2)?.Trim() ?? "";
                                    string c = Convert.ToString(targetSheet.Cells[sR, 3].Value2)?.Trim() ?? "";
                                    hasRealData = (!string.IsNullOrWhiteSpace(b) || !string.IsNullOrWhiteSpace(c));
                                }
                                // 若具备真实数据或具有底表明细块，方可计入有效最大汇总行
                                if (hasRealData || (pair.Value?.Det != null && pair.Value.DetRow > 0))
                                {
                                    if (sR > maxValidSumRow) maxValidSumRow = sR;
                                }
                            }
                        }
                        catch { }
                    }
                }
            }

            // 4. 双轨合并计算顶部汇总区最后一个有效箱柜基准物理行 (优先物理真实内容)
            int effectiveLastSumRow = Math.Max(lastPhysicalCabinetRow, maxValidSumRow);

            // 核心业务分支：若明确指定了末尾追加模式（如跨表复制/导出到其他分类表）
            // 彻底脱离光标依赖，直接在现有最后一个有效箱柜汇总行后紧随追加 (+1)，杜绝跳行产生空行
            if (appendToEnd)
            {
                if (effectiveLastSumRow >= 6)
                {
                    int appendRow = effectiveLastSumRow + 1;
                    if (appendRow >= minDetRow)
                    {
                        return Math.Max(6, minDetRow - 1);
                    }
                    return appendRow;
                }
                return cabSumStartRow;
            }

            // 5. 获取当前活动光标行号
            int cursorRow = preferredCursorRow;
            if (cursorRow <= 0)
            {
                try
                {
                    dynamic activeCell = app?.ActiveCell;
                    if (activeCell != null)
                    {
                        cursorRow = Convert.ToInt32(activeCell.Row);
                    }
                }
                catch { }
            }

            // 6. 判定光标位置并计算插入行号：
            // 分支 A: 光标明确落在现有有效箱柜区间内 (行号 >= 6 且 <= effectiveLastSumRow 且小于 minDetRow)
            // 在光标行上方插入
            if (cursorRow >= 6 && cursorRow <= effectiveLastSumRow && cursorRow < minDetRow)
            {
                return cursorRow;
            }

            // 分支 B: 光标处于底表明细区内部 (行号 >= minDetRow)
            if (cursorRow >= minDetRow && validCabinets != null && validCabinets.Count > 0)
            {
                try
                {
                    KeyValuePair<int, Models.CabinetAnchorModel>? activeCab = Tool.GetActiveCabinet((object)app, validCabinets, fallbackSingle: false);
                    if (activeCab != null && activeCab.Value.Value?.Sum != null)
                    {
                        int sumR = Convert.ToInt32(activeCab.Value.Value.Sum.Row);
                        if (sumR >= 6 && sumR < minDetRow)
                        {
                            return sumR;
                        }
                    }
                }
                catch { }
            }

            // 分支 C: 兜底情况 (光标在表头 1-5 行、点在已有箱柜下方的空白行、或游离光标)
            // 坚决在顶部汇总区最后一个有效箱柜之后紧随追加，杜绝跳行产生空行！
            if (effectiveLastSumRow >= 6)
            {
                int appendRow = effectiveLastSumRow + 1;
                if (appendRow >= minDetRow)
                {
                    return Math.Max(6, minDetRow - 1);
                }
                return appendRow;
            }

            return cabSumStartRow;
        }

        /// <summary>
        /// 核心统一执行引擎：执行箱柜完整插入、跨表克隆与定义名称系统注册
        /// 严格遵循规则 6、7、8，支持名称不变、末尾追加、策略A计费区克隆与保留CAD句柄
        /// </summary>
        /// <param name="app">Excel Application 实例</param>
        /// <param name="wb">Excel Workbook 实例</param>
        /// <param name="clip">待插入/克隆的箱柜完整快照实体</param>
        /// <param name="targetSheet">目标工作表 COM 实例</param>
        /// <param name="keepName">是否保持名称不变 (true 时不加“-副本”)</param>
        /// <param name="appendToEnd">是否严格在目标表汇总与明细末尾追加</param>
        /// <param name="preferredCursorRow">显式优先指定的光标行号 (<=0 则动态获取)</param>
        /// <returns>创建后的箱柜关键行号信息实体，失败返回 null</returns>
        public static Models.CabinetCreatedInfo? ExecuteCabinetInsertOrClone(
            dynamic app,
            dynamic wb,
            Models.CopiedCabinetContext clip,
            dynamic targetSheet,
            bool keepName = true,
            bool appendToEnd = false,
            int preferredCursorRow = 0)
        {
            if (clip == null || targetSheet == null) return null;

            // 备份并关闭屏幕渲染与事件
            bool prevUpdating = false;
            bool prevAlerts = false;
            bool prevEvents = false;
            try
            {
                prevUpdating = Convert.ToBoolean(app.ScreenUpdating);
                prevAlerts = Convert.ToBoolean(app.DisplayAlerts);
                prevEvents = Convert.ToBoolean(app.EnableEvents);
            }
            catch { }

            app.ScreenUpdating = false;
            app.DisplayAlerts = false;
            app.EnableEvents = false;

            try
            {
                // 1. 规则 8 守门：操作前强制检查并自愈目标分类表的定义名称健康状态
                Tool.FixAndFillCabinetNamesForSheet(targetSheet, forceRebuild: false);
                // 核心安全清理：剔除目标表中指向空行且无明细的幽灵残留名称 (如孤立的 Cab_Sum_6)，防止污染序号与行号
                Tool.PurgeGhostNamesForSheet(targetSheet);
                // 失效目标表箱柜映射缓存，确保结构实时同步
                Tool.InvalidateSheetCabinetsCache(Convert.ToString(targetSheet.Name) ?? "");

                // 2. 配置前缀模型
                var cfg = ConfigManager.Instance.Current?.Excel;
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = cfg?.Prefixes ?? CabinetPrefixConfig.Default;

                // 3. 规则 6: 分配目标工作表下一个全局递增唯一序号 newK
                int newK = GetNextCabinetIndex(wb, targetSheet);

                // 4. 获取源工作表句柄对象
                dynamic srcWs = null;
                try
                {
                    if (!string.IsNullOrWhiteSpace(clip.SourceSheetName))
                    {
                        srcWs = wb.Worksheets[clip.SourceSheetName];
                    }
                }
                catch { }
                if (srcWs == null) srcWs = targetSheet;
                // 判断是否在同一个工作表内操作
                bool isSameSheet = (Convert.ToString(srcWs.Name) == Convert.ToString(targetSheet.Name));

                // 核心安全守门：智能判定追加模式与光标前置插入模式
                // 1. 若调用方明确指定 appendToEnd = true (如右键二级菜单后台一键推送到其他分类表)，强制末尾追加并解绑光标；
                // 2. 若未指定 appendToEnd，且提供了有效光标行 (preferredCursorRow >= 6)，则无论是同表还是跨表复制插入，均 100% 尊重光标位置执行前置插入；
                // 3. 若未提供有效光标行且不是同表操作，则平滑回退末尾安全追加，避免将源表历史游离光标带入目标表。
                if (appendToEnd)
                {
                    // 显式指定末尾追加模式
                    preferredCursorRow = 0;
                }
                else if (!isSameSheet && preferredCursorRow < 6)
                {
                    // 跨表且未提供有效光标时，自动走末尾安全追加
                    appendToEnd = true;
                    // 清空游离光标
                    preferredCursorRow = 0;
                }

                // 5. 探测目标工作表箱柜分布并精准定位汇总行物理插入行号
                var validCabinets = Tool.GetSheetValidCabinets((object)targetSheet, (object)wb);

                // 预先嗅探用户光标所选中的目标箱柜序号 (在插行前锁定，避免插行后选区漂移)
                int targetCabinetK = 0;
                if (!appendToEnd && preferredCursorRow >= 6 && validCabinets.Count > 0)
                {
                    // 优先从有效集合中寻找汇总行等于光标行的箱柜
                    var matchedByRow = validCabinets.FirstOrDefault(c => (c.Value.Sum != null ? Convert.ToInt32(c.Value.Sum.Row) : c.Value.SumRow) == preferredCursorRow);
                    if (matchedByRow.Key > 0)
                    {
                        // 命中目标箱柜序号
                        targetCabinetK = matchedByRow.Key;
                    }
                    else
                    {
                        // 若未精确匹配汇总行，使用 FindCabinetByRow 区间匹配
                        var matchedAnc = Tool.FindCabinetByRow(validCabinets, preferredCursorRow);
                        if (matchedAnc.HasValue) targetCabinetK = matchedAnc.Value.Key;
                    }
                }

                // 核心安全准则：统一调用 DetermineTopSummaryInsertRow，支持 appendToEnd 模式直接追加至目标表末尾
                int insertSumRow = DetermineTopSummaryInsertRow(app, targetSheet, validCabinets, preferredCursorRow, appendToEnd);

                // 6. 顶部汇总表物理向下平移插入 1 行
                targetSheet.Rows[$"{insertSumRow}:{insertSumRow}"].Insert(-4121);
                // 汇总行插入后立即主动失效目标表短效缓存，确保后续明细行计算使用下移后的真实物理行号
                string curSheet = Convert.ToString(targetSheet.Name) ?? "";
                Tool.InvalidateSheetCabinetsCache(curSheet);

                // 7. 动态获取源箱柜汇总行当前最新的物理行号 (通过定义名称自适应获取)
                int realSrcSumRow = clip.SourceSumRow;
                try
                {
                    dynamic? sumName = Tool.SafeGetSheetName(srcWs, $"{sumPrefix}{clip.SourceCabinetK}");
                    if (sumName != null)
                    {
                        realSrcSumRow = Convert.ToInt32(sumName.RefersToRange.Row);
                    }
                    else if (isSameSheet && insertSumRow <= clip.SourceSumRow)
                    {
                        // 同表且插在上方时自适应加 1
                        realSrcSumRow = clip.SourceSumRow + 1;
                    }
                }
                catch { }

                // 8. 从源汇总行复制完整格式与单元格样式 (边框线、背景底色、行高)
                try
                {
                    if (realSrcSumRow > 0)
                    {
                        dynamic srcSumRange = srcWs.Rows[$"{realSrcSumRow}:{realSrcSumRow}"];
                        dynamic dstSumRange = targetSheet.Rows[$"{insertSumRow}:{insertSumRow}"];
                        // 复制源汇总行整行格式与样式
                        srcSumRange.Copy(dstSumRange);
                        // 清洗复制带入的跨表外部工作表公式前缀并自动将源箱柜定义名称 Key 替换为新 Key
                        Tool.CleanRangeFormulas(dstSumRange, clip.SourceSheetName, clip.SourceCabinetK, newK);
                    }
                }
                catch (Exception exSumCopy)
                {
                    LogHelper.WriteLog($"复制汇总行格式异常: {exSumCopy.Message}");
                }

                // 9. 绑定新汇总行定义名称 Cab_Sum_newK (统一工作表级别)
                string newSumTag = $"{sumPrefix}{newK}";
                dynamic sumAnchorCell = targetSheet.Cells[insertSumRow, 1];
                Tool.SafeSetSheetName(targetSheet, curSheet, newSumTag, insertSumRow);

                // 10. 用户确认指令：名称不变 (keepName 为 true 时保持原柜号与名称，彻底不加“-副本”)
                string newCabNo = keepName ? clip.CabinetNo : (clip.IsCut ? clip.CabinetNo : $"{clip.CabinetNo}-副本"); // --硬编码: 副本后缀--

                // A 列写入自适应动态序号公式 =ROW()-ROW(A$6)
                sumAnchorCell.Formula = "=ROW()-ROW(A$6)"; // --硬编码: 序号公式--
                // B 列到 F 列写入箱柜基本属性
                targetSheet.Cells[insertSumRow, 2].Value = newCabNo;
                targetSheet.Cells[insertSumRow, 3].Value = clip.Name;
                targetSheet.Cells[insertSumRow, 4].Value = clip.Model;
                targetSheet.Cells[insertSumRow, 5].Value = clip.Unit;
                targetSheet.Cells[insertSumRow, 6].Value = clip.Quantity;

                // 记录返回的箱柜关键信息
                var createdInfo = new Models.CabinetCreatedInfo
                {
                    CabinetK = newK,
                    SumRow = insertSumRow
                };

                // 11. 区分箱柜类型：无明细箱柜 vs 普通有明细箱柜
                if (clip.Kind == Models.CabinetKind.NoDetail || clip.DetailBlockRowCount <= 0)
                {
                    // 无明细箱柜：回写单价与成本公式
                    object rawUnitPrice = clip.SumRowValues.GetLength(1) >= 7 ? clip.SumRowValues[1, 7] : 0;
                    object rawCostPrice = clip.SumRowValues.GetLength(1) >= 10 ? clip.SumRowValues[1, 10] : 0;

                    targetSheet.Cells[insertSumRow, 7].Value = rawUnitPrice;
                    targetSheet.Cells[insertSumRow, 8].Formula = $"=F{insertSumRow}*G{insertSumRow}";
                    targetSheet.Cells[insertSumRow, 10].Value = rawCostPrice;
                    targetSheet.Cells[insertSumRow, 11].Formula = $"=F{insertSumRow}*J{insertSumRow}";
                    targetSheet.Cells[insertSumRow, 12].Formula = $"=H{insertSumRow}-K{insertSumRow}";
                    targetSheet.Cells[insertSumRow, 13].Formula = $"=IF(H{insertSumRow}=0,0,L{insertSumRow}/H{insertSumRow})";
                }
                else
                {
                    // 12. 普通有明细箱柜：定位明细块插入物理行
                    // 重新获取目标表最新的有效箱柜列表 (已包含刚才插入汇总行后的真实物理下移)
                    validCabinets = Tool.GetSheetValidCabinets((object)targetSheet, (object)wb);
                    int insertDetailRow = 41;

                    // 若为末尾追加模式或目标表现存无箱柜
                    if (appendToEnd || validCabinets.Count == 0)
                    {
                        // 重新扫描最新总计行
                        var curIndexes = Tool.FindStandardCategoryRowIndexes((object)targetSheet, -1);
                        // 校验目标表是否确实存在已有的底表明细块 (总计行与信息行均大于0)
                        if (curIndexes.cabTolsumRow > 0 && curIndexes.cabDetRow > 0)
                        {
                            // 紧随最后一个总计行后 + 4 行追加大标题
                            insertDetailRow = curIndexes.cabTolsumRow + 4;
                        }
                        else
                        {
                            // 全表尚无明细块，以配置的明细基准行 (默认41) - 3 作为首个明细大标题起始行 (第38行大标题，第41行Det)
                            int defDetRow = ConfigManager.Instance.Current?.Excel?.CabDetRowIndex ?? 41;
                            insertDetailRow = Math.Max(38, defDetRow - 3);
                        }
                    }
                    else
                    {
                        // 前置对齐插入：插入在指定目标箱柜上方 (与顶部汇总光标行上方插入保持严格对齐一致)
                        int targetDetRow = 0;
                        // 优先采用插行前已锁定的目标箱柜序号，彻底杜绝插行后 Excel 选区漂移
                        int activeCabK = targetCabinetK;
                        if (activeCabK <= 0)
                        {
                            // 若插行前未锁定，回退使用当前光标实时探测
                            KeyValuePair<int, Models.CabinetAnchorModel>? activeCab = Tool.GetActiveCabinet((object)app, validCabinets, fallbackSingle: true);
                            if (activeCab != null) activeCabK = activeCab.Value.Key;
                        }

                        if (activeCabK > 0)
                        {
                            // 锁定目标箱柜在最新有效集合中的索引
                            int activeIdx = -1;
                            for (int i = 0; i < validCabinets.Count; i++)
                            {
                                if (validCabinets[i].Key == activeCabK) { activeIdx = i; break; }
                            }
                            // 从目标箱柜向后寻找首个拥有明细的箱柜 (优先命中目标箱柜自身)
                            for (int i = (activeIdx >= 0 ? activeIdx : 0); i < validCabinets.Count; i++)
                            {
                                var anc = validCabinets[i].Value;
                                if (anc != null)
                                {
                                    if (anc.Det != null)
                                    {
                                        try { targetDetRow = Convert.ToInt32(anc.Det.Row); } catch { }
                                    }
                                    if (targetDetRow <= 0 && anc.DetRow > 0)
                                    {
                                        targetDetRow = anc.DetRow;
                                    }
                                    if (targetDetRow > 0) break;
                                }
                            }
                        }

                        if (targetDetRow > 0)
                        {
                            // 插在明细大标题行 (Det - 3) 上方，确保顺序一致 --硬编码: 大标题在Det前3行--
                            insertDetailRow = Math.Max(38, targetDetRow - 3);
                        }
                        else
                        {
                            // 全表无明细或无法定位前置柜时，回退末尾追加
                            var curIndexes = Tool.FindStandardCategoryRowIndexes((object)targetSheet, -1);
                            if (curIndexes.cabTolsumRow > 0 && curIndexes.cabDetRow > 0)
                            {
                                insertDetailRow = curIndexes.cabTolsumRow + 4;
                            }
                            else
                            {
                                int defDetRow = ConfigManager.Instance.Current?.Excel?.CabDetRowIndex ?? 41;
                                insertDetailRow = Math.Max(38, defDetRow - 3);
                            }
                        }
                    }

                    // 13. 动态获取源箱柜明细块关键行号
                    int realDetRow = 0;
                    int realSubsumRow = 0;
                    int realTolsumRow = 0;
                    try
                    {
                        dynamic? detName = Tool.SafeGetSheetName(srcWs, $"{detPrefix}{clip.SourceCabinetK}");
                        if (detName != null) realDetRow = Convert.ToInt32(detName.RefersToRange.Row);
                        dynamic? subsumName = Tool.SafeGetSheetName(srcWs, $"{subsumPrefix}{clip.SourceCabinetK}");
                        if (subsumName != null) realSubsumRow = Convert.ToInt32(subsumName.RefersToRange.Row);
                        dynamic? tolsumName = Tool.SafeGetSheetName(srcWs, $"{tolsumPrefix}{clip.SourceCabinetK}");
                        if (tolsumName != null) realTolsumRow = Convert.ToInt32(tolsumName.RefersToRange.Row);
                    }
                    catch { }

                    // 兜底计算偏移 (若定义名称未解析到，优先使用快照中的真实物理行与明细区间)
                    if (realDetRow <= 0)
                    {
                        int sumOffset = (isSameSheet && insertSumRow <= clip.SourceDetStartRow) ? 1 : 0;
                        realDetRow = clip.SourceDetRow + sumOffset;
                        realSubsumRow = clip.SourceSubsumRow + sumOffset;
                        realTolsumRow = clip.SourceTolsumRow + sumOffset;
                        if (realDetRow <= 0 && clip.SourceDetStartRow > 0)
                        {
                            realDetRow = clip.SourceDetStartRow + 3;
                            realSubsumRow = clip.SourceDetEndRow - 8;
                            realTolsumRow = clip.SourceDetEndRow - 3;
                        }
                    }

                    // 推导源明细块起止行范围 (通常 detRow - 3 到 tolsumRow + 3)
                    int realDetStart = realDetRow - 3;
                    if (realDetStart < 1) realDetStart = realDetRow;
                    int realDetEnd = realTolsumRow + 3;
                    int blockRowCount = realDetEnd - realDetStart + 1;

                    // 14. 物理插入明细块空白行
                    targetSheet.Rows[$"{insertDetailRow}:{insertDetailRow + blockRowCount - 1}"].Insert(-4121);

                    // 同表偏移处理
                    if (isSameSheet && insertDetailRow <= realDetStart)
                    {
                        realDetStart += blockRowCount;
                        realDetEnd += blockRowCount;
                        realDetRow += blockRowCount;
                        realSubsumRow += blockRowCount;
                        realTolsumRow += blockRowCount;
                    }

                    // 15. 原生 Range.Copy 完整复刻边框、底色、合并单元格、行高，并自动保留 CAD 句柄
                    dynamic dstDetailRange = targetSheet.Rows[$"{insertDetailRow}:{insertDetailRow + blockRowCount - 1}"];
                    bool copySucceeded = false;
                    try
                    {
                        if (realDetStart > 0 && realDetEnd >= realDetStart)
                        {
                            dynamic srcDetailRange = srcWs.Rows[$"{realDetStart}:{realDetEnd}"];
                            // 全量克隆明细块 (策略 A: 完整原样克隆计费区与元器件；保留句柄)
                            srcDetailRange.Copy(dstDetailRange);
                            copySucceeded = true;
                        }
                    }
                    catch (Exception exCopy)
                    {
                        LogHelper.WriteLog($"Range.Copy 明细块异常: {exCopy.Message}");
                    }

                    // 降级兜底数组写入
                    if (!copySucceeded)
                    {
                        dynamic detRange = targetSheet.Range[$"A{insertDetailRow}:Q{insertDetailRow + blockRowCount - 1}"];
                        detRange.Value2 = clip.DetailBlockValues;
                        detRange.Formula = clip.DetailBlockFormulas;
                        try
                        {
                            detRange.Borders.LineStyle = 1;
                            detRange.Borders.Weight = 2;
                        }
                        catch { }
                        // 清洗数组写入公式中残留的跨表前缀与旧箱柜定义名称 Key
                        Tool.CleanRangeFormulas(detRange, clip.SourceSheetName, clip.SourceCabinetK, newK);
                    }
                    else
                    {
                        // 清洗复制产生的跨表外部引用公式，恢复本地自适应，剥离源工作表限定名并替换旧箱柜 Key
                        Tool.CleanRangeFormulas(dstDetailRange, clip.SourceSheetName, clip.SourceCabinetK, newK);
                    }

                    // 16. 计算新明细块关键行号
                    int newDetRow = insertDetailRow + (realDetRow - realDetStart);
                    int newSubsumRow = insertDetailRow + (realSubsumRow - realDetStart);
                    int newTolsumRow = insertDetailRow + (realTolsumRow - realDetStart);

                    string newDetTag = $"{detPrefix}{newK}";
                    string newSubsumTag = $"{subsumPrefix}{newK}";
                    string newTolsumTag = $"{tolsumPrefix}{newK}";

                    // 17. 注册目标表 3 个明细定义名称 (规则 6 规范)
                    Tool.SafeSetSheetName(targetSheet, curSheet, newDetTag, newDetRow);
                    Tool.SafeSetSheetName(targetSheet, curSheet, newSubsumTag, newSubsumRow);
                    Tool.SafeSetSheetName(targetSheet, curSheet, newTolsumTag, newTolsumRow);

                    // 18. 重建双向跳转超链接
                    try
                    {
                        targetSheet.Hyperlinks.Add(
                            Anchor: sumAnchorCell,
                            Address: "",
                            SubAddress: $"'{curSheet}'!{newDetTag}",
                            ScreenTip: "点击进入本箱柜明细表" // --硬编码: 屏幕提示文本--
                        );
                        sumAnchorCell.Formula = "=ROW()-ROW(A$6)"; // --硬编码: 序号公式--
                        targetSheet.Hyperlinks.Add(
                            Anchor: targetSheet.Cells[newDetRow, 1],
                            Address: "",
                            SubAddress: $"'{curSheet}'!{newSumTag}",
                            ScreenTip: "返回汇总行" // --硬编码: 屏幕提示文本--
                        );
                    }
                    catch { }

                    // 19. 更新明细表头柜号文本 (名称保持不变)
                    targetSheet.Cells[newDetRow, 2].Value = newCabNo;

                    // 19.5 核心关键修复：显式将明细表头行 (detRow + 1) A 列动态自适应绑定到新汇总行定义名称
                    // 彻底解决跨分类复制或插入箱柜时表头公式残留旧 Key 并报 #NAME? 的缺陷 (规则 6 架构规范)
                    try
                    {
                        // 明细表头行 (newDetRow + 1) A 列写入动态序号公式，对齐标准模板与规则 6
                        targetSheet.Cells[newDetRow + 1, 1].Formula = $"=\"序号\" & {newSumTag}"; // --硬编码: 明细表头序号公式--
                    }
                    catch (Exception exHdrFormula)
                    {
                        // 记录表头公式重写容错日志
                        LogHelper.WriteLog($"绑定新箱柜明细表头序号公式异常: {exHdrFormula.Message}");
                    }

                    // 20. 刷新明细块元器件 A 列序号、小计求和公式与计费区域公式
                    int compStartRow = newDetRow + 2;
                    RefreshCabinetFeeAreaFormulas(targetSheet, newDetRow, compStartRow, newSubsumRow, newTolsumRow);

                    // 21. 汇总行单价总价与成本公式联动绑定 (严谨对齐标准分类表列定义，消除 #REF!)
                    try
                    {
                        // 在 tolsum 总计行 F 列自动绑定联动公式指向汇总行 F 列 (保证台数同步)
                        targetSheet.Cells[newTolsumRow, 6].Formula = $"=F{insertSumRow}";
                        // G 列单价公式指向明细总计行单台单价 G 列
                        targetSheet.Cells[insertSumRow, 7].Formula = $"=G{newTolsumRow}";
                        // H 列总价公式 = 数量(F列) * 单价(G列)
                        targetSheet.Cells[insertSumRow, 8].Formula = $"=F{insertSumRow}*G{insertSumRow}";
                        // J 列成本总价公式指向明细总计行的成本总价 (K 列)
                        targetSheet.Cells[insertSumRow, 10].Formula = $"=K{newTolsumRow}";
                        // K 列毛利公式 = 总价(H列) - 成本总价(J列)
                        targetSheet.Cells[insertSumRow, 11].Formula = $"=H{insertSumRow}-J{insertSumRow}";
                        // L 列毛利率公式
                        targetSheet.Cells[insertSumRow, 12].Formula = $"=IF(H{insertSumRow}=0,0,K{insertSumRow}/H{insertSumRow})";
                    }
                    catch (Exception exFormula)
                    {
                        // 记录公式绑定异常日志
                        LogHelper.WriteLog($"绑定新箱柜汇总行联动公式异常: {exFormula.Message}");
                    }

                    // 安全自愈兜底：若因跨表复制导致 G 列显示为 #REF!，强制再次恢复
                    try
                    {
                        // 提取当前单元格文本
                        string gTxt = Convert.ToString(targetSheet.Cells[insertSumRow, 7].Text)?.Trim() ?? "";
                        // 判定是否含有坏引用
                        if (gTxt.Contains("#REF!"))
                        {
                            // 强制重写单价公式指向总计行
                            targetSheet.Cells[insertSumRow, 7].Formula = $"=G{newTolsumRow}";
                        }
                    }
                    catch { }

                    // 填充创建实体明细行号
                    createdInfo.DetRow = newDetRow;
                    createdInfo.SubsumRow = newSubsumRow;
                    createdInfo.TolsumRow = newTolsumRow;
                }

                // 21.5 跨表复制清理：若非同表操作且源序号与新序号不同，清理目标表中被 Range.Copy 附带复制进来的源表旧定义名称
                if (!isSameSheet && clip.SourceCabinetK > 0 && clip.SourceCabinetK != newK)
                {
                    // 安全移除目标表中残留的源表箱柜旧定义名称，防止名称冲突和序号串扰
                    Tool.SafeDeleteName(targetSheet, wb, $"{sumPrefix}{clip.SourceCabinetK}");
                    // 安全删除明细信息行定义名称
                    Tool.SafeDeleteName(targetSheet, wb, $"{detPrefix}{clip.SourceCabinetK}");
                    // 安全删除小计行定义名称
                    Tool.SafeDeleteName(targetSheet, wb, $"{subsumPrefix}{clip.SourceCabinetK}");
                    // 安全删除总计行定义名称
                    Tool.SafeDeleteName(targetSheet, wb, $"{tolsumPrefix}{clip.SourceCabinetK}");
                }

                // 核心强化：彻底清理新插入位置上除 newK 之外任何被 Range.Copy 附带带入的同位置冗余旧定义名称
                Tool.PurgeDuplicateNamesAtRow(targetSheet, wb, insertSumRow, createdInfo.DetRow, createdInfo.SubsumRow, createdInfo.TolsumRow, newK);
                // 核心自愈闭环：执行一次全表幽灵名称与同位置重复定义名称安全自愈净化
                Tool.PurgeGhostNamesForSheet(targetSheet);

                // 21.6 终极自愈保险：确保新创建的箱柜 newK 的整套定义名称 100% 存在且指向正确行号 (规则 6 架构保障)
                try
                {
                    if (insertSumRow > 0)
                    {
                        // 校验并回写汇总行定义名称
                        Tool.SafeSetSheetName(targetSheet, curSheet, newSumTag, insertSumRow);
                    }
                    if (createdInfo.DetRow > 0)
                    {
                        // 校验并回写明细信息行定义名称
                        Tool.SafeSetSheetName(targetSheet, curSheet, $"{detPrefix}{newK}", createdInfo.DetRow);
                    }
                    if (createdInfo.SubsumRow > 0)
                    {
                        // 校验并回写小计行定义名称
                        Tool.SafeSetSheetName(targetSheet, curSheet, $"{subsumPrefix}{newK}", createdInfo.SubsumRow);
                    }
                    if (createdInfo.TolsumRow > 0)
                    {
                        // 校验并回写总计行定义名称
                        Tool.SafeSetSheetName(targetSheet, curSheet, $"{tolsumPrefix}{newK}", createdInfo.TolsumRow);
                    }
                }
                catch { }

                // 22. 若为剪切模式，删除源工作表中的原始箱柜
                if (clip.IsCut && !string.IsNullOrWhiteSpace(clip.SourceSheetName))
                {
                    try
                    {
                        srcWs = wb.Worksheets[clip.SourceSheetName];
                        if (srcWs != null)
                        {
                            DeleteCabinets(app, srcWs, new List<int> { clip.SourceCabinetK });
                        }
                    }
                    catch (Exception exDel)
                    {
                        LogHelper.WriteLog($"剪切后删除原箱柜异常: {exDel.Message}");
                    }
                    _copiedCabinetContext = null;
                }

                // 主动失效目标工作表箱柜映射缓存，确保结构实时同步
                Tool.InvalidateSheetCabinetsCache(curSheet);

                return createdInfo;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[ExecuteCabinetInsertOrClone] 异常: {ex.Message}");
                return null;
            }
            finally
            {
                // 恢复屏幕渲染与系统事件
                try
                {
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }
                catch { }
            }
        }

        /// <summary>
        /// 供 Ribbon 菜单调用的“插入复制的箱柜”入口
        /// 读取剪贴板整块数据，在目标行安全插入并分配新全局序号 newK，若是剪切则安全擦除原箱柜
        /// 统一复用 ExecuteCabinetInsertOrClone 引擎，实现名称不变与末尾安全追加
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

                // 获取当前光标所在物理行号 (用户指令核心：需要在顶部汇总的光标行上方插入)
                int cursorRow = 0;
                try
                {
                    dynamic activeCell = app?.ActiveCell;
                    if (activeCell != null) cursorRow = Convert.ToInt32(activeCell.Row);
                }
                catch { }

                // 调度核心引擎 (keepName = true 保持名称不变，preferredCursorRow: cursorRow 在光标行上方插入)
                var result = ExecuteCabinetInsertOrClone(
                    app,
                    wb,
                    _copiedCabinetContext,
                    activeSheet,
                    keepName: true,
                    appendToEnd: false,
                    preferredCursorRow: cursorRow);
                if (result != null)
                {
                    // 激活新插入的箱柜
                    activeSheet.Activate();
                    if (result.DetRow > 0)
                    {
                        activeSheet.Cells[result.DetRow, 2].Select();
                    }
                    else if (result.SumRow > 0)
                    {
                        activeSheet.Cells[result.SumRow, 2].Select();
                    }
                    System.Windows.Forms.MessageBox.Show($"箱柜【{_copiedCabinetContext.CabinetNo}】插入成功！", "操作成功", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
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

                // 规则 8: 在操作 Excel 提取前执行 FixAndFillCabinetNamesForSheet 确保规则 6 架构与定义名称正确
                Tool.FixAndFillCabinetNamesForSheet(activeSheet, forceRebuild: false);
                // 失效当前工作表箱柜短效缓存，强制获取最新拓扑
                string sheetName = Convert.ToString(activeSheet.Name) ?? "";
                Tool.InvalidateSheetCabinetsCache(sheetName);

                // 扫描获取目标工作表所有有效箱柜
                var validCabinets = Tool.GetSheetValidCabinets(activeSheet, wb);
                int order = 1;
                // 初始化已访问汇总行集合以防重复映射
                var seenSumRows = new HashSet<int>();
                // 初始化已访问箱柜索引集合
                var seenIndices = new HashSet<int>();

                foreach (var cab in validCabinets)
                {
                    int k = cab.Key;
                    var anchor = cab.Value;
                    // 提取汇总行物理行号
                    int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : (anchor.SumRow > 0 ? anchor.SumRow : 0);
                    // 规则 6 强约束：顶部必须为汇总表，汇总行必须大于等于 6，若无有效汇总行则绝非合法箱柜，坚决剔除
                    if (sumRow < 6) continue;
                    // 杜绝重复物理行重复添加 (避免一个汇总行被多个损坏名称重复映射)
                    if (seenSumRows.Contains(sumRow) || seenIndices.Contains(k)) continue;

                    // 提取明细行物理行号判定是否为无明细箱柜
                    int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : (anchor.DetRow > 0 ? anchor.DetRow : 0);
                    bool isNoDet = (detRow <= 0);

                    // 读取汇总行中的柜号、名称与型号
                    string cabNo = Convert.ToString(activeSheet.Cells[sumRow, 2].Value)?.Trim() ?? "";
                    string name = Convert.ToString(activeSheet.Cells[sumRow, 3].Value)?.Trim() ?? "";
                    string model = Convert.ToString(activeSheet.Cells[sumRow, 4].Value)?.Trim() ?? "";

                    // 若柜号、名称、型号均为空，说明属于空白行或游离名称，跳过
                    if (string.IsNullOrEmpty(cabNo) && string.IsNullOrEmpty(name) && string.IsNullOrEmpty(model)) continue;

                    // 登记已收录的汇总行与箱柜索引
                    seenSumRows.Add(sumRow);
                    seenIndices.Add(k);

                    // 智能补齐默认柜号
                    if (string.IsNullOrEmpty(cabNo)) cabNo = !string.IsNullOrEmpty(name) ? name : $"箱柜{k}";

                    // 读取汇总行第 6 列数量
                    double qty = 1.0;
                    try { qty = Convert.ToDouble(activeSheet.Cells[sumRow, 6].Value); } catch { }

                    // 添加至调序输出传输模型列表
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
        /// 应用箱柜调序（ExWinner 核心内存整块物理重排算法）
        /// 完整双向联动：同步按用户指定新顺序对顶部汇总表与底部分类明细表进行零公式撕裂重排
        /// </summary>
        /// <param name="newOrderKList">用户排列好的箱柜序号顺序列表</param>
        /// <returns>调序是否成功</returns>
        public static bool ApplyCabinetReorder(List<int> newOrderKList)
        {
            // 校验输入序号列表有效性
            if (newOrderKList == null || newOrderKList.Count <= 1) return true;

            try
            {
                // 获取当前活动 Excel 运行环境上下文
                var context = Tool.GetActiveExcelContext();
                if (context == null) return false;
                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic activeSheet = context.Sheet;

                // 规则 8: 调序前调用 FixAndFillCabinetNamesForSheet 自愈并校准定义名称
                Tool.FixAndFillCabinetNamesForSheet(activeSheet, forceRebuild: false);
                string sheetName = Convert.ToString(activeSheet.Name) ?? "";
                // 刷新内存短效缓存
                Tool.InvalidateSheetCabinetsCache(sheetName);

                // 读取箱柜定义名称前缀配置对象
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

                // 扫描获取当前工作表中所有已识别的有效箱柜列表 (显式强类型接收，断开 dynamic 推导)
                List<KeyValuePair<int, Models.CabinetAnchorModel>> rawCabinets = Tool.GetSheetValidCabinets((object)activeSheet, (object)wb);
                // 规则 6 强约束：仅纳管存在有效汇总行 (SumRow >= 6) 的合规箱柜
                List<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets = rawCabinets
                    .Where(c => (c.Value.Sum != null ? Convert.ToInt32(c.Value.Sum.Row) : c.Value.SumRow) >= 6)
                    .ToList();
                if (validCabinets.Count <= 1) return true;

                // 提取当前物理排列顺序并与用户指定顺序对比
                var currentOrder = validCabinets.Select(c => c.Key).ToList();
                // 若新旧顺序完全一致则无需任何物理行变动
                if (currentOrder.SequenceEqual(newOrderKList)) return true;

                // 构造有效箱柜序号哈希集合进行存在性过滤
                var validKSet = new HashSet<int>(validCabinets.Select(c => c.Key));
                var sanitizedOrder = newOrderKList.Where(k => validKSet.Contains(k)).ToList();
                // 容错补偿：补充可能未在传入列表中声明的已有箱柜序号
                foreach (var k in currentOrder)
                {
                    // 若列表中缺少当前箱柜则追加在末尾
                    if (!sanitizedOrder.Contains(k)) sanitizedOrder.Add(k);
                }
                // 校验清洗后的有效排序集合数量
                if (sanitizedOrder.Count <= 1) return true;

                // 预先搜集所有拥有明细块的箱柜及其明细块物理总行数
                var detCabinets = validCabinets.Where(c => c.Value.Det != null).OrderBy(c => (int)c.Value.Det.Row).ToList();
                var detRowCountMap = new Dictionary<int, int>();
                int firstDetStartRow = 0;

                // 若存在拥有明细块的箱柜则计算明细区域参数
                if (detCabinets.Count > 0)
                {
                    // 获取全表首个明细块的起始物理行号 (Det行减3)
                    int firstDetRow = Convert.ToInt32(detCabinets[0].Value.Det.Row);
                    firstDetStartRow = firstDetRow - 3;
                    // 容错校验起始行必须大于等于1
                    if (firstDetStartRow < 1) firstDetStartRow = firstDetRow;

                    // 遍历所有有明细箱柜计算各自明细块完整物理行数
                    for (int i = 0; i < detCabinets.Count; i++)
                    {
                        var cab = detCabinets[i];
                        int k = cab.Key;
                        // 获取当前明细信息行与总计行物理行号
                        int curDet = Convert.ToInt32(cab.Value.Det.Row);
                        int curTol = cab.Value.Tolsum != null ? Convert.ToInt32(cab.Value.Tolsum.Row) : (curDet + 27);
                        // 计算起始行 (大标题行)
                        int curStart = curDet - 3;
                        if (curStart < 1) curStart = curDet;

                        // 计算结束行 (总计行及附注落款共3行)
                        int curEnd = curTol + 3;
                        // 若存在紧随其后的下一个明细块且中间有空行则平滑包含空行
                        if (i < detCabinets.Count - 1)
                        {
                            // 获取下一个明细块的起始行号
                            int nextDet = Convert.ToInt32(detCabinets[i + 1].Value.Det.Row);
                            int nextStart = nextDet - 3;
                            // 若下一个起始行大于当前结束行加1则将间隙包含在当前块内
                            if (nextStart > curEnd + 1)
                            {
                                curEnd = nextStart - 1;
                            }
                        }

                        // 计算并登记当前箱柜明细块总行数
                        int rowCount = curEnd - curStart + 1;
                        detRowCountMap[k] = rowCount;
                    }
                }

                // 记录首个汇总行当前的物理行号锚点
                int firstSumRow = Convert.ToInt32(validCabinets[0].Value.Sum.Row);

                // 关闭屏幕更新、弹窗拦截与事件循环以提升执行性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                try
                {
                    // ==========================================
                    // 阶段一：顶部汇总行物理整行剪切插入重排
                    // ==========================================
                    int targetSumRow = firstSumRow;
                    foreach (int k in sanitizedOrder)
                    {
                        // 实时动态获取箱柜 k 当前的最新汇总行物理行号 (定义名称自适应平移)
                        dynamic? sumName = Tool.SafeGetSheetName(activeSheet, $"{sumPrefix}{k}");
                        int curSumRow = sumName != null ? Convert.ToInt32(sumName.RefersToRange.Row) : 0;
                        if (curSumRow <= 0) continue;

                        // 若当前行号与目标行号不一致则执行原生整行剪切插入
                        if (curSumRow != targetSumRow)
                        {
                            // 提取源行与目标行范围
                            dynamic cutRow = activeSheet.Rows[$"{curSumRow}:{curSumRow}"];
                            dynamic targetRow = activeSheet.Rows[$"{targetSumRow}:{targetSumRow}"];
                            // 原生 Cut 紧随 Insert 插入剪切单元格 (-4121 对应 xlShiftDown)
                            cutRow.Cut();
                            targetRow.Insert(-4121);
                        }

                        // 目标汇总行指针单步下移
                        targetSumRow++;
                    }

                    // ==========================================
                    // 阶段二：底部分类明细块物理整块剪切插入重排
                    // ==========================================
                    if (detCabinets.Count > 0 && firstDetStartRow > 0)
                    {
                        // 筛选新排序列表中拥有明细块的箱柜序列
                        var orderedDetKs = sanitizedOrder.Where(k => detRowCountMap.ContainsKey(k)).ToList();
                        int targetDetStartRow = firstDetStartRow;

                        foreach (int k in orderedDetKs)
                        {
                            // 提取预先锁定的该箱柜明细块物理行数
                            if (!detRowCountMap.TryGetValue(k, out int blockLen) || blockLen <= 0) continue;

                            // 实时动态获取箱柜 k 当前最新的明细信息行行号
                            dynamic? detName = Tool.SafeGetSheetName(activeSheet, $"{detPrefix}{k}");
                            int curDetRow = detName != null ? Convert.ToInt32(detName.RefersToRange.Row) : 0;
                            if (curDetRow <= 0) continue;

                            // 推导当前明细块起始行与结束行范围
                            int curStart = curDetRow - 3;
                            if (curStart < 1) curStart = curDetRow;
                            int curEnd = curStart + blockLen - 1;

                            // 若当前起始行与目标起始行不一致则执行整块原生剪切插入
                            if (curStart != targetDetStartRow)
                            {
                                // 提取源明细块多行范围与目标行
                                dynamic cutBlock = activeSheet.Rows[$"{curStart}:{curEnd}"];
                                dynamic targetBlock = activeSheet.Rows[$"{targetDetStartRow}:{targetDetStartRow}"];
                                // 原生 Cut 紧随 Insert 插入整块明细行 (-4121 对应 xlShiftDown)
                                cutBlock.Cut();
                                targetBlock.Insert(-4121);
                            }

                            // 目标明细起始行指针累加当前块总行数
                            targetDetStartRow += blockLen;
                        }
                    }

                    // ==========================================
                    // 阶段三：全局校准自愈与公式超链接双向联动刷新
                    // ==========================================
                    string curSheetName = Convert.ToString(activeSheet.Name) ?? "";
                    // 重新全量扫描重排后的最新有效箱柜物理映射 (显式强类型接收，断开 dynamic 推导)
                    List<KeyValuePair<int, Models.CabinetAnchorModel>> refreshedCabinets = Tool.GetSheetValidCabinets((object)activeSheet, (object)wb);

                    foreach (var cab in refreshedCabinets)
                    {
                        int k = cab.Key;
                        var anc = cab.Value;
                        // 读取各锚点当前最新的绝对物理行号
                        int sRow = anc.Sum != null ? Convert.ToInt32(anc.Sum.Row) : 0;
                        int dRow = anc.Det != null ? Convert.ToInt32(anc.Det.Row) : 0;
                        int subRow = anc.Subsum != null ? Convert.ToInt32(anc.Subsum.Row) : 0;
                        int tolRow = anc.Tolsum != null ? Convert.ToInt32(anc.Tolsum.Row) : 0;

                        if (sRow > 0)
                        {
                            // 1. 刷新汇总行 A 列自适应动态序号公式
                            activeSheet.Cells[sRow, 1].Formula = "=ROW()-ROW(A$6)"; // --硬编码: 公式表达式--

                            // 2. 普通有明细箱柜联动公式与超链接校准
                            if (dRow > 0 && tolRow > 0)
                            {
                                // 汇总行 G 列单价指向明细总计行单台单价 G 列 (方式 B 稳健绑定，避免数量二次相乘)
                                activeSheet.Cells[sRow, 7].Formula = $"=G{tolRow}";
                                // 汇总行 H 列销售合价 = 数量 * 单价
                                activeSheet.Cells[sRow, 8].Formula = $"=F{sRow}*G{sRow}";
                                // 汇总行 J 列成本单价指向明细总计行成本总价 K 列
                                activeSheet.Cells[sRow, 10].Formula = $"=K{tolRow}";
                                // 汇总行 K 列毛利 = 销售合价 - 成本总价
                                activeSheet.Cells[sRow, 11].Formula = $"=H{sRow}-J{sRow}";
                                // 汇总行 L 列毛利率计算公式
                                activeSheet.Cells[sRow, 12].Formula = $"=IF(H{sRow}=0,0,K{sRow}/H{sRow})";

                                // 重新绑定汇总行至明细行的超链接 (保护 A 列公式)
                                try
                                {
                                    dynamic sumAnchorCell = activeSheet.Cells[sRow, 1];
                                    string detTag = $"{detPrefix}{k}";
                                    // 汇总行 A 列设置跳转明细表头超链接
                                    activeSheet.Hyperlinks.Add(
                                        Anchor: sumAnchorCell,
                                        Address: "",
                                        SubAddress: $"'{curSheetName}'!{detTag}",
                                        ScreenTip: "点击进入本箱柜明细表" // --硬编码: 屏幕提示文本--
                                    );
                                    // 恢复 A 列为自适应序号公式
                                    sumAnchorCell.Formula = "=ROW()-ROW(A$6)"; // --硬编码: 公式表达式--

                                    // 明细行 A 列设置返回顶部汇总行超链接
                                    dynamic detAnchorCell = activeSheet.Cells[dRow, 1];
                                    string sumTag = $"{sumPrefix}{k}";
                                    activeSheet.Hyperlinks.Add(
                                        Anchor: detAnchorCell,
                                        Address: "",
                                        SubAddress: $"'{curSheetName}'!{sumTag}",
                                        ScreenTip: "返回汇总行"
                                    );
                                }
                                catch { }

                                // 确保明细行 B 列无超链接 (遵循规则 6 架构规范)
                                try { activeSheet.Cells[dRow, 2].Hyperlinks.Delete(); } catch { }

                                // 刷新明细块元器件区域序号公式、小计公式及计费公式
                                int compStart = dRow + 2;
                                if (subRow > compStart && tolRow >= subRow)
                                {
                                    // 调度批量公式自愈刷新服务 (内存二维数组批量写回)
                                    RefreshCabinetFeeAreaFormulas(activeSheet, dRow, compStart, subRow, tolRow);
                                }
                            }
                        }
                    }

                    // 触发当前活动工作表公式链全量重新计算
                    activeSheet.Calculate();

                    // 重新激活工作表并聚焦首个箱柜汇总行
                    activeSheet.Activate();
                    if (firstSumRow > 0)
                    {
                        // 选中首行柜号单元格
                        activeSheet.Cells[firstSumRow, 2].Select();
                    }
                }
                finally
                {
                    // 恢复 Excel 屏幕更新、弹窗告警与事件响应
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }

                // 调序成功返回 true
                return true;
            }
            catch (Exception ex)
            {
                // 记录调序异常日志
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
