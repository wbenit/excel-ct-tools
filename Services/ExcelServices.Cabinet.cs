using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：新建箱柜与对象模型渲染
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 从模板 CabinetTemplate.xlsx 中直接复制第 41 行至 74 行（完整箱柜明细块），并在光标位置插入顶部汇总行
        /// 完整完成：顶部汇总行插入、公式联动、4 个定义名称注册及双向超链接绑定
        /// 根据规则与要求：不执行公式外部链接清洗
        /// </summary>
        /// <param name="targetSheet">目标工作表 COM 实例（为空时自动获取当前活动工作表）</param>
        /// <param name="targetStartRow">目标明细插入起始物理行号（大于0时生效，否则自动根据光标/末尾推导）</param>
        /// <param name="cabinetK">新箱柜序号 K（<=0 时自动计算下一个独立序号）</param>
        /// <param name="initialCabName">可选传入的初始箱柜名称（为空时默认为“箱柜K”）</param>
        /// <param name="explicitApp">可选传入的 Excel Application COM 实例</param>
        /// <returns>新建/复制成功的箱柜行号信息对象（CabinetCreatedInfo），失败返回 null</returns>
        public static Models.CabinetCreatedInfo? CopyCabinetDetailFromTemplate(
            dynamic? targetSheet = null,
            int targetStartRow = 0,
            int cabinetK = 0,
            string? initialCabName = null,
            dynamic? explicitApp = null)
        {
            try
            {
                // 1. 获取当前 Excel 运行环境与上下文
                var context = Tool.GetActiveExcelContext(explicitApp, targetSheet);
                // 校验运行上下文有效性
                if (context == null) return null;
                // 提取 Excel COM 核心对象
                dynamic app = context.App;
                // 提取活动工作簿对象
                dynamic wb = context.Wb;
                // 提取目标工作表对象
                dynamic activeSheet = targetSheet ?? context.Sheet;

                // 2. 检查或生成模板文件物理路径 (复用 ProjectController.EnsureCabinetTemplate)
                string templatePath = Controllers.ProjectController.EnsureCabinetTemplate(app);
                // 校验模板物理文件是否存在
                if (!System.IO.File.Exists(templatePath))
                {
                    // 记录未找到模板文件日志
                    LogHelper.WriteLog($"未找到模板文件: {templatePath}");
                    // 模板缺失返回失败
                    return null;
                }

                // 3. 读取配置参数与前缀
                var cfg = ConfigManager.Instance.Current.Excel;
                // 读取箱柜定义名称前缀值对象 (零堆分配)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = cfg.Prefixes;
                // 默认模板工作表名称
                string defaultTemplateSheet = cfg.DefaultTemplateSheet ?? "分类1";

                // 关闭屏幕刷新与系统事件以提升操作性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                try
                {
                    // 4. 扫描当前工作表有效箱柜映射与末尾分布
                    var validCabinets = Tool.GetSheetValidCabinets(activeSheet, wb);
                    // 探测工作表基准/末尾行号分布
                    var lastIndexes = Tool.FindStandardCategoryRowIndexes((object)activeSheet, -1);

                    // 5. 动态计算箱柜序号 K (若未显式传入)
                    if (cabinetK <= 0)
                    {
                        // 动态计算下一个可用独立序号
                        cabinetK = GetNextCabinetIndex(wb, activeSheet);
                    }

                    // 6. 智能识别当前光标命中的箱柜实体
                    var activeCab = Tool.GetActiveCabinet(app, validCabinets, fallbackSingle: true);
                    // 提取源箱柜锚点 (优先光标选中，回退最后一个有效箱柜)
                    var srcCabAnchor = activeCab?.Value ?? (validCabinets.Count > 0 ? validCabinets[validCabinets.Count - 1].Value : null);

                    // 7. 定位顶部汇总行目标插入位置 (紧随选中箱柜汇总行之后，未命中取基准汇总行下一行)
                    int insertSumRow = 0;
                    if (srcCabAnchor?.Sum != null)
                    {
                        // 在光标命中的箱柜汇总行下方插入
                        insertSumRow = Convert.ToInt32(srcCabAnchor.Sum.Row) + 1;
                    }
                    else
                    {
                        // 使用汇总行最大值 +1
                        insertSumRow = lastIndexes.cabSumRow + 1;
                    }

                    // 检查汇总目标行 A 列与 B 列是否包含内容，防止覆盖已有行
                    string checkSumA = Convert.ToString(activeSheet.Cells[insertSumRow, 2].Value)?.Trim() ?? "";
                    string checkSumB = Convert.ToString(activeSheet.Cells[insertSumRow, 3].Value)?.Trim() ?? "";
                    // 判定是否需要物理向下插入汇总行
                    bool needInsertSumRow = !string.IsNullOrWhiteSpace(checkSumA) || !string.IsNullOrWhiteSpace(checkSumB);
                    // 记录是否实际执行了汇总插行
                    if (needInsertSumRow)
                    {
                        // 物理插入 1 行 (-4121 对应 xlShiftDown)
                        activeSheet.Rows[$"{insertSumRow}:{insertSumRow}"].Insert(-4121);
                    }

                    // 8. 定位底部明细块目标起始插入行
                    int copyRowCount = 74 - 41 + 1; // 模板固定 41 行至 74 行，共 34 行
                    int targetDetailStartRow = targetStartRow;
                    // 若未显式传入明细起始行，则智能计算
                    if (targetDetailStartRow <= 0)
                    {
                        // 若光标命中了源箱柜，紧随源箱柜明细块（总计行+3行报价人）之后插入
                        if (srcCabAnchor?.Tolsum != null)
                        {
                            // 紧随源箱柜报价人信息之后
                            targetDetailStartRow = Convert.ToInt32(srcCabAnchor.Tolsum.Row) + 4;
                        }
                        else
                        {
                            // 回退至当前表末尾明细块之后或基准 41 行
                            targetDetailStartRow = lastIndexes.cabTolsumRow > 0 ? lastIndexes.cabTolsumRow + 4 : 41;
                        }
                    }

                    // 9. 物理向下插入 34 行空间 (-4121 对应 xlShiftDown)，确保后续内容与定义名称安全平移
                    activeSheet.Rows[$"{targetDetailStartRow}:{targetDetailStartRow + copyRowCount - 1}"].Insert(-4121);

                    // 10. 只读方式打开 CabinetTemplate.xlsx 模板工作簿并复制 41:74 行
                    dynamic templateWb = app.Workbooks.Open(templatePath, ReadOnly: true);
                    try
                    {
                        // 获取模板中的源工作表
                        dynamic templateSheet = null;
                        try
                        {
                            // 尝试按配置的工作表名称获取
                            templateSheet = templateWb.Sheets[defaultTemplateSheet];
                        }
                        catch
                        {
                            // 回退取第 2 个工作表或第 1 个工作表
                            templateSheet = templateWb.Sheets.Count >= 2 ? templateWb.Sheets[2] : templateWb.Sheets[1];
                        }

                        // 提取模板 41 行至 74 行源区域 (完整箱柜明细块)
                        dynamic srcRange = templateSheet.Rows["41:74"];
                        // 提取当前表目标写入区域
                        dynamic dstRange = activeSheet.Rows[$"{targetDetailStartRow}:{targetDetailStartRow + copyRowCount - 1}"];

                        // 执行复制操作 (完整包含格式、公式与边框)
                        srcRange.Copy(dstRange);
                    }
                    finally
                    {
                        // 复制完成后立即关闭模板工作簿句柄 (不保存)
                        templateWb.Close(false);
                    }

                    // 11. 计算新复制箱柜明细的关键物理行号映射
                    // 模板中 44 行对应 Cab_Det，相对起始行 41 的偏移为 3 行
                    int newDetRow = targetDetailStartRow + (44 - 41);
                    // 模板中 65 行对应 Cab_Subsum，相对起始行 41 的偏移为 24 行
                    int newSubsumRow = targetDetailStartRow + (65 - 41);
                    // 模板中 71 行对应 Cab_Tolsum，相对起始行 41 的偏移为 30 行
                    int newTolsumRow = targetDetailStartRow + (71 - 41);

                    // 12. 注册当前工作表的 4 个定义名称 (规则 6)
                    string curSheetName = Convert.ToString(activeSheet.Name) ?? "";
                    // 注册 Cab_Sum_{K}
                    string sumNameTag = $"{sumPrefix}{cabinetK}";
                    Tool.SafeSetSheetName(activeSheet, curSheetName, sumNameTag, insertSumRow);

                    // 注册 Cab_Det_{K}
                    string detNameTag = $"{detPrefix}{cabinetK}";
                    Tool.SafeSetSheetName(activeSheet, curSheetName, detNameTag, newDetRow);

                    // 注册 Cab_Subsum_{K}
                    string subsumNameTag = $"{subsumPrefix}{cabinetK}";
                    Tool.SafeSetSheetName(activeSheet, curSheetName, subsumNameTag, newSubsumRow);

                    // 注册 Cab_Tolsum_{K}
                    string tolsumNameTag = $"{tolsumPrefix}{cabinetK}";
                    Tool.SafeSetSheetName(activeSheet, curSheetName, tolsumNameTag, newTolsumRow);

                    // 13. 建立汇总行与明细行之间的双向超链接绑定 (规则 6)
                    try
                    {
                        dynamic sumAnchorCell = activeSheet.Cells[insertSumRow, 1];
                        dynamic detAnchorCell = activeSheet.Cells[newDetRow, 1];

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

                    // 14. 设置初始箱柜名称并绑定汇总行公式与清洗明细表头
                    string cabDisplayName = string.IsNullOrWhiteSpace(initialCabName) ? $"箱柜{cabinetK}" : initialCabName.Trim();
                    // 写入汇总行箱柜名称 (Cell B)
                    activeSheet.Cells[insertSumRow, 2].Value = cabDisplayName;
                    // 写入明细行箱柜名称 (Cell B)
                    activeSheet.Cells[newDetRow, 2].Value = cabDisplayName;
                    // 保留明细表头 C 列静态标签(型号:)，清空明细表头备注旧数据 (Cell I)
                    activeSheet.Cells[newDetRow, 9].Value = string.Empty;
                    activeSheet.Cells[insertSumRow, 5].Formula = $"台";
                    // 汇总行公式绑定至明细总计行
                    // G 列单价公式指向明细总计行的销售总价 (H 列)
                    activeSheet.Cells[insertSumRow, 7].Formula = $"=H{newTolsumRow}";
                    // H 列总价公式 = 数量(F列) * 单价(G列)
                    activeSheet.Cells[insertSumRow, 8].Formula = $"=F{insertSumRow}*G{insertSumRow}";
                    // J 列成本总价公式指向明细总计行的成本总价 (K 列)
                    activeSheet.Cells[insertSumRow, 10].Formula = $"=K{newTolsumRow}";
                    // K 列毛利公式 = 总价 - 成本总价
                    activeSheet.Cells[insertSumRow, 11].Formula = $"=H{insertSumRow}-J{insertSumRow}";
                    // L 列毛利率公式
                    activeSheet.Cells[insertSumRow, 12].Formula = $"=IF(H{insertSumRow}=0,0,K{insertSumRow}/H{insertSumRow})";

                    // 15. 激活当前工作表并聚焦光标至新汇总行 B 列
                    activeSheet.Activate();
                    activeSheet.Cells[insertSumRow, 2].Select();

                    // 16. 返回新建箱柜的关键信息实体
                    return new Models.CabinetCreatedInfo
                    {
                        // 设置箱柜序号 K
                        CabinetK = cabinetK,
                        // 设置顶部汇总行号
                        SumRow = insertSumRow,
                        // 设置箱柜信息行号
                        DetRow = newDetRow,
                        // 设置小计行号
                        SubsumRow = newSubsumRow,
                        // 设置总计行号
                        TolsumRow = newTolsumRow
                    };
                }
                finally
                {
                    // 恢复屏幕刷新与系统事件
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"CopyCabinetDetailFromTemplate 异常: {ex.Message}");
                // 弹窗提示异常
                System.Windows.Forms.MessageBox.Show(
                    $"从模板复制箱柜明细异常: {ex.Message}",
                    "系统提示",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                // 返回空对象
                return null;
            }
        }

        /// <summary>
        /// 供 Ribbon 菜单及右键快捷菜单调用的新建箱柜入口
        /// </summary>
        public static void CreateNewCabinetFromSelection()
        {
            CopyCabinetDetailFromTemplate();
            // 调度核心新建箱柜业务逻辑
            //CreateNewCabinet();
        }

        /// <summary>
        /// 供 Ribbon 菜单调用的删除箱柜入口
        /// 智能识别用户光标当前所在的箱柜或顶部汇总选中的多个箱柜，弹出确认提示后执行精准批量删除
        /// </summary>
        public static void DeleteCabinetFromSelection()
        {
            try
            {
                // 获取当前 Excel 活动运行上下文 (安全调用)
                var context = Tool.GetActiveExcelContext();
                if (context == null) return;
                dynamic app = context.App;
                dynamic activeWb = context.Wb;
                dynamic activeSheet = context.Sheet;

                // 构建当前工作表有效箱柜映射集合 (显式强类型接收，避免隐式 dynamic 传染)
                List<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets = Tool.GetSheetValidCabinets(activeSheet, activeWb);

                // 校验当前工作表是否存在有效箱柜
                if (validCabinets == null || validCabinets.Count == 0)
                {
                    // 弹出提示：未识别到箱柜
                    System.Windows.Forms.MessageBox.Show(
                        "当前工作表未识别到任何箱柜定义，无法执行删除操作。",
                        "删除箱柜提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // 智能查找当前选区或光标命中的所有目标箱柜 (支持单选与多选)
                var selectedCabinets = Tool.GetSelectedCabinets(app, validCabinets, fallbackActiveCell: true, fallbackSingle: true);

                // 若未命中任何箱柜
                if (selectedCabinets == null || selectedCabinets.Count == 0)
                {
                    // 弹出友好提示引导用户点击或框选目标箱柜
                    System.Windows.Forms.MessageBox.Show(
                        "请先在工作表中将光标移至或框选要删除的箱柜汇总行或明细区域，然后再次点击【删除箱柜】。",
                        "删除箱柜提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // 准备确认提示信息与待删除的序号列表
                var targetKs = new List<int>();
                string confirmMsg = string.Empty;

                // 区分单选与多选交互场景
                if (selectedCabinets.Count == 1)
                {
                    // 单箱柜场景
                    var targetCab = selectedCabinets[0];
                    int targetK = targetCab.Key;
                    targetKs.Add(targetK);
                    var anchor = targetCab.Value;
                    int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                    int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;

                    // 读取箱柜名称 (优先从汇总行 B 列获取，回退从明细行 B 列获取)
                    string cabName = string.Empty;
                    if (sumRow > 0)
                    {
                        try { cabName = Convert.ToString(activeSheet.Cells[sumRow, 2].Value)?.Trim() ?? ""; } catch { }
                    }
                    if (string.IsNullOrWhiteSpace(cabName) && detRow > 0)
                    {
                        try { cabName = Convert.ToString(activeSheet.Cells[detRow, 2].Value)?.Trim() ?? ""; } catch { }
                    }
                    if (string.IsNullOrWhiteSpace(cabName)) cabName = $"箱柜{targetK}";

                    // 构造单箱柜提示语
                    confirmMsg = validCabinets.Count > 1
                        ? $"确定要删除【{cabName}】(序号: {targetK}) 吗？\n\n此操作将删除该箱柜的顶部汇总行及底部完整明细区块，且不可恢复。"
                        : $"当前分类表仅有 1 台箱柜【{cabName}】。\n\n确定要清空并重置该箱柜为初始空白状态吗？";
                }
                else
                {
                    // 多箱柜批量场景：搜集所有选中箱柜的显示名称
                    var cabNames = new List<string>();
                    foreach (var cab in selectedCabinets)
                    {
                        targetKs.Add(cab.Key);
                        int sumRow = cab.Value.Sum != null ? Convert.ToInt32(cab.Value.Sum.Row) : 0;
                        int detRow = cab.Value.Det != null ? Convert.ToInt32(cab.Value.Det.Row) : 0;
                        string cabName = string.Empty;
                        if (sumRow > 0)
                        {
                            try { cabName = Convert.ToString(activeSheet.Cells[sumRow, 2].Value)?.Trim() ?? ""; } catch { }
                        }
                        if (string.IsNullOrWhiteSpace(cabName) && detRow > 0)
                        {
                            try { cabName = Convert.ToString(activeSheet.Cells[detRow, 2].Value)?.Trim() ?? ""; } catch { }
                        }
                        if (string.IsNullOrWhiteSpace(cabName)) cabName = $"箱柜{cab.Key}";
                        cabNames.Add($"【{cabName}】(序号: {cab.Key})");
                    }

                    // 拼接箱柜名称列表
                    string namesListStr = string.Join("\n", cabNames);

                    // 若全选了当前表的所有箱柜
                    if (selectedCabinets.Count >= validCabinets.Count)
                    {
                        confirmMsg = $"检测到您选中了当前工作表的全部 {selectedCabinets.Count} 台箱柜：\n\n{namesListStr}\n\n确定要清空并重置为初始的 1 台空白箱柜吗？";
                    }
                    else
                    {
                        confirmMsg = $"确定要批量删除以下 {selectedCabinets.Count} 台箱柜吗？\n\n{namesListStr}\n\n此操作将删除所选箱柜的顶部汇总行及底部完整明细区块，且不可恢复。";
                    }
                }

                // 弹出删除确认对话框
                var dialogRes = System.Windows.Forms.MessageBox.Show(
                    confirmMsg,
                    "删除箱柜确认",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question);

                // 若用户取消则直接退出
                if (dialogRes != System.Windows.Forms.DialogResult.Yes) return;

                // 执行核心批量删除业务逻辑
                DeleteCabinets(app, activeSheet, targetKs);
            }
            catch (Exception ex)
            {
                // 弹出异常提示
                System.Windows.Forms.MessageBox.Show(
                    $"删除箱柜失败: {ex.Message}",
                    "系统提示",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 核心方法：在当前分类表中批量删除指定序号的箱柜集合
        /// 遵循方案 A（自底向上精准删除明细块与汇总行，清理 4 个定义名称，全删保护机制）
        /// 支持外部显式传入 Excel COM Application 与 Worksheet
        /// </summary>
        /// <param name="explicitApp">可选显式传入的 Excel COM Application 实例</param>
        /// <param name="explicitSheet">可选显式传入的 Worksheet 实例</param>
        /// <param name="cabinetKs">待删除的箱柜序号列表</param>
        /// <returns>删除是否成功</returns>
        public static bool DeleteCabinets(
            dynamic? explicitApp,
            dynamic? explicitSheet,
            List<int> cabinetKs)
        {
            // 校验待删除序号列表有效性
            if (cabinetKs == null || cabinetKs.Count == 0) return false;

            try
            {
                // 获取当前运行的 Excel Application COM 接口实例与工作簿工作表 (复用 Tool 公共方法)
                var context = Tool.GetActiveExcelContext(explicitApp, explicitSheet);
                if (context == null) return false;
                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic activeSheet = context.Sheet;

                // 读取箱柜定义名称前缀值对象 (零堆分配)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

                // 关闭屏幕刷新与系统弹窗以提升执行性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                try
                {
                    // 1. 扫描当前工作表有效箱柜映射 (复用 Tool 公共方法)
                    var validCabinets = Tool.GetSheetValidCabinets(activeSheet, wb);
                    if (validCabinets == null || validCabinets.Count == 0) return false;

                    // 匹配所有待删除的目标箱柜实体 (使用 HashSet 去重检索)
                    var targetKSet = new HashSet<int>(cabinetKs);
                    var toDeleteList = new List<KeyValuePair<int, Models.CabinetAnchorModel>>();
                    foreach (var cab in validCabinets)
                    {
                        if (targetKSet.Contains(cab.Key))
                        {
                            toDeleteList.Add(cab);
                        }
                    }

                    // 校验是否存在命中的待删除箱柜
                    if (toDeleteList.Count == 0) return false;

                    // 2. 特殊情况：若全选了所有箱柜（即删除后工作表将无箱柜），执行首台箱柜重置保护
                    if (toDeleteList.Count >= validCabinets.Count)
                    {
                        // 若原表有多台箱柜，保留第 1 台箱柜，删除其余第 2~N 台箱柜
                        if (validCabinets.Count > 1)
                        {
                            // 提取从第 2 台开始的其余箱柜序号进行物理删除
                            var otherKs = new List<int>();
                            for (int i = 1; i < validCabinets.Count; i++)
                            {
                                otherKs.Add(validCabinets[i].Key);
                            }

                            // 递归调用删除其余箱柜
                            DeleteCabinets(app, activeSheet, otherKs);
                        }

                        // 重新获取并重置保留下来的第 1 台箱柜
                        var remainingCabinets = Tool.GetSheetValidCabinets(activeSheet, wb);
                        if (remainingCabinets.Count > 0)
                        {
                            var firstCab = remainingCabinets[0];
                            int firstK = firstCab.Key;
                            var firstAnchor = firstCab.Value;
                            int sumR = firstAnchor.Sum != null ? Convert.ToInt32(firstAnchor.Sum.Row) : 0;
                            int detR = firstAnchor.Det != null ? Convert.ToInt32(firstAnchor.Det.Row) : 0;
                            int subsumR = firstAnchor.Subsum != null ? Convert.ToInt32(firstAnchor.Subsum.Row) : (detR + 24);

                            // 清空汇总行数据与属性
                            if (sumR > 0)
                            {
                                activeSheet.Cells[sumR, 2].Value = "箱柜1";
                                activeSheet.Cells[sumR, 3].Value = string.Empty;
                                activeSheet.Cells[sumR, 4].Value = string.Empty;
                                activeSheet.Cells[sumR, 5].Value = string.Empty;
                                activeSheet.Cells[sumR, 6].Value = 1;
                                activeSheet.Cells[sumR, 13].Value = string.Empty;
                            }

                            // 清空明细表头
                            if (detR > 0)
                            {
                                activeSheet.Cells[detR, 2].Value = "箱柜1";
                                activeSheet.Cells[detR, 3].Value = string.Empty;
                                activeSheet.Cells[detR, 9].Value = string.Empty;

                                // 重新构建标准空白元器件矩阵 (规则 6 & 规则 7)
                                int compStartRow = detR + 2;
                                int compEndRow = subsumR - 1;
                                if (compEndRow >= compStartRow)
                                {
                                    object[,] compMatrix = Tool.BuildComponentRowsMatrix(compStartRow, compEndRow, detR, 17);
                                    activeSheet.Range[$"A{compStartRow}:Q{compEndRow}"].Formula = compMatrix;
                                }
                            }

                            // 激活并选中汇总行
                            activeSheet.Activate();
                            if (sumR > 0) activeSheet.Cells[sumR, 2].Select();
                        }
                        return true;
                    }

                    // 3. 部分删除模式：搜集每个待删除箱柜的物理行号范围
                    var deleteBlocks = new List<CabinetDeleteInfo>();
                    foreach (var cab in toDeleteList)
                    {
                        var anchor = cab.Value;
                        int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                        int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;
                        int tolsumRow = anchor.Tolsum != null ? Convert.ToInt32(anchor.Tolsum.Row) : (detRow + 27);

                        // 计算明细区块行范围 [detailStartRow, detailEndRow] (从大标题到总计行下方3行报价人信息)
                        int detailStartRow = detRow - 3;
                        // 若起始行小于 1 则兜底使用 detRow
                        if (detailStartRow < 1) detailStartRow = detRow;
                        // 结束行包含总计行及紧随其后的 3 行报价人信息 (完整明细块)
                        int detailEndRow = tolsumRow + 3;

                        // 检查明细块下方是否包含 1 行分隔空行，若有连同空行一起删除保持整洁
                        try
                        {
                            string nextRowCellA = Convert.ToString(activeSheet.Cells[detailEndRow + 1, 1].Value) ?? "";
                            string nextRowCellB = Convert.ToString(activeSheet.Cells[detailEndRow + 1, 2].Value) ?? "";
                            if (string.IsNullOrWhiteSpace(nextRowCellA) && string.IsNullOrWhiteSpace(nextRowCellB))
                            {
                                detailEndRow += 1;
                            }
                        }
                        catch { }

                        deleteBlocks.Add(new CabinetDeleteInfo
                        {
                            CabinetK = cab.Key,
                            SumRow = sumRow,
                            DetailStartRow = detailStartRow,
                            DetailEndRow = detailEndRow
                        });
                    }

                    // 4. 执行物理删除第一阶段：按明细块起始行号降序（自底向上，从大到小）删除所有明细区块
                    // 由于明细行全部位于汇总行下方，从下往上删除明细块不会改变上方任何明细行与汇总行的物理行号
                    deleteBlocks.Sort((a, b) => b.DetailStartRow.CompareTo(a.DetailStartRow));
                    foreach (var block in deleteBlocks)
                    {
                        if (block.DetailStartRow > 0 && block.DetailEndRow >= block.DetailStartRow)
                        {
                            // 删除底部明细行 (-4162 对应 xlShiftUp 向上移)
                            activeSheet.Rows[$"{block.DetailStartRow}:{block.DetailEndRow}"].Delete(-4162);
                        }
                    }

                    // 5. 执行物理删除第二阶段：按汇总行行号降序（自底向上，从大到小）删除所有顶部汇总行
                    // 明细块删除完毕后汇总行原始行号完好无损，从下往上删除汇总行不会改变上方汇总行的行号
                    deleteBlocks.Sort((a, b) => b.SumRow.CompareTo(a.SumRow));
                    foreach (var block in deleteBlocks)
                    {
                        if (block.SumRow > 0)
                        {
                            // 删除顶部汇总行 (-4162 对应 xlShiftUp)
                            activeSheet.Rows[$"{block.SumRow}:{block.SumRow}"].Delete(-4162);
                        }
                    }

                    // 6. 安全清理所有被删除箱柜的 4 个定义名称 (方案 A)
                    foreach (var block in deleteBlocks)
                    {
                        string sumNameTag = $"{sumPrefix}{block.CabinetK}";
                        string detNameTag = $"{detPrefix}{block.CabinetK}";
                        string subsumNameTag = $"{subsumPrefix}{block.CabinetK}";
                        string tolsumNameTag = $"{tolsumPrefix}{block.CabinetK}";

                        Tool.SafeDeleteName(activeSheet, wb, sumNameTag);
                        Tool.SafeDeleteName(activeSheet, wb, detNameTag);
                        Tool.SafeDeleteName(activeSheet, wb, subsumNameTag);
                        Tool.SafeDeleteName(activeSheet, wb, tolsumNameTag);
                    }

                    // 7. 激活当前工作表并聚焦光标至合理的汇总区域
                    activeSheet.Activate();
                    try
                    {
                        // 尝试重新扫描剩余的首个箱柜汇总行并选中
                        var remaining = Tool.GetSheetValidCabinets(activeSheet, wb);
                        if (remaining.Count > 0 && remaining[0].Value.Sum != null)
                        {
                            int firstSumR = Convert.ToInt32(remaining[0].Value.Sum.Row);
                            activeSheet.Cells[firstSumR, 2].Select();
                        }
                    }
                    catch { }

                    return true;
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
                // 记录日志并报错
                LogHelper.WriteLog($"DeleteCabinets 异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show(
                    $"批量删除箱柜发生异常: {ex.Message}",
                    "系统提示",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// 内部辅助结构：记录待删除箱柜的物理坐标与序号
        /// </summary>
        private struct CabinetDeleteInfo
        {
            public int CabinetK;
            public int SumRow;
            public int DetailStartRow;
            public int DetailEndRow;
        }

        /// <summary>
        /// 核心方法：在当前分类表中删除指定序号的箱柜
        /// 遵循方案 A（轻量精准删除，先删底部明细块再删顶部汇总行，清理 4 个定义名称）
        /// 支持外部显式传入 Excel COM Application 与 Worksheet
        /// </summary>
        /// <param name="explicitApp">可选显式传入的 Excel COM Application 实例</param>
        /// <param name="explicitSheet">可选显式传入的 Worksheet 实例</param>
        /// <param name="cabinetK">待删除的箱柜序号 K</param>
        /// <returns>删除是否成功</returns>
        public static bool DeleteCabinet(
            dynamic? explicitApp,
            dynamic? explicitSheet,
            int cabinetK)
        {
            // 重定向调用批量删除方法
            return DeleteCabinets(explicitApp, explicitSheet, new List<int> { cabinetK });
        }


        /// <summary>
        /// 核心方法：在当前分类表中新建箱柜
        /// 遵循规则 6（顶部汇总行、底部明细块、4 个定义名称及超链接）与规则 7（内存二维数组批量操作）
        /// 支持显式传入外部 COM Application 与 Worksheet（方便 AutoCAD / TuFan 等外部模块直接调用）
        /// </summary>
        /// <param name="explicitApp">可选显式传入的 Excel COM Application 实例</param>
        /// <param name="explicitSheet">可选显式传入的 Worksheet 实例</param>
        /// <param name="initialCabName">可选显式传入的箱柜名称</param>
        /// <returns>新建箱柜的核心行号及序号对象，失败返回 null</returns>
        public static Models.CabinetCreatedInfo? CreateNewCabinet(
            dynamic? explicitApp = null,
            dynamic? explicitSheet = null,
            string? initialCabName = null)
        {
            try
            {
                // 获取当前运行的 Excel Application COM 接口实例与工作簿工作表 (复用 Tool 公共方法)
                var context = Tool.GetActiveExcelContext(explicitApp, explicitSheet);
                if (context == null) return null;
                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic activeSheet = context.Sheet;

                // 读取全局配置参数
                var cfg = ConfigManager.Instance.Current.Excel;
                // 读取箱柜定义名称前缀值对象 (零堆分配)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = cfg.Prefixes;

                // 关闭屏幕刷新与系统弹窗以提升执行性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                int insertRow = 0;
                try
                {
                    // 1. 扫描当前工作表有效箱柜映射 (复用 Tool 公共方法)
                    var validCabinets = Tool.GetSheetValidCabinets(activeSheet, wb);

                    // 探测工作表末尾箱柜行号分布 (传入 -1 显式表示获取当前表中最后一个箱柜分布)
                    var lastIndexes = Tool.FindStandardCategoryRowIndexes((object)activeSheet, -1);


                    // 2. 动态计算下一个全新的独立箱柜序号 K
                    int cabinetK = GetNextCabinetIndex(wb, activeSheet);

                    // 3. 根据当前光标/选区所在行直接获取选中的箱柜实体 (复用 Tool 公共方法)
                    var activeCab = Tool.GetActiveCabinet(app, validCabinets, fallbackSingle: true);
                    // 提取目标源箱柜锚点 (优先使用当前选中的箱柜，未命中且有多台时回退取最后一个有效箱柜)
                    var srcCabAnchor = activeCab?.Value ?? (validCabinets.Count > 0 ? validCabinets[validCabinets.Count - 1].Value : null);

                    // 定位顶部汇总行目标插入位置 (紧随选中箱柜的汇总行之后插入，未命中则紧随基准起始行)
                    if (srcCabAnchor?.Sum != null)
                    {
                        // 插入在选中箱柜的汇总行下一行
                        insertRow = Convert.ToInt32(srcCabAnchor.Sum.Row) + 1;
                    }
                    else
                    {
                        // 首个箱柜使用基准起始行
                        insertRow = lastIndexes.cabSumRow + 1;
                    }

                    // 检查目标行 A 列与 B 列是否包含内容或公式，防止误覆盖“合计”行或下部表头
                    string checkCellValA = Convert.ToString(activeSheet.Cells[insertRow, 2].Value)?.Trim() ?? "";
                    // 获取 B 列单元格纯文本
                    string checkCellValB = Convert.ToString(activeSheet.Cells[insertRow, 3].Value)?.Trim() ?? "";
                    // 只要 A 列或 B 列有内容则判定需要物理插入行
                    bool needInsertSumRow = !string.IsNullOrWhiteSpace(checkCellValA) || !string.IsNullOrWhiteSpace(checkCellValB);

                    // 记录顶部汇总行是否触发了物理插行
                    bool didInsertSumRow = false;
                    // 若需要插行则调用 Excel COM API 插入整行
                    if (needInsertSumRow)
                    {
                        // 物理插入 1 行 (-4121 对应 xlShiftDown)
                        activeSheet.Rows[$"{insertRow}:{insertRow}"].Insert(-4121);
                        // 标记已执行物理插行
                        didInsertSumRow = true;
                    }

                    // 4. 定位源箱柜（当前行所在箱柜/上一个箱柜）的标准行号分布与末尾插入位置
                    // 提取源箱柜序号 K (未命中时取末尾箱柜 -1)
                    int srcK = activeCab?.Key ?? (validCabinets.Count > 0 ? validCabinets[validCabinets.Count - 1].Key : -1);
                    // 根据源箱柜序号获取其标准行号分布 (复用 Tool 公共方法)
                    var srcIndexes = Tool.FindStandardCategoryRowIndexes((object)activeSheet, srcK);
                    int srcDetRow = srcIndexes.cabDetRow;
                    int srcSubsumRow = srcIndexes.cabSubsumRow;
                    int srcTolsumRow = srcIndexes.cabTolsumRow;

                    // 若顶部汇总行执行了物理插入，下方原本扫描到的所有源箱柜明细与总计行号均已在 Excel 中物理下移 1 行，执行同步补偿
                    if (didInsertSumRow)
                    {
                        // 补偿源箱柜信息行号
                        srcDetRow += 1;
                        // 补偿源箱柜小计行号
                        srcSubsumRow += 1;
                        // 补偿源箱柜总计行号
                        srcTolsumRow += 1;
                    }

                    // 动态获取源箱柜计费区域跨度 (从 Cab_Subsum 到 Cab_Tolsum 的总行数，包含小计与总计行)
                    int feeSpan = srcTolsumRow - srcSubsumRow + 1;
                    // 兜底校验计费区域跨度有效性 --硬编码--
                    if (feeSpan < 2) feeSpan = 6;

                    // 获取源箱柜元器件实际行数
                    int srcCompStart = srcDetRow + 2;
                    int srcCompEnd = srcSubsumRow - 1;
                    int srcCompRows = Math.Max(0, srcCompEnd - srcCompStart + 1);

                    // 动态推导标准元器件行数 (基于 appsettings.json 基础行号)
                    // 配置文件硬编码: CabDetRowIndex (基准明细行44), CabTolsumRowIndex (基准总计行71) --硬编码--
                    int cfgDet = cfg.CabDetRowIndex > 0 ? cfg.CabDetRowIndex : 44;
                    int cfgTol = cfg.CabTolsumRowIndex > 0 ? cfg.CabTolsumRowIndex : 71;
                    // 计算标准明细块从 Det 到 Tolsum 的总跨度 (默认 71-44+1 = 28)
                    int standardTotalSpan = cfgTol - cfgDet + 1;
                    // 动态推导标准元器件行数: 标准总跨度 - 表头2行(Det行+列标题行) - 动态计费跨度(feeSpan)
                    int standardCompRows = standardTotalSpan - 2 - feeSpan;
                    // 兜底保护标准元器件行数 --硬编码--
                    if (standardCompRows <= 0) standardCompRows = 19;

                    // 动态计算源明细块大标题物理行 (Cab_Det 上方 3 行) 与完整结束行 (包含总计行下方3行报价人信息)
                    int copyStartRow = srcDetRow - 3;
                    int copyEndRow = srcTolsumRow + 3;
                    // 计算待复制的明细块总行数
                    int copyRowCount = copyEndRow - copyStartRow + 1;
                    // 异常兜底校验行数有效性 --硬编码--
                    if (copyRowCount <= 0) copyRowCount = cfg.TemplateDetailBlockTotalRows + 3;

                    // 5. 根据汇总行所选中的源箱柜，定位新明细块的目标起始行 (紧随源箱柜明细块的结束行之后)
                    int newDetailStartRow = copyEndRow + 1;

                    // 物理向下插入 copyRowCount 行空间 (-4121 对应 xlShiftDown)，确保后续已存在的箱柜明细块安全下移不被覆盖
                    activeSheet.Rows[$"{newDetailStartRow}:{newDetailStartRow + copyRowCount - 1}"].Insert(-4121);

                    // 提取源明细块区域与目标插入区域
                    dynamic copyRange = activeSheet.Rows[$"{copyStartRow}:{copyEndRow}"];
                    // 提取目标写入区域范围
                    dynamic targetRange = activeSheet.Rows[$"{newDetailStartRow}:{newDetailStartRow + copyRowCount - 1}"];
                    // 将包含完整格式、公式、边框与报价人信息的明细区块完整复制到新插入的区域
                    copyRange.Copy(targetRange);

                    // 清洗复制公式中的外部工作簿文件路径引用 (显式跳过 A 列以保护定义名称与超链接)
                    Tool.CleanRangeFormulas(targetRange);

                    // 6. 计算新箱柜的关键行号初始映射
                    int newDetRow = newDetailStartRow + 3; // 箱柜信息行 (大标题后第 3 行)
                    int newCompStartRow = newDetRow + 2;   // 元器件起始行 (规则 6: Cab_Det + 2)
                    int newSubsumRow = newDetailStartRow + (srcSubsumRow - copyStartRow); // 初始小计行
                    int newTolsumRow = newDetailStartRow + (srcTolsumRow - copyStartRow); // 初始总计行

                    // 计算被复制箱柜中多余的元器件行数
                    int extraRows = srcCompRows - standardCompRows;
                    // 若被复制箱柜元器件行数超出标准行数，物理删除多余行
                    if (extraRows > 0)
                    {
                        // 确定待删除多余行的物理起始与终止行号 (在元器件区域末尾删除)
                        int delStartRow = newCompStartRow + standardCompRows;
                        int delEndRow = newSubsumRow - 1;
                        // 执行物理向上删除多余行 (-4162 对应 xlShiftUp)
                        activeSheet.Rows[$"{delStartRow}:{delEndRow}"].Delete(-4162);

                        // 删行后，下方的小计行、总计行及报价人信息行号均同步向上偏移 extraRows
                        newSubsumRow -= extraRows;
                        newTolsumRow -= extraRows;
                    }

                    // 最终计算确立元器件终止行 (规则 6: Cab_Subsum - 1)
                    int newCompEndRow = newSubsumRow - 1;

                    // 7. 规则 7: 调用公共方法一次性批量重写新箱柜的元器件区域 (覆盖 A~Q 列自适应空行公式)
                    if (newCompEndRow >= newCompStartRow)
                    {
                        // 内存构建标准纯净的二维元器件数据与公式矩阵
                        object[,] compMatrix = Tool.BuildComponentRowsMatrix(newCompStartRow, newCompEndRow, newDetRow, 17);
                        // 一次性批量写回工作表
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

                    // 11. 写入初始箱柜名称并同步公式与清洗明细表头旧数据
                    // 设置初始箱柜名称 (优先使用传入的名称)
                    string cabDisplayName = string.IsNullOrWhiteSpace(initialCabName) ? $"箱柜{cabinetK}" : initialCabName.Trim();
                    // 写入汇总行箱柜名称
                    activeSheet.Cells[insertRow, 2].Value = cabDisplayName;
                    activeSheet.Cells[insertRow, 5].Value = "台";
                    // 写入明细行箱柜名称
                    activeSheet.Cells[newDetRow, 2].Value = cabDisplayName;
                    // 保留明细表头 C 列静态标签(型号:)，清空明细表头备注旧数据
                    activeSheet.Cells[newDetRow, 9].Value = string.Empty;

                    // 汇总行公式绑定至明细总计行
                    // G 列单价公式指向明细总计行的销售总价 (H 列)
                    activeSheet.Cells[insertRow, 7].Formula = $"=H{newTolsumRow}";
                    // H 列总价公式 = 数量(F列) * 单价(G列)
                    activeSheet.Cells[insertRow, 8].Formula = $"=F{insertRow}*G{insertRow}";
                    // J 列成本总价公式指向明细总计行的成本总价 (K 列)
                    activeSheet.Cells[insertRow, 10].Formula = $"=K{newTolsumRow}";
                    // K 列毛利公式 = 总价 - 成本总价
                    activeSheet.Cells[insertRow, 11].Formula = $"=H{insertRow}-J{insertRow}";
                    // L 列毛利率公式
                    activeSheet.Cells[insertRow, 12].Formula = $"=IF(H{insertRow}=0,0,K{insertRow}/H{insertRow})";

                    // 12. 激活原工作表并选中新插入的汇总行
                    activeSheet.Activate();
                    activeSheet.Cells[insertRow, 2].Select();

                    // 返回新建箱柜的关键信息实体对象
                    return new Models.CabinetCreatedInfo
                    {
                        CabinetK = cabinetK,
                        SumRow = insertRow,
                        DetRow = newDetRow,
                        SubsumRow = newSubsumRow,
                        TolsumRow = newTolsumRow
                    };
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

            // 读取箱柜定义名称前缀值对象 (零堆分配)
            var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

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
        /// 从 Excel 工作表中反向解析指定箱柜的 CabinetObject 实体数据模型
        /// </summary>
        public static Models.CabinetObject? ParseCabinetObjectFromSheet(dynamic sheet, int cabinetIndex)
        {
            if (sheet == null || cabinetIndex <= 0) return null;

            try
            {
                // 读取箱柜定义名称前缀值对象 (零堆分配)
                var prefixes = CabinetPrefixConfig.Current;
                // 获取明细信息行定义名称
                string detTagName = prefixes.GetDetName(cabinetIndex);
                // 获取汇总行定义名称
                string sumTagName = prefixes.GetSumName(cabinetIndex);

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
        /// <summary>
        /// 批量导入/导出箱柜集合至 Excel 活动工作簿 (极速优化版)
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
            // 记录批量导出入口日志及箱柜总数
            LogHelper.WriteLog($"[BatchExport] 开始批量导出，传入待导出箱柜总数: {cabinets?.Count ?? 0}");

            // 校验待导出集合是否有效
            if (cabinets == null || cabinets.Count == 0)
            {
                // 记录入参为空警告日志
                LogHelper.WriteLog("[BatchExport] 待导出箱柜列表为空，终止导出流程");
                return 0;
            }

            // 获取 Excel COM Application 接口 (带存活探测与自动重连保护)
            dynamic? app = null;
            if (explicitApp != null)
            {
                try
                {
                    // 探测传入的 COM 句柄是否依然存活有效
                    var _ = explicitApp.Version;
                    app = explicitApp;
                }
                catch
                {
                    // 记录传入句柄断开日志
                    LogHelper.WriteLog("[BatchExport] 传入的 explicitApp COM 句柄已断开(0x80010114)，尝试重新获取当前运行中的 Excel 实例");
                }
            }
            if (app == null)
            {
                try
                {
                    // 重新从 ROT 获取活动 Excel 实例
                    app = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch { }
            }
            if (app == null)
            {
                // 回退 ExcelDnaSafeAccessor
                app = ExcelDnaSafeAccessor.GetApplication();
            }

            if (app == null)
            {
                // 记录获取 Excel Application 失败日志
                LogHelper.WriteLog("[BatchExport] 获取 Excel Application COM 实例失败，终止导出");
                return 0;
            }

            // 获取活动工作簿
            dynamic? activeWb = null;
            try
            {
                activeWb = app.ActiveWorkbook;
            }
            catch (Exception exWb)
            {
                LogHelper.WriteLog($"[BatchExport] 获取 ActiveWorkbook 异常: {exWb.Message}");
            }

            if (activeWb == null)
            {
                // 记录获取活动工作簿失败日志
                LogHelper.WriteLog("[BatchExport] 获取活动工作簿 ActiveWorkbook 失败，终止导出");
                return 0;
            }

            // 1. 前向填充空白 Category，确保子箱柜准确继承父分类工作表名称
            string runningCategory = string.Empty;
            foreach (var c in cabinets)
            {
                if (c?.Header == null) continue;
                if (!string.IsNullOrWhiteSpace(c.Header.Category))
                {
                    runningCategory = c.Header.Category.Trim();
                }
                else
                {
                    c.Header.Category = runningCategory;
                }
            }

            // 临时保存 Excel 环境状态以支持还原
            bool prevUpdating = true;
            bool prevAlerts = true;
            bool prevEvents = true;
            int prevCalculation = -4105; // 默认 xlCalculationAutomatic (-4105)
            try { prevUpdating = app.ScreenUpdating; } catch { }
            try { prevAlerts = app.DisplayAlerts; } catch { }
            try { prevEvents = app.EnableEvents; } catch { }
            try { prevCalculation = Convert.ToInt32(app.Calculation); } catch { }

            int successCount = 0;

            try
            {
                // 2. 关闭屏幕刷新与事件调度，锁定手动计算模式以大幅提升批量写入性能
                try { app.ScreenUpdating = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }
                try { app.EnableEvents = false; } catch { }
                try { app.Calculation = -4135; /* xlCalculationManual */ } catch { }

                // 3. 按 Category 分组处理箱柜列表，消除工作表频繁交替切换开销
                var categoryGroups = cabinets
                    .GroupBy(c => string.IsNullOrWhiteSpace(c.Header.Category) ? string.Empty : c.Header.Category.Trim())
                    .ToList();

                // 维护每个分类工作表的内存状态上下文字典
                var sheetContextMap = new Dictionary<string, BatchCategorySheetContext>(StringComparer.OrdinalIgnoreCase);
                string lastSheetName = string.Empty;
                dynamic? lastSheet = null;
                int totalProcessed = 0;

                // 4. 遍历所有分类分组
                foreach (var grp in categoryGroups)
                {
                    string targetCategory = grp.Key;
                    var cabListInGroup = grp.ToList();
                    int cabCountInGroup = cabListInGroup.Count;
                    if (cabCountInGroup == 0) continue;

                    // 分类表路由与新建：智能获取已有表或新建分类工作表
                    dynamic? currentSheet = ResolveCategorySheet(app, activeWb, targetCategory, lastSheetName, ref lastSheetName);
                    if (currentSheet == null)
                    {
                        // 记录解析工作表失败日志
                        LogHelper.WriteLog($"[BatchExport] 分类【{targetCategory}】解析工作表失败，跳过该分类下 {cabCountInGroup} 个箱柜");
                        continue;
                    }
                    lastSheet = currentSheet;

                    // 获取或初始化该分类工作表的内存状态上下文
                    string actualSheetName = Convert.ToString(currentSheet.Name) ?? lastSheetName;
                    if (!sheetContextMap.TryGetValue(actualSheetName, out var sheetCtx))
                    {
                        // 首次进入该分类表，初始化纯净母版与预扩容
                        sheetCtx = InitCategorySheetContext(app, activeWb, currentSheet, actualSheetName, cabCountInGroup);
                        sheetContextMap[actualSheetName] = sheetCtx;
                    }

                    // 5. 顺序遍历写入当前分类下的所有箱柜
                    for (int gIdx = 0; gIdx < cabListInGroup.Count; gIdx++)
                    {
                        var cab = cabListInGroup[gIdx];
                        totalProcessed++;
                        // 触发进度回调通知
                        int progress = (int)(totalProcessed * 100.0 / cabinets.Count);
                        progressCallback?.Invoke(progress, cab.Header.Name);

                        // 调用纯净母版克隆版写入单个箱柜
                        bool ok = ExportSingleCabinetOptimized(
                            app,
                            activeWb,
                            currentSheet,
                            cab,
                            sheetCtx);

                        if (ok) successCount++;
                        else
                        {
                            // 记录单箱柜导出失败日志
                            LogHelper.WriteLog($"[BatchExport] 箱柜【{cab.Header.Name}】写入失败，分类【{actualSheetName}】");
                        }
                    }
                }

                // 6. 统一静默清理各分类工作表中的纯净母版占位 (自底向上物理删除，触发 Excel 自动平移无缝对接)
                foreach (var kvp in sheetContextMap)
                {
                    var sheetCtx = kvp.Value;
                    if (sheetCtx == null || !sheetCtx.HasTemplateToClean) continue;

                    dynamic ws = sheetCtx.Sheet;
                    try
                    {
                        // 先删除纯净母表明细块 (34 行) --硬编码--
                        int tmplDetStart = sheetCtx.TemplateDetailBlockStartRow;
                        int tmplDetCount = 34; // --硬编码--
                        ws.Rows[$"{tmplDetStart}:{tmplDetStart + tmplDetCount - 1}"].Delete(-4119); // -4119 对应 xlShiftUp

                        // 若存在母版汇总行 (全新表场景)，再删除母版汇总行 (1 行)
                        if (sheetCtx.TemplateSumRow > 0)
                        {
                            int tmplSumRow = sheetCtx.TemplateSumRow;
                            ws.Rows[$"{tmplSumRow}:{tmplSumRow}"].Delete(-4119); // -4119 对应 xlShiftUp
                        }
                    }
                    catch (Exception exClean)
                    {
                        // 记录清理母版异常日志
                        LogHelper.WriteLog($"[BatchExport] 清理分类【{kvp.Key}】临时母版异常: {exClean.Message}");
                    }
                }

                // 激活最终操作的分类工作表
                if (lastSheet != null)
                {
                    try { lastSheet.Activate(); } catch { }
                }

                // 记录批量导出完成汇总日志
                LogHelper.WriteLog($"[BatchExport] 极速批量导出全部完成，成功导出 {successCount}/{cabinets.Count} 个箱柜");
            }
            catch (Exception ex)
            {
                // 记录批量导出异常日志
                LogHelper.WriteLog($"[BatchExport] BatchExportCabinets 异常: {ex.Message}");
            }
            finally
            {
                // 7. 一次性触发全局重算并恢复原有 Excel 计算与显示环境
                try { app.Calculate(); } catch { }
                try { app.Calculation = prevCalculation; } catch { }
                try { app.ScreenUpdating = prevUpdating; } catch { }
                try { app.DisplayAlerts = prevAlerts; } catch { }
                try { app.EnableEvents = prevEvents; } catch { }
            }

            // 返回成功导出的箱柜数量
            return successCount;
        }

        /// <summary>
        /// 批量导入分类工作表内存状态上下文模型
        /// </summary>
        private class BatchCategorySheetContext
        {
            // 目标工作表 COM 引用
            public dynamic Sheet { get; set; }
            // 工作表名称
            public string SheetName { get; set; } = string.Empty;
            // 汇总表插入下一行的物理行号
            public int NextSumRow { get; set; }
            // 底部明细块下一行的起始物理行号
            public int NextDetailStartRow { get; set; }
            // 纯净母版明细块的起始物理行号 (34行纯净模板) --硬编码--
            public int TemplateDetailBlockStartRow { get; set; } = 41;
            // 纯净母版汇总行的物理行号 (<=0 表示无母版汇总行)
            public int TemplateSumRow { get; set; } = 0;
            // 是否有纯净母版需要在批量导出完成后统一物理删除
            public bool HasTemplateToClean { get; set; }
            // 下一个递增箱柜序号 K
            public int NextCabinetK { get; set; }

            public BatchCategorySheetContext(dynamic sheet, string sheetName)
            {
                Sheet = sheet;
                SheetName = sheetName;
            }
        }

        /// <summary>
        /// 首次进入分类工作表时初始化内存状态上下文
        /// 严格实施“纯净母版槽位隔离”策略：首槽位绝不填写数据，专供克隆，导出完毕后统一物理清除
        /// </summary>
        private static BatchCategorySheetContext InitCategorySheetContext(
            dynamic app,
            dynamic activeWb,
            dynamic sheet,
            string sheetName,
            int cabCountInGroup)
        {
            var ctx = new BatchCategorySheetContext(sheet, sheetName);

            // 智能探测当前分类表的基准行号分布
            var baseIndexes = Tool.FindStandardCategoryRowIndexes((object)sheet);
            int baseSumRow = baseIndexes.cabSumRow;
            int baseDetRow = baseIndexes.cabDetRow;
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

            if (!isFirstSlotOccupied)
            {
                // 【场景 1：全新工作表或首槽位为空】
                // 保留第 7 行汇总与 41:74 行明细作为纯净母版，绝不直接在上面写入数据
                ctx.HasTemplateToClean = true;
                ctx.TemplateSumRow = baseSumRow; // 默认 7 --硬编码--
                ctx.TemplateDetailBlockStartRow = 41; // 模板首台明细块起始行 --硬编码--
                ctx.NextCabinetK = 1;

                // 汇总区一次性批量预扩容 N 行（插入在母版汇总行下方）
                if (cabCountInGroup > 0)
                {
                    int insertStart = baseSumRow + 1;
                    sheet.Rows[$"{insertStart}:{insertStart + cabCountInGroup - 1}"].Insert(-4121);
                    // 汇总区扩容后，下方的母版明细块整体平移 cabCountInGroup 行
                    ctx.TemplateDetailBlockStartRow += cabCountInGroup;
                }

                // 真实箱柜的汇总行起始位置：母版汇总行下方
                ctx.NextSumRow = baseSumRow + 1;
                // 真实箱柜的明细块起始位置：母版明细块（34行）下方
                ctx.NextDetailStartRow = ctx.TemplateDetailBlockStartRow + 34; // --硬编码--
            }
            else
            {
                // 【场景 2：已有数据的旧分类表】
                // 扫描末尾有效箱柜以确定插入点
                var lastIndexes = Tool.FindStandardCategoryRowIndexes((object)sheet, -1);
                int lastSumRow = lastIndexes.cabSumRow;
                int lastTolsumRow = lastIndexes.cabTolsumRow > 0 ? lastIndexes.cabTolsumRow : 71;

                int insertSumStart = lastSumRow + 1;
                // 汇总区一次性批量预扩容 N 行
                if (cabCountInGroup > 0)
                {
                    sheet.Rows[$"{insertSumStart}:{insertSumStart + cabCountInGroup - 1}"].Insert(-4121);
                }

                ctx.NextSumRow = insertSumStart;
                ctx.TemplateSumRow = 0; // 旧表无需删除母版汇总行

                // 计算表尾明细插入点 (受汇总扩容平移 cabCountInGroup 行)
                int templatePos = lastTolsumRow + 4 + cabCountInGroup; // --硬编码--

                // 从 CabinetTemplate.xlsx 外部模板只读复制 1 次 34 行标准明细块到表尾作为临时纯净母版
                string templatePath = Controllers.ProjectController.EnsureCabinetTemplate(app);
                dynamic templateWb = app.Workbooks.Open(templatePath, ReadOnly: true);
                try
                {
                    // 默认模板工作表 --硬编码--
                    dynamic templateSheet = templateWb.Sheets.Count >= 2 ? templateWb.Sheets[2] : templateWb.Sheets[1];
                    dynamic srcTmplRange = templateSheet.Rows["41:74"]; // --硬编码--
                    dynamic dstTmplRange = sheet.Rows[$"{templatePos}:{templatePos + 34 - 1}"]; // --硬编码--
                    srcTmplRange.Copy(dstTmplRange);
                }
                finally
                {
                    // 立即关闭模板工作簿
                    templateWb.Close(false);
                }

                ctx.HasTemplateToClean = true;
                ctx.TemplateDetailBlockStartRow = templatePos;
                ctx.NextDetailStartRow = templatePos + 34; // --硬编码--
                ctx.NextCabinetK = GetNextCabinetIndex(activeWb, sheet);
            }

            return ctx;
        }

        /// <summary>
        /// 极致极速版：在指定的分类工作表中渲染并写入单个箱柜对象（纯净母版直拷 + 内存推导 + 二维数组批量写入）
        /// </summary>
        private static bool ExportSingleCabinetOptimized(
            dynamic app,
            dynamic targetWb,
            dynamic sheet,
            Models.CabinetObject cab,
            BatchCategorySheetContext ctx)
        {
            if (sheet == null || cab == null || ctx == null) return false;

            try
            {
                string sheetName = ctx.SheetName;
                string safeBoxName = string.IsNullOrWhiteSpace(cab.Header.Name)
                    ? (string.IsNullOrWhiteSpace(cab.Header.CabinetNo) ? "箱柜" : cab.Header.CabinetNo)
                    : cab.Header.Name.Trim();

                // 提取安装方式 (优先读取 InstallMode，回退 Remark)
                string installMode = string.IsNullOrWhiteSpace(cab.Header.InstallMode)
                    ? (cab.Header.Remark ?? string.Empty)
                    : cab.Header.InstallMode.Trim();

                // 1. 为新箱柜分配独立序号与物理行
                int cabinetK = ctx.NextCabinetK++;
                int sumRow = ctx.NextSumRow++;
                int targetDetailStartRow = ctx.NextDetailStartRow;

                // 2. 纯净模板直拷：从纯净母版明细块复制 34 行到表尾目标区 (零平移开销，结构 100% 完整纯净)
                int copyRowCount = 34; // --硬编码--
                int srcStart = ctx.TemplateDetailBlockStartRow;
                dynamic srcRange = sheet.Rows[$"{srcStart}:{srcStart + copyRowCount - 1}"];
                dynamic dstRange = sheet.Rows[$"{targetDetailStartRow}:{targetDetailStartRow + copyRowCount - 1}"];
                srcRange.Copy(dstRange);

                // 计算新建箱柜的关键行号映射
                int detRow = targetDetailStartRow + 3; // --硬编码--
                int subsumRow = detRow + 22; // --硬编码--
                int tolsumRow = detRow + 27; // --硬编码--

                // 3. 计算元器件容量并根据规则 6 处理差额行
                int compStartRow = detRow + 2;
                int defaultCompCount = subsumRow - compStartRow;
                int actualCompCount = cab.Components != null ? cab.Components.Count : 0;

                if (actualCompCount > defaultCompCount)
                {
                    // 计算需要物理插入的差额行数
                    int insertCount = actualCompCount - defaultCompCount;
                    int insertRowStart = compStartRow + defaultCompCount;

                    // 执行批量物理插行 (仅影响当前箱柜自身明细及后续，绝不影响上方的母版)
                    sheet.Rows[$"{insertRowStart}:{insertRowStart + insertCount - 1}"].Insert(-4121);

                    // 更新行号偏移
                    subsumRow += insertCount;
                    tolsumRow += insertCount;
                }

                // 更新上下文中的下一明细块起始物理行号 (总计行 + 3行报价人 + 1)
                ctx.NextDetailStartRow = tolsumRow + 4;
                int compEndRow = subsumRow - 1;

                // 4. 规则 7：调用 Tool.BuildComponentRowsMatrix 一次性批量生成并写入 17 列元器件矩阵
                if (compEndRow >= compStartRow)
                {
                    // 生成 17 列包含自适应公式与元件属性的二维矩阵
                    object[,] compMatrix = Tool.BuildComponentRowsMatrix(compStartRow, compEndRow, detRow, 21, cab.Components);
                    // 批量写入元器件完整区域
                    sheet.Range[$"A{compStartRow}:U{compEndRow}"].Formula = compMatrix;

                    // 批量写入 AA 列 CAD 句柄扩展列
                    int rowCount = compEndRow - compStartRow + 1;
                    object[,] handleMatrix = new object[rowCount, 1];
                    for (int r = 0; r < rowCount; r++)
                    {
                        if (cab.Components != null && r < cab.Components.Count && !string.IsNullOrWhiteSpace(cab.Components[r].Handle))
                        {
                            // 填充 CAD 实体句柄
                            handleMatrix[r, 0] = cab.Components[r].Handle;
                        }
                        else
                        {
                            // 空句柄占位
                            handleMatrix[r, 0] = string.Empty;
                        }
                    }
                    sheet.Range[$"AA{compStartRow}:AA{compEndRow}"].Value2 = handleMatrix;
                }

                // 5. 批量组装汇总行 1 行 13 列数据矩阵 (A~M 列一次性单次 Range 批量写入)
                object[,] sumRowMatrix = new object[1, 13];
                sumRowMatrix[0, 0] = "=ROW()-ROW(A$6)"; // A 列: 序号 --硬编码--
                sumRowMatrix[0, 1] = safeBoxName; // B 列: 箱柜名称
                sumRowMatrix[0, 2] = string.IsNullOrWhiteSpace(cab.Header.Model) ? safeBoxName : cab.Header.Model; // C 列: 型号
                sumRowMatrix[0, 3] = string.Empty; // D 列: 图号
                sumRowMatrix[0, 4] = string.Empty; // E 列: 备注
                sumRowMatrix[0, 5] = cab.Header.Quantity > 0 ? cab.Header.Quantity : 1; // F 列: 数量
                sumRowMatrix[0, 6] = $"=H{tolsumRow - 1}"; // G 列: 单价公式 (指向单台合计行)
                sumRowMatrix[0, 7] = $"=F{sumRow}*G{sumRow}"; // H 列: 总价公式
                sumRowMatrix[0, 8] = string.Empty; // I 列
                sumRowMatrix[0, 9] = $"=K{tolsumRow}"; // J 列: 成本总价公式 (指向总计行)
                sumRowMatrix[0, 10] = $"=H{sumRow}-J{sumRow}"; // K 列: 毛利公式
                sumRowMatrix[0, 11] = $"=IF(H{sumRow}=0,0,K{sumRow}/H{sumRow})"; // L 列: 毛利率公式
                sumRowMatrix[0, 12] = installMode; // M 列: 安装方式备注
                // 单次 COM 批量写入汇总行
                sheet.Range[$"A{sumRow}:M{sumRow}"].Formula = sumRowMatrix;

                // 6. 写入明细信息行表头属性
                sheet.Cells[detRow, 2].Value2 = safeBoxName;
                sheet.Cells[detRow, 9].Value2 = installMode;
                if (cab.Header.MinMaxPoints != null && cab.Header.MinMaxPoints.Count > 0)
                {
                    // 记录图纸范围坐标
                    sheet.Cells[detRow, 27].Value2 = string.Join("-", cab.Header.MinMaxPoints);
                }

                // 7. 极致加速：直接注册 4 个定义名称 (零异常抛接，消灭 400 次 COM 异常)
                string sumNameTag = $"Cab_Sum_{cabinetK}";
                string detNameTag = $"Cab_Det_{cabinetK}";
                string subsumNameTag = $"Cab_Subsum_{cabinetK}";
                string tolsumNameTag = $"Cab_Tolsum_{cabinetK}";

                try { sheet.Names.Add(sumNameTag, $"='{sheetName}'!$A${sumRow}"); } catch { }
                try { sheet.Names.Add(detNameTag, $"='{sheetName}'!$A${detRow}"); } catch { }
                try { sheet.Names.Add(subsumNameTag, $"='{sheetName}'!$A${subsumRow}"); } catch { }
                try { sheet.Names.Add(tolsumNameTag, $"='{sheetName}'!$A${tolsumRow}"); } catch { }

                // 8. 建立双向超链接绑定 (规则 6)
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

                // 9. 刷新小计行自适应求和公式
                try
                {
                    sheet.Cells[subsumRow, 8].Formula = $"=ROUND(SUM(H{compStartRow}:INDEX(H:H,ROW()-1)),2)";
                    sheet.Cells[subsumRow, 11].Formula = $"=ROUND(SUM(K{compStartRow}:INDEX(K:K,ROW()-1)),2)";
                }
                catch { }

                return true;
            }
            catch (Exception ex)
            {
                // 记录单箱柜导出异常日志
                LogHelper.WriteLog($"ExportSingleCabinetOptimized 异常: {ex.Message}");
                return false;
            }
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

    }
}
