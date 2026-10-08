using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Forms;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：总价一键调整 (包含动态扫描、双向联动计算、报出系数调整及计费区比例系数重算)
    /// </summary>
    public static partial class ExcelServices
    {
        // 总价一键调整窗口静态单例引用 (保持非模态单例)
        private static TotalPriceAdjustForm? _totalPriceAdjustFormInstance;

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“总价一键调整”窗口 (非模态，可自由交互 Excel)
        /// </summary>
        public static void ShowTotalPriceAdjustDialog()
        {
            try
            {
                // 若窗体实例已存在且未释放
                if (_totalPriceAdjustFormInstance != null && !_totalPriceAdjustFormInstance.IsDisposed)
                {
                    // 若窗体最小化则恢复正常显示
                    if (_totalPriceAdjustFormInstance.WindowState == FormWindowState.Minimized)
                    {
                        // 还原窗体
                        _totalPriceAdjustFormInstance.WindowState = FormWindowState.Normal;
                    }
                    // 激活并置前窗口
                    _totalPriceAdjustFormInstance.BringToFront();
                    _totalPriceAdjustFormInstance.Activate();
                    return;
                }

                // 实例化全新的总价一键调整窗口
                _totalPriceAdjustFormInstance = new TotalPriceAdjustForm();
                // 绑定释放事件清空单例引用
                _totalPriceAdjustFormInstance.FormClosed += (s, e) => _totalPriceAdjustFormInstance = null;

                // 以非模态方式展示窗体，保持 Excel 处于可交互编辑状态
                _totalPriceAdjustFormInstance.Show();
            }
            catch (Exception ex)
            {
                // 记录日志并弹窗警示
                LogHelper.WriteLog($"[总价一键调整] 弹出窗口发生异常: {ex.Message}");
                MessageBox.Show($"弹出总价一键调整窗口失败: {ex.Message}", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 读取当前工作簿的总价一键调整初始化数据 (包含所有分类表、累计总价、总成本及动态提取的 Q 列分摊标签)
        /// </summary>
        /// <returns>包含分类表与标签聚合结果的初始化实体</returns>
        public static TotalPriceAdjustInitData GetTotalPriceAdjustInitData()
        {
            // 初始化返回实体包
            var initData = new TotalPriceAdjustInitData();

            try
            {
                // 获取当前活动 Excel Application 实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    // 标记失败消息
                    initData.Success = false;
                    initData.Message = "未检测到运行中的 Excel 应用程序实例";
                    return initData;
                }

                // 获取活动工作簿
                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null)
                {
                    // 标记无工作簿打开
                    initData.Success = false;
                    initData.Message = "当前未打开任何 Excel 工作簿";
                    return initData;
                }

                // 记录工程名称
                initData.ProjectName = Convert.ToString(activeWb.Name) ?? "未命名工程";

                // 字典用于按标签名称聚合各分类的累计金额与频次
                var tagAggDict = new Dictionary<string, (decimal amount, bool hasFormula, int occurrences, string itemType)>(StringComparer.OrdinalIgnoreCase);

                // 预置默认的【元件】标签 (元器件行 Q 列固定为元件)
                tagAggDict["元件"] = (0m, false, 0, "component");

                // 遍历活动工作簿中的所有工作表
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    // 获取工作表名称
                    string sheetName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;

                    // 排除系统辅助保留表 --硬编码--
                    if (string.Equals(sheetName, "项目信息", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sheetName, "材料分布表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sheetName, "元件汇总分布表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sheetName, "元件汇总表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sheetName, "元件汇总调价清单", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sheetName, "屏柜汇总表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sheetName, "屏柜分项表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sheetName, "元器件数据管理", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // 规则 8: 操作前强制自愈并校准当前分类表的 4 个定义名称
                    try
                    {
                        Tool.FixAndFillCabinetNamesForSheet(sheet);
                    }
                    catch { }

                    // 调用公共方法读取当前分类表的有效箱柜锚点字典 (规则 6)
                    var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);
                    if (validCabinets == null || validCabinets.Count == 0)
                    {
                        continue;
                    }

                    // 统计当前分类表的有效箱柜数量
                    int cabCount = validCabinets.Count;

                    // 规则 7: 单次 COM 调用抓取该分类表整表矩形数据矩阵到内存 (0 次频繁 COM 交互)
                    int maxUsedRow = 200;
                    try
                    {
                        dynamic uRange = sheet.UsedRange;
                        if (uRange != null)
                        {
                            int uRow = Convert.ToInt32(uRange.Row);
                            int uCount = Convert.ToInt32(uRange.Rows.Count);
                            maxUsedRow = Math.Max(uRow + uCount + 10, 200);
                        }
                    }
                    catch { }

                    // 一次性读取整表 A1:AD{maxUsedRow} 矩形矩阵 (值与公式)
                    dynamic sheetDataRange = sheet.Range[$"A1:AD{maxUsedRow}"];
                    object[,] sheetValMatrix = (object[,])sheetDataRange.Value2;
                    object[,] sheetFormulaMatrix = (object[,])sheetDataRange.Formula;
                    int maxArrRow = sheetValMatrix.GetLength(0);
                    int maxArrCol = sheetValMatrix.GetLength(1);

                    // 内存安全读取闭包辅助
                    object? GetVal(int r, int c)
                    {
                        if (r >= 1 && r <= maxArrRow && c >= 1 && c <= maxArrCol) return sheetValMatrix[r, c];
                        return null;
                    }
                    string GetFormulaStr(int r, int c)
                    {
                        if (r >= 1 && r <= maxArrRow && c >= 1 && c <= maxArrCol)
                        {
                            var f = sheetFormulaMatrix[r, c];
                            return f != null ? Convert.ToString(f)?.Trim() ?? string.Empty : string.Empty;
                        }
                        return string.Empty;
                    }

                    // 本分类表当前的累计销售额与成本额
                    decimal sheetSalesTotal = 0m;
                    decimal sheetCostTotal = 0m;

                    // 遍历该分类表下的每一个有效箱柜进行内存切片聚合
                    foreach (var cab in validCabinets)
                    {
                        // 缺少明细行锚点跳过
                        if (cab.Value.Det == null) continue;

                        int detRow = Convert.ToInt32(cab.Value.Det.Row);
                        // 小计行号
                        int subsumRow = cab.Value.Subsum != null ? Convert.ToInt32(cab.Value.Subsum.Row) : detRow + 27;
                        // 总计行号
                        int tolsumRow = cab.Value.Tolsum != null ? Convert.ToInt32(cab.Value.Tolsum.Row) : subsumRow + 8;
                        // 顶部汇总行号
                        int sumRow = cab.Value.Sum != null ? Convert.ToInt32(cab.Value.Sum.Row) : 0;

                        // 提取箱柜台数 (默认 1 台)
                        int cabQty = 1;
                        if (sumRow > 0)
                        {
                            object? qVal = GetVal(sumRow, 6) ?? GetVal(sumRow, 5);
                            if (qVal != null && int.TryParse(Convert.ToString(qVal), out int parsedQty) && parsedQty > 0)
                            {
                                cabQty = parsedQty;
                            }
                        }

                        // 累计该箱柜的销售总额 (优先从汇总行 H 列提取单柜总额，乘台数)
                        decimal cabSalesAmount = 0m;
                        if (sumRow > 0)
                        {
                            // 汇总行 H 列为该箱柜总价
                            object? sumH = GetVal(sumRow, 8);
                            if (sumH != null && decimal.TryParse(Convert.ToString(sumH), out decimal shVal))
                            {
                                cabSalesAmount = shVal;
                            }
                        }
                        // 若汇总行无有效数值，则回退从小计/总计行提取
                        if (cabSalesAmount <= 0)
                        {
                            object? tolH = GetVal(tolsumRow, 8) ?? GetVal(tolsumRow, 7);
                            if (tolH != null && decimal.TryParse(Convert.ToString(tolH), out decimal thVal))
                            {
                                cabSalesAmount = thVal * cabQty;
                            }
                        }
                        sheetSalesTotal += cabSalesAmount;

                        // 累计成本总额 (汇总行 K 列)
                        if (sumRow > 0)
                        {
                            object? sumK = GetVal(sumRow, 11) ?? GetVal(sumRow, 10);
                            if (sumK != null && decimal.TryParse(Convert.ToString(sumK), out decimal skVal))
                            {
                                sheetCostTotal += skVal;
                            }
                        }

                        // 1. 扫描元器件区域 (行 detRow + 2 至 subsumRow - 1)
                        int compStartRow = detRow + 2;
                        int compEndRow = subsumRow - 1;
                        if (compEndRow >= compStartRow)
                        {
                            for (int r = compStartRow; r <= compEndRow; r++)
                            {
                                // 提取元件名称与规格型号
                                string bName = Convert.ToString(GetVal(r, 2))?.Trim() ?? string.Empty;
                                string cSpec = Convert.ToString(GetVal(r, 3))?.Trim() ?? string.Empty;
                                // 提取 Q 列类别标记 (第 17 列)
                                string qCategory = Convert.ToString(GetVal(r, 17))?.Trim() ?? string.Empty;

                                // 只有名称、型号与 Q 列标记全为空时才跳过
                                if (string.IsNullOrWhiteSpace(bName) && string.IsNullOrWhiteSpace(cSpec) && string.IsNullOrWhiteSpace(qCategory))
                                {
                                    continue;
                                }

                                // 提取 H 列销售总价与 F 列数量
                                decimal itemTotal = ParseDecimalSafe(GetVal(r, 8));
                                if (itemTotal <= 0)
                                {
                                    // 回退尝试数量 * 单价
                                    decimal fQty = ParseDecimalSafe(GetVal(r, 6));
                                    if (fQty <= 0) fQty = 1m;
                                    decimal gPrice = ParseDecimalSafe(GetVal(r, 7));
                                    itemTotal = fQty * gPrice;
                                }

                                // 确定该行的归属类别：若 Q 列指定了非“元件”的类型，按该类型单独归类，否则归为核心【元件】
                                string compTag = (!string.IsNullOrWhiteSpace(qCategory) && !string.Equals(qCategory, "元件", StringComparison.OrdinalIgnoreCase))
                                    ? qCategory
                                    : "元件";

                                // 乘以箱柜台数累计到对应标签中
                                decimal compTotalWithCabinet = itemTotal * cabQty;
                                if (tagAggDict.TryGetValue(compTag, out var existingComp))
                                {
                                    tagAggDict[compTag] = (
                                        existingComp.amount + compTotalWithCabinet,
                                        existingComp.hasFormula,
                                        existingComp.occurrences + 1,
                                        compTag == "元件" ? "component" : "fee"
                                    );
                                }
                                else
                                {
                                    tagAggDict[compTag] = (compTotalWithCabinet, false, 1, "fee");
                                }
                            }
                        }

                        // 2. 扫描计费区域 (从 subsumRow 行开始向下扫描至 tolsumRow)
                        // 起始行设为 subsumRow，并自动跳过“小计”行，确保紧贴小计行的首项费用绝不漏扫
                        int feeStartRow = subsumRow > 0 ? subsumRow : (compEndRow + 1);
                        int feeEndRow = tolsumRow > feeStartRow ? tolsumRow : (feeStartRow + 25);
                        if (feeEndRow >= feeStartRow)
                        {
                            for (int r = feeStartRow; r <= feeEndRow; r++)
                            {
                                // 提取 A 列序号、B 列费用名称
                                string noVal = Convert.ToString(GetVal(r, 1))?.Trim() ?? string.Empty;
                                string feeName = Convert.ToString(GetVal(r, 2))?.Trim() ?? string.Empty;
                                // 提取 Q 列类别标记 (第 17 列)
                                string qCategory = Convert.ToString(GetVal(r, 17))?.Trim() ?? string.Empty;

                                // 严密识别总计/单台合计行，一旦遇到立即结束当前箱柜计费区扫描
                                bool isTolsumLine = noVal.Contains("总计") || feeName.Contains("总计") ||
                                                    noVal.Contains("合计") || feeName.Contains("合计") ||
                                                    noVal.Contains("单台合计") || feeName.Contains("单台合计");
                                if (isTolsumLine)
                                {
                                    break;
                                }

                                // 严密识别小计求和汇总行，跳过该行自身，防止将小计误当作计费项
                                bool isSubsumLine = noVal.Contains("小计") || feeName.Contains("小计");
                                if (isSubsumLine)
                                {
                                    continue;
                                }

                                // 跳过全空行 (名称与 Q 列均为空)
                                if (string.IsNullOrWhiteSpace(feeName) && string.IsNullOrWhiteSpace(qCategory))
                                {
                                    continue;
                                }

                                // 确定当前行的分摊标签 (优先级: Q列标记 -> B列费用名 -> 【其他费用】兜底)
                                string currentTag = string.Empty;
                                if (!string.IsNullOrWhiteSpace(qCategory))
                                {
                                    currentTag = qCategory;
                                }
                                else if (!string.IsNullOrWhiteSpace(feeName))
                                {
                                    currentTag = feeName;
                                }
                                else
                                {
                                    currentTag = "其他费用"; // --硬编码: 兜底标签名称--
                                }

                                // 提取当前费用行的金额 (优先取 H 列销售总价，次选取 G 列单价)
                                decimal feeAmount = ParseDecimalSafe(GetVal(r, 8));
                                if (feeAmount <= 0)
                                {
                                    feeAmount = ParseDecimalSafe(GetVal(r, 7));
                                }

                                // 检查该行是否包含比例公式 (例如 =ROUND(H20*0.08, 2) 或包含乘号 *)
                                string formH = GetFormulaStr(r, 8);
                                string formG = GetFormulaStr(r, 7);
                                bool rowHasRatioFormula = (formH.StartsWith("=") && formH.Contains("*")) ||
                                                          (formG.StartsWith("=") && formG.Contains("*"));

                                // 累计金额 (乘以箱柜台数)
                                decimal feeTotalWithCabinet = feeAmount * cabQty;
                                if (tagAggDict.TryGetValue(currentTag, out var existingTag))
                                {
                                    tagAggDict[currentTag] = (
                                        existingTag.amount + feeTotalWithCabinet,
                                        existingTag.hasFormula || rowHasRatioFormula,
                                        existingTag.occurrences + 1,
                                        "fee"
                                    );
                                }
                                else
                                {
                                    tagAggDict[currentTag] = (feeTotalWithCabinet, rowHasRatioFormula, 1, "fee");
                                }
                            }
                        }
                    }

                    // 记录该分类表项并默认选中
                    initData.CategorySheets.Add(new TotalPriceCategorySheetItem
                    {
                        SheetName = sheetName,
                        CabinetCount = cabCount,
                        SalesTotalPrice = sheetSalesTotal,
                        CostTotalPrice = sheetCostTotal,
                        IsSelected = true
                    });

                    // 累加整簿总金额
                    initData.TotalSalesPrice += sheetSalesTotal;
                    initData.TotalCostPrice += sheetCostTotal;
                }

                // 计算各标签在总金额中的原始占比，并构建前端展示列表
                decimal projectGrandTotal = initData.TotalSalesPrice > 0 ? initData.TotalSalesPrice : 1m;
                foreach (var kvp in tagAggDict)
                {
                    // 彻底排除包含小计、合计、总计等求和关键字的标签
                    if (kvp.Key.Contains("小计") || kvp.Key.Contains("合计") || kvp.Key.Contains("总计"))
                    {
                        continue;
                    }

                    // 只要在表格中出现过 (occurrences > 0) 或为核心元件标签，均展示在分摊列表中，允许用户分配调价分摊比例
                    if (kvp.Value.occurrences > 0 || kvp.Key == "元件")
                    {
                        // 计算占比百分比 (保留2位小数)
                        decimal ratioPercent = projectGrandTotal > 0 ? Math.Round((kvp.Value.amount / projectGrandTotal) * 100m, 2) : 0m;

                        // 默认设置：元器件默认分摊 1.0 (100%)，其余费用项默认分摊 0.0
                        bool isDefaultComponent = kvp.Key == "元件";
                        initData.Tags.Add(new TotalPriceAdjustTagItem
                        {
                            Tag = kvp.Key,
                            ItemType = kvp.Value.itemType,
                            CurrentAmount = kvp.Value.amount,
                            OriginalRatio = ratioPercent,
                            AllocationRatio = isDefaultComponent ? 1.0m : 0.0m,
                            IsSelected = isDefaultComponent,
                            HasRatioFormula = kvp.Value.hasFormula,
                            CabinetOccurrences = kvp.Value.occurrences
                        });
                    }
                }

                // 排序：元件排在首位，其余费用按金额从大到小降序排列
                initData.Tags = initData.Tags
                    .OrderByDescending(t => t.ItemType == "component")
                    .ThenByDescending(t => t.CurrentAmount)
                    .ToList();

                // 记录扫描完成详细日志供排查
                LogHelper.WriteLog($"[总价一键调整-V2.2] 扫描完成: 共检测到 {initData.Tags.Count} 个分摊项: {string.Join(", ", initData.Tags.Select(t => $"{t.Tag}({t.CurrentAmount:F2}元)"))}");

                // 标记操作成功并附带明确版本号
                initData.Success = true;
                initData.Message = $"扫描成功(V2.2)：共检测到 {initData.CategorySheets.Count} 个分类表，包含 {initData.Tags.Count} 个分摊类别项。";
            }
            catch (Exception ex)
            {
                // 异常捕获与日志记录
                initData.Success = false;
                initData.Message = $"读取调价初始化数据异常: {ex.Message}";
                LogHelper.WriteLog($"[总价一键调整] GetTotalPriceAdjustInitData 异常: {ex.Message}\r\n{ex.StackTrace}");
            }

            return initData;
        }

        /// <summary>
        /// 核心业务：执行总价一键调整 (按 Q 列动态分摊、更新元件 L 列报出系数、自适应更新计费区乘数比例公式)
        /// </summary>
        /// <param name="request">前端提交的调价请求参数载荷</param>
        /// <returns>调价执行结果实体</returns>
        public static TotalPriceAdjustResult ExecuteTotalPriceAdjust(TotalPriceAdjustRequest request)
        {
            var result = new TotalPriceAdjustResult();

            try
            {
                // 参数校验: 目标分类表不能为空
                if (request.SelectedSheets == null || request.SelectedSheets.Count == 0)
                {
                    result.Success = false;
                    result.Message = "请至少选择一个参与调价的分类工作表！";
                    return result;
                }

                // 参数校验: 分摊比例总和必须严格等于 1.0 (容差 0.001)
                decimal sumRatios = request.TagAllocations != null ? request.TagAllocations.Sum(t => t.Ratio) : 0m;
                if (Math.Abs(sumRatios - 1.0m) > 0.005m)
                {
                    result.Success = false;
                    result.Message = $"分摊比例总和必须严格等于 1.0 (当前总和为: {sumRatios:F3})！";
                    return result;
                }

                // 获取 Excel Application 实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    result.Success = false;
                    result.Message = "未检测到运行中的 Excel 应用程序实例";
                    return result;
                }

                // 获取活动工作簿
                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null)
                {
                    result.Success = false;
                    result.Message = "当前无活动工作簿";
                    return result;
                }

                // 另存为副本处理逻辑
                string? savedCopyPath = null;
                if (request.SaveAsCopy)
                {
                    try
                    {
                        // 计算副本文件保存路径
                        string originalPath = Convert.ToString(activeWb.FullName) ?? "";
                        string originalDir = Path.GetDirectoryName(originalPath) ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        string baseName = Path.GetFileNameWithoutExtension(originalPath);
                        string ext = Path.GetExtension(originalPath);
                        if (string.IsNullOrEmpty(ext)) ext = ".xlsx";

                        string copyFileName = string.IsNullOrWhiteSpace(request.CopyFileName)
                            ? $"{baseName}_总价调整_{DateTime.Now:yyyyMMddHHmmss}{ext}"
                            : (request.CopyFileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase) ? request.CopyFileName : $"{request.CopyFileName}{ext}");

                        savedCopyPath = Path.Combine(originalDir, copyFileName);
                        // 保存副本文件至磁盘 --硬编码--
                        activeWb.SaveCopyAs(savedCopyPath);
                        result.SavedCopyFilePath = savedCopyPath;
                    }
                    catch (Exception exCopy)
                    {
                        LogHelper.WriteLog($"[总价一键调整] 另存为副本异常: {exCopy.Message}");
                    }
                }

                // 1. 重新扫描所选分类表的基准数据与各标签现有总额
                decimal selectedSheetsCurrentTotal = 0m;
                decimal selectedSheetsCostTotal = 0m;
                var tagAmounts = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

                // 快速收集选定表的有效箱柜
                var targetSheetsList = new List<dynamic>();
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    string sName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;
                    if (request.SelectedSheets.Contains(sName, StringComparer.OrdinalIgnoreCase))
                    {
                        // 规则 8: 操作前强制修复定义名称
                        Tool.FixAndFillCabinetNamesForSheet(sheet);
                        targetSheetsList.Add(sheet);
                    }
                }

                // 统计选定分类表初始总价与各标签初始总金额
                foreach (var sheet in targetSheetsList)
                {
                    var validCabs = Tool.GetSheetValidCabinets(sheet, activeWb);
                    if (validCabs == null || validCabs.Count == 0) continue;

                    // 读取该表 UsedRange 矩形大矩阵 (规则 7)
                    int maxUsedRow = 200;
                    try
                    {
                        dynamic uRange = sheet.UsedRange;
                        if (uRange != null)
                        {
                            int uRow = Convert.ToInt32(uRange.Row);
                            int uCount = Convert.ToInt32(uRange.Rows.Count);
                            maxUsedRow = Math.Max(uRow + uCount + 10, 200);
                        }
                    }
                    catch { }

                    dynamic sheetDataRange = sheet.Range[$"A1:AD{maxUsedRow}"];
                    object[,] sheetValMatrix = (object[,])sheetDataRange.Value2;
                    int maxArrRow = sheetValMatrix.GetLength(0);
                    int maxArrCol = sheetValMatrix.GetLength(1);

                    object? GetVal(int r, int c)
                    {
                        if (r >= 1 && r <= maxArrRow && c >= 1 && c <= maxArrCol) return sheetValMatrix[r, c];
                        return null;
                    }

                    foreach (var cab in validCabs)
                    {
                        if (cab.Value.Det == null) continue;
                        int detRow = Convert.ToInt32(cab.Value.Det.Row);
                        int subsumRow = cab.Value.Subsum != null ? Convert.ToInt32(cab.Value.Subsum.Row) : detRow + 27;
                        int tolsumRow = cab.Value.Tolsum != null ? Convert.ToInt32(cab.Value.Tolsum.Row) : subsumRow + 8;
                        int sumRow = cab.Value.Sum != null ? Convert.ToInt32(cab.Value.Sum.Row) : 0;

                        int cabQty = 1;
                        if (sumRow > 0)
                        {
                            object? qVal = GetVal(sumRow, 6) ?? GetVal(sumRow, 5);
                            if (qVal != null && int.TryParse(Convert.ToString(qVal), out int pq) && pq > 0) cabQty = pq;
                        }

                        // 汇总行 H 列销售总价
                        decimal cabSales = 0m;
                        if (sumRow > 0)
                        {
                            object? shVal = GetVal(sumRow, 8);
                            if (shVal != null && decimal.TryParse(Convert.ToString(shVal), out decimal csv)) cabSales = csv;
                        }
                        if (cabSales <= 0)
                        {
                            object? thVal = GetVal(tolsumRow, 8) ?? GetVal(tolsumRow, 7);
                            if (thVal != null && decimal.TryParse(Convert.ToString(thVal), out decimal ctv)) cabSales = ctv * cabQty;
                        }
                        selectedSheetsCurrentTotal += cabSales;

                        // 汇总行 K 列成本
                        if (sumRow > 0)
                        {
                            object? skVal = GetVal(sumRow, 11);
                            if (skVal != null && decimal.TryParse(Convert.ToString(skVal), out decimal csk)) selectedSheetsCostTotal += csk;
                        }

                        // 统计元器件
                        int compStartRow = detRow + 2;
                        int compEndRow = subsumRow - 1;
                        if (compEndRow >= compStartRow)
                        {
                            for (int r = compStartRow; r <= compEndRow; r++)
                            {
                                string bName = Convert.ToString(GetVal(r, 2))?.Trim() ?? string.Empty;
                                string cSpec = Convert.ToString(GetVal(r, 3))?.Trim() ?? string.Empty;
                                string qCat = Convert.ToString(GetVal(r, 17))?.Trim() ?? string.Empty;
                                if (string.IsNullOrWhiteSpace(bName) && string.IsNullOrWhiteSpace(cSpec) && string.IsNullOrWhiteSpace(qCat)) continue;

                                decimal it = ParseDecimalSafe(GetVal(r, 8));
                                if (it <= 0)
                                {
                                    decimal f = ParseDecimalSafe(GetVal(r, 6));
                                    if (f <= 0) f = 1m;
                                    decimal g = ParseDecimalSafe(GetVal(r, 7));
                                    it = f * g;
                                }

                                string compTag = (!string.IsNullOrWhiteSpace(qCat) && !string.Equals(qCat, "元件", StringComparison.OrdinalIgnoreCase))
                                    ? qCat
                                    : "元件";

                                decimal totalComp = it * cabQty;
                                tagAmounts[compTag] = (tagAmounts.ContainsKey(compTag) ? tagAmounts[compTag] : 0m) + totalComp;
                            }
                        }

                        // 统计计费区
                        int feeStartRow = subsumRow > 0 ? subsumRow : (compEndRow + 1);
                        int feeEndRow = tolsumRow > feeStartRow ? tolsumRow : (feeStartRow + 25);
                        if (feeEndRow >= feeStartRow)
                        {
                            for (int r = feeStartRow; r <= feeEndRow; r++)
                            {
                                string noVal = Convert.ToString(GetVal(r, 1))?.Trim() ?? string.Empty;
                                string fName = Convert.ToString(GetVal(r, 2))?.Trim() ?? string.Empty;
                                string qCat = Convert.ToString(GetVal(r, 17))?.Trim() ?? string.Empty;

                                // 拦截总计/单台合计行
                                if (noVal.Contains("总计") || fName.Contains("总计") ||
                                    noVal.Contains("合计") || fName.Contains("合计") ||
                                    noVal.Contains("单台合计") || fName.Contains("单台合计"))
                                {
                                    break;
                                }

                                // 拦截小计行
                                if (noVal.Contains("小计") || fName.Contains("小计"))
                                {
                                    continue;
                                }

                                if (string.IsNullOrWhiteSpace(fName) && string.IsNullOrWhiteSpace(qCat)) continue;

                                string tag = !string.IsNullOrWhiteSpace(qCat) ? qCat : (!string.IsNullOrWhiteSpace(fName) ? fName : "其他费用");
                                // 排除求和类标签
                                if (tag.Contains("小计") || tag.Contains("合计") || tag.Contains("总计")) continue;

                                decimal fAmt = ParseDecimalSafe(GetVal(r, 8));
                                if (fAmt <= 0) fAmt = ParseDecimalSafe(GetVal(r, 7));

                                decimal totalFee = fAmt * cabQty;
                                tagAmounts[tag] = (tagAmounts.ContainsKey(tag) ? tagAmounts[tag] : 0m) + totalFee;
                            }
                        }
                    }
                }

                // 安全校验: 目标总价不能低于成本总价
                if (request.TargetPrice <= selectedSheetsCostTotal && selectedSheetsCostTotal > 0)
                {
                    result.Success = false;
                    result.Message = $"目标总价 ({request.TargetPrice:N2}元) 必须高于当前成本总价 ({selectedSheetsCostTotal:N2}元)，防止亏本报价！";
                    return result;
                }

                // 计算整项调价总差额 (新目标总价 - 原选定总价)
                decimal totalDiff = request.TargetPrice - selectedSheetsCurrentTotal;
                result.OldTotalAmount = selectedSheetsCurrentTotal;

                // 2. 计算各个分摊标签的加权缩放比例 kTag
                var tagScalingFactors = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
                foreach (var alloc in request.TagAllocations)
                {
                    if (alloc.Ratio <= 0) continue;

                    // 计算该标签应当承担的差额
                    decimal diffForTag = totalDiff * alloc.Ratio;
                    // 读取该标签的原累计金额
                    decimal originTagAmt = tagAmounts.ContainsKey(alloc.Tag) ? tagAmounts[alloc.Tag] : 0m;

                    // 计算缩放倍率: k = (原金额 + 承担差额) / 原金额
                    if (originTagAmt > 0)
                    {
                        decimal k = (originTagAmt + diffForTag) / originTagAmt;
                        tagScalingFactors[alloc.Tag] = k;
                    }
                    else
                    {
                        // 若原金额为 0 则不缩放
                        tagScalingFactors[alloc.Tag] = 1.0m;
                    }
                }

                // 准备整簿撤销快照切片列表 (支持 Ctrl+Z 100% 秒级无损撤销)
                var undoSlices = new List<RangeDeltaSlice>();

                // 计数器
                int updatedSheetCount = 0;
                int updatedCabCount = 0;
                int updatedCompCount = 0;
                int updatedFeeCount = 0;

                // 暂停屏幕更新加速执行 (规则 7 & 启发式经验)
                app.ScreenUpdating = false;
                app.Calculation = -4135; // xlCalculationManual 手动重算加速

                try
                {
                    // 3. 逐表逐箱柜执行调价回写
                    foreach (var sheet in targetSheetsList)
                    {
                        string sheetName = Convert.ToString(sheet.Name) ?? string.Empty;
                        var validCabs = Tool.GetSheetValidCabinets(sheet, activeWb);
                        if (validCabs == null || validCabs.Count == 0) continue;

                        updatedSheetCount++;

                        foreach (var cab in validCabs)
                        {
                            if (cab.Value.Det == null) continue;
                            int detRow = Convert.ToInt32(cab.Value.Det.Row);
                            int subsumRow = cab.Value.Subsum != null ? Convert.ToInt32(cab.Value.Subsum.Row) : detRow + 27;
                            int tolsumRow = cab.Value.Tolsum != null ? Convert.ToInt32(cab.Value.Tolsum.Row) : subsumRow + 8;

                            updatedCabCount++;

                            // (A) 处理元器件区域: 调整 L 列加价/报出系数
                            int compStartRow = detRow + 2;
                            int compEndRow = subsumRow - 1;
                            if (compEndRow >= compStartRow && tagScalingFactors.TryGetValue("元件", out decimal kComp) && kComp > 0)
                            {
                                int compRows = compEndRow - compStartRow + 1;
                                // 抓取元器件 B、C、L 列进行内存矩阵读写
                                dynamic bRange = sheet.Range[$"B{compStartRow}:B{compEndRow}"];
                                dynamic cRange = sheet.Range[$"C{compStartRow}:C{compEndRow}"];
                                dynamic lRange = sheet.Range[$"L{compStartRow}:L{compEndRow}"];

                                object[,] bVals = (object[,])bRange.Value2;
                                object[,] cVals = (object[,])cRange.Value2;
                                object[,] lVals = (object[,])lRange.Value2;
                                object[,] lFormulas = (object[,])lRange.Formula;

                                // 构建撤销快照切片 (记录修改前的 L 列)
                                var lSlice = new RangeDeltaSlice
                                {
                                    SheetName = sheetName,
                                    RangeAddress = $"L{compStartRow}:L{compEndRow}",
                                    OldValues = (object[,])lVals.Clone(),
                                    OldFormulas = (object[,])lFormulas.Clone()
                                };

                                object[,] newLVals = (object[,])lVals.Clone();

                                for (int r = 1; r <= compRows; r++)
                                {
                                    string b = Convert.ToString(bVals[r, 1])?.Trim() ?? string.Empty;
                                    string c = Convert.ToString(cVals[r, 1])?.Trim() ?? string.Empty;
                                    // 仅对有效器件行调整系数
                                    if (string.IsNullOrWhiteSpace(b) && string.IsNullOrWhiteSpace(c))
                                    {
                                        continue;
                                    }

                                    // 读取当前行 L 列报出系数值 (缺省默认为 1.0)
                                    decimal currentL = 1.0m;
                                    object? lVal = lVals[r, 1];
                                    if (lVal != null && decimal.TryParse(Convert.ToString(lVal), out decimal parsedL) && parsedL > 0)
                                    {
                                        currentL = parsedL;
                                    }

                                    // 计算并保留 4 位精度报出系数: newL = oldL * kComp
                                    decimal newL = Math.Round(currentL * kComp, 4);
                                    newLVals[r, 1] = (double)newL;
                                    updatedCompCount++;
                                }

                                // 单次 COM 调用批量回写 L 列数值
                                lRange.Value2 = newLVals;

                                lSlice.NewValues = newLVals;
                                undoSlices.Add(lSlice);
                            }

                            // (B) 处理计费区域: 更新比例系数公式或纯固定数值
                            int feeStartRow = subsumRow + 1;
                            int feeEndRow = tolsumRow - 1;
                            if (feeEndRow >= feeStartRow)
                            {
                                int feeRows = feeEndRow - feeStartRow + 1;
                                dynamic feeRange = sheet.Range[$"A{feeStartRow}:Q{feeEndRow}"];
                                object[,] feeVals = (object[,])feeRange.Value2;
                                object[,] feeFormulas = (object[,])feeRange.Formula;

                                // 记录计费区域撤销快照
                                var feeSlice = new RangeDeltaSlice
                                {
                                    SheetName = sheetName,
                                    RangeAddress = $"A{feeStartRow}:Q{feeEndRow}",
                                    OldValues = (object[,])feeVals.Clone(),
                                    OldFormulas = (object[,])feeFormulas.Clone()
                                };

                                object[,] newFeeVals = (object[,])feeVals.Clone();
                                object[,] newFeeFormulas = (object[,])feeFormulas.Clone();

                                bool feeModified = false;

                                for (int r = 1; r <= feeRows; r++)
                                {
                                    int physicalRow = feeStartRow + r - 1;
                                    string fName = Convert.ToString(feeVals[r, 2])?.Trim() ?? string.Empty;
                                    string qCat = Convert.ToString(feeVals[r, 17])?.Trim() ?? string.Empty;
                                    if (string.IsNullOrWhiteSpace(fName) && string.IsNullOrWhiteSpace(qCat)) continue;

                                    string tag = !string.IsNullOrWhiteSpace(qCat) ? qCat : (!string.IsNullOrWhiteSpace(fName) ? fName : "其他费用");

                                    // 若该标签具有缩放系数
                                    if (tagScalingFactors.TryGetValue(tag, out decimal kFee) && kFee > 0 && Math.Abs(kFee - 1.0m) > 0.0001m)
                                    {
                                        // 检查 H 列 (第 8 列) 或 G 列 (第 7 列) 公式
                                        string formH = Convert.ToString(feeFormulas[r, 8])?.Trim() ?? string.Empty;
                                        string formG = Convert.ToString(feeFormulas[r, 7])?.Trim() ?? string.Empty;

                                        bool formulaHandled = false;

                                        // 1. 尝试识别并更新 H 列中的乘数比例系数
                                        if (formH.StartsWith("=") && formH.Contains("*"))
                                        {
                                            string updatedFormH = UpdateFormulaMultiplier(formH, kFee);
                                            if (updatedFormH != formH)
                                            {
                                                newFeeFormulas[r, 8] = updatedFormH;
                                                formulaHandled = true;
                                                feeModified = true;
                                                updatedFeeCount++;
                                            }
                                        }

                                        // 2. 尝试识别并更新 G 列中的乘数比例系数
                                        if (formG.StartsWith("=") && formG.Contains("*"))
                                        {
                                            string updatedFormG = UpdateFormulaMultiplier(formG, kFee);
                                            if (updatedFormG != formG)
                                            {
                                                newFeeFormulas[r, 7] = updatedFormG;
                                                formulaHandled = true;
                                                feeModified = true;
                                                updatedFeeCount++;
                                            }
                                        }

                                        // 3. 若无公式或公式为单纯加减/纯数字，则直接调整数值
                                        if (!formulaHandled)
                                        {
                                            object? currentNumVal = feeVals[r, 8] ?? feeVals[r, 7];
                                            if (currentNumVal != null && decimal.TryParse(Convert.ToString(currentNumVal), out decimal currentNum) && currentNum > 0)
                                            {
                                                decimal newNum = Math.Round(currentNum * kFee, 2);
                                                // 覆盖数值到 H 列或 G 列
                                                if (feeVals[r, 8] != null)
                                                {
                                                    newFeeVals[r, 8] = (double)newNum;
                                                    newFeeFormulas[r, 8] = (double)newNum;
                                                }
                                                else
                                                {
                                                    newFeeVals[r, 7] = (double)newNum;
                                                    newFeeFormulas[r, 7] = (double)newNum;
                                                }
                                                feeModified = true;
                                                updatedFeeCount++;
                                            }
                                        }
                                    }
                                }

                                if (feeModified)
                                {
                                    // 批量写回计费区域
                                    feeRange.Formula = newFeeFormulas;
                                    feeSlice.NewFormulas = newFeeFormulas;
                                    feeSlice.NewValues = newFeeVals;
                                    undoSlices.Add(feeSlice);
                                }
                            }
                        }
                    }
                }
                finally
                {
                    // 恢复 Excel 计算与渲染状态
                    app.Calculation = -4105; // xlCalculationAutomatic 恢复自动计算
                    app.Calculate();        // 立即触发全局重算
                    app.ScreenUpdating = true;
                }

                // 4. 重算并提取调价后的最新工程总价
                decimal newGrandTotal = 0m;
                foreach (var sheet in targetSheetsList)
                {
                    var validCabs = Tool.GetSheetValidCabinets(sheet, activeWb);
                    if (validCabs == null || validCabs.Count == 0) continue;

                    foreach (var cab in validCabs)
                    {
                        if (cab.Value.Sum != null)
                        {
                            int sumR = Convert.ToInt32(cab.Value.Sum.Row);
                            object? h = sheet.Cells[sumR, 8].Value2;
                            if (h != null && decimal.TryParse(Convert.ToString(h), out decimal v)) newGrandTotal += v;
                        }
                    }
                }

                result.NewTotalAmount = newGrandTotal > 0 ? newGrandTotal : request.TargetPrice;
                result.DifferenceAmount = result.NewTotalAmount - result.OldTotalAmount;
                result.UpdatedSheetsCount = updatedSheetCount;
                result.UpdatedCabinetsCount = updatedCabCount;
                result.UpdatedComponentCount = updatedCompCount;
                result.UpdatedFeeItemCount = updatedFeeCount;

                // 5. 将本次调价完整切片压入 Undo 栈，支持一键撤销
                if (undoSlices.Count > 0)
                {
                    var adjustCmd = new RangeDeltaCommand($"总价一键调整 (目标: {request.TargetPrice:N2}元)", undoSlices);
                    UndoRedoManager.Instance.PushCommand(adjustCmd);
                }

                // 标记操作成功
                result.Success = true;
                result.Message = $"总价一键调整执行成功！共同步 {updatedSheetCount} 个分类表、{updatedCabCount} 台箱柜、{updatedCompCount} 项元器件报出系数与 {updatedFeeCount} 项费用公式/数值。新总价为 {result.NewTotalAmount:N2} 元 (可通过功能区【撤销】随时撤回)。";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"执行总价调整异常: {ex.Message}";
                LogHelper.WriteLog($"[总价一键调整] ExecuteTotalPriceAdjust 异常: {ex.Message}\r\n{ex.StackTrace}");
            }

            return result;
        }

        /// <summary>
        /// 辅助正则匹配替换公式中的乘数比例系数 (如 *0.08 或 *8% 替换为新的系数)
        /// </summary>
        /// <param name="originalFormula">原始 Excel 公式字符串</param>
        /// <param name="scalingFactor">缩放比例倍率</param>
        /// <returns>替换更新乘数后的新公式字符串</returns>
        private static string UpdateFormulaMultiplier(string originalFormula, decimal scalingFactor)
        {
            if (string.IsNullOrWhiteSpace(originalFormula) || scalingFactor <= 0) return originalFormula;

            // 正则匹配乘号后的比例系数常数: 例如 * 0.08、*8%、*0.05
            var regex = new Regex(@"(?<=\*\s*)(\d+(?:\.\d+)?%?|\.\d+)");
            var match = regex.Match(originalFormula);

            if (match.Success)
            {
                string matchedStr = match.Value;
                bool isPercent = matchedStr.EndsWith("%");

                decimal rawRate = 0m;
                string cleanNum = isPercent ? matchedStr.TrimEnd('%') : matchedStr;

                if (decimal.TryParse(cleanNum, out decimal parsedNum))
                {
                    // 若带百分号，计算实际小数
                    rawRate = isPercent ? parsedNum / 100m : parsedNum;

                    // 计算调整后的新比例系数
                    decimal newRate = rawRate * scalingFactor;

                    // 重新格式化乘数常数字符串
                    string newMultiplierStr;
                    if (isPercent)
                    {
                        // 保留 2 位小数百分比格式 (例如 7.2%)
                        decimal newPercentVal = Math.Round(newRate * 100m, 2);
                        newMultiplierStr = $"{newPercentVal}%";
                    }
                    else
                    {
                        // 保留 4 位小数浮点格式 (例如 0.072)
                        decimal newDecimalVal = Math.Round(newRate, 4);
                        newMultiplierStr = newDecimalVal.ToString("G");
                    }

                    // 仅替换首个匹配到的比例乘数
                    return regex.Replace(originalFormula, newMultiplierStr, 1);
                }
            }

            return originalFormula;
        }

        /// <summary>
        /// 安全解析单元格可能包含的浮点数/货币格式数值，去除货币符号与千分位
        /// </summary>
        /// <param name="val">原始单元格对象</param>
        /// <returns>解析后的十进制金额</returns>
        private static decimal ParseDecimalSafe(object? val)
        {
            // 空对象直接返回 0
            if (val == null) return 0m;
            // 直接十进制数强转
            if (val is decimal dec) return dec;
            // 浮点与整数兼容转换
            if (val is double d) return (decimal)d;
            if (val is float f) return (decimal)f;
            if (val is int i) return (decimal)i;
            if (val is long l) return (decimal)l;
            // 提取字符串清洗货币符号与千分位
            string s = Convert.ToString(val)?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(s)) return 0m;
            // 剥离人民币、美元符号与逗号
            s = s.Replace("￥", "").Replace("¥", "").Replace("$", "").Replace(",", "").Trim();
            // 按照通用文化格式解析数值
            if (decimal.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal res))
            {
                return res;
            }
            return 0m;
        }
    }
}
