using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelAddInDemo.Models;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 服务公共分部类：报价准确性智能核验与自愈服务
    /// 遵循规则 6、7、8，保障成套工程报价零漏项、零公式损毁、同物同价与 CAD 联动参数健全
    /// </summary>
    public static partial class ExcelServices
    {
        // 报价智能核验窗口单例引用
        private static Forms.QuotationCheckForm? _quotationCheckForm;

        /// <summary>
        /// 弹出报价智能核验与自愈向导窗体 (非模态，保持 Excel 可交互)
        /// </summary>
        public static void ShowQuotationCheckDialog()
        {
            try
            {
                // 以非模态方式唤起向导，允许用户与 Excel 表格实时并排对照
                ShowModelessForm(ref _quotationCheckForm, () => new Forms.QuotationCheckForm());
            }
            catch (Exception ex)
            {
                // 记录错误日志防止异常外溢
                LogHelper.WriteLog($"弹出报价核验窗口失败: {ex.Message}");
            }
        }

        // 匹配 Excel 常见公式错误符号的正则表达式
        private static readonly Regex FormulaErrorRegex = new Regex(@"#(REF!|VALUE!|DIV/0!|NAME\?|N/A|NUM!|NULL!)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// 全量执行当前工作簿各分类表的报价准确性智能体检
        /// </summary>
        /// <param name="wb">活动工作簿句柄 (为 null 时自动提取全局活动工作簿)</param>
        /// <returns>报价核验诊断报告与冲突分组数据</returns>
        public static QuotationCheckResult ExecuteQuotationAudit(Workbook? wb = null)
        {
            // 初始化返回结果对象
            var result = new QuotationCheckResult();

            try
            {
                // 获取 Excel 宿主全局 Application 对象
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                // 若宿主未就绪则直接终止返回错误
                if (app == null)
                {
                    result.Success = false;
                    result.Message = "Excel 宿主 Application 对象未就绪";
                    return result;
                }

                // 提取目标工作簿句柄 (若未传则取当前活动工作簿)
                Workbook? targetWb = wb ?? app.ActiveWorkbook;
                // 校验工作簿有效性
                if (targetWb == null)
                {
                    result.Success = false;
                    result.Message = "未检测到已打开的有效 Excel 工作簿";
                    return result;
                }

                // 规则 8：刷新并获取当前工程中所有合法的分类表工作表名称白名单
                List<string> categorySheets = Tool.GetProjectCategorySheetNames(targetWb, forceRefresh: true).ToList();
                // 若未探测到任何分类表工作表
                if (categorySheets == null || categorySheets.Count == 0)
                {
                    result.Success = false;
                    result.Message = "当前工作簿未检测到【分类表】！请确认是否已创建成套分类工作表。";
                    return result;
                }

                // 记录扫描涉及的所有分类表
                result.SheetNames = new List<string>(categorySheets);

                // 暂存全项目提取到的所有有效元器件记录，用于同型号聚类冲突分析
                var allComponents = new List<AuditedComponentItem>();

                // 循环遍历每一个分类工作表进行深度检查
                foreach (string sheetName in categorySheets)
                {
                    // 安全提取 dynamic 工作表对象 (避免 COM 强转失败)
                    dynamic? sheet = null;
                    try { sheet = targetWb.Worksheets[sheetName]; } catch { }
                    // 若工作表对象为空则跳过
                    if (sheet == null) continue;

                    // 规则 8：在操作前确保规则 6 定义名称完整性健全自愈
                    try { Tool.FixAndFillCabinetNamesForSheet(sheet); } catch { }

                    // 获取当前工作表中所有有效的箱柜锚点数据 (规则 6 架构)
                    var cabinets = Tool.GetSheetValidCabinets((object)sheet, (object)targetWb);
                    // 若当前分类表无箱柜则跳过
                    if (cabinets == null || cabinets.Count == 0) continue;

                    // 累计统计箱柜总数
                    result.Summary.TotalCabinets += cabinets.Count;

                    // 获取工作表使用区域的最大行号与列号，规则 7 一次性读取数据
                    dynamic usedRange = sheet.UsedRange;
                    // 一次性读取全部单元格的值矩阵至内存 (规则 7)
                    object[,]? valuesMatrix = null;
                    // 一次性读取全部单元格的公式矩阵至内存 (规则 7)
                    object[,]? formulasMatrix = null;

                    try
                    {
                        // 批量转入内存二维数组
                        valuesMatrix = usedRange.Value2 as object[,];
                        formulasMatrix = usedRange.Formula as object[,];
                    }
                    catch { }

                    // 获取使用区域起始行与起始列偏移行
                    int usedRowOffset = usedRange.Row;
                    int usedColOffset = usedRange.Column;
                    int maxRows = valuesMatrix != null ? valuesMatrix.GetLength(0) : 0;
                    int maxCols = valuesMatrix != null ? valuesMatrix.GetLength(1) : 0;

                    // 逐台箱柜进行元器件区域与计费区域扫描
                    foreach (var kvp in cabinets)
                    {
                        // 箱柜序号索引
                        int cabIndex = kvp.Key;
                        // 箱柜锚点实体模型
                        var cab = kvp.Value;

                        // 安全提取箱柜物理行号 (规则 6 架构)
                        int sumRow = cab.Sum != null ? Convert.ToInt32(cab.Sum.Row) : 0;
                        int detRow = cab.Det != null ? Convert.ToInt32(cab.Det.Row) : 0;
                        int subsumRow = cab.Subsum != null ? Convert.ToInt32(cab.Subsum.Row) : 0;
                        int tolsumRow = cab.Tolsum != null ? Convert.ToInt32(cab.Tolsum.Row) : 0;

                        // 提取箱柜名称 (明细表箱柜行 B 列)
                        string cabName = string.Empty;
                        try { cabName = Convert.ToString(sheet.Cells[detRow, 2].Value2)?.Trim() ?? $"箱柜{cabIndex}"; }
                        catch { cabName = $"箱柜{cabIndex}"; }

                        // 规则 6: 元器件起始行为 detRow + 2，终止行为 subsumRow - 1
                        int compStartRow = detRow + 2;
                        int compEndRow = subsumRow - 1;

                        // 1. 扫描元器件区域各行数据
                        if (compEndRow >= compStartRow)
                        {
                            for (int r = compStartRow; r <= compEndRow; r++)
                            {
                                // 计算在内存大数组中的相对二维行索引
                                int arrRow = r - usedRowOffset + 1;
                                // 数组越界安全校验
                                if (arrRow < 1 || arrRow > maxRows) continue;

                                // 提取 B 列 (元件名称) 与 C 列 (型号规格)
                                string name = GetMatrixString(valuesMatrix, arrRow, 2 - usedColOffset + 1);
                                string model = GetMatrixString(valuesMatrix, arrRow, 3 - usedColOffset + 1);

                                // 若 B 列名称与 C 列型号均为空，则视为空行直接跳过
                                if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(model))
                                {
                                    continue;
                                }

                                // 累加有效元器件计数
                                result.Summary.TotalComponents++;

                                // 提取各项核心属性
                                string brand = GetMatrixString(valuesMatrix, arrRow, 4 - usedColOffset + 1); // D列 品牌
                                string unit = GetMatrixString(valuesMatrix, arrRow, 5 - usedColOffset + 1);  // E列 单位

                                // 提取 F 列 数量
                                decimal qty = GetMatrixDecimal(valuesMatrix, arrRow, 6 - usedColOffset + 1);
                                object? rawQtyObj = GetMatrixObject(valuesMatrix, arrRow, 6 - usedColOffset + 1);

                                // 提取 G 列 销售单价 (值与公式)
                                decimal saleUnitPrice = GetMatrixDecimal(valuesMatrix, arrRow, 7 - usedColOffset + 1);
                                string formulaG = GetMatrixFormula(formulasMatrix, arrRow, 7 - usedColOffset + 1);

                                // 提取 H 列 销售总价 (值与公式)
                                decimal saleTotalPrice = GetMatrixDecimal(valuesMatrix, arrRow, 8 - usedColOffset + 1);
                                string formulaH = GetMatrixFormula(formulasMatrix, arrRow, 8 - usedColOffset + 1);

                                // 提取 J 列 成本单价 (值与公式)
                                decimal costUnitPrice = GetMatrixDecimal(valuesMatrix, arrRow, 10 - usedColOffset + 1);
                                string formulaJ = GetMatrixFormula(formulasMatrix, arrRow, 10 - usedColOffset + 1);

                                // 提取 K 列 成本总价 (值与公式)
                                decimal costTotalPrice = GetMatrixDecimal(valuesMatrix, arrRow, 11 - usedColOffset + 1);
                                string formulaK = GetMatrixFormula(formulasMatrix, arrRow, 11 - usedColOffset + 1);

                                // 提取 L 列 报价加价系数
                                decimal markupFactor = GetMatrixDecimal(valuesMatrix, arrRow, 12 - usedColOffset + 1);
                                if (markupFactor <= 0) markupFactor = 1.0m;

                                // 提取 M 列 面价/表价 (值与公式)
                                decimal basePrice = GetMatrixDecimal(valuesMatrix, arrRow, 13 - usedColOffset + 1);
                                object? rawBasePriceObj = GetMatrixObject(valuesMatrix, arrRow, 13 - usedColOffset + 1);
                                string formulaM = GetMatrixFormula(formulasMatrix, arrRow, 13 - usedColOffset + 1);

                                // 提取 N 列 采购折扣 (值与公式)
                                decimal discount = GetMatrixDecimal(valuesMatrix, arrRow, 14 - usedColOffset + 1);
                                object? rawDiscountObj = GetMatrixObject(valuesMatrix, arrRow, 14 - usedColOffset + 1);
                                string formulaN = GetMatrixFormula(formulasMatrix, arrRow, 14 - usedColOffset + 1);

                                // 提取 AA 列 图纸名 / 图块名称 (第 27 列，用户核心要求检查！)
                                string blockDwg = GetMatrixString(valuesMatrix, arrRow, 27 - usedColOffset + 1);
                                // 提取 AB 列 目录分类名 / 扩展参数2 (第 28 列，用户核心要求检查！)
                                string blockDir = GetMatrixString(valuesMatrix, arrRow, 28 - usedColOffset + 1);

                                // 构建本行审计条目上下文
                                var auditedComp = new AuditedComponentItem
                                {
                                    SheetName = sheetName,
                                    CabinetName = cabName,
                                    CabinetIndex = cabIndex,
                                    Row = r,
                                    ComponentName = name,
                                    Model = model,
                                    Brand = brand,
                                    Quantity = qty,
                                    SaleUnitPrice = saleUnitPrice,
                                    SaleTotalPrice = saleTotalPrice,
                                    CostUnitPrice = costUnitPrice,
                                    MarkupFactor = markupFactor,
                                    BasePrice = basePrice,
                                    Discount = discount,
                                    BlockDwg = blockDwg,
                                    BlockDir = blockDir,
                                    FormulaG = formulaG,
                                    FormulaH = formulaH,
                                    FormulaJ = formulaJ,
                                    FormulaK = formulaK
                                };
                                allComponents.Add(auditedComp);

                                // ----------------------------------------------------
                                // 维度 A: AA 列与 AB 列空值检测 (CAD 图纸与目录联动完整性)
                                // ----------------------------------------------------
                                bool isDwgEmpty = string.IsNullOrWhiteSpace(blockDwg);
                                bool isDirEmpty = string.IsNullOrWhiteSpace(blockDir);

                                if (isDwgEmpty && isDirEmpty)
                                {
                                    // AA 与 AB 均为空
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{r}_AA_AB",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = r,
                                        ComponentName = name,
                                        Model = model,
                                        Brand = brand,
                                        Quantity = qty,
                                        BasePrice = basePrice,
                                        Discount = discount,
                                        BlockDwg = blockDwg,
                                        BlockDir = blockDir,
                                        TargetCol = "AA",
                                        Category = QuotationIssueCategories.MissingBothCadParams,
                                        Severity = "info",
                                        Title = "AA与AB列CAD图纸参数均为空",
                                        Detail = "图纸名称(AA列)与目录分类(AB列)均未填写，将导致 CAD 无法自动联动出图或无法插入图块",
                                        SuggestedFix = "在明细表 AA 列填入图纸名、AB 列填入目录名，或使用参数匹配工具绑定",
                                        CanAutoFix = false
                                    });
                                    result.Summary.CadMissingCount++;
                                }
                                else if (isDwgEmpty)
                                {
                                    // 仅 AA 图纸名为空
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{r}_AA",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = r,
                                        ComponentName = name,
                                        Model = model,
                                        Brand = brand,
                                        Quantity = qty,
                                        BasePrice = basePrice,
                                        Discount = discount,
                                        BlockDwg = blockDwg,
                                        BlockDir = blockDir,
                                        TargetCol = "AA",
                                        Category = QuotationIssueCategories.MissingBlockDwg,
                                        Severity = "info",
                                        Title = "AA列图纸/图块名称为空",
                                        Detail = "该元件未绑定 CAD 图纸或图块名称(AA列)，可能影响后续电气系统图出图",
                                        SuggestedFix = "在 AA 列补录 DWG 图纸名称",
                                        CanAutoFix = false
                                    });
                                    result.Summary.CadMissingCount++;
                                }
                                else if (isDirEmpty)
                                {
                                    // 仅 AB 目录分类为空
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{r}_AB",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = r,
                                        ComponentName = name,
                                        Model = model,
                                        Brand = brand,
                                        Quantity = qty,
                                        BasePrice = basePrice,
                                        Discount = discount,
                                        BlockDwg = blockDwg,
                                        BlockDir = blockDir,
                                        TargetCol = "AB",
                                        Category = QuotationIssueCategories.MissingBlockDir,
                                        Severity = "info",
                                        Title = "AB列图块分类目录为空",
                                        Detail = "该元件未绑定图块所属目录分类(AB列)，无法在方案库中精确归类",
                                        SuggestedFix = "在 AB 列补录所属目录分类",
                                        CanAutoFix = false
                                    });
                                    result.Summary.CadMissingCount++;
                                }

                                // ----------------------------------------------------
                                // 维度 B: 核心价格与数量漏填检测 (漏项漏价隐患)
                                // ----------------------------------------------------
                                if (basePrice <= 0 && string.IsNullOrWhiteSpace(formulaM))
                                {
                                    // 面价为空或 0
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{r}_M",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = r,
                                        ComponentName = name,
                                        Model = model,
                                        Brand = brand,
                                        Quantity = qty,
                                        BasePrice = basePrice,
                                        Discount = discount,
                                        TargetCol = "M",
                                        Category = QuotationIssueCategories.MissingPrice,
                                        Severity = "error",
                                        Title = "元件面价(M列)缺失或为0",
                                        Detail = $"已录入规格型号【{model}】，但 M 列表价为空或 0，可能导致对外漏报价、严重亏损！",
                                        SuggestedFix = "及时补齐面价或使用【在线查价】/【汇总调价】套价",
                                        CanAutoFix = false
                                    });
                                    result.Summary.MissingPriceCount++;
                                }

                                if (discount <= 0 && string.IsNullOrWhiteSpace(formulaN))
                                {
                                    // 采购折扣为空或 0
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{r}_N",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = r,
                                        ComponentName = name,
                                        Model = model,
                                        Brand = brand,
                                        Quantity = qty,
                                        BasePrice = basePrice,
                                        Discount = discount,
                                        TargetCol = "N",
                                        Category = QuotationIssueCategories.MissingDiscount,
                                        Severity = "error",
                                        Title = "采购折扣(N列)缺失或为0",
                                        Detail = "N 列采购折扣为空或为0，导致销售单价乘算为0或无法体现采购成本",
                                        SuggestedFix = "在 N 列填入采购折扣 (如 0.55 或 1.0)",
                                        CanAutoFix = false
                                    });
                                    result.Summary.WarningCount++;
                                }

                                if (qty <= 0)
                                {
                                    // 数量缺失
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{r}_F",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = r,
                                        ComponentName = name,
                                        Model = model,
                                        Brand = brand,
                                        Quantity = qty,
                                        BasePrice = basePrice,
                                        Discount = discount,
                                        TargetCol = "F",
                                        Category = QuotationIssueCategories.MissingQuantity,
                                        Severity = "error",
                                        Title = "元件数量(F列)小于等于0",
                                        Detail = "F 列数量为 0 或未填写，该元件将无法计入总报价与采购清单",
                                        SuggestedFix = "在 F 列补录合法元件数量",
                                        CanAutoFix = false
                                    });
                                    result.Summary.ErrorCount++;
                                }

                                // ----------------------------------------------------
                                // 维度 C: 公式损毁与错误值检测 (破坏动态联动)
                                // ----------------------------------------------------
                                // 检查销售单价 G 列公式
                                bool gHasFormula = !string.IsNullOrWhiteSpace(formulaG) && formulaG.StartsWith("=");
                                if (!gHasFormula)
                                {
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{r}_G_Broken",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = r,
                                        ComponentName = name,
                                        Model = model,
                                        Brand = brand,
                                        Quantity = qty,
                                        SaleUnitPrice = saleUnitPrice,
                                        CostUnitPrice = costUnitPrice,
                                        BasePrice = basePrice,
                                        Discount = discount,
                                        MarkupFactor = markupFactor,
                                        TargetCol = "G",
                                        Category = QuotationIssueCategories.FormulaBroken,
                                        Severity = "warning",
                                        Title = "销售单价(G列)公式被手写数字覆写",
                                        Detail = $"当前 G 列为常数 [{saleUnitPrice}] 而非联动公式，后续调整面价或折扣时单价无法联动更新！",
                                        SuggestedFix = "点击【一键恢复标准公式】重新注入 =ROUND(M*L*N, 2)",
                                        CanAutoFix = true,
                                        FixAction = "fixFormula"
                                    });
                                    result.Summary.FormulaBrokenCount++;
                                }

                                // 检查销售总价 H 列公式
                                bool hHasFormula = !string.IsNullOrWhiteSpace(formulaH) && formulaH.StartsWith("=");
                                if (!hHasFormula)
                                {
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{r}_H_Broken",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = r,
                                        ComponentName = name,
                                        Model = model,
                                        Brand = brand,
                                        Quantity = qty,
                                        SaleUnitPrice = saleUnitPrice,
                                        SaleTotalPrice = saleTotalPrice,
                                        TargetCol = "H",
                                        Category = QuotationIssueCategories.FormulaBroken,
                                        Severity = "warning",
                                        Title = "销售总价(H列)公式被覆写破坏",
                                        Detail = $"当前 H 列为常数 [{saleTotalPrice}]，修改单价或数量时总价不会重新计算！",
                                        SuggestedFix = "点击【一键恢复标准公式】重新注入 =ROUND(F*G, 2)",
                                        CanAutoFix = true,
                                        FixAction = "fixFormula"
                                    });
                                    result.Summary.FormulaBrokenCount++;
                                }

                                // 检查公式错误值 (#REF!, #VALUE! 等)
                                if (FormulaErrorRegex.IsMatch(formulaG) || FormulaErrorRegex.IsMatch(formulaH) ||
                                    FormulaErrorRegex.IsMatch(formulaJ) || FormulaErrorRegex.IsMatch(formulaK))
                                {
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{r}_Formula_Err",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = r,
                                        ComponentName = name,
                                        Model = model,
                                        Brand = brand,
                                        TargetCol = "G",
                                        Category = QuotationIssueCategories.FormulaError,
                                        Severity = "error",
                                        Title = "单元格包含公式计算错误 (#REF!/#VALUE!)",
                                        Detail = "公式引用断裂或参数类型不匹配，导致 Excel 抛出严重公式错误，阻断合计汇总",
                                        SuggestedFix = "点击【一键恢复标准公式】重置，或检查引用单元格",
                                        CanAutoFix = true,
                                        FixAction = "fixFormula"
                                    });
                                    result.Summary.ErrorCount++;
                                }

                                // ----------------------------------------------------
                                // 维度 D: 业务合理性与利润倒挂检测 (亏损防呆)
                                // ----------------------------------------------------
                                if (saleUnitPrice > 0 && costUnitPrice > 0 && saleUnitPrice < costUnitPrice)
                                {
                                    // 销售单价小于成本单价：致命倒挂
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{r}_Inverted",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = r,
                                        ComponentName = name,
                                        Model = model,
                                        Brand = brand,
                                        Quantity = qty,
                                        SaleUnitPrice = saleUnitPrice,
                                        CostUnitPrice = costUnitPrice,
                                        BasePrice = basePrice,
                                        Discount = discount,
                                        MarkupFactor = markupFactor,
                                        TargetCol = "G",
                                        Category = QuotationIssueCategories.PriceInverted,
                                        Severity = "error",
                                        Title = "销售单价低于成本单价(价格倒挂亏本)",
                                        Detail = $"销售单价 [{saleUnitPrice:F2}] < 成本单价 [{costUnitPrice:F2}]，毛利率为负！加价系数应 >= 1.0",
                                        SuggestedFix = "请调高加价系数(L列)或核查采购折扣(N列)",
                                        CanAutoFix = false
                                    });
                                    result.Summary.InvertedPriceCount++;
                                    result.Summary.ErrorCount++;
                                }

                                if (discount > 1.0m)
                                {
                                    // 采购折扣大于 1 (如把 0.55 输成了 55)
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{r}_Discount_Warn",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = r,
                                        ComponentName = name,
                                        Model = model,
                                        Brand = brand,
                                        Quantity = qty,
                                        BasePrice = basePrice,
                                        Discount = discount,
                                        TargetCol = "N",
                                        Category = QuotationIssueCategories.DiscountExceeded,
                                        Severity = "warning",
                                        Title = "采购折扣大于1.0 (疑输入百分数失误)",
                                        Detail = $"当前折扣为 [{discount}]，若本意为 {discount} 折，标准输入应为 [{discount / 100m}]",
                                        SuggestedFix = "核实折扣量级是否输错（标准成套采购折扣一般介于 0.1~1.0）",
                                        CanAutoFix = false
                                    });
                                    result.Summary.WarningCount++;
                                }

                                // ----------------------------------------------------
                                // 维度 E: 文本型数字检查 (带单引号无法 SUM)
                                // ----------------------------------------------------
                                if (rawBasePriceObj is string strM && decimal.TryParse(strM, out _) && !string.IsNullOrWhiteSpace(strM))
                                {
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{r}_TextNum_M",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = r,
                                        ComponentName = name,
                                        Model = model,
                                        Brand = brand,
                                        TargetCol = "M",
                                        Category = QuotationIssueCategories.TextNumber,
                                        Severity = "warning",
                                        Title = "面价(M列)被存为了文本型数字",
                                        Detail = "单元格格式为文本，可能导致公式计算失效或格式不统一",
                                        SuggestedFix = "点击【一键清洗数字】自动转为标准纯数字",
                                        CanAutoFix = true,
                                        FixAction = "cleanTextNumber"
                                    });
                                    result.Summary.WarningCount++;
                                }
                            }
                        }

                        // 2. 规则 6 计费区域扫描 (subsumRow 到 tolsumRow - 1)
                        if (tolsumRow > subsumRow)
                        {
                            for (int fr = subsumRow; fr < tolsumRow; fr++)
                            {
                                int fArrRow = fr - usedRowOffset + 1;
                                if (fArrRow < 1 || fArrRow > maxRows) continue;

                                string feeName = GetMatrixString(valuesMatrix, fArrRow, 2 - usedColOffset + 1);
                                decimal feeAmt = GetMatrixDecimal(valuesMatrix, fArrRow, 7 - usedColOffset + 1);

                                // 规则 6: 计费区域不能有空行
                                if (string.IsNullOrWhiteSpace(feeName) && feeAmt == 0)
                                {
                                    result.Issues.Add(new QuotationCheckItem
                                    {
                                        Id = $"{sheetName}_R{fr}_Fee_Blank",
                                        SheetName = sheetName,
                                        CabinetName = cabName,
                                        CabinetIndex = cabIndex,
                                        Row = fr,
                                        ComponentName = "计费区域空行",
                                        TargetCol = "B",
                                        Category = QuotationIssueCategories.FeeEmptyOrBlankRow,
                                        Severity = "warning",
                                        Title = "计费区域存在空行 (违背规则 6)",
                                        Detail = "计费区域不能有空行，空行会破坏计费连续性及公式引用",
                                        SuggestedFix = "删除该计费空行，保证计费区域连续",
                                        CanAutoFix = false
                                    });
                                    result.Summary.WarningCount++;
                                }
                            }
                        }
                    }
                }

                // ----------------------------------------------------
                // 维度 F: 同型号全局聚类冲突分析 (同物不同价对齐)
                // ----------------------------------------------------
                var conflictGroups = AnalyzeModelConflicts(allComponents);
                result.ConflictGroups = conflictGroups;
                result.Summary.ConflictCount = conflictGroups.Count;

                // 为同型号冲突行注入问题明细
                foreach (var group in conflictGroups)
                {
                    // 找到所有该型号的条目
                    var matchingItems = allComponents.Where(c => BuildModelKey(c.Model, c.Brand) == group.ModelKey).ToList();
                    foreach (var item in matchingItems)
                    {
                        result.Issues.Add(new QuotationCheckItem
                        {
                            Id = $"{item.SheetName}_R{item.Row}_Conflict",
                            SheetName = item.SheetName,
                            CabinetName = item.CabinetName,
                            CabinetIndex = item.CabinetIndex,
                            Row = item.Row,
                            ComponentName = item.ComponentName,
                            Model = item.Model,
                            Brand = item.Brand,
                            Quantity = item.Quantity,
                            BasePrice = item.BasePrice,
                            Discount = item.Discount,
                            MarkupFactor = item.MarkupFactor,
                            SaleUnitPrice = item.SaleUnitPrice,
                            CostUnitPrice = item.CostUnitPrice,
                            TargetCol = "M",
                            Category = QuotationIssueCategories.PriceConflict,
                            Severity = "warning",
                            Title = $"同型号价格不一致 ({group.VariantCount}种价格版本)",
                            Detail = $"全项目【{item.Model}】存在 {group.VariantCount} 种不同价格/折扣！当前行: 表价 {item.BasePrice}, 折扣 {item.Discount}, 加价 {item.MarkupFactor}",
                            SuggestedFix = group.RecommendedVariant != null
                                ? $"建议统一对齐为推荐主流值: 表价 {group.RecommendedVariant.BasePrice}, 折扣 {group.RecommendedVariant.Discount}"
                                : "在冲突面板中选择基准一键刷平",
                            CanAutoFix = true,
                            FixAction = "alignModel"
                        });
                    }
                }

                // 汇总各类问题数量
                result.Summary.TotalIssues = result.Issues.Count;
                result.Summary.ErrorCount = result.Issues.Count(i => i.Severity == "error");
                result.Summary.WarningCount = result.Issues.Count(i => i.Severity == "warning");
                result.Summary.InfoCount = result.Issues.Count(i => i.Severity == "info");

                // 计算综合健康得分 (扣分制)
                int score = 100;
                score -= result.Summary.ErrorCount * 6;      // 严重错误每个扣 6 分
                score -= result.Summary.WarningCount * 2;    // 警告每个扣 2 分
                score -= (int)(result.Summary.InfoCount * 0.5); // 提示每个扣 0.5 分
                if (score < 0) score = 0;
                result.Summary.HealthScore = score;

                result.Success = true;
                result.Message = $"体检完成！共扫描 {result.Summary.TotalCabinets} 台箱柜、{result.Summary.TotalComponents} 行元器件，发现 {result.Summary.TotalIssues} 项需关注内容。";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"报价体检执行异常: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// 一键批量修复指定行或全量公式损坏行的计算公式
        /// 重新写入标准公式：
        /// G 列: =IF(AND(B{row}="",C{row}=""),"",ROUND(M{row}*L{row}*N{row},2))
        /// H 列: =IF(AND(B{row}="",C{row}=""),"",ROUND(F{row}*G{row},2))
        /// J 列: =IF(AND(B{row}="",C{row}=""),"",ROUND(M{row}*N{row},2))
        /// K 列: =IF(AND(B{row}="",C{row}=""),"",ROUND(J{row}*F{row},2))
        /// </summary>
        /// <param name="issueIds">需修复的问题 ID 列表 (为空时自动修复全表所有公式损坏行)</param>
        /// <returns>修复结果与处理行数统计</returns>
        public static (bool Success, string Message, int FixedCount) FixQuotationFormulas(List<string>? issueIds = null)
        {
            try
            {
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null) return (false, "Excel 宿主对象未就绪", 0);
                Workbook? wb = app.ActiveWorkbook;
                if (wb == null) return (false, "未检测到有效工作簿", 0);

                // 关屏与事件加速
                app.ScreenUpdating = false;
                app.EnableEvents = false;

                int fixedCount = 0;

                try
                {
                    // 重新跑一次审计获取最新需要修复的公式行
                    var audit = ExecuteQuotationAudit(wb);
                    var brokenItems = audit.Issues
                        .Where(i => i.Category == QuotationIssueCategories.FormulaBroken || i.Category == QuotationIssueCategories.FormulaError)
                        .ToList();

                    // 若传了特定的 ID 集合则过滤
                    if (issueIds != null && issueIds.Count > 0)
                    {
                        var idSet = new HashSet<string>(issueIds);
                        brokenItems = brokenItems.Where(i => idSet.Contains(i.Id)).ToList();
                    }

                    // 按工作表分组批量处理
                    var sheetGroups = brokenItems.GroupBy(i => i.SheetName);
                    foreach (var grp in sheetGroups)
                    {
                        // 动态获取工作表对象，避免 COM 强转失败
                        dynamic? sheet = null;
                        try { sheet = wb.Worksheets[grp.Key]; } catch { }
                        if (sheet == null) continue;

                        // 提取独特的物理行号
                        var rows = grp.Select(i => i.Row).Distinct().ToList();
                        foreach (int r in rows)
                        {
                            // 重新刷回标准动态公式 (规则定义)
                            sheet.Cells[r, 7].Formula = $"=IF(AND(B{r}=\"\",C{r}=\"\"),\"\",ROUND(M{r}*L{r}*N{r},2))";
                            sheet.Cells[r, 8].Formula = $"=IF(AND(B{r}=\"\",C{r}=\"\"),\"\",ROUND(F{r}*G{r},2))";
                            sheet.Cells[r, 10].Formula = $"=IF(AND(B{r}=\"\",C{r}=\"\"),\"\",ROUND(M{r}*N{r},2))";
                            sheet.Cells[r, 11].Formula = $"=IF(AND(B{r}=\"\",C{r}=\"\"),\"\",ROUND(J{r}*F{r},2))";
                            fixedCount++;
                        }
                    }
                }
                finally
                {
                    app.ScreenUpdating = true;
                    app.EnableEvents = true;
                }

                return (true, $"成功恢复 {fixedCount} 行的标准计算公式！", fixedCount);
            }
            catch (Exception ex)
            {
                return (false, $"公式修复异常: {ex.Message}", 0);
            }
        }

        /// <summary>
        /// 同型号价格批量对齐广播服务
        /// 将指定型号的面价、折扣、加价系数批量刷入全工作簿匹配的所有分类表明细行
        /// </summary>
        /// <param name="req">对齐配置请求</param>
        /// <returns>对齐结果与受影响行数</returns>
        public static (bool Success, string Message, int UpdatedCount) AlignModelPriceAcrossSheets(AlignModelPriceRequest req)
        {
            if (req == null || string.IsNullOrWhiteSpace(req.Model))
            {
                return (false, "目标型号规格不能为空", 0);
            }

            try
            {
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null) return (false, "Excel 宿主对象未就绪", 0);
                Workbook? wb = app.ActiveWorkbook;
                if (wb == null) return (false, "未检测到有效工作簿", 0);

                app.ScreenUpdating = false;
                app.EnableEvents = false;

                int updatedCount = 0;

                try
                {
                    List<string> categorySheets = Tool.GetProjectCategorySheetNames(wb, forceRefresh: false).ToList();
                    string targetModelKey = BuildModelKey(req.Model, req.Brand);

                    foreach (string sheetName in categorySheets)
                    {
                        // 动态获取工作表对象
                        dynamic? sheet = null;
                        try { sheet = wb.Worksheets[sheetName]; } catch { }
                        if (sheet == null) continue;

                        var cabinets = Tool.GetSheetValidCabinets((object)sheet, (object)wb);
                        if (cabinets == null) continue;

                        foreach (var kvp in cabinets)
                        {
                            var cab = kvp.Value;
                            int detRow = cab.Det != null ? Convert.ToInt32(cab.Det.Row) : 0;
                            int subsumRow = cab.Subsum != null ? Convert.ToInt32(cab.Subsum.Row) : 0;
                            int compStart = detRow + 2;
                            int compEnd = subsumRow - 1;

                            if (compEnd < compStart) continue;

                            for (int r = compStart; r <= compEnd; r++)
                            {
                                string cModel = Convert.ToString(sheet.Cells[r, 3].Value2)?.Trim() ?? string.Empty;
                                string dBrand = Convert.ToString(sheet.Cells[r, 4].Value2)?.Trim() ?? string.Empty;

                                if (BuildModelKey(cModel, dBrand) == targetModelKey)
                                {
                                    // 回写面价 (M 列)
                                    sheet.Cells[r, 13].Value2 = (double)req.TargetBasePrice;
                                    // 回写折扣 (N 列)
                                    sheet.Cells[r, 14].Value2 = (double)req.TargetDiscount;
                                    // 回写加价系数 (L 列)
                                    sheet.Cells[r, 12].Value2 = (double)req.TargetMarkupFactor;

                                    // 可选重置公式，保证联动
                                    if (req.ResetFormulas)
                                    {
                                        sheet.Cells[r, 7].Formula = $"=IF(AND(B{r}=\"\",C{r}=\"\"),\"\",ROUND(M{r}*L{r}*N{r},2))";
                                        sheet.Cells[r, 8].Formula = $"=IF(AND(B{r}=\"\",C{r}=\"\"),\"\",ROUND(F{r}*G{r},2))";
                                        sheet.Cells[r, 10].Formula = $"=IF(AND(B{r}=\"\",C{r}=\"\"),\"\",ROUND(M{r}*N{r},2))";
                                        sheet.Cells[r, 11].Formula = $"=IF(AND(B{r}=\"\",C{r}=\"\"),\"\",ROUND(J{r}*F{r},2))";
                                    }

                                    updatedCount++;
                                }
                            }
                        }
                    }
                }
                finally
                {
                    app.ScreenUpdating = true;
                    app.EnableEvents = true;
                }

                return (true, $"已将型号【{req.Model}】的价格与折扣成功同步对齐至 {updatedCount} 处回路！", updatedCount);
            }
            catch (Exception ex)
            {
                return (false, $"同型号价格对齐异常: {ex.Message}", 0);
            }
        }

        /// <summary>
        /// 一键清洗文本型数字为纯数值格式 (防止 Excel SUM 漏算)
        /// </summary>
        public static (bool Success, string Message, int CleanedCount) CleanQuotationTextNumbers()
        {
            try
            {
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null) return (false, "Excel 宿主对象未就绪", 0);
                Workbook? wb = app.ActiveWorkbook;
                if (wb == null) return (false, "未检测到有效工作簿", 0);

                app.ScreenUpdating = false;
                app.EnableEvents = false;
                int cleanedCount = 0;

                try
                {
                    var audit = ExecuteQuotationAudit(wb);
                    var textItems = audit.Issues.Where(i => i.Category == QuotationIssueCategories.TextNumber).ToList();

                    foreach (var item in textItems)
                    {
                        // 动态提取工作表引用
                        dynamic? sheet = null;
                        try { sheet = wb.Worksheets[item.SheetName]; } catch { }
                        if (sheet == null) continue;

                        // 转换列字母为物理索引 (1-based)
                        int colIdx = GetColIndex(item.TargetCol);
                        if (colIdx <= 0) continue;

                        // 读取当前单元格值
                        object? val = sheet.Cells[item.Row, colIdx].Value2;
                        // 若为文本数值则强转写回双精度浮点数
                        if (val is string str && decimal.TryParse(str, out decimal numVal))
                        {
                            sheet.Cells[item.Row, colIdx].Value2 = (double)numVal;
                            cleanedCount++;
                        }
                    }
                }
                finally
                {
                    app.ScreenUpdating = true;
                    app.EnableEvents = true;
                }

                return (true, $"已成功将 {cleanedCount} 处文本数字清洗为标准纯数值！", cleanedCount);
            }
            catch (Exception ex)
            {
                return (false, $"文本清洗异常: {ex.Message}", 0);
            }
        }

        /// <summary>
        /// Excel 视口跳转与高亮定位
        /// </summary>
        /// <param name="sheetName">目标工作表名称</param>
        /// <param name="row">目标物理行号</param>
        /// <param name="colLetter">目标列字母 (如 "AA", "G", "M")</param>
        public static bool NavigateToQuotationCell(string sheetName, int row, string colLetter)
        {
            try
            {
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null) return false;
                Workbook? wb = app.ActiveWorkbook;
                if (wb == null) return false;

                // 动态获取工作表
                dynamic? sheet = null;
                try { sheet = wb.Worksheets[sheetName]; } catch { }
                if (sheet == null) return false;

                // 激活工作表
                sheet.Activate();

                // 选定目标单元格并滚动至视口可见
                dynamic targetCell = sheet.Range[$"{colLetter}{row}"];
                targetCell.Select();

                // 适度调整视口滚动行，保留上下文
                dynamic? win = app.ActiveWindow;
                if (win != null && row > 5)
                {
                    win.ScrollRow = Math.Max(1, row - 4);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        // ==========================================
        // 内部辅助私有方法
        // ==========================================

        /// <summary>
        /// 聚类分析同型号不同价格冲突
        /// </summary>
        private static List<ModelConflictGroup> AnalyzeModelConflicts(List<AuditedComponentItem> components)
        {
            var conflictGroups = new List<ModelConflictGroup>();

            // 按 (型号 + 品牌) 复合键分组
            var groups = components
                .Where(c => !string.IsNullOrWhiteSpace(c.Model))
                .GroupBy(c => BuildModelKey(c.Model, c.Brand));

            foreach (var grp in groups)
            {
                // 该组元器件总数
                int totalRows = grp.Count();
                if (totalRows <= 1) continue; // 仅 1 行不可能冲突

                // 聚合提炼不同的价格组合 (BasePrice, Discount, MarkupFactor)
                var variants = grp
                    .GroupBy(c => new {
                        BasePrice = Math.Round(c.BasePrice, 2),
                        Discount = Math.Round(c.Discount, 4),
                        MarkupFactor = Math.Round(c.MarkupFactor, 2)
                    })
                    .Select(vg => new ModelPriceVariant
                    {
                        BasePrice = vg.Key.BasePrice,
                        Discount = vg.Key.Discount,
                        MarkupFactor = vg.Key.MarkupFactor,
                        DerivedSalePrice = Math.Round(vg.Key.BasePrice * vg.Key.Discount * vg.Key.MarkupFactor, 2),
                        Occurrences = vg.Count(),
                        SampleLocation = $"{vg.First().SheetName} > {vg.First().CabinetName} > 第{vg.First().Row}行",
                        SampleSheet = vg.First().SheetName,
                        SampleRow = vg.First().Row
                    })
                    .OrderByDescending(v => v.Occurrences) // 按频次倒序排列
                    .ToList();

                // 若存在 2 种及以上价格组合，则判定为价格冲突
                if (variants.Count > 1)
                {
                    var firstItem = grp.First();
                    conflictGroups.Add(new ModelConflictGroup
                    {
                        ModelKey = grp.Key,
                        Model = firstItem.Model,
                        Brand = firstItem.Brand,
                        TotalRows = totalRows,
                        VariantCount = variants.Count,
                        Variants = variants,
                        RecommendedVariant = variants[0] // 频次最高者推荐为主流基准
                    });
                }
            }

            return conflictGroups;
        }

        /// <summary>
        /// 构造型号与品牌的复合去重键 (不区分大小写)
        /// </summary>
        private static string BuildModelKey(string model, string brand)
        {
            string m = (model ?? string.Empty).Trim().ToUpper();
            string b = (brand ?? string.Empty).Trim().ToUpper();
            return string.IsNullOrEmpty(b) ? m : $"{m} | {b}";
        }

        /// <summary>
        /// 从二维数组安全提取字符串
        /// </summary>
        private static string GetMatrixString(object[,]? matrix, int r, int c)
        {
            if (matrix == null) return string.Empty;
            if (r < 1 || r > matrix.GetLength(0) || c < 1 || c > matrix.GetLength(1)) return string.Empty;
            object? val = matrix[r, c];
            return Convert.ToString(val)?.Trim() ?? string.Empty;
        }

        /// <summary>
        /// 从二维数组安全提取浮点数值
        /// </summary>
        private static decimal GetMatrixDecimal(object[,]? matrix, int r, int c)
        {
            if (matrix == null) return 0;
            if (r < 1 || r > matrix.GetLength(0) || c < 1 || c > matrix.GetLength(1)) return 0;
            object? val = matrix[r, c];
            if (val == null) return 0;
            if (decimal.TryParse(Convert.ToString(val), out decimal d)) return d;
            return 0;
        }

        /// <summary>
        /// 从二维数组安全提取原生对象
        /// </summary>
        private static object? GetMatrixObject(object[,]? matrix, int r, int c)
        {
            if (matrix == null) return null;
            if (r < 1 || r > matrix.GetLength(0) || c < 1 || c > matrix.GetLength(1)) return null;
            return matrix[r, c];
        }

        /// <summary>
        /// 从公式矩阵安全提取公式字符串
        /// </summary>
        private static string GetMatrixFormula(object[,]? formulaMatrix, int r, int c)
        {
            if (formulaMatrix == null) return string.Empty;
            if (r < 1 || r > formulaMatrix.GetLength(0) || c < 1 || c > formulaMatrix.GetLength(1)) return string.Empty;
            object? val = formulaMatrix[r, c];
            return Convert.ToString(val)?.Trim() ?? string.Empty;
        }

        /// <summary>
        /// 将列字母字符 (如 "A", "G", "M", "AA", "AB") 解析转换为 1-based 物理列索引
        /// </summary>
        /// <param name="colLetter">列标字符</param>
        /// <returns>1-based 列号索引</returns>
        private static int GetColIndex(string colLetter)
        {
            // 若列字母为空或空白则返回 0
            if (string.IsNullOrWhiteSpace(colLetter)) return 0;
            // 累加列号
            int col = 0;
            // 循环遍历字母序列
            foreach (char ch in colLetter.ToUpper())
            {
                // 仅处理大写字母
                if (ch >= 'A' && ch <= 'Z')
                {
                    // 26 进制位移叠加 (A=1, B=2, ..., Z=26, AA=27)
                    col = col * 26 + (ch - 'A' + 1);
                }
            }
            // 返回计算出的物理列号
            return col;
        }

        /// <summary>
        /// 内部元器件审计暂存轻量结构
        /// </summary>
        private class AuditedComponentItem
        {
            public string SheetName { get; set; } = string.Empty;
            public string CabinetName { get; set; } = string.Empty;
            public int CabinetIndex { get; set; }
            public int Row { get; set; }
            public string ComponentName { get; set; } = string.Empty;
            public string Model { get; set; } = string.Empty;
            public string Brand { get; set; } = string.Empty;
            public decimal Quantity { get; set; }
            public decimal SaleUnitPrice { get; set; }
            public decimal SaleTotalPrice { get; set; }
            public decimal CostUnitPrice { get; set; }
            public decimal MarkupFactor { get; set; }
            public decimal BasePrice { get; set; }
            public decimal Discount { get; set; }
            public string BlockDwg { get; set; } = string.Empty;
            public string BlockDir { get; set; } = string.Empty;
            public string FormulaG { get; set; } = string.Empty;
            public string FormulaH { get; set; } = string.Empty;
            public string FormulaJ { get; set; } = string.Empty;
            public string FormulaK { get; set; } = string.Empty;
        }
    }
}
