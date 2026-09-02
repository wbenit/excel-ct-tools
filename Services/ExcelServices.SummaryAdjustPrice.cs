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
                // 弹出异常提示信息
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

                    // 1. 调用公共方法获取当前工作表的有效箱柜锚点列表 (规则 6)
                    var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);

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
            catch
            {
                // 捕获异常静默处理
            }

            return result;
        }

        /// <summary>
        /// 内存汇总聚合模型辅助类（支持 16 列完整字段与本体/附件拆分）
        /// </summary>
        private class AggregatedComponent
        {
            // 元件名称 (对应 Excel B 列)
            public string Name { get; set; } = string.Empty;

            // 型号规格 (对应 Excel C 列)
            public string Model { get; set; } = string.Empty;

            // 生产厂家 / 品牌 (对应 Excel H 列 / 原明细 D 列)
            public string Manufacturer { get; set; } = string.Empty;

            // 计量单位 (对应 Excel D 列 / 原明细 E 列)
            public string Unit { get; set; } = string.Empty;

            // 数量 (包含箱柜台数乘积计算后的汇总总数量，对应 Excel E 列)
            public decimal Quantity { get; set; } = 0;

            // 销售单价 (加权平均单价，对应 Excel F 列)
            public decimal UnitPrice { get; set; } = 0;

            // 销售总价 (包含箱柜台数乘积计算后的汇总总金额，对应 Excel G 列)
            public decimal TotalPrice { get; set; } = 0;

            // 成本单价 (加权平均成本单价，对应 Excel I 列)
            public decimal CostUnitPrice { get; set; } = 0;

            // 成本总价 (包含箱柜台数乘积计算后的汇总成本金额)
            public decimal CostTotalPrice { get; set; } = 0;

            // 报出系数 (销售单价/成本单价 或 原明细 L 列加价系数，对应 Excel J 列)
            public decimal MarkupFactor { get; set; } = 1.0m;

            // 本体表价 (从明细 M 列公式第一项解析提取，对应 Excel K 列)
            public decimal BasePrice { get; set; } = 0;

            // 本体折扣 (从明细 N 列公式解析提取，对应 Excel L 列)
            public decimal BaseDiscount { get; set; } = 1.0m;

            // 附件表价 (从明细 M 列公式第二项及之后项累加解析，对应 Excel M 列)
            public decimal AccessoryPrice { get; set; } = 0;

            // 附件折扣 (从明细 N 列公式解析提取，对应 Excel N 列)
            public decimal AccessoryDiscount { get; set; } = 1.0m;

            // 备注说明 (对应 Excel O 列 / 原明细 I 列)
            public string Remark { get; set; } = string.Empty;

            // 所属类别 (对应 Excel P 列 / 原明细 Q 列，默认“元件”)
            public string Category { get; set; } = "元件";

            // 原始型号规格 (预留扩展)
            public string OriginalModel { get; set; } = string.Empty;

            // 明细表 U 列原始型号 (汇总表 R 列对应，相同规格合并时用“，”连接)
            public string RawModelFromU { get; set; } = string.Empty;

            // 额定电流 (对应明细表 W 列 / 汇总表 T 列)
            public string Current { get; set; } = string.Empty;

            // 极数 (对应明细表 X 列 / 汇总表 U 列)
            public string Poles { get; set; } = string.Empty;

            // 脱扣方式 (对应明细表 Y 列 / 汇总表 V 列)
            public string Tripping { get; set; } = string.Empty;

            // 配套附件 (对应明细表 Z 列 / 汇总表 W 列)
            public string Accessory { get; set; } = string.Empty;

            // 扩展参数 1 / 图块名称 (对应明细表 AA 列 / 汇总表 X 列)
            public string Param1 { get; set; } = string.Empty;

            // 扩展参数 2 / 图块类别 (对应明细表 AB 列 / 汇总表 Y 列)
            public string Param2 { get; set; } = string.Empty;

            // CAD 图元句柄 (对应明细表 AD 列 / 汇总表 AD 列)
            public string Handle { get; set; } = string.Empty;

            // 所属分类工作表 Sheet 名称
            public string SheetName { get; set; } = string.Empty;
        }

        /// <summary>
        /// 从明细 M 列（面价/表价）中拆分本体表价与附件表价
        /// 规则：形如 =159.05+336.2 时，第一个加数 159.05 为本体，后续所有加数 336.2 累加为附件
        /// </summary>
        /// <param name="valM">单元格计算值 (Value2)</param>
        /// <param name="formulaM">单元格公式字符串 (Formula)</param>
        /// <param name="basePrice">输出解析出的本体表价</param>
        /// <param name="accPrice">输出解析出的附件表价</param>
        private static void ParseBaseAndAccessoryPrice(
            object? valM,
            object? formulaM,
            out decimal basePrice,
            out decimal accPrice)
        {
            // 初始化默认本体表价为 0
            basePrice = 0;
            // 初始化默认附件表价为 0
            accPrice = 0;

            // 提取公式字符串与数值字符串
            string formulaStr = Convert.ToString(formulaM)?.Trim() ?? string.Empty;
            // 提取纯数值字符串
            string valStr = Convert.ToString(valM)?.Trim() ?? string.Empty;

            // 优先分析公式字符串：若公式以 "=" 开头且包含加号 "+"
            if (!string.IsNullOrEmpty(formulaStr) && formulaStr.StartsWith("="))
            {
                // 去除前导等号
                string cleanFormula = formulaStr.Substring(1).Trim();
                // 若包含加号说明有本体和附件组合
                if (cleanFormula.Contains("+"))
                {
                    // 按加号分割各项表达式
                    string[] parts = cleanFormula.Split('+');
                    if (parts.Length > 0)
                    {
                        // 第一项为本体表价
                        if (decimal.TryParse(parts[0].Trim(), out decimal p0))
                        {
                            basePrice = p0;
                        }

                        // 第二项及后续所有项累加为附件表价
                        decimal sumAcc = 0;
                        for (int i = 1; i < parts.Length; i++)
                        {
                            if (decimal.TryParse(parts[i].Trim(), out decimal pi))
                            {
                                sumAcc += pi;
                            }
                        }
                        // 赋值附件总表价
                        accPrice = sumAcc;
                        return;
                    }
                }
                else
                {
                    // 单项公式，如 =22000
                    if (decimal.TryParse(cleanFormula, out decimal pSingle))
                    {
                        basePrice = pSingle;
                        accPrice = 0;
                        return;
                    }
                }
            }

            // 若无公式或公式为纯数值，直接解析计算后的数值 Value2
            if (decimal.TryParse(valStr, out decimal directVal))
            {
                basePrice = directVal;
                accPrice = 0;
            }
        }

        /// <summary>
        /// 从明细 N 列（折扣）中拆分本体折扣与附件折扣
        /// 规则：形如 =(159.05*1*0.5+336.2*1*1)/495.25 时，拆分提取本体折扣 0.5 与附件折扣 1.0
        /// </summary>
        /// <param name="valN">单元格计算值 (Value2)</param>
        /// <param name="formulaN">单元格公式字符串 (Formula)</param>
        /// <param name="basePrice">已解析出的本体表价</param>
        /// <param name="accPrice">已解析出的附件表价</param>
        /// <param name="baseDiscount">输出解析出的本体折扣</param>
        /// <param name="accDiscount">输出解析出的附件折扣</param>
        private static void ParseBaseAndAccessoryDiscount(
            object? valN,
            object? formulaN,
            decimal basePrice,
            decimal accPrice,
            out decimal baseDiscount,
            out decimal accDiscount)
        {
            // 初始化默认本体折扣为 1.0
            baseDiscount = 1.0m;
            // 初始化默认附件折扣为 1.0
            accDiscount = 1.0m;

            // 提取公式字符串与数值字符串
            string formulaStr = Convert.ToString(formulaN)?.Trim() ?? string.Empty;
            // 提取纯数值字符串
            string valStr = Convert.ToString(valN)?.Trim() ?? string.Empty;

            // 优先分析公式字符串
            if (!string.IsNullOrEmpty(formulaStr) && formulaStr.StartsWith("="))
            {
                // 去除前导等号
                string expr = formulaStr.Substring(1).Trim();

                // 兼容外层 ROUND(...) 函数包裹：如 =ROUND((159.05*1*0.5+336.2*1*1)/495.25 ,2)
                if (expr.StartsWith("ROUND(", StringComparison.OrdinalIgnoreCase) && expr.EndsWith(")"))
                {
                    // 去除前导 "ROUND(" 和末尾 ")"
                    string inner = expr.Substring(6, expr.Length - 7).Trim();
                    // 查找末尾的小数位参数分隔逗号（如 ,2 或 , 2）
                    int lastComma = inner.LastIndexOf(',');
                    if (lastComma > 0)
                    {
                        expr = inner.Substring(0, lastComma).Trim();
                    }
                    else
                    {
                        expr = inner;
                    }
                }

                // 检查是否为复合公式，形如 (A*B*C + D*E*F)/G 或 A*B*C + D*E*F
                if (expr.Contains("*") && expr.Contains("+"))
                {
                    // 提取分子部分：若包含除号 "/" 则截取分子，并去除外层括号
                    string numerator = expr;
                    int slashIdx = expr.LastIndexOf('/');
                    if (slashIdx > 0)
                    {
                        numerator = expr.Substring(0, slashIdx).Trim();
                    }
                    // 清理外层可能存在的圆括号
                    while (numerator.StartsWith("(") && numerator.EndsWith(")"))
                    {
                        numerator = numerator.Substring(1, numerator.Length - 2).Trim();
                    }

                    // 按加号分割各加数项
                    string[] terms = numerator.Split('+');
                    if (terms.Length > 0)
                    {
                        // 解析第一项（本体项）：如 159.05*1*0.5，保留 2 位小数
                        baseDiscount = Math.Round(ExtractDiscountFromTerm(terms[0], basePrice), 2);

                        // 解析第二项及后续项（附件项）：如 336.2*1*1
                        if (terms.Length > 1)
                        {
                            decimal totalAccWeighted = 0;
                            decimal totalAccBase = 0;
                            for (int i = 1; i < terms.Length; i++)
                            {
                                string[] termFactors = terms[i].Split('*');
                                decimal d = ExtractDiscountFromTerm(terms[i], 0);
                                decimal p = 0;
                                if (termFactors.Length > 0 && decimal.TryParse(termFactors[0].Trim().Trim('(', ')'), out decimal pFactor))
                                {
                                    p = pFactor;
                                }
                                decimal qty = 1;
                                if (termFactors.Length > 2 && decimal.TryParse(termFactors[1].Trim().Trim('(', ')'), out decimal qFactor))
                                {
                                    qty = qFactor;
                                }

                                totalAccWeighted += (p * qty * d);
                                totalAccBase += (p * qty);
                            }

                            if (totalAccBase > 0)
                            {
                                // 附件折扣按加权平均计算，保留 2 位小数 (如 0.85)
                                accDiscount = Math.Round(totalAccWeighted / totalAccBase, 2);
                            }
                            else
                            {
                                accDiscount = Math.Round(ExtractDiscountFromTerm(terms[1], accPrice), 2);
                            }
                        }
                        return;
                    }
                }
                else
                {
                    // 简单公式，如 =0.85 或 =1 或 =ROUND(0.852, 2)
                    string cleanSimple = expr.Trim();
                    if (decimal.TryParse(cleanSimple, out decimal simpleD))
                    {
                        baseDiscount = Math.Round(simpleD, 2);
                        accDiscount = accPrice > 0 ? baseDiscount : 1.0m;
                        return;
                    }
                }
            }

            // 若无公式或为纯数值，直接解析 Value2 并保留 2 位小数
            if (decimal.TryParse(valStr, out decimal parsedVal))
            {
                baseDiscount = Math.Round(parsedVal, 2);
                accDiscount = accPrice > 0 ? baseDiscount : 1.0m;
            }
        }

        /// <summary>
        /// 从单项乘积表达式（如 159.05*1*0.5 或 159.05*0.5 或 0.5）中提取折扣系数
        /// </summary>
        /// <param name="termStr">单项表达式字符串</param>
        /// <param name="expectedPrice">预期的表价值（辅助匹配）</param>
        /// <returns>提取出的折扣数值</returns>
        private static decimal ExtractDiscountFromTerm(string termStr, decimal expectedPrice)
        {
            // 清洗首尾空格与圆括号
            termStr = termStr.Trim().Trim('(', ')');
            if (string.IsNullOrEmpty(termStr)) return 1.0m;

            // 按乘号分割因子
            string[] factors = termStr.Split('*');
            if (factors.Length == 1)
            {
                // 仅一个因子，直接解析为折扣
                if (decimal.TryParse(factors[0].Trim(), out decimal dSingle))
                {
                    return dSingle;
                }
                return 1.0m;
            }

            if (factors.Length == 2)
            {
                // 两个因子，形如 159.05*0.5 或 0.5*159.05
                if (decimal.TryParse(factors[0].Trim(), out decimal f0) && decimal.TryParse(factors[1].Trim(), out decimal f1))
                {
                    if (expectedPrice > 0 && Math.Abs(f0 - expectedPrice) < 0.01m)
                    {
                        return f1;
                    }
                    if (expectedPrice > 0 && Math.Abs(f1 - expectedPrice) < 0.01m)
                    {
                        return f0;
                    }
                    // 默认取较小者为折扣
                    return Math.Min(f0, f1);
                }
            }

            if (factors.Length >= 3)
            {
                // 三个及以上因子，标准结构为 价格 * 数量 * 折扣
                // 优先取第 3 个因子（索引 2）
                if (decimal.TryParse(factors[2].Trim(), out decimal f2))
                {
                    return f2;
                }
                // 兜底取最后一个因子
                if (decimal.TryParse(factors[factors.Length - 1].Trim(), out decimal fLast))
                {
                    return fLast;
                }
            }

            // 默认返回 1.0
            return 1.0m;
        }

        /// <summary>
        /// 执行元件数据提取、内存合并聚合并在当前工作簿生成“元件汇总表” (17 列双层表头版，包含“原型号规格”基准参考列)
        /// </summary>
        public static bool GenerateComponentSummarySheet(Controllers.GenerateSummaryRequest request)
        {
            // 校验请求参数有效性
            if (request == null || request.SelectedSheets == null || request.SelectedSheets.Count == 0) return false;

            try
            {
                // 获取当前活动的 Excel Application 实例 (安全访问)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                // 若未获取到 Excel 实例直接返回失败
                if (app == null) return false;

                // 获取活动工作簿
                dynamic activeWb = app.ActiveWorkbook;
                // 若工作簿为空则返回失败
                if (activeWb == null) return false;

                // 临时关闭屏幕刷新与提示警告以提升处理速度
                app.ScreenUpdating = false;
                // 关闭 Excel 警告弹窗
                app.DisplayAlerts = false;

                // 原始元器件提取列表
                var rawComponents = new List<AggregatedComponent>();

                // 遍历用户选中的所有分类工作表
                foreach (string sheetName in request.SelectedSheets)
                {
                    // 定义工作表动态对象
                    dynamic? sheet = null;
                    try
                    {
                        // 按名称获取工作表 COM 实例
                        sheet = activeWb.Worksheets[sheetName];
                    }
                    catch { continue; }
                    // 校验工作表非空
                    if (sheet == null) continue;

                    // 调用公共方法构建当前分类表的有效箱柜锚点字典 (规则 6)
                    var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);

                    // 遍历当前分类表中的每个箱柜
                    foreach (var cab in validCabinets)
                    {
                        // 缺少明细锚点则跳过
                        if (cab.Value.Det == null) continue;

                        // 获取明细起始行号与小计行号
                        int detRow = Convert.ToInt32(cab.Value.Det.Row);
                        // 小计行号默认为 detRow + 27 --硬编码--
                        int subsumRow = cab.Value.Subsum != null ? Convert.ToInt32(cab.Value.Subsum.Row) : detRow + 27;
                        // 汇总行号
                        int sumRow = cab.Value.Sum != null ? Convert.ToInt32(cab.Value.Sum.Row) : 0;
                        // 默认单台箱柜台数
                        int cabQty = 1;

                        // 获取柜体台数（规则 6 & 8，优先取汇总行 F 列数量）
                        if (sumRow > 0)
                        {
                            try
                            {
                                // 优先读取汇总行数量
                                object qVal = sheet.Cells[sumRow, 6].Value ?? sheet.Cells[sumRow, 5].Value;
                                // 转换台数整数
                                if (qVal != null && int.TryParse(Convert.ToString(qVal), out int pQty) && pQty > 0)
                                {
                                    cabQty = pQty;
                                }
                            }
                            catch { }
                        }

                        // 元器件区域起始行与终止行
                        int compStartRow = detRow + 2;
                        // 终止行为小计行前一行 (规则 6)
                        int compEndRow = subsumRow - 1;
                        // 行号倒置则跳过
                        if (compEndRow < compStartRow) continue;

                        // 计算区域总行数
                        int rowCount = compEndRow - compStartRow + 1;

                        // 采用 2D 数组一次性批量读入内存（包含值矩阵与公式矩阵，覆盖 A~AD 列共 30 列，规则 7）
                        dynamic compRange = sheet.Range[$"A{compStartRow}:AD{compEndRow}"];
                        // 读取值矩阵
                        object[,] valMatrix = (object[,])compRange.Value2;
                        // 读取公式矩阵
                        object[,] formulaMatrix = (object[,])compRange.Formula;

                        // 逐行解析元器件字段
                        for (int r = 1; r <= rowCount; r++)
                        {
                            // B 列 (索引 2): 元件名称
                            string compName = Convert.ToString(valMatrix[r, 2]) ?? string.Empty;
                            // C 列 (索引 3): 型号规格 (从分类表中提取时的原始型号规格)
                            string model = Convert.ToString(valMatrix[r, 3]) ?? string.Empty;

                            // 跳过名称与型号均为空的空白行
                            if (string.IsNullOrWhiteSpace(compName) && string.IsNullOrWhiteSpace(model))
                            {
                                continue;
                            }

                            // D 列 (索引 4): 生产厂家
                            string mfg = Convert.ToString(valMatrix[r, 4]) ?? string.Empty;
                            // E 列 (索引 5): 单位
                            string unit = Convert.ToString(valMatrix[r, 5]) ?? string.Empty;

                            // F 列 (索引 6): 数量
                            decimal qty = 0;
                            // 解析数量数值
                            if (decimal.TryParse(Convert.ToString(valMatrix[r, 6]), out decimal q)) qty = q;

                            // G 列 (索引 7): 销售单价
                            decimal unitPrice = 0;
                            // 解析单价值
                            if (decimal.TryParse(Convert.ToString(valMatrix[r, 7]), out decimal p)) unitPrice = p;

                            // H 列 (索引 8): 销售总价
                            decimal totalP = 0;
                            // 解析销售总价
                            if (decimal.TryParse(Convert.ToString(valMatrix[r, 8]), out decimal tp)) totalP = tp;
                            // 若总价未填写则自动计算
                            else totalP = qty * unitPrice;

                            // I 列 (索引 9): 备注
                            string remark = Convert.ToString(valMatrix[r, 9]) ?? string.Empty;

                            // J 列 (索引 10): 成本单价
                            decimal costUnitPrice = 0;
                            // 解析成本单价
                            if (decimal.TryParse(Convert.ToString(valMatrix[r, 10]), out decimal cp)) costUnitPrice = cp;

                            // K 列 (索引 11): 成本总价
                            decimal costTotalPrice = 0;
                            // 解析成本总价
                            if (decimal.TryParse(Convert.ToString(valMatrix[r, 11]), out decimal ctp)) costTotalPrice = ctp;
                            // 缺省时自动计算成本总价
                            else costTotalPrice = qty * costUnitPrice;

                            // L 列 (索引 12): 报出系数 / 加价系数
                            decimal markupFactor = 1.0m;
                            // 解析报出系数值
                            if (decimal.TryParse(Convert.ToString(valMatrix[r, 12]), out decimal mf) && mf > 0) markupFactor = mf;
                            // 若未指定则根据售价与成本计算加价系数
                            else if (costUnitPrice > 0 && unitPrice > 0) markupFactor = Math.Round(unitPrice / costUnitPrice, 2);

                            // M 列 (索引 13): 表价 / 面价公式与拆分
                            object valM = valMatrix[r, 13];
                            // 获取 M 列公式
                            object formulaM = formulaMatrix[r, 13];
                            // 拆分本体与附件表价
                            ParseBaseAndAccessoryPrice(valM, formulaM, out decimal basePrice, out decimal accPrice);

                            // N 列 (索引 14): 折扣公式与拆分
                            object valN = valMatrix[r, 14];
                            // 获取 N 列公式
                            object formulaN = formulaMatrix[r, 14];
                            // 拆分本体与附件折扣
                            ParseBaseAndAccessoryDiscount(valN, formulaN, basePrice, accPrice, out decimal baseDiscount, out decimal accDiscount);

                            // Q 列 (索引 17): 类别
                            string compCategory = Convert.ToString(valMatrix[r, 17]) ?? string.Empty;
                            // 默认分类名称为“元件” --硬编码--
                            if (string.IsNullOrWhiteSpace(compCategory)) compCategory = "元件";

                            // U 列 (索引 21): 原始型号 (分类明细表中该元件的原始信息)
                            string uColModel = string.Empty;
                            // 安全检查矩阵第二维宽度是否达到 21 列
                            if (valMatrix.GetLength(1) >= 21)
                            {
                                // 读取 U 列单元格文本值
                                uColModel = Convert.ToString(valMatrix[r, 21]) ?? string.Empty;
                            }

                            // W 列 (索引 23): Current 额定电流
                            string currentVal = string.Empty;
                            if (valMatrix.GetLength(1) >= 23)
                            {
                                currentVal = Convert.ToString(valMatrix[r, 23])?.Trim() ?? string.Empty;
                            }

                            // X 列 (索引 24): Poles 极数
                            string polesVal = string.Empty;
                            if (valMatrix.GetLength(1) >= 24)
                            {
                                polesVal = Convert.ToString(valMatrix[r, 24])?.Trim() ?? string.Empty;
                            }

                            // Y 列 (索引 25): trip 脱扣方式
                            string tripVal = string.Empty;
                            if (valMatrix.GetLength(1) >= 25)
                            {
                                tripVal = Convert.ToString(valMatrix[r, 25])?.Trim() ?? string.Empty;
                            }

                            // Z 列 (索引 26): Accessory 配套附件
                            string accVal = string.Empty;
                            if (valMatrix.GetLength(1) >= 26)
                            {
                                accVal = Convert.ToString(valMatrix[r, 26])?.Trim() ?? string.Empty;
                            }

                            // AA 列 (索引 27): BlockName 扩展参数1 / 图块名称
                            string param1Val = string.Empty;
                            if (valMatrix.GetLength(1) >= 27)
                            {
                                param1Val = Convert.ToString(valMatrix[r, 27])?.Trim() ?? string.Empty;
                            }

                            // AB 列 (索引 28): BlockCategory 扩展参数2 / 图块类别
                            string param2Val = string.Empty;
                            if (valMatrix.GetLength(1) >= 28)
                            {
                                param2Val = Convert.ToString(valMatrix[r, 28])?.Trim() ?? string.Empty;
                            }

                            // AD 列 (索引 30): CAD 句柄
                            string handleVal = string.Empty;
                            if (valMatrix.GetLength(1) >= 30)
                            {
                                handleVal = Convert.ToString(valMatrix[r, 30])?.Trim() ?? string.Empty;
                            }

                            // 计算经过箱柜台数放大后的实际总数量与实际总金额
                            decimal totalQty = qty * cabQty;
                            // 计算放大后的销售总金额
                            decimal finalTotalPrice = totalP * cabQty;
                            // 计算放大后的成本总金额
                            decimal finalCostTotalPrice = costTotalPrice * cabQty;

                            // 将提取到的元件明细添加至临时列表
                            rawComponents.Add(new AggregatedComponent
                            {
                                Name = compName.Trim(),
                                OriginalModel = model.Trim(), // 记录分类表中的原始型号规格，供基准对照
                                RawModelFromU = uColModel.Trim(), // 记录明细表 U 列原始型号
                                Model = model.Trim(), // 当前可用于汇总调价的型号规格
                                Manufacturer = mfg.Trim(),
                                Unit = unit.Trim(),
                                Quantity = totalQty,
                                UnitPrice = unitPrice,
                                TotalPrice = finalTotalPrice,
                                CostUnitPrice = costUnitPrice,
                                CostTotalPrice = finalCostTotalPrice,
                                MarkupFactor = markupFactor,
                                BasePrice = basePrice,
                                BaseDiscount = baseDiscount,
                                AccessoryPrice = accPrice,
                                AccessoryDiscount = accDiscount,
                                Current = currentVal,
                                Poles = polesVal,
                                Tripping = tripVal,
                                Accessory = accVal,
                                Param1 = param1Val,
                                Param2 = param2Val,
                                Handle = handleVal,
                                Remark = remark.Trim(),
                                Category = compCategory.Trim(),
                                SheetName = sheetName
                            });
                        }
                    }
                }

                // 内存合并聚合
                var mergeConditions = request.MergeConditions;
                // 根据前端指定的合并条件进行分组
                var grouped = rawComponents.GroupBy(c =>
                {
                    // 构建分组键片段列表
                    var keyParts = new List<string>();
                    // 按名称合并 (支持勾选与取消)
                    if (mergeConditions.ByName)
                    {
                        keyParts.Add(c.Name);
                    }
                    // 按型号合并
                    if (mergeConditions.ByModel)
                    {
                        keyParts.Add(c.Model);
                    }
                    // 按厂家合并
                    if (mergeConditions.ByManufacturer)
                    {
                        keyParts.Add((mergeConditions.IncludeNoManufacturer && string.IsNullOrWhiteSpace(c.Manufacturer)) ? "" : c.Manufacturer);
                    }
                    // 按价格合并
                    if (mergeConditions.ByPrice)
                    {
                        keyParts.Add(c.UnitPrice.ToString());
                    }
                    // 按备注合并
                    if (mergeConditions.ByRemark)
                    {
                        keyParts.Add(c.Remark);
                    }
                    // 按原图型号合并
                    if (mergeConditions.ByOriginalModel)
                    {
                        keyParts.Add(c.OriginalModel);
                    }
                    // 兜底：若所有选项均未勾选，默认按型号分组
                    if (keyParts.Count == 0)
                    {
                        keyParts.Add(c.Model);
                    }
                    // 拼接完整分组键
                    return string.Join("||", keyParts);
                }).Select(g =>
                {
                    // 获取组内首个元素
                    var first = g.First();
                    // 汇总总数量
                    decimal sumQty = g.Sum(x => x.Quantity);
                    // 汇总销售总额
                    decimal sumTotal = g.Sum(x => x.TotalPrice);
                    // 汇总成本总额
                    decimal sumCost = g.Sum(x => x.CostTotalPrice);
                    // 加权平均单价计算
                    decimal avgPrice = sumQty > 0 ? Math.Round(sumTotal / sumQty, 2) : first.UnitPrice;
                    // 加权平均成本单价计算
                    decimal avgCostPrice = sumQty > 0 ? Math.Round(sumCost / sumQty, 2) : first.CostUnitPrice;
                    // 加权报出系数计算
                    decimal markupFactor = avgCostPrice > 0 ? Math.Round(avgPrice / avgCostPrice, 2) : first.MarkupFactor;

                    // 提取所有不为空的明细 U 列原始型号，去重并以全角逗号“，”连接 (需求规定)
                    var distinctUModels = g.Select(x => x.RawModelFromU)
                        .Where(u => !string.IsNullOrWhiteSpace(u))
                        .Distinct()
                        .ToList();
                    // 用全角逗号拼接所有不重复的原始型号
                    string combinedUModel = string.Join("，", distinctUModels);

                    // 构建聚合后的元件实体
                    return new AggregatedComponent
                    {
                        Name = first.Name,
                        OriginalModel = first.OriginalModel, // 保留原型号规格基准
                        RawModelFromU = combinedUModel, // 原始型号 (U 列多项逗号拼接)
                        Model = first.Model,
                        Manufacturer = first.Manufacturer,
                        Unit = first.Unit,
                        Quantity = sumQty,
                        UnitPrice = avgPrice,
                        TotalPrice = sumTotal,
                        CostUnitPrice = avgCostPrice,
                        CostTotalPrice = sumCost,
                        MarkupFactor = markupFactor,
                        BasePrice = first.BasePrice,
                        BaseDiscount = first.BaseDiscount,
                        AccessoryPrice = first.AccessoryPrice,
                        AccessoryDiscount = first.AccessoryDiscount,
                        Current = g.Select(x => x.Current).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty,
                        Poles = g.Select(x => x.Poles).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty,
                        Tripping = g.Select(x => x.Tripping).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty,
                        Accessory = g.Select(x => x.Accessory).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty,
                        Param1 = g.Select(x => x.Param1).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty,
                        Param2 = g.Select(x => x.Param2).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty,
                        Handle = g.Select(x => x.Handle).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty,
                        Remark = string.Join(";", g.Select(x => x.Remark).Where(r => !string.IsNullOrWhiteSpace(r)).Distinct()),
                        Category = string.IsNullOrWhiteSpace(first.Category) ? "元件" : first.Category,
                        SheetName = first.SheetName
                    };
                }).ToList();

                // 排序处理
                if (request.SortSettings.SortType == "mfg_name_model")
                {
                    // 按厂家、名称、型号排序
                    grouped = grouped.OrderBy(x => x.Manufacturer).ThenBy(x => x.Name).ThenBy(x => x.Model).ToList();
                }
                else
                {
                    // 默认按工作表、名称、厂家、型号排序
                    grouped = grouped.OrderBy(x => x.SheetName).ThenBy(x => x.Name).ThenBy(x => x.Manufacturer).ThenBy(x => x.Model).ToList();
                }

                // 创建或清空“元件汇总表”
                string summarySheetName = "元件汇总表"; // --硬编码--
                dynamic summarySheet = null;
                try
                {
                    // 尝试获取已有的“元件汇总表”
                    summarySheet = activeWb.Worksheets[summarySheetName];
                    // 清空既有内容与样式
                    summarySheet.Cells.Clear();
                }
                catch
                {
                    // 不存在则在工作簿最后新增工作表
                    summarySheet = activeWb.Worksheets.Add(After: activeWb.Worksheets[activeWb.Worksheets.Count]);
                    // 重命名为“元件汇总表”
                    summarySheet.Name = summarySheetName;
                }

                // 写入大标题 (扩展为覆盖 A1:R1 共 18 列)
                // 持久化保存本次汇总的选表范围与合并条件配置至隐藏单元格 AB1 (供一键更新精准对齐)
                try
                {
                    var buildConfig = new SummaryBuildConfig
                    {
                        SelectedSheets = request.SelectedSheets ?? new List<string>(),
                        MergeConditions = request.MergeConditions ?? new Controllers.MergeConditionsDto()
                    };
                    string jsonConfig = System.Text.Json.JsonSerializer.Serialize(buildConfig);
                    summarySheet.Range["AB1"].Value2 = jsonConfig;
                }
                catch { }
                summarySheet.Cells[1, 1].Value = "元件汇总调价清单";
                // 合并大标题区域
                dynamic titleRange = summarySheet.Range["A1:R1"];
                titleRange.Merge();
                // 设置标题字号与加粗
                titleRange.Font.Size = 14;
                titleRange.Font.Bold = true;
                // 居中对齐
                titleRange.HorizontalAlignment = -4108; // 居中 xlCenter --硬编码--

                // 设置表头第 4 行与第 5 行行高
                summarySheet.Range["4:4"].RowHeight = 22; // --硬编码--
                summarySheet.Range["5:5"].RowHeight = 22; // --硬编码--

                // 合并第 4~5 行的常规列 (A~J 列)
                // A 列: 序号
                summarySheet.Range["A4:A5"].Merge();
                summarySheet.Cells[4, 1].Value = "序号";

                // B 列: 元件名称
                summarySheet.Range["B4:B5"].Merge();
                summarySheet.Cells[4, 2].Value = "元件名称";

                // C 列: 原型号规格 (放置在元件名称后，基准参考)
                summarySheet.Range["C4:C5"].Merge();
                summarySheet.Cells[4, 3].Value = "原型号规格";

                // D 列: 型号规格 (调价列，允许工程师修改替换)
                summarySheet.Range["D4:D5"].Merge();
                summarySheet.Cells[4, 4].Value = "型号规格";

                // E 列: 单位
                summarySheet.Range["E4:E5"].Merge();
                summarySheet.Cells[4, 5].Value = "单位";

                // F 列: 数量
                summarySheet.Range["F4:F5"].Merge();
                summarySheet.Cells[4, 6].Value = "数量";

                // G 列: 单价
                summarySheet.Range["G4:G5"].Merge();
                summarySheet.Cells[4, 7].Value = "单价";

                // H 列: 总价
                summarySheet.Range["H4:H5"].Merge();
                summarySheet.Cells[4, 8].Value = "总价";

                // I 列: 生产厂家
                summarySheet.Range["I4:I5"].Merge();
                summarySheet.Cells[4, 9].Value = "生产厂家";

                // J 列: 成本单价
                summarySheet.Range["J4:J5"].Merge();
                summarySheet.Cells[4, 10].Value = "成本单价";

                // K 列: 报出系数 (合并第 4~5 行，粉色表头)
                summarySheet.Range["K4:K5"].Merge();
                summarySheet.Cells[4, 11].Value = "报出系数";

                // L~M 列: 本体复合表头 (第 4 行合并“本体”，第 5 行“表价”与“折扣”)
                summarySheet.Range["L4:M4"].Merge();
                summarySheet.Cells[4, 12].Value = "本体";
                summarySheet.Cells[5, 12].Value = "表价";
                summarySheet.Cells[5, 13].Value = "折扣";

                // N~O 列: 附件复合表头 (第 4 行合并“附件”，第 5 行“表价”与“折扣”)
                summarySheet.Range["N4:O4"].Merge();
                summarySheet.Cells[4, 14].Value = "附件";
                summarySheet.Cells[5, 14].Value = "表价";
                summarySheet.Cells[5, 15].Value = "折扣";

                // P 列: 备注
                summarySheet.Range["P4:P5"].Merge();
                summarySheet.Cells[4, 16].Value = "备注";

                // Q 列: 类别
                summarySheet.Range["Q4:Q5"].Merge();
                summarySheet.Cells[4, 17].Value = "类别";

                // R 列: 原始型号 (合并第 4~5 行，对应明细表 U 列内容，多项逗号拼接)
                summarySheet.Range["R4:R5"].Merge();
                summarySheet.Cells[4, 18].Value = "原始型号";

                // T 列: 额定电流 (第 20 列)
                summarySheet.Range["T4:T5"].Merge();
                summarySheet.Cells[4, 20].Value = "额定电流";

                // U 列: 极数 (第 21 列)
                summarySheet.Range["U4:U5"].Merge();
                summarySheet.Cells[4, 21].Value = "极数";

                // V 列: 脱扣方式 (第 22 列)
                summarySheet.Range["V4:V5"].Merge();
                summarySheet.Cells[4, 22].Value = "脱扣方式";

                // W 列: 附件 (第 23 列)
                summarySheet.Range["W4:W5"].Merge();
                summarySheet.Cells[4, 23].Value = "附件";

                // X 列: 参数1 (第 24 列)
                summarySheet.Range["X4:X5"].Merge();
                summarySheet.Cells[4, 24].Value = "参数1";

                // Y 列: 参数2 (第 25 列)
                summarySheet.Range["Y4:Y5"].Merge();
                summarySheet.Cells[4, 25].Value = "参数2";

                // 设置全表头基础文字与对齐样式 (覆盖 A4:R5 与 T4:Y5 共 24 列)
                dynamic headerAllRange = summarySheet.Range["A4:R5"];
                headerAllRange.Font.Name = "宋体"; // --硬编码--
                headerAllRange.Font.Size = 10; // --硬编码--
                headerAllRange.Font.Bold = true;
                headerAllRange.HorizontalAlignment = -4108; // 居中 xlCenter --硬编码--
                headerAllRange.VerticalAlignment = -4108; // 居中 xlCenter --硬编码--
                headerAllRange.WrapText = true;

                // 设置扩展表头基础样式 (T4:Y5)
                dynamic headerExtRange = summarySheet.Range["T4:Y5"];
                headerExtRange.Font.Name = "宋体"; // --硬编码--
                headerExtRange.Font.Size = 10; // --硬编码--
                headerExtRange.Font.Bold = true;
                headerExtRange.HorizontalAlignment = -4108; // 居中
                headerExtRange.VerticalAlignment = -4108;
                headerExtRange.WrapText = true;

                // 常规列灰色背景 (A4:J5, P4:R5 与 T4:Y5) --硬编码--
                int headerGrayColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#E8ECEF"));
                summarySheet.Range["A4:J5"].Interior.Color = headerGrayColor;
                summarySheet.Range["P4:R5"].Interior.Color = headerGrayColor;
                summarySheet.Range["T4:Y5"].Interior.Color = headerGrayColor;

                // 调价核心列粉色背景 (K4:O5，包含报出系数、本体与附件) --硬编码--
                int headerPinkColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#E6A8A8"));
                summarySheet.Range["K4:O5"].Interior.Color = headerPinkColor;

                // 柔和浅蓝色强调底色 (基准参考列专属色) --硬编码--
                int softBlueColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#A8C7EB"));

                // 原型号规格表头 (C4:C5) 显式设置为浅蓝色底色，与后续数据列保持蓝底一致风格
                summarySheet.Range["C4:C5"].Interior.Color = softBlueColor;

                // 批量回写 18 列数据 (从第 6 行开始)
                int dataRowCount = grouped.Count;
                int startDataRow = 6;
                int endDataRow = startDataRow + dataRowCount - 1;

                if (dataRowCount > 0)
                {
                    // 构建 18 列数据输出矩阵
                    object[,] outMatrix = new object[dataRowCount, 18];
                    // 构建 T~Y 列扩展参数输出矩阵 (6 列) 与 AD 列 CAD 句柄矩阵
                    object[,] paramMatrix = new object[dataRowCount, 6];
                    object[,] handleMatrix = new object[dataRowCount, 1];

                    for (int i = 0; i < dataRowCount; i++)
                    {
                        var comp = grouped[i];
                        // 计算当前数据行在 Excel 中的实际物理行号 (从 startDataRow 起算)
                        int currentRow = startDataRow + i;
                        outMatrix[i, 0] = i + 1; // A: 序号
                        outMatrix[i, 1] = comp.Name; // B: 元件名称
                        outMatrix[i, 2] = comp.OriginalModel; // C: 原型号规格 (基准参考列)
                        // D: 型号规格 (汇总生成时保持留空，不填内容)
                        outMatrix[i, 3] = string.Empty;
                        outMatrix[i, 4] = comp.Unit; // E: 单位
                        outMatrix[i, 5] = (double)comp.Quantity; // F: 数量
                        // G: 单价 (采用公式计算: 报出系数 * (本体表价 * 本体折扣 + 附件表价 * 附件折扣))
                        outMatrix[i, 6] = $"=ROUND(K{currentRow}*(L{currentRow}*M{currentRow}+N{currentRow}*O{currentRow}),2)";
                        // H: 总价 (采用公式计算: 数量 * 单价，F列为数量，G列为单价)
                        outMatrix[i, 7] = $"=ROUND(F{currentRow}*G{currentRow},2)";
                        outMatrix[i, 8] = comp.Manufacturer; // I: 生产厂家
                        // J: 成本单价 (采用公式计算: 本体表价 * 本体折扣 + 附件表价 * 附件折扣)
                        outMatrix[i, 9] = $"=ROUND(L{currentRow}*M{currentRow}+N{currentRow}*O{currentRow},2)";
                        outMatrix[i, 10] = (double)comp.MarkupFactor; // K: 报出系数
                        outMatrix[i, 11] = comp.BasePrice > 0 ? (object)(double)comp.BasePrice : string.Empty; // L: 本体表价
                        outMatrix[i, 12] = comp.BaseDiscount > 0 ? (object)(double)comp.BaseDiscount : 1; // M: 本体折扣
                        outMatrix[i, 13] = comp.AccessoryPrice > 0 ? (object)(double)comp.AccessoryPrice : string.Empty; // N: 附件表价
                        outMatrix[i, 14] = (comp.AccessoryPrice > 0 && comp.AccessoryDiscount > 0) ? (object)(double)comp.AccessoryDiscount : string.Empty; // O: 附件折扣
                        outMatrix[i, 15] = comp.Remark; // P: 备注
                        outMatrix[i, 16] = comp.Category; // Q: 类别
                        outMatrix[i, 17] = comp.RawModelFromU; // R: 原始型号 (对应明细表 U 列内容，逗号拼接)

                        // T~Y 列扩展参数
                        paramMatrix[i, 0] = comp.Current;  // T 列: 额定电流
                        paramMatrix[i, 1] = comp.Poles;    // U 列: 极数
                        paramMatrix[i, 2] = comp.Tripping; // V 列: 脱扣方式
                        paramMatrix[i, 3] = string.Empty;  // W 列: 附件 (默认留空供选附件时写入脱扣方式)
                        paramMatrix[i, 4] = comp.Param1;   // X 列: 参数1
                        paramMatrix[i, 5] = comp.Param2;   // Y 列: 参数2

                        // AD 列 CAD 句柄
                        handleMatrix[i, 0] = comp.Handle;
                    }

                    // 一次性批量写入数据与公式矩阵 (覆盖 A6:R{endDataRow}) (规则 7)
                    summarySheet.Range[$"A{startDataRow}:R{endDataRow}"].Formula = outMatrix;

                    // 一次性批量写入 T~Y 列扩展参数与 AD 列句柄矩阵 (规则 7)
                    summarySheet.Range[$"T{startDataRow}:Y{endDataRow}"].Value2 = paramMatrix;
                    summarySheet.Range[$"AD{startDataRow}:AD{endDataRow}"].Value2 = handleMatrix;

                    // 设置数据行行高
                    summarySheet.Range[$"{startDataRow}:{endDataRow}"].RowHeight = 20; // --硬编码--

                    // 设置单元格对齐方式 (完整对应 18 列布局)
                    summarySheet.Range[$"A{startDataRow}:A{endDataRow}"].HorizontalAlignment = -4108; // A 序号居中
                    summarySheet.Range[$"B{startDataRow}:D{endDataRow}"].HorizontalAlignment = -4131; // B~D 名称、原型号、型号靠左
                    summarySheet.Range[$"E{startDataRow}:F{endDataRow}"].HorizontalAlignment = -4108; // E~F 单位、数量居中
                    summarySheet.Range[$"G{startDataRow}:H{endDataRow}"].HorizontalAlignment = -4152; // G~H 单价、总价靠右
                    summarySheet.Range[$"I{startDataRow}:I{endDataRow}"].HorizontalAlignment = -4131; // I 厂家靠左
                    summarySheet.Range[$"J{startDataRow}:J{endDataRow}"].HorizontalAlignment = -4152; // J 成本单价靠右
                    summarySheet.Range[$"K{startDataRow}:K{endDataRow}"].HorizontalAlignment = -4108; // K 报出系数居中
                    summarySheet.Range[$"L{startDataRow}:L{endDataRow}"].HorizontalAlignment = -4152; // L 本体表价靠右
                    summarySheet.Range[$"M{startDataRow}:M{endDataRow}"].HorizontalAlignment = -4108; // M 本体折扣居中
                    summarySheet.Range[$"N{startDataRow}:N{endDataRow}"].HorizontalAlignment = -4152; // N 附件表价靠右
                    summarySheet.Range[$"O{startDataRow}:O{endDataRow}"].HorizontalAlignment = -4108; // O 附件折扣居中
                    summarySheet.Range[$"P{startDataRow}:P{endDataRow}"].HorizontalAlignment = -4131; // P 备注靠左
                    summarySheet.Range[$"Q{startDataRow}:Q{endDataRow}"].HorizontalAlignment = -4108; // Q 类别居中
                    summarySheet.Range[$"R{startDataRow}:R{endDataRow}"].HorizontalAlignment = -4131; // R 原始型号靠左

                    // T~V 列对齐居中，W~Y 列靠左
                    summarySheet.Range[$"T{startDataRow}:V{endDataRow}"].HorizontalAlignment = -4108; // 电流、极数、脱扣居中
                    summarySheet.Range[$"W{startDataRow}:Y{endDataRow}"].HorizontalAlignment = -4131; // 附件、参数1、参数2靠左

                    // 设置基准参考列的浅蓝色强调底色 (序号 A 列、原型号规格 C 列、数量 F 列、成本单价 J 列) --硬编码--
                    summarySheet.Range[$"A{startDataRow}:A{endDataRow}"].Interior.Color = softBlueColor;
                    summarySheet.Range[$"C{startDataRow}:C{endDataRow}"].Interior.Color = softBlueColor; // 原型号规格蓝色底色
                    summarySheet.Range[$"F{startDataRow}:F{endDataRow}"].Interior.Color = softBlueColor; // 数量列蓝色底色
                    summarySheet.Range[$"J{startDataRow}:J{endDataRow}"].Interior.Color = softBlueColor; // 成本单价蓝色底色

                    // 设置各数值列的显示格式（双轨安全容错机制，兼容中英文 Excel） --硬编码--
                    // F 列数量格式：优先使用中文本地化通用格式 G/通用格式，容错回退 General
                    try { summarySheet.Range[$"F{startDataRow}:F{endDataRow}"].NumberFormatLocal = "G/通用格式"; }
                    catch { try { summarySheet.Range[$"F{startDataRow}:F{endDataRow}"].NumberFormat = "General"; } catch { } }

                    // G~H 列：销售单价与总价格式
                    try { summarySheet.Range[$"G{startDataRow}:H{endDataRow}"].NumberFormat = "#,##0.00"; } catch { }
                    // J 列：成本单价格式
                    try { summarySheet.Range[$"J{startDataRow}:J{endDataRow}"].NumberFormat = "#,##0.00"; } catch { }
                    // K 列：报出系数格式
                    try { summarySheet.Range[$"K{startDataRow}:K{endDataRow}"].NumberFormat = "0.00"; } catch { }
                    // L 列：本体表价格式
                    try { summarySheet.Range[$"L{startDataRow}:L{endDataRow}"].NumberFormat = "#,##0.00"; } catch { }
                    // M 列：本体折扣格式
                    try { summarySheet.Range[$"M{startDataRow}:M{endDataRow}"].NumberFormat = "0.####"; } catch { }
                    // N 列：附件表价格式
                    try { summarySheet.Range[$"N{startDataRow}:N{endDataRow}"].NumberFormat = "#,##0.00"; } catch { }
                    // O 列：附件折扣格式
                    try { summarySheet.Range[$"O{startDataRow}:O{endDataRow}"].NumberFormat = "0.####"; } catch { }

                    // 写入底部合计行
                    int totalRow = endDataRow + 1;
                    summarySheet.Range[$"{totalRow}:{totalRow}"].RowHeight = 22; // --硬编码--
                    // 合计行标签写入 B 列 (元件名称)
                    summarySheet.Cells[totalRow, 2].Value = "合计";
                    summarySheet.Cells[totalRow, 2].Font.Bold = true;
                    summarySheet.Cells[totalRow, 2].HorizontalAlignment = -4108; // 居中

                    // 数量总计公式写入 F 列
                    summarySheet.Cells[totalRow, 6].Formula = $"=SUM(F{startDataRow}:F{endDataRow})";
                    summarySheet.Cells[totalRow, 6].Font.Bold = true;
                    summarySheet.Cells[totalRow, 6].HorizontalAlignment = -4108;
                    // 合计行 F 列数量格式
                    try { summarySheet.Cells[totalRow, 6].NumberFormatLocal = "G/通用格式"; }
                    catch { try { summarySheet.Cells[totalRow, 6].NumberFormat = "General"; } catch { } }

                    // 销售总价合计公式写入 H 列
                    summarySheet.Cells[totalRow, 8].Formula = $"=SUM(H{startDataRow}:H{endDataRow})";
                    summarySheet.Cells[totalRow, 8].Font.Bold = true;
                    summarySheet.Cells[totalRow, 8].HorizontalAlignment = -4152;
                    try { summarySheet.Cells[totalRow, 8].NumberFormat = "#,##0.00"; } catch { }

                    // 设置表格全区域边框 (覆盖 A4:R{totalRow} 与 U4:Y{endDataRow}) --硬编码--
                    dynamic tableRange = summarySheet.Range[$"A4:R{totalRow}"];
                    tableRange.Borders.LineStyle = 1; // 实线 xlContinuous
                    tableRange.Borders.Weight = 2; // 细线 xlThin

                    dynamic tableExtRange = summarySheet.Range[$"T4:Y{endDataRow}"];
                    tableExtRange.Borders.LineStyle = 1;
                    tableExtRange.Borders.Weight = 2;
                }
                else
                {
                    // 若无数据，至少设置表头区域边框 (A4:R5 与 T4:Y5)
                    dynamic tableRange = summarySheet.Range["A4:R5"];
                    tableRange.Borders.LineStyle = 1;
                    tableRange.Borders.Weight = 2;

                    dynamic tableExtRange = summarySheet.Range["T4:Y5"];
                    tableExtRange.Borders.LineStyle = 1;
                    tableExtRange.Borders.Weight = 2;
                }

                // 显式精准设置各列宽度（保证界面美观无挤压，注意宽度） --硬编码--
                summarySheet.Columns[1].ColumnWidth = 6;   // A: 序号
                summarySheet.Columns[2].ColumnWidth = 18;  // B: 元件名称
                summarySheet.Columns[3].ColumnWidth = 24;  // C: 原型号规格 (基准参考列)
                summarySheet.Columns[4].ColumnWidth = 24;  // D: 型号规格
                summarySheet.Columns[5].ColumnWidth = 6;   // E: 单位
                summarySheet.Columns[6].ColumnWidth = 8;   // F: 数量
                summarySheet.Columns[7].ColumnWidth = 14;  // G: 单价
                summarySheet.Columns[8].ColumnWidth = 15;  // H: 总价
                summarySheet.Columns[9].ColumnWidth = 18;  // I: 生产厂家
                summarySheet.Columns[10].ColumnWidth = 14; // J: 成本单价
                summarySheet.Columns[11].ColumnWidth = 9;  // K: 报出系数
                summarySheet.Columns[12].ColumnWidth = 12; // L: 本体表价
                summarySheet.Columns[13].ColumnWidth = 8;  // M: 本体折扣
                summarySheet.Columns[14].ColumnWidth = 12; // N: 附件表价
                summarySheet.Columns[15].ColumnWidth = 8;  // O: 附件折扣
                summarySheet.Columns[16].ColumnWidth = 18; // P: 备注
                summarySheet.Columns[17].ColumnWidth = 8;  // Q: 类别
                summarySheet.Columns[18].ColumnWidth = 24; // R: 原始型号 (明细表 U 列多项拼接)
                summarySheet.Columns[20].ColumnWidth = 10; // T: 额定电流
                summarySheet.Columns[21].ColumnWidth = 8;  // U: 极数
                summarySheet.Columns[22].ColumnWidth = 10; // V: 脱扣方式
                summarySheet.Columns[23].ColumnWidth = 14; // W: 附件
                summarySheet.Columns[24].ColumnWidth = 12; // X: 参数1
                summarySheet.Columns[25].ColumnWidth = 12; // Y: 参数2

                // 挂载 AutoFilter 自动筛选下拉箭头至第 5 行 (A5:Y5)
                try
                {
                    summarySheet.Range["A5:Y5"].AutoFilter();
                }
                catch { }

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
        /// 一键更新同步操作结果实体模型
        /// </summary>
        /// <summary>
        /// 汇总表构建与回写配置持久化模型
        /// </summary>
        public class SummaryBuildConfig
        {
            // 参与汇总的分类表名称列表
            public List<string> SelectedSheets { get; set; } = new List<string>();

            // 汇总时的合并条件
            public Controllers.MergeConditionsDto MergeConditions { get; set; } = new Controllers.MergeConditionsDto();
        }

        public class SummaryUpdateResult
        {
            // 是否成功
            public bool Success { get; set; } = false;

            // 结果提示消息
            public string Message { get; set; } = string.Empty;

            // 成功更新的分类工作表总数
            public int UpdatedSheetCount { get; set; } = 0;

            // 成功更新的箱柜总台数
            public int UpdatedCabinetCount { get; set; } = 0;

            // 成功更新的元器件明细项总数
            public int UpdatedComponentCount { get; set; } = 0;
        }

        /// <summary>
        /// 汇总调价反向同步单项数据模型
        /// </summary>
        private class SummaryAdjustItem
        {
            // 单位 (E 列)
            public string Unit { get; set; } = string.Empty;
            // 元件名称
            public string Name { get; set; } = string.Empty;

            // 原型号规格 (基准匹配键)
            public string OriginalModel { get; set; } = string.Empty;

            // 新型号规格 (D 列修改选型)
            public string NewModel { get; set; } = string.Empty;

            // 生产厂家
            public string Manufacturer { get; set; } = string.Empty;

            // 报出系数
            public decimal MarkupFactor { get; set; } = 1.0m;

            // 回写至分类表 M 列（表价）的内容（公式或数值）
            public object MColContent { get; set; } = string.Empty;

            // 回写至分类表 N 列（折扣）的内容（数值或公式）
            public object NColContent { get; set; } = 1.0m;

            // 备注说明
            public string Remark { get; set; } = string.Empty;

            // 额定电流 (T 列)
            public string Current { get; set; } = string.Empty;

            // 极数 (U 列)
            public string Poles { get; set; } = string.Empty;

            // 脱扣方式 (V 列)
            public string Tripping { get; set; } = string.Empty;

            // 附件说明与脱扣 (W 列)
            public string Accessory { get; set; } = string.Empty;

            // 扩展参数 1 (X 列)
            public string Param1 { get; set; } = string.Empty;

            // 扩展参数 2 (Y 列)
            public string Param2 { get; set; } = string.Empty;

            // U 列原始型号字符串 (多项逗号拼接)
            public string RawModelFromU { get; set; } = string.Empty;
        }

        /// <summary>
        /// 一键更新：从“元件汇总表”将修改后的型号规格、厂家、报出系数、面价公式、折扣、备注等反向同步更新至各分类表箱柜明细块
        /// </summary>
        public static SummaryUpdateResult UpdateFromComponentSummarySheet()
        {
            var result = new SummaryUpdateResult();
            try
            {
                // 获取当前活动工作簿
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    result.Message = "无法连接到 Excel 应用程序实例";
                    return result;
                }

                dynamic wb = app.ActiveWorkbook;
                if (wb == null)
                {
                    result.Message = "当前未打开任何工作簿";
                    return result;
                }

                // 查找是否存在“元件汇总表”工作表
                dynamic? summarySheet = null;
                foreach (dynamic ws in wb.Worksheets)
                {
                    string wsName = Convert.ToString(ws.Name) ?? string.Empty;
                    if (string.Equals(wsName.Trim(), ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        summarySheet = ws;
                        break;
                    }
                }

                if (summarySheet == null)
                {
                    result.Message = $"未找到名为【{ComponentMatchDefaults.ComponentSummarySheetName}】的工作表";
                    return result;
                }

                // 尝试从隐藏单元格 AB1 读取持久化的构建配置（参与汇总工作表列表与合并条件）
                var buildConfig = new SummaryBuildConfig();
                try
                {
                    // 读取 AB1 单元格中的 JSON 配置字符串
                    string rawConfig = Convert.ToString(summarySheet.Range["AB1"].Value2)?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(rawConfig))
                    {
                        // 反序列化配置模型
                        var cfg = System.Text.Json.JsonSerializer.Deserialize<SummaryBuildConfig>(rawConfig);
                        if (cfg != null)
                        {
                            // 赋值生效配置
                            buildConfig = cfg;
                        }
                    }
                }
                catch { }

                // 1. 从“元件汇总表”收集所有调价记录
                int startDataRow = 6; // 数据起始行
                int summaryUsedRows = summarySheet.UsedRange.Rows.Count;
                // 数据读取最大截止行
                int maxReadRow = Math.Max(summaryUsedRows + 5, 20);

                // 一次性批量读取汇总表 A6:Y{maxReadRow} 的值矩阵与公式矩阵 (覆盖 A~Y 共 25 列，规则 7)
                dynamic summaryRange = summarySheet.Range[$"A{startDataRow}:Y{maxReadRow}"];
                object[,] summaryValMatrix = (object[,])summaryRange.Value2;
                object[,] summaryFormulaMatrix = (object[,])summaryRange.Formula;

                if (summaryValMatrix == null || summaryFormulaMatrix == null)
                {
                    result.Message = "读取【元件汇总表】数据矩阵失败";
                    return result;
                }

                int totalSummaryRows = summaryValMatrix.GetLength(0);
                var adjustItems = new List<SummaryAdjustItem>();

                for (int r = 1; r <= totalSummaryRows; r++)
                {
                    // B 列 (索引 2): 元件名称
                    string name = Convert.ToString(summaryValMatrix[r, 2])?.Trim() ?? string.Empty;
                    // 若名称包含“合计”或遇到连续空行则终止读取 --硬编码--
                    if (name.Contains("合计"))
                    {
                        break;
                    }

                    // C 列 (索引 3): 原型号规格 (基准参考列)
                    string origModel = Convert.ToString(summaryValMatrix[r, 3])?.Trim() ?? string.Empty;

                    // 若名称与原型号均为空，跳过空行
                    if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(origModel))
                    {
                        continue;
                    }

                    // D 列 (索引 4): 新型号规格 (工程师修改选型或追加附件后的内容)
                    string newModel = Convert.ToString(summaryValMatrix[r, 4])?.Trim() ?? string.Empty;
                    // E 列 (索引 5): 单位
                    string unit = Convert.ToString(summaryValMatrix[r, 5])?.Trim() ?? string.Empty;

                    // I 列 (索引 9): 生产厂家
                    string mfr = Convert.ToString(summaryValMatrix[r, 9])?.Trim() ?? string.Empty;

                    // K 列 (索引 11): 报出系数
                    decimal markupFactor = 1.0m;
                    if (decimal.TryParse(Convert.ToString(summaryValMatrix[r, 11]), out decimal mf) && mf > 0)
                    {
                        markupFactor = mf;
                    }

                    // L 列 (索引 12): 本体表价 (值与公式)
                    object valL = summaryValMatrix[r, 12];
                    object formulaL = summaryFormulaMatrix[r, 12];
                    decimal basePrice = 0;
                    if (decimal.TryParse(Convert.ToString(valL), out decimal bp)) basePrice = bp;

                    // M 列 (索引 13): 本体折扣
                    decimal baseDiscount = 1.0m;
                    if (decimal.TryParse(Convert.ToString(summaryValMatrix[r, 13]), out decimal bd) && bd > 0) baseDiscount = bd;

                    // N 列 (索引 14): 附件表价 (值与公式)
                    object valN = summaryValMatrix[r, 14];
                    object formulaN = summaryFormulaMatrix[r, 14];
                    decimal accPrice = 0;
                    if (decimal.TryParse(Convert.ToString(valN), out decimal ap)) accPrice = ap;

                    // O 列 (索引 15): 附件折扣
                    decimal accDiscount = 1.0m;
                    if (decimal.TryParse(Convert.ToString(summaryValMatrix[r, 15]), out decimal ad) && ad > 0) accDiscount = ad;

                    // P 列 (索引 16): 备注
                    string remark = Convert.ToString(summaryValMatrix[r, 16])?.Trim() ?? string.Empty;

                    // R 列 (索引 18): 原始型号 (U 列对应内容，全角逗号拼接)
                    string rawModelU = Convert.ToString(summaryValMatrix[r, 18])?.Trim() ?? string.Empty;

                    // T 列 (索引 20): 额定电流
                    string currentT = string.Empty;
                    if (summaryValMatrix.GetLength(1) >= 20) currentT = Convert.ToString(summaryValMatrix[r, 20])?.Trim() ?? string.Empty;

                    // U 列 (索引 21): 极数
                    string polesU = string.Empty;
                    if (summaryValMatrix.GetLength(1) >= 21) polesU = Convert.ToString(summaryValMatrix[r, 21])?.Trim() ?? string.Empty;

                    // V 列 (索引 22): 脱扣方式
                    string trippingV = string.Empty;
                    if (summaryValMatrix.GetLength(1) >= 22) trippingV = Convert.ToString(summaryValMatrix[r, 22])?.Trim() ?? string.Empty;

                    // W 列 (索引 23): 附件 (附件脱扣方式)
                    string accessoryW = string.Empty;
                    if (summaryValMatrix.GetLength(1) >= 23) accessoryW = Convert.ToString(summaryValMatrix[r, 23])?.Trim() ?? string.Empty;

                    // X 列 (索引 24): 参数1
                    string param1X = string.Empty;
                    if (summaryValMatrix.GetLength(1) >= 24) param1X = Convert.ToString(summaryValMatrix[r, 24])?.Trim() ?? string.Empty;

                    // Y 列 (索引 25): 参数2
                    string param2Y = string.Empty;
                    if (summaryValMatrix.GetLength(1) >= 25) param2Y = Convert.ToString(summaryValMatrix[r, 25])?.Trim() ?? string.Empty;

                    // 构造准备回写到分类表 M 列（表价）与 N 列（折扣）的内容
                    object mColContent;
                    object nColContent;

                    // 判断是否存在有效附件表价
                    string strFormulaN = Convert.ToString(formulaN)?.Trim() ?? string.Empty;
                    bool hasAccessory = accPrice > 0 || (!string.IsNullOrEmpty(strFormulaN) && strFormulaN.StartsWith("="));

                    if (hasAccessory)
                    {
                        // 1. 构建 M 列表价公式：本体表价 + 附件表价 (例如 =159.05+336.2*2)
                        string baseExpr = basePrice > 0 ? basePrice.ToString("0.##") : "0";
                        string strFormulaL = Convert.ToString(formulaL)?.Trim() ?? string.Empty;
                        if (!string.IsNullOrEmpty(strFormulaL) && strFormulaL.StartsWith("="))
                        {
                            baseExpr = strFormulaL.TrimStart('=').Trim();
                        }

                        string accExpr = accPrice > 0 ? accPrice.ToString("0.##") : "0";
                        if (!string.IsNullOrEmpty(strFormulaN) && strFormulaN.StartsWith("="))
                        {
                            accExpr = strFormulaN.TrimStart('=').Trim();
                        }

                        // 组合生成加法公式
                        mColContent = $"={baseExpr}+{accExpr}";

                        // 2. 构建 N 列折扣：生成 =ROUND(({本体表价}*{本体折扣}+{附件表价}*{附件折扣})/{总表价},2) 标准公式样式
                        string baseDiscStr = baseDiscount.ToString("0.####");
                        // 格式化附件折扣字符串
                        string accDiscStr = accDiscount.ToString("0.####");
                        // 计算本体与附件总表价 (分母)
                        decimal totalListPrice = Math.Round(basePrice + accPrice, 2);
                        // 格式化分母总表价字符串
                        string totalListStr = totalListPrice > 0 ? totalListPrice.ToString("0.##") : "1";

                        // 组装本体项 (例如 159.05*0.5)
                        string baseTerm = $"{baseExpr}*{baseDiscStr}";
                        string accTerm;
                        // 若附件表达式包含多个加项 (如 150+80*2)
                        if (accExpr.Contains("+"))
                        {
                            // 拆分多个加项并逐项追加附件折扣乘数
                            string[] accParts = accExpr.Split('+');
                            accTerm = string.Join("+", accParts.Select(p => $"{p.Trim()}*{accDiscStr}"));
                        }
                        else
                        {
                            // 单项附件乘上附件折扣 (如 736.2*1 或 336.2*2*1)
                            accTerm = $"{accExpr}*{accDiscStr}";
                        }

                        // 组合生成完整的标准 ROUND 加权折扣公式
                        nColContent = $"=ROUND(({baseTerm}+{accTerm})/{totalListStr},2)";
                    }
                    else
                    {
                        // 无附件时：直接写入本体表价与本体折扣
                        string strFormulaL = Convert.ToString(formulaL)?.Trim() ?? string.Empty;
                        if (!string.IsNullOrEmpty(strFormulaL) && strFormulaL.StartsWith("="))
                        {
                            mColContent = strFormulaL;
                        }
                        else
                        {
                            mColContent = basePrice > 0 ? (object)(double)basePrice : string.Empty;
                        }

                        nColContent = baseDiscount > 0 ? (object)(double)baseDiscount : 1;
                    }

                    // 添加至调价规则列表
                    adjustItems.Add(new SummaryAdjustItem
                    {
                        Name = name,
                        OriginalModel = origModel,
                        NewModel = newModel,
                        Unit = unit,
                        Manufacturer = mfr,
                        MarkupFactor = markupFactor,
                        MColContent = mColContent,
                        NColContent = nColContent,
                        Remark = remark,
                        Current = currentT,
                        Poles = polesU,
                        Tripping = trippingV,
                        Accessory = accessoryW,
                        Param1 = param1X,
                        Param2 = param2Y,
                        RawModelFromU = rawModelU
                    });
                }

                // 校验是否存在调价项
                if (adjustItems.Count == 0)
                {
                    result.Message = "【元件汇总表】中未提取到有效的元器件调价数据";
                    return result;
                }

                // 2. 准备遍历所有分类工作表执行反向回写
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                int updatedSheetCount = 0;
                int updatedCabinetCount = 0;
                int updatedComponentCount = 0;

                try
                {
                    // 遍历工作簿中的所有工作表
                    foreach (dynamic sheet in wb.Worksheets)
                    {
                        string sheetName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;

                        // 排除系统特殊表 --硬编码--
                        if (string.Equals(sheetName, "项目信息", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(sheetName, "元件汇总表", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(sheetName, "元器件数据管理", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        // 核心：若配置指定了参与汇总的工作表列表，仅处理包含在列表中的工作表 (未选中的工作表保持原样)
                        if (buildConfig.SelectedSheets != null && buildConfig.SelectedSheets.Count > 0)
                        {
                            if (!buildConfig.SelectedSheets.Contains(sheetName))
                            {
                                continue;
                            }
                        }

                        // 获取当前分类表的有效箱柜锚点 (规则 6)
                        var validCabinets = Tool.GetSheetValidCabinets(sheet, wb);
                        if (validCabinets.Count == 0) continue;

                        bool sheetModified = false;

                        // 遍历各个箱柜
                        foreach (var cab in validCabinets)
                        {
                            if (cab.Value.Det == null) continue;

                            int detRow = Convert.ToInt32(cab.Value.Det.Row);
                            int subsumRow = cab.Value.Subsum != null ? Convert.ToInt32(cab.Value.Subsum.Row) : detRow + 27;

                            // 元器件起始行与终止行 (规则 6)
                            int compStartRow = detRow + 2;
                            int compEndRow = subsumRow - 1;
                            if (compEndRow < compStartRow) continue;

                            int cabinetRowCount = compEndRow - compStartRow + 1;

                            // 采用 2D 数组一次性批量读入内存 (覆盖 A~AD 共 30 列，规则 7)
                            dynamic compRange = sheet.Range[$"A{compStartRow}:AD{compEndRow}"];
                            object[,] valMatrix = (object[,])compRange.Value2;
                            object[,] formulaMatrix = (object[,])compRange.Formula;

                            bool cabinetModified = false;

                            // 遍历箱柜元器件每一行
                            for (int r = 1; r <= cabinetRowCount; r++)
                            {
                                // B 列 (索引 2): 元件名称
                                string cName = Convert.ToString(valMatrix[r, 2])?.Trim() ?? string.Empty;
                                // C 列 (索引 3): 型号规格
                                string cModel = Convert.ToString(valMatrix[r, 3])?.Trim() ?? string.Empty;

                                // 跳过空行
                                if (string.IsNullOrWhiteSpace(cName) && string.IsNullOrWhiteSpace(cModel))
                                {
                                    continue;
                                }

                                // U 列 (索引 21): 原始型号
                                string cRawU = string.Empty;
                                if (valMatrix.GetLength(1) >= 21)
                                {
                                    cRawU = Convert.ToString(valMatrix[r, 21])?.Trim() ?? string.Empty;
                                }

                                // 匹配汇总调价规则项：按原型号或 U列原始型号匹配
                                // D 列 (索引 4): 生产厂家
                                string cMfr = Convert.ToString(valMatrix[r, 4])?.Trim() ?? string.Empty;
                                // I 列 (索引 9): 备注
                                string cRemark = Convert.ToString(valMatrix[r, 9])?.Trim() ?? string.Empty;

                                // 根据汇总时的合并条件进行精准多维度匹配
                                var mergeCond = buildConfig.MergeConditions ?? new Controllers.MergeConditionsDto();
                                var matchedItem = adjustItems.FirstOrDefault(item =>
                                {
                                    // 1. 型号基准比对 (支持原型号匹配或 U 列原始型号匹配)
                                    bool modelMatched = string.Equals(item.OriginalModel, cModel, StringComparison.OrdinalIgnoreCase) ||
                                        (!string.IsNullOrEmpty(cRawU) && !string.IsNullOrEmpty(item.RawModelFromU) && item.RawModelFromU.Contains(cRawU));
                                    if (!modelMatched) return false;

                                    // 2. 若汇总时勾选了【按名称合并】：要求名称必须一致
                                    if (mergeCond.ByName)
                                    {
                                        if (!string.Equals(item.Name, cName, StringComparison.OrdinalIgnoreCase))
                                        {
                                            return false;
                                        }
                                    }

                                    // 3. 若汇总时勾选了【按厂家合并】：要求厂家必须一致
                                    if (mergeCond.ByManufacturer)
                                    {
                                        if (mergeCond.IncludeNoManufacturer && string.IsNullOrWhiteSpace(item.Manufacturer) && string.IsNullOrWhiteSpace(cMfr))
                                        {
                                            // 均为空厂家时视为匹配
                                        }
                                        else if (!string.Equals(item.Manufacturer, cMfr, StringComparison.OrdinalIgnoreCase))
                                        {
                                            return false;
                                        }
                                    }

                                    // 4. 若汇总时勾选了【按备注合并】：要求备注一致
                                    if (mergeCond.ByRemark)
                                    {
                                        if (!string.Equals(item.Remark, cRemark, StringComparison.OrdinalIgnoreCase))
                                        {
                                            return false;
                                        }
                                    }

                                    // 5. 若汇总时勾选了【按原图型号合并】：要求 U 列原始型号匹配
                                    if (mergeCond.ByOriginalModel)
                                    {
                                        if (!string.IsNullOrEmpty(cRawU) && !string.IsNullOrEmpty(item.RawModelFromU))
                                        {
                                            if (!item.RawModelFromU.Contains(cRawU))
                                            {
                                                return false;
                                            }
                                        }
                                    }

                                    return true;
                                });
                                if (matchedItem != null)
                                {
                                    // 0. 若合并条件不包含名称（取消名称勾选），连同名称一起覆盖更新至 B 列 (索引 2)
                                    if (!mergeCond.ByName && !string.IsNullOrWhiteSpace(matchedItem.Name))
                                    {
                                        formulaMatrix[r, 2] = matchedItem.Name;
                                    }

                                    // 1. 若汇总表 D 列指定了新型号，则更新分类表 C 列 (索引 3)
                                    if (!string.IsNullOrWhiteSpace(matchedItem.NewModel))
                                    {
                                        formulaMatrix[r, 3] = matchedItem.NewModel;
                                    }

                                    // 2. 更新生产厂家 D 列 (索引 4)
                                    if (!string.IsNullOrWhiteSpace(matchedItem.Manufacturer))
                                    {
                                        formulaMatrix[r, 4] = matchedItem.Manufacturer;
                                    }

                                    // 3. 不管什么条件，更新单位 E 列 (索引 5)
                                    if (!string.IsNullOrWhiteSpace(matchedItem.Unit))
                                    {
                                        formulaMatrix[r, 5] = matchedItem.Unit;
                                    }

                                    // 4. 更新备注 I 列 (索引 9)
                                    formulaMatrix[r, 9] = matchedItem.Remark;

                                    // 5. 更新报出系数 / 加价系数 L 列 (索引 12)
                                    if (matchedItem.MarkupFactor > 0)
                                    {
                                        formulaMatrix[r, 12] = (double)matchedItem.MarkupFactor;
                                    }

                                    // 6. 更新表价 / 面价 M 列 (索引 13)
                                    formulaMatrix[r, 13] = matchedItem.MColContent;

                                    // 7. 更新折扣 N 列 (索引 14)
                                    formulaMatrix[r, 14] = matchedItem.NColContent;

                                    // 8. 更新额定电流: W 列 (Current, 索引 23)
                                    if (!string.IsNullOrWhiteSpace(matchedItem.Current) && formulaMatrix.GetLength(1) >= 23)
                                    {
                                        // 回写额定电流至明细表 W 列
                                        formulaMatrix[r, 23] = matchedItem.Current;
                                    }

                                    // 9. 更新极数: X 列 (Poles, 索引 24)
                                    if (!string.IsNullOrWhiteSpace(matchedItem.Poles) && formulaMatrix.GetLength(1) >= 24)
                                    {
                                        // 回写极数至明细表 X 列
                                        formulaMatrix[r, 24] = matchedItem.Poles;
                                    }

                                    // 10. 更新脱扣方式: Y 列 (trip, 索引 25)
                                    if (!string.IsNullOrWhiteSpace(matchedItem.Tripping) && formulaMatrix.GetLength(1) >= 25)
                                    {
                                        // 回写脱扣方式至明细表 Y 列
                                        formulaMatrix[r, 25] = matchedItem.Tripping;
                                    }

                                    // 11. 更新配套附件: Z 列 (Accessory, 索引 26)
                                    if (!string.IsNullOrWhiteSpace(matchedItem.Accessory) && formulaMatrix.GetLength(1) >= 26)
                                    {
                                        // 回写配套附件至明细表 Z 列
                                        formulaMatrix[r, 26] = matchedItem.Accessory;
                                    }

                                    // 12. 更新图块名称: AA 列 (BlockName / 扩展参数1, 索引 27)
                                    if (!string.IsNullOrWhiteSpace(matchedItem.Param1) && formulaMatrix.GetLength(1) >= 27)
                                    {
                                        // 回写图块名称至明细表 AA 列
                                        formulaMatrix[r, 27] = matchedItem.Param1;
                                    }

                                    // 13. 更新图块类别: AB 列 (BlockCategory / 扩展参数2, 索引 28)
                                    if (!string.IsNullOrWhiteSpace(matchedItem.Param2) && formulaMatrix.GetLength(1) >= 28)
                                    {
                                        // 回写图块类别至明细表 AB 列
                                        formulaMatrix[r, 28] = matchedItem.Param2;
                                    }

                                    // 标记当前箱柜已修改
                                    cabinetModified = true;
                                    updatedComponentCount++;
                                }
                            }

                            // 若箱柜有更新，一次性批量回写 2D 矩阵 (规则 7)
                            if (cabinetModified)
                            {
                                compRange.Formula = formulaMatrix;
                                updatedCabinetCount++;
                                sheetModified = true;
                            }
                        }

                        // 若分类表有修改则递增计数
                        if (sheetModified)
                        {
                            updatedSheetCount++;
                        }
                    }
                }
                finally
                {
                    // 恢复屏幕刷新、警告与系统事件
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }

                // 组装成功结果报文
                result.Success = true;
                result.UpdatedSheetCount = updatedSheetCount;
                result.UpdatedCabinetCount = updatedCabinetCount;
                result.UpdatedComponentCount = updatedComponentCount;
                result.Message = "一键更新成功：共同步 " + updatedSheetCount + " 个分类表、" + updatedCabinetCount + " 台箱柜、" + updatedComponentCount + " 项元器件明细！";

                // 记录成功日志
                LogHelper.WriteLog(result.Message);
                return result;
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"UpdateFromComponentSummarySheet 异常: {ex.Message}");
                result.Success = false;
                result.Message = $"一键更新遇到异常: {ex.Message}";
                return result;
            }
        }

        /// <summary>
        /// 切换指定列区域的隐藏与显示状态（如 F:P 或 Q:W）
        /// </summary>
        /// <param name="columnRange">列区域范围字符串，例如 "F:P" 或 "Q:W"</param>
        /// <param name="hidden">是否隐藏，true 为隐藏，false 为显示</param>
        /// <returns>操作是否成功</returns>
        public static bool SetSheetColumnsHidden(string columnRange, bool hidden)
        {
            try
            {
                // 验证列区域参数是否有效
                if (string.IsNullOrWhiteSpace(columnRange)) return false;

                // 获取当前活动 Excel 应用程序实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return false;

                // 获取当前活动工作表
                dynamic activeSheet = app.ActiveSheet;
                if (activeSheet == null) return false;

                // 优化 Excel 渲染性能，临时关闭屏幕刷新
                app.ScreenUpdating = false;

                try
                {
                    // 设置目标列区域的 Hidden 属性实现隐藏或展开
                    activeSheet.Columns[columnRange].Hidden = hidden;
                }
                finally
                {
                    // 恢复 Excel 屏幕刷新显示
                    app.ScreenUpdating = true;
                }

                // 返回操作成功标志
                return true;
            }
            catch (Exception ex)
            {
                // 记录异常日志并返回失败
                LogHelper.WriteLog($"SetSheetColumnsHidden 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取当前活动工作表中价格列 (G:O) 的隐藏状态
        /// </summary>
        /// <returns>价格列是否处于隐藏状态</returns>
        public static bool GetSheetColumnsHiddenStatus()
        {
            try
            {
                // 获取当前活动 Excel 应用程序实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return false;

                // 获取当前活动工作表
                dynamic activeSheet = app.ActiveSheet;
                if (activeSheet == null) return false;

                // 读取 G 列的 Hidden 状态作为价格列 (G:O) 的代表状态 (因新增原型号规格列整体向右顺延一列) --硬编码--
                bool isPriceHidden = Convert.ToBoolean(activeSheet.Columns["G"].Hidden);

                // 返回价格列隐藏状态
                return isPriceHidden;
            }
            catch
            {
                // 异常情况下默认返回未隐藏
                return false;
            }
        }
    }
}
