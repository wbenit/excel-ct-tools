using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelAddInDemo.Forms;
using ExcelAddInDemo.Models;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 公共服务分部类：专门处理分类表明细元器件提取与添加到本地个人库业务
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，严禁在窗口类中操作 Excel，所有硬编码显式标明
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 从当前活动工作表的选区中提取元器件，核对个人库排重状态，并调起核对确认弹窗
        /// </summary>
        public static void OpenAddToPersonalDbDialog()
        {
            try
            {
                // 获取 Excel 顶层应用实例
                Microsoft.Office.Interop.Excel.Application? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveWorkbook == null || app.ActiveSheet == null)
                {
                    MessageBox.Show("未检测到有效的 Excel 工作簿或工作表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); // --硬编码: 提示语--
                    return;
                }

                // 提取当前选中区域中合法的元器件候选列表 (内部已完整执行规则 6、7、8)
                var candidates = ExtractSelectedComponentsForPersonalDb(app);
                if (candidates == null || candidates.Count == 0)
                {
                    // 若无有效候选数据，直接退出 (提取方法内已给出针对性提示)
                    return;
                }

                // 调用本地个人库服务执行批量排重与价格变动核对
                Services.PersonalComponentDbService.CheckAndMatchCandidates(candidates);

                // 调起轻量现代化核对弹窗 (Vue 3 + Element Plus，主色调 #009688)
                AddToPersonalDbForm.ShowForm(candidates);
            }
            catch (Exception ex)
            {
                // 记录调度异常日志
                LogHelper.WriteLog($"[ExcelServices] OpenAddToPersonalDbDialog 异常: {ex.Message}");
                MessageBox.Show($"调起添加到个人库窗口失败: {ex.Message}", "系统错误", MessageBoxButtons.OK, MessageBoxIcon.Error); // --硬编码: 异常提示--
            }
        }

        /// <summary>
        /// 从当前分类表的选区中提取有效元器件列表 (严格遵守规则 6 箱柜明细结构、规则 7 二维数组一次性读取、规则 8 正确性预检)
        /// </summary>
        /// <param name="app">Excel Application 实例</param>
        /// <returns>待入库元器件候选列表</returns>
        public static List<AddToPersonalDbCandidateItem> ExtractSelectedComponentsForPersonalDb(Microsoft.Office.Interop.Excel.Application app)
        {
            var candidates = new List<AddToPersonalDbCandidateItem>();

            // 提取当前活动工作表
            Worksheet? ws = app.ActiveSheet as Worksheet;
            if (ws == null)
            {
                MessageBox.Show("未能识别当前活动工作表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); // --硬编码: 提示语--
                return candidates;
            }

            // 获取当前用户的活动选区
            Range? selection = app.Selection as Range;
            if (selection == null)
            {
                MessageBox.Show("请先在表格中选择需要入库的元器件所在行！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); // --硬编码: 提示语--
                return candidates;
            }

            try
            {
                // ------------------ 【规则 8：预检并自愈规则 6 箱柜定义名称】 ------------------
                // 在读取表格数据前，优先执行 FixAndFillCabinetNamesForSheet 确保 Cab_Det 与 Cab_Subsum 定义名称健全
                try
                {
                    Tool.FixAndFillCabinetNamesForSheet(ws);
                }
                catch { }

                // 获取当前工作表所属的工作簿对象
                Workbook? wb = ws.Parent as Workbook;
                string currentSheetName = Convert.ToString(ws.Name) ?? string.Empty;

                // ------------------ 【扫描当前表内的箱柜明细起止边界】 ------------------
                // 记录各个箱柜的明细区间：(int detRow, int compStart, int compEnd, int subsumRow)
                var cabinetRanges = new List<(int detRow, int compStart, int compEnd, int subsumRow)>();

                // 提取 det 前缀与 subsum 前缀
                string detPrefix = CabinetPrefixConfig.Current.DetPrefix;
                string subPrefix = CabinetPrefixConfig.Current.SubsumPrefix;

                // 遍历工作簿名称集合，收集属于当前 Sheet 的箱柜定义名称
                var detDict = new Dictionary<int, int>(); // cabinetIndex -> detRow
                var subDict = new Dictionary<int, int>(); // cabinetIndex -> subRow

                if (wb != null && wb.Names != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Name n in wb.Names)
                    {
                        try
                        {
                            string nName = Convert.ToString(n.Name) ?? string.Empty;
                            string cleanName = Tool.ExtractCleanNameStr(nName);
                            Range? r = null;
                            try { r = n.RefersToRange; } catch { }
                            if (r != null && r.Worksheet != null && string.Equals(Convert.ToString(r.Worksheet.Name), currentSheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                int k = Tool.ExtractIndexFromName(cleanName);
                                if (k > 0)
                                {
                                    if (cleanName.StartsWith(detPrefix, StringComparison.OrdinalIgnoreCase))
                                    {
                                        detDict[k] = r.Row;
                                    }
                                    else if (cleanName.StartsWith(subPrefix, StringComparison.OrdinalIgnoreCase))
                                    {
                                        subDict[k] = r.Row;
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }

                // 匹配成对的箱柜明细区间 (规则 6：元器件起始行 = Cab_Det + 2，元器件结束行 = Cab_Subsum - 1)
                foreach (var kvp in detDict)
                {
                    int k = kvp.Key;
                    int dRow = kvp.Value;
                    if (subDict.TryGetValue(k, out int sRow))
                    {
                        int cStart = dRow + 2;
                        int cEnd = sRow - 1;
                        if (cEnd >= cStart)
                        {
                            cabinetRanges.Add((dRow, cStart, cEnd, sRow));
                        }
                    }
                }

                // 若未识别到任何有效箱柜明细区间，说明当前可能不是分类明细表
                if (cabinetRanges.Count == 0)
                {
                    MessageBox.Show("当前工作表中未检测到合法的箱柜明细区域，请在分类表中操作！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); // --硬编码: 提示语--
                    return candidates;
                }

                // ------------------ 【计算选区行与元器件有效区间的交集】 ------------------
                // 收集选区覆盖的所有物理行号集合
                var selectedPhysicalRows = new HashSet<int>();
                foreach (Range area in selection.Areas)
                {
                    int areaStart = area.Row;
                    int areaEnd = areaStart + area.Rows.Count - 1;
                    for (int r = areaStart; r <= areaEnd; r++)
                    {
                        selectedPhysicalRows.Add(r);
                    }
                }

                // 过滤仅保留位于元器件插槽合法范围内的物理行
                var validCompRows = new List<int>();
                foreach (int r in selectedPhysicalRows)
                {
                    // 必须满足落在某个箱柜的 [compStart, compEnd] 闭区间内
                    if (cabinetRanges.Any(cab => r >= cab.compStart && r <= cab.compEnd))
                    {
                        validCompRows.Add(r);
                    }
                }

                // 排序物理行号
                validCompRows.Sort();

                if (validCompRows.Count == 0)
                {
                    MessageBox.Show("您所选择的行不在箱柜的元器件区域内！\n(请选择箱柜表头下方、小计行上方的有效元器件明细行)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); // --硬编码: 提示语--
                    return candidates;
                }

                // ------------------ 【规则 7：利用二维数组一次性批量读取内存数据】 ------------------
                int minRow = validCompRows.First();
                int maxRow = validCompRows.Last();

                // 一次性读取从 minRow 到 maxRow、第 1 列(A)至第 17 列(Q)的二维矩阵
                Range targetRange = ws.Range[ws.Cells[minRow, 1], ws.Cells[maxRow, 17]];
                object[,] matrix = (object[,])targetRange.Value2;

                int itemIndex = 1;
                foreach (int r in validCompRows)
                {
                    // 计算在二维矩阵中的相对行索引 (1-based)
                    int matrixRow = r - minRow + 1;

                    // B 列 (col 2): 元件名称
                    string name = Convert.ToString(matrix[matrixRow, 2])?.Trim() ?? string.Empty;
                    // C 列 (col 3): 规格型号
                    string model = Convert.ToString(matrix[matrixRow, 3])?.Trim() ?? string.Empty;
                    // D 列 (col 4): 生产厂家/品牌
                    string brand = Convert.ToString(matrix[matrixRow, 4])?.Trim() ?? string.Empty;
                    // I 列 (col 9): 备注
                    string remark = Convert.ToString(matrix[matrixRow, 9])?.Trim() ?? string.Empty;
                    // M 列 (col 13): 面价/单价
                    decimal price = 0.0m;
                    string priceRaw = Convert.ToString(matrix[matrixRow, 13])?.Trim() ?? string.Empty;
                    if (decimal.TryParse(priceRaw, out decimal pVal))
                    {
                        price = Math.Round(pVal, 2);
                    }
                    // Q 列 (col 17): 类别 (如“元件”)
                    string category = Convert.ToString(matrix[matrixRow, 17])?.Trim() ?? "元件"; // --硬编码: 默认类别--

                    // 过滤空行：如果名称和规格型号都为空，视为预留空白行，直接跳过
                    if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(model))
                    {
                        continue;
                    }

                    // 启发式参数推导：额定电流、极数、脱扣特性
                    int? current = ExtractCurrentFromText(model + " " + remark);
                    string poles = ExtractPolesFromText(model + " " + remark);
                    string tripping = ExtractTrippingFromText(model + " " + remark);

                    // 若型号非空但品牌为空，默认标记为“通用” --硬编码: 兜底品牌--
                    if (string.IsNullOrWhiteSpace(brand))
                    {
                        brand = "通用";
                    }

                    // 构建候选实体
                    var candidate = new AddToPersonalDbCandidateItem
                    {
                        Index = itemIndex++,
                        ExcelRow = r,
                        Brand = brand,
                        Name = name,
                        Model = model,
                        Price = price,
                        Remark = remark,
                        Category = category,
                        Current = current,
                        Poles = poles,
                        Tripping = tripping,
                        IsSelected = !string.IsNullOrWhiteSpace(model)
                    };

                    candidates.Add(candidate);
                }

                if (candidates.Count == 0)
                {
                    MessageBox.Show("选中的元器件区域中未发现任何有效物料行(名称或型号为空)！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); // --硬编码: 提示语--
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[ExcelServices] ExtractSelectedComponentsForPersonalDb 异常: {ex.Message}");
                MessageBox.Show($"提取选区元器件数据失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); // --硬编码: 异常提示--
            }

            return candidates;
        }

        /// <summary>
        /// 从型号或备注文本中启发式提取电流整型数字 (如 "32A", "100A", "/63" ➔ 32, 100, 63)
        /// </summary>
        private static int? ExtractCurrentFromText(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            // 优先匹配 100A / 32A 等带 A 单位的模式
            var m = Regex.Match(text, @"(\d+)\s*A\b", RegexOptions.IgnoreCase);
            if (m.Success && int.TryParse(m.Groups[1].Value, out int a)) return a;
            // 匹配斜杠后的电流数字 (如 NM1-125S/3300 100A 或 /100)
            m = Regex.Match(text, @"[/_\-](\d{2,4})\b");
            if (m.Success && int.TryParse(m.Groups[1].Value, out int b)) return b;
            return null;
        }

        /// <summary>
        /// 从型号或备注中提取极数 (如 "3P", "4P", "3极" ➔ "3", "4")
        /// </summary>
        private static string ExtractPolesFromText(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            var m = Regex.Match(text, @"([1-4])\s*(?:P|极)\b", RegexOptions.IgnoreCase);
            if (m.Success) return m.Groups[1].Value;
            return string.Empty;
        }

        /// <summary>
        /// 从型号或备注中提取脱扣特性 (如 "C", "D", "电子脱扣")
        /// </summary>
        private static string ExtractTrippingFromText(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            string clean = text!;
            if (Regex.IsMatch(clean, @"\bC\d+", RegexOptions.IgnoreCase) || clean.Contains("C型")) return "C";
            if (Regex.IsMatch(clean, @"\bD\d+", RegexOptions.IgnoreCase) || clean.Contains("D型")) return "D";
            if (clean.Contains("电子") || clean.Contains("MIC")) return "电子脱扣";
            return string.Empty;
        }
    }
}
