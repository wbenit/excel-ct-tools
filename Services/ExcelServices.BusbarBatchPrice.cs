using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：母排一键改价全项目扫描与回写引擎
    /// </summary>
    public static partial class ExcelServices
    {
        // 铜排理论金属密度 (单位: kg / (mm² * m))
        private const double CopperDensity = 0.0089;

        // 铝排理论金属密度 (单位: kg / (mm² * m))
        private const double AluminumDensity = 0.0027;

        /// <summary>
        /// 唤起基于 WebView2 + Vue 3 的“母排一键改价”弹窗工作台
        /// </summary>
        public static void ShowBusbarBatchPriceDialog()
        {
            try
            {
                // 获取当前活动 Excel Application 宿主对象
                dynamic app = ExcelDnaUtil.Application;
                // 获取活动工作表
                dynamic activeSheet = app.ActiveSheet;
                // 若活动表为空则阻断操作
                if (activeSheet == null)
                {
                    MessageBox.Show("未检测到当前活动的 Excel 工作表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 规则 8: 操作前对活动工作表执行规则 6 自愈检查
                Tool.FixAndFillCabinetNamesForSheet(activeSheet);

                // 获取 Excel 主窗口的 Win32 HWND 句柄，避免模态线程冲突
                IntPtr excelHwnd = ExcelDnaUtil.WindowHandle;
                // 实例化基于 WebView2 的无边框工作台窗体
                using var form = new Forms.BusbarBatchPriceForm();
                // 以模态对话框模式展示在 Excel 窗体前端
                form.ShowDialog(new ExcelWin32Window(excelHwnd));
            }
            catch (Exception ex)
            {
                // 捕获异常并弹出错误提示
                MessageBox.Show($"打开母排一键改价工作台失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 扫描指定分类工作表集合中的全部母排物料行
        /// </summary>
        public static BusbarBatchPriceScanResultDto ScanProjectBusbars(BusbarBatchPriceConfigDto config)
        {
            // 初始化返回结果对象
            var result = new BusbarBatchPriceScanResultDto
            {
                Config = config ?? new BusbarBatchPriceConfigDto()
            };

            try
            {
                // 获取 Excel 宿主对象与活动工作簿
                dynamic app = ExcelDnaUtil.Application;
                dynamic activeWb = app.ActiveWorkbook;
                dynamic activeSheet = app.ActiveSheet;
                // 校验工作簿有效性
                if (activeWb == null)
                {
                    result.Success = false;
                    result.Message = "未找到活动工作簿";
                    return result;
                }

                // 记录当前活动工作表名称
                result.CurrentSheetName = Convert.ToString(activeSheet?.Name)?.Trim() ?? string.Empty;

                // 收集全项目所有有效分类工作表
                var allCategorySheets = new List<string>();
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    string wsName = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                    // 过滤排除系统保留辅助表 --硬编码: 系统辅助表白名单过滤--
                    if (IsSystemAuxiliarySheet(wsName)) continue;
                    allCategorySheets.Add(wsName);
                }
                result.AvailableSheets = allCategorySheets;

                // 确定实际需要扫描的工作表清单
                var sheetsToScan = new List<string>();
                if (string.Equals(config?.ScopeMode, "all", StringComparison.OrdinalIgnoreCase))
                {
                    sheetsToScan.AddRange(allCategorySheets);
                }
                else
                {
                    // 仅扫描当前工作表或配置指定的表
                    if (config?.TargetSheets != null && config.TargetSheets.Count > 0)
                    {
                        sheetsToScan.AddRange(config.TargetSheets.Where(s => allCategorySheets.Contains(s)));
                    }
                    else if (!string.IsNullOrEmpty(result.CurrentSheetName) && allCategorySheets.Contains(result.CurrentSheetName))
                    {
                        sheetsToScan.Add(result.CurrentSheetName);
                    }
                    else if (allCategorySheets.Count > 0)
                    {
                        sheetsToScan.Add(allCategorySheets[0]);
                    }
                }

                // 涉及的有效箱柜集合
                var touchedCabinets = new HashSet<string>();

                // 逐个工作表扫描母排物料
                foreach (string sheetName in sheetsToScan)
                {
                    dynamic ws = null;
                    try { ws = activeWb.Worksheets[sheetName]; } catch { continue; }
                    if (ws == null) continue;

                    // 规则 8: 扫描前自愈箱柜定义名称
                    try { Tool.FixAndFillCabinetNamesForSheet(ws); } catch { }

                    // 提取该表的有效箱柜锚点字典 (规则 6)
                    var validCabinets = Tool.GetSheetValidCabinets(ws, activeWb);
                    if (validCabinets == null || validCabinets.Count == 0) continue;

                    // 提取当前工作表最大使用行数
                    int maxUsedRow = 500;
                    try { maxUsedRow = ws.UsedRange.Rows.Count + ws.UsedRange.Row; } catch { maxUsedRow = 500; }
                    if (maxUsedRow < 20) maxUsedRow = 500;

                    // 规则 7: 单次 COM 调用抓取 A 到 N 列数据矩阵至内存 (N 为第 14 列，M 为第 13 列表价值)
                    dynamic fullRange = ws.Range[$"A1:N{maxUsedRow}"];
                    object[,] matrix = ConvertTo2DArray(fullRange.Value2, maxUsedRow, 14);

                    // 遍历各箱柜扫描元器件明细区
                    foreach (var kvp in validCabinets)
                    {
                        int cabIndex = kvp.Key;
                        var cabInfo = kvp.Value;
                        if (cabInfo == null) continue;

                        // 提取箱柜明细行与小计行物理行号 (安全兼顾动态属性与内存行号缓存)
                        int detRow = cabInfo.Det != null ? Convert.ToInt32(cabInfo.Det.Row) : cabInfo.DetRow;
                        int subsumRow = cabInfo.Subsum != null ? Convert.ToInt32(cabInfo.Subsum.Row) : cabInfo.SubsumRow;
                        if (detRow <= 0 || subsumRow <= 0) continue;

                        // 元器件起始行为 detRow + 2，终止行为 subsumRow - 1
                        int compStartRow = detRow + 2;
                        int compEndRow = subsumRow - 1;

                        // 提取箱柜名称 (位于 detRow 的 B 列)
                        string cabName = matrix[detRow, 2]?.ToString()?.Trim() ?? $"箱柜 {cabIndex}";
                        string cabUniqueKey = $"{sheetName}_{cabIndex}";

                        // 循环元器件区间各物理行
                        for (int r = compStartRow; r <= compEndRow && r <= maxUsedRow; r++)
                        {
                            // 提取 B 列名称、C 列型号、E 列单位、F 列数量、M 列基准单价
                            string name = matrix[r, 2]?.ToString()?.Trim() ?? string.Empty;
                            string model = matrix[r, 3]?.ToString()?.Trim() ?? string.Empty;
                            string unit = matrix[r, 5]?.ToString()?.Trim() ?? "kg";
                            string remark = matrix[r, 9]?.ToString()?.Trim() ?? string.Empty;

                            // 校验是否属于母排材料行
                            if (!IsBusbarMaterialRow(name, model)) continue;

                            // 解析用量数量
                            double qty = ParseDoubleSafe(matrix[r, 6]);
                            // 解析 M 列原基准表价单价
                            double oldUnitPrice = ParseDoubleSafe(matrix[r, 13]);
                            // 原合价
                            double oldTotalPrice = Math.Round(oldUnitPrice * qty, 2);

                            // 判定母排材质归属 (铜排 / 铝排)
                            string matType = DetermineMaterialType(name, model);

                            // 判断是否为按米计价模式
                            bool isMeter = IsMeterUnit(unit);
                            // 理论每米单重计算
                            double meterWeight = CalculateBusbarTheoreticalMeterWeight(model, matType);

                            // 依据配置测算新单价
                            double newUnitPrice = 0.0;
                            double baseMetalPrice = string.Equals(matType, "aluminum", StringComparison.OrdinalIgnoreCase)
                                ? config.AlTotalUnitPrice
                                : config.CuTotalUnitPrice;

                            // 若为米制单位且配置了自动米制换算
                            if (isMeter && config.AutoConvertMeter && meterWeight > 0)
                            {
                                // 新米价 = 每公斤综合金属单价 * 理论米重
                                newUnitPrice = Math.Round(baseMetalPrice * meterWeight, 2);
                            }
                            else
                            {
                                // 默认公斤单价或未识别米重直接采用综合金属单价
                                newUnitPrice = baseMetalPrice;
                            }

                            // 测算新合价与差额
                            double newTotalPrice = Math.Round(newUnitPrice * qty, 2);
                            double diffAmount = Math.Round(newTotalPrice - oldTotalPrice, 2);

                            // 创建行条目模型
                            var rowItem = new BusbarMaterialRowDto
                            {
                                SheetName = sheetName,
                                RowIndex = r,
                                CabinetIndex = cabIndex,
                                CabinetName = cabName,
                                Name = string.IsNullOrEmpty(name) ? (matType == "aluminum" ? "矩形铝母线" : "矩形铜母线") : name,
                                Model = model,
                                MaterialType = matType,
                                Unit = string.IsNullOrEmpty(unit) ? "kg" : unit,
                                Qty = qty,
                                OldUnitPrice = oldUnitPrice,
                                OldTotalPrice = oldTotalPrice,
                                NewUnitPrice = newUnitPrice,
                                NewTotalPrice = newTotalPrice,
                                DiffAmount = diffAmount,
                                IsMeterMode = isMeter,
                                TheoreticalKgPerMeter = meterWeight,
                                Remark = remark,
                                Selected = true
                            };

                            // 加入明细集合并记录统计
                            result.Rows.Add(rowItem);
                            touchedCabinets.Add(cabUniqueKey);

                            // 累计总重量 (若单位为米则通过米重折算)
                            double rowWeight = isMeter ? (qty * meterWeight) : qty;
                            if (matType == "aluminum")
                            {
                                result.TotalAluminumWeight += rowWeight;
                            }
                            else
                            {
                                result.TotalCopperWeight += rowWeight;
                            }

                            // 累计金额
                            result.OldTotalAmount += oldTotalPrice;
                            result.NewTotalAmount += newTotalPrice;
                        }
                    }
                }

                // 统计汇总
                result.TotalCabinetCount = touchedCabinets.Count;
                result.TotalRowCount = result.Rows.Count;
                result.TotalDiffAmount = Math.Round(result.NewTotalAmount - result.OldTotalAmount, 2);
                result.TotalCopperWeight = Math.Round(result.TotalCopperWeight, 2);
                result.TotalAluminumWeight = Math.Round(result.TotalAluminumWeight, 2);
                result.OldTotalAmount = Math.Round(result.OldTotalAmount, 2);
                result.NewTotalAmount = Math.Round(result.NewTotalAmount, 2);

                result.Message = $"扫描完成，共在 {sheetsToScan.Count} 个工作表中找到 {result.TotalRowCount} 处母排物料。";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"扫描母排物料异常: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// 一键将调整后的母排单价批量回写至各 Excel 分类工作表的 M 列
        /// </summary>
        public static (bool Success, string Message, int UpdatedCount) ApplyBusbarBatchPriceToExcel(BusbarApplyPricePayload payload)
        {
            // 参数合法性前置检查
            if (payload == null || payload.Items == null || payload.Items.Count == 0)
            {
                return (false, "未提供需回写的母排条目", 0);
            }

            // 获取用户勾选的条目列表
            var selectedItems = payload.Items.Where(it => it.Selected).ToList();
            if (selectedItems.Count == 0)
            {
                return (false, "请先勾选需要改价的母排条目", 0);
            }

            try
            {
                // 获取 Excel 宿主对象
                dynamic app = ExcelDnaUtil.Application;
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null) return (false, "未找到活动工作簿", 0);

                // 暂停屏幕刷新与弹窗警报，加速 COM 批量写入并防闪烁
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                int updatedCount = 0;

                // 按工作表对更新项进行分组处理
                var groupedBySheet = selectedItems.GroupBy(it => it.SheetName);

                foreach (var group in groupedBySheet)
                {
                    string sheetName = group.Key;
                    dynamic ws = null;
                    try { ws = activeWb.Worksheets[sheetName]; } catch { continue; }
                    if (ws == null) continue;

                    // 循环写入当前工作表各行的 M 列 (第 13 列: 表价/基准单价)
                    foreach (var item in group)
                    {
                        if (item.RowIndex > 0)
                        {
                            // 直接更新 M 列数值
                            ws.Cells[item.RowIndex, 13].Value = item.NewUnitPrice;
                            updatedCount++;
                        }
                    }
                }

                // 强制触发 Excel 全局公式链条重算 (驱动 G/H/J/K 列与箱柜小计自适应更新)
                app.Calculate();

                // 恢复屏幕刷新
                app.ScreenUpdating = true;
                app.DisplayAlerts = true;

                return (true, $"成功更新 {groupedBySheet.Count()} 个工作表共 {updatedCount} 处母排单价！", updatedCount);
            }
            catch (Exception ex)
            {
                // 恢复屏幕刷新避免 Excel 假死
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                }
                catch { }

                return (false, $"批量回写 Excel 失败: {ex.Message}", 0);
            }
        }

        /// <summary>
        /// 判定当前物料行是否属于母排材料行
        /// </summary>
        private static bool IsBusbarMaterialRow(string name, string model)
        {
            // 排除空项
            if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(model)) return false;

            // 严格排除非母排的附件与加工辅料 --硬编码: 母排辅材排除过滤词--
            string[] excludeKeywords = new[] { "夹", "套管", "热缩", "端子", "铜鼻子", "绝缘子", "支架", "连接排配件", "保护罩", "螺栓", "母线框" };
            foreach (var kw in excludeKeywords)
            {
                if (name.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0) return false;
            }

            // 包含母排关键名称识别 --硬编码: 母排核心关键字--
            string[] busbarNameKeywords = new[] { "铜排", "母排", "铜母线", "母线", "矩形母线", "铝排", "铝母线", "TMY", "LMY", "接地铜排", "零地排", "PE排", "引下排", "分支排", "水平主排", "主母排", "零排", "地排" };
            foreach (var kw in busbarNameKeywords)
            {
                if (name.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }

            // 从型号匹配母排特征 (以 TMY / LMY 开头，或包含 TMY/LMY 标识)
            if (model.IndexOf("TMY", StringComparison.OrdinalIgnoreCase) >= 0 ||
                model.IndexOf("LMY", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 识别母排材质归属 ("aluminum" 或 "copper")
        /// </summary>
        private static string DetermineMaterialType(string name, string model)
        {
            // 优先判定铝材质
            if (name.IndexOf("铝", StringComparison.OrdinalIgnoreCase) >= 0 ||
                model.IndexOf("LMY", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "aluminum";
            }
            // 默认为铜排
            return "copper";
        }

        /// <summary>
        /// 判定计量单位是否为米制
        /// </summary>
        private static bool IsMeterUnit(string unit)
        {
            if (string.IsNullOrWhiteSpace(unit)) return false;
            string u = unit.Trim().ToLowerInvariant();
            return u == "米" || u == "m" || u == "meter";
        }

        /// <summary>
        /// 解析母排型号规格并计算理论每米重量 (kg/m)
        /// 支持格式示例: TMY-30*4, TMY-2*(100*10), TMY-3*(80*8)+1*(80*8), TMY 60x6
        /// </summary>
        public static double CalculateBusbarTheoreticalMeterWeight(string model, string materialType)
        {
            if (string.IsNullOrWhiteSpace(model)) return 0.0;

            // 依据材质选择金属密度
            double density = string.Equals(materialType, "aluminum", StringComparison.OrdinalIgnoreCase)
                ? AluminumDensity
                : CopperDensity;

            // 统一乘号与格式
            string cleaned = model.ToUpperInvariant()
                .Replace("X", "*")
                .Replace("×", "*")
                .Replace(" ", "");

            double totalSectionMm2 = 0.0;

            try
            {
                // 正则模式 1: 匹配组合拼数型，如 3*(80*8) 或 1*(60*6) 或 2*(100*10)
                var multiMatch = Regex.Matches(cleaned, @"(\d+)\*\((\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)\)");
                if (multiMatch.Count > 0)
                {
                    foreach (Match m in multiMatch)
                    {
                        double count = double.Parse(m.Groups[1].Value);
                        double width = double.Parse(m.Groups[2].Value);
                        double thick = double.Parse(m.Groups[3].Value);
                        totalSectionMm2 += count * width * thick;
                    }
                }
                else
                {
                    // 正则模式 2: 匹配简单乘积型，如 80*8 或 2*80*8 或 60*6
                    var simpleMatch = Regex.Match(cleaned, @"(?:TMY|LMY)?-?(?:(\d+)\*)?(\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)");
                    if (simpleMatch.Success)
                    {
                        double count = 1.0;
                        if (!string.IsNullOrEmpty(simpleMatch.Groups[1].Value))
                        {
                            count = double.Parse(simpleMatch.Groups[1].Value);
                        }
                        double width = double.Parse(simpleMatch.Groups[2].Value);
                        double thick = double.Parse(simpleMatch.Groups[3].Value);
                        totalSectionMm2 = count * width * thick;
                    }
                }

                // 若计算出截面积，则计算理论米重: 截面积 (mm²) * 密度 (kg/(mm²·m))
                if (totalSectionMm2 > 0)
                {
                    return Math.Round(totalSectionMm2 * density, 3);
                }
            }
            catch
            {
                // 解析失败兜底返回 0
            }

            return 0.0;
        }

        /// <summary>
        /// 判定工作表是否为系统内部保留的辅助表
        /// </summary>
        private static bool IsSystemAuxiliarySheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) return true;
            string s = sheetName.Trim();
            // 排除核心系统表 --硬编码--
            return string.Equals(s, "项目信息", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(s, "材料分布表", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(s, "元件汇总分布表", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(s, "元件汇总表", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(s, "元件汇总调价清单", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(s, "屏柜汇总表", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(s, "屏柜分项表", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(s, "元器件数据管理", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(s, "汇总调价表", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 安全解析浮点数值
        /// </summary>
        private static double ParseDoubleSafe(object val)
        {
            if (val == null || val == DBNull.Value) return 0.0;
            if (double.TryParse(val.ToString(), out double d)) return d;
            return 0.0;
        }
    }
}

