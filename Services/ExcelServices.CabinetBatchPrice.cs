using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：箱体一键改价全项目扫描、多模式测算与回写引擎
    /// 遵循规范：新增代码每 3 行至少包含 1 行中文注释
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 唤起基于 WebView2 + Vue 3 的“箱体一键改价”弹窗工作台
        /// </summary>
        public static void ShowCabinetBatchPriceDialog()
        {
            try
            {
                // 获取当前活动 Excel Application 宿主对象
                dynamic app = ExcelDnaUtil.Application;
                dynamic activeSheet = app?.ActiveSheet;
                // 若活动表为空则阻断操作
                if (activeSheet == null)
                {
                    MessageBox.Show("未检测到当前活动的 Excel 工作表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 规则 8: 操作前对活动工作表执行箱柜锚点自愈检查
                try { Tool.FixAndFillCabinetNamesForSheet(activeSheet); } catch { }

                // 获取 Excel 主窗口 Win32 HWND 句柄，避免模态线程死锁
                IntPtr excelHwnd = ExcelDnaUtil.WindowHandle;
                // 实例化基于 WebView2 的无边框箱体改价工作台窗体
                using var form = new Forms.CabinetBatchPriceForm();
                // 以模态对话框模式展示在 Excel 窗体前端
                form.ShowDialog(new ExcelWin32Window(excelHwnd));
            }
            catch (Exception ex)
            {
                // 捕获异常并弹出错误提示
                MessageBox.Show($"打开箱体一键改价工作台失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogHelper.WriteLog($"[CabinetBatchPrice] ShowCabinetBatchPriceDialog 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 扫描指定分类工作表集合中的全部箱体/壳体物料行
        /// </summary>
        /// <param name="config">前端传入的箱体改价规则配置</param>
        /// <returns>扫描结果传输模型</returns>
        public static CabinetBatchPriceScanResultDto ScanProjectCabinets(CabinetBatchPriceConfigDto config)
        {
            // 初始化返回结果传输实体
            var result = new CabinetBatchPriceScanResultDto();
            config ??= new CabinetBatchPriceConfigDto();

            // 提取/加载系统全局定额规则，保障与【辅材壳体计算】完全同源一致
            QuotationRules quotationRules = null;
            try { quotationRules = LoadQuotationRules(); } catch { }
            quotationRules ??= new QuotationRules();

            // 提取壳体名称匹配标记，默认 "箱体" --硬编码: 默认壳体匹配名--
            string shellMatchName = quotationRules.ShellRules?.ShellMatchName?.Trim() ?? "箱体";
            if (string.IsNullOrWhiteSpace(shellMatchName)) shellMatchName = "箱体";

            // 提取壳体价格定额规则配置 (材质单价库、高度阶梯、面数)
            ShellPriceConfig shellPriceRules = config.ShellPriceRules;
            if (shellPriceRules == null || shellPriceRules.MaterialPrices == null || shellPriceRules.MaterialPrices.Count == 0)
            {
                shellPriceRules = quotationRules.ShellPriceRules ?? new ShellPriceConfig();
            }
            config.ShellPriceRules = shellPriceRules;
            result.ShellPriceRules = shellPriceRules;
            result.DefaultMaterial = !string.IsNullOrWhiteSpace(shellPriceRules.DefaultMaterial) ? shellPriceRules.DefaultMaterial : "镀锌板";

            try
            {
                // 获取 Excel 宿主对象与活动工作簿 (支持 ExcelDna 内部与外部自动化安全兜底)
                dynamic app = null;
                try { app = ExcelDnaUtil.Application; } catch { }
                if (app == null)
                {
                    try { app = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"); } catch { }
                }

                // 获取活动工作簿，若 ActiveWorkbook 暂时失去焦点则安全尝试 Workbooks[1]
                dynamic activeWb = null;
                try { activeWb = app?.ActiveWorkbook; } catch { }
                if (activeWb == null)
                {
                    try { if (app != null && app.Workbooks != null && app.Workbooks.Count > 0) activeWb = app.Workbooks[1]; } catch { }
                }
                dynamic activeSheet = null;
                try { activeSheet = app?.ActiveSheet ?? activeWb?.ActiveSheet; } catch { }

                if (activeWb == null)
                {
                    result.Success = false;
                    result.Message = "未检测到已打开的活动工作簿！";
                    return result;
                }

                // 记录当前活动分类工作表名称
                result.CurrentSheetName = Convert.ToString(activeSheet?.Name)?.Trim() ?? string.Empty;

                // 收集全项目所有非系统辅助分类工作表
                var allCategorySheets = new List<string>();
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    string wsName = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                    // 过滤排除系统保留表 (如 项目信息, 汇总表, 字典表等)
                    if (IsSystemAuxiliarySheet(wsName)) continue;
                    allCategorySheets.Add(wsName);
                }
                result.AvailableSheets = allCategorySheets;

                // 确定待扫描的工作表清单
                var sheetsToScan = new List<string>();
                if (string.Equals(config.ScopeMode, "all", StringComparison.OrdinalIgnoreCase))
                {
                    sheetsToScan.AddRange(allCategorySheets);
                }
                else
                {
                    // 仅扫描当前工作表或配置指定的表
                    if (config.TargetSheets != null && config.TargetSheets.Count > 0)
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

                // 记录纳管的箱柜总数 (去重统计)
                var touchedCabinets = new HashSet<string>();

                // 逐个工作表极速扫描箱体物料
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

                    // 抓取 A 到 Q 列数据矩阵至内存 (17 列，N 为销售系数，M 为基准表价，Q 为成本类别)
                    dynamic fullRange = ws.Range[$"A1:Q{maxUsedRow}"];
                    object[,] matrix = ConvertTo2DArray(fullRange.Value2, maxUsedRow, 17);

                    // 遍历各箱柜扫描箱体物料行
                    foreach (var kvp in validCabinets)
                    {
                        int cabIndex = kvp.Key;
                        var cabInfo = kvp.Value;
                        if (cabInfo == null) continue;

                        // 提取箱柜明细行、小计行与总计行物理行号
                        int detRow = cabInfo.Det != null ? Convert.ToInt32(cabInfo.Det.Row) : cabInfo.DetRow;
                        int subsumRow = cabInfo.Subsum != null ? Convert.ToInt32(cabInfo.Subsum.Row) : cabInfo.SubsumRow;
                        int tolsumRow = cabInfo.Tolsum != null ? Convert.ToInt32(cabInfo.Tolsum.Row) : cabInfo.TolsumRow;
                        if (detRow <= 0 || subsumRow <= 0 || subsumRow <= detRow) continue;

                        // 若总计行锚点未定义或小于小计行，自动延伸至小计后 15 行或工作表使用末尾
                        if (tolsumRow <= subsumRow)
                        {
                            tolsumRow = Math.Min(subsumRow + 15, maxUsedRow); // --硬编码: 计费区兜底探测深度--
                        }

                        // 记录箱柜唯一标识
                        string cabUniqueKey = $"{sheetName}_{cabIndex}";
                        touchedCabinets.Add(cabUniqueKey);

                        // 提取箱柜名称 (位于 detRow 的 B 列)
                        string cabName = matrix[detRow, 2]?.ToString()?.Trim() ?? $"箱柜 {cabIndex}";
                        // 提取箱柜主行规格型号 (位于 detRow 的 C 列，清洗前缀标签)
                        string cabMainModel = CleanHeaderModelText(matrix[detRow, 3]?.ToString()?.Trim());
                        // 提取箱柜台数 (位于 detRow 的 F 列)
                        double cabQty = ParseDoubleSafe(matrix[detRow, 6]);
                        if (cabQty <= 0) cabQty = 1.0;

                        // 标记当前箱柜是否已在计费区命中箱体
                        bool foundCabinetInBilling = false;

                        // 第一阶段：优先扫描【计费区域】(subsumRow 至 tolsumRow)
                        for (int r = subsumRow; r <= tolsumRow && r <= maxUsedRow; r++)
                        {
                            // 提取 B 列名称、C 列型号、E 列单位、F 列数量、M 列基准单价、Q 列成本分摊项
                            string name = matrix[r, 2]?.ToString()?.Trim() ?? string.Empty;
                            string model = matrix[r, 3]?.ToString()?.Trim() ?? string.Empty;
                            string unit = matrix[r, 5]?.ToString()?.Trim() ?? "台";
                            string qCategory = matrix[r, 17]?.ToString()?.Trim() ?? string.Empty;

                            // 选项 A 核心守门：严格全字比对壳体匹配名，排除小计、人工劳务及辅材
                            if (!IsCabinetMaterialRow(name, model, qCategory, shellMatchName)) continue;

                            // 解析数量 (计费区 F 列若为空或0，智能继承箱柜主台数)
                            double qty = ParseDoubleSafe(matrix[r, 6]);
                            if (qty <= 0) qty = cabQty;

                            // 解析 M 列原基准表价值
                            double oldUnitPrice = ParseDoubleSafe(matrix[r, 13]);
                            // 若 M 列表价为空，尝试从 G 列销售单价或 J 列成本单价兜底
                            if (oldUnitPrice <= 0)
                            {
                                oldUnitPrice = ParseDoubleSafe(matrix[r, 7]);
                                if (oldUnitPrice <= 0) oldUnitPrice = ParseDoubleSafe(matrix[r, 10]);
                            }
                            double oldTotalPrice = Math.Round(oldUnitPrice * qty, 2);

                            // 构建单个箱体明细模型
                            var item = new CabinetBatchItemDto
                            {
                                Id = $"{sheetName}_R{r}",
                                SheetName = sheetName,
                                CabinetIndex = cabIndex,
                                CabinetName = cabName,
                                RowIndex = r,
                                Name = string.IsNullOrWhiteSpace(name) ? "箱体" : name,
                                Model = model,
                                Unit = string.IsNullOrWhiteSpace(unit) ? "台" : unit,
                                Qty = Math.Round(qty, 2),
                                OldUnitPrice = Math.Round(oldUnitPrice, 2),
                                OldTotalPrice = oldTotalPrice,
                                Selected = true
                            };

                            // 识别箱柜是否含有“二层板”结构特征
                            bool hasSecondPlate = (!string.IsNullOrWhiteSpace(cabName) && cabName.Contains("二层板")) ||
                                                  (!string.IsNullOrWhiteSpace(model) && model.Contains("二层板"));
                            item.HasSecondPlate = hasSecondPlate;

                            // 解析箱体长宽高尺寸或 PZ30 模数 (支持继承箱柜主型号与箱柜名，并联动图一高度深度梯度表)
                            ParseCabinetDimensions(item, cabMainModel, cabName, quotationRules.ShellRules);

                            // 自动推导材质 (复用辅材壳体计算的 DetermineMaterial 机制)
                            item.Material = DetermineMaterial(cabName, config.ShellPriceRules);

                            // 自动根据高度推导板厚 (高度阶梯决策法: <=800:1.2mm, <=1600:1.5mm, >1600:2.0mm)
                            if (item.HasValidSize && item.Height > 0)
                            {
                                item.Thickness = DetermineThickness((int)item.Height, config.ShellPriceRules);
                            }
                            else
                            {
                                item.Thickness = 1.5; // 默认 1.5mm 兜底 --硬编码--
                            }

                            // 根据当前配置计算新单价与差额
                            CalculateSingleCabinetPrice(item, config);

                            // 收集至扫描结果集并标记已命中
                            result.Rows.Add(item);
                            foundCabinetInBilling = true;
                        }

                        // 第二阶段：若计费区未命中，兼顾扫描【元器件明细区】(detRow + 1 至 subsumRow - 1)
                        if (!foundCabinetInBilling)
                        {
                            int compStartRow = detRow + 1;
                            int compEndRow = subsumRow - 1;

                            for (int r = compStartRow; r <= compEndRow && r <= maxUsedRow; r++)
                            {
                                string name = matrix[r, 2]?.ToString()?.Trim() ?? string.Empty;
                                string model = matrix[r, 3]?.ToString()?.Trim() ?? string.Empty;
                                string unit = matrix[r, 5]?.ToString()?.Trim() ?? "台";
                                string qCategory = matrix[r, 17]?.ToString()?.Trim() ?? string.Empty;

                                // 同样执行选项 A 全字匹配
                                if (!IsCabinetMaterialRow(name, model, qCategory, shellMatchName)) continue;

                                double qty = ParseDoubleSafe(matrix[r, 6]);
                                if (qty <= 0) qty = 1.0;

                                double oldUnitPrice = ParseDoubleSafe(matrix[r, 13]);
                                if (oldUnitPrice <= 0) oldUnitPrice = ParseDoubleSafe(matrix[r, 7]);
                                double oldTotalPrice = Math.Round(oldUnitPrice * qty, 2);

                                var item = new CabinetBatchItemDto
                                {
                                    Id = $"{sheetName}_R{r}",
                                    SheetName = sheetName,
                                    CabinetIndex = cabIndex,
                                    CabinetName = cabName,
                                    RowIndex = r,
                                    Name = string.IsNullOrWhiteSpace(name) ? "箱体" : name,
                                    Model = model,
                                    Unit = string.IsNullOrWhiteSpace(unit) ? "台" : unit,
                                    Qty = Math.Round(qty, 2),
                                    OldUnitPrice = Math.Round(oldUnitPrice, 2),
                                    OldTotalPrice = oldTotalPrice,
                                    Selected = true
                                };

                                bool hasSecondPlate = (!string.IsNullOrWhiteSpace(cabName) && cabName.Contains("二层板")) ||
                                                      (!string.IsNullOrWhiteSpace(model) && model.Contains("二层板"));
                                item.HasSecondPlate = hasSecondPlate;

                                // 解析箱体长宽高尺寸或模数并联动图一高度深度梯度表
                                ParseCabinetDimensions(item, cabMainModel, cabName, quotationRules.ShellRules);
                                item.Material = DetermineMaterial(cabName, config.ShellPriceRules);
                                item.Thickness = (item.HasValidSize && item.Height > 0)
                                    ? DetermineThickness((int)item.Height, config.ShellPriceRules)
                                    : 1.5;

                                CalculateSingleCabinetPrice(item, config);
                                result.Rows.Add(item);
                            }
                        }
                    }
                }

                // 汇总反馈
                result.Success = true;
                result.TotalCabinetCount = touchedCabinets.Count;
                result.Message = $"已成功检测到 {result.Rows.Count} 处箱体/壳体物料（涉及 {touchedCabinets.Count} 台箱柜）";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"扫描箱体异常: {ex.Message}";
                LogHelper.WriteLog($"[CabinetBatchPrice] ScanProjectCabinets 异常: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 清洗箱柜主信息行 detRow 中的型号文本 (去除 "型号:" 等前缀)
        /// </summary>
        private static string CleanHeaderModelText(string rawModel)
        {
            if (string.IsNullOrWhiteSpace(rawModel)) return string.Empty;
            // 正则清除开头的 "型号:"、"规格:"、"型号：" 等
            string cleaned = Regex.Replace(rawModel.Trim(), @"^(?:型号|规格|型式)[:：\s]*", "", RegexOptions.IgnoreCase).Trim();
            return cleaned;
        }

        /// <summary>
        /// 判定指定物料行是否为箱体/壳体材料行 (选项 A 规则：严格全字比对 ShellMatchName 与备选词，杜绝模糊包含误判)
        /// </summary>
        private static bool IsCabinetMaterialRow(string bName, string cModel, string qCategory, string shellMatchName = "箱体")
        {
            // 排除小计、总计、成套费、税金、辅材等综合项
            if (bName.Contains("小计") || bName.Contains("合计") || bName.Contains("总计") ||
                bName.Contains("成套费") || bName.Contains("税金") || bName.Contains("辅材") ||
                bName.Contains("管理费") || bName.Contains("利润") || bName.Contains("运费"))
            {
                return false;
            }

            // 排除人工费、制作工价等劳务行
            if (bName.Contains("制作") || bName.Contains("人工") || bName.Contains("工时") || bName.Contains("组装"))
            {
                return false;
            }

            // 排除支架、母线框、绝缘子等五金配件
            if (bName.Contains("支架") || bName.Contains("螺栓") || bName.Contains("导轨") || bName.Contains("母线框") || bName.Contains("铜排"))
            {
                return false;
            }

            // 1. 选项 A (核心规范)：全字严格比对用户配置的壳体名称，备选词 "箱体"、"壳体"、"柜体"
            // 绝不再做包含“配电箱/动力箱”的模糊包含匹配，彻底消灭“配电箱安装”、“动力箱制作劳务”等误判
            if (string.Equals(bName, shellMatchName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(bName, "箱体", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(bName, "壳体", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(bName, "柜体", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // 2. 若 Q 列显式标注为“箱体”或“壳体”或“柜体” (全字严格比对)
            if (string.Equals(qCategory, "箱体", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(qCategory, "壳体", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(qCategory, "柜体", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // 3. 若 C 列以常见标准箱体型号开头 (如 JXF, PZ30, GGD, GCK, KYN28)
            if (!string.IsNullOrWhiteSpace(cModel) && Regex.IsMatch(cModel, @"^(JXF|PZ30|GGD|GCK|GCS|MNS|KYN28|XL-21)", RegexOptions.IgnoreCase))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 解析箱体尺寸（多级回退匹配：自身型号 -> 箱柜主型号 -> 箱柜名称，并联动图一深度推荐规则）
        /// </summary>
        private static void ParseCabinetDimensions(CabinetBatchItemDto item, string fallbackMainModel = "", string fallbackCabName = "", ShellConfig shellConfig = null)
        {
            // 构建候选文本检查列表 (按优先级降序排序)
            var candidates = new List<string>();
            if (!string.IsNullOrWhiteSpace(item.Model) && item.Model != "型号:")
            {
                candidates.Add(item.Model.Trim());
            }
            if (!string.IsNullOrWhiteSpace(fallbackMainModel))
            {
                candidates.Add(fallbackMainModel.Trim());
            }
            if (!string.IsNullOrWhiteSpace(fallbackCabName))
            {
                candidates.Add(fallbackCabName.Trim());
            }

            // 逐个候选文本尝试提取尺寸
            foreach (var text in candidates)
            {
                // 尝试提取尺寸并传入壳体深度选型规则
                if (TryExtractDimensionsFromText(text, item, shellConfig))
                {
                    // 若当前明细的型号为空或仅为标签，自动回填识别到尺寸的有效型号供用户查看
                    if (string.IsNullOrWhiteSpace(item.Model) || item.Model == "型号:")
                    {
                        item.Model = text;
                    }
                    return;
                }
            }

            // 所有候选均未能解析出有效尺寸
            item.HasValidSize = false;
        }

        /// <summary>
        /// 从指定文本中尝试提取箱体长宽高尺寸或模数位数 (二维尺寸缺失深度时联动图一深度梯度表)
        /// </summary>
        private static bool TryExtractDimensionsFromText(string text, CabinetBatchItemDto item, ShellConfig shellConfig = null)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;

            // 1. 匹配 PZ30 模数箱特征 (如 18位, 24位)
            var pzMatch = Regex.Match(text, @"(\d{1,2})\s*位", RegexOptions.IgnoreCase);
            if (pzMatch.Success && int.TryParse(pzMatch.Groups[1].Value, out int digits))
            {
                item.IsModular = true;
                item.ModularDigits = digits;
                item.HasValidSize = true;
                return true;
            }

            // 2. 匹配三维长宽高尺寸或二维长宽尺寸 (如 800*1200*300 或 800*1800 或 W800*H1200*D300)
            var sizeMatch = Regex.Match(text, @"(?:W|宽)?\s*(\d{2,4})\s*[\*xX×\*]\s*(?:H|高)?\s*(\d{2,4})(?:\s*[\*xX×\*]\s*(?:D|厚|深)?\s*(\d{2,4}))?", RegexOptions.IgnoreCase);
            if (sizeMatch.Success)
            {
                double.TryParse(sizeMatch.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double w);
                double.TryParse(sizeMatch.Groups[2].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double h);
                double d = 0;

                // 若三维中显式注明深度且大于等于 50mm
                if (sizeMatch.Groups[3].Success && double.TryParse(sizeMatch.Groups[3].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedD))
                {
                    if (parsedD >= 50) d = parsedD;
                }

                // 核心修复：若型号未注明深度 (如 800*1800)，严格根据图一【辅材壳体计算】高度深度梯度表推导推荐深度
                if (d <= 0 && h > 0)
                {
                    // 调用已有的 DeriveDepthFromHeight，H<=1800 命中 400mm 动力配电柜
                    d = DeriveDepthFromHeight((int)h, shellConfig);
                }

                // 极端无规则兜底 --硬编码: 深度兜底--
                if (d <= 0) d = 200.0;

                // 校验尺寸有效性 (长宽均须大于等于 100mm)
                if (w >= 100 && h >= 100)
                {
                    item.Width = w;
                    item.Height = h;
                    item.Depth = d;
                    item.HasValidSize = true;

                    // 六面体封闭展开面积 (㎡): 2 * (W*H + W*D + H*D) / 1000000.0
                    double area = 2.0 * (w * h + w * d + h * d) / 1000000.0;
                    item.SurfaceArea = Math.Round(area, 3);
                    return true;
                }
            }

            // 3. 匹配 8~60 之间的纯数字模数 (如 12, 18, 24)
            if (int.TryParse(text.Trim(), out int pureNum) && pureNum >= 8 && pureNum <= 60)
            {
                item.IsModular = true;
                item.ModularDigits = pureNum;
                item.HasValidSize = true;
                return true;
            }

            return false;
        }

        /// <summary>
        /// 根据改价配置与箱体特征，计算单台箱体新单价、新合价与差额 (选项 A 默认：融合材质厚度矩阵与二层板加面)
        /// </summary>
        public static void CalculateSingleCabinetPrice(CabinetBatchItemDto item, CabinetBatchPriceConfigDto config)
        {
            double newUnit = item.OldUnitPrice;
            string remark = "";

            switch (config.Mode)
            {
                case "area":
                    // 模式 1：钣金表面积造价重算 (融合辅材壳体计算定额矩阵)
                    if (item.IsModular)
                    {
                        // PZ30 模数箱按位数折算
                        newUnit = Math.Max(item.ModularDigits * config.Pz30PricePerDigit, config.MinUnitPrice);
                        remark = $"PZ30模数箱: {item.ModularDigits}位 × {config.Pz30PricePerDigit}元/位";
                    }
                    else if (item.HasValidSize && item.Width > 0 && item.Height > 0)
                    {
                        var priceConfig = config.ShellPriceRules ?? new ShellPriceConfig();

                        // 1. 面数与面积计算 (采用封闭六面体标准 + 二层板智能加面)
                        double faceWD = priceConfig.FaceCountWD > 0 ? priceConfig.FaceCountWD : 2.0;
                        double faceHW = priceConfig.FaceCountHW > 0 ? priceConfig.FaceCountHW : 2.0;
                        double faceHD = priceConfig.FaceCountHD > 0 ? priceConfig.FaceCountHD : 2.0;

                        if (item.HasSecondPlate)
                        {
                            // 若含二层板，自动叠加面数
                            double extraHW = priceConfig.SecondPlateExtraHW > 0 ? priceConfig.SecondPlateExtraHW : 1.0;
                            faceHW += extraHW;
                        }

                        // 计算展开面积 (m²)
                        double areaWD = faceWD * (item.Width * item.Depth) / 1000000.0;
                        double areaHW = faceHW * (item.Height * item.Width) / 1000000.0;
                        double areaHD = faceHD * (item.Height * item.Depth) / 1000000.0;
                        double totalArea = Math.Round(areaWD + areaHW + areaHD, 3);
                        item.SurfaceArea = totalArea;

                        // 2. 检索对应材质和厚度的每平米单价
                        double pricePerSq = GetMaterialThicknessPrice(item.Material, item.Thickness, priceConfig);
                        item.PricePerSqMeter = pricePerSq;

                        // 3. 拼装标准型号
                        string secondSuffix = item.HasSecondPlate ? "(含二层板)" : "";
                        item.StandardModel = $"{item.Width}*{item.Height}*{item.Depth}-{item.Thickness:F1}mm {item.Material}{secondSuffix}";

                        // 4. 算得新单价
                        double calcPrice = Math.Round(totalArea * pricePerSq, 2);
                        newUnit = Math.Max(calcPrice, config.MinUnitPrice);
                        remark = $"【{item.Material} {item.Thickness:F1}mm】展开 {totalArea}㎡ × {pricePerSq}元/㎡{secondSuffix}";
                    }
                    else
                    {
                        // 尺寸缺失时使用配置的保底基准价
                        newUnit = config.MissingSizePrice;
                        remark = $"未检测到标准尺寸，按缺失保底价 {config.MissingSizePrice}元 计费";
                    }
                    break;

                case "percent":
                    // 模式 2：比例浮动模式
                    double factor = 1.0 + (config.PercentRatio / 100.0);
                    newUnit = Math.Max(item.OldUnitPrice * factor, config.MinUnitPrice);
                    remark = $"按原单价 {item.OldUnitPrice}元 浮动 {config.PercentRatio:+0.##;-0.##;0}%";
                    break;

                case "fixed":
                    // 模式 3：统一固定单价 (一口价)
                    newUnit = config.FixedUnitPrice;
                    remark = $"统一固定单价 {config.FixedUnitPrice}元/台";
                    break;

                default:
                    newUnit = item.OldUnitPrice;
                    remark = "保持当前单价";
                    break;
            }

            // 四舍五入保留两位小数
            newUnit = Math.Round(newUnit, 2);
            item.NewUnitPrice = newUnit;
            item.NewTotalPrice = Math.Round(newUnit * item.Qty, 2);
            item.DiffAmount = Math.Round(item.NewTotalPrice - item.OldTotalPrice, 2);
            item.Remark = remark;
        }

        /// <summary>
        /// 从壳体价格定额规则中根据材质和厚度检索每平米单价 (支持精准厚度与最接近厚度)
        /// </summary>
        public static double GetMaterialThicknessPrice(string material, double thickness, ShellPriceConfig priceConfig)
        {
            if (priceConfig?.MaterialPrices == null || priceConfig.MaterialPrices.Count == 0)
            {
                // 默认兜底价格 (1.2mm 58, 1.5mm 72, 2.0mm 96) --硬编码--
                return thickness <= 1.2 ? 58.0 : (thickness <= 1.5 ? 72.0 : 96.0);
            }

            // 查找材质项
            var mat = priceConfig.MaterialPrices.FirstOrDefault(m => string.Equals(m.MaterialName, material, StringComparison.OrdinalIgnoreCase));
            if (mat == null || mat.ThicknessPrices == null || mat.ThicknessPrices.Count == 0)
            {
                // 尝试以基准材质查找
                string defMat = !string.IsNullOrWhiteSpace(priceConfig.DefaultMaterial) ? priceConfig.DefaultMaterial : "镀锌板";
                mat = priceConfig.MaterialPrices.FirstOrDefault(m => string.Equals(m.MaterialName, defMat, StringComparison.OrdinalIgnoreCase));
            }

            if (mat != null && mat.ThicknessPrices != null && mat.ThicknessPrices.Count > 0)
            {
                // 优先精确匹配厚度 (误差0.05以内)
                var exact = mat.ThicknessPrices.FirstOrDefault(t => Math.Abs(t.Thickness - thickness) < 0.05);
                if (exact != null && exact.PricePerSqMeter > 0) return exact.PricePerSqMeter;

                // 未精确命中时取最接近厚度
                var nearest = mat.ThicknessPrices.OrderBy(t => Math.Abs(t.Thickness - thickness)).First();
                if (nearest != null && nearest.PricePerSqMeter > 0) return nearest.PricePerSqMeter;
            }

            // 兜底价格 --硬编码--
            return thickness <= 1.2 ? 58.0 : (thickness <= 1.5 ? 72.0 : 96.0);
        }

        /// <summary>
        /// 执行箱体批量改价回写至 Excel 工作表 (仅写 M 列表价，彻底保护公式链)
        /// </summary>
        /// <param name="payload">改价回写载荷实体</param>
        /// <returns>操作元组 (成功状态, 提示信息, 更新条数)</returns>
        public static (bool Success, string Message, int UpdatedCount) ApplyCabinetBatchPriceToExcel(CabinetApplyPricePayload payload)
        {
            if (payload == null || payload.Items == null || payload.Items.Count == 0)
            {
                return (false, "待改价回写的箱体明细集合为空！", 0);
            }

            int count = 0;

            try
            {
                // 获取当前活动 Excel 应用程序 (支持 ExcelDna 内部与外部自动化安全兜底)
                dynamic app = null;
                try { app = ExcelDnaUtil.Application; } catch { }
                if (app == null)
                {
                    try { app = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"); } catch { }
                }

                // 获取活动工作簿，若 ActiveWorkbook 暂时失去焦点则安全尝试 Workbooks[1]
                dynamic activeWb = null;
                try { activeWb = app?.ActiveWorkbook; } catch { }
                if (activeWb == null)
                {
                    try { if (app != null && app.Workbooks != null && app.Workbooks.Count > 0) activeWb = app.Workbooks[1]; } catch { }
                }
                if (activeWb == null) return (false, "未检测到已打开的活动工作簿！", 0);

                // 临时挂起屏幕刷新与自动重算，提升批量写入吞吐速度
                bool oldScreenUpdating = app.ScreenUpdating;
                dynamic oldCalculation = app.Calculation;
                app.ScreenUpdating = false;
                app.Calculation = -4135; // xlCalculationManual --硬编码: Excel COM 常量--

                try
                {
                    // 按工作表分组批量写入
                    var groupedBySheet = new Dictionary<string, List<CabinetBatchItemDto>>();
                    foreach (var item in payload.Items)
                    {
                        if (!item.Selected) continue;
                        if (!groupedBySheet.ContainsKey(item.SheetName))
                        {
                            groupedBySheet[item.SheetName] = new List<CabinetBatchItemDto>();
                        }
                        groupedBySheet[item.SheetName].Add(item);
                    }

                    // 遍历每个工作表进行安全回写
                    foreach (var kvp in groupedBySheet)
                    {
                        string sheetName = kvp.Key;
                        var items = kvp.Value;

                        dynamic ws = null;
                        try { ws = activeWb.Worksheets[sheetName]; } catch { }
                        if (ws == null) continue;

                        foreach (var itm in items)
                        {
                            // 校验物理行号合法性
                            if (itm.RowIndex <= 0) continue;

                            // 写入第 13 列 (M 列: 基准表价值)
                            ws.Cells[itm.RowIndex, 13].Value2 = itm.NewUnitPrice;
                            count++;
                        }
                    }
                }
                finally
                {
                    // 恢复屏幕刷新
                    app.ScreenUpdating = oldScreenUpdating;
                    // 恢复原计算模式
                    app.Calculation = oldCalculation;
                    // 强制全工作簿公式链级联刷新计算 (刷新 G/H/J/K 及箱柜小计)
                    app.Calculate();
                }

                return (true, $"已成功批量更新 {count} 处箱体单价，并完成全表公式级联重算！", count);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CabinetBatchPrice] ApplyCabinetBatchPriceToExcel 异常: {ex.Message}");
                return (false, $"批量回写失败: {ex.Message}", count);
            }
        }
    }
}

