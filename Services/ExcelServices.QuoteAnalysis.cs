using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelAddInDemo.Forms;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：工程报价全景分析与绚丽可视化大屏数据引擎
    /// 遵循规则 3 (公共业务收拢)、规则 7 (内存二维数组批量读取) 与每 3 行 1 行规范注释
    /// </summary>
    public static partial class ExcelServices
    {
        // 缓存报价全景分析大屏窗口实例 (支持单例与防重复弹出)
        private static QuoteAnalysisForm? _quoteAnalysisFormInstance;

        /// <summary>
        /// 供 Ribbon 菜单或快捷指令调用的报价全景分析大屏入口方法
        /// </summary>
        public static void ShowQuoteAnalysisDialog()
        {
            try
            {
                // 获取 Excel 主窗口的操作系统底层 HWND 句柄
                IntPtr excelHwnd = ExcelDnaSafeAccessor.GetWindowHandle();

                // 若窗口已实例化且未释放，直接激活并拉至前台
                if (_quoteAnalysisFormInstance != null && !_quoteAnalysisFormInstance.IsDisposed)
                {
                    // 激活已有窗口
                    _quoteAnalysisFormInstance.Activate();
                    return;
                }

                // 实例化全新的报价全景分析大屏窗口
                _quoteAnalysisFormInstance = new QuoteAnalysisForm();

                // 在 Excel 主窗口上以非模态置顶或模态呈现 (传入 ExcelWin32Window 宿主父句柄)
                if (excelHwnd != IntPtr.Zero)
                {
                    // 显示窗口并以 Excel 窗口为宿主父窗体
                    _quoteAnalysisFormInstance.Show(new ExcelWin32Window(excelHwnd));
                }
                else
                {
                    // 降级以普通 Show 形式弹出
                    _quoteAnalysisFormInstance.Show();
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志避免中断 Excel 消息循环
                LogHelper.WriteLog($"[报价分析大屏] 弹出窗口发生异常: {ex.Message}");
                MessageBox.Show($"打开报价分析大屏失败: {ex.Message}", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// 从当前活动工作簿中全量提取报价分析数据模型 (严格遵循规则 7 采用数组一次性读入内存)
        /// </summary>
        /// <returns>结构化报价全景分析传输对象</returns>
        public static QuoteAnalysisDto GetQuoteAnalysisData()
        {
            // 初始化数据传输对象根节点
            var dto = new QuoteAnalysisDto();

            try
            {
                // 获取当前正在运行的 Excel Application COM 接口
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return dto;

                // 获取当前活动的工程算价工作簿
                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null) return dto;

                // 基础概况初值：默认取文件名作为备选项目名
                string fallbackProjName = Path.GetFileNameWithoutExtension(Convert.ToString(activeWb.Name) ?? "");
                dto.Summary.ProjectName = fallbackProjName;

                // 收集活动工作簿中所有定义名称字典 (用于快速匹配 Cab_Det_K 行号)
                var definedNameRowMap = ExtractDefinedNamesMap(activeWb);

                // 缓存所有分类工作表引用
                var categorySheets = new List<dynamic>();
                dynamic? projectInfoSheet = null;
                dynamic? componentSumSheet = null;

                // 遍历当前工作簿下全部工作表
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    // 读取并清洗工作表名称
                    string sName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;

                    // 识别【项目信息】工作表 --硬编码: 系统固定工作表名称--
                    if (string.Equals(sName, "项目信息", StringComparison.OrdinalIgnoreCase))
                    {
                        projectInfoSheet = sheet;
                    }
                    // 识别【元件汇总表】或【元件汇总表１】 --硬编码: 系统固定工作表名称--
                    else if (sName.StartsWith("元件汇总表", StringComparison.OrdinalIgnoreCase))
                    {
                        if (componentSumSheet == null) componentSumSheet = sheet;
                    }
                    // 识别以【分类】开头的箱柜分类工作表 (如 分类1, 分类2) --硬编码: 分类表前缀--
                    else if (sName.StartsWith("分类", StringComparison.OrdinalIgnoreCase))
                    {
                        categorySheets.Add(sheet);
                    }
                }

                // 步骤 1：解析【项目信息】工作表
                if (projectInfoSheet != null)
                {
                    ParseProjectInfoSheet(projectInfoSheet, dto.Summary);
                }

                // 步骤 2：全量解析各分类工作表中的箱柜汇总与明细构成 (遵循规则 7 内存二维数组读取)
                foreach (dynamic catSheet in categorySheets)
                {
                    ParseCategorySheetCabinets(catSheet, dto, definedNameRowMap);
                }

                // 步骤 3：解析【元件汇总表】提取供应链排名与品牌体量
                if (componentSumSheet != null)
                {
                    ParseComponentSummarySheet(componentSumSheet, dto);
                }

                // 步骤 4：基于提取的箱柜矩阵全量汇总宏观指标与成本五大支柱
                FinalizeAnalysisMetrics(dto);
            }
            catch (Exception ex)
            {
                // 异常统一记录日志
                LogHelper.WriteLog($"[报价分析大屏] 提取全景数据异常: {ex.Message}\r\n{ex.StackTrace}");
            }

            return dto;
        }

        /// <summary>
        /// 解析【项目信息】工作表中的宏观商务与财务汇总参数 (规则 7: 二维数组批量读入)
        /// </summary>
        private static void ParseProjectInfoSheet(dynamic sheet, QuoteProjectSummaryDto summary)
        {
            try
            {
                // 获取已用单元格区域
                dynamic used = sheet.UsedRange;
                if (used == null) return;

                // 一次性批量读取为内存 object[,] 二维数组 (规避多次 COM 跨进程开销)
                object[,] values = used.Value2 as object[,];
                if (values == null) return;

                int rowCount = values.GetLength(0);
                int colCount = values.GetLength(1);

                // 扫描前 30 行提取项目名称、单号、公司等字段
                for (int r = 1; r <= Math.Min(rowCount, 30); r++)
                {
                    for (int c = 1; c <= Math.Min(colCount, 10); c++)
                    {
                        string val = Convert.ToString(values[r, c])?.Trim() ?? string.Empty;
                        if (string.IsNullOrEmpty(val)) continue;

                        // 识别“项目名称:”或“项目名称”标签 --硬编码: 单元格标签关键词--
                        if (val.Contains("项目名称") && string.IsNullOrEmpty(summary.ProjectName))
                        {
                            // 若当前单元格直接包含冒号，截取冒号后文字
                            int colonIdx = val.IndexOfAny(new[] { ':', '：' });
                            if (colonIdx >= 0 && colonIdx < val.Length - 1)
                            {
                                summary.ProjectName = val.Substring(colonIdx + 1).Trim();
                            }
                            else if (c + 1 <= colCount)
                            {
                                // 否则取右侧相邻单元格
                                summary.ProjectName = Convert.ToString(values[r, c + 1])?.Trim() ?? string.Empty;
                            }
                        }

                        // 识别“报价序号:”或“报价单号:” --硬编码: 单元格标签关键词--
                        if ((val.Contains("报价序号") || val.Contains("报价单号")) && string.IsNullOrEmpty(summary.QuoteNumber))
                        {
                            int colonIdx = val.IndexOfAny(new[] { ':', '：' });
                            if (colonIdx >= 0 && colonIdx < val.Length - 1)
                            {
                                summary.QuoteNumber = val.Substring(colonIdx + 1).Trim();
                            }
                            else if (c + 1 <= colCount)
                            {
                                summary.QuoteNumber = Convert.ToString(values[r, c + 1])?.Trim() ?? string.Empty;
                            }
                        }

                        // 识别单位名称 (如: 扬州华科智能科技有限公司) --硬编码: 单元格标签关键词--
                        if (val.Contains("单位名称") && string.IsNullOrEmpty(summary.CompanyName))
                        {
                            if (c + 1 <= colCount)
                            {
                                summary.CompanyName = Convert.ToString(values[r, c + 1])?.Trim() ?? string.Empty;
                            }
                        }
                    }
                }

                // 扫描分类汇总行 (通常在第 25~55 行，包含“总价”、“小计”或“合计”)
                for (int r = 25; r <= Math.Min(rowCount, 60); r++)
                {
                    string label = Convert.ToString(values[r, 2])?.Trim() ?? string.Empty; // B 列
                    // 识别“合计”行或“小计”行 --硬编码: 汇总行标签--
                    if (label == "合计" || label == "小计")
                    {
                        // D 列总价、E 列成本、F 列毛利、G 列毛利率
                        double totPrice = ConvertToDouble(values[r, 4]);
                        double cost = ConvertToDouble(values[r, 5]);
                        double profit = ConvertToDouble(values[r, 6]);
                        double profitRate = ConvertToDouble(values[r, 7]);

                        if (totPrice > 0 && summary.TotalQuotePrice == 0)
                        {
                            summary.TotalQuotePrice = totPrice;
                            summary.TotalCostPrice = cost;
                            summary.TotalGrossProfit = profit;
                            summary.GrossProfitRate = profitRate;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[报价分析大屏] 解析项目信息表异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 解析分类工作表（箱柜顶部汇总区及底部明细计费区） (规则 7: 二维数组批量读入)
        /// </summary>
        private static void ParseCategorySheetCabinets(dynamic sheet, QuoteAnalysisDto dto, Dictionary<string, int> definedNameRowMap)
        {
            try
            {
                string sheetName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;
                dynamic used = sheet.UsedRange;
                if (used == null) return;

                // 规则 7: 一次性读取全表二维数组到内存
                object[,] values = used.Value2 as object[,];
                if (values == null) return;

                int rowCount = values.GetLength(0);
                int colCount = values.GetLength(1);

                // 寻找顶部箱柜汇总行表头 (通常在第 6 行，包含“序号”、“柜号”、“箱柜名称”)
                int headerRow = 6;
                for (int r = 1; r <= Math.Min(rowCount, 15); r++)
                {
                    string colB = Convert.ToString(values[r, 2])?.Trim() ?? "";
                    string colC = Convert.ToString(values[r, 3])?.Trim() ?? "";
                    if (colB == "柜号" || colC == "箱柜名称" || colB == "箱柜名称")
                    {
                        headerRow = r;
                        break;
                    }
                }

                // 从表头下一行起读取箱柜汇总行，直到遇到明细区标识或连续空行
                int cabIndex = 1;
                for (int r = headerRow + 1; r <= rowCount; r++)
                {
                    // A 列为序号
                    string seqStr = Convert.ToString(values[r, 1])?.Trim() ?? "";
                    // B 列为柜号或箱柜名称
                    string cabName = Convert.ToString(values[r, 2])?.Trim() ?? "";

                    // 若 B 列为空或出现“柜号:”明细特征，则说明顶部汇总区结束
                    if (string.IsNullOrEmpty(cabName) || cabName.StartsWith("柜号:") || cabName.StartsWith("柜号："))
                    {
                        break;
                    }

                    // 尝试解析序号数字
                    if (!int.TryParse(seqStr, out int seqNum))
                    {
                        seqNum = cabIndex;
                    }

                    // 读取各列字段
                    string cabModel = Convert.ToString(values[r, 4])?.Trim() ?? ""; // D 列型号
                    if (string.IsNullOrEmpty(cabModel)) cabModel = Convert.ToString(values[r, 3])?.Trim() ?? ""; // 备选 C 列
                    double qty = ConvertToDouble(values[r, 6]);                      // F 列数量
                    if (qty <= 0) qty = 1;
                    double unitPrice = ConvertToDouble(values[r, 7]);                // G 列单价
                    double totalPrice = ConvertToDouble(values[r, 8]);               // H 列总价
                    if (totalPrice <= 0 && unitPrice > 0) totalPrice = unitPrice * qty;
                    double costPrice = ConvertToDouble(values[r, 10]);               // J 列成本总价
                    double grossProfit = ConvertToDouble(values[r, 11]);             // K 列毛利
                    double profitRate = ConvertToDouble(values[r, 12]);              // L 列毛利率
                    string size = Convert.ToString(values[r, 13])?.Trim() ?? "";     // M 列尺寸或安装方式

                    // 构建箱柜模型
                    var cabItem = new CabinetAnalysisItemDto
                    {
                        Index = seqNum,
                        SheetName = sheetName,
                        Name = cabName,
                        Model = cabModel,
                        Quantity = qty,
                        UnitPrice = unitPrice,
                        TotalPrice = totalPrice,
                        CostPrice = costPrice,
                        GrossProfit = grossProfit,
                        GrossProfitRate = profitRate,
                        Size = size,
                        InstallType = DetectInstallType(cabName, size),
                        FunctionalCategory = ClassifyCabinetCategory(cabName)
                    };

                    // 尝试匹配对应箱柜明细锚点行号 (优先从定义名称 Cab_Det_{Index} 查找)
                    string detTagKey = $"{sheetName}!Cab_Det_{seqNum}";
                    if (definedNameRowMap.TryGetValue(detTagKey, out int detRow))
                    {
                        cabItem.DetRow = detRow;
                    }
                    else if (definedNameRowMap.TryGetValue($"Cab_Det_{seqNum}", out int globalDetRow))
                    {
                        cabItem.DetRow = globalDetRow;
                    }

                    // 步骤 2.2：从二维数组中提取该箱柜明细段的计费构成 (元器件、箱体、辅材、人工、成套费)
                    ExtractCabinetCostBreakdown(values, rowCount, colCount, cabItem);

                    dto.Cabinets.Add(cabItem);
                    cabIndex++;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[报价分析大屏] 解析分类表箱柜异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 从二维数组中扫描提取单台箱柜底部的五大计费支柱数据 (规则 7)
        /// </summary>
        private static void ExtractCabinetCostBreakdown(object[,] values, int rowCount, int colCount, CabinetAnalysisItemDto cabItem)
        {
            if (cabItem.DetRow <= 0 || cabItem.DetRow >= rowCount) return;

            int startRow = cabItem.DetRow;
            // 向下扫描至多 120 行寻找计费小计与费用项
            for (int r = startRow; r <= Math.Min(rowCount, startRow + 120); r++)
            {
                // B 列标签
                string label = Convert.ToString(values[r, 2])?.Trim() ?? "";

                // 若遇到下一个箱柜的起始标志，提前终止
                if (r > startRow + 5 && (label.StartsWith("柜号:") || label.StartsWith("柜号：")))
                {
                    break;
                }

                // 辅材费用行 --硬编码: 计费标签--
                if (label.Contains("辅材") || label.Contains("辅料"))
                {
                    cabItem.AuxMaterialsCost = ConvertToDouble(values[r, 8]); // H 列总价
                }
                // 箱体费用行 --硬编码: 计费标签--
                else if (label.Contains("箱体") || label.Contains("壳体"))
                {
                    cabItem.CabinetBoxCost = ConvertToDouble(values[r, 8]); // H 列总价
                }
                // 小计行 (元器件 + 辅材 + 箱体小计) --硬编码: 计费标签--
                else if (label == "小计")
                {
                    double subsum = ConvertToDouble(values[r, 8]);
                    // 元器件小计估算 = 小计 - 辅材 - 箱体
                    cabItem.ComponentsCost = Math.Max(0, subsum - cabItem.AuxMaterialsCost - cabItem.CabinetBoxCost);
                }
                // 人工费行 --硬编码: 计费标签--
                else if (label.Contains("人工"))
                {
                    cabItem.LaborCost = ConvertToDouble(values[r, 8]);
                }
                // 综合成套费行 --硬编码: 计费标签--
                else if (label.Contains("成套费") || label.Contains("综合") || label.Contains("管理费"))
                {
                    cabItem.CompleteSetFee = ConvertToDouble(values[r, 8]);
                }
            }

            // 若元器件小计为 0，用成本总价减去箱体/人工/辅材兜底推导
            if (cabItem.ComponentsCost <= 0 && cabItem.CostPrice > 0)
            {
                cabItem.ComponentsCost = Math.Max(0, cabItem.CostPrice - cabItem.CabinetBoxCost - cabItem.AuxMaterialsCost - cabItem.LaborCost);
            }
        }

        /// <summary>
        /// 解析【元件汇总表】提取全案供应链前 20 核心物料及品牌分布 (规则 7: 二维数组批量读入)
        /// </summary>
        private static void ParseComponentSummarySheet(dynamic sheet, QuoteAnalysisDto dto)
        {
            try
            {
                dynamic used = sheet.UsedRange;
                if (used == null) return;

                // 规则 7: 一次性读取全表二维数组
                object[,] values = used.Value2 as object[,];
                if (values == null) return;

                int rowCount = values.GetLength(0);
                int colCount = values.GetLength(1);

                // 寻找表头行 (通常第 4 行或第 5 行，含“元件名称”)
                int headerRow = 4;
                for (int r = 1; r <= Math.Min(rowCount, 10); r++)
                {
                    string colB = Convert.ToString(values[r, 2])?.Trim() ?? "";
                    if (colB == "元件名称")
                    {
                        headerRow = r;
                        break;
                    }
                }

                var compList = new List<ComponentRankingDto>();
                var brandMap = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

                // 从表头下一行或第 6 行开始读取元器件物料
                int startRow = headerRow + 1;
                if (Convert.ToString(values[startRow, 1])?.Trim() == "序号") startRow++;

                for (int r = startRow; r <= rowCount; r++)
                {
                    string name = Convert.ToString(values[r, 2])?.Trim() ?? ""; // B 列元件名称
                    if (string.IsNullOrEmpty(name)) continue;

                    string model = Convert.ToString(values[r, 4])?.Trim() ?? ""; // D 列型号规格
                    if (string.IsNullOrEmpty(model)) model = Convert.ToString(values[r, 3])?.Trim() ?? ""; // C 列原型号
                    double qty = ConvertToDouble(values[r, 6]);                   // F 列数量
                    double price = ConvertToDouble(values[r, 7]);                 // G 列单价
                    double totPrice = ConvertToDouble(values[r, 8]);              // H 列总价
                    if (totPrice <= 0 && qty > 0 && price > 0) totPrice = qty * price;
                    string brand = Convert.ToString(values[r, 9])?.Trim() ?? "通用国优"; // I 列厂家/品牌 --硬编码: 默认品牌--

                    // 纳入元器件列表
                    compList.Add(new ComponentRankingDto
                    {
                        Name = name,
                        Model = model,
                        Manufacturer = string.IsNullOrEmpty(brand) ? "国优" : brand,
                        TotalQuantity = qty,
                        UnitPrice = price,
                        TotalPrice = totPrice
                    });

                    // 累计品牌采购总额
                    string cleanBrand = string.IsNullOrEmpty(brand) ? "通用/国标" : brand;
                    if (!brandMap.ContainsKey(cleanBrand)) brandMap[cleanBrand] = 0;
                    brandMap[cleanBrand] += totPrice;
                }

                dto.Summary.TotalComponentCount = compList.Count;

                // 计算元器件采购总池金额
                double allCompTotal = compList.Sum(c => c.TotalPrice);
                if (allCompTotal <= 0) allCompTotal = 1;

                // 排序并取前 20 核心物料
                var topList = compList.OrderByDescending(c => c.TotalPrice).Take(20).ToList();
                int rank = 1;
                foreach (var item in topList)
                {
                    item.Rank = rank++;
                    item.Percentage = Math.Round(item.TotalPrice / allCompTotal, 4);
                }
                dto.TopComponents = topList;

                // 构建品牌分布列表
                double brandTotal = brandMap.Values.Sum();
                if (brandTotal <= 0) brandTotal = 1;

                dto.BrandDistributions = brandMap
                    .Select(kv => new BrandDistributionDto
                    {
                        BrandName = kv.Key,
                        TotalPrice = Math.Round(kv.Value, 2),
                        Percentage = Math.Round(kv.Value / brandTotal, 4)
                    })
                    .OrderByDescending(b => b.TotalPrice)
                    .ToList();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[报价分析大屏] 解析元件汇总表异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 全量聚合财务指标、五大成本支柱与功能分类分布 (完备性对齐)
        /// </summary>
        private static void FinalizeAnalysisMetrics(QuoteAnalysisDto dto)
        {
            try
            {
                // 若 Summary 缺少总价或台数，从 Cabinets 列表重新汇总
                if (dto.Cabinets.Count > 0)
                {
                    dto.Summary.TotalCabinetCount = dto.Cabinets.Count;

                    if (dto.Summary.TotalQuotePrice <= 0)
                    {
                        dto.Summary.TotalQuotePrice = Math.Round(dto.Cabinets.Sum(c => c.TotalPrice), 2);
                        dto.Summary.TotalCostPrice = Math.Round(dto.Cabinets.Sum(c => c.CostPrice), 2);
                        dto.Summary.TotalGrossProfit = Math.Round(dto.Summary.TotalQuotePrice - dto.Summary.TotalCostPrice, 2);
                        if (dto.Summary.TotalQuotePrice > 0)
                        {
                            dto.Summary.GrossProfitRate = Math.Round(dto.Summary.TotalGrossProfit / dto.Summary.TotalQuotePrice, 4);
                        }
                    }

                    // 聚合成本五大支柱
                    double compSum = dto.Cabinets.Sum(c => c.ComponentsCost);
                    double boxSum = dto.Cabinets.Sum(c => c.CabinetBoxCost);
                    double auxSum = dto.Cabinets.Sum(c => c.AuxMaterialsCost);
                    double laborSum = dto.Cabinets.Sum(c => c.LaborCost);
                    double setSum = dto.Cabinets.Sum(c => c.CompleteSetFee);

                    // 若细项总和过小，基于总成本按行业黄金比例弹性推导补齐 --硬编码: 行业参考比例推导--
                    if (compSum + boxSum + laborSum <= 0 && dto.Summary.TotalCostPrice > 0)
                    {
                        compSum = dto.Summary.TotalCostPrice * 0.65; // 元器件约 65%
                        boxSum = dto.Summary.TotalCostPrice * 0.18;  // 箱体约 18%
                        laborSum = dto.Summary.TotalCostPrice * 0.10; // 人工约 10%
                        auxSum = dto.Summary.TotalCostPrice * 0.07;  // 辅材约 7%
                        setSum = dto.Summary.TotalQuotePrice - dto.Summary.TotalCostPrice;
                    }

                    dto.CostStructure = new CostBreakdownDto
                    {
                        ComponentsTotal = Math.Round(compSum, 2),
                        CabinetBoxTotal = Math.Round(boxSum, 2),
                        AuxMaterialsTotal = Math.Round(auxSum, 2),
                        LaborTotal = Math.Round(laborSum, 2),
                        CompleteSetTotal = Math.Round(setSum, 2),
                        ProfitTotal = Math.Round(dto.Summary.TotalGrossProfit, 2)
                    };

                    // 聚合功能分类分布 (动力控制、照明配电、双电源、消防等)
                    double totalMoney = dto.Summary.TotalQuotePrice > 0 ? dto.Summary.TotalQuotePrice : 1;
                    dto.CategoryDistributions = dto.Cabinets
                        .GroupBy(c => c.FunctionalCategory)
                        .Select(g => new CategoryDistributionDto
                        {
                            CategoryName = g.Key,
                            CabinetCount = g.Count(),
                            TotalPrice = Math.Round(g.Sum(c => c.TotalPrice), 2),
                            Percentage = Math.Round(g.Sum(c => c.TotalPrice) / totalMoney, 4)
                        })
                        .OrderByDescending(c => c.TotalPrice)
                        .ToList();
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[报价分析大屏] 汇总聚合指标异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 穿透跳转至指定工作表的指定箱柜行并在 Excel 视口中高亮居中
        /// </summary>
        /// <param name="sheetName">目标工作表名称</param>
        /// <param name="rowNumber">目标行号</param>
        /// <returns>是否定位成功</returns>
        public static bool NavigateToCabinetInExcel(string sheetName, int rowNumber)
        {
            if (string.IsNullOrWhiteSpace(sheetName) || rowNumber <= 0) return false;

            try
            {
                // 利用 ExcelAsyncUtil.QueueAsMacro 在主线程宏队列中安全执行 COM 视口切换
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                        if (app == null) return;

                        dynamic? activeWb = app.ActiveWorkbook;
                        if (activeWb == null) return;

                        // 寻址目标工作表
                        dynamic targetSheet = activeWb.Sheets[sheetName];
                        if (targetSheet != null)
                        {
                            // 激活工作表
                            targetSheet.Activate();

                            // 选中目标行
                            dynamic targetRange = targetSheet.Rows[rowNumber];
                            targetRange.Select();

                            // 将视口垂直居中滚动到该行上方 3 行位置
                            app.ActiveWindow.ScrollRow = Math.Max(1, rowNumber - 3);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogHelper.WriteLog($"[报价分析大屏] 定位箱柜宏异常: {ex.Message}");
                    }
                });

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[报价分析大屏] 调度定位箱柜异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 提取活动工作簿所有 Defined Names 的映射字典 (Name -> Row)
        /// </summary>
        private static Dictionary<string, int> ExtractDefinedNamesMap(dynamic workbook)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (dynamic nameObj in workbook.Names)
                {
                    string nameTag = Convert.ToString(nameObj.Name)?.Trim() ?? "";
                    if (string.IsNullOrEmpty(nameTag)) continue;

                    try
                    {
                        // 读取引用的 Range 对象的行号
                        dynamic refRange = nameObj.RefersToRange;
                        if (refRange != null)
                        {
                            int row = Convert.ToInt32(refRange.Row);
                            string sheetName = Convert.ToString(refRange.Worksheet?.Name) ?? "";

                            // 存储纯名称与带表名前缀复合 Key
                            map[nameTag] = row;
                            if (!string.IsNullOrEmpty(sheetName))
                            {
                                map[$"{sheetName}!{nameTag}"] = row;
                            }
                        }
                    }
                    catch
                    {
                        // 忽略无法解析范围的无效定义名称
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[报价分析大屏] 扫描定义名称异常: {ex.Message}");
            }
            return map;
        }

        /// <summary>
        /// 智能推断箱柜功能大类 (动力控制 / 照明配电 / 消防应急 / 电梯系统 / 专用设备)
        /// </summary>
        private static string ClassifyCabinetCategory(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "通用配电箱"; // --硬编码: 默认分类名称--

            // 消防/排烟/稳压类 --硬编码: 关键词匹配--
            if (name.Contains("消防") || name.Contains("排烟") || name.Contains("稳压") || name.Contains("APW") || name.Contains("APY"))
                return "消防应急控制";

            // 电梯/货梯类 --硬编码: 关键词匹配--
            if (name.Contains("电梯") || name.Contains("货梯") || name.Contains("DT") || name.Contains("HT"))
                return "电梯动力系统";

            // 双电源系统 --硬编码: 关键词匹配--
            if (name.Contains("双电源") || name.Contains("ATS") || name.Contains("ALp") || name.Contains("ALEX"))
                return "双电源互投箱";

            // 照明类 --硬编码: 关键词匹配--
            if (name.Contains("照明") || name.Contains("ALB") || name.Contains("ALZ") || name.Contains("AL1") || name.Contains("AL2") || name.Contains("ALa") || name.Contains("ALb") || name.Contains("ALc"))
                return "照明配电系统";

            // 泵房与水处理动力 --硬编码: 关键词匹配--
            if (name.Contains("泵") || name.Contains("潜污") || name.Contains("PS"))
                return "泵房动力系统";

            // 特殊控制箱 (BXF等) --硬编码: 关键词匹配--
            if (name.Contains("BXF") || name.Contains("控制") || name.Contains("ADYZ"))
                return "智能控制箱";

            return "动力与综合配电"; // --硬编码: 兜底分类名称--
        }

        /// <summary>
        /// 智能推断箱柜安装方式与环境特征 (暗装 / 明装 / 户外 / 落地)
        /// </summary>
        private static string DetectInstallType(string name, string size)
        {
            string combined = (name + " " + size).ToLower();
            if (combined.Contains("户外")) return "户外型"; // --硬编码: 安装环境标签--
            if (combined.Contains("暗") || combined.Contains("嵌")) return "嵌入暗装"; // --硬编码: 安装方式标签--
            if (combined.Contains("落地")) return "落地安装"; // --硬编码: 安装方式标签--
            return "明装壁挂"; // --硬编码: 默认安装方式--
        }

        /// <summary>
        /// 安全数值转换助手函数
        /// </summary>
        private static double ConvertToDouble(object? val)
        {
            if (val == null) return 0;
            if (val is double d) return d;
            if (val is float f) return f;
            if (val is int i) return i;
            if (val is long l) return l;
            if (val is decimal m) return (double)m;

            string s = Convert.ToString(val)?.Trim() ?? "";
            if (string.IsNullOrEmpty(s)) return 0;

            // 过滤货币符号或多余逗号 --硬编码: 字符过滤--
            s = s.Replace("￥", "").Replace("¥", "").Replace(",", "").Trim();
            if (double.TryParse(s, out double res)) return res;
            return 0;
        }
    }
}
