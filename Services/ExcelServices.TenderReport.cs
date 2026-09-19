using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：常规样式甲方投标报表生成与导出
    /// </summary>
    public static partial class ExcelServices
    {
        // 投标报表导出窗口静态单例引用 (可空)
        private static Forms.TenderReportRegularForm? _tenderReportRegularForm;

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“常规样式投标报表”导出向导窗口 (非模态，可编辑 Excel)
        /// </summary>
        public static void ShowTenderReportRegularDialog()
        {
            try
            {
                // 以非模态方式展示常规报表窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _tenderReportRegularForm, () => new Forms.TenderReportRegularForm());
            }
            catch (Exception ex)
            {
                // 弹出异常提示信息
                System.Windows.Forms.MessageBox.Show($"弹出常规样式投标报表窗口失败: {ex.Message}", "错误提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 扫描当前活动工作簿，抓取项目基础信息及所有分类工作表中的箱柜统计
        /// </summary>
        /// <returns>工程信息与分类概况列表元组</returns>
        public static (TenderReportProjectInfo ProjectInfo, List<TenderReportCategoryGroup> Categories) GetTenderReportInitialData()
        {
            // 初始化返回的工程基础信息对象
            var projectInfo = new TenderReportProjectInfo();
            // 初始化返回的分类工作表概况列表
            var categories = new List<TenderReportCategoryGroup>();

            try
            {
                // 安全获取当前 Excel 宿主应用程序 COM 实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return (projectInfo, categories);

                // 获取当前活动工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null) return (projectInfo, categories);

                // 默认使用工作簿文件名作为项目备用名称
                string wbName = Path.GetFileNameWithoutExtension(Convert.ToString(activeWb.Name) ?? "");
                projectInfo.ProjectName = wbName;
                projectInfo.QuotationDate = DateTime.Now.ToString("yyyy-MM-dd");

                // 1. 尝试从【项目信息】工作表一次性读取工程与客户信息
                try
                {
                    dynamic? infoSheet = null;
                    try { infoSheet = activeWb.Sheets["项目信息"]; } catch { }

                    if (infoSheet != null)
                    {
                        // 一次性读取项目信息工作表 A1:C30 区域至内存二维数组 (规则 7 / 规则 12)
                        dynamic infoRange = infoSheet.Range["A1:C30"];
                        object[,] infoMatrix = (object[,])infoRange.Value2;

                        // 提取工程名称 (Cell B5)
                        string pName = Convert.ToString(infoMatrix[5, 2]) ?? "";
                        if (!string.IsNullOrWhiteSpace(pName)) projectInfo.ProjectName = pName;

                        // 提取报价文件编号 / 单号 (Cell B7)
                        projectInfo.QuotationNumber = Convert.ToString(infoMatrix[7, 2]) ?? "";
                        // 提取报价编制人 (Cell B8)
                        projectInfo.Bidder = Convert.ToString(infoMatrix[8, 2]) ?? "";
                        // 提取报价日期 (Cell B9)
                        string pDate = Convert.ToString(infoMatrix[9, 2]) ?? "";
                        if (!string.IsNullOrWhiteSpace(pDate)) projectInfo.QuotationDate = pDate;
                        // 提取项目描述 / 备注 (Cell B12)
                        projectInfo.ProjectRemark = Convert.ToString(infoMatrix[12, 2]) ?? "";

                        // 提取客户甲方单位名称 (Cell B14)
                        projectInfo.CustomerName = Convert.ToString(infoMatrix[14, 2]) ?? "";
                        // 提取客户联系人 (Cell B15)
                        projectInfo.CustomerContact = Convert.ToString(infoMatrix[15, 2]) ?? "";
                        // 提取客户联系电话 (Cell B16)
                        projectInfo.CustomerPhone = Convert.ToString(infoMatrix[16, 2]) ?? "";

                        // 提取供货单位本单位名称 (Cell B22 或 Cell B1)
                        string fName = Convert.ToString(infoMatrix[22, 2]) ?? Convert.ToString(infoMatrix[1, 2]) ?? "";
                        projectInfo.FactoryName = fName;
                        // 提取供货单位联系人 (Cell B24)
                        projectInfo.FactoryContact = Convert.ToString(infoMatrix[24, 2]) ?? "";
                        // 提取供货单位电话 (Cell B25)
                        projectInfo.FactoryPhone = Convert.ToString(infoMatrix[25, 2]) ?? "";
                    }
                }
                catch (Exception exInfo)
                {
                    // 记录读取项目信息异常
                    LogHelper.WriteLog($"[TenderReport] 读取项目信息工作表异常: {exInfo.Message}");
                }

                // 2. 遍历所有工作表，提取有效分类及箱柜基础台数统计
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    string sheetName = Convert.ToString(sheet.Name) ?? string.Empty;
                    string trimmedName = sheetName.Trim();

                    // 过滤已知的系统非分类辅助表 --硬编码--
                    if (string.Equals(trimmedName, "项目信息", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, "元件汇总表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, "汇总调价表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, "封面", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, "屏柜汇总表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, "屏柜分项表", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // 调用全局通用工具方法提取当前表的有效箱柜锚点 (规则 6 / 规则 11)
                    var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);

                    int cabTotalQty = 0;
                    decimal cabTotalAmount = 0;
                    var cabinetList = new List<TenderReportCabinetItem>();

                    if (validCabinets.Count > 0)
                    {
                        // 扫描每个箱柜的汇总行或信息行进行汇总计算
                        for (int i = 0; i < validCabinets.Count; i++)
                        {
                            var cab = validCabinets[i];
                            int cabQty = 1;
                            decimal cabPrice = 0;
                            string cabNo = "";
                            string cabName = "";
                            string cabModel = "";
                            string cabRemark = "";

                            try
                            {
                                // 若具备顶部汇总行锚点，直接读取汇总行数据
                                if (cab.Value.Sum != null)
                                {
                                    int sumRow = Convert.ToInt32(cab.Value.Sum.Row);
                                    // 一次性读取汇总行 A 列到 J 列数据数组 (规则 12)
                                    dynamic sumRange = sheet.Range[sheet.Cells[sumRow, 1], sheet.Cells[sumRow, 10]];
                                    object[,] sumMatrix = (object[,])sumRange.Value2;

                                    cabNo = Convert.ToString(sumMatrix[1, 2]) ?? "";      // 柜号 (Col 2)
                                    cabName = Convert.ToString(sumMatrix[1, 3]) ?? "";    // 箱柜名称 (Col 3)
                                    cabModel = Convert.ToString(sumMatrix[1, 4]) ?? "";   // 箱柜型号 (Col 4)

                                    // 读取箱柜数量 (优先 Col 6，兜底 Col 5)
                                    object qVal = sumMatrix[1, 6] ?? sumMatrix[1, 5];
                                    if (qVal != null && int.TryParse(Convert.ToString(qVal), out int parsedQ) && parsedQ > 0)
                                    {
                                        cabQty = parsedQ;
                                    }

                                    // 读取单价与总价 (Col 7 单价，Col 8 总价)
                                    object pVal = sumMatrix[1, 7];
                                    if (pVal != null && decimal.TryParse(Convert.ToString(pVal), out decimal parsedP))
                                    {
                                        cabPrice = parsedP;
                                    }

                                    cabRemark = Convert.ToString(sumMatrix[1, 9]) ?? "";  // 备注 (Col 9)
                                }
                                else if (cab.Value.Det != null)
                                {
                                    // 兜底从明细信息行读取
                                    int detRow = Convert.ToInt32(cab.Value.Det.Row);
                                    cabNo = Convert.ToString(sheet.Cells[detRow, 2].Value2) ?? "";
                                    cabName = Convert.ToString(sheet.Cells[detRow, 8].Value2) ?? "";
                                    cabModel = Convert.ToString(sheet.Cells[detRow, 4].Value2) ?? "";
                                }
                            }
                            catch { }

                            cabTotalQty += cabQty;
                            cabTotalAmount += (cabPrice * cabQty);

                            // 构建箱柜简要项加入分类
                            cabinetList.Add(new TenderReportCabinetItem
                            {
                                Index = i + 1,
                                CabinetNo = cabNo,
                                CabinetName = cabName,
                                CabinetModel = cabModel,
                                Quantity = cabQty,
                                UnitPrice = cabPrice,
                                TotalPrice = cabPrice * cabQty,
                                Remark = cabRemark
                            });
                        }
                    }

                    // 记录该工作表的扫描详情日志
                    LogHelper.WriteLog($"[TenderReport] 扫描工作表 [{sheetName}]: 锚点数={validCabinets.Count}, 箱柜台数={cabTotalQty}, 合价={cabTotalAmount}");

                    // 仅当扫描到有效箱柜时纳入分类工作表导出候选
                    if (validCabinets.Count > 0 || cabinetList.Count > 0)
                    {
                        // 添加分类分组
                        categories.Add(new TenderReportCategoryGroup
                        {
                            CategoryName = sheetName,
                            TotalCabinetCount = cabTotalQty,
                            TotalAmount = cabTotalAmount,
                            IsSelected = true,
                            Cabinets = cabinetList
                        });
                    }
                    else
                    {
                        // 若未扫描到锚点但不是系统表，保留作为候选分类 (台数0) 供用户按需选勾
                        categories.Add(new TenderReportCategoryGroup
                        {
                            CategoryName = sheetName,
                            TotalCabinetCount = 0,
                            TotalAmount = 0,
                            IsSelected = true,
                            Cabinets = cabinetList
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                // 全局捕获异常记录日志
                LogHelper.WriteLog($"[TenderReport] 抓取初始报表数据异常: {ex.Message}");
            }

            // 返回提取的项目信息与分类列表
            return (projectInfo, categories);
        }

        /// <summary>
        /// 核心生成引擎：执行常规样式投标报表导出并输出为全新独立的 Excel 工作簿
        /// </summary>
        /// <param name="config">前端确认提交的导出参数配置</param>
        /// <returns>包含成功状态与文件路径的导出结果</returns>
        public static TenderReportExportResult ExportTenderReportRegular(TenderReportExportConfig config)
        {
            var result = new TenderReportExportResult();

            if (config == null)
            {
                result.Message = "导出配置对象为空，无法执行报表导出";
                return result;
            }

            try
            {
                // 获取 Excel 宿主应用程序实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    result.Message = "无法连接至 Excel 宿主应用程序";
                    return result;
                }

                // 获取当前活动算价工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null)
                {
                    result.Message = "当前没有打开的活动算价工作簿";
                    return result;
                }

                // 1. 定位常规样式报表模板文件全路径 (Resources\TenderReport_Regular.xlsx)
                string appDir = Tool.GetAppDirectory();
                string templatePath = Path.Combine(appDir, "Resources", "TenderReport_Regular.xlsx");

                // 若首选路径不存在，尝试从基目录或当前目录检索兜底
                if (!File.Exists(templatePath))
                {
                    templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "TenderReport_Regular.xlsx");
                }

                // 再次校验模板文件是否存在
                if (!File.Exists(templatePath))
                {
                    result.Message = $"未找到常规样式报表模板文件: {templatePath}";
                    return result;
                }

                // 2. 深度采集选中分类中的完整箱柜及元器件明细数据 (内存全量提取，无频繁 COM 穿梭)
                List<TenderReportCategoryGroup> exportCategories = CollectFullCategoryDetails((object)activeWb, config.SelectedCategories);
                if (exportCategories.Count == 0)
                {
                    result.Message = "选中的分类中未找到有效的箱柜或元器件数据，已终止导出";
                    return result;
                }

                // 3. 计算生成目标文件路径
                string targetPath = config.TargetFilePath;
                if (string.IsNullOrWhiteSpace(targetPath))
                {
                    // 默认生成在工程同级目录下，按项目名称及时间戳命名
                    string activeWbDir = "";
                    try { activeWbDir = activeWb.Path; } catch { }
                    if (string.IsNullOrWhiteSpace(activeWbDir))
                    {
                        activeWbDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    }

                    string safeProjName = string.IsNullOrWhiteSpace(config.ProjectInfo.ProjectName) ? "工程" : config.ProjectInfo.ProjectName;
                    // 过滤文件名中的非法字符
                    foreach (char invalidChar in Path.GetInvalidFileNameChars())
                    {
                        safeProjName = safeProjName.Replace(invalidChar, '_');
                    }

                    string fileName = $"{safeProjName}_甲方投标报表_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                    targetPath = Path.Combine(activeWbDir, fileName);
                }

                // 临时拷贝模板作为全新的目标报表工作簿物理文件
                File.Copy(templatePath, targetPath, true);

                // 通过 Excel COM 打开该目标工作簿进行全量数据填充
                dynamic reportWb = app.Workbooks.Open(targetPath);
                if (reportWb == null)
                {
                    result.Message = "打开目标报表工作簿失败";
                    return result;
                }

                // 临时挂起屏幕刷新、事件响应与自动计算提升填充性能
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.Calculation = -4135; // xlCalculationManual 手动重算

                try
                {
                    // 提取用户前端提交的高级偏好配置
                    var settings = config.Settings ?? new TenderReportSettings();

                    // 4. 填充《封面》工作表
                    PopulateCoverWorksheet(reportWb, config.ProjectInfo, config.IncludeCover);

                    // 5. 填充《屏柜汇总表》工作表 (记录各箱柜在汇总表的物理行，用于一键定位双向跳转)
                    Dictionary<int, int> cabSumRowMap = PopulateSummaryWorksheet(reportWb, config.ProjectInfo, exportCategories, config.IncludeSummary, settings);

                    // 6. 填充《屏柜分项表》工作表 (支持按分类独立分 Sheet 与单表输出两种模式)
                    PopulateDetailWorksheets(reportWb, config.ProjectInfo, exportCategories, config.IncludeDetail, settings, cabSumRowMap);

                    // 恢复自动计算并执行一次全局重算（调用 Application.Calculate 避免 COM 动态调度异常）
                    app.Calculation = -4105; // xlCalculationAutomatic
                    try { app.Calculate(); } catch { }

                    // 激活首选可见工作表
                    ActivateFirstVisibleSheet(reportWb);

                    // 保存填充完成的目标工作簿
                    reportWb.Save();

                    // 汇总统计输出结果指标
                    result.Success = true;
                    result.OutputFilePath = targetPath;
                    result.ExportedCategoryCount = exportCategories.Count;
                    result.ExportedCabinetCount = exportCategories.Sum(c => c.TotalCabinetCount);
                    result.ExportedTotalAmount = exportCategories.Sum(c => c.TotalAmount);
                    result.Message = $"常规样式投标报表导出成功！共导出 {result.ExportedCategoryCount} 个分类，{result.ExportedCabinetCount} 台箱柜。";

                    // 7. 若勾选自动展示且工作簿已存在，切换至前台展示
                    if (config.AutoOpen)
                    {
                        reportWb.Activate();
                    }
                    else
                    {
                        // 若未勾选自动展示，静默关闭已保存的工作簿
                        reportWb.Close(false);
                    }
                }
                finally
                {
                    // 恢复 Excel 全局事件响应机制
                    try { app.EnableEvents = true; } catch { }
                    // 恢复屏幕刷新
                    try { app.ScreenUpdating = true; } catch { }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"报表生成失败: {ex.Message}";
                LogHelper.WriteLog($"[TenderReport] 导出常规投标报表失败: {ex.ToString()}");
            }

            return result;
        }

        /// <summary>
        /// 深度提取选中分类工作表中的所有箱柜及内部元器件明细（采用二维数组一次性读取，严格遵循规则 7 / 规则 12）
        /// </summary>
        private static List<TenderReportCategoryGroup> CollectFullCategoryDetails(dynamic activeWb, List<string> selectedCategoryNames)
        {
            var result = new List<TenderReportCategoryGroup>();
            var targetNameSet = new HashSet<string>(selectedCategoryNames ?? new List<string>(), StringComparer.OrdinalIgnoreCase);

            foreach (dynamic sheet in activeWb.Worksheets)
            {
                string sName = Convert.ToString(sheet.Name) ?? "";
                // 排除系统保留表与报表表，严禁作为分类提取
                if (Tool.IsReservedOrReportSheet(sName))
                {
                    continue;
                }
                if (targetNameSet.Count > 0 && !targetNameSet.Contains(sName))
                {
                    continue;
                }

                // 规则 8: 在操作 Excel 提取前执行 FixAndFillCabinetNamesForSheet 确保规则 6 定义名称与计费起止行正确
                Tool.FixAndFillCabinetNamesForSheet(sheet);

                // 获取有效箱柜列表 (规则 6 / 规则 11)
                var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);
                if (validCabinets.Count == 0) continue;

                var categoryGroup = new TenderReportCategoryGroup
                {
                    CategoryName = sName,
                    IsSelected = true
                };

                for (int i = 0; i < validCabinets.Count; i++)
                {
                    var cabAnchor = validCabinets[i].Value;
                    var cabinetItem = new TenderReportCabinetItem { Index = i + 1 };

                    // 1. 提取箱柜基础汇总属性 (柜号、名称、型号、台数、单价、合价、备注)
                    if (cabAnchor.Sum != null)
                    {
                        int sumRow = Convert.ToInt32(cabAnchor.Sum.Row);
                        // 一次性读取汇总行 1 到 10 列数组 (规则 12)
                        dynamic sumRange = sheet.Range[sheet.Cells[sumRow, 1], sheet.Cells[sumRow, 10]];
                        object[,] sumMatrix = (object[,])sumRange.Value2;

                        cabinetItem.CabinetNo = Convert.ToString(sumMatrix[1, 2]) ?? "";
                        cabinetItem.CabinetName = Convert.ToString(sumMatrix[1, 3]) ?? "";
                        cabinetItem.CabinetModel = Convert.ToString(sumMatrix[1, 4]) ?? "";

                        object qVal = sumMatrix[1, 6] ?? sumMatrix[1, 5];
                        if (qVal != null && int.TryParse(Convert.ToString(qVal), out int q) && q > 0)
                        {
                            cabinetItem.Quantity = q;
                        }

                        object pVal = sumMatrix[1, 7];
                        if (pVal != null && decimal.TryParse(Convert.ToString(pVal), out decimal p))
                        {
                            cabinetItem.UnitPrice = p;
                        }

                        cabinetItem.TotalPrice = cabinetItem.UnitPrice * cabinetItem.Quantity;
                        cabinetItem.Remark = Convert.ToString(sumMatrix[1, 9]) ?? "";
                    }

                    // 2. 提取明细箱柜信息行属性 (尺寸、图号)
                    if (cabAnchor.Det != null)
                    {
                        int detRow = Convert.ToInt32(cabAnchor.Det.Row);
                        try
                        {
                            // 读取明细行尺寸与图号
                            cabinetItem.Size = Convert.ToString(sheet.Cells[detRow, 5].Value2) ?? "";
                            cabinetItem.CadDrawingNo = Convert.ToString(sheet.Cells[detRow, 10].Value2) ?? "";
                        }
                        catch { }

                        // 3. 提取元器件明细列表 (起点 detRow+2，终点 subsumRow-1，规则 6 / 规则 11)
                        if (cabAnchor.Subsum != null)
                        {
                            int subsumRow = Convert.ToInt32(cabAnchor.Subsum.Row);
                            int startCompRow = detRow + 2;
                            int endCompRow = subsumRow - 1;

                            if (endCompRow >= startCompRow)
                            {
                                int compRowCount = endCompRow - startCompRow + 1;
                                // 一次性将元器件整块区域 A 列至 I 列读入内存二维数组 (规则 7 / 规则 12)
                                dynamic compRange = sheet.Range[sheet.Cells[startCompRow, 1], sheet.Cells[endCompRow, 9]];
                                object[,] compMatrix = (object[,])compRange.Value2;

                                int compSeq = 1;
                                for (int r = 1; r <= compRowCount; r++)
                                {
                                    string compName = Convert.ToString(compMatrix[r, 2]) ?? "";
                                    string compModel = Convert.ToString(compMatrix[r, 3]) ?? "";

                                    // 过滤空行与空白无效行
                                    if (string.IsNullOrWhiteSpace(compName) && string.IsNullOrWhiteSpace(compModel))
                                    {
                                        continue;
                                    }

                                    string compManufacturer = Convert.ToString(compMatrix[r, 4]) ?? "";
                                    string compUnit = Convert.ToString(compMatrix[r, 5]) ?? "台";

                                    decimal compQty = 1;
                                    object cQVal = compMatrix[r, 6];
                                    if (cQVal != null && decimal.TryParse(Convert.ToString(cQVal), out decimal cq))
                                    {
                                        compQty = cq;
                                    }

                                    decimal compPrice = 0;
                                    object cPVal = compMatrix[r, 7];
                                    if (cPVal != null && decimal.TryParse(Convert.ToString(cPVal), out decimal cp))
                                    {
                                        compPrice = cp;
                                    }

                                    decimal compTotal = compQty * compPrice;
                                    object cTVal = compMatrix[r, 8];
                                    if (cTVal != null && decimal.TryParse(Convert.ToString(cTVal), out decimal ct) && ct > 0)
                                    {
                                        compTotal = ct;
                                    }

                                    string compRemark = Convert.ToString(compMatrix[r, 9]) ?? "";

                                    cabinetItem.Components.Add(new TenderReportComponentItem
                                    {
                                        Index = compSeq++,
                                        Name = compName,
                                        Model = compModel,
                                        Manufacturer = compManufacturer,
                                        Unit = string.IsNullOrWhiteSpace(compUnit) ? "台" : compUnit,
                                        Quantity = compQty,
                                        UnitPrice = compPrice,
                                        TotalPrice = compTotal,
                                        Remark = compRemark
                                    });
                                }
                            }

                            // 4. 提取计费区域费用项清单 (起点 subsumRow，终点 tolsumRow-1，规则 6 / 规则 7)
                            int tolsumRow = cabAnchor.Tolsum != null ? Convert.ToInt32(cabAnchor.Tolsum.Row) : 0;
                            // 若 Tolsum 未直接锁定，容错探测向下寻找总计行
                            if (tolsumRow == 0)
                            {
                                int maxScan = Math.Min(subsumRow + 20, (int)sheet.UsedRange.Rows.Count);
                                for (int r = subsumRow + 1; r <= maxScan; r++)
                                {
                                    string aText = Convert.ToString(sheet.Cells[r, 1].Value2) ?? "";
                                    string bText = Convert.ToString(sheet.Cells[r, 2].Value2) ?? "";
                                    if (aText.Contains("总计") || bText.Contains("总计") || aText.StartsWith("Cab_Det_"))
                                    {
                                        tolsumRow = r;
                                        break;
                                    }
                                }
                            }

                            // 计算计费区域起止行号 (规则 6: subsumRow 至 tolsumRow - 1)
                            int startFeeRow = subsumRow;
                            int endFeeRow = tolsumRow > 0 ? tolsumRow - 1 : subsumRow;

                            if (endFeeRow >= startFeeRow)
                            {
                                int feeRowCount = endFeeRow - startFeeRow + 1;
                                // 一次性读取计费区域 A 列至 I 列二维数组 (规则 7 / 规则 12)
                                dynamic feeRange = sheet.Range[sheet.Cells[startFeeRow, 1], sheet.Cells[endFeeRow, 9]];
                                object[,] feeMatrix = (object[,])feeRange.Value2;

                                int feeSeq = 1;
                                for (int r = 1; r <= feeRowCount; r++)
                                {
                                    string feeName = Convert.ToString(feeMatrix[r, 2]) ?? "";
                                    // 规则 6: 计费区域不能有空行，跳过名称为空的无效行
                                    if (string.IsNullOrWhiteSpace(feeName)) continue;

                                    string feeModel = Convert.ToString(feeMatrix[r, 3]) ?? "";
                                    string feeMfr = Convert.ToString(feeMatrix[r, 4]) ?? "";
                                    string feeUnit = Convert.ToString(feeMatrix[r, 5]) ?? "";

                                    decimal feeQty = 0;
                                    object fqVal = feeMatrix[r, 6];
                                    if (fqVal != null && decimal.TryParse(Convert.ToString(fqVal), out decimal fq) && fq > 0)
                                    {
                                        feeQty = fq;
                                    }

                                    decimal feePrice = 0;
                                    object fpVal = feeMatrix[r, 7];
                                    if (fpVal != null && decimal.TryParse(Convert.ToString(fpVal), out decimal fp) && fp > 0)
                                    {
                                        feePrice = fp;
                                    }

                                    decimal feeTotal = 0;
                                    object ftVal = feeMatrix[r, 8];
                                    if (ftVal != null && decimal.TryParse(Convert.ToString(ftVal), out decimal ft))
                                    {
                                        feeTotal = ft;
                                    }
                                    else if (feeQty > 0 && feePrice > 0)
                                    {
                                        feeTotal = feeQty * feePrice;
                                    }

                                    string feeRemark = Convert.ToString(feeMatrix[r, 9]) ?? "";

                                    // 加入计费项集合
                                    cabinetItem.FeeItems.Add(new TenderReportComponentItem
                                    {
                                        Index = feeSeq++,
                                        Name = feeName,
                                        Model = feeModel,
                                        Manufacturer = feeMfr,
                                        Unit = feeUnit,
                                        Quantity = feeQty,
                                        UnitPrice = feePrice,
                                        TotalPrice = feeTotal,
                                        Remark = feeRemark
                                    });
                                }
                            }
                        }
                    }

                    categoryGroup.Cabinets.Add(cabinetItem);
                }

                categoryGroup.TotalCabinetCount = categoryGroup.Cabinets.Sum(c => c.Quantity);
                categoryGroup.TotalAmount = categoryGroup.Cabinets.Sum(c => c.TotalPrice);
                result.Add(categoryGroup);
            }

            return result;
        }

        /// <summary>
        /// 填充《封面》工作表占位符
        /// </summary>
        private static void PopulateCoverWorksheet(dynamic reportWb, TenderReportProjectInfo proj, bool includeCover)
        {
            dynamic? coverSheet = null;
            try { coverSheet = reportWb.Sheets["封面"]; } catch { }
            if (coverSheet == null) return;

            // 若用户未勾选导出封面，直接删除封面工作表
            if (!includeCover)
            {
                try { coverSheet.Delete(); } catch { }
                return;
            }

            try
            {
                // 一次性读取 A1:K45 区域公式与内容数组 (规则 12)
                dynamic coverRange = coverSheet.Range["A1:K45"];
                object[,] matrix = (object[,])coverRange.Value2;

                int rowCount = matrix.GetLength(0);
                int colCount = matrix.GetLength(1);

                // 内存遍历并精准替换占位符
                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        string val = Convert.ToString(matrix[r, c]) ?? "";
                        if (string.IsNullOrWhiteSpace(val)) continue;

                        bool changed = false;
                        if (val.Contains("[项目名称]")) { val = val.Replace("[项目名称]", proj.ProjectName); changed = true; }
                        if (val.Contains("[报价日期]")) { val = val.Replace("[报价日期]", proj.QuotationDate); changed = true; }
                        if (val.Contains("[报价单号]")) { val = val.Replace("[报价单号]", proj.QuotationNumber); changed = true; }
                        if (val.Contains("[本单位名称]")) { val = val.Replace("[本单位名称]", proj.FactoryName); changed = true; }
                        if (val.Contains("[客户单位]")) { val = val.Replace("[客户单位]", proj.CustomerName); changed = true; }
                        if (val.Contains("[本单位联系人]")) { val = val.Replace("[本单位联系人]", proj.FactoryContact); changed = true; }
                        if (val.Contains("[本单位电话]")) { val = val.Replace("[本单位电话]", proj.FactoryPhone); changed = true; }
                        if (val.Contains("[报价人]")) { val = val.Replace("[报价人]", proj.Bidder); changed = true; }

                        if (changed)
                        {
                            matrix[r, c] = val;
                        }
                    }
                }

                // 一次性将替换后的结果回写至 Excel 区域 (规则 12)
                coverRange.Value2 = matrix;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[TenderReport] 填充封面工作表异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充《屏柜汇总表》工作表，并返回箱柜在汇总表中的物理行号映射字典
        /// </summary>
        /// <param name="reportWb">目标报表 Excel 工作簿</param>
        /// <param name="proj">项目基础信息</param>
        /// <param name="categories">分类与箱柜数据列表</param>
        /// <param name="includeSummary">是否输出汇总表</param>
        /// <param name="settings">用户自由配置偏好设置</param>
        /// <returns>箱柜全局索引到汇总表物理行号的映射字典</returns>
        private static Dictionary<int, int> PopulateSummaryWorksheet(
            dynamic reportWb,
            TenderReportProjectInfo proj,
            List<TenderReportCategoryGroup> categories,
            bool includeSummary,
            TenderReportSettings settings)
        {
            // 初始化箱柜全局索引到汇总表真实行号的映射字典
            var cabSumRowMap = new Dictionary<int, int>();

            dynamic? sumSheet = null;
            try { sumSheet = reportWb.Sheets["屏柜汇总表"]; } catch { }
            if (sumSheet == null) return cabSumRowMap;

            // 若用户未勾选导出汇总表，直接安全删除该工作表
            if (!includeSummary)
            {
                try { sumSheet.Delete(); } catch { }
                return cabSumRowMap;
            }

            try
            {
                // 1. 替换表头工程与供需信息 (Row 1 至 Row 10)
                dynamic headRange = sumSheet.Range["A1:I10"];
                object[,] headMatrix = (object[,])headRange.Value2;

                for (int r = 1; r <= 10; r++)
                {
                    for (int c = 1; c <= 9; c++)
                    {
                        string val = Convert.ToString(headMatrix[r, c]) ?? "";
                        if (string.IsNullOrWhiteSpace(val)) continue;

                        bool changed = false;
                        if (val.Contains("[项目名称]")) { val = val.Replace("[项目名称]", proj.ProjectName); changed = true; }
                        if (val.Contains("[报价日期]")) { val = val.Replace("[报价日期]", proj.QuotationDate); changed = true; }
                        if (val.Contains("[报价单号]")) { val = val.Replace("[报价单号]", proj.QuotationNumber); changed = true; }
                        if (val.Contains("[本单位名称]")) { val = val.Replace("[本单位名称]", proj.FactoryName); changed = true; }
                        if (val.Contains("[客户单位]")) { val = val.Replace("[客户单位]", proj.CustomerName); changed = true; }
                        if (val.Contains("[本单位联系人]")) { val = val.Replace("[本单位联系人]", proj.FactoryContact); changed = true; }
                        if (val.Contains("[本单位电话]")) { val = val.Replace("[本单位电话]", proj.FactoryPhone); changed = true; }

                        if (changed) { headMatrix[r, c] = val; }
                    }
                }
                headRange.Value2 = headMatrix;

                // 2. 创建临时工作表，暂存母版行（格式、字体、背景色与单元格合并）
                dynamic app = reportWb.Application;
                // 新建临时母版工作表用于存放纯净母版格式
                dynamic tempWs = reportWb.Worksheets.Add();
                // 确保临时工作表删除时不被警告弹窗打扰 --硬编码--
                try { app.DisplayAlerts = false; } catch { }

                try
                {
                    // 模板原结构：
                    // Row 11: 分类标题行 [分类名称] (绿色底色)
                    sumSheet.Rows[11].Copy(tempWs.Rows[1]);
                    // Row 12: 分类表头行 (灰底带边框)
                    sumSheet.Rows[12].Copy(tempWs.Rows[2]);
                    // Row 13: 箱柜数据模板行 (白底细网格边框，包含各字段占位符)
                    sumSheet.Rows[13].Copy(tempWs.Rows[3]);
                    // Row 14: 分类合计行 (浅绿底色，包含 [分类台数]、[分类总价])
                    sumSheet.Rows[14].Copy(tempWs.Rows[4]);
                    // Row 15: 项目总计行 (深绿底色，包含 [项目总价]、[项目台数]、[项目总价大写])
                    sumSheet.Rows[15].Copy(tempWs.Rows[5]);
                    // Row 16: 报价说明标签
                    sumSheet.Rows[16].Copy(tempWs.Rows[6]);
                    // Row 17: 报价说明内容 ([报价说明])
                    sumSheet.Rows[17].Copy(tempWs.Rows[7]);
                    // Row 18: 签名落款行 ([报价人]、[报价日期])
                    sumSheet.Rows[18].Copy(tempWs.Rows[8]);

                    // 清空 Row 11 及以后的全部旧行内容与格式，避免原有文字残留错位
                    sumSheet.Range["A11:Z1000"].Clear();

                    // 【核心机制：占位符动态列映射】：扫描母版数据行，自动识别字段对应物理列号
                    dynamic tplRowRange = tempWs.Range["A3:Z3"];
                    object[,] tplRowMatrix = (object[,])tplRowRange.Value2;
                    int maxCol = Math.Max(9, tplRowMatrix.GetLength(1));

                    // 默认兜底列号 (1-based)
                    int colCabSeq = 1;     // [箱柜序号]
                    int colCabNo = 2;      // [箱柜柜号] / [柜号]
                    int colCabName = 3;    // [箱柜名称]
                    int colCabModel = 4;   // [箱柜型号]
                    int colCabUnit = 5;    // [单位]
                    int colCabQty = 6;     // [数量]
                    int colCabPrice = 7;   // [箱柜单价] / [单价]
                    int colCabTotal = 8;   // [箱柜总价] / [总价]
                    int colCabRemark = 9;  // [备注]

                    // 遍历母版数据行，依据单元格中的占位符标签自适应解析列号
                    for (int c = 1; c <= maxCol; c++)
                    {
                        string tag = Convert.ToString(tplRowMatrix[1, c]) ?? "";
                        if (string.IsNullOrWhiteSpace(tag)) continue;
                        if (tag.Contains("[箱柜序号]") || tag.Contains("[序号]")) colCabSeq = c;
                        else if (tag.Contains("[箱柜柜号]") || tag.Contains("[柜号]")) colCabNo = c;
                        else if (tag.Contains("[箱柜名称]")) colCabName = c;
                        else if (tag.Contains("[箱柜型号]")) colCabModel = c;
                        else if (tag.Contains("[单位]")) colCabUnit = c;
                        else if (tag.Contains("[数量]")) colCabQty = c;
                        else if (tag.Contains("[箱柜单价]") || tag.Contains("[单价]")) colCabPrice = c;
                        else if (tag.Contains("[箱柜总价]") || tag.Contains("[总价]")) colCabTotal = c;
                        else if (tag.Contains("[备注]")) colCabRemark = c;
                    }

                    // 动态计算数量与单价所在的 Excel 列字母 (例如 F 列、G 列、H 列)
                    string qtyColLetter = GetExcelColumnLetter(colCabQty);
                    string priceColLetter = GetExcelColumnLetter(colCabPrice);
                    string totalColLetter = GetExcelColumnLetter(colCabTotal);

                    // 记录所有生成的小计行行号，用于项目总计求和
                    var subtotalRows = new List<int>();
                    // 当前行游标，从第 11 行开始自顶向下单向克隆
                    int currentRow = 11;
                    // 全局箱柜序号计数器
                    int globalIndex = 1;

                    // 循环填充各个分类块
                    for (int catIdx = 0; catIdx < categories.Count; catIdx++)
                    {
                        var cat = categories[catIdx];

                        // ① 克隆分类标题行 (浅绿底色)
                        tempWs.Rows[1].Copy(sumSheet.Rows[currentRow]);
                        // 自适应替换单元格中的 [分类名称] 占位符
                        ReplaceRowPlaceholders(sumSheet.Range[$"A{currentRow}:Z{currentRow}"], new Dictionary<string, string>
                        {
                            { "[分类名称]", $"分类：{cat.CategoryName}" }
                        }, fallbackCol: 1, fallbackValue: $"分类：{cat.CategoryName}");
                        currentRow++;

                        // ② 克隆表头行 (灰底)
                        tempWs.Rows[2].Copy(sumSheet.Rows[currentRow]);
                        currentRow++;

                        // 记录当前分类数据起始行号
                        int catDataStartRow = currentRow;
                        int catItemCount = cat.Cabinets.Count;

                        if (catItemCount > 0)
                        {
                            // 批量复制母版数据行格式至目标多行区域 (继承白底与细网格线)
                            dynamic targetDataRange = sumSheet.Range[sumSheet.Rows[currentRow], sumSheet.Rows[currentRow + catItemCount - 1]];
                            tempWs.Rows[3].Copy();
                            // 复制母版行的单元格样式、边框与对齐方式 --硬编码--
                            targetDataRange.PasteSpecial(-4104);
                            try { app.CutCopyMode = false; } catch { }

                            // 准备二维数组批量灌入当前分类的全部箱柜行 (规则 7 / 规则 12)
                            object[,] catMatrix = new object[catItemCount, maxCol];
                            for (int k = 0; k < catItemCount; k++)
                            {
                                var cab = cat.Cabinets[k];
                                int realRow = currentRow + k;

                                // 记录当前箱柜全局序号到汇总表物理行的映射，供分项表双向超链接跳转
                                cabSumRowMap[globalIndex] = realRow;

                                // 智能清洗前缀，杜绝重复叠加
                                string cleanCabNo = System.Text.RegularExpressions.Regex.Replace(cab.CabinetNo ?? "", @"^柜号[:：]\s*", "");
                                string cleanModel = System.Text.RegularExpressions.Regex.Replace(cab.CabinetModel ?? "", @"^型号[:：]\s*", "");
                                string cleanName = System.Text.RegularExpressions.Regex.Replace(cab.CabinetName ?? "", @"^名称[:：]\s*", "");
                                string cleanRemark = System.Text.RegularExpressions.Regex.Replace(cab.Remark ?? "", @"^备注[:：]\s*", "");

                                // 处理箱柜单价与总价取整到元配置
                                decimal cabPrice = settings.RoundCabinetUnitPrice ? Math.Round(cab.UnitPrice, 0) : cab.UnitPrice;

                                // 严格根据自适应扫描到的列号灌入数据 (0-based)
                                if (colCabSeq > 0) catMatrix[k, colCabSeq - 1] = globalIndex;
                                if (colCabNo > 0) catMatrix[k, colCabNo - 1] = cleanCabNo;
                                if (colCabName > 0) catMatrix[k, colCabName - 1] = cleanName;
                                if (colCabModel > 0) catMatrix[k, colCabModel - 1] = cleanModel;
                                if (colCabUnit > 0) catMatrix[k, colCabUnit - 1] = string.IsNullOrWhiteSpace(cab.Unit) ? "台" : cab.Unit;
                                if (colCabQty > 0) catMatrix[k, colCabQty - 1] = cab.Quantity;
                                if (colCabPrice > 0) catMatrix[k, colCabPrice - 1] = cabPrice;

                                // 总价公式或纯数值判定
                                if (colCabTotal > 0)
                                {
                                    if (settings.TotalWithFormula)
                                    {
                                        // 若带公式且勾选取整到元，使用 ROUND(..., 0)
                                        int totalDecimals = settings.RoundCabinetTotal ? 0 : 2;
                                        catMatrix[k, colCabTotal - 1] = $"=ROUND({qtyColLetter}{realRow}*{priceColLetter}{realRow}, {totalDecimals})";
                                    }
                                    else
                                    {
                                        // 若不带公式，直接写入计算好的纯数值
                                        decimal totalVal = cabPrice * cab.Quantity;
                                        catMatrix[k, colCabTotal - 1] = settings.RoundCabinetTotal ? Math.Round(totalVal, 0) : Math.Round(totalVal, 2);
                                    }
                                }
                                if (colCabRemark > 0) catMatrix[k, colCabRemark - 1] = cleanRemark;

                                globalIndex++;
                            }

                            // 批量写入数据区域 (规则 12)
                            dynamic dataRange = sumSheet.Range[sumSheet.Cells[currentRow, 1], sumSheet.Cells[currentRow + catItemCount - 1, maxCol]];
                            dataRange.Value2 = catMatrix;

                            // 若开启文字自动换行
                            if (settings.TextAutoWrap)
                            {
                                try { dataRange.WrapText = true; } catch { }
                            }

                            // 若开启一键定位：为汇总表序号列绑定跳转到分项表对应箱柜定义名称的超链接
                            if (settings.EnableOneKeyLocate)
                            {
                                int tempStartGlobal = globalIndex - catItemCount;
                                for (int k = 0; k < catItemCount; k++)
                                {
                                    int curGlobal = tempStartGlobal + k;
                                    int curRow = currentRow + k;
                                    try
                                    {
                                        // 汇总表跳转至分项表对应箱柜定义名称 CabDetRef_{curGlobal}
                                        sumSheet.Hyperlinks.Add(
                                            Anchor: sumSheet.Cells[curRow, colCabSeq],
                                            Address: "",
                                            SubAddress: $"CabDetRef_{curGlobal}",
                                            ScreenTip: "点击跳转至该柜分项明细",
                                            TextToDisplay: curGlobal.ToString());
                                    }
                                    catch { }
                                }
                            }

                            currentRow += catItemCount;
                        }

                        // 记录当前分类数据终止行号
                        int catDataEndRow = currentRow - 1;

                        // ③ 克隆分类合计行 (浅绿底色)
                        tempWs.Rows[4].Copy(sumSheet.Rows[currentRow]);
                        dynamic subtotalRange = sumSheet.Range[$"A{currentRow}:Z{currentRow}"];
                        object[,] subtotalMatrix = (object[,])subtotalRange.Value2;

                        string catQtyFormula = catDataEndRow >= catDataStartRow ? $"=SUM({qtyColLetter}{catDataStartRow}:{qtyColLetter}{catDataEndRow})" : "0";
                        string catTotalFormula = catDataEndRow >= catDataStartRow
                            ? (settings.RoundCabinetTotal
                                ? $"=ROUND(SUM({totalColLetter}{catDataStartRow}:{totalColLetter}{catDataEndRow}), 0)"
                                : $"=SUM({totalColLetter}{catDataStartRow}:{totalColLetter}{catDataEndRow})")
                            : "0";

                        // 遍历合计行单元格，依据占位符自适应注入公式或纯数值
                        for (int sc = 1; sc <= maxCol; sc++)
                        {
                            string sVal = Convert.ToString(subtotalMatrix[1, sc]) ?? "";
                            if (sVal.Contains("[分类台数]") || sVal.Contains("[台数]") || (sc == colCabQty && sVal.Contains("[")))
                            {
                                subtotalRange.Cells[1, sc].Formula = catQtyFormula;
                            }
                            else if (sVal.Contains("[分类总价]") || sVal.Contains("[总价]") || (sc == colCabTotal && sVal.Contains("[")))
                            {
                                if (settings.TotalWithFormula)
                                {
                                    subtotalRange.Cells[1, sc].Formula = catTotalFormula;
                                }
                                else
                                {
                                    // 纯数值输出分类合计金额
                                    decimal catTotalVal = cat.TotalAmount;
                                    subtotalRange.Cells[1, sc].Value2 = settings.RoundCabinetTotal ? Math.Round(catTotalVal, 0) : Math.Round(catTotalVal, 2);
                                }
                            }
                        }

                        // 记录该小计行行号，用于汇总最终总计
                        subtotalRows.Add(currentRow);
                        currentRow++;
                    }

                    // 3. 写入项目总计行 (克隆深绿底色母版)
                    int grandTotalRow = currentRow;
                    tempWs.Rows[5].Copy(sumSheet.Rows[grandTotalRow]);
                    dynamic grandRange = sumSheet.Range[$"A{grandTotalRow}:Z{grandTotalRow}"];
                    object[,] grandMatrix = (object[,])grandRange.Value2;

                    string qtySumFormula = subtotalRows.Count > 0 ? "=" + string.Join("+", subtotalRows.Select(r => $"{qtyColLetter}{r}")) : "0";
                    string amtSumFormula = subtotalRows.Count > 0 ? "=" + string.Join("+", subtotalRows.Select(r => $"{totalColLetter}{r}")) : "0";
                    decimal grandTotalAmount = categories.Sum(c => c.TotalAmount);
                    // 人民币大写金额转换 (若取整到元则先截断四舍五入)
                    decimal dispGrandAmount = settings.RoundProjectTotal ? Math.Round(grandTotalAmount, 0) : grandTotalAmount;
                    string upperAmount = ConvertAmountToChineseUpper(dispGrandAmount);

                    // 遍历总计行单元格，依据占位符自适应填充
                    for (int gc = 1; gc <= maxCol; gc++)
                    {
                        string gVal = Convert.ToString(grandMatrix[1, gc]) ?? "";
                        if (gVal.Contains("[项目台数]") || gVal.Contains("[总台数]"))
                        {
                            grandRange.Cells[1, gc].Formula = qtySumFormula;
                        }
                        else if (gVal.Contains("[项目总价大写]") || gVal.Contains("[大写]"))
                        {
                            grandRange.Cells[1, gc].Value2 = upperAmount;
                        }
                        else if (gVal.Contains("[项目总价]") || gVal.Contains("[总金额]"))
                        {
                            if (settings.TotalWithFormula)
                            {
                                grandRange.Cells[1, gc].Formula = amtSumFormula;
                            }
                            else
                            {
                                // 纯数值输出项目总价
                                grandRange.Cells[1, gc].Value2 = dispGrandAmount;
                            }
                        }
                    }
                    currentRow++;

                    // 4. 写入报价说明与签名尾栏 (受 settings.OutputNotesInSummary 控制)
                    if (settings.OutputNotesInSummary)
                    {
                        // 复制报价说明标签行
                        tempWs.Rows[6].Copy(sumSheet.Rows[currentRow]);
                        currentRow++;

                        // 复制报价说明内容行
                        tempWs.Rows[7].Copy(sumSheet.Rows[currentRow]);
                        ReplaceRowPlaceholders(sumSheet.Range[$"A{currentRow}:Z{currentRow}"], new Dictionary<string, string>
                        {
                            { "[报价说明]", proj.ProjectRemark }
                        }, fallbackCol: 2, fallbackValue: proj.ProjectRemark);

                        // 若开启报价说明自动行高
                        if (settings.NotesAutoFitRowHeight)
                        {
                            try { sumSheet.Rows[currentRow].AutoFit(); } catch { }
                        }
                        currentRow++;

                        // 空一行留出视觉间距
                        currentRow++;
                    }

                    // 复制签名落款行
                    tempWs.Rows[8].Copy(sumSheet.Rows[currentRow]);
                    ReplaceRowPlaceholders(sumSheet.Range[$"A{currentRow}:Z{currentRow}"], new Dictionary<string, string>
                    {
                        { "[报价人]", proj.Bidder },
                        { "[报价日期]", proj.QuotationDate }
                    });
                }
                finally
                {
                    // 安全移除临时母版工作表，防止残留工作簿
                    try
                    {
                        tempWs.Delete();
                        app.DisplayAlerts = true;
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[TenderReport] 填充屏柜汇总表异常: {ex.Message}");
            }

            // 返回全局箱柜行号映射
            return cabSumRowMap;
        }

        /// <summary>
        /// 调度生成《屏柜分项表》工作表 (支持按“分类工作表”分别输出独立 Sheet 与单个分项表两种模式)
        /// </summary>
        /// <param name="reportWb">目标报表 Excel 工作簿</param>
        /// <param name="proj">项目基础信息</param>
        /// <param name="categories">分类与箱柜元器件数据列表</param>
        /// <param name="includeDetail">是否输出分项明细表</param>
        /// <param name="settings">用户配置偏好设置</param>
        /// <param name="cabSumRowMap">箱柜在汇总表中的物理行号映射字典</param>
        private static void PopulateDetailWorksheets(
            dynamic reportWb,
            TenderReportProjectInfo proj,
            List<TenderReportCategoryGroup> categories,
            bool includeDetail,
            TenderReportSettings settings,
            Dictionary<int, int> cabSumRowMap)
        {
            // 若用户未勾选导出分项明细表，安全删除模板表后返回
            if (!includeDetail)
            {
                try { reportWb.Sheets["屏柜分项表"].Delete(); } catch { }
                return;
            }

            dynamic? templateDetSheet = null;
            try { templateDetSheet = reportWb.Sheets["屏柜分项表"]; } catch { }
            if (templateDetSheet == null) return;

            dynamic app = reportWb.Application;
            int runningGlobalCabIndex = 1;

            // 核心判定：是否启用按“分类工作表”分别输出独立 Sheet
            bool splitBySheet = settings.SplitDetailBySheet && categories.Count > 0;

            if (splitBySheet)
            {
                // 筛选包含有效箱柜的分类
                var validCats = categories.Where(c => c.Cabinets != null && c.Cabinets.Count > 0).ToList();
                if (validCats.Count == 0)
                {
                    validCats = categories;
                }

                // 遍历每个分类，独立克隆一份分项明细 Sheet
                for (int i = 0; i < validCats.Count; i++)
                {
                    var cat = validCats[i];
                    // 从模板克隆工作表至末尾
                    templateDetSheet.Copy(After: reportWb.Sheets[reportWb.Sheets.Count]);
                    dynamic catSheet = reportWb.Sheets[reportWb.Sheets.Count];

                    // 规范化 Sheet 名称，剔除 Excel 不支持的特殊字符 \ / ? * : [ ]
                    string safeCatName = cat.CategoryName ?? "分项表";
                    foreach (char invalidChar in new char[] { '\\', '/', '?', '*', ':', '[', ']' })
                    {
                        safeCatName = safeCatName.Replace(invalidChar, '_');
                    }
                    if (safeCatName.Length > 25) safeCatName = safeCatName.Substring(0, 25);
                    string finalSheetName = safeCatName;
                    int dupIndex = 1;
                    // 防止工作表重名报错
                    while (SheetExists(reportWb, finalSheetName))
                    {
                        finalSheetName = $"{safeCatName}_{dupIndex++}";
                    }
                    catSheet.Name = finalSheetName;

                    // 填充该独立分类工作表
                    PopulateSingleDetailSheet(reportWb, catSheet, proj, new List<TenderReportCategoryGroup> { cat }, settings, cabSumRowMap, ref runningGlobalCabIndex);
                }

                // 所有独立 Sheet 生成完毕后，安全移除原模板工作表
                try
                {
                    app.DisplayAlerts = false;
                    templateDetSheet.Delete();
                    app.DisplayAlerts = true;
                }
                catch { }
            }
            else
            {
                // 单表模式：所有分类输出在同一个《屏柜分项表》中
                PopulateSingleDetailSheet(reportWb, templateDetSheet, proj, categories, settings, cabSumRowMap, ref runningGlobalCabIndex);
            }
        }

        /// <summary>
        /// 填充单张分项明细工作表的核心数据生成引擎
        /// </summary>
        private static void PopulateSingleDetailSheet(
            dynamic reportWb,
            dynamic detSheet,
            TenderReportProjectInfo proj,
            List<TenderReportCategoryGroup> categories,
            TenderReportSettings settings,
            Dictionary<int, int> cabSumRowMap,
            ref int globalCabIndex)
        {
            try
            {
                // 1. 替换表头工程与供需信息 (Row 1 至 Row 10)
                dynamic headRange = detSheet.Range["A1:I10"];
                object[,] headMatrix = (object[,])headRange.Value2;

                for (int r = 1; r <= 10; r++)
                {
                    for (int c = 1; c <= 9; c++)
                    {
                        string val = Convert.ToString(headMatrix[r, c]) ?? "";
                        if (string.IsNullOrWhiteSpace(val)) continue;

                        bool changed = false;
                        if (val.Contains("[项目名称]")) { val = val.Replace("[项目名称]", proj.ProjectName); changed = true; }
                        if (val.Contains("[报价日期]")) { val = val.Replace("[报价日期]", proj.QuotationDate); changed = true; }
                        if (val.Contains("[报价单号]")) { val = val.Replace("[报价单号]", proj.QuotationNumber); changed = true; }
                        if (val.Contains("[本单位名称]")) { val = val.Replace("[本单位名称]", proj.FactoryName); changed = true; }
                        if (val.Contains("[客户单位]")) { val = val.Replace("[客户单位]", proj.CustomerName); changed = true; }
                        if (val.Contains("[本单位联系人]")) { val = val.Replace("[本单位联系人]", proj.FactoryContact); changed = true; }
                        if (val.Contains("[本单位电话]")) { val = val.Replace("[本单位电话]", proj.FactoryPhone); changed = true; }

                        if (changed) { headMatrix[r, c] = val; }
                    }
                }
                headRange.Value2 = headMatrix;

                // 2. 创建临时工作表，暂存母版行 (浅蓝柜头/灰底表头/白底元件行/橙底总计行)
                dynamic app = reportWb.Application;
                // 新建临时母版工作表用于存放明细表纯净母版格式
                dynamic tempWs = reportWb.Worksheets.Add();
                // 确保临时工作表删除时不被警告弹窗打扰 --硬编码--
                try { app.DisplayAlerts = false; } catch { }

                try
                {
                    // 模板原结构：
                    // Row 11: 箱柜信息行 (浅蓝底色与单元格合并，包含 柜号：[柜号]、型号：[箱柜型号] 等)
                    detSheet.Rows[11].Copy(tempWs.Rows[1]);
                    // Row 12: 元件表头行 (灰底带边框)
                    detSheet.Rows[12].Copy(tempWs.Rows[2]);
                    // Row 13: 元件明细行模板 (白底细网格边框，包含各元件占位符)
                    detSheet.Rows[13].Copy(tempWs.Rows[3]);
                    // Row 14: 箱柜总计行 (橙底加粗带边框，包含 [箱柜数量]、[箱柜总价])
                    detSheet.Rows[14].Copy(tempWs.Rows[4]);

                    // 清空 Row 11 及以后的全部旧行内容与格式，防止残留
                    detSheet.Range["A11:Z10000"].Clear();

                    // 【核心机制：占位符动态列映射】：扫描第 3 行母版数据行，自适应识别各列号
                    dynamic tplCompRange = tempWs.Range["A3:Z3"];
                    object[,] tplCompMatrix = (object[,])tplCompRange.Value2;
                    int maxCompCol = Math.Max(9, tplCompMatrix.GetLength(1));

                    // 默认兜底列号 (1-based)
                    int colCompSeq = 1;     // [元件序号] / [序号]
                    int colCompName = 2;    // [元件名称]
                    int colCompModel = 3;   // [元件型号]
                    int colCompUnit = 4;    // [元件单位] / [单位]
                    int colCompQty = 5;     // [元件数量] / [数量]
                    int colCompPrice = 6;   // [元件单价] / [单价]
                    int colCompTotal = 7;   // [元件总价] / [总价]
                    int colCompMfr = 8;     // [生产厂家] / [厂家]
                    int colCompRemark = 9;  // [元件备注] / [备注]

                    // 遍历母版元件行，依据占位符自适应锁定列号
                    for (int c = 1; c <= maxCompCol; c++)
                    {
                        string tag = Convert.ToString(tplCompMatrix[1, c]) ?? "";
                        if (string.IsNullOrWhiteSpace(tag)) continue;
                        if (tag.Contains("[元件序号]") || tag.Contains("[序号]")) colCompSeq = c;
                        else if (tag.Contains("[元件名称]")) colCompName = c;
                        else if (tag.Contains("[元件型号]")) colCompModel = c;
                        else if (tag.Contains("[元件单位]") || tag.Contains("[单位]")) colCompUnit = c;
                        else if (tag.Contains("[元件数量]") || tag.Contains("[数量]")) colCompQty = c;
                        else if (tag.Contains("[元件单价]") || tag.Contains("[单价]")) colCompPrice = c;
                        else if (tag.Contains("[元件总价]") || tag.Contains("[总价]")) colCompTotal = c;
                        else if (tag.Contains("[生产厂家]") || tag.Contains("[厂家]")) colCompMfr = c;
                        else if (tag.Contains("[元件备注]") || tag.Contains("[备注]")) colCompRemark = c;
                    }

                    // 动态计算元件数量与单价、总价的 Excel 列字母
                    string compQtyColLetter = GetExcelColumnLetter(colCompQty);
                    string compPriceColLetter = GetExcelColumnLetter(colCompPrice);
                    string compTotalColLetter = GetExcelColumnLetter(colCompTotal);

                    // 当前行游标，从第 11 行开始自顶向下单向克隆
                    int currentRow = 11;

                    // 循环每个分类中的每个箱柜进行展开
                    for (int cIdx = 0; cIdx < categories.Count; cIdx++)
                    {
                        var cat = categories[cIdx];

                        for (int cabIdx = 0; cabIdx < cat.Cabinets.Count; cabIdx++)
                        {
                            var cab = cat.Cabinets[cabIdx];
                            int thisCabStartRow = currentRow;

                            // ① 克隆箱柜信息栏 (浅蓝底色与单元格合并)
                            tempWs.Rows[1].Copy(detSheet.Rows[currentRow]);

                            // 为该箱柜顶部行在工作簿中定义名称 CabDetRef_{globalCabIndex}，供汇总表精准直达跳转
                            try
                            {
                                reportWb.Names.Add($"CabDetRef_{globalCabIndex}", detSheet.Cells[currentRow, 1]);
                            }
                            catch { }

                            // 智能清洗源数据可能自带的重复前缀
                            string cleanCabNo = System.Text.RegularExpressions.Regex.Replace(cab.CabinetNo ?? "", @"^柜号[:：]\s*", "");
                            string cleanCabModel = System.Text.RegularExpressions.Regex.Replace(cab.CabinetModel ?? "", @"^型号[:：]\s*", "");
                            string cleanCabName = System.Text.RegularExpressions.Regex.Replace(cab.CabinetName ?? "", @"^名称[:：]\s*", "");
                            string cleanCabRemark = System.Text.RegularExpressions.Regex.Replace(cab.Remark ?? "", @"^备注[:：]\s*", "");

                            // 【核心机制：占位符自适应替换】：整行按占位符自然替换
                            ReplaceRowPlaceholders(detSheet.Range[$"A{currentRow}:Z{currentRow}"], new Dictionary<string, string>
                            {
                                { "[箱柜序号]", globalCabIndex.ToString() },
                                { "[柜号]", cleanCabNo },
                                { "[箱柜型号]", cleanCabModel },
                                { "[箱柜名称]", cleanCabName },
                                { "[箱柜备注]", cleanCabRemark }
                            });

                            // 若开启“一键定位”：在箱柜信息栏绑定返回汇总表的超链接
                            if (settings.EnableOneKeyLocate && cabSumRowMap.TryGetValue(globalCabIndex, out int sumRow))
                            {
                                try
                                {
                                    // 给箱柜信息行第 1 个单元格添加反向跳转超链接
                                    detSheet.Hyperlinks.Add(
                                        Anchor: detSheet.Cells[currentRow, 1],
                                        Address: "",
                                        SubAddress: $"'屏柜汇总表'!A{sumRow}",
                                        ScreenTip: "点击返回汇总表该柜所在行",
                                        TextToDisplay: Convert.ToString(detSheet.Cells[currentRow, 1].Value2));
                                }
                                catch { }
                            }

                            globalCabIndex++;
                            currentRow++;

                            // ② 克隆元件表头栏 (灰底)
                            tempWs.Rows[2].Copy(detSheet.Rows[currentRow]);
                            currentRow++;

                            // ③ 批量写入元器件与计费区域明细行 (白底细网格，费用项紧随元器件尾部输出)
                            int compCount = cab.Components.Count;
                            int feeCount = cab.FeeItems.Count;
                            int totalDetailCount = compCount + feeCount;
                            int detailStartRow = currentRow;
                            int singleTotalRow = 0;
                            int subtotalRow = 0;

                            if (totalDetailCount > 0)
                            {
                                // 批量复制母版数据行格式至目标多行区域 (继承白色底与细网格线)
                                dynamic compTargetRange = detSheet.Range[detSheet.Rows[currentRow], detSheet.Rows[currentRow + totalDetailCount - 1]];
                                tempWs.Rows[3].Copy();
                                // 复制母版行的单元格样式与边框 --硬编码--
                                compTargetRange.PasteSpecial(-4104);
                                try { app.CutCopyMode = false; } catch { }

                                // 准备二维数组一次性回写 (规则 7 / 规则 12)
                                object[,] detailMatrix = new object[totalDetailCount, maxCompCol];

                                // 3.1 写入元器件明细行
                                for (int ci = 0; ci < compCount; ci++)
                                {
                                    var item = cab.Components[ci];
                                    int realRow = currentRow + ci;

                                    // 单价取整到元控制
                                    decimal compPrice = settings.RoundDetailUnitPriceTotal ? Math.Round(item.UnitPrice, 0) : item.UnitPrice;

                                    if (colCompSeq > 0) detailMatrix[ci, colCompSeq - 1] = item.Index;
                                    if (colCompName > 0) detailMatrix[ci, colCompName - 1] = item.Name;
                                    if (colCompModel > 0) detailMatrix[ci, colCompModel - 1] = item.Model;
                                    if (colCompUnit > 0) detailMatrix[ci, colCompUnit - 1] = item.Unit;
                                    if (colCompQty > 0) detailMatrix[ci, colCompQty - 1] = item.Quantity;
                                    if (colCompPrice > 0) detailMatrix[ci, colCompPrice - 1] = compPrice;

                                    // 元器件总价公式或纯数值判定
                                    if (colCompTotal > 0)
                                    {
                                        if (settings.TotalWithFormula)
                                        {
                                            // 带公式输出：若勾选取整到元使用 0 位小数
                                            int compDecimals = settings.RoundDetailUnitPriceTotal ? 0 : 2;
                                            detailMatrix[ci, colCompTotal - 1] = $"=ROUND({compQtyColLetter}{realRow}*{compPriceColLetter}{realRow}, {compDecimals})";
                                        }
                                        else
                                        {
                                            // 纯数值输出
                                            decimal compTotalVal = item.Quantity * compPrice;
                                            detailMatrix[ci, colCompTotal - 1] = settings.RoundDetailUnitPriceTotal ? Math.Round(compTotalVal, 0) : Math.Round(compTotalVal, 2);
                                        }
                                    }
                                    if (colCompMfr > 0) detailMatrix[ci, colCompMfr - 1] = item.Manufacturer;
                                    if (colCompRemark > 0) detailMatrix[ci, colCompRemark - 1] = item.Remark;
                                }

                                // 3.2 写入计费区域费用项 (明细尾部输出，费用项不带公式直接存入数值)
                                for (int fi = 0; fi < feeCount; fi++)
                                {
                                    var fee = cab.FeeItems[fi];
                                    int mIdx = compCount + fi;
                                    int realRow = currentRow + mIdx;

                                    // 识别“单台合计”或“小计”物理行号用于总计公式联动
                                    if (fee.Name.Contains("单台合计") || fee.Name.Contains("单台总计"))
                                    {
                                        singleTotalRow = realRow;
                                    }
                                    else if (fee.Name.Contains("小计") && subtotalRow == 0)
                                    {
                                        subtotalRow = realRow;
                                    }

                                    // 费用项取整控制
                                    decimal feeTotalVal = settings.RoundDetailUnitPriceTotal ? Math.Round(fee.TotalPrice, 0) : Math.Round(fee.TotalPrice, 2);

                                    if (colCompSeq > 0) detailMatrix[mIdx, colCompSeq - 1] = compCount + fi + 1;
                                    if (colCompName > 0) detailMatrix[mIdx, colCompName - 1] = fee.Name;
                                    if (colCompModel > 0) detailMatrix[mIdx, colCompModel - 1] = fee.Model;
                                    if (colCompUnit > 0) detailMatrix[mIdx, colCompUnit - 1] = fee.Unit;
                                    if (colCompQty > 0) detailMatrix[mIdx, colCompQty - 1] = fee.Quantity > 0 ? (object)fee.Quantity : "";
                                    if (colCompPrice > 0) detailMatrix[mIdx, colCompPrice - 1] = fee.UnitPrice > 0 ? (object)fee.UnitPrice : "";
                                    // 费用项不带公式直接存入纯数值，杜绝跨表引用报错
                                    if (colCompTotal > 0) detailMatrix[mIdx, colCompTotal - 1] = feeTotalVal;
                                    if (colCompMfr > 0) detailMatrix[mIdx, colCompMfr - 1] = fee.Manufacturer;
                                    if (colCompRemark > 0) detailMatrix[mIdx, colCompRemark - 1] = fee.Remark;
                                }

                                dynamic detailRange = detSheet.Range[detSheet.Cells[currentRow, 1], detSheet.Cells[currentRow + totalDetailCount - 1, maxCompCol]];
                                // 使用 Formula 属性兼顾公式解析与数值存入
                                detailRange.Formula = detailMatrix;

                                // 文字自动换行支持
                                if (settings.TextAutoWrap)
                                {
                                    try { detailRange.WrapText = true; } catch { }
                                }

                                currentRow += totalDetailCount;
                            }

                            int detailEndRow = currentRow - 1;

                            // ④ 克隆箱柜总计栏 (橙底加粗带边框)
                            tempWs.Rows[4].Copy(detSheet.Rows[currentRow]);
                            dynamic cabTotalRange = detSheet.Range[$"A{currentRow}:Z{currentRow}"];
                            object[,] cabTotalMatrix = (object[,])cabTotalRange.Value2;

                            // 动态识别总计行中的数量列字母
                            string cabQtyColLetter = compQtyColLetter;
                            for (int tc = 1; tc <= maxCompCol; tc++)
                            {
                                string tVal = Convert.ToString(cabTotalMatrix[1, tc]) ?? "";
                                if (tVal.Contains("[箱柜数量]") || tVal.Contains("[数量]"))
                                {
                                    cabQtyColLetter = GetExcelColumnLetter(tc);
                                    cabTotalRange.Cells[1, tc].Value2 = cab.Quantity;
                                }
                            }

                            // 动态识别总计行中的总价列并写入求和/联动计算公式或数值
                            int cabTotalDecimals = settings.RoundCabinetTotal ? 0 : 2;
                            for (int tc = 1; tc <= maxCompCol; tc++)
                            {
                                string tVal = Convert.ToString(cabTotalMatrix[1, tc]) ?? "";
                                if (tVal.Contains("[箱柜总价]") || tVal.Contains("[总价]"))
                                {
                                    if (settings.TotalWithFormula)
                                    {
                                        if (singleTotalRow > 0)
                                        {
                                            // 模式 1：存在“单台合计”行，总价公式 = 单台合计 * 箱柜数量
                                            cabTotalRange.Cells[1, tc].Formula = $"=ROUND({compTotalColLetter}{singleTotalRow}*{cabQtyColLetter}{currentRow}, {cabTotalDecimals})";
                                        }
                                        else if (subtotalRow > 0 && feeCount > 1)
                                        {
                                            // 模式 2：存在“小计”行与费用项，总价公式 = (小计至最后一项费用之和) * 箱柜数量
                                            cabTotalRange.Cells[1, tc].Formula = $"=ROUND(SUM({compTotalColLetter}{subtotalRow}:{compTotalColLetter}{detailEndRow})*{cabQtyColLetter}{currentRow}, {cabTotalDecimals})";
                                        }
                                        else if (totalDetailCount > 0)
                                        {
                                            // 模式 3：无单台合计/小计，全明细求和 * 箱柜数量
                                            cabTotalRange.Cells[1, tc].Formula = $"=ROUND(SUM({compTotalColLetter}{detailStartRow}:{compTotalColLetter}{detailEndRow})*{cabQtyColLetter}{currentRow}, {cabTotalDecimals})";
                                        }
                                        else
                                        {
                                            // 模式 4：兜底保底，直接填入箱柜总价数值
                                            cabTotalRange.Cells[1, tc].Value2 = settings.RoundCabinetTotal ? Math.Round(cab.TotalPrice, 0) : Math.Round(cab.TotalPrice, 2);
                                        }
                                    }
                                    else
                                    {
                                        // 若不带公式，直接写入计算好的纯数值
                                        cabTotalRange.Cells[1, tc].Value2 = settings.RoundCabinetTotal ? Math.Round(cab.TotalPrice, 0) : Math.Round(cab.TotalPrice, 2);
                                    }
                                }
                            }
                            currentRow++;

                            // ⑤ 打印设置：每台箱柜分页打印 (插入水平分页符)
                            bool isLastCabInSheet = (cIdx == categories.Count - 1) && (cabIdx == cat.Cabinets.Count - 1);
                            if (settings.PrintSetting == "Paginated" && !isLastCabInSheet)
                            {
                                try
                                {
                                    // 在下一个箱柜的起始位置插入水平分页符
                                    detSheet.HPageBreaks.Add(detSheet.Cells[currentRow + 1, 1]);
                                }
                                catch { }
                            }

                            // 箱柜块之间空一行保持美观间距
                            currentRow++;
                        }
                    }
                }
                finally
                {
                    // 安全移除临时母版工作表，防止残留工作簿
                    try
                    {
                        tempWs.Delete();
                        app.DisplayAlerts = true;
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[TenderReport] 填充屏柜分项表异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查工作簿中是否存在指定名称的工作表
        /// </summary>
        /// <param name="wb">Excel 工作簿实例</param>
        /// <param name="sheetName">待检测的工作表名称</param>
        /// <returns>存在返回 true，否则 false</returns>
        private static bool SheetExists(dynamic wb, string sheetName)
        {
            try
            {
                foreach (dynamic ws in wb.Worksheets)
                {
                    if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 将 1-based 数字列号转换为标准 Excel 列字母 (例如 1->A, 2->B, 26->Z, 27->AA)
        /// </summary>
        /// <param name="colIndex">1-based 数字列号</param>
        /// <returns>Excel 列字母大写字符串</returns>
        internal static string GetExcelColumnLetter(int colIndex)
        {
            // 兜底防御，若列号非法默认返回首列 A --硬编码--
            if (colIndex <= 0) return "A";
            int div = colIndex;
            string colLetter = string.Empty;

            // 循环整除 26 进制基数
            while (div > 0)
            {
                // 计算当前位对应的 ASCII 字母
                int mod = (div - 1) % 26;
                colLetter = (char)('A' + mod) + colLetter;
                // 计算下一位商
                div = (div - mod) / 26;
            }

            // 返回最终计算出的列英文字母
            return colLetter;
        }

        /// <summary>
        /// 在指定单行区域内自适应匹配并替换占位符文本 (保留模板原汁原味的排版、前缀与格式)
        /// </summary>
        /// <param name="rowRange">目标单行 Excel 区域 (如 A11:Z11)</param>
        /// <param name="replacements">占位符与目标值映射字典 (例如 ["[柜号]"] = "SW-APKT")</param>
        /// <param name="fallbackCol">兜底列号 (可选，若整行未找到占位符时兜底写入的列号)</param>
        /// <param name="fallbackValue">兜底写入内容 (可选)</param>
        private static void ReplaceRowPlaceholders(
            dynamic rowRange,
            Dictionary<string, string> replacements,
            int fallbackCol = 0,
            string fallbackValue = "")
        {
            try
            {
                // 一次性将该单行读入内存二维数组 (规则 7 / 规则 12)
                object[,] matrix = (object[,])rowRange.Value2;
                int cols = matrix.GetLength(1);
                bool matchedAny = false;

                // 遍历每个单元格，执行占位符精准替换
                for (int c = 1; c <= cols; c++)
                {
                    string cellText = Convert.ToString(matrix[1, c]) ?? "";
                    if (string.IsNullOrWhiteSpace(cellText)) continue;

                    bool changed = false;
                    foreach (var kv in replacements)
                    {
                        // 若单元格内容包含当前占位符
                        if (cellText.Contains(kv.Key))
                        {
                            // 执行自适应原地字符串替换
                            cellText = cellText.Replace(kv.Key, kv.Value ?? "");
                            changed = true;
                            matchedAny = true;
                        }
                    }

                    if (changed)
                    {
                        // 单元格值发生变更，回写到对应单元格
                        rowRange.Cells[1, c].Value2 = cellText;
                    }
                }

                // 若整行均未命中且提供了兜底列号，执行安全回退写入
                if (!matchedAny && fallbackCol > 0 && !string.IsNullOrWhiteSpace(fallbackValue))
                {
                    rowRange.Cells[1, fallbackCol].Value2 = fallbackValue;
                }
            }
            catch (Exception ex)
            {
                // 捕获异常记录日志
                LogHelper.WriteLog($"[TenderReport] 替换单行占位符异常: {ex.Message}");
            }
        }
        /// 激活工作簿中第一个可见工作表
        /// </summary>
        private static void ActivateFirstVisibleSheet(dynamic wb)
        {
            try
            {
                foreach (dynamic ws in wb.Worksheets)
                {
                    if (ws.Visible == -1) // xlSheetVisible
                    {
                        ws.Activate();
                        break;
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 人民币金额数值转换为标准中文大写金额字符串 (如: 1250.50 -> 壹仟贰佰伍拾元伍角整)
        /// </summary>
        public static string ConvertAmountToChineseUpper(decimal amount)
        {
            if (amount == 0) return "零元整";

            var digits = new[] { "零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖" };
            var radices = new[] { "", "拾", "佰", "仟" };
            var bigRadices = new[] { "", "万", "亿", "兆" };

            string sign = amount < 0 ? "负" : "";
            amount = Math.Abs(amount);

            long integral = (long)Math.Truncate(amount);
            int decimalPart = (int)Math.Round((amount - integral) * 100);

            var sb = new StringBuilder();
            sb.Append(sign);

            // 整数部分转换
            if (integral > 0)
            {
                string intStr = integral.ToString();
                int len = intStr.Length;
                bool zero = false;

                for (int i = 0; i < len; i++)
                {
                    int n = intStr[i] - '0';
                    int p = len - i - 1;
                    int bigIndex = p / 4;
                    int subIndex = p % 4;

                    if (n != 0)
                    {
                        if (zero)
                        {
                            sb.Append("零");
                            zero = false;
                        }
                        sb.Append(digits[n]);
                        sb.Append(radices[subIndex]);
                    }
                    else
                    {
                        if (subIndex == 0 && bigIndex > 0)
                        {
                            // 跨万级或亿级处理
                        }
                        else
                        {
                            zero = true;
                        }
                    }

                    if (subIndex == 0 && bigIndex < bigRadices.Length)
                    {
                        sb.Append(bigRadices[bigIndex]);
                        zero = false;
                    }
                }

                sb.Append("元");
            }

            // 小数部分转换 (角、分)
            if (decimalPart == 0)
            {
                sb.Append("整");
            }
            else
            {
                int jiao = decimalPart / 10;
                int fen = decimalPart % 10;

                if (jiao > 0)
                {
                    sb.Append(digits[jiao] + "角");
                }
                else if (integral > 0)
                {
                    sb.Append("零");
                }

                if (fen > 0)
                {
                    sb.Append(digits[fen] + "分");
                }
            }

            return sb.ToString();
        }
    }
}
