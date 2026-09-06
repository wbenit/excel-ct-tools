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

                // 临时挂起屏幕刷新与自动计算提升填充性能
                app.ScreenUpdating = false;
                app.Calculation = -4135; // xlCalculationManual 手动重算

                try
                {
                    // 4. 填充《封面》工作表
                    PopulateCoverWorksheet(reportWb, config.ProjectInfo, config.IncludeCover);

                    // 5. 填充《屏柜汇总表》工作表
                    PopulateSummaryWorksheet(reportWb, config.ProjectInfo, exportCategories, config.IncludeSummary);

                    // 6. 填充《屏柜分项表》工作表
                    PopulateDetailWorksheet(reportWb, config.ProjectInfo, exportCategories, config.IncludeDetail);

                    // 恢复自动计算并执行一次全局重算
                    app.Calculation = -4105; // xlCalculationAutomatic
                    reportWb.Calculate();

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
                    // 恢复屏幕刷新
                    app.ScreenUpdating = true;
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
                if (targetNameSet.Count > 0 && !targetNameSet.Contains(sName))
                {
                    continue;
                }

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
        /// 填充《屏柜汇总表》工作表
        /// </summary>
        private static void PopulateSummaryWorksheet(dynamic reportWb, TenderReportProjectInfo proj, List<TenderReportCategoryGroup> categories, bool includeSummary)
        {
            dynamic? sumSheet = null;
            try { sumSheet = reportWb.Sheets["屏柜汇总表"]; } catch { }
            if (sumSheet == null) return;

            if (!includeSummary)
            {
                try { sumSheet.Delete(); } catch { }
                return;
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

                // 2. 获取模板中的母版行锚点
                // 模板原结构：Row 11 为分类标题，Row 12 为表头，Row 13 为箱柜数据行模板，Row 14 为小计模板，Row 15 为总计模板
                int templateCatTitleRow = 11;
                int templateHeaderRow = 12;

                // 记录所有生成的小计行行号，用于项目总计求和
                var subtotalRows = new List<int>();
                int currentRow = templateCatTitleRow;
                int globalIndex = 1;

                // 循环填充各个分类块
                for (int catIdx = 0; catIdx < categories.Count; catIdx++)
                {
                    var cat = categories[catIdx];

                    // 写入分类标题行
                    sumSheet.Rows[currentRow].Value2 = sumSheet.Rows[templateCatTitleRow].Value2;
                    sumSheet.Cells[currentRow, 1].Value2 = $"分类：{cat.CategoryName}";
                    currentRow++;

                    // 写入表头行
                    sumSheet.Rows[currentRow].Value2 = sumSheet.Rows[templateHeaderRow].Value2;
                    currentRow++;

                    int catDataStartRow = currentRow;
                    int catItemCount = cat.Cabinets.Count;

                    if (catItemCount > 0)
                    {
                        // 准备二维数组批量灌入当前分类的全部箱柜行 (规则 12)
                        object[,] catMatrix = new object[catItemCount, 9];
                        for (int k = 0; k < catItemCount; k++)
                        {
                            var cab = cat.Cabinets[k];
                            catMatrix[k, 0] = globalIndex++;                  // 序号 (Col 1)
                            catMatrix[k, 1] = cab.CabinetNo;                   // 柜号 (Col 2)
                            catMatrix[k, 2] = cab.CabinetName;                 // 箱柜名称 (Col 3)
                            catMatrix[k, 3] = cab.CabinetModel;                // 箱柜型号 (Col 4)
                            catMatrix[k, 4] = string.IsNullOrWhiteSpace(cab.Unit) ? "台" : cab.Unit; // 单位 (Col 5)
                            catMatrix[k, 5] = cab.Quantity;                    // 数量 (Col 6)
                            catMatrix[k, 6] = cab.UnitPrice;                   // 单价 (Col 7)
                            catMatrix[k, 7] = $"=F{currentRow + k}*G{currentRow + k}"; // 总价动态公式 (Col 8)
                            catMatrix[k, 8] = cab.Remark;                      // 备注 (Col 9)
                        }

                        // 批量写入数据区域 (规则 12)
                        dynamic dataRange = sumSheet.Range[sumSheet.Cells[currentRow, 1], sumSheet.Cells[currentRow + catItemCount - 1, 9]];
                        dataRange.Value2 = catMatrix;
                        currentRow += catItemCount;
                    }

                    int catDataEndRow = currentRow - 1;

                    // 写入分类小计行
                    sumSheet.Cells[currentRow, 2].Value2 = "合计";
                    sumSheet.Cells[currentRow, 5].Value2 = "台";
                    if (catDataEndRow >= catDataStartRow)
                    {
                        sumSheet.Cells[currentRow, 6].Formula = $"=SUM(F{catDataStartRow}:F{catDataEndRow})";
                        sumSheet.Cells[currentRow, 8].Formula = $"=SUM(H{catDataStartRow}:H{catDataEndRow})";
                    }
                    else
                    {
                        sumSheet.Cells[currentRow, 6].Value2 = 0;
                        sumSheet.Cells[currentRow, 8].Value2 = 0;
                    }

                    subtotalRows.Add(currentRow);
                    currentRow++;
                }

                // 3. 写入项目总计行
                int grandTotalRow = currentRow;
                sumSheet.Cells[grandTotalRow, 2].Value2 = "总计";
                sumSheet.Cells[grandTotalRow, 5].Value2 = "台";

                if (subtotalRows.Count > 0)
                {
                    // 拼接所有分类合计行的求和公式
                    string qtySumFormula = "=" + string.Join("+", subtotalRows.Select(r => $"F{r}"));
                    string amtSumFormula = "=" + string.Join("+", subtotalRows.Select(r => $"H{r}"));
                    sumSheet.Cells[grandTotalRow, 6].Formula = qtySumFormula;
                    sumSheet.Cells[grandTotalRow, 3].Formula = amtSumFormula; // 总计小计数字 (Col 3)

                    // 计算总金额大写并写入 Col 7
                    decimal grandTotalAmount = categories.Sum(c => c.TotalAmount);
                    sumSheet.Cells[grandTotalRow, 7].Value2 = ConvertAmountToChineseUpper(grandTotalAmount);
                }
                currentRow++;

                // 4. 写入报价说明与签名尾栏 (报价人、审核人等)
                sumSheet.Cells[currentRow, 1].Value2 = "报价说明:";
                currentRow++;
                sumSheet.Cells[currentRow, 2].Value2 = proj.ProjectRemark;
                currentRow++;
                sumSheet.Cells[currentRow, 2].Value2 = $"报价人：{proj.Bidder}";
                sumSheet.Cells[currentRow, 4].Value2 = $"报价日期：{proj.QuotationDate}";
                sumSheet.Cells[currentRow, 8].Value2 = "审核人：";
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[TenderReport] 填充屏柜汇总表异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充《屏柜分项表》工作表
        /// </summary>
        private static void PopulateDetailWorksheet(dynamic reportWb, TenderReportProjectInfo proj, List<TenderReportCategoryGroup> categories, bool includeDetail)
        {
            dynamic? detSheet = null;
            try { detSheet = reportWb.Sheets["屏柜分项表"]; } catch { }
            if (detSheet == null) return;

            if (!includeDetail)
            {
                try { detSheet.Delete(); } catch { }
                return;
            }

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

                // 2. 母版基准行 (Row 11 箱柜信息，Row 12 元件表头，Row 13 元件模板，Row 14 箱柜总计)
                int templateCabInfoRow = 11;

                int currentRow = templateCabInfoRow;
                int globalCabIndex = 1;

                // 循环每个分类中的每个箱柜进行展开
                for (int cIdx = 0; cIdx < categories.Count; cIdx++)
                {
                    var cat = categories[cIdx];

                    for (int cabIdx = 0; cabIdx < cat.Cabinets.Count; cabIdx++)
                    {
                        var cab = cat.Cabinets[cabIdx];

                        // ① 写入箱柜信息栏 (Row 11 样式)
                        detSheet.Cells[currentRow, 1].Value2 = globalCabIndex++;
                        detSheet.Cells[currentRow, 2].Value2 = $"柜号：{cab.CabinetNo}";
                        detSheet.Cells[currentRow, 3].Value2 = "型号：";
                        detSheet.Cells[currentRow, 4].Value2 = cab.CabinetModel;
                        detSheet.Cells[currentRow, 7].Value2 = "名称：";
                        detSheet.Cells[currentRow, 8].Value2 = cab.CabinetName;
                        detSheet.Cells[currentRow, 9].Value2 = string.IsNullOrWhiteSpace(cab.Remark) ? "" : $"备注:{cab.Remark}";
                        currentRow++;

                        // ② 写入元件表头栏 (序号 | 元件名称 | 型号规格 | 单位 | 数量 | 单价 | 总价 | 生产厂家 | 备注)
                        detSheet.Cells[currentRow, 1].Value2 = "序号";
                        detSheet.Cells[currentRow, 2].Value2 = "元件名称";
                        detSheet.Cells[currentRow, 3].Value2 = "型号规格";
                        detSheet.Cells[currentRow, 4].Value2 = "单位";
                        detSheet.Cells[currentRow, 5].Value2 = "数量";
                        detSheet.Cells[currentRow, 6].Value2 = "单   价";
                        detSheet.Cells[currentRow, 7].Value2 = "总   价";
                        detSheet.Cells[currentRow, 8].Value2 = "生产厂家";
                        detSheet.Cells[currentRow, 9].Value2 = "备注";
                        currentRow++;

                        // ③ 批量灌入元器件明细行 (采用二维数组一次性回写，规则 12)
                        int compCount = cab.Components.Count;
                        int compStartRow = currentRow;

                        if (compCount > 0)
                        {
                            object[,] compMatrix = new object[compCount, 9];
                            for (int ci = 0; ci < compCount; ci++)
                            {
                                var item = cab.Components[ci];
                                compMatrix[ci, 0] = item.Index;
                                compMatrix[ci, 1] = item.Name;
                                compMatrix[ci, 2] = item.Model;
                                compMatrix[ci, 3] = item.Unit;
                                compMatrix[ci, 4] = item.Quantity;
                                compMatrix[ci, 5] = item.UnitPrice;
                                compMatrix[ci, 6] = $"=E{currentRow + ci}*F{currentRow + ci}"; // 动态总价公式
                                compMatrix[ci, 7] = item.Manufacturer;
                                compMatrix[ci, 8] = item.Remark;
                            }

                            dynamic compRange = detSheet.Range[detSheet.Cells[currentRow, 1], detSheet.Cells[currentRow + compCount - 1, 9]];
                            compRange.Value2 = compMatrix;
                            currentRow += compCount;
                        }

                        int compEndRow = currentRow - 1;

                        // ④ 写入箱柜总计栏 (数量、总价)
                        detSheet.Cells[currentRow, 2].Value2 = "总计";
                        detSheet.Cells[currentRow, 4].Value2 = "台";
                        detSheet.Cells[currentRow, 5].Value2 = cab.Quantity; // 箱柜台数

                        // 若有元件明细，总价采用对上方求和并乘以台数的公式，或直接写入计算金额
                        if (compEndRow >= compStartRow)
                        {
                            detSheet.Cells[currentRow, 7].Formula = $"=SUM(G{compStartRow}:G{compEndRow})*E{currentRow}";
                        }
                        else
                        {
                            detSheet.Cells[currentRow, 7].Value2 = cab.TotalPrice;
                        }

                        currentRow += 2; // 箱柜块之间空一行保持美观间距
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[TenderReport] 填充屏柜分项表异常: {ex.Message}");
            }
        }

        /// <summary>
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
