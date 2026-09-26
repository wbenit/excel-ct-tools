using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：成套设备【成品交接单】母版分列式自动导出
    /// </summary>
    public static partial class ExcelServices
    {
        // 声明成品交接单分类选择向导窗口静态单例引用
        private static Forms.FinishedHandoverForm? _finishedHandoverForm;

        /// <summary>
        /// 扫描当前活动工作簿，获取所有有效分类明细表及箱柜统计，同时检测是否已存在成品交接单
        /// </summary>
        /// <returns>成品交接单初始化数据模型</returns>
        public static FinishedHandoverInitDataDto GetFinishedHandoverInitData()
        {
            // 初始化返回的初始化数据 DTO
            var initData = new FinishedHandoverInitDataDto();

            try
            {
                // 获取 Excel 宿主应用程序实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return initData;

                // 获取活动算价工作簿
                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null) return initData;

                // 提取工程项目名称 (默认取工作簿文件名)
                string projectName = Path.GetFileNameWithoutExtension(Convert.ToString(activeWb.Name) ?? "");

                // 遍历活动工作簿下的所有工作表
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    // 获取工作表名称并清洗两端空白
                    string sName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;

                    // 检查是否已包含【成品交接单】工作表 --硬编码: 系统工作表名称--
                    if (string.Equals(sName, "成品交接单", StringComparison.OrdinalIgnoreCase))
                    {
                        initData.HasExistingFinishedHandover = true;
                    }

                    // 尝试从【项目信息】工作表提取工程信息 --硬编码: 项目信息单元格--
                    if (string.Equals(sName, "项目信息", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            // B5 单元格为项目名称
                            string pName = Convert.ToString(sheet.Range["B5"].Value2)?.Trim() ?? "";
                            if (!string.IsNullOrWhiteSpace(pName)) projectName = pName;

                            // B14 单元格为客户单位名称 (客户信息)
                            string custName = Convert.ToString(sheet.Range["B14"].Value2)?.Trim() ?? "";
                            if (!string.IsNullOrWhiteSpace(custName)) initData.CustomerName = custName;

                            // B22 单元格为本企业单位名称
                            string compName = Convert.ToString(sheet.Range["B22"].Value2)?.Trim() ?? "";
                            if (!string.IsNullOrWhiteSpace(compName)) initData.CompanyName = compName;
                        }
                        catch { }
                    }

                    // 排除系统保留与非分类工作表 --硬编码: 系统保留表名--
                    if (string.Equals(sName, "项目信息", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "采购清单", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "领料清单", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "人工清单", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "成品交接单", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "材料分布表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "元件汇总分布表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "元件汇总表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "屏柜汇总表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "元器件数据管理", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // 白名单安全守门校验：是否为分类表
                    if (!Tool.IsProjectCategorySheet(sheet)) continue;

                    // 规则 8: 校验前确保规则 6 拓扑健康
                    Tool.FixAndFillCabinetNamesForSheet(sheet);

                    // 提取当前分类表下有效箱柜数量
                    var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);
                    int cabCount = validCabinets != null ? validCabinets.Count : 0;

                    // 若该分类下存在箱柜，添加至返回列表中
                    if (cabCount > 0)
                    {
                        initData.Categories.Add(new FinishedHandoverCategoryDto
                        {
                            SheetName = sName,
                            CabinetCount = cabCount,
                            IsSelected = true
                        });
                    }
                }

                // 回填提取到的项目名称
                initData.ProjectName = projectName;
            }
            catch (Exception ex)
            {
                // 异常日志捕获记录
                LogHelper.WriteLog($"[成品交接单] 初始化数据获取异常: {ex.Message}");
            }

            return initData;
        }

        /// <summary>
        /// 功能区 Ribbon 入口：根据分类表数量智能分流处理
        /// 若仅有 1 个分类表，则一键极简直达；若有多个分类表，则弹出选择向导
        /// </summary>
        public static void ShowFinishedHandoverDialogOrExport()
        {
            try
            {
                // 获取当前工作簿的分类表初始化数据
                var initData = GetFinishedHandoverInitData();

                // 校验是否存在有效分类表
                if (initData.Categories.Count == 0)
                {
                    MessageBox.Show("当前工作簿未检测到任何包含箱柜的有效分类工作表！\n请先新建分类并添加箱柜后再导出成品交接单。",
                        "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 场景 1: 若当前仅有 1 个分类表
                if (initData.Categories.Count == 1)
                {
                    // 若已存在成品交接单，采用弹窗防误触提醒确认
                    if (initData.HasExistingFinishedHandover)
                    {
                        var confirmResult = MessageBox.Show(
                            "当前工作簿已存在【成品交接单】工作表！\n重新生成将覆盖并重置该表，是否继续执行？",
                            "覆盖确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (confirmResult != DialogResult.Yes) return;
                    }

                    // 单表场景下，直接取该分类表名称一键秒级导出
                    var selectedSheets = new List<string> { initData.Categories[0].SheetName };
                    var exportResult = ExportFinishedHandoverToCurrentWorkbook(selectedSheets);

                    if (exportResult.Success)
                    {
                        MessageBox.Show($"成套设备【成品交接单】生成成功！\n共计生成 {exportResult.CabinetCount} 台箱柜成品交接明细。",
                            "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show($"导出成品交接单失败: {exportResult.Message}",
                            "导出失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }

                // 场景 2: 若包含多个分类表，弹出优雅的 WebView2 + Vue 3 选择向导窗口
                ShowModelessForm(ref _finishedHandoverForm, () => new Forms.FinishedHandoverForm());
            }
            catch (Exception ex)
            {
                // 顶层异常捕获与提示
                MessageBox.Show($"弹出成品交接单功能失败: {ex.Message}", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 核心导出服务：基于选中的一个或多个分类表，生成全新的【成品交接单】标准工作表 (母版分列式)
        /// 纯静默执行，零 Win32 阻塞弹窗，避免 WebView2 IPC 假死
        /// </summary>
        /// <param name="selectedCategorySheets">选中的分类工作表名称列表</param>
        /// <returns>导出执行结果模型</returns>
        public static FinishedHandoverExportResult ExportFinishedHandoverToCurrentWorkbook(List<string> selectedCategorySheets)
        {
            // 初始化返回的执行结果模型
            var result = new FinishedHandoverExportResult();

            // 前置入参安全防御校验
            if (selectedCategorySheets == null || selectedCategorySheets.Count == 0)
            {
                result.Message = "未选择任何需要导出成品交接单的分类工作表！";
                return result;
            }

            try
            {
                // 1. 获取 Excel 宿主应用程序实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    result.Message = "无法获取当前 Excel 宿主应用程序实例！";
                    return result;
                }

                // 2. 获取当前活动算价工作簿
                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null)
                {
                    result.Message = "当前未打开任何活动的 Excel 算价工作簿！";
                    return result;
                }

                // 3. 提取项目名称、客户名称与公司名称
                string projectName = Path.GetFileNameWithoutExtension(Convert.ToString(activeWb.Name) ?? "");
                string customerName = string.Empty;
                string companyName = string.Empty;

                foreach (dynamic ws in activeWb.Worksheets)
                {
                    if (string.Equals(Convert.ToString(ws.Name)?.Trim(), "项目信息", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            string pName = Convert.ToString(ws.Range["B5"].Value2)?.Trim() ?? "";
                            if (!string.IsNullOrWhiteSpace(pName)) projectName = pName;

                            string cust = Convert.ToString(ws.Range["B14"].Value2)?.Trim() ?? "";
                            if (!string.IsNullOrWhiteSpace(cust)) customerName = cust;

                            string comp = Convert.ToString(ws.Range["B22"].Value2)?.Trim() ?? "";
                            if (!string.IsNullOrWhiteSpace(comp)) companyName = comp;
                        }
                        catch { }
                        break;
                    }
                }

                // 加载管道决策规则配置供兜底判定箱柜型号
                var pipeConfig = LoadCabinetPipelineConfig();

                // 4. 遍历选中的分类表，提取所有箱柜及其明细数据
                var handoverItems = new List<FinishedHandoverItem>();
                int globalIndex = 0;

                foreach (string sheetName in selectedCategorySheets)
                {
                    // 定位分类工作表实例
                    dynamic? sheet = null;
                    foreach (dynamic ws in activeWb.Worksheets)
                    {
                        if (string.Equals(Convert.ToString(ws.Name)?.Trim(), sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            sheet = ws;
                            break;
                        }
                    }

                    if (sheet == null) continue;

                    // 规则 8: 操作前保证拓扑健康
                    Tool.FixAndFillCabinetNamesForSheet(sheet);

                    // 提取该分类表下全部有效箱柜锚点 (规则 6)
                    var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);
                    if (validCabinets == null || validCabinets.Count == 0) continue;

                    // 遍历各箱柜锚点
                    foreach (var kvp in validCabinets)
                    {
                        int cabK = kvp.Key;
                        var anchor = kvp.Value;
                        if (anchor == null) continue;
                        globalIndex++;

                        // 安全提取箱柜 4 个关键定义名称锚点的物理行号
                        int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                        int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;
                        int subsumRow = anchor.Subsum != null ? Convert.ToInt32(anchor.Subsum.Row) : 0;
                        int tolsumRow = anchor.Tolsum != null ? Convert.ToInt32(anchor.Tolsum.Row) : 0;

                        // 准备箱柜项基础属性
                        string cabName = "配电箱"; // 默认箱柜名称 --硬编码: 默认箱柜名称--
                        string cabCode = string.Empty; // 柜号 (C 列)
                        string cabModelD = string.Empty; // 汇总行 D 列型号
                        string cabModelN = string.Empty; // 汇总行 N 列三箱型号
                        decimal cabQty = 1; // 箱柜数量
                        string cabRemark = string.Empty; // 备注
                        string rawDimension = string.Empty; // 尺寸文本
                        int width = 0, height = 0, depth = 0;

                        // 步骤 A: 从顶部汇总行 (Cab_Sum_k) 提取基础信息
                        if (sumRow > 0)
                        {
                            try
                            {
                                // 规则 7: 一次性读取汇总行 A:N 列 (覆盖至第 14 列 N 列)
                                dynamic sumRange = sheet.Range[$"A{sumRow}:N{sumRow}"];
                                object[,] sumMatrix = (object[,])sumRange.Value2;
                                int sumCols = sumMatrix.GetLength(1);

                                // B 列为柜号 (第 2 列)
                                string bVal = Convert.ToString(sumMatrix[1, 2])?.Trim() ?? "";
                                if (!string.IsNullOrWhiteSpace(bVal)) cabCode = bVal;

                                // C 列为箱柜名称 (第 3 列)
                                string cVal = Convert.ToString(sumMatrix[1, 3])?.Trim() ?? "";
                                if (!string.IsNullOrWhiteSpace(cVal)) cabName = cVal;

                                // D 列为原箱柜型号 (第 4 列)
                                string dVal = Convert.ToString(sumMatrix[1, 4])?.Trim() ?? "";
                                if (!string.IsNullOrWhiteSpace(dVal)) cabModelD = dVal;

                                // F 列为数量 (第 6 列)
                                object fObj = sumMatrix[1, 6];
                                if (fObj != null && decimal.TryParse(fObj.ToString(), out decimal parsedQty) && parsedQty > 0)
                                {
                                    cabQty = parsedQty;
                                }

                                // I 列为备注 (第 9 列)
                                if (sumCols >= 9)
                                {
                                    string iVal = Convert.ToString(sumMatrix[1, 9])?.Trim() ?? "";
                                    if (!string.IsNullOrWhiteSpace(iVal)) cabRemark = iVal;
                                }

                                // M 列为箱柜尺寸 (第 13 列)
                                if (sumCols >= 13)
                                {
                                    string mVal = Convert.ToString(sumMatrix[1, 13])?.Trim() ?? "";
                                    if (!string.IsNullOrWhiteSpace(mVal))
                                    {
                                        rawDimension = mVal;
                                        if (TryParseShellDimensions(mVal, out int pw, out int ph, out int pd))
                                        {
                                            width = pw;
                                            height = ph;
                                            depth = pd;
                                        }
                                    }
                                }

                                // N 列为三箱型号回填值 (第 14 列)
                                if (sumCols >= 14)
                                {
                                    string nVal = Convert.ToString(sumMatrix[1, 14])?.Trim() ?? "";
                                    if (!string.IsNullOrWhiteSpace(nVal)) cabModelN = nVal;
                                }
                            }
                            catch (Exception exSum)
                            {
                                LogHelper.WriteLog($"[成品交接单] 提取箱柜 {cabK} 汇总行容错: {exSum.Message}");
                            }
                        }

                        // 若柜号为空，尝试从明细表信息行 (Cab_Det_k) 补充
                        if (detRow > 0 && string.IsNullOrWhiteSpace(cabCode))
                        {
                            try
                            {
                                dynamic detRange = sheet.Range[$"A{detRow}:D{detRow}"];
                                object[,] detMatrix = (object[,])detRange.Value2;
                                string bVal = Convert.ToString(detMatrix[1, 2])?.Trim() ?? "";
                                if (!string.IsNullOrWhiteSpace(bVal)) cabCode = bVal;
                            }
                            catch { }
                        }

                        // 步骤 B: 用户指令明确要求：直接读取 det + 2 行的 W 列 (第 23 列)，不需要容错，如果为空则填 0
                        int compStartRow = detRow > 0 ? detRow + 2 : 0;
                        int compEndRow = subsumRow > 0 ? subsumRow - 1 : 0;

                        // 初始化电流默认值为 0 (--硬编码: 为空默认填 0--)
                        string extractedCurrent = "0";
                        if (compStartRow > 0)
                        {
                            // 直接读取 det + 2 行第 23 列 (W 列) 的单元格值
                            object wObj = sheet.Cells[compStartRow, 23].Value2;
                            // 转换为去除首尾空格的字符串
                            string wVal = Convert.ToString(wObj)?.Trim() ?? "";
                            // 如果单元格有值则直接赋值，若为空则填 0 (--硬编码: 为空填 0--)
                            extractedCurrent = string.IsNullOrWhiteSpace(wVal) ? "0" : wVal;
                        }

                        // 元器件名称与型号列表 (供步骤 C 管道决策流兜底箱柜型号使用)
                        var compNameList = new List<string>();
                        var compModelList = new List<string>();

                        // 2. 扫描元器件区域提取名称与型号 (供步骤 C 管道决策流兜底箱柜型号使用)
                        if (compStartRow > 0 && compEndRow >= compStartRow)
                        {
                            try
                            {
                                // 规则 7: 二维数组一次性读取元器件区域 B 列(名称)与 C 列(型号规格)
                                dynamic compRange = sheet.Range[$"B{compStartRow}:C{compEndRow}"];
                                object[,] compMatrix = (object[,])compRange.Value2;
                                int cRowCount = compMatrix.GetLength(0);

                                for (int r = 1; r <= cRowCount; r++)
                                {
                                    // 第 1 列为元件名称 (B 列)
                                    string cName = Convert.ToString(compMatrix[r, 1])?.Trim() ?? "";
                                    // 第 2 列为型号规格 (C 列)
                                    string cModel = Convert.ToString(compMatrix[r, 2])?.Trim() ?? "";

                                    // 排除空行或表头无效行
                                    if (string.IsNullOrWhiteSpace(cName) && string.IsNullOrWhiteSpace(cModel)) continue;
                                    if (cName == "元件名称" || cName == "名称" || cName == "型号规格") continue;

                                    if (!string.IsNullOrWhiteSpace(cName)) compNameList.Add(cName);
                                    if (!string.IsNullOrWhiteSpace(cModel)) compModelList.Add(cModel);
                                }
                            }
                            catch (Exception exComp)
                            {
                                LogHelper.WriteLog($"[成品交接单] 提取元器件数据容错: {exComp.Message}");
                            }
                        }

                        // 步骤 C: 判定 F 列【箱柜型号】
                        // 用户指令：型号优先取汇总行的 N 列；若为空则自动通过管道判定兜底
                        string finalCabinetModel = string.Empty;
                        if (!string.IsNullOrWhiteSpace(cabModelN))
                        {
                            finalCabinetModel = cabModelN;
                        }
                        else
                        {
                            // 构造管道判定实体执行决策流推导
                            var pipeItem = new CabinetPipelineItemDto
                            {
                                SheetName = sheetName,
                                CabinetIndex = cabK,
                                CabinetNo = cabCode,
                                CabinetName = cabName,
                                CurrentModel = cabModelD,
                                Width = width,
                                Height = height,
                                Depth = depth,
                                RawDimensionText = rawDimension,
                                ComponentNames = compNameList,
                                ComponentModels = compModelList
                            };

                            // 执行单台箱柜管道决策流
                            EvaluateSingleCabinetInPipeline(pipeItem, pipeConfig);

                            if (!string.IsNullOrWhiteSpace(pipeItem.FinalModel) && pipeItem.FinalModel != "待确认")
                            {
                                finalCabinetModel = pipeItem.FinalModel;
                            }
                            else if (!string.IsNullOrWhiteSpace(cabModelD))
                            {
                                finalCabinetModel = cabModelD;
                            }
                            else
                            {
                                finalCabinetModel = "配电箱"; // --硬编码: 兜底箱柜型号--
                            }
                        }

                        // 步骤 D: 判定 H 列【防护等级】
                        // 用户指令：防护等级默认标准 IP30（若有特殊标注则提取）
                        string combinedText = $"{cabName} {cabCode} {cabModelD} {cabRemark} {rawDimension}";
                        string finalProtection = ExtractProtectionLevel(combinedText, "IP30"); // --硬编码: 默认防护等级 IP30--

                        // 组装成品交接单明细项
                        handoverItems.Add(new FinishedHandoverItem
                        {
                            Index = globalIndex,
                            Name = cabName,
                            CabinetCode = !string.IsNullOrWhiteSpace(cabCode) ? cabCode : (string.IsNullOrWhiteSpace(cabModelD) ? $"箱柜{cabK}" : cabModelD),
                            Quantity = cabQty,
                            Unit = "台", // --硬编码: 单位--
                            CabinetModel = finalCabinetModel,
                            Current = extractedCurrent,
                            ProtectionLevel = finalProtection,
                            Remark = !string.IsNullOrWhiteSpace(cabRemark) ? cabRemark : sheetName,
                            SheetName = sheetName,
                            DetailStartRow = detRow > 0 ? detRow : (sumRow > 0 ? sumRow : 1),
                            DetailEndRow = tolsumRow > 0 ? tolsumRow : (subsumRow > 0 ? subsumRow : (detRow > 0 ? detRow : 1))
                        });
                    }
                }

                // 校验提取结果有效性
                if (handoverItems.Count == 0)
                {
                    result.Message = "选中的分类表中未提取到任何有效的箱柜数据！";
                    return result;
                }

                // 5. 挂起屏幕刷新、事件与系统提示，杜绝事件交叉死锁
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.DisplayAlerts = false;

                try
                {
                    // 6. 查找已有成品交接单并安全删除旧表 (纯净克隆重置机制)
                    dynamic? handoverSheet = null;
                    foreach (dynamic ws in activeWb.Worksheets)
                    {
                        if (string.Equals(Convert.ToString(ws.Name)?.Trim(), "成品交接单", StringComparison.OrdinalIgnoreCase))
                        {
                            handoverSheet = ws;
                            break;
                        }
                    }

                    if (handoverSheet != null)
                    {
                        try
                        {
                            handoverSheet.Delete();
                            handoverSheet = null;
                        }
                        catch (Exception exDel)
                        {
                            LogHelper.WriteLog($"[成品交接单] 移除旧成品交接单工作表容错: {exDel.Message}");
                        }
                    }

                    // 7. 打开标准母版克隆纯净的标准【成品交接单】工作表
                    string templatePath = Controllers.ProjectController.EnsureCabinetTemplate(app);
                    if (!File.Exists(templatePath))
                    {
                        result.Message = $"未找到标准母版文件: {templatePath}";
                        return result;
                    }

                    dynamic templateWb = app.Workbooks.Open(templatePath, ReadOnly: true);
                    try
                    {
                        dynamic? tplSheet = null;
                        foreach (dynamic sh in templateWb.Worksheets)
                        {
                            if (string.Equals(Convert.ToString(sh.Name)?.Trim(), "成品交接单", StringComparison.OrdinalIgnoreCase))
                            {
                                tplSheet = sh;
                                break;
                            }
                        }

                        if (tplSheet == null)
                        {
                            result.Message = "母版 CabinetTemplate.xlsx 中未包含【成品交接单】标准工作表！";
                            return result;
                        }

                        // 复制模板工作表至当前算价工作簿最末尾
                        tplSheet.Copy(After: activeWb.Sheets[activeWb.Sheets.Count]);
                        handoverSheet = activeWb.Sheets[activeWb.Sheets.Count];
                        try { handoverSheet.Name = "成品交接单"; } catch { }
                    }
                    finally
                    {
                        // 复制完成后立即关闭模板工作簿句柄
                        try { templateWb.Close(false); } catch { }
                    }

                    // 8. 替换表头信息占位符 (第 1~5 行)
                    try
                    {
                        // A2 单元格：公司名称替换
                        if (!string.IsNullOrWhiteSpace(companyName))
                        {
                            string a2Val = Convert.ToString(handoverSheet.Range["A2"].Value2) ?? "";
                            if (a2Val.Contains("[公司名称]"))
                            {
                                handoverSheet.Range["A2"].Value2 = a2Val.Replace("[公司名称]", companyName);
                            }
                            else if (!string.IsNullOrWhiteSpace(a2Val))
                            {
                                handoverSheet.Range["A2"].Value2 = companyName;
                            }
                        }

                        // A4 单元格：单位名称 (客户名称)
                        if (!string.IsNullOrWhiteSpace(customerName))
                        {
                            handoverSheet.Range["A4"].Value2 = $"单位名称：{customerName}";
                        }

                        // A5 单元格：项目名称替换
                        string a5Val = Convert.ToString(handoverSheet.Range["A5"].Value2) ?? "";
                        if (a5Val.Contains("[[项目名称]]"))
                        {
                            handoverSheet.Range["A5"].Value2 = a5Val.Replace("[[项目名称]]", projectName);
                        }
                        else if (a5Val.Contains("[项目名称]"))
                        {
                            handoverSheet.Range["A5"].Value2 = a5Val.Replace("[项目名称]", projectName);
                        }
                        else
                        {
                            handoverSheet.Range["A5"].Value2 = $"项目名称：{projectName}";
                        }
                    }
                    catch (Exception exHeader)
                    {
                        LogHelper.WriteLog($"[成品交接单] 表头信息替换容错: {exHeader.Message}");
                    }

                    // 9. 自适应排版扩展 (模板预留 6 行: 7~12 行)
                    int itemCount = handoverItems.Count;
                    int startRow = 7; // 首个数据行 --硬编码: 首行--
                    int defaultTemplateRows = 6; // 默认预留 6 行 --硬编码: 预留行数--

                    if (itemCount > defaultTemplateRows)
                    {
                        int needInsert = itemCount - defaultTemplateRows;
                        int insertStartRow = startRow + defaultTemplateRows; // 第 13 行 (合计行)
                        int insertEndRow = insertStartRow + needInsert - 1;

                        // 整块向下推挤插入空行 (xlShiftDown = -4121)
                        dynamic insertRows = handoverSheet.Rows[$"{insertStartRow}:{insertEndRow}"];
                        insertRows.Insert(-4121);

                        // 复制第 7 行单元格格式 (xlPasteFormats = -4122)
                        handoverSheet.Rows[7].Copy();
                        dynamic formatTarget = handoverSheet.Rows[$"{insertStartRow}:{insertEndRow}"];
                        formatTarget.PasteSpecial(-4122);
                        app.CutCopyMode = 0;
                    }

                    // 计算数据实际末行与合计行行号
                    int endDataRow = startRow + Math.Max(itemCount, defaultTemplateRows) - 1;
                    int actualItemEndRow = startRow + itemCount - 1;
                    int totalRow = endDataRow + 1; // 合计行物理行号

                    // 10. 构造数据写入二维矩阵并注入超链接与各列数据
                    // 母版分列式: actualItemCount 行 x 9 列 (A~I 列)
                    object[,] writeMatrix = new object[itemCount, 9];

                    for (int i = 0; i < itemCount; i++)
                    {
                        var it = handoverItems[i];

                        // A 列 (序号 + 超链接公式，用户指令: 与人工清单一样生成 HYPERLINK 公式)
                        string linkFormula = $"=HYPERLINK(\"#'{it.SheetName}'!A{it.DetailStartRow}:H{it.DetailEndRow}\", {it.Index})";
                        writeMatrix[i, 0] = linkFormula;

                        // B 列 (名称: 配电箱)
                        writeMatrix[i, 1] = it.Name;

                        // C 列 (型 号: 柜号，如 2-ALZ11)
                        writeMatrix[i, 2] = it.CabinetCode;

                        // D 列 (数量: 台数)
                        writeMatrix[i, 3] = it.Quantity;

                        // E 列 (单位: 台)
                        writeMatrix[i, 4] = it.Unit;

                        // F 列 (箱柜型号: 三箱型号，如 PZ30, XL21)
                        writeMatrix[i, 5] = it.CabinetModel;

                        // G 列 (电流: 主进线电流，如 100A, 63A)
                        writeMatrix[i, 6] = it.Current;

                        // H 列 (防护等级: 如 IP30, IP54)
                        writeMatrix[i, 7] = it.ProtectionLevel;

                        // I 列 (备注: 楼栋编号或分类名)
                        writeMatrix[i, 8] = it.Remark;
                    }

                    // 规则 7: 内存二维矩阵一次性回写数据与公式 (通过 Formula 属性保证公式生效)
                    dynamic targetDataRange = handoverSheet.Range[handoverSheet.Cells[startRow, 1], handoverSheet.Cells[actualItemEndRow, 9]];
                    targetDataRange.Formula = writeMatrix;

                    // 若实际数据小于预留的 6 行，清空剩余预留行的占位符内容，保持网格线纯净
                    if (itemCount < defaultTemplateRows)
                    {
                        int cleanStartRow = actualItemEndRow + 1;
                        dynamic cleanRange = handoverSheet.Range[handoverSheet.Cells[cleanStartRow, 1], handoverSheet.Cells[endDataRow, 9]];
                        cleanRange.ClearContents();
                    }

                    // 11. 批量设置行高、边框与文本对齐样式
                    try
                    {
                        dynamic allDataRange = handoverSheet.Range[handoverSheet.Cells[startRow, 1], handoverSheet.Cells[endDataRow, 9]];
                        allDataRange.Borders.LineStyle = 1; // 细实线边框
                        allDataRange.Borders.Weight = 2;
                        allDataRange.RowHeight = 28; // 舒适行高

                        // 居中对齐列: A~H 列 (xlCenter = -4108)
                        for (int c = 1; c <= 8; c++)
                        {
                            handoverSheet.Range[handoverSheet.Cells[startRow, c], handoverSheet.Cells[endDataRow, c]].HorizontalAlignment = -4108;
                        }
                        // 左对齐列: I 列 (备注) (xlLeft = -4131)
                        handoverSheet.Range[handoverSheet.Cells[startRow, 9], handoverSheet.Cells[endDataRow, 9]].HorizontalAlignment = -4131;
                    }
                    catch (Exception exStyle)
                    {
                        LogHelper.WriteLog($"[成品交接单] 样式设置容错: {exStyle.Message}");
                    }

                    // 12. 自适应刷新合计行公式与清理多余公式
                    try
                    {
                        // D 列数量合计公式: =SUM(D6:D{endDataRow})
                        handoverSheet.Range[$"D{totalRow}"].Formula = $"=SUM(D6:D{endDataRow})";

                        // 清除 H 列可能残留的求和公式 (防范显示 0)
                        try { handoverSheet.Range[$"H{totalRow}"].ClearContents(); } catch { }
                    }
                    catch (Exception exTotal)
                    {
                        LogHelper.WriteLog($"[成品交接单] 刷新合计行容错: {exTotal.Message}");
                    }

                    // 13. 激活工作表并定位到 A1
                    try
                    {
                        handoverSheet.Activate();
                        handoverSheet.Range["A1"].Select();
                    }
                    catch { }

                    // 回填成功导出结果
                    result.Success = true;
                    result.CabinetCount = itemCount;
                    result.Message = $"成套设备【成品交接单】生成成功，共计 {itemCount} 台箱柜。";
                }
                finally
                {
                    // 14. 恢复 Excel 宿主交互与屏幕刷新
                    app.ScreenUpdating = true;
                    app.EnableEvents = true;
                    app.DisplayAlerts = true;
                }
            }
            catch (Exception exExport)
            {
                // 导出顶层异常捕获
                result.Success = false;
                result.Message = $"生成成品交接单异常: {exExport.Message}";
                LogHelper.WriteLog($"[成品交接单] 导出全流程异常: {exExport}");
            }

            return result;
        }



        /// <summary>
        /// 智能提取防护等级 (如 IP30, IP54, IP55, IP65)
        /// </summary>
        /// <param name="text">候选文本</param>
        /// <param name="defaultIp">默认防护等级</param>
        /// <returns>防护等级字符串</returns>
        private static string ExtractProtectionLevel(string text, string defaultIp = "IP30") // --硬编码: 默认防护等级 IP30--
        {
            if (string.IsNullOrWhiteSpace(text)) return defaultIp;

            // 匹配 IPxx 格式
            var m = Regex.Match(text, @"\b(IP\s*\d{2})\b", RegexOptions.IgnoreCase);
            if (m.Success)
            {
                // 去除内部空格并转为大写标准格式
                return m.Groups[1].Value.Replace(" ", "").ToUpperInvariant();
            }

            return defaultIp;
        }
    }
}
