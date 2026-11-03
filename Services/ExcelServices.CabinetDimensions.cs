using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：成套设备【尺寸安装方式】标准报表自动生成与导出
    /// </summary>
    public static partial class ExcelServices
    {
        // 声明尺寸安装方式分类选择向导窗口静态单例引用
        private static Forms.CabinetDimensionsForm? _cabinetDimensionsForm;

        /// <summary>
        /// 扫描当前活动工作簿，获取所有有效分类明细表及箱柜统计，检测是否已存在尺寸安装方式表
        /// </summary>
        /// <returns>尺寸安装方式初始化数据模型</returns>
        public static CabinetDimensionsInitDataDto GetCabinetDimensionsInitData()
        {
            // 初始化返回的初始化数据 DTO 实例
            var initData = new CabinetDimensionsInitDataDto();

            try
            {
                // 获取 Excel 宿主应用程序实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return initData;

                // 获取当前活动的工程算价工作簿
                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null) return initData;

                // 提取工程项目名称 (默认取工作簿文件名)
                string projectName = Path.GetFileNameWithoutExtension(Convert.ToString(activeWb.Name) ?? "");

                // 遍历活动工作簿下的所有工作表
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    // 获取工作表名称并清洗首尾空白
                    string sName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;

                    // 检查是否已包含【尺寸安装方式】工作表 --硬编码: 系统工作表名称--
                    if (string.Equals(sName, "尺寸安装方式", StringComparison.OrdinalIgnoreCase))
                    {
                        initData.HasExistingDimensionsSheet = true;
                    }

                    // 尝试从【项目信息】工作表提取工程信息 --硬编码: 项目信息单元格--
                    if (string.Equals(sName, "项目信息", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            // 规则 7: 一次性读取 A1:C30 区域至内存二维数组
                            dynamic infoRange = sheet.Range["A1:C30"];
                            object[,] infoMatrix = (object[,])infoRange.Value2;

                            // B5 单元格为项目名称 (Row 5, Col 2)
                            string pName = Convert.ToString(infoMatrix[5, 2])?.Trim() ?? "";
                            if (!string.IsNullOrWhiteSpace(pName)) projectName = pName;

                            // 遍历扫描客户名称 (优先 Cell B14，若空则扫描 A 列包含“客户”的行)
                            string custName = Convert.ToString(infoMatrix[14, 2])?.Trim() ?? "";
                            if (string.IsNullOrWhiteSpace(custName))
                            {
                                for (int r = 1; r <= 30; r++)
                                {
                                    string aText = Convert.ToString(infoMatrix[r, 1])?.Trim() ?? "";
                                    if (aText.Contains("客户"))
                                    {
                                        string bVal = Convert.ToString(infoMatrix[r, 2])?.Trim() ?? "";
                                        if (!string.IsNullOrWhiteSpace(bVal)) { custName = bVal; break; }
                                    }
                                }
                            }
                            if (!string.IsNullOrWhiteSpace(custName)) initData.CustomerName = custName;

                            // B22 单元格为本企业单位名称，若空回退 B1
                            string compName = Convert.ToString(infoMatrix[22, 2])?.Trim() ?? "";
                            if (string.IsNullOrWhiteSpace(compName)) compName = Convert.ToString(infoMatrix[1, 2])?.Trim() ?? "";
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
                        string.Equals(sName, "尺寸安装方式", StringComparison.OrdinalIgnoreCase) ||
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

                    // 提取当前分类表下有效箱柜列表
                    var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);
                    // 仅统计汇总行具有有效数量 (> 0) 的箱柜
                    int cabCount = 0;
                    if (validCabinets != null && validCabinets.Count > 0)
                    {
                        // 遍历各个箱柜检测汇总行数量
                        foreach (var kvp in validCabinets)
                        {
                            // 检查是否存在汇总行锚点
                            if (kvp.Value.Sum != null)
                            {
                                int sRow = Convert.ToInt32(kvp.Value.Sum.Row);
                                // 读取汇总行 F 列 (第 6 列) 数量
                                object? qVal = sheet.Cells[sRow, 6].Value2;
                                if (qVal != null && decimal.TryParse(Convert.ToString(qVal)?.Trim(), out decimal q) && q > 0)
                                {
                                    // 仅当数量存在且大于 0 时计入
                                    cabCount++;
                                }
                            }
                        }
                    }

                    // 若该分类下存在有效箱柜，添加至返回列表中
                    if (cabCount > 0)
                    {
                        initData.Categories.Add(new CabinetDimensionsCategoryDto
                        {
                            SheetName = sName,
                            CabinetCount = cabCount,
                            IsSelected = true
                        });
                    }
                }

                // 回填项目名称
                initData.ProjectName = projectName;
            }
            catch (Exception ex)
            {
                // 异常日志捕获记录
                LogHelper.WriteLog($"[尺寸安装方式] 初始化数据获取异常: {ex.Message}");
            }

            return initData;
        }

        /// <summary>
        /// 功能区 Ribbon 入口：根据分类表数量智能分流处理
        /// 若仅有 1 个分类表，则一键极简直达；若有多个分类表，则弹出选择向导
        /// </summary>
        public static void ShowCabinetDimensionsDialogOrExport()
        {
            try
            {
                // 获取当前工作簿的分类表初始化数据
                var initData = GetCabinetDimensionsInitData();

                // 校验是否存在有效分类表
                if (initData.Categories.Count == 0)
                {
                    MessageBox.Show("当前工作簿未检测到任何包含箱柜的有效分类工作表！\n请先新建分类并添加箱柜后再生成尺寸安装方式表。",
                        "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 场景 1: 若当前仅有 1 个分类表
                if (initData.Categories.Count == 1)
                {
                    // 若已存在尺寸安装方式表，采用弹窗防误触提醒确认
                    if (initData.HasExistingDimensionsSheet)
                    {
                        var confirmResult = MessageBox.Show(
                            "当前工作簿已存在【尺寸安装方式】工作表！\n重新生成将覆盖并重置该表，是否继续执行？",
                            "覆盖确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (confirmResult != DialogResult.Yes) return;
                    }

                    // 单表场景下，直接取该分类表名称一键秒级导出
                    var selectedSheets = new List<string> { initData.Categories[0].SheetName };
                    var exportResult = ExportCabinetDimensionsToCurrentWorkbook(selectedSheets);

                    if (exportResult.Success)
                    {
                        MessageBox.Show($"成套设备【尺寸安装方式】生成成功！\n共计生成 {exportResult.CabinetCount} 台箱柜明细。",
                            "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show($"导出尺寸安装方式表失败: {exportResult.Message}",
                            "导出失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }

                // 场景 2: 若包含多个分类表，弹出基于 WebView2 + Vue 3 的选择向导窗口
                ShowModelessForm(ref _cabinetDimensionsForm, () => new Forms.CabinetDimensionsForm());
            }
            catch (Exception ex)
            {
                // 顶层异常捕获与提示
                MessageBox.Show($"弹出尺寸安装方式功能失败: {ex.Message}", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 核心导出服务：基于选中的一个或多个分类表，生成全新的【尺寸安装方式】标准工作表
        /// 遵循规则 6、7、8、9，使用内存数组批量读写
        /// </summary>
        /// <param name="selectedCategorySheets">选中的分类工作表名称列表</param>
        /// <returns>导出执行结果模型</returns>
        public static CabinetDimensionsExportResult ExportCabinetDimensionsToCurrentWorkbook(List<string> selectedCategorySheets)
        {
            // 初始化返回的执行结果模型
            var result = new CabinetDimensionsExportResult();

            // 前置入参安全防御校验
            if (selectedCategorySheets == null || selectedCategorySheets.Count == 0)
            {
                result.Message = "未选择任何需要导出尺寸安装方式的分类工作表！";
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
                            // 规则 7: 一次性读取 A1:C30 区域至内存二维数组
                            dynamic infoRange = ws.Range["A1:C30"];
                            object[,] infoMatrix = (object[,])infoRange.Value2;

                            // B5 单元格为项目名称 (Row 5, Col 2)
                            string pName = Convert.ToString(infoMatrix[5, 2])?.Trim() ?? "";
                            if (!string.IsNullOrWhiteSpace(pName)) projectName = pName;

                            // 遍历扫描客户名称 (优先 Cell B14，若空则扫描 A 列包含“客户”的行)
                            string cust = Convert.ToString(infoMatrix[14, 2])?.Trim() ?? "";
                            if (string.IsNullOrWhiteSpace(cust))
                            {
                                for (int r = 1; r <= 30; r++)
                                {
                                    string aText = Convert.ToString(infoMatrix[r, 1])?.Trim() ?? "";
                                    if (aText.Contains("客户"))
                                    {
                                        string bVal = Convert.ToString(infoMatrix[r, 2])?.Trim() ?? "";
                                        if (!string.IsNullOrWhiteSpace(bVal)) { cust = bVal; break; }
                                    }
                                }
                            }
                            if (!string.IsNullOrWhiteSpace(cust)) customerName = cust;

                            // B22 单元格为本企业单位名称，若空回退 B1
                            string comp = Convert.ToString(infoMatrix[22, 2])?.Trim() ?? "";
                            if (string.IsNullOrWhiteSpace(comp)) comp = Convert.ToString(infoMatrix[1, 2])?.Trim() ?? "";
                            if (!string.IsNullOrWhiteSpace(comp)) companyName = comp;
                        }
                        catch { }
                        break;
                    }
                }

                // 4. 遍历选中的分类表，提取所有箱柜及其汇总行数据
                var dimensionsItems = new List<CabinetDimensionsItem>();
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

                        // 安全提取箱柜 4 个关键定义名称锚点的物理行号
                        int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                        int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;
                        int subsumRow = anchor.Subsum != null ? Convert.ToInt32(anchor.Subsum.Row) : 0;
                        int tolsumRow = anchor.Tolsum != null ? Convert.ToInt32(anchor.Tolsum.Row) : 0;

                        // 准备箱柜项基础属性
                        string cabName = "配电箱"; // 默认箱柜名称 --硬编码: 默认箱柜名称--
                        string cabCode = string.Empty; // 柜号 (B 列)
                        string cabModelD = string.Empty; // 汇总行 D 列型号/安装方式
                        decimal cabQty = 1; // 箱柜数量
                        bool skipCabinet = false;
                        string cabRemark = string.Empty; // 备注 (I 列)
                        string cabDimension = string.Empty; // 箱柜尺寸 (M 列)

                        // 步骤 A: 从顶部汇总行 (Cab_Sum_k) 提取基础信息
                        if (sumRow > 0)
                        {
                            try
                            {
                                // 规则 7: 一次性读取汇总行 A:M 列 (至少覆盖到第 13 列 M 列)
                                dynamic sumRange = sheet.Range[$"A{sumRow}:M{sumRow}"];
                                object[,] sumMatrix = (object[,])sumRange.Value2;
                                int sumCols = sumMatrix.GetLength(1);

                                // B 列为柜号 (第 2 列)
                                string bVal = Convert.ToString(sumMatrix[1, 2])?.Trim() ?? "";
                                if (!string.IsNullOrWhiteSpace(bVal)) cabCode = bVal;

                                // C 列为箱柜名称 (第 3 列)
                                string cVal = Convert.ToString(sumMatrix[1, 3])?.Trim() ?? "";
                                if (!string.IsNullOrWhiteSpace(cVal)) cabName = cVal;

                                // D 列为箱柜型号/安装方式 (第 4 列)
                                // 用户明确指令：[箱柜型号]直接取D列数据，为空也填空
                                string dVal = Convert.ToString(sumMatrix[1, 4])?.Trim() ?? "";
                                cabModelD = dVal;

                                // F 列为数量 (第 6 列)
                                object fObj = sumMatrix[1, 6];
                                if (fObj != null && decimal.TryParse(fObj.ToString()?.Trim(), out decimal parsedQty) && parsedQty > 0)
                                {
                                    cabQty = parsedQty;
                                }
                                else
                                {
                                    // 数量为空、null或<=0均跳过不纳入尺寸安装方式表
                                    skipCabinet = true;
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
                                    cabDimension = mVal;
                                }
                            }
                            catch (Exception exSum)
                            {
                                LogHelper.WriteLog($"[尺寸安装方式] 提取箱柜 {cabK} 汇总行容错: {exSum.Message}");
                            }
                        }

                        // 若该箱柜数量明确为 0 或未填，直接跳过
                        if (skipCabinet) continue;

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

                        globalIndex++;

                        // 组装单台箱柜尺寸安装方式项
                        dimensionsItems.Add(new CabinetDimensionsItem
                        {
                            Index = globalIndex,
                            Name = cabName,
                            CabinetCode = !string.IsNullOrWhiteSpace(cabCode) ? cabCode : $"箱柜{cabK}",
                            Quantity = cabQty,
                            Unit = "台", // --硬编码: 单位--
                            Dimensions = cabDimension,
                            InstallMode = cabModelD, // 用户明确指令：直接取 D 列数据
                            ProtectionLevel = "IP30", // 用户明确指令：统一填默认值 IP30 --硬编码--
                            Remark = cabRemark,
                            SheetName = sheetName,
                            DetailStartRow = detRow > 0 ? detRow : (sumRow > 0 ? sumRow : 1),
                            DetailEndRow = tolsumRow > 0 ? tolsumRow : (subsumRow > 0 ? subsumRow : (detRow > 0 ? detRow : 1))
                        });
                    }
                }

                // 校验提取结果有效性
                if (dimensionsItems.Count == 0)
                {
                    result.Message = "选中的分类表中未提取到任何有效的箱柜数据！";
                    return result;
                }

                // 5. 挂起屏幕刷新、事件与系统提示，提升执行性能并杜绝事件冲突
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.DisplayAlerts = false;

                try
                {
                    // 6. 查找已有尺寸安装方式表并安全删除旧表 (纯净克隆重置机制)
                    dynamic? targetSheet = null;
                    foreach (dynamic ws in activeWb.Worksheets)
                    {
                        if (string.Equals(Convert.ToString(ws.Name)?.Trim(), "尺寸安装方式", StringComparison.OrdinalIgnoreCase))
                        {
                            targetSheet = ws;
                            break;
                        }
                    }

                    if (targetSheet != null)
                    {
                        try
                        {
                            targetSheet.Delete();
                            targetSheet = null;
                        }
                        catch (Exception exDel)
                        {
                            LogHelper.WriteLog($"[尺寸安装方式] 移除旧尺寸安装方式工作表容错: {exDel.Message}");
                        }
                    }

                    // 7. 打开标准母版克隆纯净的标准工作表
                    string templatePath = Controllers.ProjectController.EnsureCabinetTemplate(app);
                    if (!File.Exists(templatePath))
                    {
                        result.Message = $"未找到标准母版文件: {templatePath}";
                        return result;
                    }

                    dynamic templateWb = app.Workbooks.Open(templatePath, ReadOnly: true);
                    bool needAdjustHeaders = false;

                    try
                    {
                        dynamic? tplSheet = null;
                        // 优先寻找【尺寸安装方式】母版
                        foreach (dynamic sh in templateWb.Worksheets)
                        {
                            if (string.Equals(Convert.ToString(sh.Name)?.Trim(), "尺寸安装方式", StringComparison.OrdinalIgnoreCase))
                            {
                                tplSheet = sh;
                                break;
                            }
                        }

                        // 若母版中未包含【尺寸安装方式】，自愈以【成品交接单】为基础克隆
                        if (tplSheet == null)
                        {
                            foreach (dynamic sh in templateWb.Worksheets)
                            {
                                if (string.Equals(Convert.ToString(sh.Name)?.Trim(), "成品交接单", StringComparison.OrdinalIgnoreCase))
                                {
                                    tplSheet = sh;
                                    needAdjustHeaders = true;
                                    break;
                                }
                            }
                        }

                        if (tplSheet == null)
                        {
                            result.Message = "母版 CabinetTemplate.xlsx 中未包含可用模板工作表！";
                            return result;
                        }

                        // 复制模板工作表至当前算价工作簿最末尾
                        tplSheet.Copy(After: activeWb.Sheets[activeWb.Sheets.Count]);
                        targetSheet = activeWb.Sheets[activeWb.Sheets.Count];
                        try { targetSheet.Name = "尺寸安装方式"; } catch { }
                    }
                    finally
                    {
                        // 复制完成后立即关闭模板工作簿句柄
                        try { templateWb.Close(false); } catch { }
                    }

                    // 8. 若自愈克隆自成品交接单，调整标题与表头
                    if (needAdjustHeaders)
                    {
                        targetSheet.Range["A1"].Value2 = "附件3:";
                        targetSheet.Range["A3"].Value2 = "尺寸安装方式";
                        targetSheet.Range["F6"].Value2 = "宽高深";
                        targetSheet.Range["G6"].Value2 = "安装方式";
                    }

                    // 9. 替换表头信息占位符 (第 1~5 行)
                    try
                    {
                        // A2 单元格：公司名称替换 (为空置空占位符)
                        string a2Val = Convert.ToString(targetSheet.Range["A2"].Value2) ?? "";
                        string finalCompanyText = !string.IsNullOrWhiteSpace(companyName) ? companyName : "";
                        if (a2Val.Contains("[公司名称]"))
                        {
                            targetSheet.Range["A2"].Value2 = a2Val.Replace("[公司名称]", finalCompanyText);
                        }
                        else if (a2Val.Contains("[[公司名称]]"))
                        {
                            targetSheet.Range["A2"].Value2 = a2Val.Replace("[[公司名称]]", finalCompanyText);
                        }
                        else if (!string.IsNullOrWhiteSpace(finalCompanyText))
                        {
                            targetSheet.Range["A2"].Value2 = finalCompanyText;
                        }

                        // A4 单元格：单位名称 (客户名称) - 用户明确指令：[客户名称]为空也需要置空
                        string a4Val = Convert.ToString(targetSheet.Range["A4"].Value2) ?? "";
                        string finalCustomerText = !string.IsNullOrWhiteSpace(customerName) ? customerName : "";
                        if (a4Val.Contains("[客户名称]"))
                        {
                            targetSheet.Range["A4"].Value2 = a4Val.Replace("[客户名称]", finalCustomerText);
                        }
                        else if (a4Val.Contains("[[客户名称]]"))
                        {
                            targetSheet.Range["A4"].Value2 = a4Val.Replace("[[客户名称]]", finalCustomerText);
                        }
                        else
                        {
                            // 若模板不含占位符 (如直接为写死的旧企业文本)，按是否有客户名称规范写回
                            targetSheet.Range["A4"].Value2 = string.IsNullOrWhiteSpace(finalCustomerText) ? "单位名称: " : $"单位名称: {finalCustomerText}";
                        }

                        // A5 单元格：项目名称替换 (为空置空占位符)
                        string a5Val = Convert.ToString(targetSheet.Range["A5"].Value2) ?? "";
                        string finalProjectText = !string.IsNullOrWhiteSpace(projectName) ? projectName : "";
                        if (a5Val.Contains("[[项目名称]]"))
                        {
                            targetSheet.Range["A5"].Value2 = a5Val.Replace("[[项目名称]]", finalProjectText);
                        }
                        else if (a5Val.Contains("[项目名称]"))
                        {
                            targetSheet.Range["A5"].Value2 = a5Val.Replace("[项目名称]", finalProjectText);
                        }
                        else
                        {
                            targetSheet.Range["A5"].Value2 = string.IsNullOrWhiteSpace(finalProjectText) ? "项目名称: " : $"项目名称: {finalProjectText}";
                        }
                    }
                    catch (Exception exHeader)
                    {
                        LogHelper.WriteLog($"[尺寸安装方式] 表头信息替换容错: {exHeader.Message}");
                    }

                    // 10. 自适应排版扩展 (模板预留 6 行: 7~12 行)
                    int itemCount = dimensionsItems.Count;
                    int startRow = 7; // 首个数据行 --硬编码: 首行--
                    int defaultTemplateRows = 6; // 默认预留 6 行 --硬编码: 预留行数--

                    if (itemCount > defaultTemplateRows)
                    {
                        int needInsert = itemCount - defaultTemplateRows;
                        int insertStartRow = startRow + defaultTemplateRows; // 第 13 行 (合计行)
                        int insertEndRow = insertStartRow + needInsert - 1;

                        // 整块向下推挤插入空行 (xlShiftDown = -4121)
                        dynamic insertRows = targetSheet.Rows[$"{insertStartRow}:{insertEndRow}"];
                        insertRows.Insert(-4121);

                        // 复制第 7 行单元格格式 (xlPasteFormats = -4122)
                        targetSheet.Rows[7].Copy();
                        dynamic formatTarget = targetSheet.Rows[$"{insertStartRow}:{insertEndRow}"];
                        formatTarget.PasteSpecial(-4122);
                        app.CutCopyMode = 0;
                    }

                    // 计算数据实际末行与合计行行号
                    int endDataRow = startRow + Math.Max(itemCount, defaultTemplateRows) - 1;
                    int actualItemEndRow = startRow + itemCount - 1;
                    int totalRow = endDataRow + 1; // 合计行物理行号

                    // 11. 规则 7: 构造数据写入二维矩阵并注入超链接与各列数据
                    // 矩阵维度: actualItemCount 行 x 9 列 (A~I 列)
                    object[,] writeMatrix = new object[itemCount, 9];

                    for (int i = 0; i < itemCount; i++)
                    {
                        var it = dimensionsItems[i];

                        // A 列 (序号 + 超链接公式，直达所属分类表该箱柜明细行)
                        string linkFormula = $"=HYPERLINK(\"#'{it.SheetName}'!A{it.DetailStartRow}:I{it.DetailEndRow}\", {it.Index})";
                        writeMatrix[i, 0] = linkFormula;

                        // B 列 (名称: 配电箱或提取名称)
                        writeMatrix[i, 1] = !string.IsNullOrWhiteSpace(it.Name) ? it.Name : "配电箱";

                        // C 列 (型 号: 对应模板[柜号])
                        writeMatrix[i, 2] = it.CabinetCode;

                        // D 列 (数量: 台数数值)
                        writeMatrix[i, 3] = it.Quantity;

                        // E 列 (单位: 台)
                        writeMatrix[i, 4] = it.Unit;

                        // F 列 (宽高深: 对应模板[箱柜尺寸])
                        writeMatrix[i, 5] = it.Dimensions;

                        // G 列 (安装方式: 对应模板[箱柜型号]，取 D 列数据)
                        writeMatrix[i, 6] = it.InstallMode;

                        // H 列 (防护等级: 统一填默认值 IP30)
                        writeMatrix[i, 7] = it.ProtectionLevel;

                        // I 列 (备注: 汇总行 I 列)
                        writeMatrix[i, 8] = it.Remark;
                    }

                    // 规则 7: 一次性批量写入 A~I 列真实数据与超链接
                    dynamic targetDataRange = targetSheet.Range[$"A{startRow}:I{actualItemEndRow}"];
                    targetDataRange.Formula = writeMatrix;

                    // 若实际箱柜数量少于预留行数，清空多余预留行数据 (保留边框网格)
                    if (itemCount < defaultTemplateRows)
                    {
                        int emptyStartRow = actualItemEndRow + 1;
                        dynamic emptyRange = targetSheet.Range[$"A{emptyStartRow}:I{endDataRow}"];
                        emptyRange.ClearContents();
                    }

                    // 12. 刷新合计行公式与静态值
                    try
                    {
                        // D 列数量合计公式
                        targetSheet.Cells[totalRow, 4].Formula = $"=SUM(D{startRow}:D{actualItemEndRow})";
                        // E 列单位标注为“台”
                        targetSheet.Cells[totalRow, 5].Value2 = "台";
                    }
                    catch (Exception exTotal)
                    {
                        LogHelper.WriteLog($"[尺寸安装方式] 刷新合计行公式容错: {exTotal.Message}");
                    }

                    // 13. 激活新生成的尺寸安装方式表为当前活动主视图
                    try
                    {
                        targetSheet.Activate();
                        targetSheet.Range["A1"].Select();
                    }
                    catch { }

                    // 标记导出成功
                    result.Success = true;
                    result.CabinetCount = dimensionsItems.Count;
                    result.Message = $"已成功导出成套设备【尺寸安装方式】表，共计 {dimensionsItems.Count} 台箱柜。";
                }
                finally
                {
                    // 恢复 Excel 屏幕刷新、事件响应与警告提示
                    app.ScreenUpdating = true;
                    app.EnableEvents = true;
                    app.DisplayAlerts = true;
                }
            }
            catch (Exception ex)
            {
                // 全局未捕获异常记录
                result.Success = false;
                result.Message = $"导出尺寸安装方式表发生未捕获异常: {ex.Message}";
                LogHelper.WriteLog($"[尺寸安装方式] 核心导出发生异常: {ex}");
            }

            return result;
        }
    }
}
