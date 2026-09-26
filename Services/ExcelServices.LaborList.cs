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
    /// Excel 业务服务分部类：成套设备装配【人工清单】自动导出与分类表选择
    /// </summary>
    public static partial class ExcelServices
    {
        // 声明人工清单分类选择向导窗口静态单例引用
        private static Forms.LaborListForm? _laborListForm;

        /// <summary>
        /// 扫描当前活动工作簿，获取所有有效分类明细表及箱柜统计，同时检测是否已存在人工清单工作表
        /// 供前端向导界面多选框展示、或单表极简一键直达使用
        /// </summary>
        /// <returns>人工清单初始化数据模型</returns>
        public static LaborListInitDataDto GetLaborListInitData()
        {
            // 初始化返回的初始化数据 DTO
            var initData = new LaborListInitDataDto();

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

                    // 检查是否已包含【人工清单】工作表 --硬编码: 系统工作表名称--
                    if (string.Equals(sName, "人工清单", StringComparison.OrdinalIgnoreCase))
                    {
                        initData.HasExistingLaborList = true;
                    }

                    // 尝试从【项目信息】工作表 B5 单元格提取工程项目名称 --硬编码: B5 项目名称单元格--
                    if (string.Equals(sName, "项目信息", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            string pName = Convert.ToString(sheet.Range["B5"].Value2)?.Trim() ?? "";
                            if (!string.IsNullOrWhiteSpace(pName)) projectName = pName;
                        }
                        catch { }
                    }

                    // 排除系统保留与非分类工作表 --硬编码: 系统保留表名--
                    if (string.Equals(sName, "项目信息", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "采购清单", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "领料清单", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "人工清单", StringComparison.OrdinalIgnoreCase) ||
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
                        initData.Categories.Add(new LaborListCategoryDto
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
                LogHelper.WriteLog($"[人工清单] 初始化数据获取异常: {ex.Message}");
            }

            return initData;
        }

        /// <summary>
        /// 功能区 Ribbon 入口：根据分类表数量智能分流处理
        /// 若仅有 1 个分类表，则一键极简直达；若有多个分类表，则弹出选择向导
        /// </summary>
        public static void ShowLaborListDialogOrExport()
        {
            try
            {
                // 获取当前工作簿的分类表初始化数据
                var initData = GetLaborListInitData();

                // 校验是否存在有效分类表
                if (initData.Categories.Count == 0)
                {
                    MessageBox.Show("当前工作簿未检测到任何包含箱柜的有效分类工作表！\n请先新建分类并添加箱柜后再导出人工清单。",
                        "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 场景 1: 若当前仅有 1 个分类表
                if (initData.Categories.Count == 1)
                {
                    // 若已存在人工清单，采用弹窗防误触提醒确认
                    if (initData.HasExistingLaborList)
                    {
                        var confirmResult = MessageBox.Show(
                            "当前工作簿已存在【人工清单】工作表！\n重新生成将覆盖并重置该表，是否继续执行？",
                            "覆盖确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (confirmResult != DialogResult.Yes) return;
                    }

                    // 单表场景下，直接取该分类表名称一键秒级导出
                    var selectedSheets = new List<string> { initData.Categories[0].SheetName };
                    var exportResult = ExportLaborListToCurrentWorkbook(selectedSheets);

                    if (exportResult.Success)
                    {
                        MessageBox.Show($"成套设备【人工清单】导出成功！\n共计生成 {exportResult.CabinetCount} 台箱柜，工费合计 {exportResult.TotalAmount:N0} 元。",
                            "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show($"导出人工清单失败: {exportResult.Message}",
                            "导出失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return;
                }

                // 场景 2: 若包含多个分类表，弹出优雅的 WebView2 + Vue 3 选择向导窗口
                ShowModelessForm(ref _laborListForm, () => new Forms.LaborListForm());
            }
            catch (Exception ex)
            {
                // 顶层异常捕获与提示
                MessageBox.Show($"弹出人工清单功能失败: {ex.Message}", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 核心导出服务：基于选中的一个或多个分类表，生成全新的【人工清单】标准工作表
        /// 纯静默执行，零 Win32 阻塞弹窗，避免 WebView2 IPC 假死
        /// </summary>
        /// <param name="selectedCategorySheets">选中的分类工作表名称列表</param>
        /// <returns>导出执行结果模型</returns>
        public static LaborListExportResult ExportLaborListToCurrentWorkbook(List<string> selectedCategorySheets)
        {
            // 初始化返回的执行结果模型
            var result = new LaborListExportResult();

            // 前置入参安全防御校验
            if (selectedCategorySheets == null || selectedCategorySheets.Count == 0)
            {
                result.Message = "未选择任何需要导出人工清单的分类工作表！";
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

                // 3. 提取项目名称 (优先从【项目信息】表 B5 单元格提取)
                string projectName = Path.GetFileNameWithoutExtension(Convert.ToString(activeWb.Name) ?? "");
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    if (string.Equals(Convert.ToString(ws.Name)?.Trim(), "项目信息", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            string pName = Convert.ToString(ws.Range["B5"].Value2)?.Trim() ?? "";
                            if (!string.IsNullOrWhiteSpace(pName)) projectName = pName;
                        }
                        catch { }
                        break;
                    }
                }

                // 4. 遍历选中的分类表，提取所有箱柜及计费区人工费信息
                var laborItems = new List<LaborListItem>();
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
                        var anchor = kvp.Value;
                        if (anchor == null) continue;
                        globalIndex++;

                        // 安全提取箱柜 4 个关键定义名称锚点的物理行号
                        int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : 0;
                        int detRow = anchor.Det != null ? Convert.ToInt32(anchor.Det.Row) : 0;
                        int subsumRow = anchor.Subsum != null ? Convert.ToInt32(anchor.Subsum.Row) : 0;
                        int tolsumRow = anchor.Tolsum != null ? Convert.ToInt32(anchor.Tolsum.Row) : 0;

                        // 准备箱柜项属性
                        string cabName = "配电箱"; // 默认箱柜名称 --硬编码: 默认箱柜名称--
                        string cabModel = string.Empty; // 柜号/型号
                        decimal cabQty = 1; // 箱柜数量

                        // 优先从汇总行 (Cab_Sum_k) 提取基础信息
                        if (sumRow > 0)
                        {
                            try
                            {
                                // 一次性读取汇总行 A:F 列
                                dynamic sumRange = sheet.Range[$"A{sumRow}:F{sumRow}"];
                                object[,] sumMatrix = (object[,])sumRange.Value2;

                                // B 列为柜号
                                string bVal = Convert.ToString(sumMatrix[1, 2])?.Trim() ?? "";
                                // C 列为箱柜名称
                                string cVal = Convert.ToString(sumMatrix[1, 3])?.Trim() ?? "";
                                // D 列为箱柜型号
                                string dVal = Convert.ToString(sumMatrix[1, 4])?.Trim() ?? "";
                                // F 列为数量
                                object fObj = sumMatrix[1, 6];

                                if (!string.IsNullOrWhiteSpace(cVal)) cabName = cVal;

                                // 柜号与型号拼接，优先保证柜号完整
                                if (!string.IsNullOrWhiteSpace(bVal) && !string.IsNullOrWhiteSpace(dVal))
                                {
                                    cabModel = $"{bVal} {dVal}".Trim();
                                }
                                else if (!string.IsNullOrWhiteSpace(bVal))
                                {
                                    cabModel = bVal;
                                }
                                else if (!string.IsNullOrWhiteSpace(dVal))
                                {
                                    cabModel = dVal;
                                }

                                if (fObj != null && decimal.TryParse(fObj.ToString(), out decimal parsedQty) && parsedQty > 0)
                                {
                                    cabQty = parsedQty;
                                }
                            }
                            catch (Exception exSum)
                            {
                                LogHelper.WriteLog($"[人工清单] 提取箱柜 {kvp.Key} 汇总行容错: {exSum.Message}");
                            }
                        }

                        // 若汇总行未能提取到完整信息，从明细表信息行 (Cab_Det_k) 补充
                        if (detRow > 0 && string.IsNullOrWhiteSpace(cabModel))
                        {
                            try
                            {
                                // 读取明细表信息行 A:G 列
                                dynamic detRange = sheet.Range[$"A{detRow}:G{detRow}"];
                                object[,] detMatrix = (object[,])detRange.Value2;

                                string bVal = Convert.ToString(detMatrix[1, 2])?.Trim() ?? "";
                                if (!string.IsNullOrWhiteSpace(bVal)) cabModel = bVal;
                            }
                            catch { }
                        }

                        // 5. 从明细表底部的计费区域提取【人工费】金额
                        // 计费区域范围: subsumRow + 1 到 tolsumRow - 1 (规则 6)
                        decimal totalLaborAmount = 0;
                        int feeStartRow = subsumRow + 1;
                        int feeEndRow = tolsumRow - 1;

                        if (feeEndRow >= feeStartRow && feeStartRow > 0)
                        {
                            try
                            {
                                // 一次性读取计费区域 B 列至 K 列 (规则 7 批量读入内存)
                                dynamic feeRange = sheet.Range[$"B{feeStartRow}:K{feeEndRow}"];
                                object[,] feeMatrix = (object[,])feeRange.Value2;
                                int feeRowCount = feeMatrix.GetLength(0);

                                for (int r = 1; r <= feeRowCount; r++)
                                {
                                    // 第 1 列对应 B 列 (费用项名称)
                                    string feeItemName = Convert.ToString(feeMatrix[r, 1])?.Trim() ?? "";

                                    // 模糊匹配包含“人工”两个字的行 (例如: 人工费、配线人工、成套制作人工)
                                    if (feeItemName.Contains("人工"))
                                    {
                                        // 依次探测 H 列(第 7 列), G 列(第 6 列), K 列(第 10 列)
                                        decimal foundVal = 0;
                                        int[] candidateCols = new int[] { 7, 6, 10 }; // 对应 H, G, K 列 --硬编码: 费用候选金额列--

                                        foreach (int colIdx in candidateCols)
                                        {
                                            object valObj = feeMatrix[r, colIdx];
                                            if (valObj != null && decimal.TryParse(valObj.ToString(), out decimal parsedVal) && parsedVal > 0)
                                            {
                                                foundVal = parsedVal;
                                                break;
                                            }
                                        }

                                        // 若找到有效人工费数值，记录并跳出
                                        if (foundVal > 0)
                                        {
                                            totalLaborAmount = foundVal;
                                            break;
                                        }
                                    }
                                }
                            }
                            catch (Exception exFee)
                            {
                                LogHelper.WriteLog($"[人工清单] 提取箱柜 {kvp.Key} 计费区人工费容错: {exFee.Message}");
                            }
                        }

                        // 6. 计算单台人工费 (选项 A: 直接取明细原值，若为成套总费用则除以台数)
                        decimal unitLaborPrice = 0;
                        if (totalLaborAmount > 0)
                        {
                            // 单台人工单价 = 总人工费 / 台数，四舍五入取整对齐工程实际
                            unitLaborPrice = cabQty > 0 ? Math.Round(totalLaborAmount / cabQty, 0) : totalLaborAmount;
                        }

                        // 组装箱柜人工清单项
                        laborItems.Add(new LaborListItem
                        {
                            Index = globalIndex,
                            Name = cabName,
                            Model = cabModel,
                            Quantity = cabQty,
                            Unit = "台",
                            UnitPrice = unitLaborPrice,
                            TotalPrice = unitLaborPrice * cabQty,
                            SheetName = sheetName,
                            DetailStartRow = detRow > 0 ? detRow : (sumRow > 0 ? sumRow : 1),
                            DetailEndRow = tolsumRow > 0 ? tolsumRow : (subsumRow > 0 ? subsumRow : (detRow > 0 ? detRow : 1))
                        });
                    }
                }

                // 校验提取结果有效性
                if (laborItems.Count == 0)
                {
                    result.Message = "选中的分类表中未提取到任何有效的箱柜数据！";
                    return result;
                }

                // 7. 挂起屏幕刷新、事件与系统提示，杜绝事件交叉死锁
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.DisplayAlerts = false;

                try
                {
                    // 8. 查找已有人工清单并安全删除旧表 (纯净克隆重置机制)
                    dynamic? laborSheet = null;
                    foreach (dynamic ws in activeWb.Worksheets)
                    {
                        if (string.Equals(Convert.ToString(ws.Name)?.Trim(), "人工清单", StringComparison.OrdinalIgnoreCase))
                        {
                            laborSheet = ws;
                            break;
                        }
                    }

                    if (laborSheet != null)
                    {
                        try
                        {
                            laborSheet.Delete();
                            laborSheet = null;
                        }
                        catch (Exception exDel)
                        {
                            LogHelper.WriteLog($"[人工清单] 移除旧人工清单工作表容错: {exDel.Message}");
                        }
                    }

                    // 9. 打开标准母版克隆纯净的标准【人工清单】工作表
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
                            if (string.Equals(Convert.ToString(sh.Name)?.Trim(), "人工清单", StringComparison.OrdinalIgnoreCase))
                            {
                                tplSheet = sh;
                                break;
                            }
                        }

                        if (tplSheet == null)
                        {
                            result.Message = "母版 CabinetTemplate.xlsx 中未包含【人工清单】标准工作表！";
                            return result;
                        }

                        // 复制模板工作表至当前算价工作簿最末尾
                        tplSheet.Copy(After: activeWb.Sheets[activeWb.Sheets.Count]);
                        laborSheet = activeWb.Sheets[activeWb.Sheets.Count];
                        try { laborSheet.Name = "人工清单"; } catch { }
                    }
                    finally
                    {
                        // 启发式规则 12: 复制完成后立即关闭模板工作簿句柄
                        try { templateWb.Close(false); } catch { }
                    }

                    // 10. 解析模板第 7 行占位符与列定义 (自适应零硬编码)
                    var colMap = new LaborListTemplateColumnMap();
                    try
                    {
                        dynamic row7Range = laborSheet.Range["A7:H7"];
                        object[,] row7Matrix = (object[,])row7Range.Value2;

                        for (int col = 1; col <= 8; col++)
                        {
                            string cellText = Convert.ToString(row7Matrix[1, col])?.Trim() ?? "";

                            if (cellText.Contains("[箱柜序号]") || cellText.Contains("[序号]"))
                            {
                                colMap.IndexCol = col;
                            }
                            else if (cellText.Contains("[箱柜名称]") || cellText.Contains("[名称]"))
                            {
                                colMap.NameCol = col;
                            }
                            else if (cellText.Contains("[柜号]") || cellText.Contains("[型号]"))
                            {
                                colMap.ModelCol = col;
                            }
                            else if (cellText.Contains("[数量]"))
                            {
                                colMap.QuantityCol = col;
                            }
                            else if (cellText.Contains("[单位]") || cellText.Equals("台"))
                            {
                                colMap.UnitCol = col;
                            }
                            else if (cellText.Contains("[人工费]") || cellText.Contains("[单价]"))
                            {
                                colMap.UnitPriceCol = col;
                            }
                        }
                    }
                    catch (Exception exMap)
                    {
                        LogHelper.WriteLog($"[人工清单] 模板列动态解析容错: {exMap.Message}");
                    }

                    // 11. 替换表头项目名称占位符 (A5 单元格)
                    try
                    {
                        string a5Val = Convert.ToString(laborSheet.Range["A5"].Value2) ?? "";
                        if (a5Val.Contains("[[项目名称]]"))
                        {
                            laborSheet.Range["A5"].Value2 = a5Val.Replace("[[项目名称]]", projectName);
                        }
                        else if (a5Val.Contains("[项目名称]"))
                        {
                            laborSheet.Range["A5"].Value2 = a5Val.Replace("[项目名称]", projectName);
                        }
                    }
                    catch (Exception exA5)
                    {
                        LogHelper.WriteLog($"[人工清单] 替换项目名称容错: {exA5.Message}");
                    }

                    // 12. 自适应排版扩展 (模板预留 6 行: 7~12 行)
                    int itemCount = laborItems.Count;
                    int startRow = 7; // 首个数据行 --硬编码: 首行--
                    int defaultTemplateRows = 6; // 默认预留 6 行 --硬编码: 预留行数--

                    if (itemCount > defaultTemplateRows)
                    {
                        int needInsert = itemCount - defaultTemplateRows;
                        int insertStartRow = startRow + defaultTemplateRows; // 第 13 行 (合计行)
                        int insertEndRow = insertStartRow + needInsert - 1;

                        // 整块向下推挤插入空行 (xlShiftDown = -4121)
                        dynamic insertRows = laborSheet.Rows[$"{insertStartRow}:{insertEndRow}"];
                        insertRows.Insert(-4121);

                        // 复制第 7 行单元格格式 (xlPasteFormats = -4122)
                        laborSheet.Rows[7].Copy();
                        dynamic formatTarget = laborSheet.Rows[$"{insertStartRow}:{insertEndRow}"];
                        formatTarget.PasteSpecial(-4122);
                        app.CutCopyMode = 0;
                    }

                    // 计算数据实际末行与合计行行号
                    int endDataRow = startRow + Math.Max(itemCount, defaultTemplateRows) - 1;
                    int actualItemEndRow = startRow + itemCount - 1;
                    int totalRow = endDataRow + 1; // 合计行物理行号
                    int upperRow = totalRow + 1; // 大写金额行物理行号

                    // 13. 构造数据写入二维矩阵并注入超链接与算价公式
                    // 矩阵大小: actualItemCount 行 x 8 列
                    object[,] writeMatrix = new object[itemCount, 8];

                    for (int i = 0; i < itemCount; i++)
                    {
                        var it = laborItems[i];
                        int curRow = startRow + i;

                        // A 列 (序号 + 超链接公式，选项 B)
                        // 例如: =HYPERLINK("#'分类1'!A44:H70", 1)
                        string linkFormula = $"=HYPERLINK(\"#'{it.SheetName}'!A{it.DetailStartRow}:H{it.DetailEndRow}\", {it.Index})";
                        writeMatrix[i, 0] = linkFormula;

                        // B 列 (名称: 配电箱)
                        writeMatrix[i, 1] = it.Name;

                        // C 列 (柜号/型号)
                        writeMatrix[i, 2] = it.Model;

                        // D 列 (数量台数)
                        writeMatrix[i, 3] = it.Quantity;

                        // E 列 (单位: 台)
                        writeMatrix[i, 4] = it.Unit;

                        // F 列 (单价: 人工费单价)
                        writeMatrix[i, 5] = it.UnitPrice;

                        // G 列 (金额公式: =F*D)
                        writeMatrix[i, 6] = $"=F{curRow}*D{curRow}";

                        // H 列 (备注: 留空)
                        writeMatrix[i, 7] = it.Remark;
                    }

                    // 规则 7: 内存二维矩阵一次性回写数据与公式 (通过 Formula 属性保证公式生效)
                    dynamic targetDataRange = laborSheet.Range[laborSheet.Cells[startRow, 1], laborSheet.Cells[actualItemEndRow, 8]];
                    targetDataRange.Formula = writeMatrix;

                    // 若实际数据小于预留的 6 行，清空剩余预留行的占位符内容，保持网格线纯净
                    if (itemCount < defaultTemplateRows)
                    {
                        int cleanStartRow = actualItemEndRow + 1;
                        dynamic cleanRange = laborSheet.Range[laborSheet.Cells[cleanStartRow, 1], laborSheet.Cells[endDataRow, 8]];
                        cleanRange.ClearContents();
                    }

                    // 14. 批量设置行高、边框与文本对齐样式
                    try
                    {
                        dynamic allDataRange = laborSheet.Range[laborSheet.Cells[startRow, 1], laborSheet.Cells[endDataRow, 8]];
                        allDataRange.Borders.LineStyle = 1; // 细实线边框
                        allDataRange.Borders.Weight = 2;
                        allDataRange.RowHeight = 28; // 舒适行高

                        // 居中对齐列: A 列(序号), D 列(数量), E 列(单位) (xlCenter = -4108)
                        laborSheet.Range[laborSheet.Cells[startRow, 1], laborSheet.Cells[endDataRow, 1]].HorizontalAlignment = -4108;
                        laborSheet.Range[laborSheet.Cells[startRow, 4], laborSheet.Cells[endDataRow, 4]].HorizontalAlignment = -4108;
                        laborSheet.Range[laborSheet.Cells[startRow, 5], laborSheet.Cells[endDataRow, 5]].HorizontalAlignment = -4108;
                        // 右对齐列: F 列(单价), G 列(金额) (xlRight = -4152)
                        laborSheet.Range[laborSheet.Cells[startRow, 6], laborSheet.Cells[endDataRow, 6]].HorizontalAlignment = -4152;
                        laborSheet.Range[laborSheet.Cells[startRow, 7], laborSheet.Cells[endDataRow, 7]].HorizontalAlignment = -4152;
                    }
                    catch (Exception exStyle)
                    {
                        LogHelper.WriteLog($"[人工清单] 样式设置容错: {exStyle.Message}");
                    }

                    // 15. 自适应刷新合计行与大写 RMB 金额公式
                    try
                    {
                        // D 列数量合计公式: =SUM(D6:D{endDataRow})
                        laborSheet.Range[$"D{totalRow}"].Formula = $"=SUM(D6:D{endDataRow})";

                        // G 列金额合计公式: =ROUND(SUM(G6:G{endDataRow}),0)
                        laborSheet.Range[$"G{totalRow}"].Formula = $"=ROUND(SUM(G6:G{endDataRow}),0)";

                        // A 列大写人民币联动公式更新 (自适应引用当前合计行单元格 G{totalRow})
                        string rmbFormula = $"=\"合计RMB：\"& TEXT(SUBSTITUTE(IF(G{totalRow}=\"\",\"\",IF(MOD(G{totalRow},1)=0,TEXT(INT(G{totalRow}),\"[DBNum2][$-804]G/通用格式元整\"),(TEXT(INT(G{totalRow}),\"[DBNum2][$-804]G/通用格式元\")&TEXT((INT(G{totalRow}*10)-INT(G{totalRow})*10),\"[DBNum2][$-804]G/通用格式角\")&TEXT((INT(G{totalRow}*100)-INT(G{totalRow}*10)*10),\"[DBNum2][$-804]G/通用格式分\")))),\"零角\",\"零\"),)";
                        laborSheet.Range[$"A{upperRow}"].Formula = rmbFormula;
                    }
                    catch (Exception exFormula)
                    {
                        LogHelper.WriteLog($"[人工清单] 合计公式自适应更新容错: {exFormula.Message}");
                    }

                    // 16. 平滑聚焦与激活人工清单工作表
                    try
                    {
                        laborSheet.Activate();
                        laborSheet.Range["A1"].Select();
                    }
                    catch { }

                    // 回填执行成功结果
                    result.Success = true;
                    result.CabinetCount = itemCount;
                    result.TotalAmount = laborItems.Sum(x => x.TotalPrice);
                    result.Message = "成套设备【人工清单】已成功生成并完成排版！";
                }
                finally
                {
                    // 恢复 Excel 屏幕更新与事件监听
                    app.CutCopyMode = 0;
                    app.ScreenUpdating = true;
                    app.EnableEvents = true;
                    app.DisplayAlerts = true;
                }
            }
            catch (Exception exGlobal)
            {
                // 全局未捕获异常记录
                result.Success = false;
                result.Message = $"生成人工清单过程发生异常: {exGlobal.Message}";
                LogHelper.WriteLog($"[人工清单] 严重异常: {exGlobal}");
            }

            return result;
        }
    }
}
