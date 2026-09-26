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
    /// Excel 业务服务分部类：材料统计与采购清单自动导出
    /// </summary>
    public static partial class ExcelServices
    {
        // 声明材料统计分类选择向导窗口静态单例引用
        private static Forms.MaterialStatForm? _materialStatForm;

        /// <summary>
        /// 弹出基于 WebView2 + Vue 3 + Element Plus 的“材料统计与分类选择”向导窗口
        /// </summary>
        public static void ShowMaterialStatDialog()
        {
            try
            {
                // 以非模态方式展示向导窗口，确保 Excel 处于可交互状态
                ShowModelessForm(ref _materialStatForm, () => new Forms.MaterialStatForm());
            }
            catch (Exception ex)
            {
                // 异常提示
                MessageBox.Show($"弹出材料统计向导窗口失败: {ex.Message}", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 扫描当前活动工作簿，获取所有有效分类明细表及箱柜统计，同时检测是否已存在采购清单工作表
        /// 供前端向导界面复选框动态展示与覆盖检测使用
        /// </summary>
        /// <returns>材料统计初始化数据模型</returns>
        public static MaterialStatInitDataDto GetMaterialStatInitData()
        {
            // 初始化返回的初始化数据 DTO
            var initData = new MaterialStatInitDataDto();

            try
            {
                // 获取 Excel 宿主应用程序实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return initData;

                // 获取活动算价工作簿
                dynamic? activeWb = app.ActiveWorkbook;
                if (activeWb == null) return initData;

                // 遍历活动工作簿下的所有工作表
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    // 获取工作表名称并清洗两端空白
                    string sName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;

                    // 检查是否已包含【采购清单】工作表 --硬编码: 系统工作表名称--
                    if (string.Equals(sName, "采购清单", StringComparison.OrdinalIgnoreCase))
                    {
                        initData.HasExistingPurchaseList = true;
                    }

                    // 排除系统保留工作表 --硬编码: 系统保留表名--
                    if (string.Equals(sName, "项目信息", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "采购清单", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "领料清单", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "材料分布表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "元件汇总分布表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "元件汇总表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "屏柜汇总表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "元器件数据管理", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // 白名单安全守门校验
                    if (!Tool.IsProjectCategorySheet(sheet)) continue;

                    // 规则 8: 校验前确保规则 6 拓扑健康
                    Tool.FixAndFillCabinetNamesForSheet(sheet);

                    // 提取当前分类表下有效箱柜数量
                    var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);
                    int cabCount = validCabinets != null ? validCabinets.Count : 0;

                    // 添加至返回列表中
                    initData.Categories.Add(new MaterialStatCategoryDto
                    {
                        SheetName = sName,
                        CabinetCount = cabCount,
                        IsSelected = true
                    });
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[GetMaterialStatInitData] 获取分类初始化数据异常: {ex.Message}");
            }

            return initData;
        }

        /// <summary>
        /// 扫描当前活动工作簿，获取所有有效分类明细表及对应箱柜数量统计
        /// 供兼容调用
        /// </summary>
        /// <returns>分类工作表信息列表</returns>
        public static List<MaterialStatCategoryDto> GetMaterialStatCategories()
        {
            // 复用统一初始化扫描并返回分类列表
            return GetMaterialStatInitData().Categories;
        }

        /// <summary>
        /// 自动扫描当前项目所选分类箱柜元器件，严格只按【名称+型号规格】综合累计采购总数量，
        /// 并在当前工作簿生成/刷新【采购清单】工作表。遵循规则 6、7、8，零硬编码自适应模板占位符
        /// </summary>
        /// <param name="selectedCategories">用户勾选的分类工作表名称列表，若为空则全量汇总</param>
        /// <param name="sortBy">排序规则: brand_name_model (默认: 品牌+名称+型号) 或 name_model (名称+型号)</param>
        /// <param name="isSilent">是否静默导出 (默认为 true，零 MessageBox 阻塞)</param>
        /// <returns>操作执行结果对象</returns>
        public static PurchaseListExportResult ExportPurchaseListToCurrentWorkbook(List<string>? selectedCategories = null, string sortBy = "brand_name_model", bool isSilent = true)
        {
            // 初始化返回结果对象
            var result = new PurchaseListExportResult();

            try
            {
                // 1. 安全获取当前 Excel 宿主应用程序实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    // 无法连接 Excel 记录错误并直接返回 (绝不调用 MessageBox.Show 阻塞 IPC)
                    result.Message = "无法连接至 Excel 宿主应用程序！";
                    return result;
                }

                // 2. 获取当前活动算价工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null)
                {
                    // 无活动工作簿记录警告并直接返回
                    result.Message = "当前没有打开的活动工作簿，无法执行材料统计！";
                    return result;
                }

                // 3. 提取项目基础信息 (公司名称与项目名称)
                string projectName = Path.GetFileNameWithoutExtension(Convert.ToString(activeWb.Name) ?? "");
                string companyName = string.Empty;

                try
                {
                    // 尝试从【项目信息】工作表读取工程信息 --硬编码: 工作表名称--
                    dynamic? infoSheet = null;
                    try { infoSheet = activeWb.Sheets["项目信息"]; } catch { }

                    if (infoSheet != null)
                    {
                        // 规则 7: 一次性读取 A1:C30 区域至内存二维数组
                        dynamic infoRange = infoSheet.Range["A1:C30"];
                        object[,] infoMatrix = (object[,])infoRange.Value2;

                        // 提取项目名称 (Cell B5)
                        string pName = Convert.ToString(infoMatrix[5, 2])?.Trim() ?? "";
                        if (!string.IsNullOrWhiteSpace(pName)) projectName = pName;

                        // 提取公司名称 (优先 Cell B22 供货单位，次选 Cell B1 公司名称)
                        string comp = Convert.ToString(infoMatrix[22, 2])?.Trim() ?? "";
                        if (string.IsNullOrWhiteSpace(comp)) comp = Convert.ToString(infoMatrix[1, 2])?.Trim() ?? "";
                        companyName = comp;
                    }
                }
                catch (Exception exInfo)
                {
                    // 记录读取项目信息微小异常但继续流程
                    LogHelper.WriteLog($"[采购清单] 读取项目信息容错: {exInfo.Message}");
                }

                // 构建所选分类快速查找哈希表
                HashSet<string>? selectedSheetSet = null;
                if (selectedCategories != null && selectedCategories.Count > 0)
                {
                    selectedSheetSet = new HashSet<string>(selectedCategories.Select(s => s.Trim()), StringComparer.OrdinalIgnoreCase);
                }

                // 4. 遍历分类工作表，聚合元器件采购总数量
                // 重点：严格按照用户最新明确指示：汇总只按照【型号+名称】汇总，不根据品牌与单位做隔离区分
                // 聚合字典 Key: 名称___型号 (大小写不敏感)
                var aggregatedDict = new Dictionary<string, PurchaseListItem>(StringComparer.OrdinalIgnoreCase);

                // 遍历活动工作簿下的所有工作表
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    // 获取工作表名称并清洗两端空白
                    string sName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;

                    // 精准排除系统保留工作表与非分类报表 --硬编码: 系统保留表名--
                    if (string.Equals(sName, "项目信息", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "采购清单", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "领料清单", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "材料分布表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "元件汇总分布表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "元件汇总表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "屏柜汇总表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(sName, "元器件数据管理", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // 若用户指定了勾选的分类工作表，未勾选的直接跳过
                    if (selectedSheetSet != null && !selectedSheetSet.Contains(sName))
                    {
                        continue;
                    }

                    // 校验是否为有效分类明细表 (白名单或特征过滤)
                    if (!Tool.IsProjectCategorySheet(sheet))
                    {
                        continue;
                    }

                    // 规则 8: 在操作 Excel 提取前执行 FixAndFillCabinetNamesForSheet 确保规则 6 拓扑健康
                    Tool.FixAndFillCabinetNamesForSheet(sheet);

                    // 获取当前分类表中的所有有效箱柜锚点 (规则 6 / 规则 11)
                    var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);
                    if (validCabinets == null || validCabinets.Count == 0) continue;

                    // 逐个扫描该分类表下的箱柜
                    for (int i = 0; i < validCabinets.Count; i++)
                    {
                        var cabAnchor = validCabinets[i].Value;

                        // 提取箱柜台数 (从汇总行 F 列读取)
                        int cabinetQty = 1;
                        if (cabAnchor.Sum != null)
                        {
                            int sumRow = Convert.ToInt32(cabAnchor.Sum.Row);
                            // 规则 7: 一次性读取汇总行第 1~10 列
                            dynamic sumRange = sheet.Range[sheet.Cells[sumRow, 1], sheet.Cells[sumRow, 10]];
                            object[,]? sumMatrix = sumRange.Value2 as object[,];

                            if (sumMatrix != null)
                            {
                                // F 列对应第 6 列 (箱柜数量)
                                object qVal = sumMatrix[1, 6] ?? sumMatrix[1, 5];
                                if (qVal != null && int.TryParse(Convert.ToString(qVal), out int q) && q > 0)
                                {
                                    cabinetQty = q;
                                }
                            }
                        }

                        // 提取箱柜内部元器件明细 (起点 detRow+2，终点 subsumRow-1，规则 6)
                        if (cabAnchor.Det != null && cabAnchor.Subsum != null)
                        {
                            int detRow = Convert.ToInt32(cabAnchor.Det.Row);
                            int subsumRow = Convert.ToInt32(cabAnchor.Subsum.Row);
                            int startCompRow = detRow + 2;
                            int endCompRow = subsumRow - 1;

                            if (endCompRow >= startCompRow)
                            {
                                int compRowCount = endCompRow - startCompRow + 1;
                                // 规则 7: 一次性将元器件区域 A 列至 I 列读入内存二维数组
                                dynamic compRange = sheet.Range[sheet.Cells[startCompRow, 1], sheet.Cells[endCompRow, 9]];
                                object[,]? compMatrix = compRange.Value2 as object[,];

                                if (compMatrix != null)
                                {
                                    for (int r = 1; r <= compRowCount; r++)
                                    {
                                        // 提取元器件名称 (B 列 / col 2)
                                        string compName = Convert.ToString(compMatrix[r, 2])?.Trim() ?? "";
                                        // 提取元器件型号规格 (C 列 / col 3)
                                        string compModel = Convert.ToString(compMatrix[r, 3])?.Trim() ?? "";

                                        // 若名称与型号皆为空，视为无效空白行跳过
                                        if (string.IsNullOrWhiteSpace(compName) && string.IsNullOrWhiteSpace(compModel))
                                        {
                                            continue;
                                        }

                                        // 提取生产厂家/品牌 (D 列 / col 4)
                                        string compBrand = Convert.ToString(compMatrix[r, 4])?.Trim() ?? "";
                                        // 提取单位 (E 列 / col 5)，默认补 "台" --硬编码: 默认单位--
                                        string compUnit = Convert.ToString(compMatrix[r, 5])?.Trim() ?? "台";
                                        if (string.IsNullOrWhiteSpace(compUnit)) compUnit = "台";

                                        // 提取单台数量 (F 列 / col 6)
                                        decimal singleQty = 1;
                                        object cQVal = compMatrix[r, 6];
                                        if (cQVal != null && decimal.TryParse(Convert.ToString(cQVal), out decimal cq) && cq > 0)
                                        {
                                            singleQty = cq;
                                        }

                                        // 计算该元器件在该箱柜中的总用量: 单台数量 * 箱柜台数
                                        decimal subTotalQty = singleQty * cabinetQty;

                                        // 核心聚合键：严格按照【名称+型号】唯一确定采购项，彻底杜绝因单位或品牌微小差异导致同物料被拆分
                                        string aggKey = $"{compName}___{compModel}";

                                        if (aggregatedDict.TryGetValue(aggKey, out var existingItem))
                                        {
                                            // 采购总数量累计求和
                                            existingItem.TotalQuantity += subTotalQty;

                                            // 智能单位优先保障: 若后续出现标准"台"或原单位为空，优先更新
                                            if (string.Equals(compUnit, "台", StringComparison.OrdinalIgnoreCase) && !string.Equals(existingItem.Unit, "台", StringComparison.OrdinalIgnoreCase))
                                            {
                                                existingItem.Unit = "台";
                                            }

                                            // 品牌多源并存智能拼接: 若存在不同品牌，去重合并至品牌列
                                            if (!string.IsNullOrWhiteSpace(compBrand))
                                            {
                                                if (string.IsNullOrWhiteSpace(existingItem.Brand))
                                                {
                                                    existingItem.Brand = compBrand;
                                                }
                                                else
                                                {
                                                    // 检测是否已包含当前品牌 (大小写不敏感)
                                                    var existingBrands = existingItem.Brand.Split(new[] { '/', ',', '、' }, StringSplitOptions.RemoveEmptyEntries).Select(b => b.Trim());
                                                    if (!existingBrands.Contains(compBrand, StringComparer.OrdinalIgnoreCase))
                                                    {
                                                        existingItem.Brand = $"{existingItem.Brand}/{compBrand}";
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // 新增聚合记录项
                                            aggregatedDict[aggKey] = new PurchaseListItem
                                            {
                                                Name = compName,
                                                Model = compModel,
                                                Unit = compUnit,
                                                Brand = compBrand,
                                                TotalQuantity = subTotalQty
                                            };
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // 校验扫描到的聚合结果有效性
                if (aggregatedDict.Count == 0)
                {
                    result.Message = "所选分类明细表中未扫描到任何有效的元器件物料！";
                    return result;
                }

                // 依据用户选择的排序策略进行自然排序
                List<PurchaseListItem> sortedItems;
                if (string.Equals(sortBy, "name_model", StringComparison.OrdinalIgnoreCase))
                {
                    // 排序模式 1: 按【名称 + 型号规格】排序
                    sortedItems = aggregatedDict.Values
                        .OrderBy(x => x.Name)
                        .ThenBy(x => x.Model)
                        .ToList();
                }
                else
                {
                    // 排序模式 2 (默认): 按【品牌 + 名称 + 型号规格】多级优先排序
                    sortedItems = aggregatedDict.Values
                        .OrderBy(x => x.Brand)
                        .ThenBy(x => x.Name)
                        .ThenBy(x => x.Model)
                        .ToList();
                }

                // 赋予连续自增序号 (1, 2, 3...)
                for (int idx = 0; idx < sortedItems.Count; idx++)
                {
                    sortedItems[idx].Index = idx + 1;
                }

                // 5. 挂起屏幕刷新、事件响应与警告弹窗，提升表格操作性能与杜绝事件交叉干扰
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.DisplayAlerts = false;

                try
                {
                    // 6. 定位或从母版克隆【采购清单】工作表
                    dynamic? purchaseSheet = null;
                    foreach (dynamic ws in activeWb.Worksheets)
                    {
                        if (string.Equals(Convert.ToString(ws.Name)?.Trim(), "采购清单", StringComparison.OrdinalIgnoreCase))
                        {
                            purchaseSheet = ws;
                            break;
                        }
                    }

                    // 若当前工作簿已存在旧的【采购清单】，静默安全移除旧工作表
                    // 彻底解决旧数据行残留堆叠、第 7 行占位符失效以及品牌列丢失的缺陷
                    if (purchaseSheet != null)
                    {
                        try
                        {
                            purchaseSheet.Delete();
                            purchaseSheet = null;
                        }
                        catch (Exception exDel)
                        {
                            LogHelper.WriteLog($"[采购清单] 移除旧采购清单工作表容错: {exDel.Message}");
                        }
                    }

                    // 统一从标准母版动态克隆全新纯净的【采购清单】工作表
                    // 获取模板物理文件路径 (统一检索 CabinetTemplate.xlsx)
                    string templatePath = Controllers.ProjectController.EnsureCabinetTemplate(app);
                    if (!File.Exists(templatePath))
                    {
                        result.Message = $"未找到标准模板文件: {templatePath}";
                        return result;
                    }

                    // 只读方式快速打开母版工作簿
                    dynamic templateWb = app.Workbooks.Open(templatePath, ReadOnly: true);
                    try
                    {
                        dynamic? tplSheet = null;
                        foreach (dynamic sh in templateWb.Worksheets)
                        {
                            if (string.Equals(Convert.ToString(sh.Name)?.Trim(), "采购清单", StringComparison.OrdinalIgnoreCase))
                            {
                                tplSheet = sh;
                                break;
                            }
                        }

                        if (tplSheet == null)
                        {
                            result.Message = "母版 CabinetTemplate.xlsx 中未包含【采购清单】模板工作表！";
                            return result;
                        }

                        // 将母版中的【采购清单】工作表复制到当前工作簿最末尾
                        tplSheet.Copy(After: activeWb.Sheets[activeWb.Sheets.Count]);
                        // 获取最新复制的工作表实例
                        purchaseSheet = activeWb.Sheets[activeWb.Sheets.Count];
                        try { purchaseSheet.Name = "采购清单"; } catch { }
                    }
                    finally
                    {
                        // 严格遵循启发式规则 12: 复制完成后立即关闭模板工作簿句柄
                        try { templateWb.Close(false); } catch { }
                    }

                    // 7. 解析第 7 行模板占位符 (自适应零硬编码列映射)
                    var colMap = new PurchaseListTemplateColumnMap();
                    try
                    {
                        // 一次性读取第 7 行 A 列至 O 列 (前 15 列)
                        dynamic row7Range = purchaseSheet.Range["A7:O7"];
                        object[,] row7Matrix = (object[,])row7Range.Value2;

                        for (int col = 1; col <= 15; col++)
                        {
                            string cellText = Convert.ToString(row7Matrix[1, col])?.Trim() ?? "";

                            // 匹配元器件名称占位符 [元器件名称] 或 [名称]
                            if (cellText.Contains("[元器件名称]") || cellText.Contains("[名称]"))
                            {
                                colMap.NameCol = col;
                            }
                            // 匹配型号规格占位符 [型号] 或 [型号规格]
                            else if (cellText.Contains("[型号]") || cellText.Contains("[型号规格]"))
                            {
                                colMap.ModelCol = col;
                            }
                            // 匹配计量单位占位符 [单位]
                            else if (cellText.Contains("[单位]"))
                            {
                                colMap.UnitCol = col;
                            }
                            // 匹配采购数量占位符 [数量] 或 [单数]
                            else if (cellText.Contains("[数量]") || cellText.Contains("[单数]"))
                            {
                                colMap.QuantityCol = col;
                            }
                            // 匹配生产厂家/品牌占位符 [品牌] 或 [厂家] 或 [生产厂家]
                            else if (cellText.Contains("[品牌]") || cellText.Contains("[厂家]") || cellText.Contains("[生产厂家]"))
                            {
                                colMap.BrandCol = col;
                            }
                        }

                        // 动态计算最大受控列宽
                        int detectedMaxCol = Math.Max(colMap.MaxCol, Math.Max(colMap.QuantityCol, colMap.BrandCol));
                        colMap.MaxCol = Math.Max(9, detectedMaxCol);
                    }
                    catch (Exception exMap)
                    {
                        LogHelper.WriteLog($"[采购清单] 模板列解析容错: {exMap.Message}");
                    }

                    // 8. 填充抬头大标题与工程信息
                    try
                    {
                        // 替换 A1 公司名称占位符 (若已读取到实际公司名称)
                        if (!string.IsNullOrWhiteSpace(companyName))
                        {
                            string a1Val = Convert.ToString(purchaseSheet.Range["A1"].Value2) ?? "";
                            if (a1Val.Contains("[公司名称]") || string.IsNullOrWhiteSpace(a1Val))
                            {
                                purchaseSheet.Range["A1"].Value2 = companyName;
                            }
                        }

                        // 替换 A4 项目名称占位符
                        string a4Val = Convert.ToString(purchaseSheet.Range["A4"].Value2) ?? "";
                        if (a4Val.Contains("[项目名称]"))
                        {
                            // 保留原有前缀并替换占位符
                            purchaseSheet.Range["A4"].Value2 = a4Val.Replace("[项目名称]", projectName);
                        }
                        else if (string.IsNullOrWhiteSpace(a4Val))
                        {
                            purchaseSheet.Range["A4"].Value2 = $"项目名称: {projectName}";
                        }
                    }
                    catch (Exception exHeader)
                    {
                        LogHelper.WriteLog($"[采购清单] 替换标题信息容错: {exHeader.Message}");
                    }

                    // 9. 数据行智能排版与批量回写 (规则 7)
                    int itemCount = sortedItems.Count;
                    int startRow = 7; // 模板首个数据行起始物理行号 --硬编码: 首行--
                    int defaultTemplateRows = 6; // 模板中预留的 7~12 行共 6 行网格 --硬编码: 预留行数--

                    // 若数据项超过模板预留的 6 行，在第 12 行下方批量插入行并沿用第 7 行样式
                    if (itemCount > defaultTemplateRows)
                    {
                        int needInsert = itemCount - defaultTemplateRows;
                        int insertStartRow = startRow + defaultTemplateRows; // 第 13 行
                        int insertEndRow = insertStartRow + needInsert - 1;

                        // 整块向下推挤插入空行 (xlShiftDown = -4121)
                        dynamic insertRows = purchaseSheet.Rows[$"{insertStartRow}:{insertEndRow}"];
                        insertRows.Insert(-4121);

                        // 复制第 7 行的单元格格式 (xlPasteFormats = -4122)
                        purchaseSheet.Rows[7].Copy();
                        dynamic formatTarget = purchaseSheet.Rows[$"{insertStartRow}:{insertEndRow}"];
                        formatTarget.PasteSpecial(-4122);
                        app.CutCopyMode = 0;
                    }

                    // 构造数据写入矩阵 (大小: itemCount 行 x maxCol 列)
                    object[,] writeMatrix = new object[itemCount, colMap.MaxCol];

                    for (int i = 0; i < itemCount; i++)
                    {
                        var it = sortedItems[i];

                        // A 列 (序号)
                        if (colMap.IndexCol >= 1 && colMap.IndexCol <= colMap.MaxCol)
                        {
                            writeMatrix[i, colMap.IndexCol - 1] = it.Index;
                        }

                        // B 列 (名称)
                        if (colMap.NameCol >= 1 && colMap.NameCol <= colMap.MaxCol)
                        {
                            writeMatrix[i, colMap.NameCol - 1] = it.Name;
                        }

                        // C 列 (型号规格)
                        if (colMap.ModelCol >= 1 && colMap.ModelCol <= colMap.MaxCol)
                        {
                            writeMatrix[i, colMap.ModelCol - 1] = it.Model;
                        }

                        // D 列 (单位)
                        if (colMap.UnitCol >= 1 && colMap.UnitCol <= colMap.MaxCol)
                        {
                            writeMatrix[i, colMap.UnitCol - 1] = it.Unit;
                        }

                        // E 列 (采购总数量)
                        if (colMap.QuantityCol >= 1 && colMap.QuantityCol <= colMap.MaxCol)
                        {
                            writeMatrix[i, colMap.QuantityCol - 1] = it.TotalQuantity;
                        }

                        // 若模板匹配到了品牌列 (如 I 列备注下 [品牌])
                        if (colMap.BrandCol >= 1 && colMap.BrandCol <= colMap.MaxCol)
                        {
                            writeMatrix[i, colMap.BrandCol - 1] = it.Brand;
                        }
                    }

                    // 规则 7: 采用内存二维数组一次性回写全部采购物料数据
                    int endDataRow = startRow + itemCount - 1;
                    dynamic targetDataRange = purchaseSheet.Range[purchaseSheet.Cells[startRow, 1], purchaseSheet.Cells[endDataRow, colMap.MaxCol]];
                    targetDataRange.Value2 = writeMatrix;

                    // 设置统一行高与边框自愈
                    try
                    {
                        // 细实线边框 (xlContinuous = 1, xlThin = 2)
                        targetDataRange.Borders.LineStyle = 1;
                        targetDataRange.Borders.Weight = 2;
                        // 统一设置舒适行高
                        targetDataRange.RowHeight = 22;

                        // 序号、单位、数量列居中对齐 (xlCenter = -4108)
                        if (colMap.IndexCol >= 1) purchaseSheet.Range[purchaseSheet.Cells[startRow, colMap.IndexCol], purchaseSheet.Cells[endDataRow, colMap.IndexCol]].HorizontalAlignment = -4108;
                        if (colMap.UnitCol >= 1) purchaseSheet.Range[purchaseSheet.Cells[startRow, colMap.UnitCol], purchaseSheet.Cells[endDataRow, colMap.UnitCol]].HorizontalAlignment = -4108;
                        if (colMap.QuantityCol >= 1) purchaseSheet.Range[purchaseSheet.Cells[startRow, colMap.QuantityCol], purchaseSheet.Cells[endDataRow, colMap.QuantityCol]].HorizontalAlignment = -4108;
                        // 品牌列居中对齐
                        if (colMap.BrandCol >= 1) purchaseSheet.Range[purchaseSheet.Cells[startRow, colMap.BrandCol], purchaseSheet.Cells[endDataRow, colMap.BrandCol]].HorizontalAlignment = -4108;

                        // 名称与型号列靠左对齐 (xlLeft = -4131)
                        if (colMap.NameCol >= 1) purchaseSheet.Range[purchaseSheet.Cells[startRow, colMap.NameCol], purchaseSheet.Cells[endDataRow, colMap.NameCol]].HorizontalAlignment = -4131;
                        if (colMap.ModelCol >= 1) purchaseSheet.Range[purchaseSheet.Cells[startRow, colMap.ModelCol], purchaseSheet.Cells[endDataRow, colMap.ModelCol]].HorizontalAlignment = -4131;
                    }
                    catch (Exception exStyle)
                    {
                        LogHelper.WriteLog($"[采购清单] 格式与边框美化容错: {exStyle.Message}");
                    }

                    // 10. 若实际项数小于模板预留的 6 行，清空多余预留占位行的残留文本 (保留边框网格)
                    if (itemCount < defaultTemplateRows)
                    {
                        int cleanStartRow = startRow + itemCount;
                        int cleanEndRow = startRow + defaultTemplateRows - 1;
                        dynamic cleanRange = purchaseSheet.Range[purchaseSheet.Cells[cleanStartRow, 1], purchaseSheet.Cells[cleanEndRow, colMap.MaxCol]];
                        cleanRange.ClearContents();
                    }

                    // 11. 激活聚焦到【采购清单】工作表并滚屏至首行
                    purchaseSheet.Activate();
                    try
                    {
                        app.ActiveWindow.ScrollRow = 1;
                        app.ActiveWindow.ScrollColumn = 1;
                    }
                    catch { }

                    // 标记成功并记录项数
                    result.Success = true;
                    result.ItemCount = itemCount;
                    result.Message = $"采购清单自动导出成功！\n共汇总工程物料 {itemCount} 项，已完成所有箱柜台数的综合累计。";
                }
                finally
                {
                    // 恢复 Excel 屏幕渲染、事件响应与警告对话框
                    try
                    {
                        app.DisplayAlerts = true;
                        app.ScreenUpdating = true;
                        app.EnableEvents = true;
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                // 异常处理与日志记录
                LogHelper.WriteLog($"[采购清单] 导出全流程异常: {ex}");
                result.Success = false;
                result.Message = $"导出采购清单失败: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// 直接根据当前工作簿已有的【采购清单】工作表，全量提取元器件物料数据，
        /// 并将采购清单 J 列内容精准映射至领料单 K 列 [库存标记]，自动克隆并生成【领料清单】工作表
        /// 严格遵循规则 6、7、8，零硬编码自适应模板网格排版，彻底杜绝死锁与旧数据堆叠
        /// </summary>
        /// <returns>操作执行结果对象</returns>
        public static PurchaseListExportResult ExportPickListFromCurrentWorkbook()
        {
            // 初始化返回结果对象
            var result = new PurchaseListExportResult();

            try
            {
                // 1. 安全获取当前 Excel 宿主应用程序实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    // 无法连接 Excel 提示错误
                    result.Message = "无法连接至 Excel 宿主应用程序！";
                    MessageBox.Show(result.Message, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return result;
                }

                // 2. 获取当前活动算价工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null)
                {
                    // 无活动工作簿提示警告
                    result.Message = "当前没有打开的活动工作簿，无法生成领料清单！";
                    MessageBox.Show(result.Message, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return result;
                }

                // 3. 严格遵循前置依赖拦截规则：检索当前工作簿是否已存在【采购清单】
                dynamic? purchaseSheet = null;
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    // 匹配采购清单工作表名 --硬编码: 系统工作表名称--
                    if (string.Equals(Convert.ToString(ws.Name)?.Trim(), "采购清单", StringComparison.OrdinalIgnoreCase))
                    {
                        purchaseSheet = ws;
                        break;
                    }
                }

                // 若未检测到采购清单，弹出友好拦截提示并终止
                if (purchaseSheet == null)
                {
                    result.Message = "当前工作簿未检测到【采购清单】工作表！\n\n请先点击功能区【材料统计】生成采购清单后，再执行本功能生成领料单。";
                    MessageBox.Show(result.Message, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return result;
                }

                // 4. 扫描【采购清单】数据范围 (从第 7 行起扫描至数据末行)
                int startScanRow = 7; // 采购清单首个数据行物理行号 --硬编码: 首行--
                int maxScanRow = 7;

                try
                {
                    // 获取已用区域的总行数作为安全探测上界
                    int usedRows = Convert.ToInt32(purchaseSheet.UsedRange.Rows.Count);
                    int usedStartRow = Convert.ToInt32(purchaseSheet.UsedRange.Row);
                    maxScanRow = Math.Max(7, usedStartRow + usedRows - 1);
                }
                catch
                {
                    // 容错兜底探测行数
                    maxScanRow = 600;
                }

                // 规则 7: 一次性读取采购清单第 7 行至末行的第 1~10 列 (A 列至 J 列) 到内存二维数组
                dynamic pRange = purchaseSheet.Range[purchaseSheet.Cells[startScanRow, 1], purchaseSheet.Cells[maxScanRow, 10]];
                object[,]? pMatrix = pRange.Value2 as object[,];

                if (pMatrix == null)
                {
                    result.Message = "读取【采购清单】数据矩阵失败！";
                    MessageBox.Show(result.Message, "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return result;
                }

                // 逐行解析采购清单中的有效物料项 (全量带入)
                var pickItems = new List<PurchaseListItem>();
                int totalScanRows = maxScanRow - startScanRow + 1;

                for (int r = 1; r <= totalScanRows; r++)
                {
                    // 读取名称 (B 列 / col 2)
                    string name = Convert.ToString(pMatrix[r, 2])?.Trim() ?? "";
                    // 读取型号规格 (C 列 / col 3)
                    string model = Convert.ToString(pMatrix[r, 3])?.Trim() ?? "";
                    // 读取序号 (A 列 / col 1)
                    string seq = Convert.ToString(pMatrix[r, 1])?.Trim() ?? "";

                    // 若名称、型号与序号皆为空白，视为有效数据已结束，跳出循环
                    if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(model) && string.IsNullOrWhiteSpace(seq))
                    {
                        break;
                    }

                    // 读取计量单位 (D 列 / col 4)
                    string unit = Convert.ToString(pMatrix[r, 4])?.Trim() ?? "台";
                    if (string.IsNullOrWhiteSpace(unit)) unit = "台";

                    // 读取采购数量 (E 列 / col 5)
                    decimal qty = 1;
                    object qVal = pMatrix[r, 5];
                    if (qVal != null && decimal.TryParse(Convert.ToString(qVal), out decimal q) && q > 0)
                    {
                        qty = q;
                    }

                    // 读取生产厂家/品牌 (I 列 / col 9)
                    string brand = Convert.ToString(pMatrix[r, 9])?.Trim() ?? "";

                    // 核心逻辑：读取采购清单 J 列内容作为 [库存标记] (J 列 / col 10)
                    string stockTag = Convert.ToString(pMatrix[r, 10])?.Trim() ?? "";

                    // 将物料项加入领料清单列表中
                    pickItems.Add(new PurchaseListItem
                    {
                        Index = pickItems.Count + 1,
                        Name = name,
                        Model = model,
                        Unit = unit,
                        TotalQuantity = qty,
                        Brand = brand,
                        Remark = stockTag // 采购清单 J 列内容直接写入领料单备注列 (K列)
                    });
                }

                // 校验物料有效性
                if (pickItems.Count == 0)
                {
                    result.Message = "【采购清单】中未扫描到任何有效的元器件物料数据！";
                    MessageBox.Show(result.Message, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return result;
                }

                // 5. 提取项目工程信息 (公司名称与项目名称)
                string projectName = Path.GetFileNameWithoutExtension(Convert.ToString(activeWb.Name) ?? "");
                string companyName = string.Empty;

                try
                {
                    // 尝试从【项目信息】工作表读取工程信息 --硬编码: 工作表名称--
                    dynamic? infoSheet = null;
                    try { infoSheet = activeWb.Sheets["项目信息"]; } catch { }

                    if (infoSheet != null)
                    {
                        // 规则 7: 一次性读取 A1:C30 区域至内存二维数组
                        dynamic infoRange = infoSheet.Range["A1:C30"];
                        object[,] infoMatrix = (object[,])infoRange.Value2;

                        // 提取项目名称 (Cell B5)
                        string pName = Convert.ToString(infoMatrix[5, 2])?.Trim() ?? "";
                        if (!string.IsNullOrWhiteSpace(pName)) projectName = pName;

                        // 提取供货单位或公司名称 (Cell B22 或 B1)
                        string comp = Convert.ToString(infoMatrix[22, 2])?.Trim() ?? "";
                        if (string.IsNullOrWhiteSpace(comp)) comp = Convert.ToString(infoMatrix[1, 2])?.Trim() ?? "";
                        companyName = comp;
                    }
                }
                catch (Exception exInfo)
                {
                    // 记录读取工程信息容错
                    LogHelper.WriteLog($"[领料清单] 读取项目信息容错: {exInfo.Message}");
                }

                // 6. 挂起屏幕刷新、事件响应与系统警告，提升操作性能并杜绝事件冲突
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.DisplayAlerts = false;

                try
                {
                    // 7. 定位或从母版克隆【领料清单】工作表
                    dynamic? pickSheet = null;
                    foreach (dynamic ws in activeWb.Worksheets)
                    {
                        if (string.Equals(Convert.ToString(ws.Name)?.Trim(), "领料清单", StringComparison.OrdinalIgnoreCase))
                        {
                            pickSheet = ws;
                            break;
                        }
                    }

                    // 若当前工作簿已存在旧的【领料清单】，静默安全将其移除
                    // 彻底根除旧表数据堆叠膨胀、格式错乱以及死锁缺陷
                    if (pickSheet != null)
                    {
                        try
                        {
                            pickSheet.Delete();
                            pickSheet = null;
                        }
                        catch (Exception exDel)
                        {
                            LogHelper.WriteLog($"[领料清单] 移除旧领料清单工作表容错: {exDel.Message}");
                        }
                    }

                    // 统一从标准母版动态克隆全新纯净的【领料清单】工作表
                    // 获取模板物理文件路径 (统一检索 CabinetTemplate.xlsx)
                    string templatePath = Controllers.ProjectController.EnsureCabinetTemplate(app);
                    if (!File.Exists(templatePath))
                    {
                        result.Message = $"未找到标准模板文件: {templatePath}";
                        MessageBox.Show(result.Message, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return result;
                    }

                    // 只读方式快速打开母版工作簿
                    dynamic templateWb = app.Workbooks.Open(templatePath, ReadOnly: true);
                    try
                    {
                        dynamic? tplSheet = null;
                        foreach (dynamic sh in templateWb.Worksheets)
                        {
                            if (string.Equals(Convert.ToString(sh.Name)?.Trim(), "领料清单", StringComparison.OrdinalIgnoreCase))
                            {
                                tplSheet = sh;
                                break;
                            }
                        }

                        if (tplSheet == null)
                        {
                            result.Message = "母版 CabinetTemplate.xlsx 中未包含【领料清单】模板工作表！";
                            MessageBox.Show(result.Message, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return result;
                        }

                        // 将母版中的【领料清单】工作表复制到当前工作簿最末尾
                        tplSheet.Copy(After: activeWb.Sheets[activeWb.Sheets.Count]);
                        // 获取最新复制的工作表实例
                        pickSheet = activeWb.Sheets[activeWb.Sheets.Count];
                        try { pickSheet.Name = "领料清单"; } catch { }
                    }
                    finally
                    {
                        // 严格遵循启发式规则 12: 复制完成后立即关闭模板工作簿句柄
                        try { templateWb.Close(false); } catch { }
                    }

                    // 8. 填充抬头大标题与工程信息
                    try
                    {
                        // 替换 A1 公司名称占位符 (若已读取到实际公司名称)
                        if (!string.IsNullOrWhiteSpace(companyName))
                        {
                            string a1Val = Convert.ToString(pickSheet.Range["A1"].Value2) ?? "";
                            if (a1Val.Contains("[公司名称]") || string.IsNullOrWhiteSpace(a1Val))
                            {
                                pickSheet.Range["A1"].Value2 = companyName;
                            }
                        }

                        // 替换 A3 工程名称占位符
                        string a3Val = Convert.ToString(pickSheet.Range["A3"].Value2) ?? "";
                        if (a3Val.Contains("[项目名称]"))
                        {
                            // 替换工程名称
                            pickSheet.Range["A3"].Value2 = a3Val.Replace("[项目名称]", projectName);
                        }
                        else if (string.IsNullOrWhiteSpace(a3Val))
                        {
                            pickSheet.Range["A3"].Value2 = $"工程名称 ： {projectName}";
                        }

                        // 用户明确确认：Row 4 制表日期不自动填入日期，队组名称保持星号不变
                    }
                    catch (Exception exHeader)
                    {
                        LogHelper.WriteLog($"[领料清单] 替换标题信息容错: {exHeader.Message}");
                    }

                    // 9. 数据行智能排版与批量插入行 (规则 6 / 规则 7)
                    int itemCount = pickItems.Count;
                    int startRow = 7; // 首个物料起始行 --硬编码: 首行--
                    int defaultTemplateRows = 4; // 模板中第 7~10 行预留 4 行网格，第 11 行为【合计】行 --硬编码: 预留行数--

                    // 若实际物料项数超过模板预留的 4 行，在第 11 行 (合计行) 上方批量插入行
                    if (itemCount > defaultTemplateRows)
                    {
                        int needInsert = itemCount - defaultTemplateRows;
                        int insertStartRow = 11; // 在第 11 行合计行处向上推移插入
                        int insertEndRow = insertStartRow + needInsert - 1;

                        // 整块向下推挤插入空行 (xlShiftDown = -4121)
                        dynamic insertRows = pickSheet.Rows[$"{insertStartRow}:{insertEndRow}"];
                        insertRows.Insert(-4121);

                        // 复制第 7 行的单元格格式 (xlPasteFormats = -4122)
                        pickSheet.Rows[7].Copy();
                        dynamic formatTarget = pickSheet.Rows[$"{insertStartRow}:{insertEndRow}"];
                        formatTarget.PasteSpecial(-4122);
                        app.CutCopyMode = 0;
                    }

                    // 10. 构造数据写入矩阵 (大小: itemCount 行 x 11 列 / A 列至 K 列)
                    object[,] writeMatrix = new object[itemCount, 11];

                    for (int i = 0; i < itemCount; i++)
                    {
                        var it = pickItems[i];

                        // A 列: 序号 (第 1 列)
                        writeMatrix[i, 0] = it.Index;
                        // B 列: 名称 (第 2 列)
                        writeMatrix[i, 1] = it.Name;
                        // C 列: 型号 (第 3 列)
                        writeMatrix[i, 2] = it.Model;
                        // D 列: 单位 (第 4 列)
                        writeMatrix[i, 3] = it.Unit;
                        // E 列: 数量 (第 5 列)
                        writeMatrix[i, 4] = it.TotalQuantity;
                        // F 列: 生产厂家 (第 6 列)
                        writeMatrix[i, 5] = it.Brand;
                        // G 列: 领料人1数量 (第 7 列，留空)
                        writeMatrix[i, 6] = string.Empty;
                        // H 列: 领料人1签字 (第 8 列，留空)
                        writeMatrix[i, 7] = string.Empty;
                        // I 列: 领料人2数量 (第 9 列，留空)
                        writeMatrix[i, 8] = string.Empty;
                        // J 列: 领料人2签字 (第 10 列，留空)
                        writeMatrix[i, 9] = string.Empty;
                        // K 列: 备注 (第 11 列，直接写入采购清单 J 列 [库存标记] 内容)
                        writeMatrix[i, 10] = it.Remark;
                    }

                    // 规则 7: 采用内存二维数组一次性回写全部领料物料数据至 A7:K{endRow}
                    int endDataRow = startRow + itemCount - 1;
                    dynamic targetDataRange = pickSheet.Range[pickSheet.Cells[startRow, 1], pickSheet.Cells[endDataRow, 11]];
                    targetDataRange.Value2 = writeMatrix;

                    // 设置统一行高与边框自愈
                    try
                    {
                        // 细实线边框 (xlContinuous = 1, xlThin = 2)
                        targetDataRange.Borders.LineStyle = 1;
                        targetDataRange.Borders.Weight = 2;
                        // 统一设置舒适行高
                        targetDataRange.RowHeight = 22;

                        // 序号、单位、数量、生产厂家、备注列居中对齐 (xlCenter = -4108)
                        pickSheet.Range[pickSheet.Cells[startRow, 1], pickSheet.Cells[endDataRow, 1]].HorizontalAlignment = -4108;
                        pickSheet.Range[pickSheet.Cells[startRow, 4], pickSheet.Cells[endDataRow, 5]].HorizontalAlignment = -4108;
                        pickSheet.Range[pickSheet.Cells[startRow, 6], pickSheet.Cells[endDataRow, 6]].HorizontalAlignment = -4108;
                        pickSheet.Range[pickSheet.Cells[startRow, 11], pickSheet.Cells[endDataRow, 11]].HorizontalAlignment = -4108;

                        // 名称与型号列靠左对齐 (xlLeft = -4131)
                        pickSheet.Range[pickSheet.Cells[startRow, 2], pickSheet.Cells[endDataRow, 3]].HorizontalAlignment = -4131;
                    }
                    catch (Exception exStyle)
                    {
                        LogHelper.WriteLog($"[领料清单] 格式与边框美化容错: {exStyle.Message}");
                    }

                    // 11. 若实际物料项数小于预留的 4 行，清空多余预留占位行的残留文本 (保留网格边框)
                    if (itemCount < defaultTemplateRows)
                    {
                        int cleanStartRow = startRow + itemCount;
                        int cleanEndRow = startRow + defaultTemplateRows - 1;
                        dynamic cleanRange = pickSheet.Range[pickSheet.Cells[cleanStartRow, 1], pickSheet.Cells[cleanEndRow, 11]];
                        cleanRange.ClearContents();
                    }

                    // 用户明确确认：合计行保持原样，不需要自动注入求和公式

                    // 12. 激活聚焦到【领料清单】工作表并滚屏至首行首列
                    pickSheet.Activate();
                    try
                    {
                        app.ActiveWindow.ScrollRow = 1;
                        app.ActiveWindow.ScrollColumn = 1;
                    }
                    catch { }

                    // 标记成功并记录项数
                    result.Success = true;
                    result.ItemCount = itemCount;
                    result.Message = $"成套设备装配领料单生成成功！\n\n已基于【采购清单】及其库存标记，全量同步导出物料 {itemCount} 项。";

                    // 弹出友好完成提示
                    MessageBox.Show(result.Message, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    // 恢复 Excel 屏幕渲染、事件响应与警告对话框
                    try
                    {
                        app.DisplayAlerts = true;
                        app.ScreenUpdating = true;
                        app.EnableEvents = true;
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                // 异常处理与日志记录
                LogHelper.WriteLog($"[领料清单] 导出全流程异常: {ex}");
                result.Success = false;
                result.Message = $"生成领料清单失败: {ex.Message}";
                MessageBox.Show(result.Message, "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }
    }
}
