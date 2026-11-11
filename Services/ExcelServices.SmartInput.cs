using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;
using ExcelDna.Integration;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 核心业务服务分部类：智能输入与选择表联动
    /// </summary>
    public static partial class ExcelServices
    {
        // 智能输入配置窗口静态单例引用 (可空)
        private static SmartInputForm? _smartInputForm;

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“智能输入配置”窗口 (非模态，保持 Excel 可交互)
        /// </summary>
        public static void ShowSmartInputDialog()
        {
            try
            {
                // 以非模态方式展示智能输入配置窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _smartInputForm, () => new SmartInputForm());
            }
            catch (Exception ex)
            {
                // 捕获弹窗异常防止 Excel 崩溃闪退
                System.Windows.Forms.MessageBox.Show($"弹出智能输入配置窗口失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// <summary>
        /// 从当前活动工作簿的所有工作表中批量提取元器件数据，并按工作表去重归纳 (规则 6 & 规则 7)
        /// 具备三级探测保障：1. 双作用域定义名称扫描；2. 智能自动补齐校准；3. 表头特征区域兜底扫描
        /// </summary>
        /// <returns>提取并去重后的元器件存储数据根对象</returns>
        public static SmartComponentsStorage ExtractComponentsFromAllSheets()
        {
            // 构造元器件存储根对象
            var storage = new SmartComponentsStorage();

            try
            {
                // 获取 Excel Application COM 接口实例 (安全调用)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return storage;

                // 获取当前活动工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null) return storage;

                // 从配置中读取箱柜定义名称前缀值对象 (零堆分配)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

                // 遍历当前工作簿中的所有工作表
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    string sheetName = Convert.ToString(sheet.Name) ?? "";

                    // 过滤内部隐藏字典表与选择表等系统管理表
                    if (string.IsNullOrWhiteSpace(sheetName) ||
                        sheetName.StartsWith("_") ||
                        sheetName.Equals("选择表", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // 收集当前工作表中的所有箱柜元器件有效行与计费区域有效行区间 (startRow, endRow)
                    var cabinetRanges = new List<(int startRow, int endRow)>();

                    // 1. 获取当前工作表有效箱柜映射 (复用 Tool 公共方法，内置空值自动智能补齐重建)
                    var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);

                    // 若解析到有效箱柜锚点，计算各箱柜元器件有效插槽行与计费区域 (规则 6 扩展)
                    if (validCabinets.Count > 0)
                    {
                        foreach (var cab in validCabinets)
                        {
                            if (cab.Value.Det != null && cab.Value.Subsum != null)
                            {
                                int detRow = Convert.ToInt32(cab.Value.Det.Row);
                                int subsumRow = Convert.ToInt32(cab.Value.Subsum.Row);
                                // 提取总计行号 (TolsumRow)
                                int tolsumRow = cab.Value.Tolsum != null ? Convert.ToInt32(cab.Value.Tolsum.Row) : 0;

                                // 依据用户指令：无论在元器件区域还是计费区域，只要 C 列有值均加入源数据
                                int startRow = detRow + 2;
                                // 计费区域延续至总计行 tolsumRow；若未定义 tolsumRow 则向下探测 10 行
                                int endRow = (tolsumRow > subsumRow) ? tolsumRow : (subsumRow + 10);

                                if (startRow <= endRow)
                                {
                                    cabinetRanges.Add((startRow, endRow));
                                }
                            }
                        }
                    }

                    // 临时存储当前工作表提取到的元器件项
                    var sheetComponents = new List<SmartComponentItem>();
                    int rawTotalCount = 0;

                    // 3. 【执行物料数据读取：优先从箱柜区域批量读取，规则 7 数组入内存】
                    if (cabinetRanges.Count > 0)
                    {
                        foreach (var (startRow, endRow) in cabinetRanges)
                        {
                            // 规则 7: 一次性读取二维数据矩阵 (A 列至 AC 列共 29 列，绝对排除第 30 列 AD 句柄列)
                            dynamic range = sheet.Range[$"A{startRow}:AC{endRow}"];
                            object[,] data = range.Value2 as object[,];

                            if (data != null)
                            {
                                int rowCount = data.GetLength(0);
                                int colCount = data.GetLength(1);

                                for (int r = 1; r <= rowCount; r++)
                                {
                                    // 读取 C 列 (规格型号, 列索引 3)
                                    string model = colCount >= 3 ? Convert.ToString(data[r, 3])?.Trim() ?? string.Empty : string.Empty;

                                    // 核心过滤：C 列非空且排除小计、总计纯汇总统计文字
                                    if (!string.IsNullOrWhiteSpace(model) &&
                                        !model.Equals("小计") && !model.Equals("总计") && !model.Equals("合计") &&
                                        !model.StartsWith("小计") && !model.StartsWith("总计"))
                                    {
                                        rawTotalCount++;

                                        // 读取 B 列 (元件名称, 列索引 2)
                                        string name = colCount >= 2 ? Convert.ToString(data[r, 2])?.Trim() ?? string.Empty : string.Empty;

                                        // 读取 D 列 (生产厂家, 列索引 4)
                                        string mfr = colCount >= 4 ? Convert.ToString(data[r, 4])?.Trim() ?? string.Empty : string.Empty;

                                        // 读取 E 列 (计量单位, 列索引 5)
                                        string unit = colCount >= 5 ? Convert.ToString(data[r, 5])?.Trim() ?? string.Empty : string.Empty;

                                        // 读取 G 列 (销售单价, 列索引 7)
                                        decimal unitPrice = 0;
                                        if (colCount >= 7 && decimal.TryParse(Convert.ToString(data[r, 7]), out decimal pVal))
                                        {
                                            unitPrice = pVal;
                                        }

                                        // 读取 I 列 (规格备注, 列索引 9)
                                        string remark = colCount >= 9 ? Convert.ToString(data[r, 9])?.Trim() ?? string.Empty : string.Empty;

                                        // 读取 J 列 (成本单价, 列索引 10)
                                        decimal costPrice = 0;
                                        if (colCount >= 10 && decimal.TryParse(Convert.ToString(data[r, 10]), out decimal cVal))
                                        {
                                            costPrice = cVal;
                                        }

                                        // 读取 L 列 (加价系数, 列索引 12)
                                        string quoteCoeff = colCount >= 12 ? Convert.ToString(data[r, 12])?.Trim() ?? string.Empty : string.Empty;

                                        // 读取 M 列 (表价值/基准面价, 列索引 13)
                                        decimal catalogPrice = 0;
                                        if (colCount >= 13 && decimal.TryParse(Convert.ToString(data[r, 13]), out decimal mVal))
                                        {
                                            catalogPrice = mVal;
                                        }

                                        // 读取 N 列 (采购折扣系数, 列索引 14)
                                        string discountCoeff = colCount >= 14 ? Convert.ToString(data[r, 14])?.Trim() ?? string.Empty : string.Empty;

                                        // 读取 Q 列 (元件类别, 列索引 17)
                                        string category = colCount >= 17 ? Convert.ToString(data[r, 17])?.Trim() ?? string.Empty : string.Empty;

                                        // 提取第 2 列至第 29 列完整单元格属性 (跳过第 6 列 F 数量列，绝不包含第 30 列 AD 句柄列)
                                        var rowCells = new Dictionary<int, string>();
                                        for (int c = 2; c <= Math.Min(colCount, 29); c++)
                                        {
                                            // 绝对排除第 6 列 (F 列数量，保留当前行已有数量不回填)
                                            if (c == 6) continue;
                                            string cellVal = Convert.ToString(data[r, c])?.Trim() ?? string.Empty;
                                            if (!string.IsNullOrEmpty(cellVal))
                                            {
                                                rowCells[c] = cellVal;
                                            }
                                        }

                                        var item = new SmartComponentItem
                                        {
                                            Model = model,
                                            Name = name,
                                            Manufacturer = mfr,
                                            Unit = unit,
                                            UnitPrice = unitPrice,
                                            CostUnitPrice = costPrice,
                                            CatalogPrice = catalogPrice,
                                            QuoteCoefficient = quoteCoeff,
                                            DiscountCoefficient = discountCoeff,
                                            Remark = remark,
                                            Category = category,
                                            SheetName = sheetName,
                                            CabinetNo = $"箱柜_{startRow}",
                                            RowCells = rowCells
                                        };

                                        sheetComponents.Add(item);
                                    }
                                }
                            }
                        }
                    }
                    // 4. 【第三级保障：若无箱柜定义名称，直接扫描 UsedRange 已用区域提取物料】
                    else
                    {
                        dynamic usedRange = sheet.UsedRange;
                        if (usedRange != null)
                        {
                            object[,] data = usedRange.Value2 as object[,];
                            if (data != null)
                            {
                                int uStartRow = Convert.ToInt32(usedRange.Row);
                                int uRows = data.GetLength(0);
                                int uCols = data.GetLength(1);

                                // 寻找表头行 (包含“型号”或“规格”的行)
                                int headerRelativeRow = 0;
                                int modelColIndex = 3; // 默认 C 列
                                int nameColIndex = 2;  // 默认 B 列
                                int mfrColIndex = 4;   // 默认 D 列
                                int unitColIndex = 5;  // 默认 E 列
                                int priceColIndex = 7; // 默认 G 列
                                int catalogPriceColIndex = 13; // 默认 M 列

                                for (int r = 1; r <= Math.Min(uRows, 15); r++)
                                {
                                    for (int c = 1; c <= uCols; c++)
                                    {
                                        string cellTxt = Convert.ToString(data[r, c])?.Trim() ?? "";
                                        if (cellTxt.Contains("型号") || cellTxt.Contains("规格"))
                                        {
                                            headerRelativeRow = r;
                                            modelColIndex = c;
                                            break;
                                        }
                                    }
                                    if (headerRelativeRow > 0) break;
                                }

                                // 若找到表头，提取表头下方的所有有效数据
                                int dataStartR = headerRelativeRow > 0 ? headerRelativeRow + 1 : 1;
                                for (int r = dataStartR; r <= uRows; r++)
                                {
                                    string model = uCols >= modelColIndex ? Convert.ToString(data[r, modelColIndex])?.Trim() ?? "" : "";
                                    // 排除空行与小计/总计合计行
                                    if (!string.IsNullOrWhiteSpace(model) && !model.Contains("小计") && !model.Contains("总计") && !model.Contains("合计"))
                                    {
                                        rawTotalCount++;
                                        string name = uCols >= nameColIndex ? Convert.ToString(data[r, nameColIndex])?.Trim() ?? "" : "";
                                        string mfr = uCols >= mfrColIndex ? Convert.ToString(data[r, mfrColIndex])?.Trim() ?? "" : "";
                                        string unit = uCols >= unitColIndex ? Convert.ToString(data[r, unitColIndex])?.Trim() ?? "" : "";
                                        decimal unitPrice = 0;
                                        if (uCols >= priceColIndex && decimal.TryParse(Convert.ToString(data[r, priceColIndex]), out decimal pVal))
                                        {
                                            unitPrice = pVal;
                                        }

                                        decimal catalogPrice = 0;
                                        if (uCols >= catalogPriceColIndex && decimal.TryParse(Convert.ToString(data[r, catalogPriceColIndex]), out decimal mVal))
                                        {
                                            catalogPrice = mVal;
                                        }

                                        // 收集全行属性映射 (排除第 1 列 A、第 6 列 F 数量、第 30 列 AD 句柄)
                                        var rowCells = new Dictionary<int, string>();
                                        for (int c = 2; c <= Math.Min(uCols, 29); c++)
                                        {
                                            if (c == 6) continue;
                                            string cellVal = Convert.ToString(data[r, c])?.Trim() ?? string.Empty;
                                            if (!string.IsNullOrEmpty(cellVal))
                                            {
                                                rowCells[c] = cellVal;
                                            }
                                        }

                                        sheetComponents.Add(new SmartComponentItem
                                        {
                                            Model = model,
                                            Name = name,
                                            Manufacturer = mfr,
                                            Unit = unit,
                                            UnitPrice = unitPrice,
                                            CatalogPrice = catalogPrice,
                                            SheetName = sheetName,
                                            CabinetNo = "明细表",
                                            RowCells = rowCells
                                        });
                                    }
                                }
                            }
                        }
                    }

                    // 5. 【对当前工作表提取的所有物料进行精准去重与属性合并】
                    if (sheetComponents.Count > 0)
                    {
                        var uniqueDict = new Dictionary<string, SmartComponentItem>(StringComparer.OrdinalIgnoreCase);
                        foreach (var comp in sheetComponents)
                        {
                            if (!uniqueDict.ContainsKey(comp.Model))
                            {
                                uniqueDict[comp.Model] = comp;
                            }
                            else
                            {
                                // 若已有项缺少厂家或单价，则用更完整的数据补充合并
                                var existing = uniqueDict[comp.Model];
                                if (string.IsNullOrEmpty(existing.Manufacturer) && !string.IsNullOrEmpty(comp.Manufacturer))
                                {
                                    existing.Manufacturer = comp.Manufacturer;
                                }
                                if (existing.UnitPrice <= 0 && comp.UnitPrice > 0)
                                {
                                    existing.UnitPrice = comp.UnitPrice;
                                }
                                if (existing.CatalogPrice <= 0 && comp.CatalogPrice > 0)
                                {
                                    existing.CatalogPrice = comp.CatalogPrice;
                                }
                                if (string.IsNullOrEmpty(existing.QuoteCoefficient) && !string.IsNullOrEmpty(comp.QuoteCoefficient))
                                {
                                    existing.QuoteCoefficient = comp.QuoteCoefficient;
                                }
                                if (string.IsNullOrEmpty(existing.DiscountCoefficient) && !string.IsNullOrEmpty(comp.DiscountCoefficient))
                                {
                                    existing.DiscountCoefficient = comp.DiscountCoefficient;
                                }
                                if (string.IsNullOrEmpty(existing.Remark) && !string.IsNullOrEmpty(comp.Remark))
                                {
                                    existing.Remark = comp.Remark;
                                }
                                if (string.IsNullOrEmpty(existing.Category) && !string.IsNullOrEmpty(comp.Category))
                                {
                                    existing.Category = comp.Category;
                                }
                                if (string.IsNullOrEmpty(existing.Name) && !string.IsNullOrEmpty(comp.Name))
                                {
                                    existing.Name = comp.Name;
                                }
                                if (comp.RowCells != null && comp.RowCells.Count > 0)
                                {
                                    if (existing.RowCells == null) existing.RowCells = new Dictionary<int, string>();
                                    foreach (var kvp in comp.RowCells)
                                    {
                                        if (!existing.RowCells.ContainsKey(kvp.Key) || string.IsNullOrEmpty(existing.RowCells[kvp.Key]))
                                        {
                                            existing.RowCells[kvp.Key] = kvp.Value;
                                        }
                                    }
                                }
                            }
                        }

                        // 构造该工作表的元器件数据包
                        var sheetData = new SheetComponentData
                        {
                            SheetName = sheetName,
                            TotalCount = rawTotalCount,
                            UniqueCount = uniqueDict.Count,
                            Components = uniqueDict.Values.OrderBy(c => c.Model).ToList()
                        };

                        // 加入全局集合
                        storage.Sheets.Add(sheetData);
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录提取元器件异常日志
                LogHelper.WriteLog($"从所有工作表提取元器件数据发生异常: {ex.Message}");
            }

            return storage;
        }

        /// <summary>
        /// 为当前活动工作表所有箱柜的 C 列批量注入 Excel 原生下拉列表数据验证 (规则 6 & 规则 7)
        /// </summary>
        /// <param name="modelList">去重后的规格型号候选词列表</param>
        /// <returns>是否注入成功</returns>
        public static bool ApplySmartDropdownToActiveSheet(List<string> modelList)
        {
            // 校验规格型号列表
            if (modelList == null || modelList.Count == 0) return false;

            try
            {
                // 获取 Excel Application COM 接口 (安全调用)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return false;

                // 获取当前活动工作簿与工作表
                dynamic wb = app.ActiveWorkbook;
                dynamic activeSheet = app.ActiveSheet;
                if (wb == null || activeSheet == null) return false;

                // 准备或获取名为“选择表”的专用字典工作表 --硬编码: 选择表--
                string dictSheetName = "选择表";
                dynamic dictSheet = null;

                // 检查工作簿中是否已有选择表
                foreach (dynamic s in wb.Worksheets)
                {
                    if (string.Equals(s.Name, dictSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        dictSheet = s;
                        break;
                    }
                }

                // 若不存在则创建新的选择表工作表
                if (dictSheet == null)
                {
                    // 在最后一张工作表后面新建
                    dynamic lastSheet = wb.Worksheets[wb.Worksheets.Count];
                    dictSheet = wb.Worksheets.Add(After: lastSheet);
                    dictSheet.Name = dictSheetName;
                }

                // 构建 N+1 行 1 列的二维数据矩阵 (规则 7)
                int totalRows = modelList.Count + 1;
                object[,] matrix = new object[totalRows, 1];
                // A1 表头
                matrix[0, 0] = "规格型号";
                // A2 至 A(N+1) 数据
                for (int i = 0; i < modelList.Count; i++)
                {
                    matrix[i + 1, 0] = modelList[i];
                }

                // 批量一次性写入选择表 A 列 (规则 7)
                dictSheet.Range[$"A1:A{totalRows}"].Value2 = matrix;

                // 从配置中读取箱柜明细定义名称前缀值对象 (零堆分配)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;
                string currentSheetName = Convert.ToString(activeSheet.Name) ?? "";

                // 构建当前活动工作表的箱柜锚点字典 (复用 Tool 公共方法，内置空值自动智能补齐重建)
                var validCabinets = Tool.GetSheetValidCabinets(activeSheet, wb);

                // 扫描当前活动工作表中的所有箱柜元器件区域
                var cabinetRanges = new List<(int startRow, int endRow)>();

                foreach (var cab in validCabinets)
                {
                    if (cab.Value.Det != null && cab.Value.Subsum != null)
                    {
                        int detRow = Convert.ToInt32(cab.Value.Det.Row);
                        int subsumRow = Convert.ToInt32(cab.Value.Subsum.Row);

                        // 依据规则 6 界定元器件起始行与终止行
                        int startRow = detRow + 2;
                        int endRow = subsumRow - 1;
                        if (startRow <= endRow)
                        {
                            cabinetRanges.Add((startRow, endRow));
                        }
                    }
                }

                // 遍历每个箱柜的元器件区域，彻底清理 C 列已有的原生 Validation 数据验证小三角 (去除 Alt+箭头)
                foreach (var (startRow, endRow) in cabinetRanges)
                {
                    try
                    {
                        dynamic cRange = activeSheet.Range[$"C{startRow}:C{endRow}"];
                        // 清理已有的数据有效性箭头
                        cRange.Validation.Delete();
                    }
                    catch { }
                }

                return true;
            }
            catch (Exception ex)
            {
                // 记录数据同步异常
                LogHelper.WriteLog($"同步物料至选择表并清理原生下拉失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 将选中的元器件属性数据回填至当前活动单元格所在行 (用户铁律: AD句柄列绝对不回填，A序号列保持公式不回填，F数量不回填，其他列全量回填)
        /// </summary>
        /// <param name="item">元器件数据实体</param>
        /// <param name="config">回填字段配置选项</param>
        /// <returns>是否回填成功</returns>
        public static bool FillComponentToActiveRow(SmartComponentItem item, SmartInputConfigModel config)
        {
            if (item == null) return false;

            try
            {
                // 获取 Excel Application (安全调用)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return false;

                // 获取当前活动单元格及所在工作表
                dynamic activeCell = app.ActiveCell;
                dynamic activeSheet = app.ActiveSheet;
                if (activeCell == null || activeSheet == null) return false;

                // 获取当前活动单元格所在的物理行号
                int row = activeCell.Row;

                // 暂停 Excel 屏幕刷新与事件监听，确保批量赋值流畅且不递归触发
                app.ScreenUpdating = false;
                app.EnableEvents = false;

                try
                {
                    // 1. 回填 C 列 (规格型号, 永远写入)
                    activeSheet.Cells[row, 3].Value = item.Model;

                    // 2. 检查是否开启全字段自动回填 (默认开启)
                    if (config.FillAllColumns)
                    {
                        // 铁律 A: 绝对不回填第 1 列 (A 列序号公式保持不变)
                        // 铁律 B: 绝对不回填第 6 列 (F 列数量，保留当前行已有数量)
                        // 铁律 C: 绝对不回填第 30 列 (AD 列 CAD 实体句柄与图纸坐标绝对保护)

                        // 回填 B 列 (元件名称)
                        if (!string.IsNullOrEmpty(item.Name))
                        {
                            activeSheet.Cells[row, 2].Value = item.Name;
                        }

                        // 回填 D 列 (生产厂家)
                        if (!string.IsNullOrEmpty(item.Manufacturer))
                        {
                            activeSheet.Cells[row, 4].Value = item.Manufacturer;
                        }

                        // 回填 E 列 (计量单位)
                        if (!string.IsNullOrEmpty(item.Unit))
                        {
                            activeSheet.Cells[row, 5].Value = item.Unit;
                        }

                        // 回填 I 列 (规格备注)
                        if (!string.IsNullOrEmpty(item.Remark))
                        {
                            activeSheet.Cells[row, 9].Value = item.Remark;
                        }

                        // 回填 M 列 (表价值/基准面价，公式链条核心基石)
                        if (item.CatalogPrice > 0)
                        {
                            activeSheet.Cells[row, 13].Value = item.CatalogPrice;
                        }
                        else if (item.UnitPrice > 0)
                        {
                            // 若仅有单价无面价，存入 M 列表价以驱动 Excel 级联公式重算
                            activeSheet.Cells[row, 13].Value = item.UnitPrice;
                        }

                        // 回填 L 列 (报出加价系数)
                        if (!string.IsNullOrEmpty(item.QuoteCoefficient))
                        {
                            activeSheet.Cells[row, 12].Value = item.QuoteCoefficient;
                        }

                        // 回填 N 列 (采购折扣系数)
                        if (!string.IsNullOrEmpty(item.DiscountCoefficient))
                        {
                            activeSheet.Cells[row, 14].Value = item.DiscountCoefficient;
                        }

                        // 回填 Q 列 (元件类别)
                        if (!string.IsNullOrEmpty(item.Category))
                        {
                            activeSheet.Cells[row, 17].Value = item.Category;
                        }

                        // 安全回填 G 列 (销售单价): 优先保护已有公式，仅在无公式或静态值时写入单价
                        string gFormula = string.Empty;
                        try { gFormula = Convert.ToString(activeSheet.Cells[row, 7].Formula)?.Trim() ?? string.Empty; } catch { }
                        if (!gFormula.StartsWith("=") && item.UnitPrice > 0)
                        {
                            activeSheet.Cells[row, 7].Value = item.UnitPrice;
                        }

                        // 回填其余第 2 至 29 列自定义扩展属性 (跳过已处理核心列与受保护列)
                        if (item.RowCells != null && item.RowCells.Count > 0)
                        {
                            foreach (var kvp in item.RowCells)
                            {
                                int col = kvp.Key;
                                string val = kvp.Value;

                                // 严格过滤受保护列：第 1 列(A序号)、第 6 列(F数量)、第 30 列及以上(AD句柄及扩展)
                                if (col == 1 || col == 3 || col == 6 || col >= 30) continue;
                                // 过滤已在上文精准处理的核心列与公式列 (H 列为总价公式, J/K 列为成本公式)
                                if (col == 2 || col == 4 || col == 5 || col == 7 || col == 8 ||
                                    col == 9 || col == 10 || col == 11 || col == 12 || col == 13 || col == 14 || col == 17)
                                {
                                    continue;
                                }

                                if (!string.IsNullOrEmpty(val))
                                {
                                    // 写入该单元格
                                    activeSheet.Cells[row, col].Value = val;
                                }
                            }
                        }
                    }
                    else
                    {
                        // 用户显式关闭全回填时回退为按需字段勾选回填
                        if (config.FillName && !string.IsNullOrEmpty(item.Name))
                        {
                            activeSheet.Cells[row, 2].Value = item.Name;
                        }
                        if (config.FillManufacturer && !string.IsNullOrEmpty(item.Manufacturer))
                        {
                            activeSheet.Cells[row, 4].Value = item.Manufacturer;
                        }
                        if (config.FillUnit && !string.IsNullOrEmpty(item.Unit))
                        {
                            activeSheet.Cells[row, 5].Value = item.Unit;
                        }
                        if (config.FillUnitPrice && item.UnitPrice > 0)
                        {
                            activeSheet.Cells[row, 7].Value = item.UnitPrice;
                        }
                    }
                }
                finally
                {
                    // 恢复 Excel 屏幕刷新与事件响应
                    app.ScreenUpdating = true;
                    app.EnableEvents = true;
                }

                return true;
            }
            catch (Exception ex)
            {
                // 记录回填异常
                LogHelper.WriteLog($"回填元器件至活动行失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 判定当前行是否属于分类表的有效元器件行或有效计费区域行 (涵盖 Cab_Det+2 至 Cab_Tolsum)
        /// 确保元器件明细区与计费区域均能触发智能输入与新物料学习
        /// </summary>
        /// <param name="sheet">目标工作表 COM 句柄</param>
        /// <param name="row">目标物理行号</param>
        /// <returns>若处于元器件区或计费区返回 true</returns>
        public static bool IsCategoryComponentOrFeeRow(dynamic sheet, int row)
        {
            if (sheet == null || row <= 0) return false;

            try
            {
                string sheetName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(sheetName) || sheetName.StartsWith("_") ||
                    string.Equals(sheetName, "选择表", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(sheetName, ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                // 优先检查是否命中元器件行长效缓存 (0ms)
                if (IsCategoryComponentRow(sheet, row)) return true;

                // 若未命中元器件行，进一步判定是否处于计费区域 (Subsum 至 Tolsum)
                var validCabinets = Tool.GetSheetValidCabinets((object)sheet, wb: null, allowAutoRebuild: false);
                if (validCabinets != null && validCabinets.Count > 0)
                {
                    foreach (var kvp in validCabinets)
                    {
                        var anc = kvp.Value;
                        if (anc?.Subsum == null) continue;
                        int subR = 0, tolR = 0;
                        try
                        {
                            subR = Convert.ToInt32(anc.Subsum.Row);
                            tolR = anc.Tolsum != null ? Convert.ToInt32(anc.Tolsum.Row) : (subR + 10);
                        }
                        catch { continue; }

                        // 计费区域：从小计行 subsumRow 到总计行 tolsumRow
                        if (row >= subR && row <= tolR)
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"IsCategoryComponentOrFeeRow 判定异常: {ex.Message}");
            }

            return false;
        }

        // 方案 B: 单元格覆盖式智能输入窗体静态单例
        private static SmartInputOverlayForm? _overlayForm;

        /// <summary>
        /// 方案 B: 当活动单元格变化时，智能判断并在 C 列元器件行或计费区域行精确覆盖原生 TextBox+ListBox 输入控件 (规则 6 & 规则 7)
        /// 100% 还原 ZhiNengEn.ShuRu(Target) 的业务逻辑与交互行为
        /// </summary>
        /// <param name="activeCell">当前选中的活动单元格 COM 实例</param>
        /// <param name="isCategoryRowValidated">是否已经前置校验过属于有效元器件或计费区行 (默认 false，为 true 时跳过二次判定)</param>
        public static void ShuRu(dynamic activeCell, bool isCategoryRowValidated = false)
        {
            if (activeCell == null) return;

            try
            {
                // 1. 校验单元格是否为 C 列 (第 3 列: 规格型号)
                int col = 0;
                try { col = Convert.ToInt32(activeCell.Column); } catch { }
                if (col != 3)
                {
                    // 离开 C 列时隐藏覆盖输入框
                    HideSmartInputOverlay();
                    return;
                }

                // 2. 读取当前智能输入配置 (优先内存直出)
                var controller = new SmartInputController();
                var config = controller.GetConfig();
                // 若用户在配置中关闭了自动弹出，则直接退出
                if (!config.AutoPopupFloatWindow)
                {
                    HideSmartInputOverlay();
                    return;
                }

                // 3. 获取活动单元格物理行号与所属工作表
                int row = 0;
                try { row = Convert.ToInt32(activeCell.Row); } catch { }
                if (row <= 0) return;

                dynamic sheet = activeCell.Worksheet;
                if (sheet == null) return;
                string sheetName = Convert.ToString(sheet.Name) ?? "";

                // 过滤内部隐藏表与字典表
                if (sheetName.StartsWith("_") || sheetName.Equals("选择表", StringComparison.OrdinalIgnoreCase))
                {
                    HideSmartInputOverlay();
                    return;
                }

                // 4. 判定当前单元格行是否处于当前表箱柜元器件行或计费区域行 (涵盖 Cab_Det+2 至 Cab_Tolsum)
                bool isComponentOrFeeRow = isCategoryRowValidated;
                if (!isComponentOrFeeRow)
                {
                    // 复用 IsCategoryComponentOrFeeRow 判定，涵盖元器件明细区与计费区域
                    isComponentOrFeeRow = IsCategoryComponentOrFeeRow(sheet, row);
                }

                if (!isComponentOrFeeRow)
                {
                    // 离开元器件行或计费行时隐藏覆盖输入框
                    HideSmartInputOverlay();
                    return;
                }

                // 5. 获取并整理所有勾选数据源表中的元器件候选列表
                var storage = controller.GetStoredComponents();
                var selectedSheets = config.SelectedSheets ?? new List<string>();
                var candidateDict = new Dictionary<string, SmartComponentItem>(StringComparer.OrdinalIgnoreCase);

                foreach (var s in storage.Sheets)
                {
                    // 过滤仅保留勾选表
                    if ((selectedSheets.Count == 0 || selectedSheets.Contains(s.SheetName)) && s.Components != null)
                    {
                        foreach (var c in s.Components)
                        {
                            if (!string.IsNullOrWhiteSpace(c.Model) && !candidateDict.ContainsKey(c.Model))
                            {
                                candidateDict[c.Model] = c;
                            }
                        }
                    }
                }

                var candidateList = candidateDict.Values.OrderBy(c => c.Model).ToList();

                // 6. 激活覆盖输入窗体并定位到当前单元格上方
                if (_overlayForm == null || _overlayForm.IsDisposed)
                {
                    _overlayForm = new SmartInputOverlayForm();
                }

                _overlayForm.ShuRu(activeCell, candidateList, config);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"ShuRu 覆盖输入触发异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 隐藏方案 B 单元格覆盖输入窗体
        /// </summary>
        public static void HideSmartInputOverlay()
        {
            try
            {
                if (_overlayForm != null && !_overlayForm.IsDisposed && _overlayForm.Visible)
                {
                    _overlayForm.SafeHide();
                }
            }
            catch { }
        }

        /// <summary>
        /// 从分类明细表指定物理行中智能提取并学习新录入的物料数据 (元器件区与计费区全域覆盖)
        /// 遵循规则 7 向量批量读入内存数组，涵盖 B 列至 AC 列全量属性 (排除 AD 列句柄与 F 列数量)
        /// </summary>
        /// <param name="sheet">目标工作表 COM 句柄</param>
        /// <param name="row">发生变动的物理行号</param>
        /// <returns>若成功学习或更新返回 true，否则返回 false</returns>
        public static bool CheckAndLearnComponentFromRow(dynamic sheet, int row)
        {
            // 校验入参有效性
            if (sheet == null || row <= 0) return false;

            try
            {
                // 获取当前工作表名称
                string sheetName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;
                // 过滤隐藏表、字典表或选择表
                if (string.IsNullOrEmpty(sheetName) || sheetName.StartsWith("_") ||
                    string.Equals(sheetName, "选择表", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(sheetName, ComponentMatchDefaults.ComponentSummarySheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                // 规则 6 架构门控扩展: 处于有效箱柜元器件行或计费区域行 (Cab_Det+2 至 Cab_Tolsum)
                if (!IsCategoryComponentOrFeeRow(sheet, row))
                {
                    return false;
                }

                // 规则 7: 一次性读取该行 B~AC 列 (第 2 列至第 29 列，共 28 列，绝不读取第 30 列 AD 句柄)
                dynamic range = sheet.Range[$"B{row}:AC{row}"];
                // 转换为 object[,] 二维数组
                object[,] data = range.Value2 as object[,];
                // 校验数据矩阵非空
                if (data == null) return false;

                int cols = data.GetLength(1);
                // 提取 C 列 (规格型号, 对应相对列索引 2: 3-2+1=2)
                string model = cols >= 2 ? Convert.ToString(data[1, 2])?.Trim() ?? string.Empty : string.Empty;
                // 型号为空直接返回
                if (string.IsNullOrWhiteSpace(model)) return false;

                // 排除小计/总计纯文字行
                if (model.Equals("小计") || model.Equals("总计") || model.Equals("合计")) return false;

                // 提取 B 列 (元件名称, 对应相对列索引 1: 2-2+1=1)
                string name = cols >= 1 ? Convert.ToString(data[1, 1])?.Trim() ?? string.Empty : string.Empty;
                // 提取 D 列 (生产厂家, 对应相对列索引 3: 4-2+1=3)
                string mfr = cols >= 3 ? Convert.ToString(data[1, 3])?.Trim() ?? string.Empty : string.Empty;
                // 提取 E 列 (计量单位, 对应相对列索引 4: 5-2+1=4)
                string unit = cols >= 4 ? Convert.ToString(data[1, 4])?.Trim() ?? string.Empty : string.Empty;

                // 提取 G 列 (销售单价, 对应相对列索引 6: 7-2+1=6)
                decimal unitPrice = 0;
                if (cols >= 6 && decimal.TryParse(Convert.ToString(data[1, 6]), out decimal pVal))
                {
                    unitPrice = pVal;
                }

                // 提取 I 列 (规格备注, 对应相对列索引 8: 9-2+1=8)
                string remark = cols >= 8 ? Convert.ToString(data[1, 8])?.Trim() ?? string.Empty : string.Empty;
                // 提取 L 列 (加价系数, 对应相对列索引 11: 12-2+1=11)
                string quoteCoeff = cols >= 11 ? Convert.ToString(data[1, 11])?.Trim() ?? string.Empty : string.Empty;
                // 提取 M 列 (表价值, 对应相对列索引 12: 13-2+1=12)
                decimal catalogPrice = 0;
                if (cols >= 12 && decimal.TryParse(Convert.ToString(data[1, 12]), out decimal mVal))
                {
                    catalogPrice = mVal;
                }
                // 提取 N 列 (折扣系数, 对应相对列索引 13: 14-2+1=13)
                string discountCoeff = cols >= 13 ? Convert.ToString(data[1, 13])?.Trim() ?? string.Empty : string.Empty;
                // 提取 Q 列 (元件类别, 对应相对列索引 16: 17-2+1=16)
                string category = cols >= 16 ? Convert.ToString(data[1, 16])?.Trim() ?? string.Empty : string.Empty;

                // 收集整行有效属性映射 (排除第 1 列 A、第 6 列 F 数量、第 30 列 AD 句柄)
                var rowCells = new Dictionary<int, string>();
                for (int c = 1; c <= cols; c++)
                {
                    int absCol = c + 1; // 转换为工作表绝对列号 (B 列对应 2)
                    if (absCol == 6 || absCol >= 30) continue; // 跳过 F 数量列与 AD 句柄列
                    string cellVal = Convert.ToString(data[1, c])?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(cellVal))
                    {
                        rowCells[absCol] = cellVal;
                    }
                }

                // 构造智能元器件实体对象
                var newItem = new SmartComponentItem
                {
                    Model = model,
                    Name = name,
                    Manufacturer = mfr,
                    Unit = unit,
                    UnitPrice = unitPrice,
                    CatalogPrice = catalogPrice,
                    QuoteCoefficient = quoteCoeff,
                    DiscountCoefficient = discountCoeff,
                    Remark = remark,
                    Category = category,
                    SheetName = sheetName,
                    CabinetNo = $"行_{row}",
                    RowCells = rowCells
                };

                // 调用控制器执行增量学习与防抖持久化
                var controller = new Controllers.SmartInputController();
                // 尝试提交学习并返回结果
                return controller.TryLearnComponent(newItem);
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"CheckAndLearnComponentFromRow 异常: {ex.Message}");
                return false;
            }
        }
    }
}
