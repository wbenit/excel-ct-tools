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

                // 从配置文件读取箱柜定义名称前缀 (规则 6)
                string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";
                string subsumPrefix = ConfigManager.Instance.Current.Excel.SubsumNamePrefix ?? "Cab_Subsum_";
                string tolsumPrefix = ConfigManager.Instance.Current.Excel.TolsumNamePrefix ?? "Cab_Tolsum_";

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

                    // 收集当前工作表中的所有箱柜元器件有效行区间 (startRow, endRow)
                    var cabinetRanges = new List<(int startRow, int endRow)>();

                    // 1. 获取当前工作表有效箱柜映射 (复用 Tool 公共方法，内置空值自动智能补齐重建)
                    var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);

                    // 若解析到有效箱柜锚点，计算各箱柜元器件有效插槽行 (规则 6)
                    if (validCabinets.Count > 0)
                    {
                        foreach (var cab in validCabinets)
                        {
                            if (cab.Value.Det != null && cab.Value.Subsum != null)
                            {
                                int detRow = Convert.ToInt32(cab.Value.Det.Row);
                                int subsumRow = Convert.ToInt32(cab.Value.Subsum.Row);

                                // 依据规则 6: Cab_Det + 2 为元器件起始行，Cab_Subsum - 1 为元器件终止行
                                int startRow = detRow + 2;
                                int endRow = subsumRow - 1;

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

                    // 3. 【执行元器件数据读取：优先从箱柜区域批量读取，规则 7 数组入内存】
                    if (cabinetRanges.Count > 0)
                    {
                        foreach (var (startRow, endRow) in cabinetRanges)
                        {
                            // 规则 7: 一次性读取二维数据矩阵
                            dynamic range = sheet.Range[$"A{startRow}:Q{endRow}"];
                            object[,] data = range.Value2 as object[,];

                            if (data != null)
                            {
                                int rowCount = data.GetLength(0);
                                int colCount = data.GetLength(1);

                                for (int r = 1; r <= rowCount; r++)
                                {
                                    // 读取 C 列 (规格型号, 列索引 3)
                                    string model = colCount >= 3 ? Convert.ToString(data[r, 3])?.Trim() ?? string.Empty : string.Empty;

                                    // 过滤规格型号为空的行
                                    if (!string.IsNullOrWhiteSpace(model))
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

                                        // 读取 J 列 (成本单价, 列索引 10)
                                        decimal costPrice = 0;
                                        if (colCount >= 10 && decimal.TryParse(Convert.ToString(data[r, 10]), out decimal cVal))
                                        {
                                            costPrice = cVal;
                                        }

                                        // 读取 Q 列 (元件类别, 列索引 17)
                                        string category = colCount >= 17 ? Convert.ToString(data[r, 17])?.Trim() ?? string.Empty : string.Empty;

                                        var item = new SmartComponentItem
                                        {
                                            Model = model,
                                            Name = name,
                                            Manufacturer = mfr,
                                            Unit = unit,
                                            UnitPrice = unitPrice,
                                            CostUnitPrice = costPrice,
                                            Category = category,
                                            SheetName = sheetName,
                                            CabinetNo = $"箱柜_{startRow}"
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

                                        sheetComponents.Add(new SmartComponentItem
                                        {
                                            Model = model,
                                            Name = name,
                                            Manufacturer = mfr,
                                            Unit = unit,
                                            UnitPrice = unitPrice,
                                            SheetName = sheetName,
                                            CabinetNo = "明细表"
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
                                if (string.IsNullOrEmpty(existing.Name) && !string.IsNullOrEmpty(comp.Name))
                                {
                                    existing.Name = comp.Name;
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

                // 从配置中读取箱柜明细定义名称前缀 (规则 6)
                string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";
                string subsumPrefix = ConfigManager.Instance.Current.Excel.SubsumNamePrefix ?? "Cab_Subsum_";
                string tolsumPrefix = ConfigManager.Instance.Current.Excel.TolsumNamePrefix ?? "Cab_Tolsum_";
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
        /// 将选中的元器件属性数据回填至当前活动单元格所在行
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

                    // 2. 依据配置决定是否回填 B 列 (元件名称)
                    if (config.FillName && !string.IsNullOrEmpty(item.Name))
                    {
                        activeSheet.Cells[row, 2].Value = item.Name;
                    }

                    // 3. 依据配置决定是否回填 D 列 (生产厂家)
                    if (config.FillManufacturer && !string.IsNullOrEmpty(item.Manufacturer))
                    {
                        activeSheet.Cells[row, 4].Value = item.Manufacturer;
                    }

                    // 4. 依据配置决定是否回填 E 列 (计量单位)
                    if (config.FillUnit && !string.IsNullOrEmpty(item.Unit))
                    {
                        activeSheet.Cells[row, 5].Value = item.Unit;
                    }

                    // 5. 依据配置决定是否回填 G 列 (销售单价)
                    if (config.FillUnitPrice && item.UnitPrice > 0)
                    {
                        activeSheet.Cells[row, 7].Value = item.UnitPrice;
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

        // 方案 B: 单元格覆盖式智能输入窗体静态单例
        private static SmartInputOverlayForm? _overlayForm;

        /// <summary>
        /// 方案 B: 当活动单元格变化时，智能判断并在 C 列元器件行精确覆盖原生 TextBox+ListBox 输入控件 (规则 6 & 规则 7)
        /// 100% 还原 ZhiNengEn.ShuRu(Target) 的业务逻辑与交互行为
        /// </summary>
        /// <param name="activeCell">当前选中的活动单元格 COM 实例</param>
        public static void ShuRu(dynamic activeCell)
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

                // 2. 读取当前智能输入配置
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

                // 4. 判定当前单元格行是否处于当前表箱柜元器件插槽行 (Cab_Det+2 至 Cab_Subsum-1)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return;
                dynamic wb = app.ActiveWorkbook;
                if (wb == null) return;

                string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";
                string subsumPrefix = ConfigManager.Instance.Current.Excel.SubsumNamePrefix ?? "Cab_Subsum_";
                string tolsumPrefix = ConfigManager.Instance.Current.Excel.TolsumNamePrefix ?? "Cab_Tolsum_";

                // 构建当前工作表箱柜字典 (复用 Tool 公共方法，内置空值自动智能补齐重建)
                var validCabinets = Tool.GetSheetValidCabinets(sheet, wb);

                // 判断是否落在某个箱柜的元器件行区间内
                bool isComponentRow = false;
                if (validCabinets.Count > 0)
                {
                    foreach (var cab in validCabinets)
                    {
                        if (cab.Value.Det != null && cab.Value.Subsum != null)
                        {
                            int detRow = Convert.ToInt32(cab.Value.Det.Row);
                            int subsumRow = Convert.ToInt32(cab.Value.Subsum.Row);
                            int startRow = detRow + 2;
                            int endRow = subsumRow - 1;

                            if (row >= startRow && row <= endRow)
                            {
                                isComponentRow = true;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    // 若无定义名称但行号大于 1 则宽松兼容支持输入联想
                    if (row > 1) isComponentRow = true;
                }

                if (!isComponentRow)
                {
                    // 离开元器件行时隐藏覆盖输入框
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
    }
}
