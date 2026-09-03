using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 公共服务分部类: 元器件数据管理（Excel 显示、筛选查询、选中行更新、删除、选中行新增）
    /// 遵循规范：所有 Excel COM 操作集中于此分部类中，所有区域读写采用二维数组一次性读入/写出
    /// </summary>
    public static partial class ExcelServices
    {
        // 窗体静态单例引用，避免重复打开
        private static Forms.ComponentManageForm? _componentManageFormInstance;

        // 默认列映射配置实例
        private static readonly ComponentManageColumnConfig _manageColCfg = new ComponentManageColumnConfig();

        /// <summary>
        /// 供 Ribbon 菜单调用的元器件数据管理入口：弹出非模态置顶操作浮窗
        /// </summary>
        public static void ShowComponentManageDialog()
        {
            try
            {
                // 获取当前正在运行的 Excel Application
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveWorkbook == null)
                {
                    // 若无工作簿打开则弹出提示
                    System.Windows.Forms.MessageBox.Show(
                        "请先打开或新建一个 Excel 工作簿！",
                        "系统提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // 以非模态方式展示元器件数据管理窗口，保证 Excel 处于可自由划选和编辑状态
                ShowModelessForm(ref _componentManageFormInstance, () => new Forms.ComponentManageForm());
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[ExcelServices] 弹出元器件管理窗口异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"弹出元器件管理窗口失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 确保并获取【元器件数据管理】工作表。若不存在则自动新建并初始化标准表头与列样式
        /// </summary>
        /// <param name="app">Excel Application COM dynamic 实例</param>
        /// <param name="wb">目标 Workbook COM dynamic 实例</param>
        /// <returns>准备就绪的 Worksheet COM dynamic 实例</returns>
        public static dynamic EnsureComponentManageWorksheet(dynamic app, dynamic wb)
        {
            string sheetName = ComponentManageDefaults.DefaultSheetName; // --硬编码-- 工作表名称

            // 1. 尝试寻找已存在的同名工作表
            foreach (dynamic sh in wb.Worksheets)
            {
                if (string.Equals(Convert.ToString(sh.Name), sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    sh.Activate();
                    return sh;
                }
            }

            // 2. 不存在则创建新工作表并置于末尾
            dynamic newSheet = wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            newSheet.Name = sheetName;
            newSheet.Activate();

            // 3. 一次性写入标准表头 (A1:L1)
            object[,] headers = new object[1, _manageColCfg.TotalColumns];
            headers[0, _manageColCfg.IdCol - 1] = "系统ID";
            headers[0, _manageColCfg.BrandCol - 1] = "品牌";
            headers[0, _manageColCfg.NameCol - 1] = "元器件名称";
            headers[0, _manageColCfg.ModelCol - 1] = "规格型号";
            headers[0, _manageColCfg.PriceCol - 1] = "参考单价(元)";
            headers[0, _manageColCfg.CurrentCol - 1] = "额定电流(A)";
            headers[0, _manageColCfg.PolesCol - 1] = "极数";
            headers[0, _manageColCfg.TrippingCol - 1] = "脱扣方式";
            headers[0, _manageColCfg.Param1Col - 1] = "扩展参数1";
            headers[0, _manageColCfg.Param2Col - 1] = "扩展参数2";
            headers[0, _manageColCfg.RemarkCol - 1] = "备注";
            headers[0, _manageColCfg.StatusCol - 1] = "操作状态";

            dynamic headerRange = newSheet.Range[newSheet.Cells[1, 1], newSheet.Cells[1, _manageColCfg.TotalColumns]];
            headerRange.Value2 = headers;

            // 4. 设置表头样式 (绿蓝主色调 #009688，白色粗体居中)
            try
            {
                headerRange.Interior.Color = ComponentManageDefaults.ThemeHeaderOleColor; // --硬编码-- 主色调
                headerRange.Font.Color = 0xFFFFFF; // 白色 --硬编码--
                headerRange.Font.Bold = true;
                headerRange.HorizontalAlignment = -4108; // xlHAlignCenter --硬编码--
                headerRange.VerticalAlignment = -4108; // xlVAlignCenter --硬编码--
                headerRange.RowHeight = 26;

                // 开启工作表自动筛选
                headerRange.AutoFilter(1);

                // 冻结首行表头
                app.ActiveWindow.SplitRow = 1;
                app.ActiveWindow.FreezePanes = true;

                // 设置各列标准默认列宽
                ((dynamic)newSheet.Columns[_manageColCfg.IdCol]).ColumnWidth = 10;
                ((dynamic)newSheet.Columns[_manageColCfg.BrandCol]).ColumnWidth = 14;
                ((dynamic)newSheet.Columns[_manageColCfg.NameCol]).ColumnWidth = 18;
                ((dynamic)newSheet.Columns[_manageColCfg.ModelCol]).ColumnWidth = 30;
                ((dynamic)newSheet.Columns[_manageColCfg.PriceCol]).ColumnWidth = 14;
                ((dynamic)newSheet.Columns[_manageColCfg.CurrentCol]).ColumnWidth = 12;
                ((dynamic)newSheet.Columns[_manageColCfg.PolesCol]).ColumnWidth = 10;
                ((dynamic)newSheet.Columns[_manageColCfg.TrippingCol]).ColumnWidth = 12;
                ((dynamic)newSheet.Columns[_manageColCfg.Param1Col]).ColumnWidth = 16;
                ((dynamic)newSheet.Columns[_manageColCfg.Param2Col]).ColumnWidth = 16;
                ((dynamic)newSheet.Columns[_manageColCfg.RemarkCol]).ColumnWidth = 20;
                ((dynamic)newSheet.Columns[_manageColCfg.StatusCol]).ColumnWidth = 22;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[ExcelServices] 初始化表头样式异常: {ex.Message}");
            }

            return newSheet;
        }

        /// <summary>
        /// 根据品牌与名称筛选条件拉取数据并一次性写入【元器件数据管理】工作表 (支持云端或本地 SQLite 个人库)
        /// </summary>
        /// <param name="brand">筛选品牌</param>
        /// <param name="nameKeyword">筛选名称关键字</param>
        /// <param name="dataSource">物料数据源 ("cloud" 或 "personal")</param>
        /// <returns>操作响应结果</returns>
        public static ComponentManageActionResult LoadComponentsToSheet(string? brand, string? nameKeyword, string dataSource = "cloud")
        {
            var result = new ComponentManageActionResult();
            try
            {
                bool isPersonal = string.Equals(dataSource, "personal", StringComparison.OrdinalIgnoreCase);

                // 1. 调用远程接口或本地 SQLite 个人物料库拉取满足条件的所有数据
                var items = isPersonal
                    ? Services.PersonalComponentDbService.QueryManageComponents(brand, nameKeyword)
                    : ComponentApiClient.QueryManageComponents(brand, nameKeyword);

                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveWorkbook == null)
                {
                    result.Success = false;
                    result.Message = "未检测到活跃的 Excel 工作簿！";
                    return result;
                }

                try
                {
                    app.ScreenUpdating = false;
                    app.EnableEvents = false;

                    dynamic sheet = EnsureComponentManageWorksheet(app, app.ActiveWorkbook);

                    // 清空旧数据行 (保留第 1 行表头)
                    int lastRow = sheet.Cells[sheet.Rows.Count, 1].End(-4162).Row; // xlUp = -4162
                    if (lastRow > 1)
                    {
                        dynamic clearRange = sheet.Range[sheet.Cells[2, 1], sheet.Cells[lastRow, _manageColCfg.TotalColumns]];
                        clearRange.Clear();
                    }

                    if (items == null || items.Count == 0)
                    {
                        result.Success = true;
                        result.SuccessCount = 0;
                        result.Message = isPersonal ? "未在本地个人库中查询到符合条件的元器件数据。" : "未在云端库中查询到符合条件的元器件数据。";
                        return result;
                    }

                    // 3. 构建二维数组以供一次性灌入 (零逐单元格 COM 损耗)
                    int rowCount = items.Count;
                    int colCount = _manageColCfg.TotalColumns;
                    object[,] dataArray = new object[rowCount, colCount];

                    string statusTag = isPersonal ? "已同步(个人库)" : "已同步(云端)"; // --硬编码-- 同步标识

                    for (int i = 0; i < rowCount; i++)
                    {
                        var item = items[i];
                        dataArray[i, _manageColCfg.IdCol - 1] = item.Id;
                        dataArray[i, _manageColCfg.BrandCol - 1] = item.Brand ?? string.Empty;
                        dataArray[i, _manageColCfg.NameCol - 1] = item.Name ?? string.Empty;
                        dataArray[i, _manageColCfg.ModelCol - 1] = item.Model ?? string.Empty;
                        dataArray[i, _manageColCfg.PriceCol - 1] = item.Price;
                        dataArray[i, _manageColCfg.CurrentCol - 1] = item.Current.HasValue ? item.Current.Value.ToString() : string.Empty;
                        dataArray[i, _manageColCfg.PolesCol - 1] = item.Poles ?? string.Empty;
                        dataArray[i, _manageColCfg.TrippingCol - 1] = item.Tripping ?? string.Empty;
                        dataArray[i, _manageColCfg.Param1Col - 1] = item.Param1 ?? string.Empty;
                        dataArray[i, _manageColCfg.Param2Col - 1] = item.Param2 ?? string.Empty;
                        dataArray[i, _manageColCfg.RemarkCol - 1] = item.Remark ?? string.Empty;
                        dataArray[i, _manageColCfg.StatusCol - 1] = statusTag;
                    }

                    // 4. 一次性灌入数据区域
                    dynamic targetRange = sheet.Range[sheet.Cells[2, 1], sheet.Cells[1 + rowCount, colCount]];
                    targetRange.Value2 = dataArray;

                    // 5. 设置格式：单价格式两位小数、ID 列浅灰居中
                    try
                    {
                        dynamic idRange = sheet.Range[sheet.Cells[2, _manageColCfg.IdCol], sheet.Cells[1 + rowCount, _manageColCfg.IdCol]];
                        idRange.Interior.Color = ComponentManageDefaults.IdColumnLightGrayOleColor;
                        idRange.HorizontalAlignment = -4108; // xlHAlignCenter --硬编码--

                        dynamic priceRange = sheet.Range[sheet.Cells[2, _manageColCfg.PriceCol], sheet.Cells[1 + rowCount, _manageColCfg.PriceCol]];
                        priceRange.NumberFormatLocal = "0.00";

                        dynamic statusRange = sheet.Range[sheet.Cells[2, _manageColCfg.StatusCol], sheet.Cells[1 + rowCount, _manageColCfg.StatusCol]];
                        statusRange.Font.Color = 0x2E7D32; // 绿色已同步 --硬编码--
                    }
                    catch { }

                    result.Success = true;
                    result.SuccessCount = rowCount;
                    result.Message = isPersonal
                        ? $"成功拉取并呈现 {rowCount} 条【本地个人物料库】数据！"
                        : $"成功拉取并呈现 {rowCount} 条【云端公共物料库】数据！";
                }
                finally
                {
                    app.ScreenUpdating = true;
                    app.EnableEvents = true;
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"拉取元器件数据异常: {ex.Message}";
                LogHelper.WriteLog($"[ExcelServices] LoadComponentsToSheet 异常: {ex.Message}");
            }
            return result;
        }

        /// <summary>
        /// 探测当前 Excel 选区中覆盖的有效数据行号
        /// </summary>
        public static SelectionDetectResult DetectCurrentSelection()
        {
            var res = new SelectionDetectResult();
            try
            {
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveWorkbook == null || app.ActiveSheet == null)
                {
                    res.Message = "未检测到活跃工作表";
                    return res;
                }

                dynamic activeSheet = app.ActiveSheet;
                res.SheetName = Convert.ToString(activeSheet.Name) ?? "";
                res.IsInManageSheet = string.Equals(res.SheetName, ComponentManageDefaults.DefaultSheetName, StringComparison.OrdinalIgnoreCase);

                dynamic? selection = app.Selection;
                if (selection == null)
                {
                    res.Message = "当前未选中任何单元格区域";
                    return res;
                }

                // 提取所有选中的物理行号（自动排除第 1 行表头）
                var rowsSet = new SortedSet<int>();
                foreach (dynamic area in selection.Areas)
                {
                    int startRow = area.Row;
                    int endRow = area.Row + area.Rows.Count - 1;
                    for (int r = startRow; r <= endRow; r++)
                    {
                        if (r > 1) // 排除第 1 行表头
                        {
                            rowsSet.Add(r);
                        }
                    }
                }

                res.RowIndices = rowsSet.ToList();
                res.SelectedRowCount = res.RowIndices.Count;
                res.Message = res.SelectedRowCount > 0 ? $"已选中 {res.SelectedRowCount} 行数据" : "当前未选中有效数据行";
            }
            catch (Exception ex)
            {
                res.Message = $"选区探测异常: {ex.Message}";
            }
            return res;
        }

        /// <summary>
        /// 对当前选中的 1 行或多行元器件执行【更新】保存操作 (支持个人物料库 SQLite 或云端 WebAPI)
        /// </summary>
        /// <param name="dataSource">物料数据源 ("cloud" 或 "personal")</param>
        public static ComponentManageActionResult UpdateSelectedComponents(string dataSource = "cloud")
        {
            var result = new ComponentManageActionResult();
            try
            {
                bool isPersonal = string.Equals(dataSource, "personal", StringComparison.OrdinalIgnoreCase);

                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveSheet == null)
                {
                    result.Success = false;
                    result.Message = "未检测到活跃的 Excel 工作表！";
                    return result;
                }

                dynamic sheet = app.ActiveSheet;
                var detect = DetectCurrentSelection();
                if (detect.SelectedRowCount == 0)
                {
                    result.Success = false;
                    result.Message = "请先在表格中用鼠标选中 1 行或多行需要更新的元器件！";
                    return result;
                }

                int successCount = 0;
                int failCount = 0;
                var updateResults = new List<(int row, bool success, string msg)>();

                // 1. 一次性读取选中行的数据
                foreach (int row in detect.RowIndices)
                {
                    dynamic rowRange = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, _manageColCfg.TotalColumns]];
                    object[,] rowValues = (object[,])rowRange.Value2;

                    // 提取主键 ID (必须存在有效 ID 才能执行 Update)
                    string idStr = Convert.ToString(rowValues[1, _manageColCfg.IdCol])?.Trim() ?? string.Empty;
                    if (!int.TryParse(idStr, out int id) || id <= 0)
                    {
                        failCount++;
                        updateResults.Add((row, false, "缺少有效系统ID，无法更新"));
                        continue;
                    }

                    string brand = Convert.ToString(rowValues[1, _manageColCfg.BrandCol])?.Trim() ?? string.Empty;
                    string name = Convert.ToString(rowValues[1, _manageColCfg.NameCol])?.Trim() ?? string.Empty;
                    string model = Convert.ToString(rowValues[1, _manageColCfg.ModelCol])?.Trim() ?? string.Empty;

                    // 校验必填字段
                    if (string.IsNullOrWhiteSpace(brand) || string.IsNullOrWhiteSpace(model))
                    {
                        failCount++;
                        updateResults.Add((row, false, "品牌或型号不能为空"));
                        continue;
                    }

                    // 解析价格
                    decimal price = 0;
                    if (decimal.TryParse(Convert.ToString(rowValues[1, _manageColCfg.PriceCol]), out decimal parsedPrice))
                    {
                        price = parsedPrice;
                    }

                    // 组装 Update DTO
                    var updateDto = new UpdateComponentApiRequest
                    {
                        Id = id,
                        Brand = brand,
                        Name = name,
                        Model = model,
                        Price = price,
                        Current = ComponentApiClient.ExtractIntegerCurrent(Convert.ToString(rowValues[1, _manageColCfg.CurrentCol])),
                        Poles = ComponentApiClient.NormalizePolesParam(Convert.ToString(rowValues[1, _manageColCfg.PolesCol])),
                        Tripping = Convert.ToString(rowValues[1, _manageColCfg.TrippingCol])?.Trim(),
                        Param1 = Convert.ToString(rowValues[1, _manageColCfg.Param1Col])?.Trim(),
                        Param2 = Convert.ToString(rowValues[1, _manageColCfg.Param2Col])?.Trim(),
                        Remark = Convert.ToString(rowValues[1, _manageColCfg.RemarkCol])?.Trim()
                    };

                    // 根据当前数据源执行更新
                    bool isOk = false;
                    if (isPersonal)
                    {
                        isOk = Services.PersonalComponentDbService.UpdateComponents(new List<UpdateComponentApiRequest> { updateDto }) > 0;
                    }
                    else
                    {
                        var updated = ComponentApiClient.UpdateComponent(updateDto);
                        isOk = updated != null;
                    }

                    if (isOk)
                    {
                        successCount++;
                        string tag = isPersonal ? "个人库" : "云端";
                        updateResults.Add((row, true, $"✓ 更新成功({tag}) {DateTime.Now:HH:mm:ss}"));
                    }
                    else
                    {
                        failCount++;
                        string errTag = isPersonal ? "❌ 本地个人库更新失败" : "❌ 云端更新失败";
                        updateResults.Add((row, false, errTag));
                    }
                }

                // 2. 回写状态列
                try
                {
                    app.ScreenUpdating = false;
                    app.EnableEvents = false;
                    foreach (var item in updateResults)
                    {
                        dynamic statusCell = sheet.Cells[item.row, _manageColCfg.StatusCol];
                        statusCell.Value2 = item.msg;
                        statusCell.Font.Color = item.success ? 0x2E7D32 : 0x0000FF; // 成功绿色，失败红色 --硬编码--
                    }
                }
                finally
                {
                    app.ScreenUpdating = true;
                    app.EnableEvents = true;
                }

                result.Success = failCount == 0;
                result.SuccessCount = successCount;
                result.FailCount = failCount;
                result.Message = $"选中行更新完成：成功 {successCount} 条，失败 {failCount} 条。";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"更新选中行异常: {ex.Message}";
                LogHelper.WriteLog($"[ExcelServices] UpdateSelectedComponents 异常: {ex.Message}");
            }
            return result;
        }

        /// <summary>
        /// 对当前选中的 1 行或多行内容执行【新增】保存操作，并自动将新分配的主键 ID 回填至 A 列 (支持个人库 SQLite 或云端 WebAPI)
        /// </summary>
        /// <param name="dataSource">物料数据源 ("cloud" 或 "personal")</param>
        public static ComponentManageActionResult CreateSelectedComponents(string dataSource = "cloud")
        {
            var result = new ComponentManageActionResult();
            try
            {
                bool isPersonal = string.Equals(dataSource, "personal", StringComparison.OrdinalIgnoreCase);

                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveSheet == null)
                {
                    result.Success = false;
                    result.Message = "未检测到活跃的 Excel 工作表！";
                    return result;
                }

                dynamic sheet = app.ActiveSheet;
                var detect = DetectCurrentSelection();
                if (detect.SelectedRowCount == 0)
                {
                    result.Success = false;
                    result.Message = "请先在表格中用鼠标选中 1 行或多行需要新增的元器件内容！";
                    return result;
                }

                int successCount = 0;
                int failCount = 0;
                var createResults = new List<(int row, bool success, int? newId, string msg)>();

                // 1. 读取选中行的数据 (忽略原先是否有 ID，作为全新物料新增)
                foreach (int row in detect.RowIndices)
                {
                    dynamic rowRange = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, _manageColCfg.TotalColumns]];
                    object[,] rowValues = (object[,])rowRange.Value2;

                    string brand = Convert.ToString(rowValues[1, _manageColCfg.BrandCol])?.Trim() ?? string.Empty;
                    string name = Convert.ToString(rowValues[1, _manageColCfg.NameCol])?.Trim() ?? string.Empty;
                    string model = Convert.ToString(rowValues[1, _manageColCfg.ModelCol])?.Trim() ?? string.Empty;

                    // 校验必填字段
                    if (string.IsNullOrWhiteSpace(brand) || string.IsNullOrWhiteSpace(model))
                    {
                        failCount++;
                        createResults.Add((row, false, null, "品牌和型号为必填项"));
                        continue;
                    }

                    // 解析单价
                    decimal price = 0;
                    if (decimal.TryParse(Convert.ToString(rowValues[1, _manageColCfg.PriceCol]), out decimal parsedPrice))
                    {
                        price = parsedPrice;
                    }

                    var createDto = new CreateComponentApiRequest
                    {
                        Brand = brand,
                        Name = name,
                        Model = model,
                        Price = price,
                        Current = ComponentApiClient.ExtractIntegerCurrent(Convert.ToString(rowValues[1, _manageColCfg.CurrentCol])),
                        Poles = ComponentApiClient.NormalizePolesParam(Convert.ToString(rowValues[1, _manageColCfg.PolesCol])),
                        Tripping = Convert.ToString(rowValues[1, _manageColCfg.TrippingCol])?.Trim(),
                        Param1 = Convert.ToString(rowValues[1, _manageColCfg.Param1Col])?.Trim(),
                        Param2 = Convert.ToString(rowValues[1, _manageColCfg.Param2Col])?.Trim(),
                        Remark = Convert.ToString(rowValues[1, _manageColCfg.RemarkCol])?.Trim()
                    };

                    // 调用 API 或本地 SQLite 新增
                    ComponentApiDto? created = null;
                    if (isPersonal)
                    {
                        var createdList = Services.PersonalComponentDbService.CreateComponents(new List<CreateComponentApiRequest> { createDto });
                        created = createdList.FirstOrDefault();
                    }
                    else
                    {
                        created = ComponentApiClient.CreateComponent(createDto);
                    }

                    if (created != null && created.Id > 0)
                    {
                        successCount++;
                        string tag = isPersonal ? "个人库" : "云端";
                        createResults.Add((row, true, created.Id, $"✓ 新增成功({tag} ID:{created.Id})"));
                    }
                    else
                    {
                        failCount++;
                        string errTag = isPersonal ? "❌ 本地个人库新增失败" : "❌ 云端新增失败";
                        createResults.Add((row, false, null, errTag));
                    }
                }

                // 2. 回写新主键 ID 到 A 列，并更新 L 列状态
                try
                {
                    app.ScreenUpdating = false;
                    app.EnableEvents = false;
                    foreach (var item in createResults)
                    {
                        if (item.success && item.newId.HasValue)
                        {
                            dynamic idCell = sheet.Cells[item.row, _manageColCfg.IdCol];
                            idCell.Value2 = item.newId.Value;
                            idCell.Interior.Color = ComponentManageDefaults.IdColumnLightGrayOleColor;
                            idCell.HorizontalAlignment = -4108; // xlHAlignCenter --硬编码--
                        }

                        dynamic statusCell = sheet.Cells[item.row, _manageColCfg.StatusCol];
                        statusCell.Value2 = item.msg;
                        statusCell.Font.Color = item.success ? 0x2E7D32 : 0x0000FF; // 绿色/红色 --硬编码--
                    }
                }
                finally
                {
                    app.ScreenUpdating = true;
                    app.EnableEvents = true;
                }

                result.Success = failCount == 0;
                result.SuccessCount = successCount;
                result.FailCount = failCount;
                string srcName = isPersonal ? "本地个人库" : "云端";
                result.Message = $"选中行新增完成：成功 {successCount} 条，失败 {failCount} 条 ({srcName})。";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"新增选中行异常: {ex.Message}";
                LogHelper.WriteLog($"[ExcelServices] CreateSelectedComponents 异常: {ex.Message}");
            }
            return result;
        }

        /// <summary>
        /// 对当前选中的 1 行或多行执行【删除】(同步调用个人库 SQLite 或云端接口并在 Excel 中移除整行)
        /// </summary>
        /// <param name="dataSource">物料数据源 ("cloud" 或 "personal")</param>
        public static ComponentManageActionResult DeleteSelectedComponents(string dataSource = "cloud")
        {
            var result = new ComponentManageActionResult();
            try
            {
                bool isPersonal = string.Equals(dataSource, "personal", StringComparison.OrdinalIgnoreCase);

                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveSheet == null)
                {
                    result.Success = false;
                    result.Message = "未检测到活跃的 Excel 工作表！";
                    return result;
                }

                dynamic sheet = app.ActiveSheet;
                var detect = DetectCurrentSelection();
                if (detect.SelectedRowCount == 0)
                {
                    result.Success = false;
                    result.Message = "请先在表格中用鼠标选中 1 行或多行需要删除的元器件！";
                    return result;
                }

                int successCount = 0;
                int failCount = 0;
                var deleteSuccessRows = new List<int>();

                // 从后向前扫描（便于后续从大到小删除行不乱序）
                var sortedRowsDesc = detect.RowIndices.OrderByDescending(r => r).ToList();

                foreach (int row in sortedRowsDesc)
                {
                    dynamic idCell = sheet.Cells[row, _manageColCfg.IdCol];
                    string idStr = Convert.ToString(idCell.Value2)?.Trim() ?? string.Empty;

                    // 若无有效 ID，说明尚未存入数据库，直接在本地移除行
                    if (!int.TryParse(idStr, out int id) || id <= 0)
                    {
                        successCount++;
                        deleteSuccessRows.Add(row);
                        continue;
                    }

                    // 调用本地个人库或云端接口执行硬删除
                    bool deleted = false;
                    if (isPersonal)
                    {
                        deleted = Services.PersonalComponentDbService.DeleteComponents(new List<int> { id }) > 0;
                    }
                    else
                    {
                        deleted = ComponentApiClient.DeleteComponent(id);
                    }

                    if (deleted)
                    {
                        successCount++;
                        deleteSuccessRows.Add(row);
                    }
                    else
                    {
                        failCount++;
                        dynamic statusCell = sheet.Cells[row, _manageColCfg.StatusCol];
                        statusCell.Value2 = isPersonal ? "❌ 本地个人库删除失败" : "❌ 云端删除失败";
                        statusCell.Font.Color = 0x0000FF; // 红色 --硬编码--
                    }
                }

                // 在 Excel 中整行物理删除成功被删掉的行 (从大到小删)
                try
                {
                    app.ScreenUpdating = false;
                    app.EnableEvents = false;
                    foreach (int r in deleteSuccessRows)
                    {
                        dynamic delRow = sheet.Rows[r];
                        delRow.Delete(-4162); // xlShiftUp = -4162 --硬编码--
                    }
                }
                finally
                {
                    app.ScreenUpdating = true;
                    app.EnableEvents = true;
                }

                result.Success = failCount == 0;
                result.SuccessCount = successCount;
                result.FailCount = failCount;
                string delSrcName = isPersonal ? "本地个人物料库" : "云端公共库";
                result.Message = $"选中行删除完成：从【{delSrcName}】及 Excel 中成功移除 {successCount} 条记录，失败 {failCount} 条。";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"删除选中行异常: {ex.Message}";
                LogHelper.WriteLog($"[ExcelServices] DeleteSelectedComponents 异常: {ex.Message}");
            }
            return result;
        }
    }
}
