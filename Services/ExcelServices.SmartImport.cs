using System;
using System.Collections.Generic;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 业务服务分部类：智能导入箱柜 BOM
    /// 严格遵循规则 6 (四大定义名称与计费区无空行)、规则 7 (数组一次性读写) 和规则 8 (事前校验纠偏)
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 弹出基于 WebView2 + Vue 3 的智能导入箱柜 BOM 模态窗口
        /// </summary>
        public static void ShowSmartImportDialog()
        {
            try
            {
                // 获取当前活动 Excel 上下文
                var context = Tool.GetActiveExcelContext();
                if (context == null || context.App == null)
                {
                    System.Windows.Forms.MessageBox.Show("无法连接 Excel 实例，请先打开有效工作簿！", "系统提示",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // 实例化基于 WebView2 的智能导入向导窗体
                using var form = new Forms.SmartImportForm();
                // 提取 Excel 主窗口 HWND 句柄，模态挂载防止焦点走失与闪退
                IntPtr excelHwnd = (IntPtr)context.App.Hwnd;
                form.ShowDialog(new ExcelWin32Window(excelHwnd));
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[ShowSmartImportDialog] 弹窗异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"打开智能导入工作台失败: {ex.Message}", "错误",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 执行智能导入核心业务：新建独立分类表并批量写入箱柜与元器件数据
        /// 采用本表接力母版复制 + 批量数组读写 + 手动计算模式，彻底消除 O(N^2) COM 卡顿
        /// </summary>
        /// <param name="request">导入请求参数</param>
        /// <param name="progressCallback">可选实时进度回调 (当前柜序, 总柜数, 当前箱柜名)</param>
        /// <returns>执行结果对象</returns>
        public static SmartImportExecuteResult ExecuteSmartImportToNewCategory(
            SmartImportExecuteRequest request,
            Action<int, int, string>? progressCallback = null)
        {
            var result = new SmartImportExecuteResult();
            // 校验基本参数有效性
            if (request == null || string.IsNullOrWhiteSpace(request.NewCategoryName))
            {
                result.Success = false;
                result.Message = "新建分类表名称不能为空！";
                return result;
            }

            try
            {
                // 1. 获取当前 Excel 运行环境与上下文
                var context = Tool.GetActiveExcelContext();
                if (context == null)
                {
                    result.Success = false;
                    result.Message = "无法连接当前 Excel 运行上下文！";
                    return result;
                }

                dynamic app = context.App;
                dynamic activeWb = context.Wb;
                string newCategoryName = request.NewCategoryName.Trim();

                // 2. 检查或全量解析待导入的箱柜数据
                List<ParsedCabinetModel> cabinets = request.Cabinets;
                if (cabinets == null || cabinets.Count == 0)
                {
                    // 若前端未传入全量箱柜，则调用引擎在后端全量解析
                    cabinets = SmartImportParserService.ParseAllCabinets(request.FilePath, request.SheetName, request.Config);
                }

                if (cabinets == null || cabinets.Count == 0)
                {
                    result.Success = false;
                    result.Message = "未能从选定的工作表中解析出任何有效箱柜数据，请检查点选规则！";
                    return result;
                }

                // 备份进入批量写入前的 Excel 宿主环境参数 (包含自动计算模式)
                bool prevScreenUpdating = true;
                bool prevDisplayAlerts = true;
                bool prevEnableEvents = true;
                bool prevAskToUpdateLinks = true;
                int prevCalculation = -4105; // 默认 xlCalculationAutomatic
                try { prevScreenUpdating = app.ScreenUpdating; } catch { }
                try { prevDisplayAlerts = app.DisplayAlerts; } catch { }
                try { prevEnableEvents = app.EnableEvents; } catch { }
                try { prevAskToUpdateLinks = app.AskToUpdateLinks; } catch { }
                try { prevCalculation = Convert.ToInt32(app.Calculation); } catch { }

                // 全程锁定极限静默性能模式：关闭屏幕刷新、告警、事件与自动重算 (xlCalculationManual=-4135)
                try { app.ScreenUpdating = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }
                try { app.EnableEvents = false; } catch { }
                try { app.Calculation = -4135; } catch { }
                try { app.AskToUpdateLinks = false; } catch { }

                try
                {
                    // 3. 调用公共分类服务创建全新的纯净分类表 Sheet (自动克隆模板并挂接项目信息主表)
                    var createCatReq = new CreateCategoryRequest
                    {
                        CategoryName = newCategoryName,
                        InitialCabinetName = cabinets[0].CabinetName
                    };
                    var catRes = CreateNewCategory(createCatReq, app);
                    if (!catRes.Success)
                    {
                        result.Success = false;
                        result.Message = $"创建分类表【{newCategoryName}】失败: {catRes.Message}";
                        return result;
                    }

                    // 4. 定位新创建的分类表 Sheet 句柄
                    dynamic newSheet = activeWb.Worksheets[newCategoryName];
                    if (newSheet == null)
                    {
                        result.Success = false;
                        result.Message = $"未找到新创建的分类表【{newCategoryName}】！";
                        return result;
                    }

                    // 再次确保静默模式与手动重算未被子方法修改
                    try { app.ScreenUpdating = false; } catch { }
                    try { app.EnableEvents = false; } catch { }
                    try { app.Calculation = -4135; } catch { }

                    // 5. 读取模板拓扑参数并精准定位刚初始化好的第 1 个纯净箱柜行号
                    var cfg = ConfigManager.Instance.Current.Excel;
                    int headerRows = cfg.DetailHeaderRows > 0 ? cfg.DetailHeaderRows : 3;
                    int footerRows = cfg.DetailFooterRows > 0 ? cfg.DetailFooterRows : 3;
                    var (sPre, dPre, subPre, tolPre) = cfg.Prefixes;

                    CabinetCreatedInfo firstCab = LocateFirstCabinetInfo(newSheet);
                    int baseCabinetK = firstCab.CabinetK > 0 ? firstCab.CabinetK : 1;
                    int baseSumRow = firstCab.SumRow;
                    int baseDetRow = firstCab.DetRow;
                    int baseSubsumRow = firstCab.SubsumRow;
                    int baseTolsumRow = firstCab.TolsumRow;

                    // 规则 6 前置清理：仅在复制前对第 1 柜纯净计费区执行 1 次空行清理，后续接力复制的箱柜天然无空行
                    CleanEmptyRowsInBillingZone(newSheet, baseSubsumRow, ref baseTolsumRow);

                    int totalCabinets = cabinets.Count;
                    int firstBlockStart = Math.Max(1, baseDetRow - headerRows);
                    int sumInsertCount = 0;

                    // 6. 顶部汇总区一次性扩容与样式复制 (最多 1 次 Insert + 1 次 Copy，彻底消除循环逐行插行推移)
                    if (totalCabinets > 1)
                    {
                        EnsureTopSummaryCapacityAndStyle(newSheet, baseSumRow, firstBlockStart, totalCabinets, out sumInsertCount);
                        // 若顶部汇总区执行了物理插行，同步顺延第 1 柜明细基准物理行号
                        if (sumInsertCount > 0)
                        {
                            baseDetRow += sumInsertCount;
                            baseSubsumRow += sumInsertCount;
                            baseTolsumRow += sumInsertCount;
                            firstBlockStart += sumInsertCount;
                        }
                    }

                    // 7. 一次性读取默认公式组详情，用于内存构建各柜计费区二维矩阵
                    var feeController = new Controllers.FormulaAdjustFeeController();
                    var defaultFeeItems = feeController.GetDefaultGroup()?.Details ?? new List<FormulaItemModel>();

                    // 计算纯净母版明细块净高度与内部相对偏移量
                    int blockLen = (baseTolsumRow + footerRows) - firstBlockStart + 1;
                    int detOffset = baseDetRow - firstBlockStart;
                    int subsumOffset = baseSubsumRow - firstBlockStart;
                    int tolsumOffset = baseTolsumRow - firstBlockStart;
                    int defaultCap = (baseSubsumRow - 1) - (baseDetRow + 2) + 1;

                    int curBlockStart = firstBlockStart;
                    int importedCompTotal = 0;
                    // 记录所有生成箱柜的拓扑坐标与基础字段，用于最终批量回写汇总区、定义名称与超链接
                    var createdRecords = new List<CabinetWriteRecord>(totalCabinets);
                    // 判断是否配置了可选箱柜名称/型号/备注锚点 (TargetCol > 0 时覆写对应列，未配置时保留模板原样)
                    bool hasNameAnchor = request.Config?.CabinetNameAnchor?.TargetCol > 0;
                    bool hasModelAnchor = request.Config?.CabinetModelAnchor?.TargetCol > 0;
                    bool hasRemarkAnchor = request.Config?.CabinetRemarkAnchor?.TargetCol > 0;
                    var extraCols = request.Config?.ColumnMapping?.ExtraColumns;

                    // 8. 底部明细区接力棒复制与逐柜矩阵写入 (全程下方仅有 1 个空骨架，线性 O(N) 复杂度)
                    for (int i = 0; i < totalCabinets; i++)
                    {
                        var cab = cabinets[i];
                        int cabinetK = baseCabinetK + i;
                        int sumRow = baseSumRow + i;

                        // 若后方仍有待导入箱柜，先将当前位置未写入数据的纯净骨架接力复制到紧邻的下一段空白行
                        if (i < totalCabinets - 1)
                        {
                            int curTemplateEnd = curBlockStart + blockLen - 1;
                            int nextBlockStart = curTemplateEnd + 1;
                            newSheet.Rows[$"{curBlockStart}:{curTemplateEnd}"].Copy(
                                newSheet.Rows[$"{nextBlockStart}:{nextBlockStart + blockLen - 1}"]);
                        }

                        // 推导当前箱柜的初始关键行号
                        int detRow = curBlockStart + detOffset;
                        int subsumRow = curBlockStart + subsumOffset;
                        int tolsumRow = curBlockStart + tolsumOffset;
                        int compStartRow = detRow + 2;
                        int actualCompCount = cab.Components != null ? cab.Components.Count : 0;

                        // 规则 6：若元器件数量多于当前容量，在小计行前一次性插入物理差额行 (下方仅 1 个空骨架，0 负担)
                        if (actualCompCount > defaultCap)
                        {
                            int insertRowsCount = actualCompCount - defaultCap;
                            newSheet.Rows[$"{subsumRow}:{subsumRow + insertRowsCount - 1}"].Insert(-4121);
                            subsumRow += insertRowsCount;
                            tolsumRow += insertRowsCount;
                        }

                        int compEndRow = subsumRow - 1;

                        // 写入明细表头柜号 (B列)；仅在配置了对应锚点时写入名称(G列)、型号(D列)、备注(I列)
                        newSheet.Cells[detRow, 2].Value2 = cab.CabinetName;
                        if (hasNameAnchor)
                        {
                            // 将箱柜名称写入明细 det 行的 G 列 (第 7 列)
                            newSheet.Cells[detRow, 7].Value2 = cab.CabinetDisplayName ?? string.Empty;
                        }
                        if (hasModelAnchor)
                        {
                            // 将箱柜型号写入明细 det 行的 D 列 (第 4 列)
                            newSheet.Cells[detRow, 4].Value2 = cab.CabinetModel ?? string.Empty;
                        }
                        // 将箱柜备注写入明细 det 行的 I 列 (第 9 列)；未配置锚点时清空模板残留
                        newSheet.Cells[detRow, 9].Value2 = hasRemarkAnchor ? (cab.CabinetRemark ?? string.Empty) : string.Empty;

                        // 规则 7：组装并一次性数组批量写入当前柜 A~U 列元器件矩阵
                        if (compEndRow >= compStartRow)
                        {
                            var compItemList = new List<ComponentItem>(actualCompCount);
                            if (cab.Components != null)
                            {
                                // 组装元器件标准模型集合
                                foreach (var pc in cab.Components)
                                {
                                    compItemList.Add(new ComponentItem
                                    {
                                        Name = pc.ItemName,
                                        Specification = pc.ItemSpec,
                                        Manufacturer = pc.Brand,
                                        Quantity = pc.Quantity,
                                        // 将外部单价写入 M 列面价
                                        UnitPrice = pc.MarkedPrice,
                                        Unit = pc.Unit,
                                        Remark = pc.Remark
                                    });
                                }
                            }

                            // 一次性写入 A~U 列标准公式与数据矩阵
                            object[,] compMatrix = Tool.BuildComponentRowsMatrix(compStartRow, compEndRow, detRow, 21, compItemList);
                            newSheet.Range[$"A{compStartRow}:U{compEndRow}"].Formula = compMatrix;
                            importedCompTotal += actualCompCount;

                            // 若配置了动态扩展列，逐列批量数组写入对应目标列 (支持 A~AF 列)
                            if (extraCols != null && extraCols.Count > 0)
                            {
                                for (int ei = 0; ei < extraCols.Count; ei++)
                                {
                                    WriteExtraColumnIfConfigured(newSheet, compStartRow, actualCompCount, cab.Components, extraCols[ei], ei);
                                }
                            }
                        }

                        // 刷新当前箱柜计费区公式矩阵 (修正复制产生的绝对行引用如 A$45，并将总计行 F 列绑定至 =F{sumRow})
                        if (defaultFeeItems.Count > 0 && (tolsumRow - subsumRow + 1) == defaultFeeItems.Count)
                        {
                            // 利用内存二维数组一次性写入整个计费区 A~Q 列
                            object[,] feeMatrix = Tool.BuildFeeMatrix(defaultFeeItems, detRow, subsumRow, compStartRow, compEndRow, 17, sumRow, $"=F{sumRow}");
                            newSheet.Range[$"A{subsumRow}:Q{tolsumRow}"].Formula = feeMatrix;
                        }
                        else
                        {
                            // 非标准计费区行数时，使用紧凑数组辅助刷新小计/总计核心公式
                            RefreshBillingZoneCompact(newSheet, detRow, compStartRow, subsumRow, tolsumRow, sumRow);
                        }

                        // 记录本柜已写入完毕的最终拓扑坐标与字段状态
                        createdRecords.Add(new CabinetWriteRecord
                        {
                            CabK = cabinetK,
                            SumRow = sumRow,
                            DetRow = detRow,
                            SubsumRow = subsumRow,
                            TolsumRow = tolsumRow,
                            CabName = cab.CabinetName,
                            // 同步保存名称、型号、备注及是否启用覆写标志
                            CabDisplayName = cab.CabinetDisplayName ?? string.Empty,
                            CabModel = cab.CabinetModel ?? string.Empty,
                            CabRemark = cab.CabinetRemark ?? string.Empty,
                            CabQty = cab.CabinetQuantity,
                            HasName = hasNameAnchor,
                            HasModel = hasModelAnchor,
                            HasRemark = hasRemarkAnchor
                        });
                        // 推进游标至下一台箱柜明细块起点
                        curBlockStart = tolsumRow + footerRows + 1;

                        // 每处理 5 台箱柜或最后 1 台时向前端推送实时进度通知
                        if (progressCallback != null && (i % 5 == 0 || i == totalCabinets - 1))
                        {
                            progressCallback(i + 1, totalCabinets, cab.CabinetName);
                        }
                    }

                    // 9. 一次性批量注册所有箱柜的 4 个定义名称 (预读取名称字典避免 COMException 与冗余日志)
                    BatchRegisterSheetDefinedNames(newSheet, newCategoryName, sPre, dPre, subPre, tolPre, createdRecords);

                    // 10. 批量挂载或原地更新汇总行与明细行双向超链接及明细表头序号公式 (杜绝调用 Hyperlinks.Delete 破坏边框)
                    for (int i = 0; i < createdRecords.Count; i++)
                    {
                        var rec = createdRecords[i];
                        string sumNameTag = $"{sPre}{rec.CabK}";
                        string detNameTag = $"{dPre}{rec.CabK}";
                        try
                        {
                            // 汇总行 A 列超链接原地更新/跳转至明细行
                            SetCellHyperlinkPreserveStyle(
                                newSheet,
                                newSheet.Cells[rec.SumRow, 1],
                                $"'{newCategoryName}'!{detNameTag}",
                                "点击进入本箱柜明细表",
                                null);

                            // 明细行 A 列超链接原地更新/返回顶部汇总行
                            SetCellHyperlinkPreserveStyle(
                                newSheet,
                                newSheet.Cells[rec.DetRow, 1],
                                $"'{newCategoryName}'!{sumNameTag}",
                                "返回汇总行",
                                "柜号:");

                            // 明细表头行 (detRow + 1) A 列绑定汇总行序号动态公式
                            newSheet.Cells[rec.DetRow + 1, 1].Formula = $"=\"序号\" & {sumNameTag}";
                        }
                        catch { }
                    }

                    // 11. 规则 7：一次性二维数组批量回写顶部汇总区 A~L 列 (保留超链接同时 1 次写入全部箱柜汇总公式)
                    BatchWriteTopSummaryMatrix(newSheet, baseSumRow, createdRecords);

                    // 失效新工作表短效箱柜缓存
                    Tool.InvalidateSheetCabinetsCache(newCategoryName);

                    // 12. 若顶部汇总区执行了物理插行导致汇总合计行下移，仅针对新建表同步刷新【项目信息】主表引用行号
                    if (sumInsertCount > 0)
                    {
                        UpdateProjectInfoCategorySummary(activeWb, newCategoryName);
                    }

                    // 13. 激活新工作表并置顶展示
                    newSheet.Activate();
                    try
                    {
                        if (activeWb.Windows.Count > 0)
                        {
                            activeWb.Windows[1].Visible = true;
                            activeWb.Windows[1].WindowState = -4137; // xlMaximized
                        }
                    }
                    catch { }

                    result.Success = true;
                    result.CreatedCategoryName = newCategoryName;
                    result.ImportedCabinetCount = totalCabinets;
                    result.ImportedComponentCount = importedCompTotal;
                    result.Message = $"成功创建分类表【{newCategoryName}】，共导入 {totalCabinets} 台箱柜、{importedCompTotal} 条元器件！";
                }
                finally
                {
                    // 严格按安全顺序恢复用户原有的 Excel 宿主设置 (恢复 Calculation 时 Excel 自动完成 1 次全表精确计算)
                    try { app.Calculation = prevCalculation; } catch { }
                    try { app.ScreenUpdating = prevScreenUpdating; } catch { }
                    try { app.DisplayAlerts = prevDisplayAlerts; } catch { }
                    try { app.EnableEvents = prevEnableEvents; } catch { }
                    try { app.AskToUpdateLinks = prevAskToUpdateLinks; } catch { }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"导入执行失败: {ex.Message}";
                LogHelper.WriteLog($"[ExecuteSmartImportToNewCategory] 异常: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 顶部汇总区一次性容量探测、批量插行与样式复制 (最多 1 次 Insert + 1 次 Copy)
        /// </summary>
        private static void EnsureTopSummaryCapacityAndStyle(
            dynamic sheet,
            int baseSumRow,
            int firstBlockStart,
            int totalCabinets,
            out int insertedRowCount)
        {
            insertedRowCount = 0;
            int neededExtra = totalCabinets - 1;
            if (neededExtra <= 0) return;

            int scanEnd = firstBlockStart - 1;
            int firstBlockerRow = firstBlockStart;

            // 规则 7：一次性将顶部汇总候选区域 A~C 列读入内存数组，探测首个非空/合计阻断行
            if (scanEnd >= baseSumRow + 1)
            {
                try
                {
                    dynamic scanRange = sheet.Range[sheet.Cells[baseSumRow + 1, 1], sheet.Cells[scanEnd, 3]];
                    object? raw = scanRange.Value2;
                    if (raw is object[,] arr)
                    {
                        int rows = arr.GetLength(0);
                        for (int r = 1; r <= rows; r++)
                        {
                            string colA = Convert.ToString(arr[r, 1])?.Trim() ?? "";
                            string colB = Convert.ToString(arr[r, 2])?.Trim() ?? "";
                            string colC = Convert.ToString(arr[r, 3])?.Trim() ?? "";
                            // 遇到合计/小计/说明或已有内容即判定为汇总区边界
                            if (colA.Contains("小计") || colB.Contains("小计") || colA.Contains("合计") || colB.Contains("合计") ||
                                colA.Contains("总计") || colB.Contains("总计") || colB.Contains("包装") || colB.Contains("运输") ||
                                colA.Contains("明细表") || colB.Contains("明细表") || colA.Contains("说明") || colB.Contains("说明") ||
                                !string.IsNullOrWhiteSpace(colB) || !string.IsNullOrWhiteSpace(colC))
                            {
                                firstBlockerRow = baseSumRow + r;
                                break;
                            }
                        }
                    }
                }
                catch { }
            }

            // 计算当前可复用空白预留行数与需物理插入的差额行数
            int freeRows = Math.Max(0, firstBlockerRow - (baseSumRow + 1));
            if (neededExtra > freeRows)
            {
                insertedRowCount = neededExtra - freeRows;
                // 在第 1 柜汇总行正下方一次性插入差额行，确保下方 SUM(...) 汇总范围自动无缝扩展
                sheet.Rows[$"{baseSumRow + 1}:{baseSumRow + insertedRowCount}"].Insert(-4121);
            }

            // 仅执行 1 次整行复制，将第 1 柜汇总行完整边框、字体与行高同步铺满第 2~N 柜汇总行
            // 注意：禁止调用 Hyperlinks.Delete()，否则 Excel 会把 A 列原有边框样式重置清除；复制带的超链接在后续原地更新 SubAddress
            sheet.Rows[baseSumRow].Copy(sheet.Rows[$"{baseSumRow + 1}:{baseSumRow + neededExtra}"]);
        }

        /// <summary>
        /// 原地设置或更新单元格超链接 (若已有超链接则直接修改 SubAddress，绝不调用 Hyperlinks.Delete 破坏单元格边框)
        /// </summary>
        private static void SetCellHyperlinkPreserveStyle(
            dynamic sheet,
            dynamic cell,
            string subAddress,
            string screenTip,
            string? textToDisplay)
        {
            // 检查当前单元格是否已经存在复制或模板遗留的超链接
            if (cell.Hyperlinks.Count > 0)
            {
                // 直接原地改写已有超链接的跳转地址与提示，100% 保留原有单元格边框样式
                dynamic existingHl = cell.Hyperlinks[1];
                existingHl.Address = "";
                existingHl.SubAddress = subAddress;
                existingHl.ScreenTip = screenTip;
                if (textToDisplay != null)
                {
                    existingHl.TextToDisplay = textToDisplay;
                }
                return;
            }

            // 单元格尚无超链接时安全新增
            if (textToDisplay != null)
            {
                sheet.Hyperlinks.Add(
                    Anchor: cell,
                    Address: "",
                    SubAddress: subAddress,
                    ScreenTip: screenTip,
                    TextToDisplay: textToDisplay);
            }
            else
            {
                sheet.Hyperlinks.Add(
                    Anchor: cell,
                    Address: "",
                    SubAddress: subAddress,
                    ScreenTip: screenTip);
            }
        }

        /// <summary>
        /// 单台箱柜拓扑坐标与扩展字段写入记录
        /// </summary>
        private sealed class CabinetWriteRecord
        {
            /// <summary>箱柜定义名称编号</summary>
            public int CabK { get; set; }
            /// <summary>顶部汇总行号</summary>
            public int SumRow { get; set; }
            /// <summary>底部明细头行号</summary>
            public int DetRow { get; set; }
            /// <summary>小计行号</summary>
            public int SubsumRow { get; set; }
            /// <summary>总计行号</summary>
            public int TolsumRow { get; set; }
            /// <summary>箱柜柜号文本</summary>
            public string CabName { get; set; } = string.Empty;
            /// <summary>箱柜名称文本</summary>
            public string CabDisplayName { get; set; } = string.Empty;
            /// <summary>箱柜型号文本</summary>
            public string CabModel { get; set; } = string.Empty;
            /// <summary>箱柜备注文本</summary>
            public string CabRemark { get; set; } = string.Empty;
            /// <summary>箱柜台数</summary>
            public decimal CabQty { get; set; }
            /// <summary>是否配置了箱柜名称锚点</summary>
            public bool HasName { get; set; }
            /// <summary>是否配置了箱柜型号锚点</summary>
            public bool HasModel { get; set; }
            /// <summary>是否配置了箱柜备注锚点</summary>
            public bool HasRemark { get; set; }
        }

        /// <summary>
        /// 规则 7：一次性构建顶部汇总区 A~L 列二维矩阵并单次写入 Excel
        /// </summary>
        private static void BatchWriteTopSummaryMatrix(
            dynamic sheet,
            int baseSumRow,
            List<CabinetWriteRecord> records)
        {
            if (records == null || records.Count == 0) return;

            int rowCount = records.Count;
            // 先读取首行模板原始公式/默认值，保留未配置可选列的原始模板内容
            object[,]? baseRowArr = null;
            try { baseRowArr = sheet.Range[$"A{baseSumRow}:L{baseSumRow}"].Formula as object[,]; } catch { }

            object[,] sumMatrix = new object[rowCount, 12];
            for (int i = 0; i < rowCount; i++)
            {
                var rec = records[i];
                // 继承首行模板默认列内容
                for (int c = 0; c < 12; c++)
                {
                    sumMatrix[i, c] = (baseRowArr != null && c + 1 <= baseRowArr.GetLength(1)) ? (baseRowArr[1, c + 1] ?? string.Empty) : string.Empty;
                }

                // 写入汇总行核心数据与公式 (A序号, B柜号, E单位, F台数, G单价, H总价, J成本, K毛利, L毛利率)
                sumMatrix[i, 0] = "=ROW()-ROW(A$6)";
                sumMatrix[i, 1] = rec.CabName;
                // 若启用了箱柜名称锚点，覆写汇总行 C 列 (第 3 列，索引 2)
                if (rec.HasName) sumMatrix[i, 2] = rec.CabDisplayName;
                // 若启用了箱柜型号锚点，覆写汇总行 D 列 (第 4 列，索引 3)
                if (rec.HasModel) sumMatrix[i, 3] = rec.CabModel;
                sumMatrix[i, 4] = "台";
                sumMatrix[i, 5] = rec.CabQty > 0 ? rec.CabQty : 1m;
                sumMatrix[i, 6] = $"=G{rec.TolsumRow}";
                sumMatrix[i, 7] = $"=F{rec.SumRow}*G{rec.SumRow}";
                // 若启用了箱柜备注锚点，覆写汇总行 I 列 (第 9 列，索引 8)
                if (rec.HasRemark) sumMatrix[i, 8] = rec.CabRemark;
                sumMatrix[i, 9] = $"=K{rec.TolsumRow}";
                sumMatrix[i, 10] = $"=H{rec.SumRow}-J{rec.SumRow}";
                sumMatrix[i, 11] = $"=IF(H{rec.SumRow}=0,0,K{rec.SumRow}/H{rec.SumRow})";
            }

            // 一次性将全部箱柜汇总行写回工作表
            sheet.Range[$"A{baseSumRow}:L{baseSumRow + rowCount - 1}"].Formula = sumMatrix;
        }

        /// <summary>
        /// 预加载已有名称字典后批量注册工作表级别 4 定义名称，避免 4*N 次 COMException 抛出与磁盘日志写入
        /// </summary>
        private static void BatchRegisterSheetDefinedNames(
            dynamic sheet,
            string sheetName,
            string sPre,
            string dPre,
            string subPre,
            string tolPre,
            List<CabinetWriteRecord> records)
        {
            // 先一次性枚举当前工作表既有名称构建查找字典
            var existingMap = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (dynamic n in sheet.Names)
                {
                    try
                    {
                        string rawName = Convert.ToString(n.Name) ?? "";
                        int bangIdx = rawName.IndexOf('!');
                        string shortName = bangIdx >= 0 ? rawName.Substring(bangIdx + 1) : rawName;
                        shortName = shortName.Trim('\'', '"', ' ');
                        if (!string.IsNullOrEmpty(shortName))
                        {
                            existingMap[shortName] = n;
                        }
                    }
                    catch { }
                }
            }
            catch { }

            // 局部辅助：零异常精准设置或更新定义名称
            void SetNameFast(string tagName, int targetRow)
            {
                string targetRef = $"='{sheetName}'!$A${targetRow}";
                try
                {
                    if (existingMap.TryGetValue(tagName, out dynamic? existing))
                    {
                        string curRef = Convert.ToString(existing.RefersTo) ?? "";
                        if (curRef.EndsWith($"$A${targetRow}", StringComparison.OrdinalIgnoreCase))
                        {
                            return;
                        }
                        existing.Delete();
                    }
                    sheet.Names.Add(tagName, targetRef);
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[BatchRegisterSheetDefinedNames] 注册 {tagName} 异常: {ex.Message}");
                }
            }

            // 批量依次注册每台箱柜的 4 个标准定义名称
            foreach (var rec in records)
            {
                SetNameFast($"{sPre}{rec.CabK}", rec.SumRow);
                SetNameFast($"{dPre}{rec.CabK}", rec.DetRow);
                SetNameFast($"{subPre}{rec.CabK}", rec.SubsumRow);
                SetNameFast($"{tolPre}{rec.CabK}", rec.TolsumRow);
            }
        }

        /// <summary>
        /// 紧凑版计费区刷新辅助 (仅在非标准公式组行数时兜底使用，仅读写本柜计费区几行)
        /// </summary>
        private static void RefreshBillingZoneCompact(
            dynamic sheet,
            int detRow,
            int compStartRow,
            int subsumRow,
            int tolsumRow,
            int sumRow)
        {
            try
            {
                int feeRows = tolsumRow - subsumRow;
                if (feeRows > 0)
                {
                    // 读取计费区 B 列定位小计行并批量修正 A 列序号公式
                    object[,]? bArr = sheet.Range[$"B{subsumRow}:B{tolsumRow - 1}"].Value2 as object[,];
                    object[,]? aRaw = sheet.Range[$"A{subsumRow}:A{tolsumRow - 1}"].Formula as object[,];
                    object[,] aNew = new object[feeRows, 1];
                    int actualSubsumRow = subsumRow;

                    for (int r = 0; r < feeRows; r++)
                    {
                        string bVal = (bArr != null && r + 1 <= bArr.GetLength(0)) ? (Convert.ToString(bArr[r + 1, 1])?.Trim() ?? "") : "";
                        if (bVal.Contains("小计")) actualSubsumRow = subsumRow + r;
                        string aFormula = (aRaw != null && r + 1 <= aRaw.GetLength(0)) ? (Convert.ToString(aRaw[r + 1, 1])?.Trim() ?? "") : "";
                        aNew[r, 0] = (aFormula.IndexOf("ROW(", StringComparison.OrdinalIgnoreCase) >= 0 || string.IsNullOrEmpty(aFormula))
                            ? $"=ROW()-ROW(A${detRow + 1})"
                            : aFormula;
                    }
                    sheet.Range[$"A{subsumRow}:A{tolsumRow - 1}"].Formula = aNew;
                    // 刷新小计行 H/K 列求和公式
                    sheet.Cells[actualSubsumRow, 8].Formula = $"=ROUND(SUM(H{compStartRow}:INDEX(H:H,ROW()-1)),2)";
                    sheet.Cells[actualSubsumRow, 11].Formula = $"=ROUND(SUM(K{compStartRow}:INDEX(K:K,ROW()-1)),2)";
                }
                // 确保总计行 A 列标签与 F 列台数联动
                sheet.Cells[tolsumRow, 1].Value2 = "总计";
                sheet.Cells[tolsumRow, 6].Formula = $"=F{sumRow}";
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[RefreshBillingZoneCompact] 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 从全新模板克隆的工作表中定位第 1 个初始箱柜的关键行号
        /// 遵循规则 6 架构，优先提取真实定义名称，杜绝硬编码行号偏位
        /// </summary>
        private static CabinetCreatedInfo LocateFirstCabinetInfo(dynamic sheet)
        {
            // 从全局配置读取标准基准物理行号 (CabDetRowIndex=44, CabTolsumRowIndex=71)
            var cfg = ConfigManager.Instance.Current.Excel;
            int defSum = cfg.CabSumRowIndex > 0 ? cfg.CabSumRowIndex : 7;
            int defDet = cfg.CabDetRowIndex > 0 ? cfg.CabDetRowIndex : 44; // 真实的箱柜信息行基准行号
            int defTol = cfg.CabTolsumRowIndex > 0 ? cfg.CabTolsumRowIndex : 71; // 真实的总计行基准行号
            int defSub = defTol - 5; // 真实的小计行基准行号

            // 构建默认兜底实体 (采用母版真实物理行号)
            var info = new CabinetCreatedInfo
            {
                CabinetK = 1,
                SumRow = defSum,
                DetRow = defDet,
                SubsumRow = defSub,
                TolsumRow = defTol
            };

            try
            {
                // 1. 优先调用 Tool.GetSheetValidCabinets 提取工作表现存首个有效箱柜 (覆盖工作表与工作簿双作用域)
                var validCabinets = Tool.GetSheetValidCabinets((object)sheet, (object)sheet.Parent);
                if (validCabinets != null && validCabinets.Count > 0)
                {
                    // 提取首个有效箱柜实体
                    var firstCab = validCabinets[0];
                    info.CabinetK = firstCab.Key;
                    var anc = firstCab.Value;
                    // 精确绑定各关键行物理行号
                    if (anc.Sum != null) info.SumRow = Convert.ToInt32(anc.Sum.Row);
                    if (anc.Det != null) info.DetRow = Convert.ToInt32(anc.Det.Row);
                    if (anc.Tolsum != null) info.TolsumRow = Convert.ToInt32(anc.Tolsum.Row);
                    if (anc.Subsum != null) info.SubsumRow = Convert.ToInt32(anc.Subsum.Row);
                    else if (info.TolsumRow > 0) info.SubsumRow = info.TolsumRow - 5;
                    return info;
                }

                // 2. 次优通过 Tool.FindStandardCategoryRowIndexes 动态嗅探首个箱柜行号分布
                var (sR, dR, subR, tR) = Tool.FindStandardCategoryRowIndexes((object)sheet, 1);
                if (dR > 0)
                {
                    // 赋值动态识别到的行号
                    info.SumRow = sR > 0 ? sR : defSum;
                    info.DetRow = dR;
                    info.SubsumRow = subR > 0 ? subR : defSub;
                    info.TolsumRow = tR > 0 ? tR : defTol;
                    return info;
                }
            }
            catch (Exception ex)
            {
                // 记录动态定位异常日志并安全回退
                LogHelper.WriteLog($"[LocateFirstCabinetInfo] 动态定位首柜异常: {ex.Message}");
            }

            return info;
        }

        /// <summary>
        /// 清理计费区域的空行 (规则 6: 计费区域不能有空行，若有必须删除；规则 7: 数组一次性读取 B:C 列)
        /// </summary>
        private static void CleanEmptyRowsInBillingZone(dynamic sheet, int subsumRow, ref int tolsumRow)
        {
            try
            {
                // 计费区域有效校验范围为 SubsumRow + 1 到 TolsumRow - 1
                int startRow = subsumRow + 1;
                int endRow = tolsumRow - 1;
                if (endRow < startRow) return;

                // 规则 7：一次性将计费区 B~C 列读入二维数组，避免逐格 COM 往返
                object? raw = sheet.Range[$"B{startRow}:C{endRow}"].Value2;
                if (raw is object[,] arr)
                {
                    int rows = arr.GetLength(0);
                    // 自底向上逆向删除空行，避免行号偏移
                    for (int r = rows; r >= 1; r--)
                    {
                        string bVal = Convert.ToString(arr[r, 1])?.Trim() ?? "";
                        string cVal = Convert.ToString(arr[r, 2])?.Trim() ?? "";
                        if (string.IsNullOrEmpty(bVal) && string.IsNullOrEmpty(cVal))
                        {
                            int physRow = startRow + r - 1;
                            sheet.Rows[physRow].Delete(-4121);
                            tolsumRow--; // 总计行上浮 1 行
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CleanEmptyRowsInBillingZone] 清理计费空行异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 方案 B：根据用户配置的扩展列映射，将额外提取的数据批量写入导入表目标列 (规则 7 数组极速覆写)
        /// 支持写入 A~AF 等任意目标列
        /// </summary>
        /// <param name="sheet">目标工作表句柄</param>
        /// <param name="startRow">元器件起始物理行号</param>
        /// <param name="compCount">实际有效元器件数量</param>
        /// <param name="components">已解析的元器件明细集合</param>
        /// <param name="extraCol">扩展列映射配置对象</param>
        /// <param name="colSlotIndex">扩展列槽位索引 (0-based，最多支持 8 列)</param>
        private static void WriteExtraColumnIfConfigured(
            dynamic sheet,
            int startRow,
            int compCount,
            List<ParsedComponentModel>? components,
            ExtraColumnMapping? extraCol,
            int colSlotIndex)
        {
            // 校验扩展列配置有效性：源列号与目标列号均必须大于 0 且存在有效组件
            if (extraCol == null || extraCol.SourceCol <= 0 || extraCol.TargetCol <= 0) return;
            // 校验行号与列表有效性
            if (components == null || compCount <= 0 || startRow <= 0) return;

            try
            {
                // 计算实际写入的物理行数 (不超过元器件集合上限)
                int writeRows = Math.Min(compCount, components.Count);
                // 构建单列二维数组 (规则 7：杜绝逐格单元格 COM 互操作，一次性批量赋值)
                object[,] colMatrix = new object[writeRows, 1];

                // 遍历提取每个元件在当前槽位的数据
                for (int r = 0; r < writeRows; r++)
                {
                    // 获取当前行元器件模型
                    var pc = components[r];
                    string val = string.Empty;

                    // 从动态扩展值列表 ExtraValues 中提取对应槽位的值
                    if (pc.ExtraValues != null && colSlotIndex < pc.ExtraValues.Count)
                    {
                        val = pc.ExtraValues[colSlotIndex];
                    }

                    // 若槽位暂无直接提取值，根据目标列意图回退取值
                    if (string.IsNullOrEmpty(val))
                    {
                        // 目标为第 5 列时优先回退 Unit 字段
                        if (extraCol.TargetCol == 5) val = pc.Unit;
                        // 目标为第 9 列时优先回退 Remark 字段
                        else if (extraCol.TargetCol == 9) val = pc.Remark;
                    }
                    // 赋入单列二维矩阵
                    colMatrix[r, 0] = val ?? string.Empty;
                }

                // 计算目标列对应的标准 Excel 字母列标 (例如 5->E, 9->I, 15->O, 27->AA, 32->AF)
                string colLetter = GetExcelColumnLetter(extraCol.TargetCol);
                // 计算截止写入行号
                int endRow = startRow + writeRows - 1;

                // 数组一次性写入目标列指定范围 (规则 7 杜绝单元格逐格写入)
                sheet.Range[$"{colLetter}{startRow}:{colLetter}{endRow}"].Value2 = colMatrix;
            }
            catch (Exception ex)
            {
                // 记录异常日志，保障导入主流程不中断
                LogHelper.WriteLog($"[WriteExtraColumnIfConfigured] 槽位 {colSlotIndex + 1} 写入目标列 {extraCol.TargetCol} 异常: {ex.Message}");
            }
        }
    }
}
