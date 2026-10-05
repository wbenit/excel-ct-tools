using System;
using System.Collections.Generic;
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
        /// </summary>
        /// <param name="request">导入请求参数</param>
        /// <returns>执行结果对象</returns>
        public static SmartImportExecuteResult ExecuteSmartImportToNewCategory(SmartImportExecuteRequest request)
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

                // 3. 调用公共分类服务创建全新的纯净分类表 Sheet (自动克隆模板并挂接项目信息主表)
                var createCatReq = new CreateCategoryRequest { CategoryName = newCategoryName };
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

                // 临时关闭屏幕刷新与提示，保障大批量写入性能与稳定性
                // 临时关闭屏幕刷新与提示，保障大批量写入性能与稳定性
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;
                // 静默外部源链接更新询问，杜绝链接安全警告弹窗
                try { app.AskToUpdateLinks = false; } catch { }

                try
                {
                    // 5. 规则 8 事前校验：纠偏并整理新表的规则 6 架构
                    RepairAllDefinedNames();

                    // 收集或建立各箱柜行号映射
                    int totalCabinets = cabinets.Count;
                    int importedCompTotal = 0;
                    // 记录首柜基准全局编号
                    int baseCabinetK = 1;

                    for (int i = 0; i < totalCabinets; i++)
                    {
                        var cab = cabinets[i];
                        CabinetCreatedInfo cabInfo = null;
                        int cabinetK = 1;

                        if (i == 0)
                        {
                            // 第 1 柜：模板中自带第 1 柜结构，动态精准定位其物理行号 (杜绝错误硬编码偏位)
                            cabInfo = LocateFirstCabinetInfo(newSheet);
                            if (cabInfo != null && cabInfo.CabinetK > 0)
                            {
                                baseCabinetK = cabInfo.CabinetK;
                            }
                            cabinetK = baseCabinetK;
                        }
                        else
                        {
                            // 第 2~N 柜：依次在当前新工作表追加新箱柜明细区域 (编号平滑连续自增)
                            cabinetK = baseCabinetK + i;
                            cabInfo = CopyCabinetDetailFromTemplate(newSheet, 0, cabinetK, cab.CabinetName, app);
                        }

                        if (cabInfo == null)
                        {
                            // 记录创建或定位箱柜失败日志
                            LogHelper.WriteLog($"[SmartImport] 创建/定位箱柜 {cabinetK} 失败，跳过该柜。");
                            continue;
                        }

                        // 6. 回写顶部汇总表行数据 (B列柜名，E列箱柜台数)
                        int sumRow = cabInfo.SumRow;
                        if (sumRow > 0)
                        {
                            // 回写 B 列箱柜名称
                            newSheet.Cells[sumRow, 2].Value2 = cab.CabinetName;
                            // 回写 E 列箱柜数量(台数)
                            newSheet.Cells[sumRow, 5].Value2 = cab.CabinetQuantity;
                        }

                        // 7. 回写底部明细箱柜信息行 (B列柜名)
                        int detRow = cabInfo.DetRow;
                        if (detRow > 0)
                        {
                            // 回写明细信息行 B 列柜名
                            newSheet.Cells[detRow, 2].Value2 = cab.CabinetName;
                        }

                        // 8. 规则 6 校验与元器件行数自适应扩容
                        int compStartRow = detRow + 2;
                        int subsumRow = cabInfo.SubsumRow;
                        int tolsumRow = cabInfo.TolsumRow;
                        // 计算当前预留的元器件行容量
                        int defaultCap = (subsumRow - 1) - compStartRow + 1;
                        int actualCompCount = cab.Components != null ? cab.Components.Count : 0;

                        // 若导入元器件数量多于预留行，在小计行前批量插入物理行 (规则 6)
                        if (actualCompCount > defaultCap)
                        {
                            // 计算需扩容插入的物理行数
                            int insertRowsCount = actualCompCount - defaultCap;
                            // 插入起点即为小计行
                            int insertRowStart = subsumRow;

                            // 执行整行向下推移插入
                            newSheet.Rows[$"{insertRowStart}:{insertRowStart + insertRowsCount - 1}"].Insert(-4121);
                            // 更新推移后的小计与总计行号
                            subsumRow += insertRowsCount;
                            tolsumRow += insertRowsCount;
                            cabInfo.SubsumRow = subsumRow;
                            cabInfo.TolsumRow = tolsumRow;

                            // 同步校准推移后的定义名称 (规则 6)
                            var (sPre, dPre, subPre, tolPre) = CabinetPrefixConfig.Current;
                            Tool.SafeSetSheetName(newSheet, newCategoryName, $"{subPre}{cabinetK}", subsumRow);
                            Tool.SafeSetSheetName(newSheet, newCategoryName, $"{tolPre}{cabinetK}", tolsumRow);
                        }

                        int compEndRow = subsumRow - 1;

                        // 9. 规则 7：利用 Tool.BuildComponentRowsMatrix 一次性批量写入元器件二维矩阵
                        if (compEndRow >= compStartRow)
                        {
                            var compItemList = new List<ComponentItem>();
                            if (cab.Components != null)
                            {
                                // 组装元器件模型集合
                                foreach (var pc in cab.Components)
                                {
                                    compItemList.Add(new ComponentItem
                                    {
                                        Name = pc.ItemName,
                                        Specification = pc.ItemSpec,
                                        Manufacturer = pc.Brand,
                                        Quantity = pc.Quantity,
                                        // 重点：将外部单价写入表价/面价 (UnitPrice 映射到 Tool 中的 M 列面价)
                                        UnitPrice = pc.MarkedPrice,
                                        // 方案 B：携带并回填计量单位
                                        Unit = pc.Unit,
                                        // 方案 B：携带并回填备注文本
                                        Remark = pc.Remark
                                    });
                                }
                            }

                            // 生成 21 列标准公式与数据矩阵 (全量覆盖预留区，彻底清空旧模板示例元件)
                            object[,] compMatrix = Tool.BuildComponentRowsMatrix(compStartRow, compEndRow, detRow, 21, compItemList);
                            // 数组一次性批量写入 Excel 区域 (规则 7)
                            newSheet.Range[$"A{compStartRow}:U{compEndRow}"].Formula = compMatrix;
                            importedCompTotal += actualCompCount;

                            // 方案 B：根据用户配置将扩展列 1 批量安全写入导入表任意指定目标列
                            WriteExtraColumnIfConfigured(newSheet, compStartRow, actualCompCount, cab.Components, request.Config?.ColumnMapping?.ExtraCol1, 1);
                            // 方案 B：根据用户配置将扩展列 2 批量安全写入导入表任意指定目标列
                            WriteExtraColumnIfConfigured(newSheet, compStartRow, actualCompCount, cab.Components, request.Config?.ColumnMapping?.ExtraCol2, 2);
                        }

                        // 10. 检查计费区域 (SubsumRow 到 TolsumRow-1) 是否存在空行，确保公式紧凑闭环 (规则 6)
                        CleanEmptyRowsInBillingZone(newSheet, subsumRow, ref tolsumRow);
                    }

                    // 10.5 全局清洗跨工作簿公式引用 (防止残留 [CabinetTemplate.xlsx] 产生外部链接警告)
                    try
                    {
                        Tool.CleanRangeFormulas(newSheet.UsedRange);
                    }
                    catch (Exception exClean)
                    {
                        LogHelper.WriteLog($"[SmartImport] 清洗外部公式引用异常: {exClean.Message}");
                    }

                    // 11. 事后全局定义名称自愈校验 (规则 8)
                    RepairAllDefinedNames();

                    // 12. 激活新工作表并置顶展示
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
                    // 恢复全局屏幕刷新与系统事件
                    try { app.ScreenUpdating = true; } catch { }
                    try { app.DisplayAlerts = true; } catch { }
                    try { app.EnableEvents = true; } catch { }
                    try { app.AskToUpdateLinks = true; } catch { }
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
        /// 清理计费区域的空行 (规则 6: 计费区域不能有空行，若有必须删除)
        /// </summary>
        private static void CleanEmptyRowsInBillingZone(dynamic sheet, int subsumRow, ref int tolsumRow)
        {
            try
            {
                // 计费区域为 SubsumRow 到 TolsumRow - 1
                int checkRow = subsumRow + 1;
                while (checkRow < tolsumRow)
                {
                    // 检查 B 列与 C 列是否为空
                    string bVal = Convert.ToString(sheet.Cells[checkRow, 2].Value2)?.Trim() ?? "";
                    string cVal = Convert.ToString(sheet.Cells[checkRow, 3].Value2)?.Trim() ?? "";

                    if (string.IsNullOrEmpty(bVal) && string.IsNullOrEmpty(cVal))
                    {
                        // 发现空行，物理剔除 (规则 6)
                        sheet.Rows[checkRow].Delete(-4121);
                        tolsumRow--; // 总计行上浮 1 行
                    }
                    else
                    {
                        checkRow++;
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
        /// </summary>
        /// <param name="sheet">目标工作表句柄</param>
        /// <param name="startRow">元器件起始物理行号</param>
        /// <param name="compCount">实际有效元器件数量</param>
        /// <param name="components">已解析的元器件明细集合</param>
        /// <param name="extraCol">扩展列映射配置对象</param>
        /// <param name="colSlotIndex">扩展列槽位 (1: ExtraCol1, 2: ExtraCol2)</param>
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
                    var pc = components[r];
                    // 依据当前槽位索引获取提取的文本
                    string val = (colSlotIndex == 1) ? pc.ExtraValue1 : pc.ExtraValue2;
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

                // 计算目标列对应的标准 Excel 字母列标 (例如 5->E, 9->I, 15->O)
                string colLetter = GetExcelColumnLetter(extraCol.TargetCol);
                // 计算截止写入行号
                int endRow = startRow + writeRows - 1;

                // 数组一次性写入目标列指定范围 (规则 7)
                sheet.Range[$"{colLetter}{startRow}:{colLetter}{endRow}"].Value2 = colMatrix;
            }
            catch (Exception ex)
            {
                // 记录异常日志，保障导入主流程不中断
                LogHelper.WriteLog($"[WriteExtraColumnIfConfigured] 槽位 {colSlotIndex} 写入目标列 {extraCol.TargetCol} 异常: {ex.Message}");
            }
        }
    }
}
