using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：公式法调费 (对应 formula_adjust_fee.html)
    /// </summary>
    public static partial class ExcelServices
    {
        // 公式法调费窗口静态单例引用 (可空)
        private static FormulaAdjustFeeForm? _formulaAdjustFeeForm;

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“公式法调费”窗口 (非模态，可编辑 Excel)
        /// </summary>
        public static void ShowFormulaAdjustFeeDialog()
        {
            try
            {
                // 以非模态方式展示公式法调费窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _formulaAdjustFeeForm, () => new FormulaAdjustFeeForm());
            }
            catch (Exception ex)
            {
                // 全局捕获异常防止 Excel 崩溃闪退
                System.Windows.Forms.MessageBox.Show($"弹出公式法调费窗口失败: {ex.Message}", "错误提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 执行“公式法调费”逻辑: 解析公式表达式并精准更新写回 Excel 目标箱柜的费用行
        /// </summary>
        /// <param name="targetScope">调费作用域 (currentCabinet/currentCategory/allCabinets/selectedCabinet)</param>
        /// <param name="groupName">选中的公式组名称</param>
        /// <param name="items">前端编辑传递的公式明细项</param>
        public static void ApplyFormulaAdjustFeeToExcel(string targetScope, string groupName, System.Collections.Generic.List<Controllers.FormulaItemModel>? items = null)
        {
            try
            {
                // 获取当前运行的 Excel Application COM 接口实例 (安全调用)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return;

                // 获取当前激活的工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null) return;

                // 获取当前活动工作表
                dynamic activeSheet = activeWb.ActiveSheet;
                if (activeSheet == null) return;

                // 若前端未显式传递 items，则从控制器读取预置公式明细
                if (items == null || items.Count == 0)
                {
                    var controller = new Controllers.FormulaAdjustFeeController();
                    items = controller.GetFormulaDetails(groupName);
                }

                if (items == null || items.Count == 0) return;

                // 读取 4 种定义名称前缀配置项 (零堆分配)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

                // 构建当前工作表有效箱柜映射 (显式强类型复用 Tool 公共方法)
                List<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets = Tool.GetSheetValidCabinets(activeSheet);

                var targetCabinets = new System.Collections.Generic.List<KeyValuePair<int, Models.CabinetAnchorModel>>();

                // 自动补齐定义名称
                if (validCabinets.Count == 0)
                {
                    // 补齐定义名称
                    Tool.FixAndFillCabinetNamesForSheet(activeSheet);
                    // 重新扫描有效箱柜
                    validCabinets = Tool.GetSheetValidCabinets(activeSheet);
                }

                // 作用域筛选
                if (targetScope == "currentCabinet")
                {
                    // 智能匹配当前光标所属箱柜 (显式强类型声明)
                    KeyValuePair<int, Models.CabinetAnchorModel>? matched = Tool.GetActiveCabinet(app, validCabinets, fallbackSingle: true);
                    // 校验是否命中箱柜 (强类型 HasValue 判定)
                    if (matched.HasValue)
                    {
                        // 加入目标箱柜列表
                        targetCabinets.Add(matched.Value);
                    }
                }
                else
                {
                    // 包含所有箱柜
                    targetCabinets.AddRange(validCabinets);
                }

                if (targetCabinets.Count == 0) return;

                // 关闭刷新提升计算与写入性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                try
                {
                    int N = items.Count;
                    foreach (var cab in targetCabinets)
                    {
                        int k = cab.Key;
                        int cabDetRow = Convert.ToInt32(cab.Value.Det.Row);
                        int oldSubsumRow = Convert.ToInt32(cab.Value.Subsum.Row);
                        int oldTolsumRow = Convert.ToInt32(cab.Value.Tolsum.Row);
                        int compStartRow = cabDetRow + 2;

                        int oldM = oldTolsumRow - oldSubsumRow + 1;
                        int delta = N - oldM;

                        // 差额插入或删除行以对齐计费行数
                        if (delta > 0)
                        {
                            activeSheet.Rows[$"{oldSubsumRow}:{oldSubsumRow + delta - 1}"].Insert(-4121);
                        }
                        else if (delta < 0)
                        {
                            int deleteCount = -delta;
                            activeSheet.Rows[$"{oldSubsumRow}:{oldSubsumRow + deleteCount - 1}"].Delete(-4121);
                        }

                        int newSubsumRow = oldSubsumRow;
                        int newTolsumRow = newSubsumRow + N - 1;
                        int compEndRow = newSubsumRow - 1;

                        // 构建计费矩阵 (规则 7)
                        object[,] feeMatrix = Tool.BuildFeeMatrix(items, cabDetRow, newSubsumRow, compStartRow, compEndRow, 17);

                        // 批量覆盖写入 Excel 计费区域
                        dynamic feeRange = activeSheet.Range[$"A{newSubsumRow}:Q{newTolsumRow}"];
                        feeRange.Formula = feeMatrix;

                        // 更新工作表级别的定义名称锚点 (规则 6)
                        string sheetName = Convert.ToString(activeSheet.Name) ?? "";
                        // 覆盖更新当前箱柜小计行定义名称
                        Tool.SafeSetSheetName(activeSheet, sheetName, $"{subsumPrefix}{k}", newSubsumRow);
                        // 覆盖更新当前箱柜总计行定义名称
                        Tool.SafeSetSheetName(activeSheet, sheetName, $"{tolsumPrefix}{k}", newTolsumRow);

                        // 同步更新顶部汇总行公式
                        if (cab.Value.Sum != null)
                        {
                            int sumRow = Convert.ToInt32(cab.Value.Sum.Row);
                            activeSheet.Cells[sumRow, 7].Formula = $"=H{newTolsumRow}";
                            activeSheet.Cells[sumRow, 10].Formula = $"=K{newTolsumRow}";
                        }
                    }
                }
                finally
                {
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"执行公式法调费异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"执行公式法调费失败: {ex.Message}", "错误提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 公式法调费设为默认：将新的默认公式组计费行写回标准模板 CabinetTemplate.xlsx 中的【分类1】工作表
        /// 遵循规则：先删除模板中旧的计费行，采用汇总行对齐的方式，写入模板新的计费行，并更新定义名称与汇总行公式
        /// </summary>
        /// <param name="items">公式明细项集合</param>
        /// <param name="groupName">公式组名称</param>
        /// <param name="explicitApp">可选传入的 Excel Application COM 实例</param>
        /// <returns>操作是否成功</returns>
        public static bool UpdateCabinetTemplateDefaultFee(
            System.Collections.Generic.List<Controllers.FormulaItemModel>? items = null,
            string? groupName = null,
            dynamic? explicitApp = null)
        {
            // 声明模板工作簿句柄
            dynamic? templateWb = null;
            // 声明 Excel Application COM 句柄
            dynamic? app = null;

            try
            {
                // 1. 获取 Excel Application COM 接口实例 (安全调用)
                app = explicitApp ?? ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return false;

                // 2. 若入参 items 为空，从控制器读取当前默认公式组明细
                if (items == null || items.Count == 0)
                {
                    // 实例化公式控制器
                    var controller = new Controllers.FormulaAdjustFeeController();
                    // 提取指定组或默认组明细
                    items = !string.IsNullOrWhiteSpace(groupName)
                        ? controller.GetFormulaDetails(groupName)
                        : controller.GetFormulaDetails("多费用公式");
                }

                // 校验明细集合有效性
                if (items == null || items.Count == 0)
                {
                    LogHelper.WriteLog("更新模板计费行失败: 公式明细集合为空");
                    return false;
                }

                // 3. 获取 CabinetTemplate.xlsx 模板物理文件路径
                string templatePath = Controllers.ProjectController.EnsureCabinetTemplate(app);
                if (!System.IO.File.Exists(templatePath))
                {
                    LogHelper.WriteLog($"未找到模板物理文件: {templatePath}");
                    return false;
                }

                // 确保模板文件不是只读属性
                try
                {
                    var fileInfo = new System.IO.FileInfo(templatePath);
                    if (fileInfo.IsReadOnly) fileInfo.IsReadOnly = false;
                }
                catch { }

                // 4. 读取配置与定义名称前缀
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;
                string defaultTemplateSheet = ConfigManager.Instance.Current.Excel.DefaultTemplateSheet ?? "分类1";

                // 5. 以可写方式打开模板工作簿
                templateWb = app.Workbooks.Open(templatePath, UpdateLinks: 0, ReadOnly: false);
                if (templateWb == null) return false;

                // 获取目标模板工作表 (优先分类1，若无取第2个或第1个)
                dynamic? catSheet = null;
                try { catSheet = templateWb.Sheets[defaultTemplateSheet]; } catch { }
                if (catSheet == null)
                {
                    try { catSheet = templateWb.Sheets.Count >= 2 ? templateWb.Sheets[2] : templateWb.Sheets[1]; } catch { }
                }
                if (catSheet == null)
                {
                    templateWb.Close(false);
                    return false;
                }

                string sheetName = Convert.ToString(catSheet.Name) ?? defaultTemplateSheet;

                // 6. 探测模板工作表中箱柜 1 的标准行号分布 (复用 Tool 公共方法)
                var rowIndexes = Tool.FindStandardCategoryRowIndexes((object)catSheet, 1);
                int cabSumRow = rowIndexes.cabSumRow;
                int cabDetRow = rowIndexes.cabDetRow;
                int oldSubsumRow = rowIndexes.cabSubsumRow;
                int oldTolsumRow = rowIndexes.cabTolsumRow;

                // 若有任意一个标准行号无效 (<= 0)，调用现成方法修复定义名称并重新获取所有标准行
                if (cabSumRow <= 0 || cabDetRow <= 0 || oldSubsumRow <= 0 || oldTolsumRow <= 0)
                {
                    // 调用现成方法自动补齐与修复当前工作表定义名称
                    Tool.FixAndFillCabinetNamesForSheet(catSheet);
                    // 重新获取所有标准行分布
                    rowIndexes = Tool.FindStandardCategoryRowIndexes((object)catSheet, 1);
                    // 重新赋值汇总行号
                    cabSumRow = rowIndexes.cabSumRow;
                    // 重新赋值明细信息行号
                    cabDetRow = rowIndexes.cabDetRow;
                    // 重新赋值小计行号
                    oldSubsumRow = rowIndexes.cabSubsumRow;
                    // 重新赋值总计行号
                    oldTolsumRow = rowIndexes.cabTolsumRow;
                }

                int compStartRow = cabDetRow + 2;
                int newFeeRows = items.Count;

                // 7. 直接清空旧计费行的 B~Q 列数据 (A 列序号与定义名称保留不删)
                if (oldTolsumRow >= oldSubsumRow && oldSubsumRow > 0)
                {
                    // 清空旧计费区域 B 列至 Q 列数据
                    catSheet.Range[$"B{oldSubsumRow}:Q{oldTolsumRow}"].ClearContents();
                }

                // 8. 采用汇总行对齐方式计算新计费行物理区间: 以总计行向上对齐计算小计行
                int newSubsumRow = oldTolsumRow - newFeeRows + 1;
                // 新总计行保持与模板汇总行对齐
                int newTolsumRow = oldTolsumRow;
                // 计算元器件终止行
                int compEndRow = newSubsumRow - 1;

                // 9. 构建计费矩阵并批量一次性写入模板计费区域 (覆盖 A 列至 Q 列)
                object[,] feeMatrix = Tool.BuildFeeMatrix(items, cabDetRow, newSubsumRow, compStartRow, compEndRow, 17);
                // 覆盖写入新计费区域
                dynamic feeRange = catSheet.Range[$"A{newSubsumRow}:Q{newTolsumRow}"];
                feeRange.Formula = feeMatrix;

                // 10. 为元器件区域 (compStartRow 到 compEndRow) 重新刷入自适应公式矩阵 (保证 A~Q 列纯净)
                int compRowCount = compEndRow - compStartRow + 1;
                if (compRowCount > 0)
                {
                    // 生成自适应公式矩阵
                    object[,] compMatrix = Tool.BuildComponentRowsMatrix(compStartRow, compEndRow, cabDetRow, 17);
                    // 批量覆盖写回元器件区域
                    catSheet.Range[$"A{compStartRow}:Q{compEndRow}"].Formula = compMatrix;
                }

                // 11. 重新注册并更新模板工作表 4 个定义名称锚点
                Tool.SafeSetSheetName(catSheet, sheetName, $"{sumPrefix}1", cabSumRow);
                Tool.SafeSetSheetName(catSheet, sheetName, $"{detPrefix}1", cabDetRow);
                Tool.SafeSetSheetName(catSheet, sheetName, $"{subsumPrefix}1", newSubsumRow);
                Tool.SafeSetSheetName(catSheet, sheetName, $"{tolsumPrefix}1", newTolsumRow);

                // 12. 采用汇总行对齐的方式: 重新对齐绑定顶部汇总行 (cabSumRow, Row 7) 引用公式与超链接
                catSheet.Cells[cabSumRow, 7].Formula = $"=H{newTolsumRow}";
                catSheet.Cells[cabSumRow, 8].Formula = $"=F{cabSumRow}*G{cabSumRow}";
                catSheet.Cells[cabSumRow, 10].Formula = $"=K{newTolsumRow}";
                catSheet.Cells[cabSumRow, 11].Formula = $"=H{cabSumRow}-J{cabSumRow}";
                catSheet.Cells[cabSumRow, 12].Formula = $"=IF(H{cabSumRow}=0,0,K{cabSumRow}/H{cabSumRow})";

                // 双向超链接绑定
                try
                {
                    // 汇总行 B 列超链接指向明细 A 列
                    catSheet.Hyperlinks.Add(
                        Anchor: catSheet.Cells[cabSumRow, 2],
                        Address: "",
                        SubAddress: $"{sheetName}!A{cabDetRow}",
                        ScreenTip: "点击跳转至明细行"
                    );
                    // 明细行 B 列超链接指向汇总 A 列
                    catSheet.Hyperlinks.Add(
                        Anchor: catSheet.Cells[cabDetRow, 2],
                        Address: "",
                        SubAddress: $"{sheetName}!A{cabSumRow}",
                        ScreenTip: "点击返回汇总行"
                    );
                }
                catch { }

                // 13. 保存模板工作簿
                templateWb.Save();
                LogHelper.WriteLog($"成功将默认公式组计费行写回模板: {templatePath} (总计行对齐至 Row {newTolsumRow})");

                // 14. 若开发目录存在 CabinetTemplate.xlsx 副本，同步更新覆盖
                try
                {
                    string currentDevPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "CabinetTemplate.xlsx");
                    if (!string.Equals(templatePath, currentDevPath, StringComparison.OrdinalIgnoreCase) && System.IO.File.Exists(templatePath))
                    {
                        System.IO.File.Copy(templatePath, currentDevPath, true);
                    }
                }
                catch { }

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"更新模板默认计费行异常: {ex.Message}");
                return false;
            }
            finally
            {
                // 确保关闭模板工作簿句柄
                if (templateWb != null)
                {
                    try { templateWb.Close(false); } catch { }
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(templateWb); } catch { }
                }
            }
        }
    }
}
