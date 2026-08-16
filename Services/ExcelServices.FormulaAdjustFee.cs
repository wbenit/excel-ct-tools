using System;
using System.Collections.Generic;
using ExcelDna.Integration;
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
                // 获取当前运行的 Excel Application COM 接口实例
                dynamic app = ExcelDnaUtil.Application;
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

                // 读取 4 种定义名称前缀配置项
                string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";
                string subsumPrefix = ConfigManager.Instance.Current.Excel.SubsumNamePrefix ?? "Cab_Subsum_";
                string tolsumPrefix = ConfigManager.Instance.Current.Excel.TolsumNamePrefix ?? "Cab_Tolsum_";

                // 收集工作簿级别和工作表级别的所有定义名称
                var allNames = new List<dynamic>();
                dynamic parentWb = null;
                try { parentWb = activeSheet.Parent; } catch { }

                if (parentWb != null && parentWb.Names != null)
                {
                    try { foreach (dynamic n in parentWb.Names) allNames.Add(n); } catch { }
                }
                if (activeSheet != null && activeSheet.Names != null)
                {
                    try { foreach (dynamic n in activeSheet.Names) allNames.Add(n); } catch { }
                }

                // 扫描定义名称并构建箱柜锚点字典
                var validCabinets = Tool.BuildCabinetMap(
                    allNames,
                    Convert.ToString(activeSheet.Name) ?? "",
                    sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);

                dynamic activeCell = app.ActiveCell;
                int activeRow = activeCell != null ? Convert.ToInt32(activeCell.Row) : 0;

                var targetCabinets = new System.Collections.Generic.List<KeyValuePair<int, Models.CabinetAnchorModel>>();

                // 自动补齐定义名称
                if (validCabinets.Count == 0)
                {
                    Tool.FixAndFillCabinetNamesForSheet(activeSheet);
                    allNames.Clear();
                    if (parentWb != null && parentWb.Names != null)
                    {
                        try { foreach (dynamic n in parentWb.Names) allNames.Add(n); } catch { }
                    }
                    if (activeSheet != null && activeSheet.Names != null)
                    {
                        try { foreach (dynamic n in activeSheet.Names) allNames.Add(n); } catch { }
                    }
                    validCabinets = Tool.BuildCabinetMap(
                        allNames,
                        Convert.ToString(activeSheet.Name) ?? "",
                        sumPrefix, detPrefix, subsumPrefix, tolsumPrefix);
                }

                // 作用域筛选
                if (targetScope == "currentCabinet")
                {
                    KeyValuePair<int, Models.CabinetAnchorModel> matched = default;
                    foreach (var c in validCabinets)
                    {
                        int sumRow = Convert.ToInt32(c.Value.Sum.Row);
                        int detRow = Convert.ToInt32(c.Value.Det.Row);
                        int tolsumRow = Convert.ToInt32(c.Value.Tolsum.Row);

                        if (activeRow == sumRow || (activeRow >= detRow - 3 && activeRow <= tolsumRow))
                        {
                            matched = c;
                            break;
                        }
                    }

                    if (matched.Key != 0)
                    {
                        targetCabinets.Add(matched);
                    }
                    else if (validCabinets.Count > 0)
                    {
                        targetCabinets.Add(validCabinets[0]);
                    }
                }
                else
                {
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

                        // 更新定义名称锚点
                        dynamic targetWb = activeSheet.Parent;
                        string sheetName = Convert.ToString(activeSheet.Name) ?? "";
                        targetWb.Names.Add($"{subsumPrefix}{k}", $"='{sheetName}'!$A${newSubsumRow}");
                        targetWb.Names.Add($"{tolsumPrefix}{k}", $"='{sheetName}'!$A${newTolsumRow}");

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
    }
}
