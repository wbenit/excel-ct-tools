using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Controllers;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：新建分类工作表、分类表通用初始化与定义名称绑定
    /// </summary>
    public static partial class ExcelServices
    {
        // 缓存单例模态/非模态窗体引用，避免多开
        private static Forms.CategoryForm? _categoryFormInstance;

        /// <summary>
        /// 供 Ribbon 菜单调用的新建分类入口：弹出基于 WebView2 + Vue 3 的新建分类窗体
        /// </summary>
        public static void ShowCategoryDialog()
        {
            try
            {
                // 获取当前正在运行的 Excel Application 对象
                dynamic app = ExcelDnaUtil.Application;
                if (app == null || app.ActiveWorkbook == null)
                {
                    // 若无工作簿打开则提示用户
                    System.Windows.Forms.MessageBox.Show(
                        "请先打开或新建一个报价项目工作簿！",
                        "系统提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // 若窗口实例已存在且未被释放，则直接置顶激活
                if (_categoryFormInstance != null && !_categoryFormInstance.IsDisposed)
                {
                    // 激活并置前窗口
                    _categoryFormInstance.BringToFront();
                    _categoryFormInstance.Activate();
                    return;
                }

                // 实例化新建分类宿主窗体
                _categoryFormInstance = new Forms.CategoryForm();

                // 获取 Excel 主窗口 Win32 句柄以模态附着弹出，防止下沉或异常
                IntPtr excelHwnd = ExcelDnaUtil.WindowHandle;
                _categoryFormInstance.ShowDialog(new ExcelWin32Window(excelHwnd));
            }
            catch (Exception ex)
            {
                // 全局捕获异常并记录日志
                LogHelper.WriteLog($"弹出新建分类窗口异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show(
                    $"弹出新建分类窗口失败: {ex.Message}",
                    "系统提示",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 扫描当前激活工作簿，探测已有分类列表并自动建议下一个可用分类名称
        /// </summary>
        /// <returns>分类建议与已有列表对象</returns>
        public static CategorySuggestInfo GetSuggestedCategoryInfo()
        {
            var info = new CategorySuggestInfo();

            try
            {
                // 获取 Excel 应用实例
                dynamic app = ExcelDnaUtil.Application;
                if (app == null) return info;

                // 获取当前活动工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null) return info;

                int maxCatIndex = 0;
                var existingNames = new List<string>();

                // 遍历当前工作簿中所有的工作表
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    string sName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;
                    // 排除系统辅助表 --硬编码--
                    if (string.Equals(sName, "项目信息", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // 记录已有分类表名
                    existingNames.Add(sName);

                    // 匹配 "分类N" 形式的序号
                    if (sName.StartsWith("分类", StringComparison.OrdinalIgnoreCase))
                    {
                        string suffix = sName.Substring("分类".Length);
                        if (int.TryParse(suffix, out int idx) && idx > maxCatIndex)
                        {
                            maxCatIndex = idx;
                        }
                    }
                }

                // 设置已有分类表列表
                info.ExistingCategories = existingNames;
                // 推荐下一个分类名称 (如已有 分类1，则推荐 分类2)
                info.SuggestedName = $"分类{maxCatIndex + 1}";
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"探测分类建议名称异常: {ex.Message}");
            }

            return info;
        }

        /// <summary>
        /// 核心方法：执行新建分类工作表（克隆模板 -> 清洗公式 -> 调用通用分类初始化 -> 联动项目信息表）
        /// </summary>
        /// <param name="request">新建分类参数请求</param>
        /// <returns>操作结果对象</returns>
        public static CategoryOperationResult CreateNewCategory(CreateCategoryRequest request)
        {
            // 校验请求对象
            if (request == null || string.IsNullOrWhiteSpace(request.CategoryName))
            {
                return new CategoryOperationResult { Success = false, Message = "分类名称不能为空" };
            }

            string newCategoryName = request.CategoryName.Trim();

            try
            {
                // 获取 Excel COM Application 接口
                dynamic app = ExcelDnaUtil.Application;
                if (app == null) return new CategoryOperationResult { Success = false, Message = "无法连接 Excel 实例" };

                // 获取活动工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null) return new CategoryOperationResult { Success = false, Message = "当前无活动工作簿" };

                // 检查是否已存在同名工作表
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    if (string.Equals(Convert.ToString(ws.Name)?.Trim(), newCategoryName, StringComparison.OrdinalIgnoreCase))
                    {
                        return new CategoryOperationResult
                        {
                            Success = false,
                            Message = $"工作簿中已存在名为【{newCategoryName}】的工作表，请更换名称！"
                        };
                    }
                }

                // 临时关闭屏幕刷新与提示警告以提高初始化速度
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                try
                {
                    // 1. 从 CabinetTemplate.xlsx 标准模板克隆纯净的【分类1】工作表 (100% 完整保留汇总表头、合计行、大写、说明及明细模板)
                    string templatePath = ProjectController.EnsureCabinetTemplate(app);
                    dynamic templateWb = app.Workbooks.Open(templatePath, ReadOnly: true);
                    dynamic newSheet = null;
                    try
                    {
                        // 获取模板工作簿中的标准分类模板表
                        dynamic templateSheet = templateWb.Sheets["分类1"];
                        // 复制克隆至目标活动工作簿末尾
                        templateSheet.Copy(After: activeWb.Sheets[activeWb.Sheets.Count]);
                        // 复制后的新工作表即为最后一个工作表
                        newSheet = activeWb.Sheets[activeWb.Sheets.Count];
                    }
                    finally
                    {
                        // 立即关闭模板工作簿句柄，防止定义名称生命周期异常
                        templateWb.Close(false);
                    }

                    if (newSheet == null)
                    {
                        return new CategoryOperationResult { Success = false, Message = "复制创建分类工作表失败" };
                    }

                    // 2. 重命名新工作表为用户指定的分类名称
                    newSheet.Name = newCategoryName;

                    // 3. 清洗跨工作簿公式引用 (将 [CabinetTemplate.xlsx] 转换为当前工作簿本地公式，保护 A 列定义名称)
                    Tool.CleanRangeFormulas((Microsoft.Office.Interop.Excel.Range)newSheet.UsedRange);

                    // 4. 动态计算下一个全局唯一的箱柜序号 K
                    Microsoft.Office.Interop.Excel.Workbook excelWb = (Microsoft.Office.Interop.Excel.Workbook)activeWb;
                    Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)newSheet;
                    int cabinetK = GetNextCabinetIndex(excelWb, excelSheet);

                    // 5. 调用公共通用分类初始化方法 (新建项目与新建分类 100% 共用此逻辑)
                    string initCabName = string.IsNullOrWhiteSpace(request.InitialCabinetName) ? "箱柜1" : request.InitialCabinetName.Trim();
                    InitializeCategorySheet(activeWb, newSheet, cabinetK, initCabName, request.FormulaGroupId);

                    // 6. 联动更新【项目信息】工作表中的【分类汇总】区域
                    UpdateProjectInfoCategorySummary(activeWb, newCategoryName);

                    // 返回操作成功结果
                    return new CategoryOperationResult
                    {
                        Success = true,
                        Message = $"分类【{newCategoryName}】创建成功！",
                        CategoryName = newCategoryName
                    };
                }
                finally
                {
                    // 恢复 Excel 屏幕刷新、告警与事件调度
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }
            }
            catch (Exception ex)
            {
                // 记录全局执行异常
                LogHelper.WriteLog($"新建分类异常: {ex.Message}");
                return new CategoryOperationResult
                {
                    Success = false,
                    Message = $"新建分类失败: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// 公共通用方法：初始化分类工作表结构（计费矩阵写入、元器件动态序号、4个定义名称锚点绑定、公式联动与双向超链接）
        /// 遵循规则 6（行号结构与 4 个定义名称）与规则 7（内存二维数组批量操作）
        /// </summary>
        /// <param name="targetWb">目标工作簿 COM 实例</param>
        /// <param name="catSheet">目标分类工作表 COM 实例</param>
        /// <param name="cabinetIndex">分配的全局箱柜编号 K</param>
        /// <param name="cabinetName">初始箱柜名称 (默认: 箱柜1)</param>
        /// <param name="formulaGroupId">选中的调费公式组 ID (可选，为空时采用默认组)</param>
        public static void InitializeCategorySheet(
            dynamic targetWb,
            dynamic catSheet,
            int cabinetIndex,
            string cabinetName = "箱柜1",
            string formulaGroupId = "")
        {
            if (targetWb == null || catSheet == null) return;

            try
            {
                string sheetName = Convert.ToString(catSheet.Name) ?? "";

                // 1. 动态智能探测当前分类表的基准行号分布 (自动适配模板实际布局，彻底杜绝硬编码行号 +1 偏移)
                var rowIndexes = Tool.FindStandardCategoryRowIndexes((object)catSheet);
                int cabSumRow = rowIndexes.cabSumRow;
                int cabDetRow = rowIndexes.cabDetRow;
                int detectedSubsumRow = rowIndexes.cabSubsumRow;
                int cabTolsumRow = rowIndexes.cabTolsumRow;

                var cfg = ConfigManager.Instance.Current.Excel;
                string sumPrefix = cfg.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = cfg.DetNamePrefix ?? "Cab_Det_";
                string subsumPrefix = cfg.SubsumNamePrefix ?? "Cab_Subsum_";
                string tolsumPrefix = cfg.TolsumNamePrefix ?? "Cab_Tolsum_";

                // 2. 加载调费公式组明细项并计算计费矩阵 (规则 6 & 7)
                int subsumRow = detectedSubsumRow; // 动态探测得到的小计行物理行号
                int compStartRow = cabDetRow + 2; // 元器件起始行物理行号
                int compEndRow = cabTolsumRow - 1; // 元器件终止行物理行号

                try
                {
                    // 实例化公式调费控制器
                    var feeController = new Controllers.FormulaAdjustFeeController();
                    FormulaGroupModel? targetGroup = null;

                    // 若传入了指定的公式组 ID 则按 ID 查找
                    if (!string.IsNullOrWhiteSpace(formulaGroupId))
                    {
                        targetGroup = feeController.GetFormulaGroups()?.FirstOrDefault(g => g.Id == formulaGroupId);
                    }
                    // 否则回退至默认公式组
                    targetGroup = targetGroup ?? feeController.GetDefaultGroup();
                    var items = targetGroup?.Details ?? new List<FormulaItemModel>();

                    // 当明细项集合有效非空时，执行向上对齐覆盖
                    if (items.Count > 0)
                    {
                        int N = items.Count;
                        // 向上对齐总计行计算小计行物理行号: 小计行 = 总计行 - N + 1
                        subsumRow = cabTolsumRow - N + 1;
                        // 元器件区域终止物理行号: 元器件终止行 = 小计行 - 1
                        compEndRow = subsumRow - 1;

                        // 调用公共服务方法构建 N 行 17 列的计费二维矩阵 (覆盖 A 列至 Q 列)
                        object[,] feeMatrix = Tool.BuildFeeMatrix(items, cabDetRow, subsumRow, compStartRow, compEndRow, 17);

                        // 将构建完成的计费二维矩阵一次性批量覆盖写入 Excel 计费区域 (规则 7)
                        dynamic feeRange = catSheet.Range[$"A{subsumRow}:Q{cabTolsumRow}"];
                        feeRange.Formula = feeMatrix;

                        // 为元器件区域 (compStartRow 到 compEndRow) 批量写入 A 列动态序号公式
                        int compRowCount = compEndRow - compStartRow + 1;
                        if (compRowCount > 0)
                        {
                            object[,] compNoMatrix = new object[compRowCount, 1];
                            for (int r = 0; r < compRowCount; r++)
                            {
                                compNoMatrix[r, 0] = $"=ROW()-ROW(A${cabDetRow + 1})";
                            }
                            catSheet.Range[$"A{compStartRow}:A{compEndRow}"].Formula = compNoMatrix;
                        }
                    }

                    // 3. 注册绑定 4 个强类型定义名称锚点 (规则 6)
                    targetWb.Names.Add($"{sumPrefix}{cabinetIndex}", $"='{sheetName}'!$A${cabSumRow}");
                    targetWb.Names.Add($"{detPrefix}{cabinetIndex}", $"='{sheetName}'!$A${cabDetRow}");
                    targetWb.Names.Add($"{subsumPrefix}{cabinetIndex}", $"='{sheetName}'!$A${subsumRow}");
                    targetWb.Names.Add($"{tolsumPrefix}{cabinetIndex}", $"='{sheetName}'!$A${cabTolsumRow}");
                }
                catch (Exception exNames)
                {
                    LogHelper.WriteLog($"绑定分类表定义名称与写入计费矩阵异常: {exNames.Message}");
                }

                // 4. 顶部汇总行 (cabSumRow) 公式与超链接联动
                try
                {
                    string safeCabName = string.IsNullOrWhiteSpace(cabinetName) ? "箱柜1" : cabinetName.Trim();

                    // A 列超链接跳转至明细信息行定义名称 (Cab_Det_K)
                    catSheet.Hyperlinks.Add(
                        Anchor: catSheet.Range[$"A{cabSumRow}"],
                        Address: "",
                        SubAddress: $"'{sheetName}'!{detPrefix}{cabinetIndex}",
                        TextToDisplay: Convert.ToString(cabinetIndex)
                    );

                    // 填入箱柜名称
                    catSheet.Cells[cabSumRow, 2].Value = safeCabName;
                    // G 列单价公式指向明细总计行的销售总价 (H 列)
                    catSheet.Cells[cabSumRow, 7].Formula = $"=H{cabTolsumRow}";
                    // H 列总价公式 = 数量(F列) * 单价(G列)
                    catSheet.Cells[cabSumRow, 8].Formula = $"=F{cabSumRow}*G{cabSumRow}";
                    // J 列成本总价公式指向明细总计行的成本总价 (K 列)
                    catSheet.Cells[cabSumRow, 10].Formula = $"=K{cabTolsumRow}";
                    // K 列毛利公式 = 总价 - 成本总价
                    catSheet.Cells[cabSumRow, 11].Formula = $"=H{cabSumRow}-J{cabSumRow}";
                    // L 列毛利率公式
                    catSheet.Cells[cabSumRow, 12].Formula = $"=IF(H{cabSumRow}=0,0,K{cabSumRow}/H{cabSumRow})";
                }
                catch { }

                // 5. 底部明细信息行 (cabDetRow) 联动与超链接
                try
                {
                    string safeCabName = string.IsNullOrWhiteSpace(cabinetName) ? "箱柜1" : cabinetName.Trim();

                    // A 列超链接跳转回顶部汇总行定义名称 (Cab_Sum_K)
                    catSheet.Hyperlinks.Add(
                        Anchor: catSheet.Range[$"A{cabDetRow}"],
                        Address: "",
                        SubAddress: $"'{sheetName}'!{sumPrefix}{cabinetIndex}",
                        TextToDisplay: "柜号:"
                    );

                    // 填入箱柜名称
                    catSheet.Cells[cabDetRow, 2].Value = safeCabName;
                }
                catch { }

                // 6. 激活当前分类工作表为当前主视口
                try
                {
                    catSheet.Activate();
                    catSheet.Range["A1"].Select();
                }
                catch { }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"初始化分类工作表异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 公共通用方法：在【项目信息】工作表的【分类汇总】区域中追加或更新指定分类的联动汇总行
        /// </summary>
        /// <param name="targetWb">目标工作簿 COM 实例</param>
        /// <param name="categorySheetName">分类工作表名称</param>
        public static void UpdateProjectInfoCategorySummary(dynamic targetWb, string categorySheetName)
        {
            if (targetWb == null || string.IsNullOrWhiteSpace(categorySheetName)) return;

            try
            {
                dynamic infoSheet = null;
                try { infoSheet = targetWb.Sheets["项目信息"]; } catch { }
                if (infoSheet == null) return;

                // 寻找【分类汇总】区域的下一个可用行 (从 Row 29 开始向下扫描)
                int targetInfoRow = 29;
                while (true)
                {
                    string cellB = Convert.ToString(infoSheet.Cells[targetInfoRow, 2].Value)?.Trim() ?? "";
                    // 找到空白行或同名分类行
                    if (string.IsNullOrEmpty(cellB) || string.Equals(cellB, categorySheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }
                    targetInfoRow++;
                }

                // 写入分类序号 (相对序号)
                infoSheet.Cells[targetInfoRow, 1].Value = targetInfoRow - 28;

                // 写入分类名称并绑定工作表 A1 跳转超链接
                infoSheet.Hyperlinks.Add(
                    Anchor: infoSheet.Range[$"B{targetInfoRow}"],
                    Address: "",
                    SubAddress: $"'{categorySheetName}'!A1",
                    TextToDisplay: categorySheetName
                );

                // 写入箱柜数量统计公式: 统计该分类下箱柜汇总行 B 列非空数量
                infoSheet.Cells[targetInfoRow, 3].Formula = $"=COUNTA('{categorySheetName}'!B6:B20)";

                // 写入总价统计公式: 汇总该分类下所有箱柜的总价 (H 列)
                infoSheet.Cells[targetInfoRow, 4].Formula = $"=SUM('{categorySheetName}'!H6:H20)";
            }
            catch (Exception ex)
            {
                // 记录更新异常
                LogHelper.WriteLog($"更新项目信息表分类汇总异常: {ex.Message}");
            }
        }
    }
}
