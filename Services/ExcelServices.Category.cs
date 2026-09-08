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

        // 剪贴板中暂存的被复制源分类工作表名称
        private static string _copiedCategorySheetName = string.Empty;

        /// <summary>
        /// 供 Ribbon 菜单调用的分类弹窗入口：支持新建、编辑与插入复制分类模式 (非模态，可交互编辑 Excel)
        /// </summary>
        /// <param name="mode">弹窗操作模式: create(新建) / edit(编辑) / insertCopied(插入复制)</param>
        public static void ShowCategoryDialog(string mode = "create")
        {
            try
            {
                // 获取当前正在运行的 Excel Application 对象 (安全调用)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                // 校验工作簿是否打开
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

                // 如果是插入复制模式，必须校验是否存在已复制的源表
                if (mode == "insertCopied")
                {
                    // 校验是否已执行过复制
                    if (string.IsNullOrWhiteSpace(_copiedCategorySheetName))
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "请先在需要复制的分类工作表中点击【复制分类】！",
                            "系统提示",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                        return;
                    }
                }

                // 以统一非模态方式展示分类窗口，传递当前指定的业务模式
                ShowModelessForm(ref _categoryFormInstance, () => new Forms.CategoryForm(mode));
            }
            catch (Exception ex)
            {
                // 全局捕获异常并记录日志
                LogHelper.WriteLog($"弹出分类窗口异常: {ex.Message}");
                // 弹出异常提示对话框
                System.Windows.Forms.MessageBox.Show(
                    $"弹出分类窗口失败: {ex.Message}",
                    "系统提示",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 供 Ribbon 菜单直接调用的“编辑分类”入口
        /// </summary>
        public static void ShowEditCategoryDialog()
        {
            // 调度通用弹窗并指定为 edit 模式
            ShowCategoryDialog("edit");
        }

        /// <summary>
        /// 供 Ribbon 菜单直接调用的“插入复制的分类”入口
        /// </summary>
        public static void ShowInsertCopiedCategoryDialog()
        {
            // 调度通用弹窗并指定为 insertCopied 模式
            ShowCategoryDialog("insertCopied");
        }

        /// <summary>
        /// 供 Ribbon 菜单直接调用的“复制分类”功能
        /// </summary>
        public static void CopyCurrentCategory()
        {
            try
            {
                // 获取当前正在运行的 Excel Application 对象 (安全调用)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                // 校验工作簿与活动工作表
                if (app == null || app.ActiveWorkbook == null || app.ActiveSheet == null)
                {
                    System.Windows.Forms.MessageBox.Show("请先打开一个项目工作簿！", "系统提示");
                    return;
                }

                // 读取当前活动工作表名称
                string sheetName = Convert.ToString(app.ActiveSheet.Name)?.Trim() ?? string.Empty;

                // 排除系统辅助工作表 (如项目信息) --硬编码--
                if (string.Equals(sheetName, "项目信息", StringComparison.OrdinalIgnoreCase))
                {
                    System.Windows.Forms.MessageBox.Show("【项目信息】是系统主表中枢，不能作为分类复制！", "系统提示");
                    return;
                }

                // 暂存被复制的分类名称
                _copiedCategorySheetName = sheetName;

                // 提示用户复制成功
                System.Windows.Forms.MessageBox.Show(
                    $"分类【{sheetName}】已成功复制！\n请切换或直接点击【插入复制的分类】添加到项目中。",
                    "复制分类成功",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"复制分类异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"复制分类失败: {ex.Message}", "系统提示");
            }
        }

        /// <summary>
        /// 供 Ribbon 菜单直接调用的“删除分类”功能 (含安全校验与项目信息联动)
        /// </summary>
        public static void DeleteCurrentCategory()
        {
            try
            {
                // 获取当前正在运行的 Excel Application 对象 (安全调用)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                // 校验工作簿与活动工作表
                if (app == null || app.ActiveWorkbook == null || app.ActiveSheet == null)
                {
                    System.Windows.Forms.MessageBox.Show("请先打开一个项目工作簿！", "系统提示");
                    return;
                }

                dynamic activeWb = app.ActiveWorkbook;
                dynamic activeSheet = app.ActiveSheet;
                string sheetName = Convert.ToString(activeSheet.Name)?.Trim() ?? string.Empty;

                // 校验是否为受保护的系统工作表 --硬编码--
                if (string.Equals(sheetName, "项目信息", StringComparison.OrdinalIgnoreCase))
                {
                    System.Windows.Forms.MessageBox.Show("【项目信息】是系统核心工作表，严禁删除！", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // 统计工作簿内非项目信息的分类表总数量
                int catCount = 0;
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    string name = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                    if (!string.Equals(name, "项目信息", StringComparison.OrdinalIgnoreCase))
                    {
                        catCount++;
                    }
                }

                // 保证项目中至少保留一个分类工作表
                if (catCount <= 1)
                {
                    System.Windows.Forms.MessageBox.Show("项目中至少必须保留一个分类工作表，无法继续删除！", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // 弹出二次确认防呆对话框
                var confirm = System.Windows.Forms.MessageBox.Show(
                    $"确定要彻底删除分类【{sheetName}】及其包含的所有箱柜与明细数据吗？\n此操作将同步移除【项目信息】中的对应汇总，且不可撤销！",
                    "确认删除分类",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question);

                if (confirm != System.Windows.Forms.DialogResult.Yes)
                {
                    return;
                }

                // 执行删除分类核心业务
                var result = DeleteCategory(sheetName, app);
                if (result.Success)
                {
                    System.Windows.Forms.MessageBox.Show(result.Message, "删除成功", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show(result.Message, "删除失败", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"删除分类外层调度异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"删除分类异常: {ex.Message}", "系统提示");
            }
        }

        /// <summary>
        /// 扫描当前激活工作簿，探测已有分类列表并自动建议下一个可用分类名称
        /// </summary>
        /// <param name="mode">当前操作模式: create / edit / insertCopied</param>
        /// <returns>分类建议与已有列表对象</returns>
        public static CategorySuggestInfo GetSuggestedCategoryInfo(string mode = "create")
        {
            var info = new CategorySuggestInfo
            {
                Mode = mode,
                CopiedCategoryName = _copiedCategorySheetName
            };

            try
            {
                // 获取 Excel 应用实例 (安全调用)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return info;

                // 获取当前活动工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null) return info;

                // 尝试执行一次工作簿分类汇总超链接自愈规范化，确保历史错乱的超链接自动对齐 ExWinner 原生样式
                NormalizeCategorySummaryLinks(activeWb);

                // 获取当前活动工作表名称
                string activeSheetName = Convert.ToString(app.ActiveSheet?.Name)?.Trim() ?? string.Empty;
                if (!string.Equals(activeSheetName, "项目信息", StringComparison.OrdinalIgnoreCase))
                {
                    // 若当前处于分类表，则记录为当前激活分类
                    info.ActiveCategoryName = activeSheetName;
                }

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

                // 根据不同的模式设定初始推荐名称
                if (mode == "edit" && !string.IsNullOrWhiteSpace(info.ActiveCategoryName))
                {
                    // 编辑模式默认使用当前表名
                    info.SuggestedName = info.ActiveCategoryName;
                }
                else if (mode == "insertCopied" && !string.IsNullOrWhiteSpace(_copiedCategorySheetName))
                {
                    // 插入复制模式推荐 "原名称_副本" 或下一个序号
                    string candidate = $"{_copiedCategorySheetName}_副本";
                    if (existingNames.Contains(candidate, StringComparer.OrdinalIgnoreCase))
                    {
                        candidate = $"分类{maxCatIndex + 1}";
                    }
                    info.SuggestedName = candidate;
                }
                else
                {
                    // 默认新建模式：推荐下一个分类名称 (如已有 分类1，则推荐 分类2)
                    info.SuggestedName = $"分类{maxCatIndex + 1}";
                }
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
        /// <param name="explicitApp">可选显式传入的 Excel COM Application 实例</param>
        /// <returns>操作结果对象</returns>
        public static CategoryOperationResult CreateNewCategory(CreateCategoryRequest request, dynamic? explicitApp = null)
        {
            // 校验请求对象
            if (request == null || string.IsNullOrWhiteSpace(request.CategoryName))
            {
                return new CategoryOperationResult { Success = false, Message = "分类名称不能为空" };
            }

            string newCategoryName = request.CategoryName.Trim();

            try
            {
                // 获取 Excel COM Application 接口 (优先支持外部传入，回退 ExcelDnaSafeAccessor)
                dynamic? app = explicitApp ?? ExcelDnaSafeAccessor.GetApplication();
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
                    Tool.CleanRangeFormulas(newSheet.UsedRange);

                    // 4. 动态计算下一个全局唯一的箱柜序号 K
                    int cabinetK = GetNextCabinetIndex(activeWb, newSheet);

                    // 5. 联动更新【项目信息】工作表中的【分类汇总】区域 (优先注册以获取正确的分类序号物理行号)
                    UpdateProjectInfoCategorySummary(activeWb, newCategoryName);

                    // 6. 调用公共通用分类初始化方法 (新建项目与新建分类 100% 共用此逻辑)
                    string initCabName = string.IsNullOrWhiteSpace(request.InitialCabinetName) ? "箱柜1" : request.InitialCabinetName.Trim();
                    InitializeCategorySheet(activeWb, newSheet, cabinetK, initCabName, request.FormulaGroupId);

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

                // 读取箱柜定义名称前缀值对象 (零堆分配)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

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

                        // 为元器件区域 (compStartRow 到 compEndRow) 批量写入 A~Q 列自适应公式矩阵 (规则 6 & 规则 7)
                        int compRowCount = compEndRow - compStartRow + 1;
                        if (compRowCount > 0)
                        {
                            // 生成 17 列包含 F/G/H/J/K/L/N/Q 自适应判断公式的二维数据矩阵
                            object[,] compMatrix = Tool.BuildComponentRowsMatrix(compStartRow, compEndRow, cabDetRow, 17);
                            // 一次性批量覆盖写入元器件完整区域 (覆盖 A 列至 Q 列)
                            catSheet.Range[$"A{compStartRow}:Q{compEndRow}"].Formula = compMatrix;
                        }
                    }

                    // 3. 注册绑定 4 个工作表级别定义名称锚点 (规则 6)
                    // 绑定箱柜汇总行工作表级定义名称
                    Tool.SafeSetSheetName(catSheet, sheetName, $"{sumPrefix}{cabinetIndex}", cabSumRow);
                    // 绑定箱柜信息行工作表级定义名称
                    Tool.SafeSetSheetName(catSheet, sheetName, $"{detPrefix}{cabinetIndex}", cabDetRow);
                    // 绑定箱柜小计行工作表级定义名称
                    Tool.SafeSetSheetName(catSheet, sheetName, $"{subsumPrefix}{cabinetIndex}", subsumRow);
                    // 绑定箱柜总计行工作表级定义名称
                    Tool.SafeSetSheetName(catSheet, sheetName, $"{tolsumPrefix}{cabinetIndex}", cabTolsumRow);
                }
                catch (Exception exNames)
                {
                    LogHelper.WriteLog($"绑定分类表定义名称与写入计费矩阵异常: {exNames.Message}");
                }

                // 4. 顶部汇总行 (cabSumRow) 公式与超链接联动
                try
                {
                    // 获取安全的箱柜名称
                    string safeCabName = string.IsNullOrWhiteSpace(cabinetName) ? "箱柜1" : cabinetName.Trim();

                    // A 列超链接跳转至明细信息行定义名称 (Cab_Det_K) 并显示箱柜序号
                    // 屏幕提示明确标注：点击进入本箱柜明细表 --硬编码: 屏幕提示文本--
                    catSheet.Hyperlinks.Add(
                        Anchor: catSheet.Range[$"A{cabSumRow}"],
                        Address: "",
                        SubAddress: $"'{sheetName}'!{detPrefix}{cabinetIndex}",
                        ScreenTip: "点击进入本箱柜明细表",
                        TextToDisplay: Convert.ToString(cabinetIndex)
                    );

                    // 填入箱柜名称
                    catSheet.Cells[cabSumRow, 2].Value = safeCabName;
                    // G 列单价公式指向明细总计行的销售总价 (H 列)
                    catSheet.Cells[cabSumRow, 7].Formula = $"=H{cabTolsumRow - 1}";
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
                    // 获取安全的箱柜名称
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
                    // 激活工作表
                    catSheet.Activate();
                    // 默认选中 A1 单元格
                    catSheet.Range["A1"].Select();
                }
                catch { }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"初始化分类工作表异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行编辑分类工作表名称核心业务（修改表名 -> 同步项目信息表超链接与公式）
        /// </summary>
        /// <param name="request">编辑分类请求对象</param>
        /// <param name="explicitApp">可选显式传入的 Excel COM Application 实例</param>
        /// <returns>操作执行结果</returns>
        public static CategoryOperationResult EditCategoryName(EditCategoryRequest request, dynamic? explicitApp = null)
        {
            // 基础入参校验
            if (request == null || string.IsNullOrWhiteSpace(request.OldCategoryName) || string.IsNullOrWhiteSpace(request.NewCategoryName))
            {
                return new CategoryOperationResult { Success = false, Message = "原分类名称与新分类名称均不能为空！" };
            }

            string oldName = request.OldCategoryName.Trim();
            string newName = request.NewCategoryName.Trim();

            // 若新旧名称一致则无需变更
            if (string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase))
            {
                return new CategoryOperationResult { Success = true, Message = "分类名称未发生改变", CategoryName = newName };
            }

            try
            {
                // 获取 Excel COM Application 接口
                dynamic? app = explicitApp ?? ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveWorkbook == null) return new CategoryOperationResult { Success = false, Message = "无法连接 Excel 实例" };

                dynamic activeWb = app.ActiveWorkbook;

                // 检查旧工作表是否存在
                dynamic? targetSheet = null;
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    if (string.Equals(Convert.ToString(ws.Name)?.Trim(), oldName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetSheet = ws;
                        break;
                    }
                }

                if (targetSheet == null)
                {
                    return new CategoryOperationResult { Success = false, Message = $"未找到原分类工作表【{oldName}】！" };
                }

                // 检查新工作表名是否已存在同名冲突
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    if (string.Equals(Convert.ToString(ws.Name)?.Trim(), newName, StringComparison.OrdinalIgnoreCase))
                    {
                        return new CategoryOperationResult { Success = false, Message = $"工作簿中已存在名为【{newName}】的工作表，请更换新名称！" };
                    }
                }

                // 临时关闭屏幕刷新与事件处理
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                try
                {
                    // 1. 重命名目标分类工作表 (Excel 会自动级联更新大部分公式中的表名引用)
                    targetSheet.Name = newName;

                    // 2. 同步更新【项目信息】表中的分类汇总行超链接与公式文本
                    RenameProjectInfoCategorySummary(activeWb, oldName, newName);

                    // 3. 重新激活并聚焦
                    targetSheet.Activate();

                    // 若被复制的暂存表恰好为原表，同步更新暂存名称
                    if (string.Equals(_copiedCategorySheetName, oldName, StringComparison.OrdinalIgnoreCase))
                    {
                        _copiedCategorySheetName = newName;
                    }

                    return new CategoryOperationResult
                    {
                        Success = true,
                        Message = $"分类【{oldName}】已成功重命名为【{newName}】！",
                        CategoryName = newName
                    };
                }
                finally
                {
                    // 恢复屏幕刷新与告警
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"编辑分类名称异常: {ex.Message}");
                return new CategoryOperationResult { Success = false, Message = $"编辑分类名称失败: {ex.Message}" };
            }
        }

        /// <summary>
        /// 执行删除指定分类工作表核心业务（联动移除项目信息表汇总行，防止公式出现 #REF! 损坏）
        /// </summary>
        /// <param name="targetCategoryName">目标删除的分类工作表名称（为空则取当前活动工作表）</param>
        /// <param name="explicitApp">可选显式传入的 Excel COM Application 实例</param>
        /// <returns>操作执行结果</returns>
        public static CategoryOperationResult DeleteCategory(string? targetCategoryName = null, dynamic? explicitApp = null)
        {
            try
            {
                // 获取 Excel COM Application 实例
                dynamic? app = explicitApp ?? ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveWorkbook == null) return new CategoryOperationResult { Success = false, Message = "当前无活动工作簿" };

                dynamic activeWb = app.ActiveWorkbook;

                // 若未指定表名则默认取当前激活表
                string catName = string.IsNullOrWhiteSpace(targetCategoryName)
                    ? (Convert.ToString(app.ActiveSheet?.Name)?.Trim() ?? string.Empty)
                    : targetCategoryName.Trim();

                if (string.IsNullOrWhiteSpace(catName))
                {
                    return new CategoryOperationResult { Success = false, Message = "未指定要删除的目标分类！" };
                }

                // 严禁删除系统主表中枢 --硬编码--
                if (string.Equals(catName, "项目信息", StringComparison.OrdinalIgnoreCase))
                {
                    return new CategoryOperationResult { Success = false, Message = "【项目信息】是系统主表中枢，严禁删除！" };
                }

                // 查找目标工作表
                dynamic? sheetToDelete = null;
                int totalCatSheets = 0;
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    string name = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                    if (string.Equals(name, catName, StringComparison.OrdinalIgnoreCase))
                    {
                        sheetToDelete = ws;
                    }
                    if (!string.Equals(name, "项目信息", StringComparison.OrdinalIgnoreCase))
                    {
                        totalCatSheets++;
                    }
                }

                if (sheetToDelete == null)
                {
                    return new CategoryOperationResult { Success = false, Message = $"未找到待删除的分类表【{catName}】！" };
                }

                // 校验项目内分类表保有量，至少保留一个分类表
                if (totalCatSheets <= 1)
                {
                    return new CategoryOperationResult { Success = false, Message = "项目中至少需要保留一个分类工作表，无法继续删除！" };
                }

                // 暂停界面刷新与删除告警对话框
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                try
                {
                    // 1. 在【项目信息】表中精准物理删除对应的分类汇总行 (整行删除，A列序号自愈，杜绝 #REF!)
                    RemoveProjectInfoCategorySummary(activeWb, catName);

                    // 2. 物理删除分类工作表
                    sheetToDelete.Delete();

                    // 若被复制的暂存表恰好是被删除的表，清空暂存
                    if (string.Equals(_copiedCategorySheetName, catName, StringComparison.OrdinalIgnoreCase))
                    {
                        _copiedCategorySheetName = string.Empty;
                    }

                    // 3. 激活第一个剩余的分类工作表或项目信息表
                    try
                    {
                        foreach (dynamic ws in activeWb.Worksheets)
                        {
                            string wName = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                            if (!string.Equals(wName, "项目信息", StringComparison.OrdinalIgnoreCase))
                            {
                                ws.Activate();
                                break;
                            }
                        }
                    }
                    catch { }

                    return new CategoryOperationResult
                    {
                        Success = true,
                        Message = $"分类【{catName}】及【项目信息】中的对应汇总已彻底删除！",
                        CategoryName = catName
                    };
                }
                finally
                {
                    // 恢复刷新与事件调度
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"删除分类业务执行异常: {ex.Message}");
                return new CategoryOperationResult { Success = false, Message = $"删除分类执行失败: {ex.Message}" };
            }
        }

        /// <summary>
        /// 执行插入复制的分类核心业务（克隆源工作表 -> 重新分配全局箱柜定义名称 -> 联动项目信息表）
        /// </summary>
        /// <param name="request">插入复制分类请求对象</param>
        /// <param name="explicitApp">可选显式传入的 Excel COM Application 实例</param>
        /// <returns>操作执行结果</returns>
        public static CategoryOperationResult InsertCopiedCategory(InsertCopiedCategoryRequest request, dynamic? explicitApp = null)
        {
            // 入参校验
            if (request == null || string.IsNullOrWhiteSpace(request.TargetCategoryName))
            {
                return new CategoryOperationResult { Success = false, Message = "目标分类名称不能为空！" };
            }

            // 获取来源分类表名，未指定则取静态暂存表名
            string srcName = string.IsNullOrWhiteSpace(request.SourceCategoryName) ? _copiedCategorySheetName : request.SourceCategoryName.Trim();
            string targetName = request.TargetCategoryName.Trim();

            if (string.IsNullOrWhiteSpace(srcName))
            {
                return new CategoryOperationResult { Success = false, Message = "未指定来源被复制的分类，请先在源分类中点击【复制分类】！" };
            }

            try
            {
                // 获取 Excel COM Application 接口
                dynamic? app = explicitApp ?? ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveWorkbook == null) return new CategoryOperationResult { Success = false, Message = "当前无活动工作簿" };

                dynamic activeWb = app.ActiveWorkbook;

                // 检查源工作表是否存在
                dynamic? srcSheet = null;
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    string name = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                    if (string.Equals(name, srcName, StringComparison.OrdinalIgnoreCase))
                    {
                        srcSheet = ws;
                    }
                    // 校验目标新表名是否重名冲突
                    if (string.Equals(name, targetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return new CategoryOperationResult { Success = false, Message = $"工作簿中已存在名为【{targetName}】的工作表，请更换名称！" };
                    }
                }

                if (srcSheet == null)
                {
                    return new CategoryOperationResult { Success = false, Message = $"未找到被复制的源工作表【{srcName}】！" };
                }

                // 暂停界面刷新与告警
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                try
                {
                    // 1. 克隆源工作表并放置在最后一个工作表之后
                    srcSheet.Copy(After: activeWb.Sheets[activeWb.Sheets.Count]);
                    dynamic newSheet = activeWb.Sheets[activeWb.Sheets.Count];
                    // 重命名为用户指定的新名称
                    newSheet.Name = targetName;

                    // 2. 清洗跨工作簿外部公式链接 (将残留的源工作簿或外部路径洗成本地公式)
                    Tool.CleanRangeFormulas(newSheet.UsedRange);

                    // 3. 动态智能探测新分类表的基准行号分布
                    var rowIndexes = Tool.FindStandardCategoryRowIndexes((object)newSheet);
                    int cabSumRow = rowIndexes.cabSumRow;
                    int cabDetRow = rowIndexes.cabDetRow;
                    int subsumRow = rowIndexes.cabSubsumRow;
                    int cabTolsumRow = rowIndexes.cabTolsumRow;

                    // 读取箱柜定义名称前缀值对象
                    var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;

                    // 4. 重新扫描全工作簿分配新的全局唯一箱柜序号 K，杜绝跨表名称冲突 (规则 6 & 启发式经验 12)
                    int newCabinetK = GetNextCabinetIndex(activeWb, newSheet);

                    // 绑定新的 4 个工作表级定义名称锚点
                    Tool.SafeSetSheetName(newSheet, targetName, $"{sumPrefix}{newCabinetK}", cabSumRow);
                    Tool.SafeSetSheetName(newSheet, targetName, $"{detPrefix}{newCabinetK}", cabDetRow);
                    Tool.SafeSetSheetName(newSheet, targetName, $"{subsumPrefix}{newCabinetK}", subsumRow);
                    Tool.SafeSetSheetName(newSheet, targetName, $"{tolsumPrefix}{newCabinetK}", cabTolsumRow);

                    // 5. 重新配置顶部汇总行与底部明细行超链接
                    try
                    {
                        // 顶部 A 列超链接指向明细行定义名称
                        // 屏幕提示明确标注：点击进入本箱柜明细表 --硬编码: 屏幕提示文本--
                        newSheet.Hyperlinks.Add(
                            Anchor: newSheet.Range[$"A{cabSumRow}"],
                            Address: "",
                            SubAddress: $"'{targetName}'!{detPrefix}{newCabinetK}",
                            ScreenTip: "点击进入本箱柜明细表",
                            TextToDisplay: Convert.ToString(newCabinetK)
                        );

                        // 底部 A 列超链接指向汇总行定义名称
                        newSheet.Hyperlinks.Add(
                            Anchor: newSheet.Range[$"A{cabDetRow}"],
                            Address: "",
                            SubAddress: $"'{targetName}'!{sumPrefix}{newCabinetK}",
                            TextToDisplay: "柜号:"
                        );
                    }
                    catch { }

                    // 6. 在【项目信息】工作表中追加或更新分类汇总行
                    UpdateProjectInfoCategorySummary(activeWb, targetName);

                    // 7. 激活聚焦新工作表
                    newSheet.Activate();
                    newSheet.Range["A1"].Select();

                    return new CategoryOperationResult
                    {
                        Success = true,
                        Message = $"分类【{targetName}】已成功复制并插入！",
                        CategoryName = targetName
                    };
                }
                finally
                {
                    // 恢复刷新与事件调度
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"插入复制分类执行异常: {ex.Message}");
                return new CategoryOperationResult { Success = false, Message = $"插入复制分类失败: {ex.Message}" };
            }
        }

        /// <summary>
        /// 公共通用方法：在【项目信息】工作表的【分类汇总】区域中追加或更新指定分类的联动汇总行
        /// 具备防覆盖插行保护机制与与 ExWinner 母版 100% 对齐的标准 C~H 列公式体系
        /// </summary>
        /// <param name="targetWb">目标工作簿 COM 实例</param>
        /// <param name="categorySheetName">分类工作表名称</param>
        public static void UpdateProjectInfoCategorySummary(dynamic targetWb, string categorySheetName)
        {
            // 基础参数校验
            if (targetWb == null || string.IsNullOrWhiteSpace(categorySheetName)) return;

            try
            {
                // 获取【项目信息】工作表
                dynamic infoSheet = null;
                try { infoSheet = targetWb.Sheets["项目信息"]; } catch { }
                if (infoSheet == null) return;

                // 读取配置中的分类汇总起始物理行号与最大扫描数
                var cfg = ConfigManager.Instance.Current.Excel;
                int startRow = cfg.ProjectInfoCategorySummaryStartRow;
                int maxScan = cfg.ProjectInfoCategorySummaryMaxScanRows;

                // 寻找【分类汇总】区域的下一个可用行 (从配置的 startRow 开始向下扫描)
                int targetInfoRow = startRow;
                int maxRow = startRow + maxScan;
                bool needInsertRow = false;

                while (targetInfoRow < maxRow)
                {
                    // 读取 B 列单元格分类名称与内容
                    string cellB = Convert.ToString(infoSheet.Cells[targetInfoRow, 2].Value)?.Trim() ?? "";

                    // 若找到同名分类行，直接复用该物理行覆盖更新
                    if (string.Equals(cellB, categorySheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }

                    // 若遇到“小计”行，说明预留的分类空行已经全部用满，必须在此处向上插入物理行以保护小计不被覆盖
                    if (cellB.Contains("小计"))
                    {
                        needInsertRow = true;
                        break;
                    }

                    // 若遇到空白行，进一步检查下一行是否是小计，确保处于安全区域
                    if (string.IsNullOrEmpty(cellB))
                    {
                        break;
                    }

                    // 行号递增继续扫描
                    targetInfoRow++;
                }

                // 若遇小计行触发插行保护：在小计行前插入一个物理行 (向下推移 -4121)，小计行动态 INDEX 自动包裹该新行
                if (needInsertRow)
                {
                    infoSheet.Rows[targetInfoRow].Insert(-4121);
                }

                // 1. 挂载 A 列超链接并写入序号公式: 对齐 ExWinner 原生机制，点击 A 列序号数字跳转分类表
                // 屏幕提示明确标注：点击进入本分类报价清单 --硬编码: 屏幕提示文本--
                int headerRowIndex = startRow - 1;
                infoSheet.Hyperlinks.Add(
                    Anchor: infoSheet.Range[$"A{targetInfoRow}"],
                    Address: "",
                    SubAddress: $"'{categorySheetName}'!A1",
                    ScreenTip: "点击进入本分类报价清单"
                );
                // 写入 A 列公式：跟随起始行上方表头自动动态计算序号，计算结果保留超链接属性与蓝色下划线
                infoSheet.Cells[targetInfoRow, 1].Formula = $"=ROW()-ROW(A${headerRowIndex})";

                // 2. 写入 B 列分类名称动态公式: 对齐 ExWinner 原生机制，根据工作表引用动态解析分类名
                // 使用 CELL("filename") 与 FIND 提取工作表名，实现随 Sheet 改名自动级联联动
                try
                {
                    // 删除 B 列可能存在的超链接，避免蓝字与下划线干扰
                    infoSheet.Range[$"B{targetInfoRow}"].Hyperlinks.Delete();
                }
                catch { }
                // 写入动态工作表名公式 --硬编码: Excel公式函数名--
                infoSheet.Cells[targetInfoRow, 2].Formula = $"=MID(CELL(\"filename\",'{categorySheetName}'!$A$1),FIND(\"]\",CELL(\"filename\",'{categorySheetName}'!$A$1))+1,31)";

                // 3. 写入 C 列箱柜数量公式: 对齐 ExWinner 母版公式，结合 H 列属性与合计行 F35 联动
                infoSheet.Cells[targetInfoRow, 3].Formula = $"=IF(H{targetInfoRow}=\"箱变\",'{categorySheetName}'!F35,'{categorySheetName}'!F31*IF('{categorySheetName}'!F35=\"\",1,'{categorySheetName}'!F35))";

                // 4. 写入 D 列总价公式: 精准直连分类表合计行销售总价 H35，保留 2 位小数
                infoSheet.Cells[targetInfoRow, 4].Formula = $"=ROUND('{categorySheetName}'!H35,2)";

                // 5. 写入 E 列成本总价公式: 直连分类表合计行成本总价 J35，保留 2 位小数
                infoSheet.Cells[targetInfoRow, 5].Formula = $"=ROUND('{categorySheetName}'!J35,2)";

                // 6. 写入 F 列毛利公式: 销售总价 - 成本总价
                infoSheet.Cells[targetInfoRow, 6].Formula = $"=ROUND((D{targetInfoRow}-E{targetInfoRow}),2)";

                // 7. 写入 G 列毛利率公式: 毛利 / 销售总价 (防 0 除保护)
                infoSheet.Cells[targetInfoRow, 7].Formula = $"=IF(D{targetInfoRow}=0,0,F{targetInfoRow}/D{targetInfoRow})";

                // 8. 写入 H 列分类属性公式: 智能判定箱变/欧变/美变/常规类型
                infoSheet.Cells[targetInfoRow, 8].Formula = $"=IF(OR(ISNUMBER(FIND(\"箱变\",B{targetInfoRow}))=TRUE,ISNUMBER(FIND(\"欧变\",B{targetInfoRow}))=TRUE,ISNUMBER(FIND(\"美变\",B{targetInfoRow}))=TRUE,ISNUMBER(FIND(\"KVA\",UPPER(B{targetInfoRow})))=TRUE,ISNUMBER(FIND(\"箱式变电\",B{targetInfoRow}))=TRUE),\"箱变\",\"常规\")";
            }
            catch (Exception ex)
            {
                // 记录更新异常日志
                LogHelper.WriteLog($"更新项目信息表分类汇总异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 公共通用方法：在【项目信息】工作表中根据旧分类名称更新超链接与公式
        /// </summary>
        /// <param name="targetWb">目标工作簿 COM 实例</param>
        /// <param name="oldName">原旧分类工作表名称</param>
        /// <param name="newName">新分类工作表名称</param>
        public static void RenameProjectInfoCategorySummary(dynamic targetWb, string oldName, string newName)
        {
            if (targetWb == null || string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName)) return;

            try
            {
                // 获取【项目信息】工作表
                dynamic infoSheet = null;
                try { infoSheet = targetWb.Sheets["项目信息"]; } catch { }
                if (infoSheet == null) return;

                // 读取配置起始行与最大扫描数
                var cfg = ConfigManager.Instance.Current.Excel;
                int startRow = cfg.ProjectInfoCategorySummaryStartRow;
                int maxScan = cfg.ProjectInfoCategorySummaryMaxScanRows;
                int headerRowIndex = startRow - 1;

                // 遍历寻找旧分类名所在的行
                for (int r = startRow; r < startRow + maxScan; r++)
                {
                    string cellB = Convert.ToString(infoSheet.Cells[r, 2].Value)?.Trim() ?? "";
                    // 判断匹配旧分类名，或者从 A 列超链接匹配
                    bool isMatched = string.Equals(cellB, oldName, StringComparison.OrdinalIgnoreCase);
                    if (!isMatched)
                    {
                        try
                        {
                            dynamic aCell = infoSheet.Cells[r, 1];
                            if (aCell.Hyperlinks != null && aCell.Hyperlinks.Count > 0)
                            {
                                string subAddr = Convert.ToString(aCell.Hyperlinks[1].SubAddress) ?? "";
                                if (subAddr.Contains($"'{oldName}'") || subAddr.Contains($"{oldName}!"))
                                {
                                    isMatched = true;
                                }
                            }
                        }
                        catch { }
                    }

                    if (isMatched)
                    {
                        // 1. 更新 A 列超链接：SubAddress 指向最新分类表，保留屏幕提示 --硬编码: 屏幕提示文本--
                        try
                        {
                            infoSheet.Hyperlinks.Add(
                                Anchor: infoSheet.Range[$"A{r}"],
                                Address: "",
                                SubAddress: $"'{newName}'!A1",
                                ScreenTip: "点击进入本分类报价清单"
                            );
                            // 保持 A 列序号公式
                            infoSheet.Cells[r, 1].Formula = $"=ROW()-ROW(A${headerRowIndex})";
                        }
                        catch { }

                        // 2. 更新 B 列分类名称动态公式：重新写入指向新分类工作表 A1 的动态联动公式
                        try
                        {
                            // 删除可能残留的超链接
                            infoSheet.Range[$"B{r}"].Hyperlinks.Delete();
                        }
                        catch { }
                        // 写入最新分类表的动态工作表名公式 --硬编码: Excel公式函数名--
                        infoSheet.Cells[r, 2].Formula = $"=MID(CELL(\"filename\",'{newName}'!$A$1),FIND(\"]\",CELL(\"filename\",'{newName}'!$A$1))+1,31)";

                        // 3. 重新写入 C/D/E 列联动公式，确保指向最新表名
                        infoSheet.Cells[r, 3].Formula = $"=IF(H{r}=\"箱变\",'{newName}'!F35,'{newName}'!F31*IF('{newName}'!F35=\"\",1,'{newName}'!F35))";
                        infoSheet.Cells[r, 4].Formula = $"=ROUND('{newName}'!H35,2)";
                        infoSheet.Cells[r, 5].Formula = $"=ROUND('{newName}'!J35,2)";

                        // 4. 重新触发行属性判断公式
                        infoSheet.Cells[r, 8].Formula = $"=IF(OR(ISNUMBER(FIND(\"箱变\",B{r}))=TRUE,ISNUMBER(FIND(\"欧变\",B{r}))=TRUE,ISNUMBER(FIND(\"美变\",B{r}))=TRUE,ISNUMBER(FIND(\"KVA\",UPPER(B{r})))=TRUE,ISNUMBER(FIND(\"箱式变电\",B{r}))=TRUE),\"箱变\",\"常规\")";
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录重命名同步异常日志
                LogHelper.WriteLog($"更新项目信息表分类重命名异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 公共通用方法：在【项目信息】工作表中彻底物理删除指定分类对应的汇总行
        /// （整行删除机制使得 A 列动态序号自动平滑递补，杜绝产生 #REF! 损坏）
        /// </summary>
        /// <param name="targetWb">目标工作簿 COM 实例</param>
        /// <param name="categorySheetName">待移除的分类工作表名称</param>
        public static void RemoveProjectInfoCategorySummary(dynamic targetWb, string categorySheetName)
        {
            if (targetWb == null || string.IsNullOrWhiteSpace(categorySheetName)) return;

            try
            {
                // 获取【项目信息】工作表
                dynamic infoSheet = null;
                try { infoSheet = targetWb.Sheets["项目信息"]; } catch { }
                if (infoSheet == null) return;

                // 读取配置起始行与最大扫描数
                var cfg = ConfigManager.Instance.Current.Excel;
                int startRow = cfg.ProjectInfoCategorySummaryStartRow;
                int maxScan = cfg.ProjectInfoCategorySummaryMaxScanRows;

                // 遍历寻找匹配该分类名的物理行
                for (int r = startRow; r < startRow + maxScan; r++)
                {
                    string cellB = Convert.ToString(infoSheet.Cells[r, 2].Value)?.Trim() ?? "";
                    if (string.Equals(cellB, categorySheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        // 找到该分类汇总行，执行整行物理删除 (后方行自动上移，序号公式自愈，小计动态包裹)
                        infoSheet.Rows[r].EntireRow.Delete();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录删除同步异常日志
                LogHelper.WriteLog($"移除项目信息表分类汇总行异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 公共通用方法：自愈并规范化【项目信息】工作表中的分类汇总超链接与公式结构
        /// 彻底修复历史文件可能存在的：A列缺超链接、B列错挂超链接、B列公式由于名称缺失报 #NAME? 等缺陷
        /// </summary>
        /// <param name="targetWb">目标工作簿 COM 实例</param>
        public static void NormalizeCategorySummaryLinks(dynamic targetWb)
        {
            if (targetWb == null) return;

            try
            {
                // 获取【项目信息】工作表
                dynamic infoSheet = null;
                try { infoSheet = targetWb.Sheets["项目信息"]; } catch { }
                if (infoSheet == null) return;

                // 读取配置中分类汇总起始物理行号与最大扫描数
                var cfg = ConfigManager.Instance.Current.Excel;
                int startRow = cfg.ProjectInfoCategorySummaryStartRow;
                int maxScan = cfg.ProjectInfoCategorySummaryMaxScanRows;
                int headerRowIndex = startRow - 1;

                // 获取工作簿中所有现存的分类工作表名称集合 (排除系统保留表)
                var existingCategorySheets = new List<string>();
                foreach (dynamic ws in targetWb.Worksheets)
                {
                    string wsName = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                    // 排除系统辅助工作表 --硬编码: 系统辅助工作表名--
                    if (!string.Equals(wsName, "项目信息", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(wsName, "封面", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(wsName, "元件汇总表", StringComparison.OrdinalIgnoreCase))
                    {
                        existingCategorySheets.Add(wsName);
                    }
                }

                // 若当前无有效分类表，直接返回
                if (existingCategorySheets.Count == 0) return;

                // 遍历扫描分类汇总区域
                for (int r = startRow; r < startRow + maxScan; r++)
                {
                    // 读取 B 列单元格的值与公式
                    string cellBVal = Convert.ToString(infoSheet.Cells[r, 2].Value)?.Trim() ?? "";
                    string cellBFormula = Convert.ToString(infoSheet.Cells[r, 2].Formula)?.Trim() ?? "";

                    // 若遇到小计行，说明分类汇总数据区域已结束
                    if (cellBVal.Contains("小计")) break;

                    // 判断该行是否为有效分类行：尝试提取对应的分类表名
                    string matchedCatName = string.Empty;

                    // 1. 如果 B 列有现成的有效分类名称且匹配现存工作表
                    if (!string.IsNullOrEmpty(cellBVal) && !cellBVal.StartsWith("#") && existingCategorySheets.Contains(cellBVal, StringComparer.OrdinalIgnoreCase))
                    {
                        matchedCatName = existingCategorySheets.First(s => string.Equals(s, cellBVal, StringComparison.OrdinalIgnoreCase));
                    }

                    // 2. 若 B 列显示为 #NAME? 或无效，尝试从 A 列或 B 列既有超链接 SubAddress 中解析分类工作表
                    if (string.IsNullOrEmpty(matchedCatName))
                    {
                        // 检查 A 列单元格超链接
                        dynamic aCell = infoSheet.Cells[r, 1];
                        if (aCell.Hyperlinks != null && aCell.Hyperlinks.Count > 0)
                        {
                            string subAddr = Convert.ToString(aCell.Hyperlinks[1].SubAddress) ?? "";
                            matchedCatName = ExtractSheetNameFromSubAddress(subAddr, existingCategorySheets);
                        }

                        // 检查 B 列单元格超链接
                        if (string.IsNullOrEmpty(matchedCatName))
                        {
                            dynamic bCell = infoSheet.Cells[r, 2];
                            if (bCell.Hyperlinks != null && bCell.Hyperlinks.Count > 0)
                            {
                                string subAddr = Convert.ToString(bCell.Hyperlinks[1].SubAddress) ?? "";
                                matchedCatName = ExtractSheetNameFromSubAddress(subAddr, existingCategorySheets);
                            }
                        }

                        // 3. 若仍未匹配且是第 r - startRow 个分类，按物理顺序匹配现存分类工作表
                        if (string.IsNullOrEmpty(matchedCatName))
                        {
                            int catSeqIdx = r - startRow;
                            if (catSeqIdx >= 0 && catSeqIdx < existingCategorySheets.Count)
                            {
                                matchedCatName = existingCategorySheets[catSeqIdx];
                            }
                        }
                    }

                    // 若未找到对应的分类表且单元格为空，说明到达空白区间，结束巡检
                    if (string.IsNullOrEmpty(matchedCatName))
                    {
                        if (string.IsNullOrEmpty(cellBVal)) break;
                        continue;
                    }

                    // 4. 规范化 A 列：挂载超链接指向 matchedCatName，并确保公式为 =ROW()-ROW(A$28)
                    try
                    {
                        // 挂载超链接并设置屏幕提示 --硬编码: 屏幕提示文本--
                        infoSheet.Hyperlinks.Add(
                            Anchor: infoSheet.Range[$"A{r}"],
                            Address: "",
                            SubAddress: $"'{matchedCatName}'!A1",
                            ScreenTip: "点击进入本分类报价清单"
                        );
                        // 写入动态序号公式
                        infoSheet.Cells[r, 1].Formula = $"=ROW()-ROW(A${headerRowIndex})";
                    }
                    catch { }

                    // 5. 规范化 B 列：坚决删除超链接，写入动态工作表名称公式，彻底消除蓝字下划线与 #NAME?
                    try
                    {
                        // 删除超链接
                        infoSheet.Range[$"B{r}"].Hyperlinks.Delete();
                    }
                    catch { }
                    // 写入指向匹配分类表的动态工作表名公式 --硬编码: Excel公式函数名--
                    infoSheet.Cells[r, 2].Formula = $"=MID(CELL(\"filename\",'{matchedCatName}'!$A$1),FIND(\"]\",CELL(\"filename\",'{matchedCatName}'!$A$1))+1,31)";

                    // 6. 确保 C~H 列公式有效对齐
                    try
                    {
                        // C 列箱柜数量公式
                        infoSheet.Cells[r, 3].Formula = $"=IF(H{r}=\"箱变\",'{matchedCatName}'!F35,'{matchedCatName}'!F31*IF('{matchedCatName}'!F35=\"\",1,'{matchedCatName}'!F35))";
                        // D 列总价公式
                        infoSheet.Cells[r, 4].Formula = $"=ROUND('{matchedCatName}'!H35,2)";
                        // E 列成本总价公式
                        infoSheet.Cells[r, 5].Formula = $"=ROUND('{matchedCatName}'!J35,2)";
                        // F 列毛利公式
                        infoSheet.Cells[r, 6].Formula = $"=ROUND((D{r}-E{r}),2)";
                        // G 列毛利率公式
                        infoSheet.Cells[r, 7].Formula = $"=IF(D{r}=0,0,F{r}/D{r})";
                        // H 列分类属性公式
                        infoSheet.Cells[r, 8].Formula = $"=IF(OR(ISNUMBER(FIND(\"箱变\",B{r}))=TRUE,ISNUMBER(FIND(\"欧变\",B{r}))=TRUE,ISNUMBER(FIND(\"美变\",B{r}))=TRUE,ISNUMBER(FIND(\"KVA\",UPPER(B{r})))=TRUE,ISNUMBER(FIND(\"箱式变电\",B{r}))=TRUE),\"箱变\",\"常规\")";
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                // 记录自愈异常日志
                LogHelper.WriteLog($"规范化分类汇总超链接与公式异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 从超链接 SubAddress 字符串中安全提取匹配的工作表名称
        /// </summary>
        /// <param name="subAddr">超链接子地址</param>
        /// <param name="sheetNames">有效工作表名称集合</param>
        /// <returns>匹配的工作表名称</returns>
        private static string ExtractSheetNameFromSubAddress(string subAddr, List<string> sheetNames)
        {
            if (string.IsNullOrWhiteSpace(subAddr) || sheetNames == null) return string.Empty;
            // 常见格式: '分类1'!A1 或 分类1!A1
            string clean = subAddr.Split('!')[0].Trim('\'', ' ');
            return sheetNames.FirstOrDefault(s => string.Equals(s, clean, StringComparison.OrdinalIgnoreCase)) ?? string.Empty;
        }
    }
}
