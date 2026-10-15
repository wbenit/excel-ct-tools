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

        // 缓存批量删除分类窗体单例引用，避免重复打开
        private static Forms.DeleteCategoryForm? _deleteCategoryFormInstance;

        // 缓存合并分类窗体单例引用，避免重复打开
        private static Forms.MergeCategoryForm? _mergeCategoryFormInstance;

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

                // 若当前聚焦在【项目信息】表，智能调起“一次性删除一个或多个分类”管理窗口
                if (string.Equals(sheetName, "项目信息", StringComparison.OrdinalIgnoreCase))
                {
                    // 弹出基于 WebView2 + Vue 3 + Element Plus 的批量删除分类窗口
                    ShowDeleteCategoryDialog();
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
        /// 供 Ribbon 菜单直接调用的“删除无数量的箱柜”入口
        /// 严格遵循规则 6 (4个定义名称拓扑)、规则 7 (数组一次性读入内存)、规则 8 (操作前预检自愈) 与规则 9 (不符合要求直接报错提醒)
        /// </summary>
        public static void DeleteZeroQuantityCabinets()
        {
            try
            {
                // 获取当前正在运行的 Excel 运行上下文 (安全调度)
                var context = Tool.GetActiveExcelContext();
                // 校验宿主上下文有效性
                if (context == null) return;
                // 提取 Excel Application COM 句柄
                dynamic app = context.App;
                // 提取活动工作簿与活动工作表
                dynamic activeWb = context.Wb;
                dynamic activeSheet = context.Sheet;

                // 读取当前活动工作表名称
                string sheetName = Convert.ToString(activeSheet.Name)?.Trim() ?? string.Empty;

                // 规则 9: 不符合要求直接报错提醒，不写防御兼容代码
                if (string.Equals(sheetName, "项目信息", StringComparison.OrdinalIgnoreCase) || Tool.IsReservedOrReportSheet(sheetName))
                {
                    // 弹出明确的不支持提示弹窗 --硬编码: 系统提示文本--
                    System.Windows.Forms.MessageBox.Show(
                        "当前工作表不是分类明细表，无法执行删除无数量箱柜操作！",
                        "系统提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // 严格遵循规则 8：在操作 Excel 表格前，检查当前表格规则 6 的正确性，先纠正自愈后再操作
                Tool.FixAndFillCabinetNamesForSheet(activeSheet, forceRebuild: false);

                // 获取当前工作表经过规则 6 校验的全部有效箱柜映射字典
                List<KeyValuePair<int, CabinetAnchorModel>> validCabinets = Tool.GetSheetValidCabinets(activeSheet, activeWb);
                // 校验工作表中是否存在有效箱柜
                if (validCabinets == null || validCabinets.Count == 0)
                {
                    // 规则 9: 未识别到箱柜直接报错提醒
                    System.Windows.Forms.MessageBox.Show(
                        "当前工作表未识别到任何箱柜，无法执行操作！",
                        "系统提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // 严格遵守规则 7：读写多个区域，采用数组一次性读到内存
                // 计算顶部汇总行的最小与最大物理行号范围
                int minSumRow = int.MaxValue;
                int maxSumRow = int.MinValue;
                // 遍历搜集所有汇总行的行号
                foreach (var cab in validCabinets)
                {
                    // 提取汇总行锚点
                    if (cab.Value.Sum != null)
                    {
                        // 转换物理行号
                        int r = Convert.ToInt32(cab.Value.Sum.Row);
                        // 更新最小行号
                        if (r < minSumRow) minSumRow = r;
                        // 更新最大行号
                        if (r > maxSumRow) maxSumRow = r;
                    }
                }

                // 规则 9: 校验行号边界，异常则直接报错提示
                if (minSumRow > maxSumRow || minSumRow < 1)
                {
                    // 弹出错误提示 --硬编码: 系统提示文本--
                    System.Windows.Forms.MessageBox.Show(
                        "未找到有效的箱柜顶部汇总行，无法执行操作！",
                        "系统提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // 一次性读取顶部汇总行区域 (A列到F列，包含B列柜号、C列名称与F列数量)
                dynamic sumRange = activeSheet.Range[$"A{minSumRow}:F{maxSumRow}"];
                // 声明承载内存数据的二维数组
                object[,] dataMatrix;
                // 校验并接收 COM 区域的二维数组
                if (sumRange.Value2 is object[,] arr)
                {
                    // 直接赋给内存矩阵
                    dataMatrix = arr;
                }
                else
                {
                    // 单单元格场景保底构建 1 行 6 列数组
                    dataMatrix = new object[2, 7];
                    // 提取第 6 列数量
                    dataMatrix[1, 6] = sumRange.Value2;
                }

                // 初始化待删除箱柜序号与展示文本集合
                var zeroQtyKs = new List<int>();
                var zeroQtyNames = new List<string>();

                // 遍历所有箱柜判定其数量
                foreach (var cab in validCabinets)
                {
                    // 获取当前箱柜序号 K
                    int k = cab.Key;
                    // 获取箱柜锚点数据
                    var anchor = cab.Value;
                    // 若无汇总锚点则跳过
                    if (anchor.Sum == null) continue;

                    // 计算在内存矩阵中的相对行号 (1-based)
                    int sumRow = Convert.ToInt32(anchor.Sum.Row);
                    int matrixRow = sumRow - minSumRow + 1;

                    // 内存读取 A 列序号(第1列) 与 B 列箱柜名称(第2列)
                    string cabNo = Convert.ToString(dataMatrix[matrixRow, 1])?.Trim() ?? string.Empty;
                    string cabName = Convert.ToString(dataMatrix[matrixRow, 2])?.Trim() ?? string.Empty;
                    // 内存读取 F 列数量对象
                    object? qtyObj = dataMatrix[matrixRow, 6];

                    // 判定是否为无数量箱柜：①未填/空白、②数值等于0、③小于0
                    bool isZeroQty = false;
                    // 检查单元格对象是否为空
                    if (qtyObj == null)
                    {
                        // 单元格空即无数量
                        isZeroQty = true;
                    }
                    else
                    {
                        // 转换为字符串进行文本与数值双重比对
                        string qtyStr = Convert.ToString(qtyObj)?.Trim() ?? string.Empty;
                        // 空字符串视为无数量
                        if (string.IsNullOrEmpty(qtyStr))
                        {
                            isZeroQty = true;
                        }
                        // 尝试数值解析
                        else if (double.TryParse(qtyStr, out double dQty))
                        {
                            // 小于或等于 0 均视为无数量
                            if (dQty <= 0.000001)
                            {
                                isZeroQty = true;
                            }
                        }
                        else
                        {
                            // 非法字符无法转为数值，按无数量处理
                            isZeroQty = true;
                        }
                    }

                    // 若判定为无数量箱柜，加入待删除队列
                    if (isZeroQty)
                    {
                        // 记录箱柜序号
                        zeroQtyKs.Add(k);
                        // 拼接规范的显示名称 (例如: 序号 1: BXF1 或 序号 1)
                        string displayName;
                        if (!string.IsNullOrWhiteSpace(cabNo) && !string.IsNullOrWhiteSpace(cabName))
                        {
                            displayName = $"序号 {cabNo}: {cabName}";
                        }
                        else if (!string.IsNullOrWhiteSpace(cabName))
                        {
                            displayName = cabName;
                        }
                        else if (!string.IsNullOrWhiteSpace(cabNo))
                        {
                            displayName = $"序号 {cabNo}";
                        }
                        else
                        {
                            displayName = $"箱柜{k}";
                        }
                        // 添加至清单列表
                        zeroQtyNames.Add($"• 【{displayName}】(定义序号: {k})");
                    }
                }

                // 校验扫描结果：若没有无数量箱柜则提示并安全返回
                if (zeroQtyKs.Count == 0)
                {
                    // 提示用户无需删除 --硬编码: 系统提示文本--
                    System.Windows.Forms.MessageBox.Show(
                        $"当前分类表【{sheetName}】中所有箱柜数量均大于 0，未发现无数量的箱柜。",
                        "系统提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // 构建带箱柜清单的提示文本 (超过15台进行折叠)
                string listText = zeroQtyNames.Count <= 15
                    ? string.Join("\n", zeroQtyNames)
                    : string.Join("\n", zeroQtyNames.Take(15)) + $"\n...等共 {zeroQtyNames.Count} 台箱柜";

                // 组装二次确认消息文本
                string confirmMsg;
                // 区分全选所有箱柜与部分删除场景
                if (zeroQtyKs.Count >= validCabinets.Count)
                {
                    // 全选场景：提醒首台重置保护
                    confirmMsg = $"在当前分类表【{sheetName}】中，检测到全部 {zeroQtyKs.Count} 台箱柜的数量均为 0 或未填写：\n\n" +
                                 $"{listText}\n\n" +
                                 "⚠️ 注意：若删除全部箱柜，系统将重置保留首台初始空白箱柜。\n\n" +
                                 "确定要执行删除并重置吗？";
                }
                else
                {
                    // 部分删除场景：提示将被物理删除的箱柜
                    confirmMsg = $"在当前分类表【{sheetName}】中，共检测到 {zeroQtyKs.Count} 台数量为 0 或未填写的箱柜：\n\n" +
                                 $"{listText}\n\n" +
                                 "此操作将删除上述箱柜的顶部汇总行及对应的底表明细块，且不可恢复。\n\n" +
                                 "确定要执行删除吗？";
                }

                // 弹出模态二次确认对话框
                var dialogRes = System.Windows.Forms.MessageBox.Show(
                    confirmMsg,
                    "确认删除无数量箱柜",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question);

                // 若用户点击【否】或取消，直接终止退出
                if (dialogRes != System.Windows.Forms.DialogResult.Yes) return;

                // 调用公共删除引擎 DeleteCabinets 执行物理删除与定义名称注销
                bool success = DeleteCabinets(app, activeSheet, zeroQtyKs);
                // 校验删除执行结果
                if (success)
                {
                    // 弹出成功提示报告 --硬编码: 成功提示文本--
                    System.Windows.Forms.MessageBox.Show(
                        $"已成功删除分类表【{sheetName}】中 {zeroQtyKs.Count} 台无数量箱柜！",
                        "删除成功",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                // 记录错误日志
                LogHelper.WriteLog($"DeleteZeroQuantityCabinets 异常: {ex.Message}");
                // 弹出异常错误弹窗 --硬编码: 异常提示文本--
                System.Windows.Forms.MessageBox.Show(
                    $"删除无数量箱柜失败: {ex.Message}",
                    "系统提示",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
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

                // 保存调用前的 Excel 环境状态，防止在批量导出等外部上下文中提前泄露恢复全局事件
                bool prevScreenUpdating = true;
                bool prevDisplayAlerts = true;
                bool prevEnableEvents = true;
                // 安全提取原始环境状态
                try { prevScreenUpdating = app.ScreenUpdating; } catch { }
                try { prevDisplayAlerts = app.DisplayAlerts; } catch { }
                try { prevEnableEvents = app.EnableEvents; } catch { }

                // 临时关闭屏幕刷新与提示警告以提高初始化速度
                try { app.ScreenUpdating = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }
                try { app.EnableEvents = false; } catch { }

                try
                {
                    // 1. 从 CabinetTemplate.xlsx 标准模板克隆纯净的【分类1】工作表 (100% 完整保留汇总表头、合计行、大写、说明及明细模板)
                    string templatePath = ProjectController.EnsureCabinetTemplate(app);
                    // 打开模板时明确指定 UpdateLinks: 0，杜绝外部源链接安全警告弹窗
                    dynamic templateWb = app.Workbooks.Open(templatePath, UpdateLinks: 0, ReadOnly: true);
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
                    // 保护性还原 Excel 屏幕刷新、告警与事件调度（遵循进入前原始状态）
                    try { app.ScreenUpdating = prevScreenUpdating; } catch { }
                    try { app.DisplayAlerts = prevDisplayAlerts; } catch { }
                    try { app.EnableEvents = prevEnableEvents; } catch { }
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

                        // 方案 3: 提取与保护原总计行 F 列数量 (在覆盖前读取旧单元格，防清空台数)
                        object? preservedTolQty = null;
                        try
                        {
                            if (cabTolsumRow > 0)
                            {
                                dynamic oldTolCell = catSheet.Cells[cabTolsumRow, 6];
                                string oldFormula = Convert.ToString(oldTolCell.Formula) ?? "";
                                object oldVal = oldTolCell.Value2;
                                if (!string.IsNullOrWhiteSpace(oldFormula) && oldFormula.Trim() != "=")
                                    preservedTolQty = oldFormula;
                                else if (oldVal != null && !string.IsNullOrWhiteSpace(oldVal.ToString()) && Convert.ToString(oldVal) != "0")
                                    preservedTolQty = oldVal;
                            }
                            // 若原单元格为空且存在汇总行，自动联动汇总行 =F{cabSumRow}
                            if (preservedTolQty == null && cabSumRow > 0)
                            {
                                preservedTolQty = $"=F{cabSumRow}";
                            }
                        }
                        catch (Exception exReadTol)
                        {
                            LogHelper.WriteLog($"InitializeCategorySheet 读取原总计行数量异常: {exReadTol.Message}");
                        }

                        // 调用公共服务方法构建 N 行 17 列的计费二维矩阵 (覆盖 A 列至 Q 列，透传 cabSumRow 与 preservedTolQty)
                        object[,] feeMatrix = Tool.BuildFeeMatrix(items, cabDetRow, subsumRow, compStartRow, compEndRow, 17, cabSumRow, preservedTolQty);

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
                    // 写入汇总行单位 (E 列即第 5 列，默认 "台") --硬编码: 箱柜单位--
                    catSheet.Cells[cabSumRow, 5].Value = "台";
                    // 写入汇总行数量 (F 列即第 6 列，默认 1) --硬编码: 第 6 列为 F 列 (数量列)--
                    catSheet.Cells[cabSumRow, 6].Value = 1;
                    // 在 tolsum 总计行 F 列自动绑定联动公式指向顶部汇总行 F 列 (双向自适应联动)
                    catSheet.Cells[cabTolsumRow, 6].Formula = $"=F{cabSumRow}";
                    // G 列单价公式指向明细总计行的单台单价 G 列 (方式 B 稳健绑定，避免数量二次相乘)
                    catSheet.Cells[cabSumRow, 7].Formula = $"=G{cabTolsumRow}";
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

                    // 同步清空项目分类工作表白名单缓存
                    Tool.InvalidateProjectCategorySheetCache();

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
        /// 启动并弹出基于 WebView2 + Vue 3 的“删除分类”窗口 (非模态，支持一次性删除一个或多个分类)
        /// </summary>
        public static void ShowDeleteCategoryDialog()
        {
            try
            {
                // 获取当前正在运行的 Excel Application 对象 (安全调用)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                // 校验工作簿是否打开
                if (app == null || app.ActiveWorkbook == null)
                {
                    // 若无工作簿打开则提示用户
                    System.Windows.Forms.MessageBox.Show("请先打开或新建一个报价项目工作簿！", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // 以统一非模态方式展示删除分类窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _deleteCategoryFormInstance, () => new Forms.DeleteCategoryForm());
            }
            catch (Exception ex)
            {
                // 记录弹出窗体异常
                LogHelper.WriteLog($"弹出删除分类窗口异常: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"弹出删除分类窗口失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 获取待删除分类列表数据及当前选区智能感知信息 (全面支持正常分类表与 #REF! 僵尸失效行)
        /// </summary>
        /// <returns>待删除分类列表与选区命中数据</returns>
        public static DeleteCategoriesDataResponse GetDeleteCategoriesData()
        {
            var response = new DeleteCategoriesDataResponse();

            try
            {
                // 获取 Excel COM Application 接口实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveWorkbook == null) return response;

                dynamic activeWb = app.ActiveWorkbook;
                dynamic? activeSheet = app.ActiveSheet;
                string activeSheetName = Convert.ToString(activeSheet?.Name)?.Trim() ?? string.Empty;

                // 判断当前活动工作表是否为【项目信息】--硬编码: 项目信息工作表名--
                bool isProjectInfo = string.Equals(activeSheetName, "项目信息", StringComparison.OrdinalIgnoreCase);
                response.IsActiveSheetProjectInfo = isProjectInfo;

                // 准备选区行范围检测集合
                var selectedRowIndices = new HashSet<int>();
                try
                {
                    // 仅当当前处于【项目信息】表时提取选区行
                    if (isProjectInfo && app.Selection != null)
                    {
                        dynamic sel = app.Selection;
                        int selStartRow = sel.Row;
                        int selRowCount = sel.Rows.Count;
                        // 记录选区覆盖的行区间
                        for (int r = selStartRow; r < selStartRow + selRowCount; r++)
                        {
                            selectedRowIndices.Add(r);
                        }
                    }
                }
                catch { }

                // 系统保留工作表黑名单 (严禁作为分类删除) --硬编码: 系统保留表名--
                var reservedSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "项目信息", "封面", "元件汇总表", "材料分布表",
                    "元件汇总分布表", "元件汇总调价清单", "屏柜汇总表", "屏柜分项表", "元器件数据管理"
                };

                // 收集工作簿中现存的所有有效分类工作表名称集合
                var existingCategorySheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    string wsName = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(wsName) && !reservedSheets.Contains(wsName))
                    {
                        existingCategorySheetNames.Add(wsName);
                    }
                }

                // 1. 优先从【项目信息】表的分类汇总区域读取分类列表与金额/台数及 #REF! 损坏行
                dynamic? infoSheet = null;
                try { infoSheet = activeWb.Sheets["项目信息"]; } catch { }

                var detectedCategories = new Dictionary<string, DeleteCategoryItemDto>(StringComparer.OrdinalIgnoreCase);
                if (infoSheet != null)
                {
                    // 读取配置中的分类汇总起始物理行号与最大扫描数
                    var cfg = ConfigManager.Instance.Current.Excel;
                    int startRow = cfg.ProjectInfoCategorySummaryStartRow;
                    int maxScan = cfg.ProjectInfoCategorySummaryMaxScanRows;

                    // 扫描项目信息表分类汇总行
                    for (int r = startRow; r < startRow + maxScan; r++)
                    {
                        // 读取 B 列单元格的值、文本与公式
                        string cellB = Convert.ToString(infoSheet.Cells[r, 2].Value)?.Trim() ?? "";
                        string textB = string.Empty;
                        try { textB = Convert.ToString(infoSheet.Cells[r, 2].Text)?.Trim() ?? ""; } catch { }
                        string formulaB = Convert.ToString(infoSheet.Cells[r, 2].Formula)?.Trim() ?? "";

                        // 读取 C 列单元格的值与文本作为辅助校验
                        string cellC = Convert.ToString(infoSheet.Cells[r, 3].Value)?.Trim() ?? "";
                        string textC = string.Empty;
                        try { textC = Convert.ToString(infoSheet.Cells[r, 3].Text)?.Trim() ?? ""; } catch { }

                        // 遇到小计行说明汇总数据结束
                        if (cellB.Contains("小计") || textB.Contains("小计")) break;

                        // 判定是否为破坏性的 #REF! 错误或底表已删除导致的断链行
                        bool isRefError = cellB.Contains("#REF") || formulaB.Contains("#REF") || textB.Contains("#REF")
                            || cellC.Contains("#REF") || textC.Contains("#REF") || cellB.StartsWith("#") || cellB == "-2146826288" || cellC == "-2146826288";

                        // 若 B 列无值且无公式且无 #REF! 错误，说明到达未启用的预留空白行，跳过继续
                        if (!isRefError && string.IsNullOrWhiteSpace(cellB) && string.IsNullOrWhiteSpace(formulaB) && string.IsNullOrWhiteSpace(textB))
                        {
                            continue;
                        }

                        // 判断当前行是否落在用户选区中
                        bool isHit = selectedRowIndices.Contains(r);

                        // 校验底表是否真实存在
                        bool sheetExists = !string.IsNullOrWhiteSpace(cellB) && existingCategorySheetNames.Contains(cellB);

                        if (isRefError || !sheetExists)
                        {
                            // 发现已损坏或底表丢失的失效残留行
                            string displayName = isRefError ? $"#REF! (第 {r} 行)" : $"{cellB} (底表已丢失，第 {r} 行)";
                            var item = new DeleteCategoryItemDto
                            {
                                CategoryName = displayName,
                                CabinetCount = 0,
                                TotalPrice = 0.0,
                                CostPrice = 0.0,
                                // 如果是失效行或选区命中，默认推荐勾选以方便用户一键清理
                                IsSelectedInSheet = isHit || isRefError,
                                InfoRowIndex = r,
                                IsInvalid = true,
                                InvalidReason = isRefError ? "#REF! 引用失效" : "对应工作表已不存在"
                            };

                            detectedCategories[$"__INVALID_ROW_{r}"] = item;
                        }
                        else
                        {
                            // 正常的现存分类汇总行
                            int cabCount = 0;
                            int.TryParse(Convert.ToString(infoSheet.Cells[r, 3].Value)?.Trim(), out cabCount);

                            double totalPrice = 0.0;
                            double.TryParse(Convert.ToString(infoSheet.Cells[r, 4].Value)?.Trim(), out totalPrice);

                            double costPrice = 0.0;
                            double.TryParse(Convert.ToString(infoSheet.Cells[r, 5].Value)?.Trim(), out costPrice);

                            var item = new DeleteCategoryItemDto
                            {
                                CategoryName = cellB,
                                CabinetCount = cabCount,
                                TotalPrice = Math.Round(totalPrice, 2),
                                CostPrice = Math.Round(costPrice, 2),
                                IsSelectedInSheet = isHit,
                                InfoRowIndex = r,
                                IsInvalid = false,
                                InvalidReason = string.Empty
                            };

                            detectedCategories[cellB] = item;
                        }
                    }
                }

                // 2. 遍历补充未在【项目信息】表中登记的孤立分类工作表
                foreach (string wsName in existingCategorySheetNames)
                {
                    if (!detectedCategories.ContainsKey(wsName))
                    {
                        bool isCurrentActive = string.Equals(wsName, activeSheetName, StringComparison.OrdinalIgnoreCase);
                        detectedCategories[wsName] = new DeleteCategoryItemDto
                        {
                            CategoryName = wsName,
                            CabinetCount = 0,
                            TotalPrice = 0.0,
                            CostPrice = 0.0,
                            IsSelectedInSheet = isCurrentActive,
                            InfoRowIndex = 0,
                            IsInvalid = false,
                            InvalidReason = string.Empty
                        };
                    }
                }

                // 装载最终待删除分类列表响应模型
                response.Categories = detectedCategories.Values.ToList();
                response.TotalCategoryCount = existingCategorySheetNames.Count;
                response.PreSelectedCount = response.Categories.Count(c => c.IsSelectedInSheet);
                response.InvalidRowsCount = response.Categories.Count(c => c.IsInvalid);
            }
            catch (Exception ex)
            {
                // 记录数据提取异常日志
                LogHelper.WriteLog($"提取待删除分类数据异常: {ex.Message}");
            }

            return response;
        }

        /// <summary>
        /// 一次性批量删除一个或多个分类工作表（支持删除正常分类与清理 #REF! 失效汇总行）
        /// </summary>
        /// <param name="categoryNames">待删除的目标分类工作表名称集合</param>
        /// <param name="explicitRowIndices">可选显式传入待整行删除的项目信息表物理行号集合</param>
        /// <param name="explicitApp">可选显式传入的 Excel COM Application 实例</param>
        /// <returns>操作执行结果</returns>
        public static CategoryOperationResult DeleteCategories(List<string> categoryNames, List<int>? explicitRowIndices = null, dynamic? explicitApp = null)
        {
            // 整理待删除行号集合与待删除表名集合
            var targetRowIndices = new HashSet<int>(explicitRowIndices ?? new List<int>());
            categoryNames = categoryNames ?? new List<string>();

            // 清洗待删除分类名称 (去重、去空、排除项目信息等保留表与纯 #REF! 标记) --硬编码: 系统保留表名--
            var validCatNames = new List<string>();
            foreach (var n in categoryNames)
            {
                if (string.IsNullOrWhiteSpace(n)) continue;
                string trimmed = n.Trim();
                if (string.Equals(trimmed, "项目信息", StringComparison.OrdinalIgnoreCase)) continue;

                // 若项名为形如 "#REF! (第 30 行)"，尝试解析提取其物理行号
                if (trimmed.StartsWith("#REF!") || trimmed.StartsWith("__INVALID_ROW_"))
                {
                    int leftParen = trimmed.IndexOf("第 ");
                    int rightParen = trimmed.IndexOf(" 行");
                    if (leftParen >= 0 && rightParen > leftParen)
                    {
                        string rowStr = trimmed.Substring(leftParen + 2, rightParen - leftParen - 2).Trim();
                        if (int.TryParse(rowStr, out int rIdx) && rIdx > 0)
                        {
                            targetRowIndices.Add(rIdx);
                        }
                    }
                }
                else
                {
                    validCatNames.Add(trimmed);
                }
            }

            // 若既无有效表名也无行号，返回提示
            if (validCatNames.Count == 0 && targetRowIndices.Count == 0)
            {
                return new CategoryOperationResult { Success = false, Message = "未指定有效待删除的分类或待清理的失效行！" };
            }

            try
            {
                // 获取 Excel COM Application 实例
                dynamic? app = explicitApp ?? ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveWorkbook == null) return new CategoryOperationResult { Success = false, Message = "当前无活动工作簿" };

                dynamic activeWb = app.ActiveWorkbook;

                // 系统保留工作表黑名单
                var reservedSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "项目信息", "封面", "元件汇总表", "材料分布表",
                    "元件汇总分布表", "元件汇总调价清单", "屏柜汇总表", "屏柜分项表", "元器件数据管理"
                };

                // 统计工作簿中当前现存有效分类表总数
                int totalCatSheets = 0;
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    string name = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(name) && !reservedSheets.Contains(name))
                    {
                        totalCatSheets++;
                    }
                }

                // 校验项目内分类表保有量：若删除了真实分类表，删除后必须至少保留一个有效分类工作表
                if (validCatNames.Count > 0 && (totalCatSheets - validCatNames.Count < 1))
                {
                    return new CategoryOperationResult
                    {
                        Success = false,
                        Message = $"项目中至少需要保留一个分类工作表！当前共有 {totalCatSheets} 个分类，无法全部删除。"
                    };
                }

                // 挂起界面重绘、删除告警与系统事件以保障极速执行
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                try
                {
                    // 1. 批量物理删除目标分类工作表
                    foreach (var catName in validCatNames)
                    {
                        try
                        {
                            dynamic? ws = null;
                            try { ws = activeWb.Sheets[catName]; } catch { }
                            if (ws != null)
                            {
                                ws.Delete();
                            }
                        }
                        catch (Exception exWs)
                        {
                            LogHelper.WriteLog($"物理删除分类工作表【{catName}】异常: {exWs.Message}");
                        }

                        // 若暂存的复制源分类恰好被删除，重置暂存变量
                        if (string.Equals(_copiedCategorySheetName, catName, StringComparison.OrdinalIgnoreCase))
                        {
                            _copiedCategorySheetName = string.Empty;
                        }
                    }

                    // 2. 在【项目信息】表中批量倒序整行物理删除对应汇总行与指定失效行 (自下而上整行删除，行号不偏移)
                    dynamic? infoSheet = null;
                    try { infoSheet = activeWb.Sheets["项目信息"]; } catch { }
                    if (infoSheet != null)
                    {
                        var cfg = ConfigManager.Instance.Current.Excel;
                        int startRow = cfg.ProjectInfoCategorySummaryStartRow;
                        int maxScan = cfg.ProjectInfoCategorySummaryMaxScanRows;
                        var rowsToDelete = new HashSet<int>(targetRowIndices);

                        // 扫描查找 B 列匹配已删除分类名或显式 #REF! 的行号
                        for (int r = startRow; r < startRow + maxScan; r++)
                        {
                            string cellB = Convert.ToString(infoSheet.Cells[r, 2].Value)?.Trim() ?? "";
                            if (cellB.Contains("小计")) break;

                            // 命中待删除分类名称集合
                            if (validCatNames.Contains(cellB, StringComparer.OrdinalIgnoreCase))
                            {
                                rowsToDelete.Add(r);
                            }
                        }

                        // 从大到小倒序整行物理删除，保证序号公式自愈且不产生行号错位
                        foreach (int r in rowsToDelete.OrderByDescending(x => x))
                        {
                            try
                            {
                                infoSheet.Rows[r].EntireRow.Delete();
                            }
                            catch (Exception exRow)
                            {
                                LogHelper.WriteLog($"整行物理删除项目信息表第 {r} 行异常: {exRow.Message}");
                            }
                        }
                    }

                    // 3. 统一全量自愈并校准剩余分类汇总行正向与反向超链接
                    NormalizeCategorySummaryLinks(activeWb);

                    // 4. 激活【项目信息】工作表或首个剩余分类表
                    try
                    {
                        if (infoSheet != null)
                        {
                            infoSheet.Activate();
                        }
                        else
                        {
                            foreach (dynamic ws in activeWb.Worksheets)
                            {
                                string wName = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                                if (!reservedSheets.Contains(wName))
                                {
                                    ws.Activate();
                                    break;
                                }
                            }
                        }
                    }
                    catch { }

                    // 拼装反馈成功提示信息
                    string msg = validCatNames.Count > 0
                        ? $"已成功彻底删除 {validCatNames.Count} 个分类工作表及其在【项目信息】中的对应汇总！"
                        : $"已成功彻底清理【项目信息】表中的失效分类残留行！";

                    // 同步清空项目分类工作表白名单缓存
                    Tool.InvalidateProjectCategorySheetCache();

                    return new CategoryOperationResult
                    {
                        Success = true,
                        Message = msg
                    };
                }
                finally
                {
                    // 恢复界面重绘、告警与系统事件调度
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    app.EnableEvents = true;
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"批量删除分类业务执行异常: {ex.Message}");
                return new CategoryOperationResult { Success = false, Message = $"批量删除分类执行失败: {ex.Message}" };
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
                    // 获取克隆生成的新工作表 COM 实例
                    dynamic newSheet = activeWb.Sheets[activeWb.Sheets.Count];
                    // 重命名为用户指定的新分类名称
                    newSheet.Name = targetName;

                    // 2. 清洗跨工作簿外部公式链接 (将残留的源工作簿或外部路径洗成本地公式)
                    Tool.CleanRangeFormulas(newSheet.UsedRange);

                    // 3. 全量清洗并重定向超链接：将新表中原本残留指向源表名的超链接 SubAddress 替换为新分类表名
                    try
                    {
                        // 遍历新克隆表中的所有单元格超链接
                        foreach (dynamic hl in newSheet.Hyperlinks)
                        {
                            try
                            {
                                // 获取已有超链接的子地址
                                string subAddr = Convert.ToString(hl.SubAddress) ?? "";
                                // 若子地址指向源分类表，精准替换为新分类表
                                if (!string.IsNullOrEmpty(subAddr) && subAddr.Contains($"'{srcName}'!"))
                                {
                                    // 更新超链接目标子地址为新表
                                    hl.SubAddress = subAddr.Replace($"'{srcName}'!", $"'{targetName}'!");
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }

                    // 4. 在【项目信息】工作表中追加或更新分类汇总行 (先在主表注册新分类并挂载双向超链接)
                    UpdateProjectInfoCategorySummary(activeWb, targetName);

                    // 5. 强制刷新分类表白名单缓存，确保新工作表立即被工程纳管识别
                    Tool.GetProjectCategorySheetNames(activeWb, forceRefresh: true);

                    // 6. 严格遵守规则 8：调用全量自愈方法重置所有箱柜的 4 个定义名称与双向超链接 (规则 6 架构保障)
                    Tool.FixAndFillCabinetNamesForSheet(newSheet, forceRebuild: true);

                    // 7. 激活聚焦新工作表并默认定位至 A1
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

                // 选项 A: 为分类表 A5 挂载反向超链接，精准跳转回【项目信息】当前分类汇总行
                SetCategorySheetBackHyperlink(targetWb, categorySheetName, targetInfoRow);

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

                // 同步清空项目分类工作表白名单缓存
                Tool.InvalidateProjectCategorySheetCache();
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

                        // 选项 A: 同步更新重命名后的新分类表 A5 反向超链接指向当前汇总行 r
                        SetCategorySheetBackHyperlink(targetWb, newName, r);

                        // 同步清空项目分类工作表白名单缓存
                        Tool.InvalidateProjectCategorySheetCache();
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
                        // 重新全量校准剩余分类汇总行正反向超链接，避免删行引发物理行错位
                        NormalizeCategorySummaryLinks(targetWb);

                        // 同步清空项目分类工作表白名单缓存
                        Tool.InvalidateProjectCategorySheetCache();
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
                    // 读取 B 列单元格的值、文本与公式
                    string cellBVal = Convert.ToString(infoSheet.Cells[r, 2].Value)?.Trim() ?? "";
                    string cellBText = string.Empty;
                    try { cellBText = Convert.ToString(infoSheet.Cells[r, 2].Text)?.Trim() ?? ""; } catch { }
                    string cellBFormula = Convert.ToString(infoSheet.Cells[r, 2].Formula)?.Trim() ?? "";

                    // 读取 C 列单元格的文本辅助判断
                    string cellCText = string.Empty;
                    try { cellCText = Convert.ToString(infoSheet.Cells[r, 3].Text)?.Trim() ?? ""; } catch { }

                    // 若遇到小计行，说明分类汇总数据区域已结束
                    if (cellBVal.Contains("小计") || cellBText.Contains("小计")) break;

                    // 判定是否为破坏性的 #REF! 错误或断链
                    bool isRefError = cellBVal.Contains("#REF") || cellBFormula.Contains("#REF") || cellBText.Contains("#REF")
                        || cellCText.Contains("#REF") || cellBVal == "-2146826288" || cellBVal.StartsWith("#");

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

                        // 3. 若仍未匹配且并非 #REF! 错误行，按物理顺序匹配现存分类工作表
                        if (string.IsNullOrEmpty(matchedCatName) && !isRefError)
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
                        // 若该行包含破坏性的 #REF! 错误且工作簿中已无匹配分类表，直接整行物理删除该废弃断链行自愈
                        if (isRefError)
                        {
                            try
                            {
                                infoSheet.Rows[r].EntireRow.Delete();
                                r--;
                                maxScan--;
                            }
                            catch { }
                            continue;
                        }

                        if (string.IsNullOrEmpty(cellBVal) && string.IsNullOrEmpty(cellBText)) break;
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

                        // 7. 同步规范化分类表 A5 反向超链接，精准跳转回【项目信息】当前物理汇总行 (选项 A)
                        SetCategorySheetBackHyperlink(targetWb, matchedCatName, r);
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

        /// <summary>
        /// 选项 A: 为分类工作表的 A5 单元格 ("项目名称：") 挂载反向超链接，精准跳转回【项目信息】该分类对应的汇总行
        /// 具备就地更新与防空保护，杜绝重复创建导致的 COM 泄漏
        /// </summary>
        /// <param name="targetWb">目标工作簿 COM 实例</param>
        /// <param name="categorySheetName">分类工作表名称</param>
        /// <param name="targetInfoRow">【项目信息】中该分类对应的汇总行物理行号</param>
        public static void SetCategorySheetBackHyperlink(dynamic targetWb, string categorySheetName, int targetInfoRow)
        {
            // 校验工作簿与分类名称参数有效性
            if (targetWb == null || string.IsNullOrWhiteSpace(categorySheetName) || targetInfoRow <= 0) return;

            try
            {
                // 获取分类工作表 COM 实例
                dynamic catSheet = null;
                try { catSheet = targetWb.Sheets[categorySheetName]; } catch { }
                if (catSheet == null) return;

                // 提取 A5 单元格 (项目名称: 标签单元格)
                dynamic a5Cell = catSheet.Range["A5"];
                // 目标跳转子地址：指向【项目信息】中该分类对应的汇总行 A 列单元格
                string targetSubAddress = $"'项目信息'!A{targetInfoRow}";
                string screenTip = "点击返回【项目信息】汇总行"; // --硬编码: 屏幕提示文本--

                // 若 A5 单元格已存在超链接，就地更新 SubAddress 与 ScreenTip，保留已有样式与动态公式
                if (a5Cell.Hyperlinks != null && a5Cell.Hyperlinks.Count > 0)
                {
                    dynamic hl = a5Cell.Hyperlinks[1];
                    hl.SubAddress = targetSubAddress;
                    hl.ScreenTip = screenTip;
                    // 注意：绝不可赋值 hl.TextToDisplay，否则 Excel COM 会强制将单元格公式抹除为纯静态文本
                }
                else
                {
                    // 挂载新超链接（不传 TextToDisplay 参数，确保 100% 保护 A5 单元格原有的 CONCATENATE 动态公式）
                    catSheet.Hyperlinks.Add(
                        Anchor: a5Cell,
                        Address: "",
                        SubAddress: targetSubAddress,
                        ScreenTip: screenTip
                    );
                }
            }
            catch (Exception exLink)
            {
                // 记录反向超链接挂载异常日志
                LogHelper.WriteLog($"绑定分类表【{categorySheetName}】A5 反向超链接异常: {exLink.Message}");
            }
        }

        /// <summary>
        /// 供 Ribbon 菜单调用的“合并分类”弹窗入口 (基于 WebView2 + Vue 3 的非模态交互窗口)
        /// </summary>
        public static void ShowMergeCategoryDialog()
        {
            try
            {
                // 获取当前正在运行的 Excel Application 接口 (安全调用)
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                // 校验工作簿是否已打开
                if (app == null || app.ActiveWorkbook == null)
                {
                    // 提示用户必须先打开工作簿
                    System.Windows.Forms.MessageBox.Show(
                        "请先打开或新建一个报价项目工作簿！",
                        "系统提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // 统计工作簿中除系统主表外的分类表总数
                dynamic wb = app.ActiveWorkbook;
                int catCount = 0;
                // 系统保留表黑名单 --硬编码: 系统保留表名--
                var reservedSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "项目信息", "封面", "元件汇总表", "材料分布表",
                    "元件汇总分布表", "元件汇总调价清单", "屏柜汇总表", "屏柜分项表", "元器件数据管理"
                };

                // 遍历统计有效分类工作表
                foreach (dynamic ws in wb.Worksheets)
                {
                    string wsName = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(wsName) && !reservedSheets.Contains(wsName))
                    {
                        catCount++;
                    }
                }

                // 规则 9 守门：必须至少存在 2 个分类工作表才能进行合并
                if (catCount < 2)
                {
                    // 弹出拦截提示
                    System.Windows.Forms.MessageBox.Show(
                        "当前项目至少需要存在 2 个有效的分类工作表才能进行合并！",
                        "系统提示",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // 以统一非模态方式展示合并分类窗口
                ShowModelessForm(ref _mergeCategoryFormInstance, () => new Forms.MergeCategoryForm());
            }
            catch (Exception ex)
            {
                // 记录打开合并分类窗口异常日志
                LogHelper.WriteLog($"弹出合并分类窗口异常: {ex.Message}");
                // 弹出异常提示
                System.Windows.Forms.MessageBox.Show(
                    $"弹出合并分类窗口失败: {ex.Message}",
                    "系统提示",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 获取工作簿中所有有效分类的简要统计数据与推荐选项 (供前端合并界面渲染)
        /// </summary>
        /// <returns>合并分类初始化响应模型</returns>
        public static MergeCategoryInitResponse GetMergeCategoriesData()
        {
            // 初始化返回响应对象
            var response = new MergeCategoryInitResponse();

            try
            {
                // 获取 Excel COM Application 接口实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveWorkbook == null) return response;

                dynamic wb = app.ActiveWorkbook;
                dynamic? activeSheet = app.ActiveSheet;
                string activeSheetName = Convert.ToString(activeSheet?.Name)?.Trim() ?? string.Empty;

                // 系统保留工作表黑名单 --硬编码: 系统保留表名--
                var reservedSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "项目信息", "封面", "元件汇总表", "材料分布表",
                    "元件汇总分布表", "元件汇总调价清单", "屏柜汇总表", "屏柜分项表", "元器件数据管理"
                };

                // 遍历收集有效分类工作表
                foreach (dynamic ws in wb.Worksheets)
                {
                    string wsName = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                    // 跳过空白表名与系统保留表
                    if (string.IsNullOrWhiteSpace(wsName) || reservedSheets.Contains(wsName)) continue;

                    // 实例化单分类摘要 DTO
                    var item = new CategoryBriefSummaryDto
                    {
                        CategoryName = wsName
                    };

                    try
                    {
                        // 规则 8: 在操作/读取分类数据前，显式调用自愈函数纠正定义名称完整性
                        try { Tool.FixAndFillCabinetNamesForSheet(ws); } catch { }

                        // 获取工作表内所有有效箱柜分布集合
                        var validCabinets = Tool.GetSheetValidCabinets((object)ws, (object)wb);

                        // 规则 6 架构标准：直接从已识别的有效箱柜映射中提取物理汇总行号 (Cab_Sum_k 锚点)
                        // 彻底杜绝 minDetRow 猜测与 500 行硬编码截断，实现 100% 精确与箱柜一一对应
                        var orderedSumRows = validCabinets
                            .Select(c => c.Value.SumRow > 0 ? c.Value.SumRow : (c.Value.Sum != null ? Convert.ToInt32(c.Value.Sum.Row) : 0))
                            .Where(r => r >= 6)
                            .Distinct()
                            .OrderBy(r => r)
                            .ToList();

                        if (orderedSumRows.Count > 0)
                        {
                            // 确定顶部汇总物理行区间 (最小行至最大行)
                            int minSumRow = orderedSumRows.First();
                            int maxSumRow = orderedSumRows.Last();

                            // 规则 7: 根据实际箱柜汇总区间，一次性将 A 列至 H 列加载入内存二维数组 (0 次单格 COM 往返)
                            dynamic sumRange = ws.Range[$"A{minSumRow}:H{maxSumRow}"];
                            object[,] vals = (object[,])sumRange.Value2;

                            // 遍历每个箱柜的精准物理汇总行提取信息
                            foreach (int sRow in orderedSumRows)
                            {
                                int rIdx = sRow - minSumRow + 1;

                                // 提取 B 列(柜号) 与 C 列(设备名称)
                                string cabNo = Convert.ToString(vals[rIdx, 2])?.Trim() ?? "";
                                string cabName = Convert.ToString(vals[rIdx, 3])?.Trim() ?? "";

                                // 累加箱柜种类数 (每行代表一种箱柜型号)
                                item.CabinetTypeCount++;

                                // 提取 F 列台数数量 (若为空或非法数值默认按 1 台计)
                                double qty = 1.0;
                                try
                                {
                                    if (vals[rIdx, 6] != null)
                                    {
                                        double parsedQty = Convert.ToDouble(vals[rIdx, 6]);
                                        if (parsedQty > 0) qty = parsedQty;
                                    }
                                }
                                catch { }
                                // 累加设备总台数
                                item.TotalQuantity += qty;

                                // 提取柜号并追加至前端预览标签展示列表 (最多展示前 10 个)
                                string displayName = !string.IsNullOrWhiteSpace(cabNo) ? cabNo : cabName;
                                if (item.CabinetNames.Count < 10 && !string.IsNullOrWhiteSpace(displayName))
                                {
                                    item.CabinetNames.Add(displayName);
                                }

                                // 累计 H 列销售总价金额
                                try
                                {
                                    if (vals[rIdx, 8] != null) item.TotalPrice += Convert.ToDouble(vals[rIdx, 8]);
                                }
                                catch { }
                            }
                        }
                        else
                        {
                            // 兜底情况：若工作表尚未生成定义名称，从第 7 行开始向下扫描实质内容
                            dynamic used = ws.UsedRange;
                            int totalRows = used.Rows.Count;
                            if (totalRows >= 7)
                            {
                                dynamic scanRange = ws.Range[$"A7:H{Math.Min(totalRows, 300)}"];
                                object[,] vals = (object[,])scanRange.Value2;
                                int rowCount = vals.GetLength(0);
                                for (int r = 1; r <= rowCount; r++)
                                {
                                    string b = Convert.ToString(vals[r, 2])?.Trim() ?? "";
                                    string c = Convert.ToString(vals[r, 3])?.Trim() ?? "";
                                    if (string.IsNullOrWhiteSpace(b) && string.IsNullOrWhiteSpace(c)) continue;
                                    if (b.Contains("小计") || b.Contains("总计") || b.Contains("合计")) break;

                                    item.CabinetTypeCount++;
                                    double qty = 1.0;
                                    try { if (vals[r, 6] != null) qty = Convert.ToDouble(vals[r, 6]); } catch { }
                                    item.TotalQuantity += (qty > 0 ? qty : 1.0);
                                    if (item.CabinetNames.Count < 10) item.CabinetNames.Add(!string.IsNullOrWhiteSpace(b) ? b : c);
                                    try { if (vals[r, 8] != null) item.TotalPrice += Convert.ToDouble(vals[r, 8]); } catch { }
                                }
                            }
                        }

                        // 赋值兼顾兼容字段：若已统计出总台数则使用总台数，否则回退种类数
                        item.CabinetCount = item.TotalQuantity > 0 ? (int)Math.Round(item.TotalQuantity) : item.CabinetTypeCount;
                    }
                    catch (Exception exScan)
                    {
                        // 记录扫描单个分类异常容错日志
                        LogHelper.WriteLog($"扫描分类【{wsName}】摘要异常: {exScan.Message}");
                    }

                    // 加入分类列表集合
                    response.Categories.Add(item);
                }

                // 智能推导默认源分类与目标分类
                if (response.Categories.Count >= 2)
                {
                    // 若当前激活工作表在候选列表中，优先将其作为源分类
                    var matchedActive = response.Categories.FirstOrDefault(c => string.Equals(c.CategoryName, activeSheetName, StringComparison.OrdinalIgnoreCase));
                    if (matchedActive != null)
                    {
                        response.DefaultSourceCategory = matchedActive.CategoryName;
                        // 寻找除源分类外的第一个分类作为目标分类
                        var other = response.Categories.FirstOrDefault(c => !string.Equals(c.CategoryName, matchedActive.CategoryName, StringComparison.OrdinalIgnoreCase));
                        response.DefaultTargetCategory = other?.CategoryName ?? string.Empty;
                    }
                    else
                    {
                        // 默认取第 1 个作为源分类，第 2 个作为目标分类
                        response.DefaultSourceCategory = response.Categories[0].CategoryName;
                        response.DefaultTargetCategory = response.Categories[1].CategoryName;
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录数据提取异常日志
                LogHelper.WriteLog($"获取合并分类数据异常: {ex.Message}");
            }

            return response;
        }

        /// <summary>
        /// 核心业务引擎：执行合并分类 (将源分类工作表的所有箱柜明细追加合并到目标分类工作表末尾，并联动更新项目信息)
        /// </summary>
        /// <param name="request">合并分类请求参数</param>
        /// <param name="explicitApp">显式传递的 Excel COM Application 句柄 (可选)</param>
        /// <param name="progressCallback">进度更新委托回调 (百分比 0-100, 描述信息)</param>
        /// <returns>合并操作执行结果</returns>
        public static CategoryOperationResult MergeCategories(
            MergeCategoryRequest request,
            dynamic? explicitApp = null,
            Action<int, string>? progressCallback = null)
        {
            // 校验请求对象有效性
            if (request == null)
            {
                return new CategoryOperationResult { Success = false, Message = "合并分类请求参数不能为空！" };
            }

            string srcName = request.SourceCategoryName?.Trim() ?? string.Empty;
            string targetName = request.TargetCategoryName?.Trim() ?? string.Empty;

            // 规则 9 守门：参数非空与非同表校验
            if (string.IsNullOrWhiteSpace(srcName) || string.IsNullOrWhiteSpace(targetName))
            {
                return new CategoryOperationResult { Success = false, Message = "源分类与目标分类名称不能为空！" };
            }
            if (string.Equals(srcName, targetName, StringComparison.OrdinalIgnoreCase))
            {
                return new CategoryOperationResult { Success = false, Message = "源分类与目标分类不能是同一个工作表！" };
            }

            try
            {
                // 获取 Excel COM Application 接口实例
                dynamic? app = explicitApp ?? ExcelDnaSafeAccessor.GetApplication();
                if (app == null || app.ActiveWorkbook == null)
                {
                    return new CategoryOperationResult { Success = false, Message = "当前没有正在运行的项目工作簿！" };
                }

                dynamic wb = app.ActiveWorkbook;

                // 定位源分类工作表与目标分类工作表
                dynamic? srcSheet = null;
                dynamic? targetSheet = null;
                try { srcSheet = wb.Worksheets[srcName]; } catch { }
                try { targetSheet = wb.Worksheets[targetName]; } catch { }

                // 校验工作表真实存在性
                if (srcSheet == null)
                {
                    return new CategoryOperationResult { Success = false, Message = $"未找到源分类工作表【{srcName}】！" };
                }
                if (targetSheet == null)
                {
                    return new CategoryOperationResult { Success = false, Message = $"未找到目标分类工作表【{targetName}】！" };
                }

                // 进度反馈：5% 开始预检
                progressCallback?.Invoke(5, "正在校验分类工作表并执行规则自愈...");

                // 1. 规则 8 守门：强制预检并自愈两表的定义名称与规则 6 架构完整性
                Tool.FixAndFillCabinetNamesForSheet(srcSheet, forceRebuild: false);
                Tool.FixAndFillCabinetNamesForSheet(targetSheet, forceRebuild: false);
                // 清洗两表中的孤立幽灵定义名称
                Tool.PurgeGhostNamesForSheet(srcSheet);
                Tool.PurgeGhostNamesForSheet(targetSheet);
                // 刷新内存缓存
                Tool.InvalidateSheetCabinetsCache(srcName);
                Tool.InvalidateSheetCabinetsCache(targetName);

                // 2. 获取源分类表中的全部有效箱柜
                var srcCabinets = Tool.GetSheetValidCabinets((object)srcSheet, (object)wb);
                // 若源分类表完全为空 (无任何有效箱柜)
                if (srcCabinets == null || srcCabinets.Count == 0)
                {
                    // 若用户选择合并后删除源表，则直接清理源表并返回
                    if (request.DeleteSourceCategory)
                    {
                        // 移除项目信息汇总行
                        RemoveProjectInfoCategorySummary(wb, srcName);
                        // 删除工作表
                        srcSheet.Delete();
                        // 规整超链接
                        NormalizeCategorySummaryLinks(wb);
                        // 刷新白名单缓存
                        Tool.InvalidateProjectCategorySheetCache();
                    }
                    progressCallback?.Invoke(100, "合并完成！");
                    return new CategoryOperationResult
                    {
                        Success = true,
                        Message = $"源分类【{srcName}】无任何箱柜明细，已完成合并处理！",
                        CategoryName = targetName
                    };
                }

                // 进度反馈：10% 读取数据
                progressCallback?.Invoke(10, $"正在读取源分类【{srcName}】全部箱柜数据快照...");

                // 3. 严格按源表物理汇总行顺序对箱柜进行物理行去重并排序 (自上而下保证顺序对齐且彻底杜绝重复合并)
                var orderedSrcCabinets = srcCabinets
                    .GroupBy(c => c.Value.SumRow > 0 ? c.Value.SumRow : (c.Value.Sum != null ? Convert.ToInt32(c.Value.Sum.Row) : 0))
                    .Where(g => g.Key > 0)
                    .Select(g => g.First())
                    .OrderBy(c => c.Value.SumRow > 0 ? c.Value.SumRow : (c.Value.Sum != null ? Convert.ToInt32(c.Value.Sum.Row) : 0))
                    .ToList();

                // 4. 规则 7: 内存批量抓取源表所有箱柜的数据快照 (二维数组装载)
                var clips = new List<Models.CopiedCabinetContext>();
                foreach (var pair in orderedSrcCabinets)
                {
                    // 抓取单个箱柜的整柜上下文模型快照
                    var clip = CaptureCabinetContext(app, srcSheet, wb, specificK: pair.Key, isCut: false);
                    if (clip != null)
                    {
                        clips.Add(clip);
                    }
                }

                // 校验提取到的有效箱柜快照数量
                if (clips.Count == 0)
                {
                    return new CategoryOperationResult { Success = false, Message = $"从源分类【{srcName}】中未提取到有效的箱柜数据！" };
                }

                // 5. 暂停 Excel 视口渲染与系统告警，开启批量加速
                bool prevUpdating = false;
                bool prevAlerts = false;
                bool prevEvents = false;
                try
                {
                    prevUpdating = Convert.ToBoolean(app.ScreenUpdating);
                    prevAlerts = Convert.ToBoolean(app.DisplayAlerts);
                    prevEvents = Convert.ToBoolean(app.EnableEvents);
                }
                catch { }

                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                int mergedCount = 0;
                int totalClips = clips.Count;
                try
                {
                    // 6. 批量追加基线初始化 (循环外部只计算一次，消除 O(N^2) 全表扫描瓶颈)
                    progressCallback?.Invoke(12, "正在准备目标分类表批量追加基准...");

                    // 提取定义名称前缀配置
                    var cfg = ConfigManager.Instance.Current?.Excel;
                    var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = cfg?.Prefixes ?? CabinetPrefixConfig.Default;

                    // 获取目标表初始有效箱柜列表
                    var targetValidCabinets = Tool.GetSheetValidCabinets((object)targetSheet, (object)wb);

                    // 确定首个汇总行的插入物理行号 (末尾追加模式)
                    int curInsertSumRow = DetermineTopSummaryInsertRow(app, targetSheet, targetValidCabinets, 0, appendToEnd: true);

                    // 确定首个明细块的插入物理行号 (末尾追加模式)
                    int curInsertDetRow = 41; // 默认兜底基准
                    var targetIndexes = Tool.FindStandardCategoryRowIndexes((object)targetSheet, -1);
                    if (targetIndexes.cabTolsumRow > 0 && targetIndexes.cabDetRow > 0)
                    {
                        // 紧贴现有最后一个明细块总计行后 + 4 行追加大标题
                        curInsertDetRow = targetIndexes.cabTolsumRow + 4;
                    }
                    else
                    {
                        // 目标表当前尚无明细块，取默认基准行 38
                        int defDetRow = ConfigManager.Instance.Current?.Excel?.CabDetRowIndex ?? 41;
                        curInsertDetRow = Math.Max(38, defDetRow - 3);
                    }

                    // 提取目标表当前最大分配唯一序号 startK
                    int startK = GetNextCabinetIndex(wb, targetSheet) - 1;

                    // 7. 核心极速循环：依次批量追加每个箱柜 (0次重复全表扫描，0次重复自愈)
                    for (int i = 0; i < totalClips; i++)
                    {
                        var clip = clips[i];
                        int newK = ++startK;

                        // 动态计算进度百分比 (15% ~ 80%)
                        int progressVal = 15 + (int)Math.Round((double)i / totalClips * 65);
                        progressCallback?.Invoke(progressVal, $"正在合并箱柜 ({i + 1}/{totalClips}): {clip.CabinetNo}...");

                        // --- A. 顶部汇总行追加 ---
                        // 在当前汇总追加指针处插入 1 行
                        targetSheet.Rows[$"{curInsertSumRow}:{curInsertSumRow}"].Insert(-4121);
                        dynamic dstSumRange = targetSheet.Rows[$"{curInsertSumRow}:{curInsertSumRow}"];

                        // 从源表汇总行复制整行格式与样式
                        try
                        {
                            if (clip.SourceSumRow > 0)
                            {
                                dynamic srcSumRange = srcSheet.Rows[$"{clip.SourceSumRow}:{clip.SourceSumRow}"];
                                srcSumRange.Copy(dstSumRange);
                                Tool.CleanRangeFormulas(dstSumRange, clip.SourceSheetName, clip.SourceCabinetK, newK);
                            }
                        }
                        catch (Exception exSumCopy)
                        {
                            LogHelper.WriteLog($"[MergeCategories] 复制汇总行格式异常: {exSumCopy.Message}");
                        }

                        // 绑定新汇总定义名称 Cab_Sum_{newK}
                        string newSumTag = $"{sumPrefix}{newK}";
                        dynamic sumAnchorCell = targetSheet.Cells[curInsertSumRow, 1];
                        Tool.SafeSetSheetName(targetSheet, targetName, newSumTag, curInsertSumRow);

                        // 填充箱柜基本属性与自适应序号公式
                        sumAnchorCell.Formula = "=ROW()-ROW(A$6)"; // --硬编码: 序号公式--
                        targetSheet.Cells[curInsertSumRow, 2].Value = clip.CabinetNo;
                        targetSheet.Cells[curInsertSumRow, 3].Value = clip.Name;
                        targetSheet.Cells[curInsertSumRow, 4].Value = clip.Model;
                        targetSheet.Cells[curInsertSumRow, 5].Value = clip.Unit;
                        targetSheet.Cells[curInsertSumRow, 6].Value = clip.Quantity;

                        // 推进汇总行指针
                        curInsertSumRow++;
                        // 关键物理联动：顶部汇总区每插入 1 行，底表明细区在物理上自动被推下 1 行
                        curInsertDetRow++;

                        // --- B. 底表明细块追加 ---
                        if (clip.Kind == Models.CabinetKind.NoDetail || clip.DetailBlockRowCount <= 0)
                        {
                            // 无明细箱柜：回写单价与成本公式
                            int currentCabSumRow = curInsertSumRow - 1;
                            object rawUnitPrice = (clip.SumRowValues != null && clip.SumRowValues.GetLength(1) >= 7) ? clip.SumRowValues[1, 7] : 0;
                            object rawCostPrice = (clip.SumRowValues != null && clip.SumRowValues.GetLength(1) >= 10) ? clip.SumRowValues[1, 10] : 0;

                            targetSheet.Cells[currentCabSumRow, 7].Value = rawUnitPrice;
                            targetSheet.Cells[currentCabSumRow, 8].Formula = $"=F{currentCabSumRow}*G{currentCabSumRow}";
                            targetSheet.Cells[currentCabSumRow, 10].Value = rawCostPrice;
                            targetSheet.Cells[currentCabSumRow, 11].Formula = $"=F{currentCabSumRow}*J{currentCabSumRow}";
                            targetSheet.Cells[currentCabSumRow, 12].Formula = $"=H{currentCabSumRow}-K{currentCabSumRow}";
                            targetSheet.Cells[currentCabSumRow, 13].Formula = $"=IF(H{currentCabSumRow}=0,0,L{currentCabSumRow}/H{currentCabSumRow})";
                        }
                        else
                        {
                            // 普通有明细箱柜：整块复制明细块
                            int blockRowCount = clip.DetailBlockRowCount;
                            int srcStart = clip.SourceDetStartRow;
                            int srcEnd = clip.SourceDetEndRow;

                            // 物理插入明细块空白行
                            targetSheet.Rows[$"{curInsertDetRow}:{curInsertDetRow + blockRowCount - 1}"].Insert(-4121);
                            dynamic dstDetailRange = targetSheet.Rows[$"{curInsertDetRow}:{curInsertDetRow + blockRowCount - 1}"];

                            // 原生 Range.Copy 完整复刻边框、底色、合并单元格、公式与 CAD 句柄
                            bool copyDetailOk = false;
                            try
                            {
                                if (srcStart > 0 && srcEnd >= srcStart)
                                {
                                    dynamic srcDetailRange = srcSheet.Rows[$"{srcStart}:{srcEnd}"];
                                    srcDetailRange.Copy(dstDetailRange);
                                    copyDetailOk = true;
                                }
                            }
                            catch (Exception exDetCopy)
                            {
                                LogHelper.WriteLog($"[MergeCategories] Range.Copy 明细块异常: {exDetCopy.Message}");
                            }

                            // 降级兜底数组写入
                            if (!copyDetailOk && clip.DetailBlockValues != null)
                            {
                                dynamic detRange = targetSheet.Range[$"A{curInsertDetRow}:Q{curInsertDetRow + blockRowCount - 1}"];
                                detRange.Value2 = clip.DetailBlockValues;
                                detRange.Formula = clip.DetailBlockFormulas;
                                Tool.CleanRangeFormulas(detRange, clip.SourceSheetName, clip.SourceCabinetK, newK);
                            }
                            else
                            {
                                Tool.CleanRangeFormulas(dstDetailRange, clip.SourceSheetName, clip.SourceCabinetK, newK);
                            }

                            // 计算新明细块关键行号
                            int newDetRow = curInsertDetRow + (clip.SourceDetRow - srcStart);
                            int newSubsumRow = curInsertDetRow + (clip.SourceSubsumRow - srcStart);
                            int newTolsumRow = curInsertDetRow + (clip.SourceTolsumRow - srcStart);

                            string newDetTag = $"{detPrefix}{newK}";
                            string newSubsumTag = $"{subsumPrefix}{newK}";
                            string newTolsumTag = $"{tolsumPrefix}{newK}";

                            // 注册目标表 3 个明细定义名称 (规则 6 规范)
                            Tool.SafeSetSheetName(targetSheet, targetName, newDetTag, newDetRow);
                            Tool.SafeSetSheetName(targetSheet, targetName, newSubsumTag, newSubsumRow);
                            Tool.SafeSetSheetName(targetSheet, targetName, newTolsumTag, newTolsumRow);

                            // 重建双向跳转超链接
                            try
                            {
                                int currentCabSumRow = curInsertSumRow - 1;
                                dynamic currentSumAnchor = targetSheet.Cells[currentCabSumRow, 1];
                                targetSheet.Hyperlinks.Add(
                                    Anchor: currentSumAnchor,
                                    Address: "",
                                    SubAddress: $"'{targetName}'!{newDetTag}",
                                    ScreenTip: "点击进入本箱柜明细表" // --硬编码: 屏幕提示文本--
                                );
                                currentSumAnchor.Formula = "=ROW()-ROW(A$6)"; // --硬编码: 序号公式--

                                targetSheet.Hyperlinks.Add(
                                    Anchor: targetSheet.Cells[newDetRow, 1],
                                    Address: "",
                                    SubAddress: $"'{targetName}'!{newSumTag}",
                                    ScreenTip: "返回汇总行" // --硬编码: 屏幕提示文本--
                                );
                            }
                            catch { }

                            // 更新明细表头柜号与序号公式
                            targetSheet.Cells[newDetRow, 2].Value = clip.CabinetNo;
                            try
                            {
                                targetSheet.Cells[newDetRow + 1, 1].Formula = $"=\"序号\" & {newSumTag}"; // --硬编码: 明细表头序号公式--
                            }
                            catch { }

                            // 推进底表明细追加指针
                            curInsertDetRow += blockRowCount;
                        }

                        mergedCount++;
                    }

                    // 进度反馈：85% 统一全局自愈
                    progressCallback?.Invoke(85, "正在规整目标分类结构并清理幽灵定义名称...");

                    // 8. 统一全量自愈与幽灵定义名称净化 (在批量合并完成后只执行一次)
                    Tool.FixAndFillCabinetNamesForSheet(targetSheet, forceRebuild: false);
                    Tool.PurgeGhostNamesForSheet(targetSheet);
                    Tool.InvalidateSheetCabinetsCache(targetName);

                    // 进度反馈：90% 更新项目信息
                    progressCallback?.Invoke(90, "正在联动刷新【项目信息】设备汇总与超链接...");

                    // 8. 联动更新【项目信息】主表中关于目标分类的设备数量、总价与成本
                    UpdateProjectInfoCategorySummary(wb, targetName);

                    // 9. 处理源分类工作表生命周期 (根据用户勾选选项)
                    if (request.DeleteSourceCategory)
                    {
                        progressCallback?.Invoke(95, $"正在安全移除源分类【{srcName}】工作表...");

                        // 从【项目信息】表中物理删除源分类汇总行 (整行删除并规整超链接)
                        RemoveProjectInfoCategorySummary(wb, srcName);

                        // 物理静默删除源工作表
                        srcSheet.Delete();

                        // 若暂存的复制分类恰好是源表，清空暂存
                        if (string.Equals(_copiedCategorySheetName, srcName, StringComparison.OrdinalIgnoreCase))
                        {
                            _copiedCategorySheetName = string.Empty;
                        }

                        // 规整【项目信息】表中所有分类的超链接
                        NormalizeCategorySummaryLinks(wb);

                        // 清空项目分类工作表白名单缓存
                        Tool.InvalidateProjectCategorySheetCache();
                    }

                    // 进度反馈：100% 达成
                    progressCallback?.Invoke(100, "合并操作完成！");
                }
                finally
                {
                    // 10. 恢复 Excel 视口渲染与事件系统
                    try
                    {
                        app.ScreenUpdating = prevUpdating;
                        app.DisplayAlerts = prevAlerts;
                        app.EnableEvents = prevEvents;
                    }
                    catch { }
                }

                // 11. 激活目标分类工作表，提供平滑的视觉反馈
                try
                {
                    targetSheet.Activate();
                }
                catch { }

                // 返回成功执行结果
                return new CategoryOperationResult
                {
                    Success = true,
                    Message = $"已成功将分类【{srcName}】的 {mergedCount} 台箱柜明细合并至【{targetName}】！",
                    CategoryName = targetName
                };
            }
            catch (Exception ex)
            {
                // 记录合并异常日志
                LogHelper.WriteLog($"[MergeCategories] 异常: {ex.Message}");
                return new CategoryOperationResult
                {
                    Success = false,
                    Message = $"合并分类失败: {ex.Message}"
                };
            }
        }
    }
}
