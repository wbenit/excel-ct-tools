using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;
using ExcelDna.Integration;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：二次回路方案与 BOM 表的 Excel 工作表智能解析导入
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，操作均在服务层完成，数组一次性读入内存
    /// </summary>
    public static partial class ExcelServices
    {
        // 声明二次方案管理窗口静态单例引用 (可空)
        private static SecondaryCircuitForm? _secondaryCircuitForm;

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“二次图回路方案管理中心”窗口 (非模态，可交互编辑 Excel)
        /// </summary>
        public static void ShowSecondaryCircuitManageDialog()
        {
            try
            {
                // 若窗体已打开且未销毁，直接还原并激活展示
                if (_secondaryCircuitForm != null && !_secondaryCircuitForm.IsDisposed)
                {
                    // 若处于最小化则还原正常大小
                    if (_secondaryCircuitForm.WindowState == FormWindowState.Minimized)
                    {
                        _secondaryCircuitForm.WindowState = FormWindowState.Normal;
                    }
                    // 推至顶层
                    _secondaryCircuitForm.BringToFront();
                    // 激活窗口焦点
                    _secondaryCircuitForm.Activate();
                    return;
                }

                // 实例化全新窗体
                _secondaryCircuitForm = new SecondaryCircuitForm();
                // 绑定关闭事件清空单例
                _secondaryCircuitForm.FormClosed += (s, e) => _secondaryCircuitForm = null;

                // 获取 Excel 主窗口 HWND 句柄以依附弹出
                IntPtr excelHwnd = ExcelDnaSafeAccessor.GetWindowHandle();
                if (excelHwnd != IntPtr.Zero)
                {
                    // 设置 Owner 为 Excel 主窗口非模态显示
                    _secondaryCircuitForm.Show(new ExcelWin32Window(excelHwnd));
                }
                else
                {
                    // 独立非模态弹出
                    _secondaryCircuitForm.Show();
                }
            }
            catch (Exception ex)
            {
                // 捕获异常提示用户
                MessageBox.Show($"打开二次方案管理中心窗口失败: {ex.Message}", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 从当前 Excel 活动工作表中一键批量识别解析二次图 BOM 表及方案定额数据并入库
        /// 遵循规则 7：二维数组一次性读入内存，不逐格 COM 交互
        /// </summary>
        /// <returns>导入执行结果与统计消息</returns>
        public static (bool Success, string Message, int SchemeCount, int BomCount) ImportSecondarySchemesFromActiveSheet()
        {
            try
            {
                // 获取当前活动 Excel Application 实例
                dynamic app = ExcelDnaUtil.Application;
                // 获取当前活动工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null)
                {
                    return (false, "当前未检测到打开的 Excel 工作簿！", 0, 0);
                }

                // 获取当前活动工作表
                dynamic sheet = activeWb.ActiveSheet;
                if (sheet == null)
                {
                    return (false, "当前未检测到活动工作表！", 0, 0);
                }

                // 获取工作表已使用数据区域 UsedRange
                dynamic usedRange = sheet.UsedRange;
                if (usedRange == null)
                {
                    return (false, "当前工作表为空，无可用数据！", 0, 0);
                }

                // 读取总行数与总列数
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;
                // 校验基础维度 (至少需要 2 行且至少 6 列)
                if (rowCount < 2 || colCount < 6)
                {
                    return (false, "当前表格数据区域过小，无法识别为二次图 BOM 表 (至少需要包含名称、型号、数量、跨门等列)！", 0, 0);
                }

                // 规则 7：采用二维数组一次性将所有单元格数据读入内存
                object[,] valMatrix = usedRange.Value2 as object[,];
                if (valMatrix == null)
                {
                    return (false, "读取工作表内存数据失败！", 0, 0);
                }

                // 解析出的方案集合
                var schemes = new List<SecondarySchemeEntity>();
                // 当前正处理的方案对象引用
                SecondarySchemeEntity? currentScheme = null;
                // 子 BOM 物料总计数
                int totalBomCount = 0;

                // 从第 2 行开始逐行向下扫描 (第 1 行为表头)
                for (int r = 2; r <= rowCount; r++)
                {
                    // 提取 B 列: 名称
                    string colB = GetSafeString(valMatrix, r, 2);
                    // 提取 C 列: 型号 / 方案名
                    string colC = GetSafeString(valMatrix, r, 3);
                    // 提取 D 列: 数量
                    string colD = GetSafeString(valMatrix, r, 4);
                    // 提取 E 列: 单位
                    string colE = GetSafeString(valMatrix, r, 5);
                    // 提取 F 列: 二次线跨门 / 单价
                    string colF = GetSafeString(valMatrix, r, 6);
                    // 提取 G 列: 二次材料费小计
                    string colG = GetSafeString(valMatrix, r, 7);
                    // 提取 H 列: 开孔
                    string colH = colCount >= 8 ? GetSafeString(valMatrix, r, 8) : string.Empty;
                    // 提取 I 列: 人工
                    string colI = colCount >= 9 ? GetSafeString(valMatrix, r, 9) : string.Empty;
                    // 提取 J 列: 二次组
                    string colJ = colCount >= 10 ? GetSafeString(valMatrix, r, 10) : string.Empty;
                    // 提取 K 列: CAD 图名 (若存在)
                    string colK = colCount >= 11 ? GetSafeString(valMatrix, r, 11) : string.Empty;

                    // 若整行关键列全部为空，则直接跳过
                    if (string.IsNullOrWhiteSpace(colB) && string.IsNullOrWhiteSpace(colC) && string.IsNullOrWhiteSpace(colF))
                    {
                        continue;
                    }

                    // 1. 判断是否为【主方案行】：
                    // 特征：B 列名称为空，且 C 列非空（方案型号代号），或 H/I 列（开孔/人工）有值
                    bool isSchemeRow = string.IsNullOrWhiteSpace(colB) && !string.IsNullOrWhiteSpace(colC);
                    // 辅助判断：若 H 列有开孔（如“圆2”）或 I 列有人工（如 35），即使 B 列有少量标记也是主方案行
                    if (!isSchemeRow && (!string.IsNullOrWhiteSpace(colH) || !string.IsNullOrWhiteSpace(colI)))
                    {
                        // 若 D 列数量为空，判定为主方案行
                        if (string.IsNullOrWhiteSpace(colD) || !double.TryParse(colD, out _))
                        {
                            isSchemeRow = true;
                        }
                    }

                    if (isSchemeRow)
                    {
                        // 实例化全新二次方案实体
                        currentScheme = new SecondarySchemeEntity();

                        // 解析 C 列同配置适用回路代号 (如: "双电源1, CA1B")
                        var codes = colC.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(s => s.Trim())
                                        .Where(s => !string.IsNullOrWhiteSpace(s))
                                        .ToList();

                        // 方案主名称取第一个代号或合并名称
                        currentScheme.SchemeName = codes.Count > 0 ? codes[0] : colC.Trim();
                        // 适用回路代号列表
                        currentScheme.ApplicableCodes = codes;

                        // 解析 F 列: 方案级二次线跨门根数 (四舍五入为整数根数)
                        if (double.TryParse(colF, out double crossDoor))
                        {
                            // 赋值整数跨门根数
                            currentScheme.CrossDoorCount = Math.Round(crossDoor);
                        }

                        // 解析 H 列: 开孔规范 (如: "圆2")
                        currentScheme.HoleSpec = colH.Trim();

                        // 解析 I 列: 人工工费 (如: 35)
                        if (double.TryParse(colI, out double labor))
                        {
                            currentScheme.LaborCost = labor;
                        }

                        // 解析 J 列: 二次排布图 (若为空则自动默认赋予通用排布图)
                        currentScheme.GroupName = string.IsNullOrWhiteSpace(colJ) ? "通用排布图" : colJ.Trim(); // --硬编码-- 默认排布图名称

                        // 解析 K 列: CAD 图名 (若有)
                        if (!string.IsNullOrWhiteSpace(colK))
                        {
                            currentScheme.CadDrawingName = colK.Trim();
                        }

                        // 将当前方案加入方案列表
                        schemes.Add(currentScheme);
                    }
                    // 2. 否则判断是否为当前方案下的【子物料 BOM 行】：
                    // 特征：存在当前活跃方案，且 B 列（名称）或 C 列（型号）有内容，且 D 列有数量
                    else if (currentScheme != null && (!string.IsNullOrWhiteSpace(colB) || !string.IsNullOrWhiteSpace(colC)))
                    {
                        // 解析数量 (默认为 1)
                        int qty = 1;
                        if (double.TryParse(colD, out double dQty))
                        {
                            qty = (int)Math.Round(dQty);
                            if (qty <= 0) qty = 1;
                        }

                        // 解析单价 (F 列在子项中对应单价)
                        double price = 0.0;
                        if (double.TryParse(colF, out double dPrice))
                        {
                            price = dPrice;
                        }

                        // 实例化子物料对象
                        var bomItem = new SecondaryBomItem
                        {
                            Name = colB.Trim(),
                            Model = colC.Trim(),
                            Quantity = qty,
                            Unit = string.IsNullOrWhiteSpace(colE) ? "只" : colE.Trim(), // --硬编码-- 默认单位为"只"
                            UnitPrice = price
                        };

                        // 追加到当前方案的子物料清单中
                        currentScheme.BomItems.Add(bomItem);
                        // 子 BOM 计数自增
                        totalBomCount++;
                    }
                }

                // 校验是否识别到有效方案
                if (schemes.Count == 0)
                {
                    return (false, "未能在当前表格中识别到符合规范的二次方案行！请检查方案行 C 列是否包含方案名且 B 列为空。", 0, 0);
                }

                // 批量保存入库到 SQLite 个人数据库中
                int savedCount = PersonalComponentDbService.BatchSaveSecondarySchemes(schemes);

                // 组装成功消息
                string msg = $"成功识别并导入 {savedCount} 个二次图回路方案，包含 {totalBomCount} 条二次 BOM 物料明细！已同步存入本地 SQLite 数据库。";
                return (true, msg, savedCount, totalBomCount);
            }
            catch (Exception ex)
            {
                // 记录日志并返回错误
                LogHelper.WriteLog($"[SecondaryImport] 导入异常: {ex.Message}");
                return (false, $"解析工作表发生异常: {ex.Message}", 0, 0);
            }
        }

        #region 内部辅助安全取值

        /// <summary>
        /// 从 Excel 2D 内存矩阵中安全提取指定行列的字符串
        /// </summary>
        private static string GetSafeString(object[,] matrix, int row, int col)
        {
            // 校验边界
            if (matrix == null || row > matrix.GetLength(0) || col > matrix.GetLength(1))
            {
                return string.Empty;
            }

            // 提取单元格对象
            object val = matrix[row, col];
            // 为空返回空串
            if (val == null) return string.Empty;

            // 转换为安全字符串
            return val.ToString()?.Trim() ?? string.Empty;
        }

        #endregion

        #region Excel 元件组与二次回路图号绑定业务

        /// <summary>
        /// Excel 二次元件组按型号聚合去重后的 DTO
        /// </summary>
        public class ExcelComponentGroupItemDto
        {
            // 元件组规格型号 (如 *KM变频器，唯一键)
            public string GroupModel { get; set; } = string.Empty;
            // 绑定的回路图号 (如 接触器变频器)
            public string BoundDwgCode { get; set; } = string.Empty;
            // 在当前工作表中出现的总次数
            public int OccurrenceCount { get; set; }
            // 涵盖的箱柜名称列表 (如 ["1AA1", "1AA2"])
            public List<string> Cabinets { get; set; } = new List<string>();
            // 对应的全部物理行号列表 (如 [346, 415, 480, 514])
            public List<int> RowIndexes { get; set; } = new List<int>();
        }

        /// <summary>
        /// 保存绑定的请求入参 DTO
        /// </summary>
        public class ComponentGroupBindingSaveDto
        {
            // 元件组型号规格
            public string GroupModel { get; set; } = string.Empty;
            // 绑定的回路图号 (如 接触器变频器)
            public string BoundDwgCode { get; set; } = string.Empty;
            // 目标物理行号列表
            public List<int>? RowIndexes { get; set; }
            // 兼容单行号传递
            public int RowIndex { get; set; }
        }

        /// <summary>
        /// 扫描当前活动 Excel 工作表中所有的二次元件组 (严格判定 B 列='元件组'，按型号去重输出)
        /// 遵循规范：每 3 行代码至少包含 1 行中文注释，操作在服务层完成
        /// </summary>
        /// <returns>去重后的二次元件组集合</returns>
        public static List<ExcelComponentGroupItemDto> ScanExcelComponentGroups()
        {
            var resultList = new List<ExcelComponentGroupItemDto>();
            try
            {
                // 使用公共上下文安全获取当前活动 Excel 实例与活动工作表
                var context = Tool.GetActiveExcelContext();
                if (context == null || context.Sheet == null)
                {
                    return resultList;
                }

                dynamic sheet = context.Sheet;
                dynamic? activeWb = context.Wb;

                // 获取当前工作表定义的所有有效箱柜
                var cabinets = Tool.GetSheetValidCabinets(sheet, activeWb);
                if (cabinets == null || cabinets.Count == 0) return resultList;

                // 使用字典按型号规格进行唯一去重聚合
                var groupMap = new Dictionary<string, ExcelComponentGroupItemDto>(StringComparer.OrdinalIgnoreCase);

                // 遍历每一个箱柜
                foreach (var cab in cabinets)
                {
                    int k = cab.Key;
                    // 获取该箱柜的标准行索引定义 (柜信息行、元器件起始行、小计行)
                    var (sumRow, detRow, subsumRow, tolsumRow) = Tool.FindStandardCategoryRowIndexes((object)sheet, k);
                    // 元器件起始行与截止行
                    int compStartRow = detRow + 2;
                    int compEndRow = subsumRow - 1;
                    if (compStartRow > compEndRow) continue;

                    // 提取箱柜物理名称 (如 1AA1)
                    string cabName = sheet.Cells[detRow, 1]?.Value?.ToString()?.Trim() ?? $"柜{k}";
                    // 若 A 列是定义标签，尝试提取其具体展示名称
                    if (cabName.StartsWith("Cab_Det_", StringComparison.OrdinalIgnoreCase))
                    {
                        cabName = sheet.Cells[detRow, 2]?.Value?.ToString()?.Trim() ?? $"柜{k}";
                    }

                    // 逐行扫描该箱柜中的元器件
                    for (int r = compStartRow; r <= compEndRow; r++)
                    {
                        // 提取 B 列: 类别 (第 2 列)
                        string category = sheet.Cells[r, 2]?.Value?.ToString()?.Trim() ?? string.Empty;
                        // 提取 C 列: 型号规格 (第 3 列)
                        string model = sheet.Cells[r, 3]?.Value?.ToString()?.Trim() ?? string.Empty;

                        // 判别准则修正 (严格识别)：
                        // 1. B 列明确等于 "元件组"
                        // 2. 若 B 列为空白且 C 列以 "*" 开头作为辅助兜底
                        // 3. 绝不使用 Contains("*")，排除任何 B 列非元件组的行 (如开孔 HK91*91)
                        bool isCategoryGroup = string.Equals(category, "元件组", StringComparison.OrdinalIgnoreCase);
                        bool isModelStar = string.IsNullOrWhiteSpace(category) && model.StartsWith("*");
                        bool isComponentGroup = (isCategoryGroup || isModelStar) && !string.IsNullOrWhiteSpace(model);

                        if (isComponentGroup)
                        {
                            // 提取第 32 列 (AF列) 中已持久化绑定的图号
                            string boundCode = sheet.Cells[r, 32]?.Value?.ToString()?.Trim() ?? string.Empty;

                            // 查找或创建该型号对应的聚合实体
                            if (!groupMap.TryGetValue(model, out var item))
                            {
                                item = new ExcelComponentGroupItemDto
                                {
                                    GroupModel = model,
                                    BoundDwgCode = boundCode,
                                    OccurrenceCount = 0,
                                    Cabinets = new List<string>(),
                                    RowIndexes = new List<int>()
                                };
                                groupMap[model] = item;
                            }

                            // 累加出现次数
                            item.OccurrenceCount++;
                            // 记录物理行号
                            item.RowIndexes.Add(r);
                            // 记录涵盖箱柜 (去重)
                            if (!string.IsNullOrWhiteSpace(cabName) && !item.Cabinets.Contains(cabName))
                            {
                                item.Cabinets.Add(cabName);
                            }
                            // 若之前未获取到图号但当前行有图号，补齐图号
                            if (string.IsNullOrWhiteSpace(item.BoundDwgCode) && !string.IsNullOrWhiteSpace(boundCode))
                            {
                                item.BoundDwgCode = boundCode;
                            }
                        }
                    }
                }

                // 将聚合去重后的结果转为列表输出
                resultList.AddRange(groupMap.Values);
            }
            catch (Exception ex)
            {
                // 记录扫描异常日志
                LogHelper.WriteLog($"[SecondaryBind] ScanExcelComponentGroups 异常: {ex.Message}");
            }

            // 返回去重后的元件组列表
            return resultList;
        }

        /// <summary>
        /// 批量将二次元件组绑定的回路图号持久化写入 Excel 对应行的第 32 列 (AF 列)
        /// 遵循规范：每 3 行代码至少包含 1 行中文注释
        /// </summary>
        /// <param name="bindings">待写入的绑定映射列表</param>
        /// <returns>写入成功行数与结果信息</returns>
        public static (bool Success, int UpdatedCount, string Message) SaveExcelComponentGroupBindings(List<ComponentGroupBindingSaveDto>? bindings)
        {
            if (bindings == null || bindings.Count == 0)
            {
                return (false, 0, "提交的绑定映射列表为空！");
            }

            try
            {
                // 使用公共上下文安全获取当前活动 Excel 实例与活动工作表
                var context = Tool.GetActiveExcelContext();
                if (context == null || context.App == null || context.Sheet == null)
                {
                    return (false, 0, "未检测到可操作的活动 Excel 工作表，请确保工作簿处于打开状态！");
                }

                dynamic app = context.App;
                dynamic sheet = context.Sheet;

                // 挂起屏幕刷新与自动计算以提速
                app.ScreenUpdating = false;
                app.Calculation = -4135; // xlCalculationManual --硬编码-- 手动计算常量

                int count = 0;
                int modelCount = 0;
                try
                {
                    // 遍历所有待写入的绑定映射
                    foreach (var item in bindings)
                    {
                        string code = item.BoundDwgCode?.Trim() ?? string.Empty;
                        bool modelUpdated = false;

                        // 1. 若提供了具体的行号列表，全量回写这些行
                        if (item.RowIndexes != null && item.RowIndexes.Count > 0)
                        {
                            foreach (int r in item.RowIndexes)
                            {
                                if (r > 0)
                                {
                                    sheet.Cells[r, 32].Value = code;
                                    count++;
                                    modelUpdated = true;
                                }
                            }
                        }
                        // 2. 兼容单行号模式
                        else if (item.RowIndex > 0)
                        {
                            sheet.Cells[item.RowIndex, 32].Value = code;
                            count++;
                            modelUpdated = true;
                        }

                        if (modelUpdated)
                        {
                            modelCount++;
                        }
                    }
                }
                finally
                {
                    // 恢复屏幕刷新与自动计算
                    app.Calculation = -4105; // xlCalculationAutomatic --硬编码-- 自动计算常量
                    app.ScreenUpdating = true;
                }

                // 返回成功结果
                return (true, count, $"成功将 {modelCount} 个元件组型号 (共计覆盖 {count} 处单元格) 持久化写入 Excel 第 32 列 (AF列)！");
            }
            catch (Exception ex)
            {
                // 记录写入异常日志
                LogHelper.WriteLog($"[SecondaryBind] SaveExcelComponentGroupBindings 异常: {ex.Message}");
                return (false, 0, $"持久化保存至 Excel 发生异常: {ex.Message}");
            }
        }

        #endregion
    }
}
