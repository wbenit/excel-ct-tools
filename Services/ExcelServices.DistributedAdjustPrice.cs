using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：分布调价与【材料分布表】二维矩阵生成与反向更新
    /// </summary>
    public static partial class ExcelServices
    {
        // 分布调价无边框宿主窗体静态单例引用 (可空)
        private static DistributedAdjustPriceForm? _distributedAdjustPriceForm;

        // 元件汇总分布表工作表默认标准名称 --硬编码--
        public const string DistributionSheetName = "元件汇总分布表";

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“分布调价”窗口 (非模态，可自由交互 Excel)
        /// </summary>
        public static void ShowDistributedAdjustPriceDialog()
        {
            try
            {
                // 以非模态方式展示分布调价窗口，允许用户在操作界面与 Excel 之间无缝切换
                ShowModelessForm(ref _distributedAdjustPriceForm, () => new DistributedAdjustPriceForm());
            }
            catch (Exception ex)
            {
                // 弹出异常提示对话框
                System.Windows.Forms.MessageBox.Show($"弹出分布调价窗口失败: {ex.Message}", "错误提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 检查当前活动工作簿中是否已存在【元件汇总分布表】
        /// </summary>
        public static bool CheckDistributionSheetExists()
        {
            try
            {
                // 获取 Excel 顶级 Application 实例
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                // 校验 Application 对象有效性
                if (app == null) return false;

                // 获取当前活动工作簿
                dynamic activeWb = app.ActiveWorkbook;
                // 校验工作簿有效性
                if (activeWb == null) return false;

                // 遍历当前工作簿的所有工作表
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    // 获取并修剪工作表名称
                    string sheetName = Convert.ToString(sheet.Name)?.Trim() ?? string.Empty;
                    // 比对是否为元件汇总分布表 --硬编码--
                    if (string.Equals(sheetName, DistributionSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        // 命中存在
                        return true;
                    }
                }

                // 遍历完成未找到返回 false
                return false;
            }
            catch
            {
                // 异常时安全兜底返回不存在
                return false;
            }
        }

        /// <summary>
        /// 遍历当前工作簿，读取所有分类工作表（Sheet）及对应箱柜台数统计
        /// </summary>
        public static List<DistributionCategoryDto> GetCategorySheetsForDistribution()
        {
            // 初始化分类返回列表
            var result = new List<DistributionCategoryDto>();

            try
            {
                // 安全获取 Excel 应用程序上下文
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                // 校验 Application 有效性
                if (app == null)
                {
                    // 记录空上下文警告日志
                    LogHelper.WriteLog("[分布调价] GetCategorySheetsForDistribution: ExcelDnaSafeAccessor.GetApplication() 返回 null");
                    return result;
                }

                // 获取活动工作簿
                dynamic activeWb = app.ActiveWorkbook;
                // 校验工作簿有效性
                if (activeWb == null)
                {
                    // 记录无活动工作簿警告
                    LogHelper.WriteLog("[分布调价] GetCategorySheetsForDistribution: 当前无活动工作簿 (app.ActiveWorkbook is null)");
                    return result;
                }

                // 记录当前正在扫描的工作簿名称
                LogHelper.WriteLog($"[分布调价] 开始扫描工作簿 [{activeWb.Name}] 的分类工作表...");

                // 遍历工作簿中的所有工作表
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    // 获取当前工作表名称
                    string sheetName = Convert.ToString(sheet.Name) ?? string.Empty;
                    // 去除首尾空白字符
                    string trimmedName = sheetName.Trim();

                    // 精准排除系统辅助工作表与分布表自身 --硬编码--
                    if (string.Equals(trimmedName, "项目信息", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, DistributionSheetName, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, "材料分布表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, "元件汇总分布表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, "元件汇总表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, "元件汇总调价清单", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, "屏柜汇总表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, "屏柜分项表", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(trimmedName, "元器件数据管理", StringComparison.OrdinalIgnoreCase))
                    {
                        // 忽略系统保留表
                        continue;
                    }

                    // 声明当前分类下的箱柜累计总台数
                    int totalCount = 0;

                    // 规则 7: 优先通过内存二维数组极速只读扫描顶部汇总表区域 A1:H60 (仅 1 次 COM 调用，1ms 级响应)
                    try
                    {
                        // 读取顶部 60 行 8 列矩形区域 (覆盖标准成套报价表汇总区)
                        dynamic topRange = sheet.Range["A1:H60"];
                        object[,]? topMatrix = topRange?.Value2 as object[,];

                        if (topMatrix != null)
                        {
                            int rowCount = topMatrix.GetLength(0);
                            int colCount = topMatrix.GetLength(1);

                            // 1. 定位顶部箱柜汇总表的表头行 (通常在前 15 行内)
                            int headerRow = 0;
                            for (int r = 1; r <= Math.Min(15, rowCount); r++)
                            {
                                // 提取前 3 列文本
                                string c1 = Convert.ToString(topMatrix[r, 1])?.Trim() ?? "";
                                string c2 = Convert.ToString(topMatrix[r, 2])?.Trim() ?? "";
                                string c3 = (colCount >= 3) ? (Convert.ToString(topMatrix[r, 3])?.Trim() ?? "") : "";

                                // 命中标准表头关键字特征
                                if (c1.Contains("序号") || c2.Contains("序号") || c2.Contains("柜号") || c2.Contains("设备") || c3.Contains("型号"))
                                {
                                    headerRow = r;
                                    break;
                                }
                            }

                            // 2. 若成功识别到汇总表头，向下遍历各箱柜汇总数据行
                            if (headerRow > 0)
                            {
                                for (int r = headerRow + 1; r <= rowCount; r++)
                                {
                                    // 提取序号、柜号与设备名称/型号
                                    string noStr = Convert.ToString(topMatrix[r, 1])?.Trim() ?? "";
                                    string cabNo = (colCount >= 2) ? (Convert.ToString(topMatrix[r, 2])?.Trim() ?? "") : "";
                                    string cabName = (colCount >= 3) ? (Convert.ToString(topMatrix[r, 3])?.Trim() ?? "") : "";

                                    // 关键截断：遇到合计、总计、小计、大写、说明或签字落款等行，立即终止扫描 (绝不穿透至落款区)
                                    if (noStr.Contains("合计") || cabNo.Contains("合计") || cabName.Contains("合计") ||
                                        noStr.Contains("总计") || cabNo.Contains("总计") || cabName.Contains("总计") ||
                                        noStr.Contains("小计") || cabNo.Contains("小计") || cabName.Contains("小计") ||
                                        cabNo.Contains("大写") || cabName.Contains("大写") || noStr.Contains("说明") ||
                                        cabNo.Contains("说明") || cabNo.Contains("审核") || cabNo.Contains("批准") ||
                                        cabNo.Contains("制表") || cabNo.Contains("编制") || cabNo.Contains("明细") ||
                                        cabName.Contains("明细") || cabNo.Contains("元件"))
                                    {
                                        break;
                                    }

                                    // 排除 A、B、C 列全为空的无效空行
                                    if (string.IsNullOrWhiteSpace(noStr) && string.IsNullOrWhiteSpace(cabNo) && string.IsNullOrWhiteSpace(cabName))
                                    {
                                        continue;
                                    }

                                    // 排除中间可能重复出现的表头文本行
                                    if (noStr.Contains("序号") || cabNo.Contains("柜号") || cabName.Contains("设备"))
                                    {
                                        continue;
                                    }

                                    // 默认单台箱柜数量为 1
                                    int cabQty = 1;
                                    // 用户明确要求：数量一定在 F 列 (第 6 列)，直接读取 F 列单元格内容，绝不尝试解析 E 列
                                    if (colCount >= 6)
                                    {
                                        object qVal = topMatrix[r, 6];
                                        // 采用 double 容错解析 (防范 Excel 单元格浮点数值格式)
                                        if (qVal != null && double.TryParse(Convert.ToString(qVal), out double dQty) && dQty > 0)
                                        {
                                            cabQty = (int)Math.Round(dQty);
                                        }
                                    }

                                    // 累加当前箱柜有效台数
                                    totalCount += cabQty;
                                }
                            }
                        }
                    }
                    catch { }

                    // 3. 兜底 1: 若顶部扫描未提取到台数，尝试轻量统计现存的 Cab_Sum_ 定义名称数量 (不触发全量重构)
                    if (totalCount == 0)
                    {
                        try
                        {
                            int existingSumCount = 0;
                            dynamic sNames = sheet.Names;
                            if (sNames != null)
                            {
                                // 遍历工作表现存定义名称
                                foreach (dynamic n in sNames)
                                {
                                    string nStr = Convert.ToString(n.Name) ?? "";
                                    // 匹配 Cab_Sum_ 汇总行定义名称前缀
                                    if (nStr.IndexOf("Cab_Sum_", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        existingSumCount++;
                                    }
                                }
                            }
                            if (existingSumCount > 0)
                            {
                                // 采用现存汇总行数量
                                totalCount = existingSumCount;
                            }
                        }
                        catch { }
                    }

                    // 4. 兜底 2: 最终保底保证至少 1 台
                    if (totalCount == 0)
                    {
                        try
                        {
                            // 检查已用区域行数
                            dynamic usedRange = sheet.UsedRange;
                            if (usedRange != null && usedRange.Rows.Count > 5)
                            {
                                // 兜底记为 1 台
                                totalCount = 1;
                            }
                        }
                        catch { }
                    }

                    // 记录扫描日志
                    LogHelper.WriteLog($"[分布调价] 识别到分类表 [{sheetName}], 箱柜台数: {totalCount}");

                    // 添加至返回结果清单
                    result.Add(new DistributionCategoryDto
                    {
                        // 设定工作表名称
                        SheetName = sheetName,
                        // 设定箱柜统计台数
                        CabinetCount = totalCount,
                        // 默认全选
                        IsSelected = true
                    });
                }

                // 记录扫描结束日志
                LogHelper.WriteLog($"[分布调价] 分类工作表扫描完毕，共纳管 {result.Count} 个分类工作表");
            }
            catch (Exception ex)
            {
                // 输出调试日志
                LogHelper.WriteLog($"[分布调价] 读取分类工作表异常: {ex.Message}");
            }

            // 返回分类列表
            return result;
        }

        #region 内部临时数据结构

        /// <summary>
        /// 箱柜横向列信息实体
        /// </summary>
        private class CabinetColumnInfo
        {
            // 箱柜全局唯一索引
            public int Index { get; set; }

            // 所属分类工作表名称
            public string SheetName { get; set; } = string.Empty;

            // 箱柜柜号 (如 1AA, 11AA)
            public string CabinetNo { get; set; } = string.Empty;

            // 箱柜名称 (如 进线柜、联络柜)
            public string CabinetName { get; set; } = string.Empty;

            // 箱柜型号 (如 GCK)
            public string CabinetModel { get; set; } = string.Empty;

            // 箱柜类别 (如 低压抽屉柜、照明配电箱)
            public string CabinetCategory { get; set; } = string.Empty;

            // 明细行定义名称 (Cab_Det_X) 供设置跳转超链接
            public string DetNameTag { get; set; } = string.Empty;

            // 箱柜台数 (如 1、2)
            public int Quantity { get; set; } = 1;

            // 汇总行单价
            public decimal UnitPrice { get; set; } = 0;

            // 汇总行总价
            public decimal TotalPrice { get; set; } = 0;

            // 备注
            public string Remark { get; set; } = string.Empty;

            // 明细起始行 (规则 6: detRow + 2)
            public int CompStartRow { get; set; }

            // 明细终止行 (规则 6: subsumRow - 1)
            public int CompEndRow { get; set; }
        }

        /// <summary>
        /// 箱柜明细元器件原始条目
        /// </summary>
        private class RawComponentItem
        {
            // 所属箱柜索引
            public int CabinetIndex { get; set; }

            // 物理所在行
            public int PhysicalRow { get; set; }

            // 元器件名称 (B列)
            public string Name { get; set; } = string.Empty;

            // 规格型号 (C列)
            public string Model { get; set; } = string.Empty;

            // 厂家 (D列)
            public string Manufacturer { get; set; } = string.Empty;

            // 单位 (E列)
            public string Unit { get; set; } = "个";

            // 单台数量 (F列)
            public decimal Quantity { get; set; } = 1;

            // 单价 (G列)
            public decimal UnitPrice { get; set; } = 0;

            // 合价 (H列)
            public decimal TotalPrice { get; set; } = 0;

            // 备注 (I列)
            public string Remark { get; set; } = string.Empty;

            // 方案名称 (AC列，第29列)
            public string TemplateName { get; set; } = string.Empty;
        }

        /// <summary>
        /// 聚合后的分布元器件行数据实体
        /// </summary>
        private class AggregatedComponentRow
        {
            // 元件名称 (A列)
            public string Name { get; set; } = string.Empty;

            // 规格型号 (B列)
            public string Model { get; set; } = string.Empty;

            // 单位 (C列)
            public string Unit { get; set; } = string.Empty;

            // 汇总总数量 (D列)
            public decimal TotalQuantity { get; set; } = 0;

            // 调价单价 (E列)
            public decimal UnitPrice { get; set; } = 0;

            // 厂家 (G列)
            public string Manufacturer { get; set; } = string.Empty;

            // 表价/面价 (H列)
            public decimal ListPrice { get; set; } = 0;

            // 折扣系数 (I列)
            public decimal Discount { get; set; } = 1.0m;

            // 实际成本 (J列)
            public decimal CostPrice { get; set; } = 0;

            // 差价 (K列)
            public decimal PriceDifference { get; set; } = 0;

            // 备注 (只读记录)
            public string Remark { get; set; } = string.Empty;

            // 方案名称 (只读记录)
            public string TemplateName { get; set; } = string.Empty;

            // 各箱柜单台数量字典 (Key: 箱柜索引 0..M-1, Value: 该箱柜单台数量)
            public Dictionary<int, decimal> CabinetQuantities { get; set; } = new Dictionary<int, decimal>();
        }

        #endregion

        /// <summary>
        /// 智能推断箱柜类别 (对应截图第 6 行: 低压抽屉柜、照明配电箱、高压开关柜等)
        /// </summary>
        private static string InferCabinetCategory(string model, string name, string sheet)
        {
            // 组合并转为大写以便特征匹配
            string combined = $"{model} {name} {sheet}".ToUpperInvariant();

            // 1. 低压抽屉柜特征 (GCK, GCS, MNS, 抽屉)
            if (combined.Contains("GCK") || combined.Contains("GCS") || combined.Contains("MNS") || combined.Contains("抽屉"))
            {
                // 命中抽屉柜
                return "低压抽屉柜";
            }

            // 2. 低压固定柜特征 (GGD, 固定)
            if (combined.Contains("GGD") || combined.Contains("固定"))
            {
                // 命中固定柜
                return "低压固定柜";
            }

            // 3. 高压开关柜特征 (KYN, KYN28, 高压, 中置, 10KV)
            if (combined.Contains("KYN") || combined.Contains("高压") || combined.Contains("中置") || combined.Contains("10KV"))
            {
                // 命中高压柜
                return "高压开关柜";
            }

            // 4. 照明配电箱/双电源箱特征 (双电源, 照明, PZ30, 配电箱, 箱)
            if (combined.Contains("双电源") || combined.Contains("照明") || combined.Contains("PZ") || combined.Contains("箱"))
            {
                // 命中照明配电箱
                return "照明配电箱";
            }

            // 5. 动力配电箱特征 (动力, XL-21)
            if (combined.Contains("动力") || combined.Contains("XL-21") || combined.Contains("XL21"))
            {
                // 命中动力箱
                return "动力配电箱";
            }

            // 6. 箱式变电站特征 (变, 美变, 欧变, 箱变)
            if (combined.Contains("变") || combined.Contains("箱变") || combined.Contains("美变") || combined.Contains("欧变"))
            {
                // 命中箱变
                return "美式箱变";
            }

            // 兜底返回所属分类工作表名或通用配电柜
            return !string.IsNullOrWhiteSpace(sheet) ? sheet : "低压抽屉柜";
        }

        /// <summary>
        /// 核心服务：生成【元件汇总分布表】(二维矩阵交叉表，完全按截图标准设计)
        /// </summary>
        public static bool GenerateComponentDistributionSheet(GenerateDistributionRequest request)
        {
            try
            {
                // 1. 获取 Excel 顶层应用对象
                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return false;

                // 获取活动工作簿
                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null) return false;

                // 2. 准备容器收集所有参与分布调价的箱柜与元器件
                var cabinetList = new List<CabinetColumnInfo>();
                var rawComponentList = new List<RawComponentItem>();

                // 冻结屏幕刷新与警告，提升批量执行速度
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                int globalCabIndex = 0;

                // 3. 遍历用户选中的各个分类工作表
                foreach (string sheetName in request.SelectedSheets)
                {
                    dynamic? sheet = null;
                    try
                    {
                        // 获取当前分类工作表
                        sheet = activeWb.Worksheets[sheetName];
                    }
                    catch
                    {
                        continue;
                    }
                    if (sheet == null) continue;

                    // 严格遵守规则 8：在操作 Excel 表格前，显式调用 FixAndFillCabinetNamesForSheet 确保规则 6 定义名称正确性
                    Tool.FixAndFillCabinetNamesForSheet(sheet);

                    // 获取当前工作表经过自愈校准后的有效箱柜列表 (规则 6)
                    var validCabinets = Tool.GetSheetValidCabinets(sheet, activeWb);
                    if (validCabinets.Count == 0) continue;

                    // 遍历当前分类表中的每个箱柜
                    foreach (var cabEntry in validCabinets)
                    {
                        var anchor = cabEntry.Value;
                        // 校验箱柜明细行锚点是否存在
                        if (anchor.Det == null) continue;

                        int detRow = Convert.ToInt32(anchor.Det.Row);
                        // 若未绑定小计行，则容错安全跳过
                        if (anchor.Subsum == null) continue;
                        int subsumRow = Convert.ToInt32(anchor.Subsum.Row);

                        // 按照规则 6 定义：detRow + 2 为元器件起始行，subsumRow - 1 为元器件终止行
                        int compStartRow = detRow + 2;
                        int compEndRow = subsumRow - 1;

                        // 提取箱柜汇总行数据
                        string cabNo = "";
                        string cabName = "";
                        string cabModel = "";
                        int cabQty = 1;
                        decimal cabUnitPrice = 0;
                        decimal cabTotalPrice = 0;
                        string cabRemark = "";

                        if (anchor.Sum != null)
                        {
                            int sumRow = Convert.ToInt32(anchor.Sum.Row);
                            // 一次性读取汇总行 A:H 区域数据 (覆盖柜号B、名称C、型号D、单位E、数量F、单价G、总价H，规则 7)
                            object[,] sumMatrix = (object[,])sheet.Range[$"A{sumRow}:H{sumRow}"].Value2;
                            cabNo = Convert.ToString(sumMatrix[1, 2])?.Trim() ?? "";
                            cabName = Convert.ToString(sumMatrix[1, 3])?.Trim() ?? "";
                            cabModel = Convert.ToString(sumMatrix[1, 4])?.Trim() ?? "";

                            // 数量一定在 F 列 (第 6 列)，直接读取 F 列，不解析 E 列
                            object qObj = sumMatrix[1, 6];
                            // 采用 double 容错解析台数
                            if (qObj != null && double.TryParse(Convert.ToString(qObj), out double parsedDQ) && parsedDQ > 0)
                            {
                                // 赋值有效箱柜台数
                                cabQty = (int)Math.Round(parsedDQ);
                            }

                            // 提取单价 (G列第7列) 与总价 (H列第8列)
                            decimal.TryParse(Convert.ToString(sumMatrix[1, 7]), out cabUnitPrice);
                            decimal.TryParse(Convert.ToString(sumMatrix[1, 8]), out cabTotalPrice);
                        }

                        // 拦截非箱柜行 (如合计、总计等)
                        if (cabNo.Contains("合计") || cabName.Contains("合计") || cabNo.Contains("总计") || cabName.Contains("总计"))
                        {
                            // 跳过非箱柜行
                            continue;
                        }

                        // 若汇总行柜号为空，则尝试从明细行提取
                        if (string.IsNullOrWhiteSpace(cabNo))
                        {
                            cabNo = Convert.ToString(sheet.Cells[detRow, 2].Value)?.Trim() ?? $"柜{globalCabIndex + 1}";
                        }
                        if (string.IsNullOrWhiteSpace(cabName))
                        {
                            cabName = cabNo;
                        }

                        // 封装箱柜列信息实体
                        var cabInfo = new CabinetColumnInfo
                        {
                            Index = globalCabIndex,
                            SheetName = sheetName,
                            CabinetNo = cabNo,
                            CabinetName = cabName,
                            CabinetModel = cabModel,
                            CabinetCategory = InferCabinetCategory(cabModel, cabName, sheetName),
                            DetNameTag = $"Cab_Det_{cabEntry.Key}",
                            Quantity = cabQty,
                            UnitPrice = cabUnitPrice,
                            TotalPrice = cabTotalPrice,
                            Remark = cabRemark,
                            CompStartRow = compStartRow,
                            CompEndRow = compEndRow
                        };
                        cabinetList.Add(cabInfo);

                        // 提取元器件区域数据（规则 6：detRow + 2 至 subsumRow - 1）
                        if (compEndRow >= compStartRow)
                        {
                            // 使用二维数组一次性读取元器件 A 列至 AC 列（第1列至第29列）(规则 7)
                            dynamic compRange = sheet.Range[$"A{compStartRow}:AC{compEndRow}"];
                            object[,] compMatrix = (object[,])compRange.Value2;
                            int totalCompRows = compEndRow - compStartRow + 1;

                            for (int r = 1; r <= totalCompRows; r++)
                            {
                                string cName = Convert.ToString(compMatrix[r, 2])?.Trim() ?? "";
                                string cModel = Convert.ToString(compMatrix[r, 3])?.Trim() ?? "";
                                string cMfg = Convert.ToString(compMatrix[r, 4])?.Trim() ?? "";
                                string cUnit = Convert.ToString(compMatrix[r, 5])?.Trim() ?? "";

                                // 若名称与型号均为空，则属于元器件区域空行，跳过提取
                                if (string.IsNullOrWhiteSpace(cName) && string.IsNullOrWhiteSpace(cModel))
                                {
                                    continue;
                                }

                                // 过滤可能误入的局部小计字样
                                if (cName.Contains("小计") || cModel.Contains("小计"))
                                {
                                    continue;
                                }

                                decimal cQty = 0;
                                decimal cPrice = 0;
                                decimal cTotal = 0;
                                decimal.TryParse(Convert.ToString(compMatrix[r, 6]), out cQty);
                                decimal.TryParse(Convert.ToString(compMatrix[r, 7]), out cPrice);
                                decimal.TryParse(Convert.ToString(compMatrix[r, 8]), out cTotal);

                                string cRemark = Convert.ToString(compMatrix[r, 9])?.Trim() ?? "";
                                // 提取 AC 列（第 29 列）方案名称
                                string cTemplate = Convert.ToString(compMatrix[r, 29])?.Trim() ?? "";

                                rawComponentList.Add(new RawComponentItem
                                {
                                    CabinetIndex = globalCabIndex,
                                    PhysicalRow = compStartRow + r - 1,
                                    Name = cName,
                                    Model = cModel,
                                    Manufacturer = cMfg,
                                    Unit = string.IsNullOrWhiteSpace(cUnit) ? "个" : cUnit,
                                    Quantity = cQty > 0 ? cQty : 1,
                                    UnitPrice = cPrice,
                                    TotalPrice = cTotal,
                                    Remark = cRemark,
                                    TemplateName = cTemplate
                                });
                            }
                        }

                        // 递增全局箱柜计数
                        globalCabIndex++;
                    }
                }

                // 校验是否存在有效箱柜
                if (cabinetList.Count == 0)
                {
                    throw new InvalidOperationException("所选分类工作表中未探测到有效的箱柜数据");
                }

                // 4. 按照合并规则聚合元器件，形成纵向分布行集合
                var cond = request.MergeConditions ?? new DistributionMergeConditionsDto();
                var groupedComponents = new Dictionary<string, AggregatedComponentRow>();

                foreach (var raw in rawComponentList)
                {
                    // 构建分组聚合 Key
                    string keyName = raw.Name.Trim();
                    string keyModel = raw.Model.Trim();
                    string keyMfg = cond.ByManufacturer ? raw.Manufacturer.Trim() : "";
                    if (!cond.IncludeNoManufacturer && string.IsNullOrWhiteSpace(keyMfg))
                    {
                        keyMfg = "--无厂家--";
                    }
                    string keyPrice = cond.ByPrice ? raw.UnitPrice.ToString("F2") : "";
                    string keyRemark = cond.ByRemark ? raw.Remark.Trim() : "";
                    string keyTemplate = cond.ByTemplateName ? raw.TemplateName.Trim() : "";

                    // 组合生成唯一复合 Key
                    string groupKey = $"{keyName}|||{keyModel}|||{keyMfg}|||{keyPrice}|||{keyRemark}|||{keyTemplate}";

                    if (!groupedComponents.TryGetValue(groupKey, out var aggRow))
                    {
                        aggRow = new AggregatedComponentRow
                        {
                            Name = raw.Name,
                            Model = raw.Model,
                            Manufacturer = raw.Manufacturer,
                            Unit = raw.Unit,
                            UnitPrice = raw.UnitPrice,
                            ListPrice = raw.UnitPrice,
                            Discount = 1.0m,
                            Remark = raw.Remark,
                            TemplateName = raw.TemplateName
                        };
                        groupedComponents[groupKey] = aggRow;
                    }

                    // 累加该箱柜内的单台数量
                    if (!aggRow.CabinetQuantities.ContainsKey(raw.CabinetIndex))
                    {
                        aggRow.CabinetQuantities[raw.CabinetIndex] = 0;
                    }
                    aggRow.CabinetQuantities[raw.CabinetIndex] += raw.Quantity;
                }

                // 计算每个聚合元器件的总汇总数量（加权箱柜台数）
                foreach (var kvp in groupedComponents)
                {
                    var agg = kvp.Value;
                    decimal sumQty = 0;
                    foreach (var cabQtyKvp in agg.CabinetQuantities)
                    {
                        int cabIdx = cabQtyKvp.Key;
                        decimal singleQty = cabQtyKvp.Value;
                        int cabCount = cabinetList[cabIdx].Quantity;
                        sumQty += singleQty * cabCount;
                    }
                    agg.TotalQuantity = sumQty;
                }

                // 转换为列表并执行排序
                var compRowList = groupedComponents.Values.ToList();
                var sortSettings = request.SortSettings ?? new DistributionSortSettingsDto();

                if (sortSettings.SortType == "mfg_name_model")
                {
                    // 按厂家、名称、型号排序
                    compRowList = compRowList.OrderBy(c => c.Manufacturer).ThenBy(c => c.Name).ThenBy(c => c.Model).ToList();
                }
                else
                {
                    // 默认按名称、型号、厂家排序
                    compRowList = compRowList.OrderBy(c => c.Name).ThenBy(c => c.Model).ThenBy(c => c.Manufacturer).ToList();
                }

                // 5. 准备在 Excel 中创建或重置【元件汇总分布表】
                dynamic? distrSheet = null;
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    if (string.Equals(Convert.ToString(ws.Name)?.Trim(), DistributionSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        distrSheet = ws;
                        break;
                    }
                }

                if (distrSheet == null)
                {
                    // 新建工作表，并放置在所有工作表最前面
                    distrSheet = activeWb.Worksheets.Add(activeWb.Worksheets[1]);
                    distrSheet.Name = DistributionSheetName;
                }
                else
                {
                    // 若已存在，先清空全表内容与格式
                    distrSheet.Cells.Clear();
                }

                // 6. 确定左侧固定列数: 11列 (A~K列，见截图设计) 与箱柜横向总列数
                int fixedColCount = 11;
                // 箱柜总列数
                int totalCabCols = cabinetList.Count;
                // 整个大矩阵总列数 (A~K + 箱柜列)
                int totalMatrixCols = fixedColCount + totalCabCols;

                // 严格按照用户截图布局设定行号：
                int rowCabSeq = 1;       // 行 1: 柜序号 (K1: 柜序号, L1起: 1, 2, 3...)
                int rowCabSheet = 2;     // 行 2: 所属分类 (K2: 所属分类, L2起: 低压配电室...)
                int rowCabName = 3;      // 行 3: 箱柜名称 (K3: 箱柜名称, L3起: 进线柜, 联络柜...)
                int rowCabModel = 4;     // 行 4: 箱柜型号 (K4: 箱柜型号, L4起: GCK...)
                int rowCabNo = 5;        // 行 5: 柜号 (K5: 柜号, L5起: 蓝色超链接 1AA, 11AA...)
                int rowCabCategory = 6;  // 行 6: 箱柜类别 (K6: 箱柜类别, L6起: 低压抽屉柜, 照明配电箱...)
                int rowCabQty = 7;       // 行 7: 箱柜数量 (K7: 箱柜数量, L7起: 2, 1, 2...；A7: 项目名称)
                int rowCompHeader = 8;   // 行 8: 元件汇总表头 (A8~K8 为字段名，L8起各箱柜列标题为“数量”，整行开启 AutoFilter)
                int rowDataStart = 9;    // 行 9: 元器件数据起始行

                // 提取项目名称
                string projectName = "成套电气设备工程";
                try
                {
                    // 尝试从【项目信息】表中读取真实工程名称
                    dynamic projSheet = activeWb.Worksheets["项目信息"];
                    if (projSheet != null)
                    {
                        // 读取 B2 单元格工程名称
                        string pName = Convert.ToString(projSheet.Cells[2, 2].Value);
                        if (!string.IsNullOrWhiteSpace(pName)) projectName = pName.Trim();
                    }
                }
                catch { }

                // 在 A7 单元格写入项目名称加粗标题 (见截图设计)
                distrSheet.Cells[rowCabQty, 1].Value = $"项目名称: {projectName}";
                distrSheet.Cells[rowCabQty, 1].Font.Bold = true;
                distrSheet.Cells[rowCabQty, 1].Font.Size = 11;

                // 在行 4 与行 6 写入红色提示说明文字 (与截图红字一致)
                distrSheet.Cells[4, 4].Value = "*排序ID：指在箱柜中的排列位置，可输入任意数字(支持小数)。";
                distrSheet.Cells[4, 4].Font.Color = ColorTranslator.ToOle(Color.Red);
                distrSheet.Cells[4, 4].Font.Size = 9;
                distrSheet.Cells[4, 4].Font.Bold = true;

                distrSheet.Cells[6, 4].Value = "*更新到项目时，数量为空删除，为0保留。";
                distrSheet.Cells[6, 4].Font.Color = ColorTranslator.ToOle(Color.Red);
                distrSheet.Cells[6, 4].Font.Size = 9;
                distrSheet.Cells[6, 4].Font.Bold = true;

                // 写入 K1:K7 表头引导标签 (居中对齐，浅灰色底)
                distrSheet.Cells[rowCabSeq, fixedColCount].Value = "柜序号";
                distrSheet.Cells[rowCabSheet, fixedColCount].Value = "所属分类";
                distrSheet.Cells[rowCabName, fixedColCount].Value = "箱柜名称";
                distrSheet.Cells[rowCabModel, fixedColCount].Value = "箱柜型号";
                distrSheet.Cells[rowCabNo, fixedColCount].Value = "柜号";
                distrSheet.Cells[rowCabCategory, fixedColCount].Value = "箱柜类别";
                distrSheet.Cells[rowCabQty, fixedColCount].Value = "箱柜数量";

                dynamic labelRange = distrSheet.Range[$"K1:K7"];
                labelRange.Font.Bold = true;
                labelRange.HorizontalAlignment = -4108; // 居中
                labelRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242));

                // 构建横向箱柜表头二维数组 (共 7 行，从 L 列展开)
                object[,] cabHeaderMatrix = new object[7, totalCabCols];
                for (int j = 0; j < totalCabCols; j++)
                {
                    // 行 1: 柜序号 (1, 2, 3...)
                    cabHeaderMatrix[0, j] = j + 1;
                    // 行 2: 所属分类工作表名称
                    cabHeaderMatrix[1, j] = cabinetList[j].SheetName;
                    // 行 3: 箱柜名称 (进线柜、联络柜等)
                    cabHeaderMatrix[2, j] = cabinetList[j].CabinetName;
                    // 行 4: 箱柜型号 (GCK等)
                    cabHeaderMatrix[3, j] = cabinetList[j].CabinetModel;
                    // 行 5: 箱柜柜号 (文本预填，随后设为超链接)
                    cabHeaderMatrix[4, j] = cabinetList[j].CabinetNo;
                    // 行 6: 箱柜类别 (低压抽屉柜、照明配电箱等)
                    cabHeaderMatrix[5, j] = cabinetList[j].CabinetCategory;
                    // 行 7: 箱柜数量 (台数)
                    cabHeaderMatrix[6, j] = cabinetList[j].Quantity;
                }

                // 转换为 Excel 列名 (第 12 列 L 列开始)
                string startCabColLetter = GetExcelColumnLetter(fixedColCount + 1);
                string endCabColLetter = GetExcelColumnLetter(totalMatrixCols);

                // 一次性写入横向箱柜表头 (规则 7)
                distrSheet.Range[$"{startCabColLetter}1:{endCabColLetter}7"].Value2 = cabHeaderMatrix;
                distrSheet.Range[$"{startCabColLetter}1:{endCabColLetter}7"].HorizontalAlignment = -4108; // 居中

                // 针对行 2(所属分类) 与 行 6(箱柜类别) 设定浅灰低饱和度字体，与截图视觉完全一致
                distrSheet.Range[$"{startCabColLetter}2:{endCabColLetter}2"].Font.Color = ColorTranslator.ToOle(Color.FromArgb(120, 120, 120));
                distrSheet.Range[$"{startCabColLetter}6:{endCabColLetter}6"].Font.Color = ColorTranslator.ToOle(Color.FromArgb(120, 120, 120));

                // 为行 5 (柜号) 添加可点击的超链接，跳转至各分类工作表该箱柜明细位置 (规则 6 联动)
                for (int j = 0; j < totalCabCols; j++)
                {
                    int col = fixedColCount + 1 + j;
                    var cab = cabinetList[j];
                    string detTarget = !string.IsNullOrWhiteSpace(cab.DetNameTag) ? cab.DetNameTag : $"Cab_Det_{j + 1}";
                    try
                    {
                        distrSheet.Hyperlinks.Add(
                            Anchor: distrSheet.Cells[rowCabNo, col],
                            Address: "",
                            SubAddress: $"'{cab.SheetName}'!{detTarget}",
                            ScreenTip: $"点击跳转至【{cab.SheetName}】{cab.CabinetNo} 明细",
                            TextToDisplay: cab.CabinetNo
                        );
                    }
                    catch { }
                }

                // 7. 写入行 8：列标题行 (A8~K8 固定列标题，L8向右每一列为“数量”)
                object[,] compHeaderRow = new object[1, totalMatrixCols];
                compHeaderRow[0, 0] = "排序ID";
                compHeaderRow[0, 1] = "元件名称";
                compHeaderRow[0, 2] = "型号规格";
                compHeaderRow[0, 3] = "生产厂家";
                compHeaderRow[0, 4] = "表价";
                compHeaderRow[0, 5] = "折扣";
                compHeaderRow[0, 6] = "报出";
                compHeaderRow[0, 7] = "备注";
                compHeaderRow[0, 8] = "单位";
                compHeaderRow[0, 9] = "类别";
                compHeaderRow[0, 10] = "元件总数";

                // 横向各箱柜列标题：截图显示全部为“数量”
                for (int j = 0; j < totalCabCols; j++)
                {
                    compHeaderRow[0, fixedColCount + j] = "数量";
                }

                // 一次性写入行 8 列标题 (规则 7)
                distrSheet.Range[$"A{rowCompHeader}:{endCabColLetter}{rowCompHeader}"].Value2 = compHeaderRow;
                distrSheet.Range[$"A{rowCompHeader}:{endCabColLetter}{rowCompHeader}"].Font.Bold = true;
                distrSheet.Range[$"A{rowCompHeader}:{endCabColLetter}{rowCompHeader}"].HorizontalAlignment = -4108; // 居中
                distrSheet.Range[$"A{rowCompHeader}:{endCabColLetter}{rowCompHeader}"].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242));
                distrSheet.Rows[rowCompHeader].RowHeight = 22;

                // 8. 构建元器件数据大矩阵并批量写入 (行 9 起)
                int totalCompRowsCount = compRowList.Count;
                if (totalCompRowsCount > 0)
                {
                    object[,] mainDataMatrix = new object[totalCompRowsCount, totalMatrixCols];

                    for (int r = 0; r < totalCompRowsCount; r++)
                    {
                        var comp = compRowList[r];

                        // A 列 (第 1 列): 排序ID (从 1 递增)
                        mainDataMatrix[r, 0] = r + 1;
                        // B 列 (第 2 列): 元件名称
                        mainDataMatrix[r, 1] = comp.Name;
                        // C 列 (第 3 列): 型号规格
                        mainDataMatrix[r, 2] = comp.Model;
                        // D 列 (第 4 列): 生产厂家
                        mainDataMatrix[r, 3] = comp.Manufacturer;
                        // E 列 (第 5 列): 表价
                        mainDataMatrix[r, 4] = comp.ListPrice > 0 ? (object)comp.ListPrice : "";
                        // F 列 (第 6 列): 折扣
                        mainDataMatrix[r, 5] = comp.Discount > 0 ? (object)comp.Discount : 1.0m;
                        // G 列 (第 7 列): 报出 (调价核心输入列)
                        mainDataMatrix[r, 6] = comp.UnitPrice;
                        // H 列 (第 8 列): 备注
                        mainDataMatrix[r, 7] = comp.Remark;
                        // I 列 (第 9 列): 单位
                        mainDataMatrix[r, 8] = comp.Unit;
                        // J 列 (第 10 列): 类别 (固定填“元件”)
                        mainDataMatrix[r, 9] = "元件";
                        // K 列 (第 11 列): 元件总数 (预填数值，后续写入 SUMPRODUCT 公式)
                        mainDataMatrix[r, 10] = comp.TotalQuantity;

                        // L 列及以后: 各箱柜单台数量
                        for (int c = 0; c < totalCabCols; c++)
                        {
                            if (comp.CabinetQuantities.TryGetValue(c, out decimal q) && q > 0)
                            {
                                mainDataMatrix[r, fixedColCount + c] = q;
                            }
                            else
                            {
                                mainDataMatrix[r, fixedColCount + c] = ""; // 无此器件留空
                            }
                        }
                    }

                    int endDataRow = rowDataStart + totalCompRowsCount - 1;
                    // 一次性将全部数据矩阵写入 Excel (规则 7)
                    distrSheet.Range[$"A{rowDataStart}:{endCabColLetter}{endDataRow}"].Value2 = mainDataMatrix;

                    // 批量写入 A 列动态行号排序公式：=ROW()-ROW(A$8)，确保用户在 Excel 移动/剪切行时序号自动联动连续
                    object[,] formulasColA = new object[totalCompRowsCount, 1];
                    // 遍历构建 A 列动态自适应序号公式
                    for (int r = 0; r < totalCompRowsCount; r++)
                    {
                        // 动态引用表头行 rowCompHeader (第8行)，实现从 1 递增
                        formulasColA[r, 0] = $"=ROW()-ROW(A${rowCompHeader})";
                    }
                    // 一次性批量写入 A 列公式 (规则 7)
                    distrSheet.Range[$"A{rowDataStart}:A{endDataRow}"].Formula = formulasColA;

                    // 一次性批量写入 K 列动态公式 (规则 7)：元件总数 = SUMPRODUCT($L$7:$末尾$7, L9:末尾9)
                    object[,] formulasColK = new object[totalCompRowsCount, 1];
                    // 遍历构建 K 列乘积求和公式
                    for (int r = 0; r < totalCompRowsCount; r++)
                    {
                        // 计算当前元器件绝对物理行号
                        int currRow = rowDataStart + r;
                        // 组装 SUMPRODUCT 公式
                        formulasColK[r, 0] = $"=SUMPRODUCT(${startCabColLetter}${rowCabQty}:${endCabColLetter}${rowCabQty}, {startCabColLetter}{currRow}:{endCabColLetter}{currRow})";
                    }
                    // 一次性批量写入 K 列公式 (规则 7)
                    distrSheet.Range[$"K{rowDataStart}:K{endDataRow}"].Formula = formulasColK;

                    // 样式设定 (白底黑字可编辑，细黑色网格边框)：
                    distrSheet.Range[$"A{rowDataStart}:{endCabColLetter}{endDataRow}"].Interior.Color = ColorTranslator.ToOle(Color.White);
                    // J 列(类别) 灰字居中
                    distrSheet.Range[$"J{rowDataStart}:J{endDataRow}"].Font.Color = ColorTranslator.ToOle(Color.FromArgb(128, 128, 128));

                    // 对齐与数字格式设定 (与截图完全一致)
                    distrSheet.Range[$"A{rowDataStart}:A{endDataRow}"].HorizontalAlignment = -4108; // 排序ID居中
                    distrSheet.Range[$"B{rowDataStart}:D{endDataRow}"].HorizontalAlignment = -4131; // 名称型号厂家靠左
                    distrSheet.Range[$"E{rowDataStart}:G{endDataRow}"].HorizontalAlignment = -4152; // 表价折扣报出靠右
                    distrSheet.Range[$"E{rowDataStart}:E{endDataRow}"].NumberFormat = "0.00";       // 表价两位小数
                    distrSheet.Range[$"G{rowDataStart}:G{endDataRow}"].NumberFormat = "0.00";       // 报出单价两位小数
                    distrSheet.Range[$"H{rowDataStart}:H{endDataRow}"].HorizontalAlignment = -4131; // 备注靠左
                    distrSheet.Range[$"I{rowDataStart}:K{endDataRow}"].HorizontalAlignment = -4108; // 单位类别总数居中
                    distrSheet.Range[$"{startCabColLetter}{rowDataStart}:{endCabColLetter}{endDataRow}"].HorizontalAlignment = -4108; // 箱柜单台数量居中

                    // 全表网格细实线边框
                    dynamic allTableRange = distrSheet.Range[$"A1:{endCabColLetter}{endDataRow}"];
                    allTableRange.Borders.LineStyle = 1; // 实线
                    allTableRange.Borders.Weight = 2;    // xlThin
                    allTableRange.Borders.Color = ColorTranslator.ToOle(Color.FromArgb(218, 220, 224));

                    // 开启第 8 行自动筛选 (AutoFilter)
                    try
                    {
                        distrSheet.Range[$"A{rowCompHeader}:{endCabColLetter}{endDataRow}"].AutoFilter();
                    }
                    catch { }
                }

                // 9. 精确列宽设定 (完美匹配截图比例)
                distrSheet.Columns[1].ColumnWidth = 7;      // A 列: 排序ID
                distrSheet.Columns[2].ColumnWidth = 16;     // B 列: 元件名称
                distrSheet.Columns[3].ColumnWidth = 20;     // C 列: 型号规格
                distrSheet.Columns[4].ColumnWidth = 14;     // D 列: 生产厂家
                distrSheet.Columns[5].ColumnWidth = 11;     // E 列: 表价
                distrSheet.Columns[6].ColumnWidth = 6;      // F 列: 折扣
                distrSheet.Columns[7].ColumnWidth = 10;     // G 列: 报出
                distrSheet.Columns[8].ColumnWidth = 10;     // H 列: 备注
                distrSheet.Columns[9].ColumnWidth = 6;      // I 列: 单位
                distrSheet.Columns[10].ColumnWidth = 6;     // J 列: 类别
                distrSheet.Columns[11].ColumnWidth = 9;     // K 列: 元件总数
                // 各箱柜列 (L列向右) 设为 8.5
                for (int c = fixedColCount + 1; c <= totalMatrixCols; c++)
                {
                    distrSheet.Columns[c].ColumnWidth = 8.5;
                }

                // 激活并选中【元件汇总分布表】
                distrSheet.Activate();

                // 冻结前 8 行表头 (1~8 行为表头与柜号，第 9 行起为元器件数据)
                try
                {
                    // 获取当前 Excel 活动窗口
                    dynamic win = app.ActiveWindow;
                    if (win != null)
                    {
                        // 先解除历史冻结状态以保证窗口干净
                        win.FreezePanes = false;
                        // 视口滚动复位至第 1 行第 1 列
                        win.ScrollRow = 1;
                        win.ScrollColumn = 1;
                        // 设定在第 8 行下方拆分窗格
                        win.SplitRow = rowCompHeader;
                        // 不进行列拆分
                        win.SplitColumn = 0;
                        // 锁定开启冻结窗格
                        win.FreezePanes = true;
                    }
                }
                catch (Exception ex)
                {
                    // 记录冻结窗格失败日志
                    LogHelper.WriteLog($"[分布调价] 冻结前8行表头窗格异常: {ex.Message}");
                }

                distrSheet.Cells[rowDataStart, 7].Select(); // 选中第一个报出单价单元格 (G9)

                // 恢复屏幕刷新与提示
                app.ScreenUpdating = true;
                app.DisplayAlerts = true;

                // 成功返回
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[分布调价] 生成材料分布表发生严重异常: {ex.Message}\r\n{ex.StackTrace}");
                return false;
            }
            finally
            {
                try
                {
                    dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                    if (app != null)
                    {
                        app.ScreenUpdating = true;
                        app.DisplayAlerts = true;
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// 箱柜反向回写元器件项实体 (供分布调价反向回写精准匹配各柜数量与参数)
        /// </summary>
        private class CabinetUpdateComponentItem
        {
            // 元件名称
            public string Name { get; set; } = string.Empty;
            // 规格型号
            public string Model { get; set; } = string.Empty;
            // 计量单位
            public string Unit { get; set; } = "个";
            // 生产厂家
            public string Manufacturer { get; set; } = string.Empty;
            // 单价
            public decimal UnitPrice { get; set; } = 0;
            // 该箱柜在分布表中配置的单台数量
            public decimal Quantity { get; set; } = 0;
            // 是否在分布表单元格中具体填写了有效数量
            public bool HasQuantitySpecified { get; set; } = false;
            // 排序ID (对应分布表 A 列，支持任意小数)
            public double SortId { get; set; } = 0;
            // 该分布项是否已在阶段 1 匹配认领，防止重复认领或在阶段 2 误当新增项
            public bool IsConsumed { get; set; } = false;
        }

        /// <summary>
        /// 核心服务：从【材料分布表】反向一键更新调价与数量数据到所有分类工作表各箱柜明细中 (支持数量修改与自动插入新增行)
        /// </summary>
        public static DistributionUpdateResult UpdateFromComponentDistributionSheet(DistributionUpdateOptions options, Action<int, string>? progressCallback = null)
        {
            var result = new DistributionUpdateResult();
            // 暂存 Excel 原始计算模式 (用于异常与正常结束时的可靠恢复)
            dynamic? prevCalc = null;
            // 暂存 Excel 原始事件启用状态
            bool prevEvents = true;

            try
            {
                // 发送初始进度通知 (5%)
                progressCallback?.Invoke(5, "正在获取 Excel 实例并定位【元件汇总分布表】...");

                dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null)
                {
                    result.Message = "无法获取 Excel 应用程序实例";
                    return result;
                }

                dynamic activeWb = app.ActiveWorkbook;
                if (activeWb == null)
                {
                    result.Message = "未打开有效的工作簿";
                    return result;
                }

                // 1. 定位【元件汇总分布表】
                dynamic? distrSheet = null;
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    if (string.Equals(Convert.ToString(ws.Name)?.Trim(), DistributionSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        distrSheet = ws;
                        break;
                    }
                }

                if (distrSheet == null)
                {
                    result.Message = $"当前工作簿中未找到【{DistributionSheetName}】，请先生成分布表";
                    return result;
                }

                // 发送解析数据中进度通知 (8%)
                progressCallback?.Invoke(8, "正在读取并解析【元件汇总分布表】调价数据...");

                // 2. 从【元件汇总分布表】中动态定位元器件表头行 (自适应第 8 行与历史排版)
                int rowCompHeader = 8;
                // 遍历前 25 行寻找“元件名称”所在行
                for (int r = 1; r <= 25; r++)
                {
                    string bVal = Convert.ToString(distrSheet.Cells[r, 2].Value)?.Trim() ?? "";
                    string aVal = Convert.ToString(distrSheet.Cells[r, 1].Value)?.Trim() ?? "";
                    if (bVal.Contains("元件名称") || aVal.Contains("元件名称") || aVal.Contains("排序ID"))
                    {
                        rowCompHeader = r;
                        break;
                    }
                }

                // 动态推导各行物理行号 (新版标准: rowCompHeader = 8)
                int rowCabSheet = Math.Max(1, rowCompHeader - 6);   // 所属分类行 (第 2 行)
                int rowCabNo = Math.Max(1, rowCompHeader - 3);      // 箱柜柜号行 (第 5 行)
                int rowCabQty = Math.Max(1, rowCompHeader - 1);     // 箱柜数量行 (第 7 行)
                int rowDataStart = rowCompHeader + 1;               // 元器件数据起始行 (第 9 行)

                dynamic usedRange = distrSheet.UsedRange;
                int lastRow = usedRange.Rows.Count + usedRange.Row - 1;

                // 如果末尾行为总合计行，排除之
                string lastCellText = Convert.ToString(distrSheet.Cells[lastRow, 1].Value)?.Trim() ?? "";
                if (lastCellText.Contains("合计"))
                {
                    lastRow--;
                }

                if (lastRow < rowDataStart)
                {
                    result.Message = "【元件汇总分布表】中未提取到有效的元器件调价数据";
                    return result;
                }

                // 确定左侧固定列数: 11列 (A~K列，见截图设计)
                int fixedColCount = 11;
                // 检测使用区域总列数
                int totalUsedCols = distrSheet.UsedRange.Columns.Count + distrSheet.UsedRange.Column - 1;
                int totalCabCols = Math.Max(0, totalUsedCols - fixedColCount);

                // 存储横向箱柜列元数据列表
                var cabColumns = new List<CabinetColumnInfo>();

                if (totalCabCols > 0)
                {
                    string startCabColLetter = GetExcelColumnLetter(fixedColCount + 1);
                    string endCabColLetter = GetExcelColumnLetter(totalUsedCols);
                    // 批量读取箱柜表头: 所属分类、箱柜数量、柜号等 (规则 7)
                    object[,] headerMatrix = (object[,])distrSheet.Range[$"{startCabColLetter}1:{endCabColLetter}{rowCabQty}"].Value2;
                    for (int c = 1; c <= totalCabCols; c++)
                    {
                        // 提取所属分类 (行 2)
                        string sName = Convert.ToString(headerMatrix[rowCabSheet, c])?.Trim() ?? "";
                        // 提取箱柜柜号 (行 5)
                        string cNo = Convert.ToString(headerMatrix[rowCabNo, c])?.Trim() ?? "";
                        // 提取箱柜台数 (行 7)
                        int q = 1;
                        int.TryParse(Convert.ToString(headerMatrix[rowCabQty, c]), out q);

                        cabColumns.Add(new CabinetColumnInfo
                        {
                            Index = c - 1,
                            SheetName = sName,
                            Quantity = q > 0 ? q : 1,
                            CabinetNo = cNo
                        });
                    }
                }

                // 预先重算分布表以确保 A 列自适应行号公式与 SUMPRODUCT 处于最新计算结果状态
                try
                {
                    // 触发单表公式刷新重算
                    distrSheet.Calculate();
                }
                catch { }

                // 一次性读取数据大矩阵 (从 A9 到最后一列) (规则 7)
                string maxDataColLetter = totalCabCols > 0 ? GetExcelColumnLetter(totalUsedCols) : "K";
                object[,] distrMatrix = (object[,])distrSheet.Range[$"A{rowDataStart}:{maxDataColLetter}{lastRow}"].Value2;
                int totalCompRows = lastRow - rowDataStart + 1;

                // 通用单价与规格规则列表 (用于全局兜底匹配)
                var globalPriceRules = new List<AggregatedComponentRow>();
                // 按 "SheetName|CabinetNo" 归集每个箱柜的具体元器件清单 (含最新单价与修改后的单台数量)
                var cabSpecificComponentsMap = new Dictionary<string, List<CabinetUpdateComponentItem>>(StringComparer.OrdinalIgnoreCase);

                for (int r = 1; r <= totalCompRows; r++)
                {
                    // A 列: 排序ID (第 1 列)
                    double sortId = r;
                    double.TryParse(Convert.ToString(distrMatrix[r, 1]), out sortId);

                    // B 列: 元件名称 (第 2 列)
                    string name = Convert.ToString(distrMatrix[r, 2])?.Trim() ?? "";
                    // C 列: 型号规格 (第 3 列)
                    string model = Convert.ToString(distrMatrix[r, 3])?.Trim() ?? "";
                    // D 列: 生产厂家 (第 4 列)
                    string mfg = Convert.ToString(distrMatrix[r, 4])?.Trim() ?? "";
                    // E 列: 表价 (第 5 列)
                    decimal listPrice = 0;
                    decimal.TryParse(Convert.ToString(distrMatrix[r, 5]), out listPrice);
                    // G 列: 报出单价 (第 7 列，核心调价输入列)
                    decimal unitPrice = 0;
                    decimal.TryParse(Convert.ToString(distrMatrix[r, 7]), out unitPrice);
                    // H 列: 备注 (第 8 列)
                    string remark = Convert.ToString(distrMatrix[r, 8])?.Trim() ?? "";
                    // I 列: 单位 (第 9 列)
                    string unit = Convert.ToString(distrMatrix[r, 9])?.Trim() ?? "个";

                    if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(model))
                    {
                        continue;
                    }

                    globalPriceRules.Add(new AggregatedComponentRow
                    {
                        Name = name,
                        Model = model,
                        Unit = unit,
                        Manufacturer = mfg,
                        UnitPrice = unitPrice,
                        ListPrice = listPrice,
                        Remark = remark
                    });

                    // 遍历横向箱柜列，提取该器件在各箱柜的单台数量
                    for (int c = 0; c < cabColumns.Count; c++)
                    {
                        var colInfo = cabColumns[c];
                        string key = $"{colInfo.SheetName}|{colInfo.CabinetNo}";
                        if (!cabSpecificComponentsMap.TryGetValue(key, out var compList))
                        {
                            compList = new List<CabinetUpdateComponentItem>();
                            cabSpecificComponentsMap[key] = compList;
                        }

                        object qObj = distrMatrix[r, fixedColCount + 1 + c];
                        decimal parsedQty = 0;
                        bool hasQty = false;
                        if (qObj != null && !string.IsNullOrWhiteSpace(Convert.ToString(qObj)))
                        {
                            if (decimal.TryParse(Convert.ToString(qObj), out parsedQty))
                            {
                                hasQty = true;
                            }
                        }

                        compList.Add(new CabinetUpdateComponentItem
                        {
                            Name = name,
                            Model = model,
                            Unit = string.IsNullOrWhiteSpace(unit) ? "个" : unit,
                            Manufacturer = mfg,
                            UnitPrice = unitPrice,
                            Quantity = parsedQty,
                            HasQuantitySpecified = hasQty, // 若单元格为空则为 false，为 0 则为 true 且值为 0
                            SortId = sortId
                        });
                    }
                }

                if (globalPriceRules.Count == 0)
                {
                    result.Message = "【元件汇总分布表】中未能解析出有效调价规则项";
                    return result;
                }

                // 3. 冻结计算、事件与屏幕刷新以极致提升反向回写性能 (核心性能优化)
                try
                {
                    // 记录原有计算模式
                    prevCalc = app.Calculation;
                }
                catch { }
                try
                {
                    // 记录原有事件启用状态
                    prevEvents = app.EnableEvents;
                }
                catch { }

                try
                {
                    // 强制设为手动计算，避免每次写入单元格或插入行时 Excel 在后台触发全簿公式重算风暴
                    app.Calculation = -4135; // XlCalculation.xlCalculationManual
                }
                catch { }

                try
                {
                    // 禁用事件触发，杜绝 COM 内部监听频繁响应
                    app.EnableEvents = false;
                }
                catch { }

                // 禁用屏幕刷新与系统提示弹窗
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                // 从分布表箱柜列中提取涉及的有效目标分类工作表名称集合
                var targetSheetNames = cabColumns
                    .Select(c => c.SheetName)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                // 收集工作簿中实际存在的有效目标工作表列表，跳过所有无关表
                var validTargetSheets = new List<dynamic>();
                // 遍历当前工作簿的所有表以定位目标表
                foreach (dynamic ws in activeWb.Worksheets)
                {
                    // 提取工作表名称
                    string wsName = Convert.ToString(ws.Name)?.Trim() ?? "";
                    // 仅收录属于目标分类且非辅助表的工作表
                    if (targetSheetNames.Contains(wsName) &&
                        !string.Equals(wsName, DistributionSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        validTargetSheets.Add(ws);
                    }
                }

                // 预先统计总箱柜数以计算平滑进度
                int totalCabCountAcrossSheets = 0;
                // 缓存各表的有效箱柜列表映射 (使用正确的 CabinetAnchorModel 类型)
                var sheetCabMap = new Dictionary<string, List<KeyValuePair<int, Models.CabinetAnchorModel>>>(StringComparer.OrdinalIgnoreCase);
                // 遍历目标表执行自愈并统计箱柜
                foreach (dynamic sheet in validTargetSheets)
                {
                    // 提取表名
                    string sName = Convert.ToString(sheet.Name)?.Trim() ?? "";
                    // 严格遵守规则 8：在回写前显式调用自愈校准箱柜名称
                    Tool.FixAndFillCabinetNamesForSheet(sheet);
                    // 获取当前工作表的有效箱柜列表
                    var validCabs = Tool.GetSheetValidCabinets(sheet, activeWb);
                    // 登记字典缓存
                    sheetCabMap[sName] = validCabs;
                    // 累加总箱柜台数
                    totalCabCountAcrossSheets += validCabs.Count;
                }

                int updatedSheetCount = 0;
                int updatedCabCount = 0;
                int updatedCompCount = 0;
                int currentCabIndex = 0;

                // 收集所有发生更新的箱柜差量切片 (用于支持整簿跨表跨箱柜一键撤销与还原)
                var undoSlices = new List<DistributionCabinetSlice>();

                // 4. 仅遍历涉及的目标分类工作表并精准回写 (避免对无关表执行空转)
                foreach (dynamic sheet in validTargetSheets)
                {
                    string sheetName = Convert.ToString(sheet.Name)?.Trim() ?? "";
                    if (!sheetCabMap.TryGetValue(sheetName, out var validCabinets) || validCabinets.Count == 0)
                    {
                        continue;
                    }

                    bool sheetModified = false;

                    // 遍历工作表下的每个箱柜
                    foreach (var cabEntry in validCabinets)
                    {
                        currentCabIndex++;
                        var anchor = cabEntry.Value;
                        if (anchor.Det == null || anchor.Subsum == null) continue;

                        int detRow = Convert.ToInt32(anchor.Det.Row);
                        int subsumRow = Convert.ToInt32(anchor.Subsum.Row);
                        int compStartRow = detRow + 2;
                        int compEndRow = subsumRow - 1;
                        // 计算当前箱柜明细表头行 (detRow + 1)，用于生成动态序号公式 =ROW()-ROW(A${headerRow})
                        int headerRow = detRow + 1;

                        if (compEndRow < compStartRow) continue;

                        // 提取箱柜柜号
                        string cabNo = "";
                        if (anchor.Sum != null)
                        {
                            // 汇总行存在时从汇总行提取柜号 (B列)
                            int sumRow = Convert.ToInt32(anchor.Sum.Row);
                            cabNo = Convert.ToString(sheet.Cells[sumRow, 2].Value)?.Trim() ?? "";
                        }
                        if (string.IsNullOrWhiteSpace(cabNo))
                        {
                            // 汇总行无柜号时从明细行提取
                            cabNo = Convert.ToString(sheet.Cells[detRow, 2].Value)?.Trim() ?? "";
                        }

                        // 计算当前箱柜进度百分比 (10% ~ 90%)
                        int cabProgress = totalCabCountAcrossSheets > 0
                            ? 10 + (int)Math.Round((double)currentCabIndex / totalCabCountAcrossSheets * 80.0)
                            : 50;
                        // 触发进度回调通知前端当前正在更新的工作表与箱柜号
                        progressCallback?.Invoke(cabProgress, $"正在更新: [{sheetName}] - {cabNo} ({currentCabIndex}/{totalCabCountAcrossSheets})...");

                        // 查找该箱柜在分布表中的具体配置项列表
                        string cabKey = $"{sheetName}|{cabNo}";
                        cabSpecificComponentsMap.TryGetValue(cabKey, out var expectedCompItems);

                        // 重置该箱柜所有分布项的消费认领状态，确保状态纯净
                        if (expectedCompItems != null)
                        {
                            // 遍历条目重置消费标记
                            foreach (var item in expectedCompItems)
                            {
                                item.IsConsumed = false;
                            }
                        }

                        int cabRowsCount = compEndRow - compStartRow + 1;
                        // 一次性读取该箱柜的全部 30 列区域 (覆盖 A 列至 AD 列 CAD 句柄，规则 7)
                        dynamic compRange = sheet.Range[$"A{compStartRow}:AD{compEndRow}"];
                        object[,] compMatrix = (object[,])compRange.Value2;
                        // 克隆修改前的完整 30 列公式矩阵，作为撤销时的底层快照 (包含原序号、原单价公式、原合价公式等)
                        object[,] origFormulaMatrix = (object[,])((object[,])compRange.Formula).Clone();
                        // 记录更新前的关键行号
                        int origCompEndRow = compEndRow;
                        int origSubsumRow = subsumRow;
                        int insertedRowCount = 0;
                        int insertedRowStart = 0;
                        bool cabModified = false;

                        // 记录可用空白行相对索引列表 (1..cabRowsCount)
                        var availableEmptyRowIndices = new List<int>();
                        // 记录箱柜内首次出现的器件复合键与对应主行行号 (用于合并相同元器件)
                        var primaryRowMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                        // 判定是否执行合并相同元件：未勾选“不合并相同元件”时默认合并
                        bool shouldMergeSameBom = options == null || !options.NotMergeSameBom;

                        // 1. 扫描并更新已有行
                        for (int r = 1; r <= cabRowsCount; r++)
                        {
                            // 提取明细表现有行器件名称 (B列)
                            string cName = Convert.ToString(compMatrix[r, 2])?.Trim() ?? "";
                            // 提取明细表现有行规格型号 (C列)
                            string cModel = Convert.ToString(compMatrix[r, 3])?.Trim() ?? "";
                            // 提取明细表现有行生产厂家 (D列)
                            string cMfg = Convert.ToString(compMatrix[r, 4])?.Trim() ?? "";

                            // 识别空白行并记录
                            if (string.IsNullOrWhiteSpace(cName) && string.IsNullOrWhiteSpace(cModel))
                            {
                                // 登记可用空白行索引
                                availableEmptyRowIndices.Add(r);
                                continue;
                            }

                            // 构建相同元器件判定复合键 (名称 + 型号 + 厂家)
                            string sameCompKey = $"{cName.ToUpperInvariant()}|||{cModel.ToUpperInvariant()}|||{cMfg.ToUpperInvariant()}";

                            // 若启用了合并相同元件，且前面已经登记过该相同器件的主行
                            if (shouldMergeSameBom && primaryRowMap.TryGetValue(sameCompKey, out int mainR))
                            {
                                // 提取重复行的 CAD 句柄 (AD 列，第 30 列)
                                string subHandle = Convert.ToString(compMatrix[r, 30])?.Trim() ?? "";
                                // 提取主行的 CAD 句柄
                                string mainHandle = Convert.ToString(compMatrix[mainR, 30])?.Trim() ?? "";

                                // 若重复行有 CAD 句柄，将其安全追加合并至主行 AD 列，杜绝句柄丢失
                                if (!string.IsNullOrWhiteSpace(subHandle))
                                {
                                    if (string.IsNullOrWhiteSpace(mainHandle))
                                    {
                                        // 主行无句柄直接继承
                                        compMatrix[mainR, 30] = subHandle;
                                    }
                                    else
                                    {
                                        // 拆分现有句柄以去重
                                        var existingHandles = mainHandle.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                            .Select(h => h.Trim())
                                            .ToList();
                                        if (!existingHandles.Contains(subHandle))
                                        {
                                            // 逗号拼接追加句柄
                                            compMatrix[mainR, 30] = $"{mainHandle},{subHandle}";
                                        }
                                    }
                                }

                                // 清空当前重复行的全部 30 列数据 (实现合并消除重复行)
                                for (int col = 1; col <= 30; col++)
                                {
                                    compMatrix[r, col] = "";
                                }

                                // 将清空后的重复行登记为可用空行
                                availableEmptyRowIndices.Add(r);
                                cabModified = true;
                                updatedCompCount++;
                                continue;
                            }

                            // 登记该元器件首次出现的主行索引 (供后续相同元器件合并)
                            if (shouldMergeSameBom)
                            {
                                primaryRowMap[sameCompKey] = r;
                            }

                            // 采用多级优先级匹配当前箱柜特定分布项清单 (未被消费认领的项)
                            CabinetUpdateComponentItem? matchedExpected = null;
                            if (expectedCompItems != null)
                            {
                                // 优先级 1: 名称 + 型号 + 厂家完全匹配，且在当前箱柜明确填有有效数量
                                matchedExpected = expectedCompItems.FirstOrDefault(item =>
                                    !item.IsConsumed &&
                                    item.HasQuantitySpecified &&
                                    string.Equals(item.Name, cName, StringComparison.OrdinalIgnoreCase) &&
                                    string.Equals(item.Model, cModel, StringComparison.OrdinalIgnoreCase) &&
                                    string.Equals(item.Manufacturer, cMfg, StringComparison.OrdinalIgnoreCase));

                                // 优先级 2: 名称 + 型号匹配，且在当前箱柜明确填有有效数量
                                if (matchedExpected == null)
                                {
                                    // 次优先名称与型号匹配
                                    matchedExpected = expectedCompItems.FirstOrDefault(item =>
                                        !item.IsConsumed &&
                                        item.HasQuantitySpecified &&
                                        string.Equals(item.Name, cName, StringComparison.OrdinalIgnoreCase) &&
                                        string.Equals(item.Model, cModel, StringComparison.OrdinalIgnoreCase));
                                }

                                // 优先级 3: 仅型号匹配，且在当前箱柜明确填有有效数量
                                if (matchedExpected == null)
                                {
                                    // 兜底优先型号匹配有数量项
                                    matchedExpected = expectedCompItems.FirstOrDefault(item =>
                                        !item.IsConsumed &&
                                        item.HasQuantitySpecified &&
                                        string.Equals(item.Model, cModel, StringComparison.OrdinalIgnoreCase));
                                }

                                // 优先级 4 (删除判定): 只有当所有未消费条目中都明确未填数量时，才匹配到空数量项执行删除
                                if (matchedExpected == null)
                                {
                                    // 匹配同名同型号的空数量项
                                    matchedExpected = expectedCompItems.FirstOrDefault(item =>
                                        !item.IsConsumed &&
                                        !item.HasQuantitySpecified &&
                                        string.Equals(item.Name, cName, StringComparison.OrdinalIgnoreCase) &&
                                        string.Equals(item.Model, cModel, StringComparison.OrdinalIgnoreCase));
                                }

                                // 优先级 5 (兜底删除判定): 型号匹配且未填数量
                                if (matchedExpected == null)
                                {
                                    // 仅型号匹配空数量项
                                    matchedExpected = expectedCompItems.FirstOrDefault(item =>
                                        !item.IsConsumed &&
                                        !item.HasQuantitySpecified &&
                                        string.Equals(item.Model, cModel, StringComparison.OrdinalIgnoreCase));
                                }
                            }

                            // 全局单价规则兜底
                            var matchedGlobal = globalPriceRules.FirstOrDefault(p =>
                                string.Equals(p.Name, cName, StringComparison.OrdinalIgnoreCase) &&
                                string.Equals(p.Model, cModel, StringComparison.OrdinalIgnoreCase)) ??
                                globalPriceRules.FirstOrDefault(p =>
                                string.Equals(p.Model, cModel, StringComparison.OrdinalIgnoreCase));

                            if (matchedExpected != null)
                            {
                                // 立即标记为已消费认领，防止后续重复认领或在阶段 2 被当成新增项
                                matchedExpected.IsConsumed = true;

                                // 🌟 核心践行截图红字规范：*更新到项目时，数量为空删除，为0保留
                                if (!matchedExpected.HasQuantitySpecified)
                                {
                                    // 单元格为空：在该箱柜中删除该器件 (清空整行全部 30 列单元格内容)
                                    for (int col = 2; col <= 30; col++)
                                    {
                                        compMatrix[r, col] = "";
                                    }

                                    // 登记清空后的可用空白行
                                    availableEmptyRowIndices.Add(r);
                                    cabModified = true;
                                    updatedCompCount++;
                                }
                                else
                                {
                                    // 回写新规格型号 (若修改了)
                                    if (!string.IsNullOrWhiteSpace(matchedExpected.Model))
                                    {
                                        compMatrix[r, 3] = matchedExpected.Model;
                                    }

                                    // 回写新厂家 (若有调整)
                                    if (!string.IsNullOrWhiteSpace(matchedExpected.Manufacturer))
                                    {
                                        compMatrix[r, 4] = matchedExpected.Manufacturer;
                                    }

                                    // 回写新单价 (取 G 列报出单价)
                                    compMatrix[r, 7] = matchedExpected.UnitPrice;

                                    // 数量更新策略：
                                    // 若允许合并 (shouldMergeSameBom)，主行采用分布表中聚合总数量；
                                    // 若用户勾选了“不合并相同元件”，则保持当前行原有的数量不变，仅调价。
                                    if (shouldMergeSameBom)
                                    {
                                        // 采用分布表聚合总数量
                                        compMatrix[r, 6] = matchedExpected.Quantity;
                                    }

                                    // 提取当前最终数量以重新计算合价
                                    decimal currentQty = 0;
                                    decimal.TryParse(Convert.ToString(compMatrix[r, 6]), out currentQty);
                                    // 重新计算合价 (合价 = 数量 * 单价)
                                    compMatrix[r, 8] = Math.Round(currentQty * matchedExpected.UnitPrice, 2);

                                    cabModified = true;
                                    updatedCompCount++;
                                }
                            }
                            else if (matchedGlobal != null)
                            {
                                // 仅更新单价、型号、厂家，保持数量不变
                                if (!string.IsNullOrWhiteSpace(matchedGlobal.Model)) compMatrix[r, 3] = matchedGlobal.Model;
                                if (!string.IsNullOrWhiteSpace(matchedGlobal.Manufacturer)) compMatrix[r, 4] = matchedGlobal.Manufacturer;
                                compMatrix[r, 7] = matchedGlobal.UnitPrice;

                                // 计算并刷新合价
                                decimal qty = 0;
                                decimal.TryParse(Convert.ToString(compMatrix[r, 6]), out qty);
                                compMatrix[r, 8] = Math.Round(qty * matchedGlobal.UnitPrice, 2);

                                cabModified = true;
                                updatedCompCount++;
                            }
                        }

                        // 将已有行的修改先一次性写回当前 30 列区域 (规则 7)
                        compRange.Value2 = compMatrix;

                        // 2. 🌟 核心处理：对于分布表中针对该箱柜有数量 (> 0) 但未在阶段 1 被消费认领的“真正新增器件”
                        if (expectedCompItems != null)
                        {
                            // 筛选当前箱柜有数量且未被消费的新增器件
                            var pendingNewItems = expectedCompItems
                                .Where(item => item.Quantity > 0 && !item.IsConsumed)
                                .ToList();

                            if (pendingNewItems.Count > 0)
                            {
                                // 计算当前元器件区域缺少多少可用空行
                                int neededRows = pendingNewItems.Count - availableEmptyRowIndices.Count;
                                if (neededRows > 0)
                                {
                                    // 记录插行起始物理行与插入行数，供撤销时倒序物理删除
                                    insertedRowStart = subsumRow;
                                    insertedRowCount = neededRows;
                                    // 严格遵守规则 6：“如果元器件数量多于区域行数，先要插入行”
                                    // 批量一次性在小计行前插入 neededRows 整行，避免循环内单行多次插入导致的反复重排与重算
                                    sheet.Range[$"{subsumRow}:{subsumRow + neededRows - 1}"].Insert(-4121); // xlDown 批量插入新行

                                    // 将新插入的多行相对行号追加登记为可用空行
                                    for (int i = 0; i < neededRows; i++)
                                    {
                                        // 计算插入行相对元器件起始行的索引
                                        int newRelIdx = (subsumRow + i) - compStartRow + 1;
                                        // 登记为可用空行
                                        availableEmptyRowIndices.Add(newRelIdx);
                                    }

                                    // 插入行后元器件终止行与小计行顺延
                                    compEndRow += neededRows;
                                    subsumRow += neededRows;
                                }

                                // 遍历逐个追加新增器件到可用空行中
                                foreach (var newItem in pendingNewItems)
                                {
                                    // 标记消费状态
                                    newItem.IsConsumed = true;

                                    // 从可用空行列表中取出空行索引
                                    if (availableEmptyRowIndices.Count > 0)
                                    {
                                        // 提取并移出第一个空行
                                        int emptyIdx = availableEmptyRowIndices[0];
                                        availableEmptyRowIndices.RemoveAt(0);

                                        // 计算目标物理行号
                                        int targetRow = compStartRow + emptyIdx - 1;
                                        // 填入新增器件各项属性与公式 (动态使用当前箱柜表头行 headerRow)
                                        sheet.Cells[targetRow, 1].Formula = $"=ROW()-ROW(A${headerRow})";
                                        sheet.Cells[targetRow, 2].Value = newItem.Name;
                                        sheet.Cells[targetRow, 3].Value = newItem.Model;
                                        sheet.Cells[targetRow, 4].Value = newItem.Manufacturer;
                                        sheet.Cells[targetRow, 5].Value = newItem.Unit;
                                        sheet.Cells[targetRow, 6].Value = newItem.Quantity;
                                        sheet.Cells[targetRow, 7].Value = newItem.UnitPrice;
                                        sheet.Cells[targetRow, 8].Formula = $"=ROUND(F{targetRow}*G{targetRow}, 2)";

                                        cabModified = true;
                                        updatedCompCount++;
                                    }
                                }
                            }
                        }

                        // 3. 元器件行紧凑排版与重排序 (第一行主元器件固定不动，空行整齐沉底，30列全数据联动)
                        // 条件：若勾选了调整排序，或者执行了相同元件合并，则触发紧凑规整与排版
                        bool needReorder = options != null && options.UpdateBomOrder;
                        if ((needReorder || shouldMergeSameBom) && cabModified && compEndRow >= compStartRow)
                        {
                            try
                            {
                                int currentCompRows = compEndRow - compStartRow + 1;
                                // 只有当行数大于 1 时才有排版必要 (第一行主开关保持不动，仅对从第 2 行起的元件整理)
                                if (currentCompRows > 1)
                                {
                                    // 一次性读取整整 30 列数据矩阵 (从 A 列至 AD 列 CAD 句柄，规则 7)
                                    dynamic fullCompRange = sheet.Range[$"A{compStartRow}:AD{compEndRow}"];
                                    object[,] fullMatrix = (object[,])fullCompRange.Value2;

                                    // 收集从第 2 行起的有效元器件 (第 1 行主器件保留原位，不参与排序)
                                    var validTailRows = new List<(double sortId, object[] rowCells)>();

                                    // 从第 2 行遍历至末尾
                                    for (int r = 2; r <= currentCompRows; r++)
                                    {
                                        string cName = Convert.ToString(fullMatrix[r, 2])?.Trim() ?? "";
                                        string cModel = Convert.ToString(fullMatrix[r, 3])?.Trim() ?? "";
                                        // 过滤合并或删除后的空白行
                                        if (string.IsNullOrWhiteSpace(cName) && string.IsNullOrWhiteSpace(cModel)) continue;

                                        // 完整提取该行的全部 30 列数据 (保证 AD 列 CAD 句柄及右侧所有属性完整跟随平移)
                                        object[] rowCells = new object[30];
                                        for (int c = 0; c < 30; c++)
                                        {
                                            rowCells[c] = fullMatrix[r, c + 1];
                                        }

                                        // 确定该元件的排序权重 (勾选排序取分布表 SortId，未勾选排序保持原相对行号)
                                        double itemSortId = needReorder ? 999999 + r : r;
                                        if (needReorder && expectedCompItems != null)
                                        {
                                            // 匹配分布表排序 ID
                                            var m = expectedCompItems.FirstOrDefault(it =>
                                                string.Equals(it.Name, cName, StringComparison.OrdinalIgnoreCase) &&
                                                string.Equals(it.Model, cModel, StringComparison.OrdinalIgnoreCase));
                                            if (m != null && m.SortId > 0) itemSortId = m.SortId;
                                        }

                                        validTailRows.Add((itemSortId, rowCells));
                                    }

                                    // 排序整理：勾选排序按 SortId 升序，未勾选排序保持原出现顺序，空行自动沉底
                                    var sortedTailRows = validTailRows.OrderBy(v => v.sortId).ToList();

                                    // 重构 30 列大矩阵写回
                                    object[,] reorderedMatrix = new object[currentCompRows, 30];

                                    // 1. 第 1 行元器件原位保留，整行 30 列数据纹丝不动 (包含 AD 列 CAD 句柄)
                                    for (int c = 0; c < 30; c++)
                                    {
                                        reorderedMatrix[0, c] = fullMatrix[1, c + 1];
                                    }

                                    // 2. 从第 2 行开始填入排序与紧凑排列后的有效元器件
                                    int availableTailRows = currentCompRows - 1;
                                    for (int i = 0; i < availableTailRows; i++)
                                    {
                                        int destR = i + 1; // 目标行索引 (1..currentCompRows-1)
                                        if (i < sortedTailRows.Count)
                                        {
                                            // 写入有效元器件的 30 列数据
                                            var rowData = sortedTailRows[i].rowCells;
                                            for (int c = 0; c < 30; c++)
                                            {
                                                reorderedMatrix[destR, c] = rowData[c];
                                            }
                                        }
                                        else
                                        {
                                            // 超出有效元件数的剩余行变为空白行 (清空整行 30 列，空行整齐沉底)
                                            for (int c = 0; c < 30; c++)
                                            {
                                                reorderedMatrix[destR, c] = "";
                                            }
                                        }
                                    }

                                    // 一次性批量写回 30 列大矩阵 (规则 7)
                                    fullCompRange.Value2 = reorderedMatrix;

                                    // 3. 统一批量重新灌入自适应动态序号公式与合价公式 (范围批量赋值，消除数十次 COM 细碎往返)
                                    object[,] formulasColA = new object[currentCompRows, 1];
                                    object[,] formulasColH = new object[currentCompRows, 1];

                                    // 刷新第 1 行序号与合价公式
                                    formulasColA[0, 0] = $"=ROW()-ROW(A${headerRow})";
                                    formulasColH[0, 0] = $"=ROUND(F{compStartRow}*G{compStartRow}, 2)";

                                    // 刷新第 2 行起有效元器件行的序号与合价公式
                                    for (int i = 0; i < sortedTailRows.Count; i++)
                                    {
                                        int relIdx = 1 + i;
                                        int realR = compStartRow + relIdx;
                                        formulasColA[relIdx, 0] = $"=ROW()-ROW(A${headerRow})";
                                        formulasColH[relIdx, 0] = $"=ROUND(F{realR}*G{realR}, 2)";
                                    }

                                    // 清空多余沉底空行的公式，保持表格纯净
                                    for (int relIdx = 1 + sortedTailRows.Count; relIdx < currentCompRows; relIdx++)
                                    {
                                        formulasColA[relIdx, 0] = "";
                                        formulasColH[relIdx, 0] = "";
                                    }

                                    // 一次性范围批量写入 A 列序号公式与 H 列合价公式 (规则 7)
                                    sheet.Range[$"A{compStartRow}:A{compEndRow}"].Formula = formulasColA;
                                    sheet.Range[$"H{compStartRow}:H{compEndRow}"].Formula = formulasColH;
                                }
                            }
                            catch (Exception ex)
                            {
                                // 记录排序与规整异常日志
                                LogHelper.WriteLog($"[分布调价] 箱柜 {cabNo} 调整元件排序/合并规整异常: {ex.Message}");
                            }
                        }

                        // 若该箱柜有元器件发生更新或新增，重新联动刷新公式 (规则 6 & 8)
                        if (cabModified)
                        {
                            // 重新刷新小计行求和公式 (规则 6: Cab_Subsum_k)
                            sheet.Cells[subsumRow, 8].Formula = $"=SUM(H{compStartRow}:H{compEndRow})";

                            // 重新联动总计行公式 (规则 6: Cab_Tolsum_k)
                            if (anchor.Tolsum != null)
                            {
                                int tolsumRow = Convert.ToInt32(anchor.Tolsum.Row);
                                // 若总计行物理位置大于小计行，联动求和
                                if (tolsumRow > subsumRow)
                                {
                                    sheet.Cells[tolsumRow, 8].Formula = $"=SUM(H{subsumRow}:H{tolsumRow - 1})";
                                }
                            }

                            sheetModified = true;
                            updatedCabCount++;

                            // 提取修改后该箱柜完整的 30 列公式矩阵快照 (覆盖 A 列至 AD 列 CAD 句柄，规则 7)
                            object[,] newFormulaMatrix = (object[,])((object[,])sheet.Range[$"A{compStartRow}:AD{compEndRow}"].Formula).Clone();

                            // 收集该箱柜的分布调价差量切片
                            undoSlices.Add(new DistributionCabinetSlice
                            {
                                // 记录工作表名称
                                SheetName = sheetName,
                                // 记录箱柜柜号
                                CabinetNo = cabNo,
                                // 记录元器件起始行
                                CompStartRow = compStartRow,
                                // 记录修改前的元器件终止行
                                OrigCompEndRow = origCompEndRow,
                                // 记录修改前的小计行
                                OrigSubsumRow = origSubsumRow,
                                // 记录插入的新行数
                                InsertedRowCount = insertedRowCount,
                                // 记录插入行的起始物理行号
                                InsertedRowStart = insertedRowStart,
                                // 记录修改后的元器件终止行
                                NewCompEndRow = compEndRow,
                                // 记录修改后的小计行
                                NewSubsumRow = subsumRow,
                                // 记录修改前的完整公式与数值快照
                                OldFormulas = origFormulaMatrix,
                                // 记录修改后的完整公式与数值快照
                                NewFormulas = newFormulaMatrix
                            });
                        }
                    }

                    if (sheetModified)
                    {
                        // 严格执行规则 8：闭环自愈校准工作表超链接与 4 个定义名称
                        Tool.FixAndFillCabinetNamesForSheet(sheet);
                        updatedSheetCount++;
                    }
                }

                // 发送最终重算与校准进度通知 (95%)
                progressCallback?.Invoke(95, "正在执行公式重算与最终工作表自愈校准...");

                // 统一触发一次全簿公式重新计算，刷新所有小计、总计与跨表联动
                try
                {
                    app.Calculate();
                }
                catch { }

                // 恢复 Excel 原始计算模式
                if (prevCalc != null)
                {
                    try { app.Calculation = prevCalc; } catch { }
                }

                // 恢复事件监听响应
                try { app.EnableEvents = prevEvents; } catch { }

                // 恢复屏幕刷新与提示
                app.ScreenUpdating = true;
                app.DisplayAlerts = true;

                // 发送 100% 完成通知
                progressCallback?.Invoke(100, "分布调价同步完成！");

                // 若有箱柜发生数据更新，打包为分布调价专有可撤销命令推入撤销栈
                if (undoSlices.Count > 0 && updatedCabCount > 0)
                {
                    // 构建整簿分布调价可逆命令
                    var distCmd = new DistributionAdjustPriceCommand($"分布调价同步 ({updatedCabCount}台箱柜)", undoSlices);
                    // 压入全局撤销管理中心
                    UndoRedoManager.Instance.PushCommand(distCmd);
                }

                result.Success = true;
                result.UpdatedSheetCount = updatedSheetCount;
                result.UpdatedCabinetCount = updatedCabCount;
                result.UpdatedComponentCount = updatedCompCount;
                result.Message = $"分布调价同步成功！共更新 {updatedSheetCount} 个分类表，{updatedCabCount} 台箱柜，{updatedCompCount} 项元器件（可随时按 Ctrl+Z 撤销）。";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"分布调价更新执行失败: {ex.Message}";
                LogHelper.WriteLog($"[分布调价] 反向更新异常: {ex.Message}\r\n{ex.StackTrace}");
            }
            finally
            {
                try
                {
                    dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                    if (app != null)
                    {
                        // 确保计算模式恢复
                        if (prevCalc != null)
                        {
                            try { app.Calculation = prevCalc; } catch { }
                        }
                        // 确保事件恢复
                        try { app.EnableEvents = prevEvents; } catch { }
                        // 确保屏幕刷新恢复
                        app.ScreenUpdating = true;
                        app.DisplayAlerts = true;
                    }
                }
                catch { }
            }

            return result;
        }
    }
}
