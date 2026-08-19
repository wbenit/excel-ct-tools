using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 箱柜顶部固定属性实体对象
    /// </summary>
    public class CabinetHeader
    {
        // 柜号 (如: 箱柜1, 箱柜2)
        public string CabinetNo { get; set; } = string.Empty;

        // 箱柜型号
        public string Model { get; set; } = string.Empty;

        // 箱柜名称
        public string Name { get; set; } = string.Empty;

        // 备注信息
        public string Remark { get; set; } = string.Empty;

        // 安装方式 (如: 暗装, 明装, 落地式)
        public string InstallMode { get; set; } = string.Empty;

        // 箱柜数量
        public int Quantity { get; set; } = 1;

        // 箱柜尺寸 (宽*高*深)
        public string Dimensions { get; set; } = string.Empty;

        // 箱柜类别 / 所属分类表名称 (如: 1F强电井)
        public string Category { get; set; } = string.Empty;

        // 图纸编号
        public string DrawingNo { get; set; } = string.Empty;

        // CAD 图元句柄/坐标范围
        public List<string> MinMaxPoints { get; set; } = new List<string>();
    }

    /// <summary>
    /// 新建箱柜关键行号与序号返回结果实体
    /// 替代可空 ValueTuple，防止在 DLR 动态调用上下文中发生 HasValue 拆箱异常
    /// </summary>
    public class CabinetCreatedInfo
    {
        // 新箱柜序号 K
        public int CabinetK { get; set; }
        // 顶部汇总行物理行号
        public int SumRow { get; set; }
        // 底部明细箱柜信息行物理行号
        public int DetRow { get; set; }
        // 底部明细小计行物理行号
        public int SubsumRow { get; set; }
        // 底部明细总计行物理行号
        public int TolsumRow { get; set; }
    }

    /// <summary>
    /// Excel 活动运行环境上下文实体对象
    /// 封装当前 Application、ActiveWorkbook 及 ActiveSheet 的 COM 句柄
    /// </summary>
    public class ExcelContext
    {
        // Excel 应用程序 COM 实例
        public dynamic App { get; set; }

        // 活动工作簿 COM 实例
        public dynamic Wb { get; set; }

        // 目标工作表 COM 实例
        public dynamic Sheet { get; set; }

        // 构造函数初始化上下文对象
        public ExcelContext(dynamic app, dynamic wb, dynamic sheet)
        {
            App = app;
            Wb = wb;
            Sheet = sheet;
        }
    }


    /// <summary>
    /// 元器件明细项实体对象
    /// </summary>
    public class ComponentItem
    {
        // 相对行序号
        public int Index { get; set; }

        // 元件名称
        public string Name { get; set; } = string.Empty;

        // 型号规格
        public string Specification { get; set; } = string.Empty;

        // 生产厂家
        public string Manufacturer { get; set; } = string.Empty;

        // 计量单位 (如: 台、个)
        public string Unit { get; set; } = string.Empty;

        // CAD 图元句柄
        public string Handle { get; set; } = string.Empty;

        // 数量
        public decimal Quantity { get; set; }

        // 销售单价
        public decimal UnitPrice { get; set; }

        // 销售总价 (计算属性: 数量 * 单价)
        public decimal TotalPrice => Quantity * UnitPrice;

        // 成本单价
        public decimal CostUnitPrice { get; set; }

        // 成本总价 (计算属性: 数量 * 成本单价)
        public decimal CostTotalPrice => Quantity * CostUnitPrice;

        // 备注
        public string Remark { get; set; } = string.Empty;

        // 表价
        public decimal ListPrice { get; set; }

        // 折扣/成套费
        public decimal SetFee { get; set; }

        // 元件类别
        public string Category { get; set; } = string.Empty;

        // 物料编码
        public string MaterialCode { get; set; } = string.Empty;

        // 产地
        public string Origin { get; set; } = string.Empty;

        // 加工费
        public decimal ProcessingFee { get; set; }

        // 订货号
        public string OrderNumber { get; set; } = string.Empty;

        // 人工费
        public decimal LaborFee { get; set; }

        // 辅料费
        public decimal AccessoriesFee { get; set; }
    }

    /// <summary>
    /// 费用项计算模式枚举
    /// </summary>
    public enum CalculationMode
    {
        // 手工输入固定金额
        FixedValue,

        // 按小计或基数的百分比计算
        Percentage,

        // 动态公式/脚本计算
        CustomFormula
    }

    /// <summary>
    /// 弹性费用项数据对象
    /// </summary>
    public class BillingFeeItem
    {
        // 费用项唯一标识 Key (如: Subtotal, LaborFee, Tax)
        public string FeeKey { get; set; } = string.Empty;

        // 界面显示名称 (如: 小计、管理费、利润)
        public string DisplayName { get; set; } = string.Empty;

        // 计算模式
        public CalculationMode Mode { get; set; } = CalculationMode.FixedValue;

        // 费率或系数 (如 0.12 表示 12%)
        public decimal Rate { get; set; }

        // 最终计算结果金额
        public decimal Amount { get; set; }

        // Excel 对应的公式字符串 (如 =ROUND(H26*0.12, 2))
        public string ExcelFormula { get; set; } = string.Empty;

        // 是否在表格底部展示
        public bool IsVisible { get; set; } = true;
    }

    /// <summary>
    /// 抽象计费策略接口 (策略模式)
    /// </summary>
    public interface IBillingStrategy
    {
        // 策略名称
        string StrategyName { get; }

        // 执行调费计算并返回结果费用项集合
        List<BillingFeeItem> Calculate(List<ComponentItem> components, List<BillingFeeItem> feeItems);
    }

    /// <summary>
    /// 公式组明细行模板定义对象 (对应“公式法调费”界面明细网格的每一行)
    /// </summary>
    public class FormulaFeeRowDefinition
    {
        // 序号标记 (如: "[序号]", "5", "总计")
        public string IndexTag { get; set; } = "[序号]";

        // 费用项名称 (如: "小计", "综合费(管理、税金、利润等)", "单台合计")
        public string Name { get; set; } = string.Empty;

        // 型号规格 (对应C列，为空调价时不修改项目内容)
        public string Specification { get; set; } = string.Empty;

        // 生产厂家 (对应D列)
        public string Manufacturer { get; set; } = string.Empty;

        // 单位 (对应E列，如: "台", "公斤")
        public string Unit { get; set; } = string.Empty;

        // 数量或数量公式 (对应F列)
        public string QuantityFormula { get; set; } = string.Empty;

        // 单价或单价公式 (对应G列)
        public string UnitPriceFormula { get; set; } = string.Empty;

        // 总价或总价公式 (对应H列，如: "[总价小计]", "=ROUND(H2*0.2, 2)")
        public string TotalPriceFormula { get; set; } = string.Empty;

        // 成本单价公式 (对应J列)
        public string CostUnitPriceFormula { get; set; } = string.Empty;

        // 成本总价公式 (对应K列，如: "[成本总价]")
        public string CostTotalPriceFormula { get; set; } = string.Empty;
    }

    /// <summary>
    /// 基于 Excel 公式组的调费策略实现对象
    /// </summary>
    public class FormulaBillingGroupStrategy : IBillingStrategy
    {
        // 策略 ID
        public string StrategyId { get; set; } = Guid.NewGuid().ToString();

        // 策略名称 (如: "简易费用公式", "多费用公式", "国网报价费用公式")
        public string StrategyName { get; set; } = "简易费用公式";

        // 是否为默认系统公式组
        public bool IsDefault { get; set; } = false;

        // 公式组包含的模板行定义列表
        public List<FormulaFeeRowDefinition> RowDefinitions { get; set; } = new List<FormulaFeeRowDefinition>();

        // 执行策略计算逻辑
        public List<BillingFeeItem> Calculate(List<ComponentItem> components, List<BillingFeeItem> feeItems)
        {
            // 依据组件与公式行模板调用转换引擎计算
            return FormulaEngine.ProcessFormulaGroup(components, RowDefinitions);
        }
    }

    /// <summary>
    /// 动态公式重映射与计算引擎
    /// </summary>
    public static class FormulaEngine
    {
        /// <summary>
        /// 处理公式组解析，根据组件计算小计并转换相对行号
        /// </summary>
        public static List<BillingFeeItem> ProcessFormulaGroup(List<ComponentItem> components, List<FormulaFeeRowDefinition> rowDefs)
        {
            // 创建计算结果列表
            var results = new List<BillingFeeItem>();

            // 如果模板行定义为空直接返回
            if (rowDefs == null || rowDefs.Count == 0) return results;

            // 计算元器件销售总价小计
            decimal subtotal = components?.Sum(c => c.TotalPrice) ?? 0m;

            // 遍历公式行定义
            foreach (var rowDef in rowDefs)
            {
                // 实例化新的费用项
                var item = new BillingFeeItem
                {
                    FeeKey = rowDef.Name,
                    DisplayName = rowDef.Name,
                    ExcelFormula = rowDef.TotalPriceFormula
                };

                // 判断是否为小计占位符
                if (rowDef.TotalPriceFormula.Contains("[总价小计]"))
                {
                    // 赋值小计金额
                    item.Amount = subtotal;
                    item.Mode = CalculationMode.FixedValue;
                }
                else
                {
                    // 标记为公式类型模式
                    item.Mode = CalculationMode.CustomFormula;
                }

                // 添加至结果集合
                results.Add(item);
            }

            // 返回转换后的费用项列表
            return results;
        }

        /// <summary>
        /// 将模板公式中的相对模板行号重映射转换为真实 Excel 行号公式
        /// </summary>
        /// <param name="rawFormula">原始模板公式 (如 =ROUND(H2*0.2, 2))</param>
        /// <param name="templateStartRow">模板起始行号 (通常为 1)</param>
        /// <param name="excelStartRow">Excel 真实落地起始行号</param>
        /// <param name="componentStartRow">元器件起始行</param>
        /// <param name="componentEndRow">元器件结束行</param>
        /// <returns>转换后的真实 Excel 公式 (如 =ROUND(H26*0.2, 2))</returns>
        public static string ConvertToExcelFormula(string rawFormula, int templateStartRow, int excelStartRow, int componentStartRow, int componentEndRow)
        {
            // 空字符串直接返回
            if (string.IsNullOrWhiteSpace(rawFormula)) return string.Empty;

            // 替换 [总价小计] 动态范围公式
            if (rawFormula.Contains("[总价小计]"))
            {
                // 生成自适应动态求和公式并保留两位小数 (如 =ROUND(SUM(H8:INDEX(H:H,ROW()-1)),2))
                return $"=ROUND(SUM(H{componentStartRow}:INDEX(H:H,ROW()-1)),2)";
            }

            // 替换 [成本总价] 或 [成本总价小计] 动态范围公式
            if (rawFormula.Contains("[成本总价]") || rawFormula.Contains("[成本总价小计]"))
            {
                // 生成自适应成本动态求和公式并保留两位小数 (如 =ROUND(SUM(K8:INDEX(K:K,ROW()-1)),2))
                return $"=ROUND(SUM(K{componentStartRow}:INDEX(K:K,ROW()-1)),2)";
            }

            // 若不是以等号开头的表达式，原样返回文本
            if (!rawFormula.StartsWith("=")) return rawFormula;

            // 正则匹配提取公式中的行号引用并动态偏移重映射
            string resultFormula = System.Text.RegularExpressions.Regex.Replace(
                rawFormula,
                @"([A-Z]+)(\d+)",
                match =>
                {
                    // 提取列名 (如 H)
                    string col = match.Groups[1].Value;

                    // 提取模板相对行号
                    if (int.TryParse(match.Groups[2].Value, out int tmplRow))
                    {
                        // 计算基于真实落地的真实 Excel 行号
                        int realRow = excelStartRow + (tmplRow - templateStartRow);

                        // 返回重映射后的单元格标识 (如 H26)
                        return $"{col}{realRow}";
                    }

                    // 无法解析时保持原样
                    return match.Value;
                }
            );

            // 返回替换完毕的真实 Excel 公式
            return resultFormula;
        }
    }

    /// <summary>
    /// 箱柜面向对象主实体 (聚合根)
    /// </summary>
    public class CabinetObject
    {
        // 箱柜对应在 Excel 工作表中的定义名称序号 (如 1, 2, 3)
        public int CabinetIndex { get; set; }

        // 明细表头锚点在 Excel 中的真实绝对行号 (Cab_Det_k.Row)
        public int DetAnchorRow { get; set; }

        // 汇总小计锚点在 Excel 中的真实绝对行号 (Cab_Sum_k.Row)
        public int SumAnchorRow { get; set; }

        // 明细小计锚点在 Excel 中的真实绝对行号 (Cab_Subsum_k.Row)
        public int SubsumAnchorRow { get; set; }

        // 明细总计锚点在 Excel 中的真实绝对行号 (Cab_Tolsum_k.Row)
        public int TolsumAnchorRow { get; set; }

        // 顶部固定信息对象
        public CabinetHeader Header { get; set; } = new CabinetHeader();

        // 中间元器件明细列表
        public List<ComponentItem> Components { get; set; } = new List<ComponentItem>();

        // 底部弹性计费项列表
        public List<BillingFeeItem> BillingFeeItems { get; set; } = new List<BillingFeeItem>();

        // 当前箱柜绑定的计费策略对象 (默认简易公式组)
        public IBillingStrategy BillingStrategy { get; set; } = new FormulaBillingGroupStrategy();

        // 重新触发计费策略计算
        public void Recalculate()
        {
            // 执行策略计算更新费用项
            if (BillingStrategy != null)
            {
                BillingFeeItems = BillingStrategy.Calculate(Components, BillingFeeItems);
            }
        }
    }

    /// <summary>
    /// 单张 Excel 工作表及其包含的所有箱柜对象容器模型
    /// </summary>
    public class CabinetSheetObject
    {
        // 工作表名称 (如 "分类1", "配电房A")
        public string SheetName { get; set; } = string.Empty;

        // 当前工作表中包含的所有箱柜对象集合
        public List<CabinetObject> Cabinets { get; set; } = new List<CabinetObject>();
    }

    /// <summary>
    /// 标准箱柜面向对象实体工厂
    /// </summary>
    public static class CabinetObjectFactory
    {
        /// <summary>
        /// 创建一个干净的标准初始箱柜对象 (预置空插槽与默认简易公式组策略)
        /// </summary>
        /// <param name="cabinetIndex">箱柜序号 (如 1, 2, 3)</param>
        /// <param name="slotCount">预置空白元件插槽行数 (默认 32)</param>
        /// <returns>全新的干净 CabinetObject 实体</returns>
        public static CabinetObject CreateCleanCabinet(int cabinetIndex, int slotCount = 32)
        {
            // 实例化全新的箱柜对象
            var cab = new CabinetObject
            {
                CabinetIndex = cabinetIndex
            };

            // 设置表头默认属性
            cab.Header.CabinetNo = $"箱柜{cabinetIndex}";
            cab.Header.Name = $"箱柜{cabinetIndex}";
            cab.Header.Model = string.Empty;
            cab.Header.Remark = string.Empty;

            // 循环生成 slotCount 个干净空白元器件插槽对象
            for (int i = 1; i <= slotCount; i++)
            {
                // 创建空白元器件实体
                cab.Components.Add(new ComponentItem
                {
                    Index = i,
                    Name = string.Empty,
                    Specification = string.Empty,
                    Manufacturer = string.Empty,
                    Unit = string.Empty,
                    Quantity = 0m,
                    UnitPrice = 0m,
                    CostUnitPrice = 0m
                });
            }

            // 创建默认【简易费用公式组策略】
            var defaultStrategy = new FormulaBillingGroupStrategy
            {
                StrategyName = "简易费用公式",
                IsDefault = true,
                RowDefinitions = new List<FormulaFeeRowDefinition>
                {
                    // 1. 小计
                    new FormulaFeeRowDefinition
                    {
                        IndexTag = "[序号]",
                        Name = "小计",
                        TotalPriceFormula = "[总价小计]",
                        CostTotalPriceFormula = "[成本总价]"
                    },
                    // 2. 综合费 (管理、税金、利润等)
                    new FormulaFeeRowDefinition
                    {
                        IndexTag = "[序号]",
                        Name = "综合费(管理、税金、利润等)",
                        TotalPriceFormula = "=ROUND(H2*0.2, 2)"
                    },
                    // 3. 单台合计
                    new FormulaFeeRowDefinition
                    {
                        IndexTag = "[序号]",
                        Name = "单台合计",
                        TotalPriceFormula = "=ROUND(SUM(H2:H3), 2)"
                    }
                }
            };

            // 绑定至箱柜策略对象
            cab.BillingStrategy = defaultStrategy;

            // 重新计算费用项
            cab.Recalculate();

            // 返回干净的初始实体对象
            return cab;
        }
    }

    /// <summary>
    /// 箱柜在 Excel 工作表中的 4 个定义名称 Range 锚点强类型模型 (Cab_Det, Cab_Sum, Cab_Subsum, Cab_Tolsum)
    /// </summary>
    public class CabinetAnchorModel
    {
        // 箱柜信息行 Range (Cab_Det)
        public dynamic Det { get; set; }

        // 顶部汇总行 Range (Cab_Sum)
        public dynamic Sum { get; set; }

        // 底部明细小计行 Range (Cab_Subsum)
        public dynamic Subsum { get; set; }

        // 底部明细总计行 Range (Cab_Tolsum)
        public dynamic Tolsum { get; set; }

        // 构造函数
        public CabinetAnchorModel(dynamic det = null, dynamic sum = null, dynamic subsum = null, dynamic tolsum = null)
        {
            Det = det;
            Sum = sum;
            Subsum = subsum;
            Tolsum = tolsum;
        }
    }
}


