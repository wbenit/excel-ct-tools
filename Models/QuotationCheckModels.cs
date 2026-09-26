using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 报价缺陷与合规性问题严重等级枚举
    /// </summary>
    public enum QuotationIssueSeverity
    {
        // 致命错误：如价格倒挂、公式错误值、严重漏价，可能导致直接亏损或计算崩溃
        Error,
        // 警告：如公式被手填写死、同型号折扣不一致、计费区空行等
        Warning,
        // 提示：如 AA/AB 列 CAD 图纸与目录名为空等工程参数提醒
        Info
    }

    /// <summary>
    /// 报价智能核验问题分类枚举常量定义
    /// </summary>
    public static class QuotationIssueCategories
    {
        // 同型号价格/折扣/加价系数冲突
        public const string PriceConflict = "PriceConflict";
        // 销售单价/总价/成本单价公式被手写数字覆写破坏
        public const string FormulaBroken = "FormulaBroken";
        // 单元格出现 #REF!、#VALUE!、#DIV/0!、#N/A 等严重公式错误
        public const string FormulaError = "FormulaError";
        // 关键面价缺失（有型号但面价为空或为0）
        public const string MissingPrice = "MissingPrice";
        // 采购折扣缺失（折扣为空）
        public const string MissingDiscount = "MissingDiscount";
        // 元件数量缺失（数量为空或小于等于0）
        public const string MissingQuantity = "MissingQuantity";
        // AA 列 CAD 图块/图纸名称缺失
        public const string MissingBlockDwg = "MissingBlockDwg";
        // AB 列 CAD 图块目录/分类名称缺失
        public const string MissingBlockDir = "MissingBlockDir";
        // AA 与 AB 列 CAD 工程联动参数均缺失
        public const string MissingBothCadParams = "MissingBothCadParams";
        // 价格倒挂（销售单价小于成本单价，出现毛利负值亏损）
        public const string PriceInverted = "PriceInverted";
        // 折扣量级超限（折扣大于 1.0 或为负数）
        public const string DiscountExceeded = "DiscountExceeded";
        // 报价系数异常（加价系数小于 1.0 或大于 5.0）
        public const string FactorAbnormal = "FactorAbnormal";
        // 计费区域存在空行或关键计费项金额为 0
        public const string FeeEmptyOrBlankRow = "FeeEmptyOrBlankRow";
        // 汇总行 Cab_Sum 与明细总计 Cab_Tolsum 金额不一致
        public const string SumMismatch = "SumMismatch";
        // 纯数值被存为了文本型格式 (容易导致 Excel SUM 漏算)
        public const string TextNumber = "TextNumber";
    }

    /// <summary>
    /// 报价核验单条明细问题记录实体
    /// </summary>
    public class QuotationCheckItem
    {
        // 唯一标识键（如 分类1_R15_G）
        public string Id { get; set; } = string.Empty;

        // 所属 Excel 工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 所属箱柜名称（如 1#动力配电箱）
        public string CabinetName { get; set; } = string.Empty;

        // 箱柜序号索引
        public int CabinetIndex { get; set; }

        // Excel 物理行号 (1-based)
        public int Row { get; set; }

        // 元器件名称 (B列)
        public string ComponentName { get; set; } = string.Empty;

        // 型号规格 (C列)
        public string Model { get; set; } = string.Empty;

        // 品牌/生产厂家 (D列)
        public string Brand { get; set; } = string.Empty;

        // 数量 (F列)
        public decimal Quantity { get; set; }

        // 销售单价 (G列当前值)
        public decimal SaleUnitPrice { get; set; }

        // 销售总价 (H列当前值)
        public decimal SaleTotalPrice { get; set; }

        // 成本单价 (J列当前值)
        public decimal CostUnitPrice { get; set; }

        // 报价加价系数 (L列当前值)
        public decimal MarkupFactor { get; set; }

        // 面价/表价 (M列当前值)
        public decimal BasePrice { get; set; }

        // 采购折扣 (N列当前值)
        public decimal Discount { get; set; }

        // AA 列图块/DWG 图纸名称
        public string BlockDwg { get; set; } = string.Empty;

        // AB 列图块所属目录分类
        public string BlockDir { get; set; } = string.Empty;

        // 建议定位选中的 Excel 列标 (如 "G", "M", "AA", "AB")
        public string TargetCol { get; set; } = "C";

        // 问题类别编码 (参见 QuotationIssueCategories)
        public string Category { get; set; } = string.Empty;

        // 严重等级："error" (红), "warning" (黄), "info" (蓝)
        public string Severity { get; set; } = "warning";

        // 简短问题标题 (如 "同型号价格不一致", "销售单价公式丢失")
        public string Title { get; set; } = string.Empty;

        // 详细说明与现状描述
        public string Detail { get; set; } = string.Empty;

        // 建议修复方案 (如 "一键恢复 ROUND 公式", "补录面价")
        public string SuggestedFix { get; set; } = string.Empty;

        // 是否具备一键自动修复能力
        public bool CanAutoFix { get; set; }

        // 自动修复的动作指令 (如 "fixFormula", "cleanTextNumber", "alignModel")
        public string FixAction { get; set; } = string.Empty;
    }

    /// <summary>
    /// 同型号价格变体与分布记录
    /// </summary>
    public class ModelPriceVariant
    {
        // 面价 / 基准价
        public decimal BasePrice { get; set; }

        // 采购折扣
        public decimal Discount { get; set; }

        // 报价加价系数
        public decimal MarkupFactor { get; set; }

        // 销售单价推导值
        public decimal DerivedSalePrice { get; set; }

        // 该组价格组合在项目中的出现频次
        public int Occurrences { get; set; }

        // 首次出现或代表性的样本位置描述 (如 "分类1 > 1#柜 > 行12")
        public string SampleLocation { get; set; } = string.Empty;

        // 样本所在工作表名
        public string SampleSheet { get; set; } = string.Empty;

        // 样本物理行号
        public int SampleRow { get; set; }
    }

    /// <summary>
    /// 同型号聚合冲突分析模型
    /// </summary>
    public class ModelConflictGroup
    {
        // 复合对齐键 (格式如 "NXM-125S/3300 100A | 正泰")
        public string ModelKey { get; set; } = string.Empty;

        // 型号规格
        public string Model { get; set; } = string.Empty;

        // 生产厂家/品牌
        public string Brand { get; set; } = string.Empty;

        // 该型号在项目中被使用的总行数
        public int TotalRows { get; set; }

        // 存在的不同价格版本数量 (若大于1则构成冲突)
        public int VariantCount { get; set; }

        // 所有不同的价格组合变体列表
        public List<ModelPriceVariant> Variants { get; set; } = new();

        // 推荐采用的主流基准变体 (频次最高或最新)
        public ModelPriceVariant? RecommendedVariant { get; set; }
    }

    /// <summary>
    /// 报价健康体检总览概括统计
    /// </summary>
    public class QuotationCheckSummary
    {
        // 健康得分 (100分制，发现严重错误与警告逐级扣分)
        public int HealthScore { get; set; } = 100;

        // 扫描发现的总问题数
        public int TotalIssues { get; set; }

        // 致命错误数量 (红色警报)
        public int ErrorCount { get; set; }

        // 需关注警告数量 (黄色警报)
        public int WarningCount { get; set; }

        // 工程参数提醒数量 (蓝色提示，如 AA/AB 为空)
        public int InfoCount { get; set; }

        // 同型号价格冲突条目数
        public int ConflictCount { get; set; }

        // 公式损坏条目数
        public int FormulaBrokenCount { get; set; }

        // 漏填面价/漏项条目数
        public int MissingPriceCount { get; set; }

        // CAD 图纸/图块参数缺失条目数 (AA/AB 列)
        public int CadMissingCount { get; set; }

        // 价格倒挂条目数 (销售单价小于成本单价)
        public int InvertedPriceCount { get; set; }

        // 扫描覆盖的有效箱柜台数
        public int TotalCabinets { get; set; }

        // 扫描覆盖的有效元器件总行数
        public int TotalComponents { get; set; }
    }

    /// <summary>
    /// 报价智能核验全量返回结果载荷实体
    /// </summary>
    public class QuotationCheckResult
    {
        // 是否扫描成功
        public bool Success { get; set; } = true;

        // 状态消息反馈
        public string Message { get; set; } = string.Empty;

        // 体检概览总评统计
        public QuotationCheckSummary Summary { get; set; } = new();

        // 检查出的全量问题清单
        public List<QuotationCheckItem> Issues { get; set; } = new();

        // 同型号冲突聚类清单 (供可视化比对与一键批量对齐)
        public List<ModelConflictGroup> ConflictGroups { get; set; } = new();

        // 扫描涉及的所有有效分类表名称列表
        public List<string> SheetNames { get; set; } = new();
    }

    /// <summary>
    /// 同型号价格批量对齐请求参数模型
    /// </summary>
    public class AlignModelPriceRequest
    {
        // 目标型号复合键
        public string ModelKey { get; set; } = string.Empty;

        // 规格型号
        public string Model { get; set; } = string.Empty;

        // 品牌/厂家
        public string Brand { get; set; } = string.Empty;

        // 对齐目标面价 (M 列)
        public decimal TargetBasePrice { get; set; }

        // 对齐目标折扣 (N 列)
        public decimal TargetDiscount { get; set; }

        // 对齐目标加价系数 (L 列)
        public decimal TargetMarkupFactor { get; set; }

        // 是否同时重置修复受影响行的计算公式
        public bool ResetFormulas { get; set; } = true;
    }

    /// <summary>
    /// 批量公式修复请求参数模型
    /// </summary>
    public class FixFormulasRequest
    {
        // 需要修复的问题行 ID 列表 (为空时表示修复全量公式被破坏的行)
        public List<string> IssueIds { get; set; } = new();
    }
}
