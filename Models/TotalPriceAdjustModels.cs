using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 分类工作表基本信息实体 (用于总价一键调整向导多选分类表)
    /// </summary>
    public class TotalPriceCategorySheetItem
    {
        // 工作表实际显示名称 (如“高压变电站”)
        public string SheetName { get; set; } = string.Empty;

        // 该分类表包含的有效箱柜台数
        public int CabinetCount { get; set; }

        // 该分类表当前的销售总金额
        public decimal SalesTotalPrice { get; set; }

        // 该分类表当前的成本总金额
        public decimal CostTotalPrice { get; set; }

        // 是否被勾选参与本次总价调整 (默认勾选)
        public bool IsSelected { get; set; } = true;
    }

    /// <summary>
    /// 动态分摊标签实体 (依据 Q 列标签及 B 列兜底聚合提取)
    /// </summary>
    public class TotalPriceAdjustTagItem
    {
        // 标签名称 (例如：“元件”、“生产工时”、“综合管理费”、“包装运输费”、“其他费用”)
        public string Tag { get; set; } = string.Empty;

        // 标签类别属性: "component" (元器件) 或 "fee" (计费区域费用)
        public string ItemType { get; set; } = "fee";

        // 当前该标签在所有选定箱柜中的累计总金额
        public decimal CurrentAmount { get; set; }

        // 该标签在原项目总金额中的占比百分比 (0~100)
        public decimal OriginalRatio { get; set; }

        // 用户配置的分摊比例 (0.0~1.0，元器件默认 1.0，其余默认 0.0，总和严格等于 1.0)
        public decimal AllocationRatio { get; set; }

        // 是否勾选参与分摊
        public bool IsSelected { get; set; }

        // 是否包含乘数比例系数公式 (例如 =ROUND(H20*0.08, 2))
        public bool HasRatioFormula { get; set; }

        // 包含此标签的箱柜频次统计
        public int CabinetOccurrences { get; set; }
    }

    /// <summary>
    /// 总价一键调整初始化数据响应包
    /// </summary>
    public class TotalPriceAdjustInitData
    {
        // 操作执行状态标志
        public bool Success { get; set; }

        // 提示或错误消息
        public string Message { get; set; } = string.Empty;

        // 当前活动工作簿工程名称
        public string ProjectName { get; set; } = string.Empty;

        // 当前所有分类表累计销售总价
        public decimal TotalSalesPrice { get; set; }

        // 当前所有分类表累计成本总价
        public decimal TotalCostPrice { get; set; }

        // 可参与调价的分类工作表列表
        public List<TotalPriceCategorySheetItem> CategorySheets { get; set; } = new List<TotalPriceCategorySheetItem>();

        // 动态扫描聚合得到的分摊标签列表
        public List<TotalPriceAdjustTagItem> Tags { get; set; } = new List<TotalPriceAdjustTagItem>();
    }

    /// <summary>
    /// 单个分摊项用户比例设定载荷
    /// </summary>
    public class TagAllocationSetting
    {
        // 目标标签名称
        public string Tag { get; set; } = string.Empty;

        // 用户指定的分摊比例 (0.0 ~ 1.0)
        public decimal Ratio { get; set; }
    }

    /// <summary>
    /// 总价一键调整前端请求载荷实体
    /// </summary>
    public class TotalPriceAdjustRequest
    {
        // 参与本次调价的目标分类工作表名称列表
        public List<string> SelectedSheets { get; set; } = new List<string>();

        // 用户设定的目标总价 (元)
        public decimal TargetPrice { get; set; }

        // 调价基准模式 ("total": 按项目总价, "cost": 按成本总价)
        public string BaselineType { get; set; } = "total";

        // 调价浮动百分比数值 (例如 5.0 代表 5%)
        public decimal AdjustPercentage { get; set; }

        // 浮动方向 ("up": 上浮, "down": 下浮)
        public string AdjustDirection { get; set; } = "down";

        // 各分摊标签项的比例设置集合 (总和必须等于 1.0)
        public List<TagAllocationSetting> TagAllocations { get; set; } = new List<TagAllocationSetting>();

        // 是否另存为调价副本新工作簿 (默认 false 为就地更新并提供撤销)
        public bool SaveAsCopy { get; set; }

        // 另存为副本时的自定义文件名 (可选)
        public string? CopyFileName { get; set; }
    }

    /// <summary>
    /// 总价一键调整执行结果实体模型
    /// </summary>
    public class TotalPriceAdjustResult
    {
        // 执行是否成功
        public bool Success { get; set; }

        // 操作结果文本提示
        public string Message { get; set; } = string.Empty;

        // 本次成功更新的分类工作表数量
        public int UpdatedSheetsCount { get; set; }

        // 本次成功更新的箱柜总台数
        public int UpdatedCabinetsCount { get; set; }

        // 本次成功重算系数的元器件行总数
        public int UpdatedComponentCount { get; set; }

        // 本次成功调整比例系数或金额的费用项总数
        public int UpdatedFeeItemCount { get; set; }

        // 调整前的原始工程总价
        public decimal OldTotalAmount { get; set; }

        // 调整后的全新工程总价
        public decimal NewTotalAmount { get; set; }

        // 本次调价的实际增减差额 (新总价 - 旧总价)
        public decimal DifferenceAmount { get; set; }

        // 另存为副本时的物理磁盘文件路径 (若开启另存)
        public string? SavedCopyFilePath { get; set; }
    }
}
