using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 从 Excel 分类表明细提取出的待入库元器件候选实体模型
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class AddToPersonalDbCandidateItem
    {
        // 候选条目唯一序号 (前端列表展示与索引对应)
        public int Index { get; set; }

        // 元器件在当前 Excel 工作表中的物理行号
        public int ExcelRow { get; set; }

        // 品牌/生产厂家 (D 列)
        public string Brand { get; set; } = string.Empty;

        // 元件名称 (B 列)
        public string Name { get; set; } = string.Empty;

        // 规格型号 (C 列，必填关键项)
        public string Model { get; set; } = string.Empty;

        // 本次表格中提取出的单价/面价 (M 列)
        public decimal Price { get; set; } = 0.0m;

        // 备注信息 (I 列)
        public string Remark { get; set; } = string.Empty;

        // 物料类别 (Q 列，如“元件”、“电线”、“材料”)
        public string Category { get; set; } = string.Empty;

        // 启发式提取的额定/整定电流 (整型安培数)
        public int? Current { get; set; }

        // 启发式提取的极数 (如 "3", "4")
        public string Poles { get; set; } = string.Empty;

        // 启发式提取的脱扣特性 (如 "C", "D")
        public string Tripping { get; set; } = string.Empty;

        // 当前排重比对状态: "New"(全新物料), "PriceChanged"(已存在但价格变动), "ExactMatch"(完全一致)
        public string Status { get; set; } = "New"; // --硬编码: 默认状态新物料--

        // 状态文字说明展示 (如 "全新物料"、"价格变动"、"完全一致")
        public string StatusText { get; set; } = "全新物料"; // --硬编码: 默认状态文案--

        // 本地个人库中已存在的原单价 (若不存在则为 null)
        public decimal? OldPrice { get; set; }

        // 本地个人库中已存在的数据库主键 ID (若不存在则为 null)
        public int? OldId { get; set; }

        // 前端表格中是否勾选入库 (默认勾选)
        public bool IsSelected { get; set; } = true;
    }

    /// <summary>
    /// 前端提交执行入库的请求参数模型
    /// </summary>
    public class AddToPersonalDbSubmitRequest
    {
        // 是否启用“新价格覆盖更新” (默认开启，符合用户决策)
        public bool OverwritePrice { get; set; } = true;

        // 待入库的元器件条目列表
        public List<AddToPersonalDbCandidateItem> Items { get; set; } = new List<AddToPersonalDbCandidateItem>();
    }

    /// <summary>
    /// 执行入库结果模型
    /// </summary>
    public class AddToPersonalDbResult
    {
        // 操作是否整体成功
        public bool Success { get; set; } = true;

        // 处理的总条目数
        public int TotalCount { get; set; } = 0;

        // 成功新增入库的条目数
        public int InsertedCount { get; set; } = 0;

        // 覆盖更新现有库数据的条目数
        public int UpdatedCount { get; set; } = 0;

        // 跳过不处理的条目数
        public int SkippedCount { get; set; } = 0;

        // 结果反馈描述文本
        public string Message { get; set; } = string.Empty;
    }
}
