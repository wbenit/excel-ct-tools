using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 元器件分布穿透查询入参模型
    /// </summary>
    public class ComponentDistributionQuery
    {
        // 元件名称 (例如: 微型断路器)
        public string ComponentName { get; set; } = string.Empty;

        // 型号规格 (汇总表中 C 列原型号或 D 列新型号)
        public string Model { get; set; } = string.Empty;

        // 原型号规格 (汇总表中 C 列原始型号基准)
        public string OriginalModel { get; set; } = string.Empty;

        // 生产厂家 (汇总表中 I 列)
        public string Manufacturer { get; set; } = string.Empty;

        // 单位 (汇总表中 E 列)
        public string Unit { get; set; } = string.Empty;

        // 汇总表当前行的汇总数量 (F 列)
        public decimal SummaryQuantity { get; set; } = 0;

        // 汇总表中的原始型号 (R 列对应明细表 U 列原始串)
        public string RawModelFromU { get; set; } = string.Empty;

        // 右键触发时所在的汇总表行号
        public int SummaryRowIndex { get; set; } = 0;

        // 参与本次汇总的源分类表名称列表 (从元件汇总表 AB1 单元格读取，只扫描此列表中的分类表)
        public List<string> SourceSheets { get; set; } = new List<string>();
    }

    /// <summary>
    /// 元器件在具体分类明细表某个箱柜中的单行分布数据项 DTO (已去除总数量与操作列)
    /// </summary>
    public class ComponentDistributionRowDto
    {
        // 行序号 (从 1 开始递增)
        public int Index { get; set; }

        // 所属分类工作表名称 (如: 1#动力箱表)
        public string SheetName { get; set; } = string.Empty;

        // 箱柜柜号 (如: 1AP1)
        public string CabinetNo { get; set; } = string.Empty;

        // 设备/箱柜名称 (如: 动力配电箱)
        public string CabinetName { get; set; } = string.Empty;

        // 单柜内该元件数量 (明细表元器件区间 F 列)
        public decimal QuantityPerCabinet { get; set; }

        // 箱柜工程台数 (汇总行 F 列读取)
        public int CabinetCount { get; set; } = 1;

        // 回路编号或备注信息 (明细行 I 列备注)
        public string Remark { get; set; } = string.Empty;

        // 该元件在明细表中的物理行号 (供穿透定位)
        public int RowIndex { get; set; }

        // 该元件在明细表中的物理列号 (默认 2，即 B 列元件名称)
        public int ColumnIndex { get; set; } = 2; // --硬编码: 默认定位列为 B 列--

        // 所属箱柜 Cab_Det 信息行物理行号 (用于穿透定位时将 det 行置于视口首行)
        public int DetRow { get; set; }
    }

    /// <summary>
    /// 按分类工作表分组的元器件分布集合 DTO
    /// </summary>
    public class ComponentDistributionCategoryGroupDto
    {
        // 分类工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 该分类表下涉及的箱柜总台数
        public int CabinetCount { get; set; }

        // 该分类表下涉及的元器件单柜数量小计
        public decimal SubtotalQuantityPerCabinet { get; set; }

        // 该分类表下的明细行列表
        public List<ComponentDistributionRowDto> Rows { get; set; } = new List<ComponentDistributionRowDto>();
    }

    /// <summary>
    /// 元器件分布穿透查询最终响应 DTO
    /// </summary>
    public class ComponentDistributionResultDto
    {
        // 是否查询成功
        public bool Success { get; set; } = true;

        // 提示或错误消息
        public string Message { get; set; } = string.Empty;

        // 目标元器件基本信息概要
        public ComponentDistributionQuery? ComponentInfo { get; set; }

        // 涉及的分类工作表总数
        public int TotalSheetCount { get; set; }

        // 涉及的箱柜总台数
        public int TotalCabinetCount { get; set; }

        // 单柜数量合计总和
        public decimal TotalQuantityPerCabinet { get; set; }

        // 按工作表分组的分布列表
        public List<ComponentDistributionCategoryGroupDto> Groups { get; set; } = new List<ComponentDistributionCategoryGroupDto>();

        // 扁平明细列表 (供直接全局搜索与过滤)
        public List<ComponentDistributionRowDto> FlatRows { get; set; } = new List<ComponentDistributionRowDto>();

        // 扫描耗时 (毫秒)
        public long ElapsedMilliseconds { get; set; }
    }

    /// <summary>
    /// 从分类明细表反向差量更新元件汇总表结果 DTO (方案 A 智能分流)
    /// </summary>
    public class UpdateSummaryResultDto
    {
        // 是否执行成功
        public bool Success { get; set; } = true;

        // 执行结果中文提示
        public string Message { get; set; } = string.Empty;

        // 是否触发了多柜分流 (true: 智能分流插入或累加新行; false: 原行直接更新)
        public bool IsSplit { get; set; } = false;

        // 更新或分流的目标行号
        public int TargetSummaryRow { get; set; }

        // 更新后的最新型号规格 (更新到 C 列原型号规格)
        public string NewModel { get; set; } = string.Empty;

        // 差量转移用量
        public decimal DeltaQuantity { get; set; }
    }
}

