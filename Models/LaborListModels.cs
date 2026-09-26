using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 人工清单箱柜项实体模型
    /// 承载从分类表中提取出的各个箱柜台套、型号、数量、人工单价与超链接信息
    /// </summary>
    public class LaborListItem
    {
        // 人工清单序号 (1-indexed)
        public int Index { get; set; }

        // 箱柜名称 (如: 配电箱、照明箱，默认"配电箱") --硬编码: 默认箱柜名称--
        public string Name { get; set; } = "配电箱";

        // 箱柜型号/柜号 (对应 Cab_Det_k 的柜号或箱柜型号)
        public string Model { get; set; } = string.Empty;

        // 箱柜数量/台数 (默认 1 台)
        public decimal Quantity { get; set; } = 1;

        // 计量单位 (固定为 "台") --硬编码: 默认单位--
        public string Unit { get; set; } = "台";

        // 单台人工费单价 (元，由计费区总人工费除以台数得出)
        public decimal UnitPrice { get; set; }

        // 单台人工费计算出的总金额 (元，在 Excel 中由 =F*D 公式联动)
        public decimal TotalPrice { get; set; }

        // 备注说明文本
        public string Remark { get; set; } = string.Empty;

        // 所属来源分类工作表名称 (用于生成超链接回跳)
        public string SheetName { get; set; } = string.Empty;

        // 箱柜在来源分类表中的明细起始行号 (Cab_Det_k 行号)
        public int DetailStartRow { get; set; }

        // 箱柜在来源分类表中的明细终止行号 (Cab_Tolsum_k 行号)
        public int DetailEndRow { get; set; }
    }

    /// <summary>
    /// 人工清单模板列动态映射模型
    /// 扫描母版第 7 行占位符，动态匹配各字段所在的物理列索引，杜绝硬编码列号
    /// </summary>
    public class LaborListTemplateColumnMap
    {
        // 序号列索引 (匹配 [箱柜序号] 或 [序号]，默认第 1 列 A 列)
        public int IndexCol { get; set; } = 1;

        // 箱柜名称列索引 (匹配 [箱柜名称] 或 [名称]，默认第 2 列 B 列)
        public int NameCol { get; set; } = 2;

        // 柜号/型号列索引 (匹配 [柜号] 或 [型号]，默认第 3 列 C 列)
        public int ModelCol { get; set; } = 3;

        // 数量台数列索引 (匹配 [数量]，默认第 4 列 D 列)
        public int QuantityCol { get; set; } = 4;

        // 单位列索引 (匹配 [单位]，默认第 5 列 E 列)
        public int UnitCol { get; set; } = 5;

        // 单价列索引 (匹配 [人工费] 或 [单价]，默认第 6 列 F 列)
        public int UnitPriceCol { get; set; } = 6;

        // 金额列索引 (默认第 7 列 G 列)
        public int TotalPriceCol { get; set; } = 7;

        // 备注列索引 (默认第 8 列 H 列)
        public int RemarkCol { get; set; } = 8;

        // 模板有效最大列宽 (默认 8 列，到 H 列)
        public int MaxCol { get; set; } = 8;
    }

    /// <summary>
    /// 人工清单分类表选择项 DTO
    /// 用于前端弹窗多选框展示分类名称、箱柜数量及勾选状态
    /// </summary>
    public class LaborListCategoryDto
    {
        // 分类工作表名称 (如 分类1、配电等)
        public string SheetName { get; set; } = string.Empty;

        // 该分类下的有效箱柜总台数
        public int CabinetCount { get; set; }

        // 是否被勾选包含生成人工清单 (默认 true)
        public bool IsSelected { get; set; } = true;
    }

    /// <summary>
    /// 人工清单初始化数据 DTO 模型
    /// 用于前端向导弹窗首屏加载时返回分类列表及是否存在已有表标志
    /// </summary>
    public class LaborListInitDataDto
    {
        // 是否已存在人工清单工作表标志
        public bool HasExistingLaborList { get; set; }

        // 工作簿工程项目名称 (从项目信息表读取)
        public string ProjectName { get; set; } = string.Empty;

        // 可供选择的分类工作表列表
        public List<LaborListCategoryDto> Categories { get; set; } = new List<LaborListCategoryDto>();
    }

    /// <summary>
    /// 人工清单导出请求 DTO 模型
    /// </summary>
    public class LaborListExportRequestDto
    {
        // 用户选中的参与生成的分类表名称列表
        public List<string> SelectedSheetNames { get; set; } = new List<string>();
    }

    /// <summary>
    /// 人工清单导出操作执行结果模型
    /// </summary>
    public class LaborListExportResult
    {
        // 导出是否成功
        public bool Success { get; set; }

        // 提示信息或异常错误消息
        public string Message { get; set; } = string.Empty;

        // 导出的箱柜总台数
        public int CabinetCount { get; set; }

        // 导出的总人工金额 (元)
        public decimal TotalAmount { get; set; }
    }
}
