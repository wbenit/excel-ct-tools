using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 采购清单物料聚合实体模型
    /// 用于承载项目所有分类箱柜中提取并汇总去重后的采购元器件数据
    /// </summary>
    public class PurchaseListItem
    {
        // 采购清单序号 (1-indexed)
        public int Index { get; set; }

        // 元器件名称 (例如: 塑壳断路器、交流接触器)
        public string Name { get; set; } = string.Empty;

        // 元器件型号规格 (例如: NM1-125S/3300 100A)
        public string Model { get; set; } = string.Empty;

        // 计量单位 (例如: 台、只、套、个，默认"台")
        public string Unit { get; set; } = "台";

        // 采购总数量 (所有箱柜中 单台数量 * 箱柜台数 的累计总和)
        public decimal TotalQuantity { get; set; }

        // 生产厂家/品牌 (例如: 正泰、德力西、常熟开关)
        public string Brand { get; set; } = string.Empty;

        // 采购备注说明 (留空或记录补充参数)
        public string Remark { get; set; } = string.Empty;
    }

    /// <summary>
    /// 采购清单模板列动态映射模型
    /// 扫描模板行占位符动态识别各属性所在的物理列索引，杜绝硬编码列号
    /// </summary>
    public class PurchaseListTemplateColumnMap
    {
        // 序号列索引 (1-indexed，通常为第 1 列 A 列)
        public int IndexCol { get; set; } = 1;

        // 名称列索引 (匹配 [元器件名称] 或 [名称])
        public int NameCol { get; set; } = 2;

        // 型号规格列索引 (匹配 [型号] 或 [型号规格])
        public int ModelCol { get; set; } = 3;

        // 单位列索引 (匹配 [单位])
        public int UnitCol { get; set; } = 4;

        // 数量列索引 (匹配 [数量] 或 [单数])
        public int QuantityCol { get; set; } = 5;

        // 品牌/厂家列索引 (匹配 [品牌] 或 [厂家]，若未匹配则为 -1)
        public int BrandCol { get; set; } = -1;

        // 模板总有效列宽 (通常到 I 列为第 9 列)
        public int MaxCol { get; set; } = 9;
    }

    /// <summary>
    /// 采购清单导出操作执行结果模型
    /// </summary>
    public class PurchaseListExportResult
    {
        // 操作是否成功完成
        public bool Success { get; set; }

        // 提示信息或异常错误文本
        public string Message { get; set; } = string.Empty;

        // 累计汇总导出的有效物料项数
        public int ItemCount { get; set; }
    }

    /// <summary>
    /// 材料统计分类工作表选择项 DTO
    /// 用于前端弹窗展示分类名称、箱柜数量及勾选状态
    /// </summary>
    public class MaterialStatCategoryDto
    {
        // 分类工作表名称 (如 终端、配电、分类1)
        public string SheetName { get; set; } = string.Empty;

        // 该分类下的有效箱柜数量
        public int CabinetCount { get; set; }

        // 是否被勾选包含在统计中 (默认 true)
        public bool IsSelected { get; set; } = true;
    }

    /// <summary>
    /// 材料统计初始化数据 DTO 模型
    /// 用于前端弹窗首屏加载时同时返回分类列表及是否存在已有采购清单标志
    /// </summary>
    public class MaterialStatInitDataDto
    {
        // 分类工作表明细列表
        public List<MaterialStatCategoryDto> Categories { get; set; } = new List<MaterialStatCategoryDto>();

        // 当前工作簿是否已存在【采购清单】工作表 (用于前端感知并执行覆盖确认)
        public bool HasExistingPurchaseList { get; set; } = false;
    }
}

