using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 智能输入元器件条目实体模型
    /// </summary>
    public class SmartComponentItem
    {
        // 元件规格型号 (C 列核心关键字，去重与选择主键)
        public string Model { get; set; } = string.Empty;

        // 元件名称 (B 列)
        public string Name { get; set; } = string.Empty;

        // 生产厂家 (D 列)
        public string Manufacturer { get; set; } = string.Empty;

        // 计量单位 (E 列)
        public string Unit { get; set; } = string.Empty;

        // 销售单价 (G 列)
        public decimal UnitPrice { get; set; }

        // 成本单价 (J 列)
        public decimal CostUnitPrice { get; set; }

        // M 列表价 / 基准价 (第 13 列，价格计算链条源头)
        public decimal CatalogPrice { get; set; }

        // L 列报出加价系数 (第 12 列)
        public string QuoteCoefficient { get; set; } = string.Empty;

        // N 列采购折扣系数 (第 14 列)
        public string DiscountCoefficient { get; set; } = string.Empty;

        // I 列规格备注说明 (第 9 列)
        public string Remark { get; set; } = string.Empty;

        // 元件类别 (Q 列)
        public string Category { get; set; } = string.Empty;

        // 所属来源工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 所属箱柜编号或名称
        public string CabinetNo { get; set; } = string.Empty;

        // 整行原始单元格映射 (Key: 1-based 列号 2~29 即 B列到AC列，严禁含 A列序号/F列数量/AD列句柄)
        public Dictionary<int, string> RowCells { get; set; } = new Dictionary<int, string>();
    }

    /// <summary>
    /// 单个工作表提取的元器件去重数据结构
    /// </summary>
    public class SheetComponentData
    {
        // 工作表名称 (例如: 分类1, 高压柜)
        public string SheetName { get; set; } = string.Empty;

        // 原始元器件总行数
        public int TotalCount { get; set; }

        // 去重后的有效元器件数量
        public int UniqueCount { get; set; }

        // 该工作表去重后的元器件列表集合
        public List<SmartComponentItem> Components { get; set; } = new List<SmartComponentItem>();
    }

    /// <summary>
    /// 全局元器件去重分类缓存文件根结构 (存储于 data/smart_components.json)
    /// </summary>
    public class SmartComponentsStorage
    {
        // 最后一次刷新提取的更新时间戳
        public string LastUpdatedTime { get; set; } = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        // 包含的所有工作表元器件数据集
        public List<SheetComponentData> Sheets { get; set; } = new List<SheetComponentData>();
    }

    /// <summary>
    /// 智能填写全局配置模型 (存储于 data/smart_input_config.json)
    /// </summary>
    public class SmartInputConfigModel
    {
        // 用户选中的数据源工作表名称列表 (只有勾选的表才参与输入候选)
        public List<string> SelectedSheets { get; set; } = new List<string>();

        // 回填字段范围: 是否回填 B 列 (元件名称)
        public bool FillName { get; set; } = true;

        // 回填字段范围: 是否回填 D 列 (生产厂家)
        public bool FillManufacturer { get; set; } = true;

        // 回填字段范围: 是否回填 E 列 (计量单位)
        public bool FillUnit { get; set; } = true;

        // 回填字段范围: 是否回填 G 列 (销售单价)
        public bool FillUnitPrice { get; set; } = true;

        // 是否启用全字段自动回填 (除了 AD 句柄列、A 序号列、F 数量列之外的所有列均自动回填，默认开启)
        public bool FillAllColumns { get; set; } = true;

        // 是否自动在 Excel 中启用原生下拉列表
        public bool AutoDropdownEnabled { get; set; } = true;

        // 是否在选中 C 列元器件行时自动启用智能覆盖输入 (默认开启)
        public bool AutoPopupFloatWindow { get; set; } = true;

        // 是否自动学习并加入新录入的元器件至候选词库 (默认开启)
        public bool AutoLearnNewComponents { get; set; } = true;
    }
}
