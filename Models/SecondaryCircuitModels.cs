using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 二次图控制回路方案与定额配置实体模型
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，配置与硬编码显式标明
    /// </summary>
    public class SecondarySchemeEntity
    {
        // 方案唯一主键自增 ID (SQLite 数据库主键)
        [JsonPropertyName("id")]
        public int Id { get; set; }

        // 所属二次组别分类 (如: "双电源组", "电动机控制组", "照明组")
        [JsonPropertyName("groupName")]
        public string GroupName { get; set; } = string.Empty;

        // 二次回路方案主名称 (如: "双电源标准控制回路方案A")
        [JsonPropertyName("schemeName")]
        public string SchemeName { get; set; } = string.Empty;

        // 共享同一套配置的适用回路代号列表 (如: ["双电源1", "CA1B", "ATS-1"])
        [JsonPropertyName("applicableCodes")]
        public List<string> ApplicableCodes { get; set; } = new List<string>();

        // 对应的 AutoCAD 图纸名称 / 图号 / 图块名称 (如: "双电源二次图_V1.dwg")
        [JsonPropertyName("cadDrawingName")]
        public string CadDrawingName { get; set; } = string.Empty;

        // 方案级二次线跨门数量或总长度 (单位: 根或米，默认 0.0)
        [JsonPropertyName("crossDoorCount")]
        public double CrossDoorCount { get; set; } = 0.0;

        // 箱体门板开孔规范描述 (如: "圆2", "圆3方1")
        [JsonPropertyName("holeSpec")]
        public string HoleSpec { get; set; } = string.Empty;

        // 二次装配与配线接线人工工费 (单位: 元，默认 0.0)
        [JsonPropertyName("laborCost")]
        public double LaborCost { get; set; } = 0.0;

        // 关联本地个人物料库的子 BOM 元器件明细清单
        [JsonPropertyName("bomItems")]
        public List<SecondaryBomItem> BomItems { get; set; } = new List<SecondaryBomItem>();

        // 方案级二次材料费计算属性：由其下所有子物料材料费实时动态求和累加，绝不存死值
        [JsonPropertyName("totalMaterialCost")]
        public double TotalMaterialCost
        {
            // 实时计算子 BOM 材料费总和并保留两位小数
            get => BomItems == null ? 0.0 : Math.Round(BomItems.Sum(item => item.SubtotalCost), 2);
        }

        // 方案级综合总费用计算属性：二次材料费 + 人工工费
        [JsonPropertyName("totalCost")]
        public double TotalCost
        {
            // 综合成本由材料与人工实时动态合并
            get => Math.Round(TotalMaterialCost + LaborCost, 2);
        }

        // 备注说明与工艺要点提示
        [JsonPropertyName("remark")]
        public string Remark { get; set; } = string.Empty;

        // 方案创建时间戳
        [JsonPropertyName("createdAt")]
        public string CreatedAt { get; set; } = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        // 方案最后修改更新时间戳
        [JsonPropertyName("updatedAt")]
        public string UpdatedAt { get; set; } = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
    }

    /// <summary>
    /// 二次方案下的单个元器件子物料明细实体模型 (与本地物料库强关联)
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class SecondaryBomItem
    {
        // 关联本地个人物料库 components 表的主键 ID (为 0 表示未入库临时元件)
        [JsonPropertyName("componentId")]
        public int ComponentId { get; set; } = 0;

        // 元器件规范名称 (如: "指示灯", "熔断器", "中间继电器")
        [JsonPropertyName("name")]
        public string Name { get; set; } = string.Empty;

        // 元器件规格型号 (如: "AD11 AC220V 红", "RT18-32 4A")
        [JsonPropertyName("model")]
        public string Model { get; set; } = string.Empty;

        // 元器件生产厂家或品牌 (如: "正泰", "施耐德")
        [JsonPropertyName("brand")]
        public string Brand { get; set; } = string.Empty;

        // 元器件使用数量 (默认 1)
        [JsonPropertyName("quantity")]
        public int Quantity { get; set; } = 1;

        // 计量单位 (如: "只", "台", "个")
        [JsonPropertyName("unit")]
        public string Unit { get; set; } = "只"; // --硬编码-- 默认单位为"只"

        // 元器件单价 (动态从本地物料库关联读取，保持全局一致性)
        [JsonPropertyName("unitPrice")]
        public double UnitPrice { get; set; } = 0.0;

        // 子项材料费合价计算属性：数量 * 单价，动态计算不存静态死值
        [JsonPropertyName("subtotalCost")]
        public double SubtotalCost
        {
            // 实时公式驱动计算并四舍五入保留 2 位小数
            get => Math.Round(Quantity * UnitPrice, 2);
        }

        // 单件二次元件特殊备注说明
        [JsonPropertyName("remark")]
        public string Remark { get; set; } = string.Empty;
    }
}
