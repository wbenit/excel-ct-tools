using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 一次图成套柜体与系统方案配置实体模型
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，配置与硬编码显式标明
    /// </summary>
    public class PrimarySchemeEntity
    {
        // 方案唯一主键自增 ID (SQLite 数据库主键)
        [JsonPropertyName("id")]
        public int Id { get; set; }

        // 所属一次分类/方案子文件夹名 (如: "低压进线方案", "电容补偿柜方案")
        [JsonPropertyName("groupName")]
        public string GroupName { get; set; } = string.Empty;

        // 一次方案主名称 (如: "MNS 低压主进线柜(2000A/3P)")
        [JsonPropertyName("schemeName")]
        public string SchemeName { get; set; } = string.Empty;

        // 共享同一套配置的适用回路或柜号代号列表 (如: ["AH1", "AH2", "进线柜"])
        [JsonPropertyName("applicableCodes")]
        public List<string> ApplicableCodes { get; set; } = new List<string>();

        // 对应的 AutoCAD 图纸名称 / 图号 / 图块名称 (如: "MNS_IN_2000A.dwg")
        [JsonPropertyName("cadDrawingName")]
        public string CadDrawingName { get; set; } = string.Empty;

        // 柜型型号 (如: "MNS", "KYN28-12", "GGD", "PZ30")
        [JsonPropertyName("cabinetModel")]
        public string CabinetModel { get; set; } = string.Empty;

        // 外形尺寸 (宽*深*高 mm，如: "800*1000*2200")
        [JsonPropertyName("dimensions")]
        public string Dimensions { get; set; } = string.Empty;

        // 额定工作电流 (单位: A，如: 2000)
        [JsonPropertyName("ratedCurrent")]
        public double RatedCurrent { get; set; } = 0.0;

        // 主母排/铜排规格 (如: "TMY 3*(2*100*10)+1*(100*10)")
        [JsonPropertyName("busbarSpec")]
        public string BusbarSpec { get; set; } = string.Empty;

        // 柜体制作与装配人工工费 (单位: 元，默认 0.0)
        [JsonPropertyName("laborCost")]
        public double LaborCost { get; set; } = 0.0;

        // 主母排加工/铜排基础成本 (单位: 元，默认 0.0)
        [JsonPropertyName("copperCost")]
        public double CopperCost { get; set; } = 0.0;

        // 方案级品牌/生产厂家 (如: "常熟开关", "施耐德", "正泰")
        [JsonPropertyName("brand")]
        public string Brand { get; set; } = string.Empty;

        // 方案详细技术描述与主要技术参数说明
        [JsonPropertyName("description")]
        public string Description { get; set; } = string.Empty;

        // 一次主元器件 BOM 清单集合 (直接复用 CloudSchemeBomItem)
        [JsonPropertyName("bomItems")]
        public List<CloudSchemeBomItem> BomItems { get; set; } = new List<CloudSchemeBomItem>();

        // 方案级材料费用计算属性：由其下所有元器件合价动态汇总
        [JsonPropertyName("totalMaterialCost")]
        public double TotalMaterialCost
        {
            // 实时汇总子元器件合价并保留两位小数
            get => BomItems == null ? 0.0 : Math.Round(BomItems.Sum(b => (double)b.TotalPrice), 2);
        }

        // 方案级综合总费用计算属性：元器件材料费 + 制作装配工费 + 铜排加工成本
        [JsonPropertyName("totalCost")]
        public double TotalCost
        {
            // 综合成本由材料、人工与母排成本动态合并
            get => Math.Round(TotalMaterialCost + LaborCost + CopperCost, 2);
        }

        // 创建时间文本 (yyyy-MM-dd HH:mm:ss)
        [JsonPropertyName("createdAt")]
        public string CreatedAt { get; set; } = string.Empty;

        // 最后更新时间文本 (yyyy-MM-dd HH:mm:ss)
        [JsonPropertyName("updatedAt")]
        public string UpdatedAt { get; set; } = string.Empty;
    }

    /// <summary>
    /// 一次方案 DWG 根目录下的直接子文件夹项 DTO
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class PrimaryFolderItemDto
    {
        // 子文件夹显示名称 (如: "低压进线方案", "电容补偿柜方案")
        public string Name { get; set; } = string.Empty;

        // 子文件夹物理绝对路径
        public string FullPath { get; set; } = string.Empty;

        // 该子文件夹下包含的 DWG 图纸数量
        public int DwgCount { get; set; } = 0;
    }

    /// <summary>
    /// 云方案中心一次回路 DWG 图纸与数据库参数一体化卡片展示 DTO
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class PrimaryFolderDwgCardDto
    {
        // 完整文件名称 (带后缀，如: "MNS_IN_2000A.dwg")
        public string FileName { get; set; } = string.Empty;

        // 去扩展名的图名代号 (如: "MNS_IN_2000A")
        public string DwgName { get; set; } = string.Empty;

        // 方案大标题显示名称 (如: "MNS 低压主进线柜(2000A/3P)")
        public string SchemeName { get; set; } = string.Empty;

        // 所属方案目录名称 (左侧选中的子文件夹名，如: "低压进线方案")
        public string FolderName { get; set; } = string.Empty;

        // 所属子文件夹物理全路径
        public string FolderPath { get; set; } = string.Empty;

        // DWG 文件物理绝对全路径
        public string FullPath { get; set; } = string.Empty;

        // 柜型型号 (如: "MNS", "KYN28-12", "GGD")
        public string CabinetModel { get; set; } = string.Empty;

        // 外形尺寸 (宽*深*高 mm，如: "800*1000*2200")
        public string Dimensions { get; set; } = string.Empty;

        // 额定工作电流 (单位: A，如: 2000)
        public double RatedCurrent { get; set; } = 0.0;

        // 主母排规格说明
        public string BusbarSpec { get; set; } = string.Empty;

        // 柜型品牌/厂家
        public string Brand { get; set; } = string.Empty;

        // 方案工艺与技术描述
        public string Description { get; set; } = string.Empty;

        // 柜体制作与装配人工工费
        public double LaborCost { get; set; } = 0.0;

        // 主母排加工/铜排基础成本
        public double CopperCost { get; set; } = 0.0;

        // 元件材料费用小计 (实时从子 BOM 汇总，单位: 元)
        public double MaterialCost { get; set; } = 0.0;

        // 综合总成本 (人工 + 铜排 + 材料)
        public double TotalCost { get; set; } = 0.0;

        // DWG 缩略图 Base64 图像流
        public string PreviewBase64 { get; set; } = string.Empty;

        // 是否在 personal_components.db 中精准匹配命中
        public bool IsMatched { get; set; } = false;

        // 命中的数据库主键 ID (未命中为 0)
        public int SchemeId { get; set; } = 0;

        // 方案最近更新时间文本
        public string UpdateTime { get; set; } = string.Empty;

        // 关联的完整一次方案实体 (供编辑弹窗直接反显)
        public PrimarySchemeEntity? SchemeData { get; set; }
    }
}
