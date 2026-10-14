using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 云物料中心通用物料数据传输对象
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class HubMaterialDto
    {
        // 物料自增主键 ID
        public int Id { get; set; }

        // 物料类别: "元件" | "母排" | "柜体" | "辅料"
        public string Category { get; set; } = "元件";

        // 品牌
        public string Brand { get; set; } = string.Empty;

        // 物料名称 (如: 矩形铜母线、XM配电箱壳体、微型断路器)
        public string Name { get; set; } = string.Empty;

        // 规格型号 (如: TMY-3*(100x5)、800*2200*1000、C65N-3P-C20A)
        public string Model { get; set; } = string.Empty;

        // 单价 (元)
        public decimal Price { get; set; }

        // 计量单位 (台 / 米 / kg / 个 / 套)
        public string Unit { get; set; } = "个";

        // 数量 (默认 1.0)
        public double Quantity { get; set; } = 1.0;

        // 认证状态: true=已认证(企业物料), false=未认证(个人物料)
        public bool IsCertified { get; set; } = false;

        // 额定电流 (A)
        public int? Current { get; set; }

        // 极数 (如: 1P, 2P, 3P, 4P)
        public string? Poles { get; set; }

        // 备注描述或物料编码
        public string? Remark { get; set; }

        // 扩展属性1 (如母排重长比、板材材质等)
        public string? Param1 { get; set; }

        // 扩展属性2 (JSON扩展参数或工艺配置)
        public string? Param2 { get; set; }
    }

    /// <summary>
    /// 物料库分页与条件查询参数模型
    /// </summary>
    public class HubMaterialQueryParams
    {
        // 目标子库类型: "元件" | "母排" | "柜体" | "辅料" | "全部"
        public string Category { get; set; } = "全部";

        // 型号或编码模糊检索关键字
        public string? ModelKeyword { get; set; }

        // 物料名称模糊检索关键字
        public string? NameKeyword { get; set; }

        // 品牌过滤
        public string? Brand { get; set; }

        // 认证状态过滤: "all" | "certified" | "uncertified"
        public string CertFilter { get; set; } = "all";

        // 分页页码 (从 1 开始)
        public int PageIndex { get; set; } = 1;

        // 分页大小
        public int PageSize { get; set; } = 50;
    }

    /// <summary>
    /// 物料库分页查询结果模型
    /// </summary>
    public class HubMaterialQueryResult
    {
        // 符合条件的总记录数
        public int TotalCount { get; set; }

        // 当前页物料数据列表
        public List<HubMaterialDto> Items { get; set; } = new List<HubMaterialDto>();
    }

    /// <summary>
    /// 国标母线载流量矩阵明细条目模型 (TMY / LMY)
    /// </summary>
    public class BusbarAmpacityItem
    {
        // 规格序号
        public int Index { get; set; }

        // 截面规格名称 (如: 15*3, 40*4, 100*10)
        public string Spec { get; set; } = string.Empty;

        // 截面宽度 (mm)
        public double Width { get; set; }

        // 截面厚度 (mm)
        public double Thickness { get; set; }

        // 单排截面积 (mm²)
        public double Area => Width * Thickness;

        // 每米单重 (kg/m, 紫铜密度按 8.9 g/cm³)
        public double UnitWeight => Math.Round(Area * 0.0089, 4);

        // 25℃ 平放最大允许持续电流 (A)
        public double Amp25Flat { get; set; }

        // 25℃ 竖放最大允许持续电流 (A)
        public double Amp25Vertical { get; set; }

        // 35℃ 平放最大允许持续电流 (A)
        public double Amp35Flat { get; set; }

        // 35℃ 竖放最大允许持续电流 (A)
        public double Amp35Vertical { get; set; }

        // 40℃ 平放最大允许持续电流 (A)
        public double Amp40Flat { get; set; }

        // 40℃ 竖放最大允许持续电流 (A)
        public double Amp40Vertical { get; set; }

        // 动态计算单价 (元/米，基于当前铜价联动)
        public decimal UnitPrice { get; set; }
    }
}
