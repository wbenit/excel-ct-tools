using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 报价全景分析主数据传输对象 (DTO)
    /// 用于向 WebView2 前端 Vue 3 + ECharts 大屏提供结构化业务数据
    /// </summary>
    public class QuoteAnalysisDto
    {
        // 项目基本财务与概况信息
        public QuoteProjectSummaryDto Summary { get; set; } = new QuoteProjectSummaryDto();

        // 全表 32 台或多台箱柜汇总明细列表
        public List<CabinetAnalysisItemDto> Cabinets { get; set; } = new List<CabinetAnalysisItemDto>();

        // 全案成本构成五大维度解构 (元件、箱体、辅材、人工、成套费)
        public CostBreakdownDto CostStructure { get; set; } = new CostBreakdownDto();

        // 核心元器件供应链排行与体量分布 (来自元件汇总表)
        public List<ComponentRankingDto> TopComponents { get; set; } = new List<ComponentRankingDto>();

        // 箱柜功能类别聚合统计 (动力控制、照明配电、双电源、消防等)
        public List<CategoryDistributionDto> CategoryDistributions { get; set; } = new List<CategoryDistributionDto>();

        // 品牌供应链采购金额分布 (正泰、国优、安科瑞等)
        public List<BrandDistributionDto> BrandDistributions { get; set; } = new List<BrandDistributionDto>();
    }

    /// <summary>
    /// 项目宏观财务概况数据模型
    /// </summary>
    public class QuoteProjectSummaryDto
    {
        // 工程名称 (如: 江苏赛格瑞)
        public string ProjectName { get; set; } = string.Empty;

        // 报价单号 (如: WB202609231710)
        public string QuoteNumber { get; set; } = string.Empty;

        // 编制单位 (如: 扬州华科智能科技有限公司)
        public string CompanyName { get; set; } = string.Empty;

        // 报出总金额 (元)
        public double TotalQuotePrice { get; set; }

        // 综合成本总额 (元)
        public double TotalCostPrice { get; set; }

        // 毛利总额 (元)
        public double TotalGrossProfit { get; set; }

        // 综合毛利率 (百分比数值，例如 0.1955 表示 19.55%)
        public double GrossProfitRate { get; set; }

        // 箱柜总台套数
        public int TotalCabinetCount { get; set; }

        // 元器件总规格型号数
        public int TotalComponentCount { get; set; }
    }

    /// <summary>
    /// 单台箱柜全维度分析明细项模型
    /// </summary>
    public class CabinetAnalysisItemDto
    {
        // 箱柜序号 (1, 2, 3...)
        public int Index { get; set; }

        // 所属分类工作表名称 (如: 分类1)
        public string SheetName { get; set; } = string.Empty;

        // 柜号或箱柜系统名称 (如: 普通货梯电源箱 HT系统)
        public string Name { get; set; } = string.Empty;

        // 箱柜型号规格
        public string Model { get; set; } = string.Empty;

        // 箱柜台数
        public double Quantity { get; set; }

        // 报出单价 (元)
        public double UnitPrice { get; set; }

        // 报出总价 (元)
        public double TotalPrice { get; set; }

        // 成本总价 (元)
        public double CostPrice { get; set; }

        // 毛利金额 (元)
        public double GrossProfit { get; set; }

        // 毛利率 (例如 0.1955)
        public double GrossProfitRate { get; set; }

        // 箱柜智能功能大类 (动力控制 / 照明配电 / 消防应急 / 电梯系统 / 专用设备)
        public string FunctionalCategory { get; set; } = string.Empty;

        // 安装方式与环境 (暗装 / 明装 / 户外 / 落地)
        public string InstallType { get; set; } = string.Empty;

        // 对应分类表中的明细锚点行号 (用于图表联动下钻定位)
        public int DetRow { get; set; }

        // 箱体外形尺寸 (高x宽x深)
        public string Size { get; set; } = string.Empty;

        // 该箱柜内部元器件小计金额
        public double ComponentsCost { get; set; }

        // 该箱柜箱体外壳金额
        public double CabinetBoxCost { get; set; }

        // 该箱柜辅材耗材金额
        public double AuxMaterialsCost { get; set; }

        // 该箱柜人工配线制作工价
        public double LaborCost { get; set; }

        // 该箱柜综合成套费
        public double CompleteSetFee { get; set; }
    }

    /// <summary>
    /// 成本构成五大支柱解构模型
    /// </summary>
    public class CostBreakdownDto
    {
        // 元器件采购总成本 (元)
        public double ComponentsTotal { get; set; }

        // 箱体外壳制作总成本 (元)
        public double CabinetBoxTotal { get; set; }

        // 铜排辅材总成本 (元)
        public double AuxMaterialsTotal { get; set; }

        // 人工装配制作总工时费 (元)
        public double LaborTotal { get; set; }

        // 综合成套服务费总额 (元)
        public double CompleteSetTotal { get; set; }

        // 净毛利总额 (元)
        public double ProfitTotal { get; set; }
    }

    /// <summary>
    /// 核心元器件供应链排行模型
    /// </summary>
    public class ComponentRankingDto
    {
        // 排序索引
        public int Rank { get; set; }

        // 元器件名称 (如: 塑壳断路器、电涌保护器)
        public string Name { get; set; } = string.Empty;

        // 规格型号 (如: NXM-125S/4300B 50A)
        public string Model { get; set; } = string.Empty;

        // 生产厂家/品牌 (如: 正泰、国优)
        public string Manufacturer { get; set; } = string.Empty;

        // 全项目累计用量
        public double TotalQuantity { get; set; }

        // 采购单价 (元)
        public double UnitPrice { get; set; }

        // 采购合价总额 (元)
        public double TotalPrice { get; set; }

        // 在元器件采购池中的金额占比 (0~1)
        public double Percentage { get; set; }
    }

    /// <summary>
    /// 箱柜功能分类分布统计模型
    /// </summary>
    public class CategoryDistributionDto
    {
        // 功能分类名称 (如: 照明配电箱)
        public string CategoryName { get; set; } = string.Empty;

        // 该类别箱柜台数
        public int CabinetCount { get; set; }

        // 该类别总金额 (元)
        public double TotalPrice { get; set; }

        // 金额占比 (0~1)
        public double Percentage { get; set; }
    }

    /// <summary>
    /// 供应链品牌分布统计模型
    /// </summary>
    public class BrandDistributionDto
    {
        // 品牌名称 (如: 正泰、国优、安科瑞)
        public string BrandName { get; set; } = string.Empty;

        // 该品牌物料采购总金额
        public double TotalPrice { get; set; }

        // 占比 (0~1)
        public double Percentage { get; set; }
    }
}
