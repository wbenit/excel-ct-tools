using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 全局计算定额与规则配置聚合根模型
    /// </summary>
    public class QuotationRules
    {
        // 通用加点与税率综合设置分节
        [JsonPropertyName("general")]
        public GeneralConfig General { get; set; } = new GeneralConfig();

        // 壳体尺寸选型与空间配置分节
        [JsonPropertyName("shellRules")]
        public ShellConfig ShellRules { get; set; } = new ShellConfig();

        // 铜排与母线计算定额配置分节
        [JsonPropertyName("copperRules")]
        public CopperConfig CopperRules { get; set; } = new CopperConfig();

        // 辅材与配线定额配置分节
        [JsonPropertyName("auxRules")]
        public AuxConfig AuxRules { get; set; } = new AuxConfig();

        // 人工工价定额配置分节
        [JsonPropertyName("laborRules")]
        public LaborConfig LaborRules { get; set; } = new LaborConfig();
    }

    /// <summary>
    /// 通用加点、综合税费与铜价基础配置模型
    /// </summary>
    public class GeneralConfig
    {
        // 元件加点系数 (对应原 VBA 中的 xishu，默认 1.0)
        [JsonPropertyName("elementMarkupRatio")]
        public double ElementMarkupRatio { get; set; } = 1.0;

        // 税金与综合管理费乘数 (默认 1.13) --硬编码--
        [JsonPropertyName("taxAndManageRatio")]
        public double TaxAndManageRatio { get; set; } = 1.13;

        // 铜排基准市场单价 (单位: 元/KG，默认 76.0) --硬编码--
        [JsonPropertyName("copperPricePerKg")]
        public double CopperPricePerKg { get; set; } = 76.0;
    }

    /// <summary>
    /// 壳体选型规则与接线空间配置模型
    /// </summary>
    public class ShellConfig
    {
        // 壳体匹配名称 (默认 "箱体"，用于计费区域匹配及回写) --硬编码--
        [JsonPropertyName("shellMatchName")]
        public string ShellMatchName { get; set; } = "箱体";

        // 配电箱体容积率安全系数 (可用面积/元件面积，默认 1.2) --硬编码--
        [JsonPropertyName("boxAreaSafetyFactor")]
        public double BoxAreaSafetyFactor { get; set; } = 1.2;

        // 落地柜体容积率安全系数 (默认 1.3) --硬编码--
        [JsonPropertyName("cabinetAreaSafetyFactor")]
        public double CabinetAreaSafetyFactor { get; set; } = 1.3;

        // 落地柜体判定最小高度门限 (单位: mm，默认 1500) --硬编码--
        [JsonPropertyName("cabinetMinHeight")]
        public int CabinetMinHeight { get; set; } = 1500;

        // 柜体判定的电流门限 (单位: A，默认 250) --硬编码--
        [JsonPropertyName("cabinetCurrentThreshold")]
        public int CabinetCurrentThreshold { get; set; } = 250;

        // 可供智能匹配的标准壳体尺寸库 (格式: 宽*高，单位 mm)
        [JsonPropertyName("standardSizes")]
        public List<string> StandardSizes { get; set; } = new List<string>
        {
            "400*500", "500*600", "600*700", "600*800", "700*900", "800*1000",
            "800*1200", "800*1600", "800*1800", "800*2000", "800*2200", "1000*2200"
        };

        // 元件电流对应的上下接线空间高度梯度表 (单位: mm)
        [JsonPropertyName("wiringSpaceGradients")]
        public List<WiringSpaceItem> WiringSpaceGradients { get; set; } = new List<WiringSpaceItem>
        {
            new WiringSpaceItem { MaxCurrent = 40, Space = 100 },
            new WiringSpaceItem { MaxCurrent = 63, Space = 110 },
            new WiringSpaceItem { MaxCurrent = 100, Space = 125 },
            new WiringSpaceItem { MaxCurrent = 140, Space = 155 },
            new WiringSpaceItem { MaxCurrent = 160, Space = 200 },
            new WiringSpaceItem { MaxCurrent = 9999, Space = 370 }
        };

        // 互感器对总开关预留高度加成 (单位: mm)
        [JsonPropertyName("transformerSpacing")]
        public TransformerSpacingConfig TransformerSpacing { get; set; } = new TransformerSpacingConfig();
    }

    /// <summary>
    /// 接线空间阶梯条目模型
    /// </summary>
    public class WiringSpaceItem
    {
        // 电流上限 (单位: A)
        [JsonPropertyName("maxCurrent")]
        public int MaxCurrent { get; set; }

        // 对应接线空间预留高度 (单位: mm)
        [JsonPropertyName("space")]
        public int Space { get; set; }
    }

    /// <summary>
    /// 互感器空间预留高度配置
    /// </summary>
    public class TransformerSpacingConfig
    {
        // 火灾互感器额外预留高度 (单位: mm，默认 50) --硬编码--
        [JsonPropertyName("fireTransformer")]
        public int FireTransformer { get; set; } = 50;

        // 1套普通互感器预留高度 (默认 100) --硬编码--
        [JsonPropertyName("oneSet")]
        public int OneSet { get; set; } = 100;

        // 2~5套互感器预留高度 (默认 350) --硬编码--
        [JsonPropertyName("twoToFiveSets")]
        public int TwoToFiveSets { get; set; } = 350;

        // 大于5套互感器预留高度 (默认 600) --硬编码--
        [JsonPropertyName("overFiveSets")]
        public int OverFiveSets { get; set; } = 600;
    }

    /// <summary>
    /// 铜排与母线计算定额规则模型 (基于 tmy.DrawIO 全新铜排制作与推导规则)
    /// </summary>
    public class CopperConfig
    {
        // 出线塑壳断路器数量门限 (单位: 台，默认 2) --硬编码--
        [JsonPropertyName("mccbCountThreshold")]
        public int MccbCountThreshold { get; set; } = 2;

        // 出线塑壳断路器电流之和门限 (单位: A，默认 250) --硬编码--
        [JsonPropertyName("mccbCurrentSumThreshold")]
        public int MccbCurrentSumThreshold { get; set; } = 250;

        // 出线分路总电流之和门限 (单位: A，默认 300) --硬编码--
        [JsonPropertyName("branchTotalCurrentThreshold")]
        public int BranchTotalCurrentThreshold { get; set; } = 300;

        // 主进线开关电流判定门限 (单位: A，默认 250) --硬编码--
        [JsonPropertyName("mainSwitchCurrentThreshold")]
        public int MainSwitchCurrentThreshold { get; set; } = 250;

        // 极数为 4 的塑壳断路器数量门限 (单位: 台，默认 1，>=该值采用 4 根水平排，否则 3 根) --硬编码--
        [JsonPropertyName("fourPoleMccbThreshold")]
        public int FourPoleMccbThreshold { get; set; } = 1;

        // 触发垂直 N 排的特殊元器件关键字列表 (可动态编辑增删) --硬编码--
        [JsonPropertyName("specialComponentKeywords")]
        public List<string> SpecialComponentKeywords { get; set; } = new List<string>
        {
            "双电源", "ATS", "火灾探测器", "火灾互感器", "电气火灾"
        };

        // 柜宽扣除边距 (单位: mm，默认 120) --硬编码--
        [JsonPropertyName("widthDeduction")]
        public int WidthDeduction { get; set; } = 120;

        // 柜高扣除边距 (单位: mm，默认 300) --硬编码--
        [JsonPropertyName("heightDeduction")]
        public int HeightDeduction { get; set; } = 300;

        // 垂直母排基准展开长 (单位: m，默认 1.2) --硬编码--
        [JsonPropertyName("verticalBaseLength")]
        public double VerticalBaseLength { get; set; } = 1.2;

        // 垂直母排负荷延伸系数 (单位: m，默认 0.1) --硬编码--
        [JsonPropertyName("loadExtensionRatio")]
        public double LoadExtensionRatio { get; set; } = 0.1;

        // 垂直母排负荷延伸步长电流基数 (单位: A，默认 150) --硬编码--
        [JsonPropertyName("loadExtensionStepCurrent")]
        public int LoadExtensionStepCurrent { get; set; } = 150;

        // 出线大电流分支铜排起算门限 (单位: A，默认 100) --硬编码--
        [JsonPropertyName("branchMinCurrent")]
        public int BranchMinCurrent { get; set; } = 100;

        // 出线分支铜排单台基准展开长 (单位: 米，默认 1.0) --硬编码--
        [JsonPropertyName("branchBusUnitLength")]
        public double BranchBusUnitLength { get; set; } = 1.0;

        // 电流区间与对应铜排截面及每米单重对照表 (kg/m)
        [JsonPropertyName("mainBusSpecTable")]
        public List<MainBusSpecItem> MainBusSpecTable { get; set; } = new List<MainBusSpecItem>
        {
            new MainBusSpecItem { MaxCurrent = 100, Spec = "TMY-20*3", WeightPerMeter = 0.534 },
            new MainBusSpecItem { MaxCurrent = 160, Spec = "TMY-25*3", WeightPerMeter = 0.668 },
            new MainBusSpecItem { MaxCurrent = 250, Spec = "TMY-30*4", WeightPerMeter = 1.068 },
            new MainBusSpecItem { MaxCurrent = 400, Spec = "TMY-40*4", WeightPerMeter = 1.424 },
            new MainBusSpecItem { MaxCurrent = 630, Spec = "TMY-50*5", WeightPerMeter = 2.225 },
            new MainBusSpecItem { MaxCurrent = 800, Spec = "TMY-60*6", WeightPerMeter = 3.204 },
            new MainBusSpecItem { MaxCurrent = 1250, Spec = "TMY-80*8", WeightPerMeter = 5.696 },
            new MainBusSpecItem { MaxCurrent = 1600, Spec = "TMY-100*10", WeightPerMeter = 8.900 },
            new MainBusSpecItem { MaxCurrent = 9999, Spec = "TMY-120*10", WeightPerMeter = 10.680 }
        };

        // 倒T结构母排电流阈值 (单位: A，兼容保留) --硬编码--
        [JsonPropertyName("invertedTCurrent")]
        public int InvertedTCurrent { get; set; } = 300;

        // I型结构母排电流阈值 (单位: A，兼容保留) --硬编码--
        [JsonPropertyName("iStructureCurrent")]
        public int IStructureCurrent { get; set; } = 140;

        // 四极主母排预留长度补偿 (单位: mm，兼容保留) --硬编码--
        [JsonPropertyName("fourPoleExtra")]
        public int FourPoleExtra { get; set; } = 1400;

        // 三极主母排预留长度补偿 (单位: mm，兼容保留) --硬编码--
        [JsonPropertyName("threePoleExtra")]
        public int ThreePoleExtra { get; set; } = 1200;

        // 双电源(ATS)倒T结构增加排长 (单位: mm，兼容保留) --硬编码--
        [JsonPropertyName("atsInvertedTExtra")]
        public int AtsInvertedTExtra { get; set; } = 4800;

        // 双电源(ATS)I型结构增加排长 (单位: mm，兼容保留) --硬编码--
        [JsonPropertyName("atsIExtra")]
        public int AtsIExtra { get; set; } = 4200;

        // 附件与特殊元器件铜排动态影响规则库 (兼容保留)
        [JsonPropertyName("attachmentRules")]
        public List<AttachmentBusbarRule> AttachmentRules { get; set; } = new List<AttachmentBusbarRule>();
    }

    /// <summary>
    /// 附件与特殊元器件铜排动态影响规则条目模型
    /// </summary>
    public class AttachmentBusbarRule
    {
        // 匹配元器件名称或型号的关键字 (如 "双电源", "ATS", "火灾互感器", "母联开关")
        [JsonPropertyName("keyword")]
        public string Keyword { get; set; } = string.Empty;

        // 适用母排结构类型: "all"(通用), "invertedT"(仅倒T型), "iStructure"(仅I型)
        [JsonPropertyName("targetStructure")]
        public string TargetStructure { get; set; } = "all";

        // 柜宽联动排数 (根数，与有效柜宽 (W-dW) 乘算)
        [JsonPropertyName("widthMultiplier")]
        public int WidthMultiplier { get; set; } = 3;

        // 柜高联动排数 (根数，与有效柜高 (H-dH) 乘算)
        [JsonPropertyName("heightMultiplier")]
        public int HeightMultiplier { get; set; } = 0;

        // 固定折弯/端头/相间预留补偿长度 (单位: mm)
        [JsonPropertyName("extraFixedLength")]
        public int ExtraFixedLength { get; set; } = 1200;

        // 铜排规格选用: true 表示采用主母排单重，false 表示采用该元件自身回路电流规格
        [JsonPropertyName("useMainBusSpec")]
        public bool UseMainBusSpec { get; set; } = true;

        // 是否启用当前动态规则
        [JsonPropertyName("isEnabled")]
        public bool IsEnabled { get; set; } = true;
    }

    /// <summary>
    /// 母排截面规格条目模型
    /// </summary>
    public class MainBusSpecItem
    {
        // 电流上限 (单位: A)
        [JsonPropertyName("maxCurrent")]
        public int MaxCurrent { get; set; }

        // 铜排规格型号
        [JsonPropertyName("spec")]
        public string Spec { get; set; } = string.Empty;

        // 每米理论重量 (单位: kg/m)
        [JsonPropertyName("weightPerMeter")]
        public double WeightPerMeter { get; set; }
    }

    /// <summary>
    /// 一次导线长度生成与折算参数配置模型
    /// </summary>
    public class PrimaryWireLengthConfig
    {
        // 基础垂直预留高度 (单位: mm，默认 130) --硬编码--
        [JsonPropertyName("baseVerticalHeight")]
        public int BaseVerticalHeight { get; set; } = 130;

        // 火灾互感器垂直高度增量 (单位: mm，默认 100) --硬编码--
        [JsonPropertyName("fireTransformerExtraHeight")]
        public int FireTransformerExtraHeight { get; set; } = 100;

        // 普通互感器垂直高度增量 (单位: mm，默认 130) --硬编码--
        [JsonPropertyName("normalTransformerExtraHeight")]
        public int NormalTransformerExtraHeight { get; set; } = 130;

        // 落地柜判定高度门限 (单位: mm，默认 1600) --硬编码--
        [JsonPropertyName("cabinetMinHeight")]
        public int CabinetMinHeight { get; set; } = 1600;

        // 落地柜一次导线柜宽系数 (默认 0.7) --硬编码--
        [JsonPropertyName("cabinetWidthFactor")]
        public double CabinetWidthFactor { get; set; } = 0.7;

        // 落地柜一次导线长度裕量放大系数 (默认 1.1) --硬编码--
        [JsonPropertyName("cabinetLengthMargin")]
        public double CabinetLengthMargin { get; set; } = 1.1;

        // 配电箱一次导线箱宽系数 (默认 0.6) --硬编码--
        [JsonPropertyName("boxWidthFactor")]
        public double BoxWidthFactor { get; set; } = 0.6;
    }

    /// <summary>
    /// 辅材与配线定额规则模型
    /// </summary>
    public class AuxConfig
    {
        // 基础辅材起步费 (单位: 元，默认 10.0) --硬编码--
        [JsonPropertyName("baseFee")]
        public double BaseFee { get; set; } = 10.0;

        // 小箱零地排固定补贴 (单位: 元，默认 30.0) --硬编码--
        [JsonPropertyName("smallBoxGroundBarFee")]
        public double SmallBoxGroundBarFee { get; set; } = 30.0;

        // 高柜(高度>1500mm)辅材补贴 (单位: 元，默认 40.0) --硬编码--
        [JsonPropertyName("highCabinetExtraFee")]
        public double HighCabinetExtraFee { get; set; } = 40.0;

        // 一次导线长度生成折算配置
        [JsonPropertyName("wireLengthConfig")]
        public PrimaryWireLengthConfig WireLengthConfig { get; set; } = new PrimaryWireLengthConfig();

        // 一次配线选型对照表 (按电流门限匹配导线截面与每米单价)
        [JsonPropertyName("primaryWireSpecTable")]
        public List<PrimaryWireSpecItem> PrimaryWireSpecTable { get; set; } = new List<PrimaryWireSpecItem>
        {
            new PrimaryWireSpecItem { MaxCurrent = 16, Spec = "BV-2.5", CrossSection = 2.5, PricePerMeter = 1.5 },
            new PrimaryWireSpecItem { MaxCurrent = 25, Spec = "BV-4.0", CrossSection = 4.0, PricePerMeter = 2.2 },
            new PrimaryWireSpecItem { MaxCurrent = 40, Spec = "BV-6.0", CrossSection = 6.0, PricePerMeter = 3.2 },
            new PrimaryWireSpecItem { MaxCurrent = 63, Spec = "BV-10", CrossSection = 10.0, PricePerMeter = 5.5 },
            new PrimaryWireSpecItem { MaxCurrent = 80, Spec = "BV-16", CrossSection = 16.0, PricePerMeter = 8.5 },
            new PrimaryWireSpecItem { MaxCurrent = 100, Spec = "BV-25", CrossSection = 25.0, PricePerMeter = 13.5 },
            new PrimaryWireSpecItem { MaxCurrent = 125, Spec = "BV-35", CrossSection = 35.0, PricePerMeter = 18.5 },
            // 大电流回路无水平排时的导线选型兜底 (160A: BV-50, 250A: BV-70, 400A+: BV-95) --硬编码--
            new PrimaryWireSpecItem { MaxCurrent = 160, Spec = "BV-50", CrossSection = 50.0, PricePerMeter = 26.5 },
            new PrimaryWireSpecItem { MaxCurrent = 250, Spec = "BV-70", CrossSection = 70.0, PricePerMeter = 38.0 },
            new PrimaryWireSpecItem { MaxCurrent = 9999, Spec = "BV-95", CrossSection = 95.0, PricePerMeter = 52.0 }
        };

        // 二次元件接线定额库 (根据元件关键字匹配配线根数、线单价与工价)
        [JsonPropertyName("secondaryElements")]
        public List<SecondaryElementRule> SecondaryElements { get; set; } = new List<SecondaryElementRule>
        {
            new SecondaryElementRule { Keyword = "接触器", WireCount = 4, WirePrice = 0.8, LaborPrice = 8.0 },
            new SecondaryElementRule { Keyword = "中间继电器", WireCount = 4, WirePrice = 0.8, LaborPrice = 6.0 },
            new SecondaryElementRule { Keyword = "热继电器", WireCount = 2, WirePrice = 0.8, LaborPrice = 4.0 },
            new SecondaryElementRule { Keyword = "按钮", WireCount = 2, WirePrice = 0.8, LaborPrice = 3.0 },
            new SecondaryElementRule { Keyword = "指示灯", WireCount = 2, WirePrice = 0.8, LaborPrice = 3.0 },
            new SecondaryElementRule { Keyword = "多功能表", WireCount = 8, WirePrice = 0.8, LaborPrice = 15.0 },
            new SecondaryElementRule { Keyword = "电度表", WireCount = 6, WirePrice = 0.8, LaborPrice = 12.0 },
            new SecondaryElementRule { Keyword = "变频器", WireCount = 8, WirePrice = 0.8, LaborPrice = 30.0 },
            new SecondaryElementRule { Keyword = "断路器附件", WireCount = 3, WirePrice = 0.8, LaborPrice = 5.0 }
        };
    }

    /// <summary>
    /// 一次配线规格选型条目模型
    /// </summary>
    public class PrimaryWireSpecItem
    {
        // 允许承载的最大电流 (单位: A)
        [JsonPropertyName("maxCurrent")]
        public int MaxCurrent { get; set; }

        // 导线规格型号 (如 "BV-2.5", "BV-6")
        [JsonPropertyName("spec")]
        public string Spec { get; set; } = string.Empty;

        // 导线截面积 (单位: mm²)
        [JsonPropertyName("crossSection")]
        public double CrossSection { get; set; }

        // 每米单价 (单位: 元/米)
        [JsonPropertyName("pricePerMeter")]
        public double PricePerMeter { get; set; }
    }

    /// <summary>
    /// 一次导线推导消耗明细条目
    /// </summary>
    public class PrimaryWireUsageItem
    {
        // 导线规格 (如 "BV-2.5")
        [JsonPropertyName("spec")]
        public string Spec { get; set; } = string.Empty;

        // 截面积 (mm²)
        [JsonPropertyName("crossSection")]
        public double CrossSection { get; set; }

        // 计算消耗总长度 (单位: 米)
        [JsonPropertyName("lengthMeters")]
        public double LengthMeters { get; set; }

        // 单价 (元/米)
        [JsonPropertyName("pricePerMeter")]
        public double PricePerMeter { get; set; }

        // 费用小计 (单位: 元)
        [JsonPropertyName("subtotalCost")]
        public double SubtotalCost { get; set; }
    }

    /// <summary>
    /// 二次元件定额规则条目
    /// </summary>
    public class SecondaryElementRule
    {
        // 匹配关键字 (如 "接触器", "按钮")
        [JsonPropertyName("keyword")]
        public string Keyword { get; set; } = string.Empty;

        // 二次接线根数 (默认 2)
        [JsonPropertyName("wireCount")]
        public int WireCount { get; set; } = 2;

        // 二次配线每米单价 (单位: 元/米，默认 0.8)
        [JsonPropertyName("wirePrice")]
        public double WirePrice { get; set; } = 0.8;

        // 单只元件装配接线工价 (单位: 元/只，默认 5.0)
        [JsonPropertyName("laborPrice")]
        public double LaborPrice { get; set; } = 5.0;
    }

    /// <summary>
    /// 人工工价定额规则模型
    /// </summary>
    public class LaborConfig
    {
        // 壳体面积平铺制作工价系数 (单位: 元/分米² 即 元/0.01㎡，默认 2.95) --硬编码--
        [JsonPropertyName("areaBaseRate")]
        public double AreaBaseRate { get; set; } = 2.95;

        // 预留回路工价折减系数 (默认 0.4) --硬编码--
        [JsonPropertyName("reservedCircuitDiscount")]
        public double ReservedCircuitDiscount { get; set; } = 0.4;
    }

    /// <summary>
    /// 箱柜元器件扫描汇总实体
    /// </summary>
    public class CabinetScanData
    {
        // 箱柜所在工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 箱柜索引编号 (如 1, 2)
        public int CabinetIndex { get; set; }

        // 箱柜在顶部汇总行中的物理行号 (Cab_Sum)
        public int SumRow { get; set; }

        // 箱柜在明细表中的信息行号 (Cab_Det)
        public int DetRow { get; set; }

        // 元器件起始行号 (Cab_Det + 2)
        public int CompStartRow { get; set; }

        // 元器件结束行号 (Cab_Subsum - 1)
        public int CompEndRow { get; set; }

        // 小计行号 (Cab_Subsum)
        public int SubsumRow { get; set; }

        // 总计行号 (Cab_Tolsum)
        public int TolsumRow { get; set; }

        // 箱柜名称或编号 (如 AP1)
        public string CabinetName { get; set; } = string.Empty;

        // 箱柜台数 (默认 1)
        public int Quantity { get; set; } = 1;

        // 扫描提取的元器件明细项列表
        public List<CabinetComponentItem> Components { get; set; } = new List<CabinetComponentItem>();
    }

    /// <summary>
    /// 箱柜扫描元器件明细项
    /// </summary>
    public class CabinetComponentItem
    {
        // 所在物理行号
        public int RowIndex { get; set; }

        // 元件名称 (Column B)
        public string Name { get; set; } = string.Empty;

        // 型号规格 (Column C)
        public string Model { get; set; } = string.Empty;

        // 数量 (Column F)
        public int Quantity { get; set; } = 1;

        // 额定电流 (Column W 或通过型号解析，单位 A)
        public int Current { get; set; }

        // 极数 (Column X 或通过型号解析，如 3, 4, 1N 等)
        public string Poles { get; set; } = "3";

        // 极数数值 (解析后的数字，如 3P -> 3, 3P+N -> 4)
        public int PoleCount { get; set; } = 3;

        // 脱扣类型/脱扣方式 (Column Y)
        public string Trip { get; set; } = string.Empty;

        // 附件描述 (Column Z)
        public string Accessory { get; set; } = string.Empty;

        // 图块名称 (Column AA)
        public string BlockName { get; set; } = string.Empty;

        // 图块类别 (Column AB)
        public string BlockCategory { get; set; } = string.Empty;

        // 是否为双电源 ATS
        public bool IsAts { get; set; }

        // 是否为火灾互感器
        public bool IsFireTransformer { get; set; }

        // 是否为普通电流互感器
        public bool IsCurrentTransformer { get; set; }

        // 是否为预留回路
        public bool IsReserved { get; set; }
    }

    /// <summary>
    /// 单个箱柜智能推导计算结果
    /// </summary>
    public class CabinetCalcResult
    {
        // 箱柜编号与名称
        [JsonPropertyName("cabinetName")]
        public string CabinetName { get; set; } = string.Empty;

        // 箱柜在明细表中的信息行号 (Cab_Det)
        [JsonPropertyName("detRow")]
        public int DetRow { get; set; }

        // 小计行号
        [JsonPropertyName("subsumRow")]
        public int SubsumRow { get; set; }

        // 总计行号
        [JsonPropertyName("tolsumRow")]
        public int TolsumRow { get; set; }

        // 元件总占用排布面积 (单位: mm²)
        [JsonPropertyName("componentArea")]
        public double ComponentArea { get; set; }

        // 最大电流 (单位: A)
        [JsonPropertyName("maxCurrent")]
        public int MaxCurrent { get; set; }

        // 是否判定为落地柜体
        [JsonPropertyName("isCabinet")]
        public bool IsCabinet { get; set; }

        // 推荐匹配的壳体尺寸 (如 "600*800" 或 "800*1800")
        [JsonPropertyName("recommendedShellSize")]
        public string RecommendedShellSize { get; set; } = string.Empty;

        // 壳体回写目标位置说明 (如 "计费区域第 45 行" 或 "Cab_Det 信息行")
        [JsonPropertyName("shellTargetLocation")]
        public string ShellTargetLocation { get; set; } = string.Empty;

        // 壳体是否命中计费区域
        [JsonPropertyName("shellMatchedInFeeArea")]
        public bool ShellMatchedInFeeArea { get; set; }

        // 铜排总重量 (单位: KG)
        [JsonPropertyName("copperWeight")]
        public double CopperWeight { get; set; }

        // 铜排数量公式 (如 "=ROUND(18.6*1*1,1)")
        [JsonPropertyName("copperQtyFormula")]
        public string CopperQtyFormula { get; set; } = string.Empty;

        // 辅材费用总计 (单位: 元)
        [JsonPropertyName("auxiliaryCost")]
        public double AuxiliaryCost { get; set; }

        // 辅材金额公式 (如 "=ROUND(268.5*1*1,1)")
        [JsonPropertyName("auxiliaryFormula")]
        public string AuxiliaryFormula { get; set; } = string.Empty;

        // 人工费用总计 (单位: 元)
        [JsonPropertyName("laborCost")]
        public double LaborCost { get; set; }

        // 人工费用算式内容与动态公式 (如 "=ROUND((6*8*2.95+5*2)*1*1.13,1)")
        [JsonPropertyName("laborFormula")]
        public string LaborFormula { get; set; } = string.Empty;

        // 一次导线用量与费用明细列表
        [JsonPropertyName("primaryWireDetails")]
        public List<PrimaryWireUsageItem> PrimaryWireDetails { get; set; } = new List<PrimaryWireUsageItem>();

        // 铜排各分项算式明细列表 (展示主母排、各动态附件排、分支排的具体计算式与尺寸联动)
        [JsonPropertyName("copperFormulaDetails")]
        public List<string> CopperFormulaDetails { get; set; } = new List<string>();

        // 推导过程与说明明细
        [JsonPropertyName("description")]
        public string Description { get; set; } = string.Empty;
    }
}
