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

        // 壳体价格计算与批量定价定额分节 (单列Tab)
        [JsonPropertyName("shellPriceRules")]
        public ShellPriceConfig ShellPriceRules { get; set; } = new ShellPriceConfig();
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

        // 参与一次导线与铜排计算的元器件自定义集合 (关键字列表) --硬编码--
        [JsonPropertyName("primaryCalcComponents")]
        public List<string> PrimaryCalcComponents { get; set; } = new List<string>
        {
            // 默认包含主流断路器、隔离开关与双电源等一次元件关键字
             "断路器", "漏电", "开关", "双电源"
        };
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

        // 纯高度驱动的箱柜深度推荐梯度表 (高度 <= maxHeight 则取 depth，与电流解耦) --硬编码--
        [JsonPropertyName("heightDepthGradients")]
        public List<HeightDepthGradientItem> HeightDepthGradients { get; set; } = new List<HeightDepthGradientItem>
        {
            // H <= 400 推荐深度 160mm
            new HeightDepthGradientItem { MaxHeight = 400, Depth = 160, CandidateDepths = new List<int>{ 120, 140, 160 }, Remark = "微型终端箱/照明箱" },
            // 400 < H <= 600 推荐深度 180mm
            new HeightDepthGradientItem { MaxHeight = 600, Depth = 180, CandidateDepths = new List<int>{ 160, 180, 200 }, Remark = "动力照明箱/基业箱" },
            // 600 < H <= 800 推荐深度 200mm
            new HeightDepthGradientItem { MaxHeight = 800, Depth = 200, CandidateDepths = new List<int>{ 180, 200, 220, 250 }, Remark = "中型动力箱/控制箱" },
            // 800 < H <= 1000 推荐深度 250mm
            new HeightDepthGradientItem { MaxHeight = 1000, Depth = 250, CandidateDepths = new List<int>{ 200, 250, 300 }, Remark = "大型壁挂配电箱" },
            // 1000 < H <= 1400 推荐深度 300mm
            new HeightDepthGradientItem { MaxHeight = 1400, Depth = 300, CandidateDepths = new List<int>{ 250, 300, 350 }, Remark = "超高挂墙箱/落地小箱" },
            // 1400 < H <= 1800 推荐深度 400mm
            new HeightDepthGradientItem { MaxHeight = 1800, Depth = 400, CandidateDepths = new List<int>{ 350, 400, 500 }, Remark = "动力配电柜(如XL-21)" },
            // 1800 < H <= 2000 推荐深度 800mm
            new HeightDepthGradientItem { MaxHeight = 2000, Depth = 800, CandidateDepths = new List<int>{ 600, 800, 1000 }, Remark = "标准低压柜(如GGD)" },
            // H > 2000 推荐深度 1000mm
            new HeightDepthGradientItem { MaxHeight = 9999, Depth = 1000, CandidateDepths = new List<int>{ 800, 1000, 1200 }, Remark = "大型低压成套开关柜" }
        };

        // 板材材质及各厚度每平米单价库 --硬编码--
        [JsonPropertyName("materialPrices")]
        public List<MaterialPriceItem> MaterialPrices { get; set; } = new List<MaterialPriceItem>
        {
            // 冷轧板单价阶梯
            new MaterialPriceItem
            {
                MaterialName = "冷轧板",
                Density = 7.85,
                ThicknessPrices = new List<MaterialThicknessPriceItem>
                {
                    // 1.0mm 单价 45.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 1.0, PricePerSqMeter = 45.0 },
                    // 1.2mm 单价 54.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 1.2, PricePerSqMeter = 54.0 },
                    // 1.5mm 单价 67.5 元/m²
                    new MaterialThicknessPriceItem { Thickness = 1.5, PricePerSqMeter = 67.5 },
                    // 2.0mm 单价 90.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 2.0, PricePerSqMeter = 90.0 },
                    // 2.5mm 单价 112.5 元/m²
                    new MaterialThicknessPriceItem { Thickness = 2.5, PricePerSqMeter = 112.5 },
                    // 3.0mm 单价 135.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 3.0, PricePerSqMeter = 135.0 }
                }
            },
            // 不锈钢201单价阶梯
            new MaterialPriceItem
            {
                MaterialName = "不锈钢201",
                Density = 7.93,
                ThicknessPrices = new List<MaterialThicknessPriceItem>
                {
                    // 1.0mm 单价 71.37 元/m²
                    new MaterialThicknessPriceItem { Thickness = 1.0, PricePerSqMeter = 71.37 },
                    // 1.2mm 单价 85.6 元/m²
                    new MaterialThicknessPriceItem { Thickness = 1.2, PricePerSqMeter = 85.6 },
                    // 1.5mm 单价 107.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 1.5, PricePerSqMeter = 107.0 },
                    // 2.0mm 单价 142.8 元/m²
                    new MaterialThicknessPriceItem { Thickness = 2.0, PricePerSqMeter = 142.8 },
                    // 2.5mm 单价 178.5 元/m²
                    new MaterialThicknessPriceItem { Thickness = 2.5, PricePerSqMeter = 178.5 },
                    // 3.0mm 单价 214.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 3.0, PricePerSqMeter = 214.0 }
                }
            },
            // 不锈钢304单价阶梯
            new MaterialPriceItem
            {
                MaterialName = "不锈钢304",
                Density = 7.93,
                ThicknessPrices = new List<MaterialThicknessPriceItem>
                {
                    // 1.0mm 单价 95.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 1.0, PricePerSqMeter = 95.0 },
                    // 1.2mm 单价 114.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 1.2, PricePerSqMeter = 114.0 },
                    // 1.5mm 单价 142.5 元/m²
                    new MaterialThicknessPriceItem { Thickness = 1.5, PricePerSqMeter = 142.5 },
                    // 2.0mm 单价 190.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 2.0, PricePerSqMeter = 190.0 },
                    // 2.5mm 单价 237.5 元/m²
                    new MaterialThicknessPriceItem { Thickness = 2.5, PricePerSqMeter = 237.5 },
                    // 3.0mm 单价 285.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 3.0, PricePerSqMeter = 285.0 }
                }
            },
            // 镀锌板单价阶梯
            new MaterialPriceItem
            {
                MaterialName = "镀锌板",
                Density = 7.85,
                ThicknessPrices = new List<MaterialThicknessPriceItem>
                {
                    // 1.0mm 单价 48.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 1.0, PricePerSqMeter = 48.0 },
                    // 1.2mm 单价 58.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 1.2, PricePerSqMeter = 58.0 },
                    // 1.5mm 单价 72.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 1.5, PricePerSqMeter = 72.0 },
                    // 2.0mm 单价 96.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 2.0, PricePerSqMeter = 96.0 },
                    // 2.5mm 单价 120.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 2.5, PricePerSqMeter = 120.0 },
                    // 3.0mm 单价 144.0 元/m²
                    new MaterialThicknessPriceItem { Thickness = 3.0, PricePerSqMeter = 144.0 }
                }
            }
        };

        // 箱体加工与折弯放量预留配置
        [JsonPropertyName("allowance")]
        public CabinetAllowanceConfig Allowance { get; set; } = new CabinetAllowanceConfig();
    }

    /// <summary>
    /// 壳体价格计算与批量定价定额配置模型 (与壳体尺寸生成彻底解耦)
    /// </summary>
    public class ShellPriceConfig
    {
        // 批量计算默认基准材质 (兜底，默认 "镀锌板") --硬编码--
        [JsonPropertyName("defaultMaterial")]
        public string DefaultMaterial { get; set; } = "镀锌板";

        // 顶底板展开面数 (W*D，默认 2.0) --硬编码--
        [JsonPropertyName("faceCountWD")]
        public double FaceCountWD { get; set; } = 2.0;

        // 门背板展开面数 (H*W，默认 2.0) --硬编码--
        [JsonPropertyName("faceCountHW")]
        public double FaceCountHW { get; set; } = 2.0;

        // 左右侧板展开面数 (H*D，默认 2.0) --硬编码--
        [JsonPropertyName("faceCountHD")]
        public double FaceCountHD { get; set; } = 2.0;

        // 二层板增加的宽高面数 (含"二层板"时自动叠加，默认 1.0) --硬编码--
        [JsonPropertyName("secondPlateExtraHW")]
        public double SecondPlateExtraHW { get; set; } = 1.0;

        // 批量计算板厚高度阶梯决策列表 (高度阶梯决策法) --硬编码--
        [JsonPropertyName("thicknessGradients")]
        public List<HeightThicknessGradientItem> ThicknessGradients { get; set; } = new List<HeightThicknessGradientItem>
        {
            // H <= 800mm (小型照明箱/终端箱) 对应 1.2mm
            new HeightThicknessGradientItem { MaxHeight = 800, Thickness = 1.2, Remark = "小型配电箱/照明箱" },
            // 800 < H <= 1600mm (中型动力箱/挂墙箱) 对应 1.5mm
            new HeightThicknessGradientItem { MaxHeight = 1600, Thickness = 1.5, Remark = "中型动力箱/挂墙箱" },
            // H > 1600mm (大型成套开关柜) 对应 2.0mm
            new HeightThicknessGradientItem { MaxHeight = 9999, Thickness = 2.0, Remark = "大型成套落地开关柜" }
        };

        // 板材材质及各厚度每平米单价库 --硬编码--
        [JsonPropertyName("materialPrices")]
        public List<MaterialPriceItem> MaterialPrices { get; set; } = new List<MaterialPriceItem>
        {
            // 镀锌板单价阶梯 (默认基准材质)
            new MaterialPriceItem
            {
                MaterialName = "镀锌板",
                Density = 7.85,
                ThicknessPrices = new List<MaterialThicknessPriceItem>
                {
                    new MaterialThicknessPriceItem { Thickness = 1.0, PricePerSqMeter = 48.0 },
                    new MaterialThicknessPriceItem { Thickness = 1.2, PricePerSqMeter = 58.0 },
                    new MaterialThicknessPriceItem { Thickness = 1.5, PricePerSqMeter = 72.0 },
                    new MaterialThicknessPriceItem { Thickness = 2.0, PricePerSqMeter = 96.0 },
                    new MaterialThicknessPriceItem { Thickness = 2.5, PricePerSqMeter = 120.0 },
                    new MaterialThicknessPriceItem { Thickness = 3.0, PricePerSqMeter = 144.0 }
                }
            },
            // 冷轧板单价阶梯
            new MaterialPriceItem
            {
                MaterialName = "冷轧板",
                Density = 7.85,
                ThicknessPrices = new List<MaterialThicknessPriceItem>
                {
                    new MaterialThicknessPriceItem { Thickness = 1.0, PricePerSqMeter = 45.0 },
                    new MaterialThicknessPriceItem { Thickness = 1.2, PricePerSqMeter = 54.0 },
                    new MaterialThicknessPriceItem { Thickness = 1.5, PricePerSqMeter = 67.5 },
                    new MaterialThicknessPriceItem { Thickness = 2.0, PricePerSqMeter = 90.0 },
                    new MaterialThicknessPriceItem { Thickness = 2.5, PricePerSqMeter = 112.5 },
                    new MaterialThicknessPriceItem { Thickness = 3.0, PricePerSqMeter = 135.0 }
                }
            },
            // 不锈钢201单价阶梯
            new MaterialPriceItem
            {
                MaterialName = "不锈钢201",
                Density = 7.93,
                ThicknessPrices = new List<MaterialThicknessPriceItem>
                {
                    new MaterialThicknessPriceItem { Thickness = 1.0, PricePerSqMeter = 71.37 },
                    new MaterialThicknessPriceItem { Thickness = 1.2, PricePerSqMeter = 85.6 },
                    new MaterialThicknessPriceItem { Thickness = 1.5, PricePerSqMeter = 107.0 },
                    new MaterialThicknessPriceItem { Thickness = 2.0, PricePerSqMeter = 142.8 },
                    new MaterialThicknessPriceItem { Thickness = 2.5, PricePerSqMeter = 178.5 },
                    new MaterialThicknessPriceItem { Thickness = 3.0, PricePerSqMeter = 214.0 }
                }
            },
            // 不锈钢304单价阶梯
            new MaterialPriceItem
            {
                MaterialName = "不锈钢304",
                Density = 7.93,
                ThicknessPrices = new List<MaterialThicknessPriceItem>
                {
                    new MaterialThicknessPriceItem { Thickness = 1.0, PricePerSqMeter = 95.0 },
                    new MaterialThicknessPriceItem { Thickness = 1.2, PricePerSqMeter = 114.0 },
                    new MaterialThicknessPriceItem { Thickness = 1.5, PricePerSqMeter = 142.5 },
                    new MaterialThicknessPriceItem { Thickness = 2.0, PricePerSqMeter = 190.0 },
                    new MaterialThicknessPriceItem { Thickness = 2.5, PricePerSqMeter = 237.5 },
                    new MaterialThicknessPriceItem { Thickness = 3.0, PricePerSqMeter = 285.0 }
                }
            }
        };
    }

    /// <summary>
    /// 高度对应板厚阶梯条目模型 (高度阶梯决策法)
    /// </summary>
    public class HeightThicknessGradientItem
    {
        // 高度区间上限 (单位: mm)
        [JsonPropertyName("maxHeight")]
        public int MaxHeight { get; set; }

        // 对应推荐板厚 (单位: mm)
        [JsonPropertyName("thickness")]
        public double Thickness { get; set; }

        // 适用箱柜类型与说明
        [JsonPropertyName("remark")]
        public string Remark { get; set; } = string.Empty;
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
    /// 元器件垂直预留高度加成映射规则模型
    /// </summary>
    public class ComponentExtraHeightRule
    {
        // 匹配元器件名称或型号的关键字 (如 "火灾", "互感器")
        [JsonPropertyName("keyword")]
        public string Keyword { get; set; } = string.Empty;

        // 增加的垂直尺寸高度 (单位: mm，默认 100) --硬编码--
        [JsonPropertyName("extraHeight")]
        public int ExtraHeight { get; set; } = 100;

        // 规则说明与用途描述
        [JsonPropertyName("description")]
        public string Description { get; set; } = string.Empty;
    }

    /// <summary>
    /// 固定长度元器件映射规则模型 (如接触器短跳线固定长度 300mm)
    /// </summary>
    public class FixedLengthComponentRule
    {
        // 匹配元器件名称或型号的关键字 (如 "接触器", "热继电器")
        [JsonPropertyName("keyword")]
        public string Keyword { get; set; } = string.Empty;

        // 单根固定接线导线长度 (单位: mm，默认 300) --硬编码--
        [JsonPropertyName("fixedLengthMm")]
        public int FixedLengthMm { get; set; } = 300;

        // 规则说明与用途描述
        [JsonPropertyName("description")]
        public string Description { get; set; } = string.Empty;
    }

    /// <summary>
    /// 一次导线长度生成与折算参数配置模型
    /// </summary>
    public class PrimaryWireLengthConfig
    {
        // 基础垂直预留高度 (单位: mm，默认 130) --硬编码--
        [JsonPropertyName("baseVerticalHeight")]
        public int BaseVerticalHeight { get; set; } = 130;

        // 火灾互感器垂直高度增量 (兼容保留字段，单位: mm，默认 100) --硬编码--
        [JsonPropertyName("fireTransformerExtraHeight")]
        public int FireTransformerExtraHeight { get; set; } = 100;

        // 普通互感器垂直高度增量 (兼容保留字段，单位: mm，默认 130) --硬编码--
        [JsonPropertyName("normalTransformerExtraHeight")]
        public int NormalTransformerExtraHeight { get; set; } = 130;

        // 元器件垂直预留高度加成映射规则列表 (单柜命中按类去重只加一次)
        [JsonPropertyName("extraHeightRules")]
        public List<ComponentExtraHeightRule> ExtraHeightRules { get; set; } = new List<ComponentExtraHeightRule>
        {
            new ComponentExtraHeightRule { Keyword = "火灾", ExtraHeight = 100, Description = "火灾漏电/探测元件垂直加成" },
            new ComponentExtraHeightRule { Keyword = "互感器", ExtraHeight = 130, Description = "普通电流互感器垂直加成" }
        };

        // 落地柜判定高度门限 (单位: mm，默认 1600) --硬编码--
        [JsonPropertyName("cabinetMinHeight")]
        public int CabinetMinHeight { get; set; } = 1600;

        // 落地柜一次导线柜宽系数 (默认 0.7) --硬编码--
        [JsonPropertyName("cabinetWidthFactor")]
        public double CabinetWidthFactor { get; set; } = 0.7;

        // 落地柜一次导线柜高比例系数 (默认 0.4) --硬编码--
        [JsonPropertyName("cabinetHeightFactor")]
        public double CabinetHeightFactor { get; set; } = 0.4;

        // 落地柜一次导线长度裕量放大系数 (默认 1.1) --硬编码--
        [JsonPropertyName("cabinetLengthMargin")]
        public double CabinetLengthMargin { get; set; } = 1.1;

        // 配电箱一次导线箱宽系数 (默认 0.6) --硬编码--
        [JsonPropertyName("boxWidthFactor")]
        public double BoxWidthFactor { get; set; } = 0.6;

        // 配电箱一次导线箱高比例系数 (默认 0.3) --硬编码--
        [JsonPropertyName("boxHeightFactor")]
        public double BoxHeightFactor { get; set; } = 0.3;

        // 配电箱一次导线长度裕量放大系数 (默认 1.05) --硬编码--
        [JsonPropertyName("boxLengthMargin")]
        public double BoxLengthMargin { get; set; } = 1.05;
    }

    /// <summary>
    /// 辅材与配线定额规则模型
    /// </summary>
    public class AuxConfig
    {
        // 辅材匹配名称 (默认 "辅材"，优先在计费区查找，找不到在元器件区查找) --硬编码--
        [JsonPropertyName("auxMatchName")]
        public string AuxMatchName { get; set; } = "辅材";

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

        // 二次配线每米单价 (单位: 元/米，默认 0.8) --硬编码--
        [JsonPropertyName("secondaryWirePrice")]
        public double SecondaryWirePrice { get; set; } = 0.8;

        // 二次跨门导线柜宽折算系数 (默认 0.8) --硬编码--
        [JsonPropertyName("secondaryWidthFactor")]
        public double SecondaryWidthFactor { get; set; } = 0.8;

        // 二次跨门导线柜高比例折算系数 (默认 0.3) --硬编码--
        [JsonPropertyName("secondaryHeightFactor")]
        public double SecondaryHeightFactor { get; set; } = 0.3;

        // 二次跨门导线端头剥线与接线预留裕量 (单位: mm，默认 300) --硬编码--
        [JsonPropertyName("secondaryMarginLength")]
        public int SecondaryMarginLength { get; set; } = 300;

        // 旧版二次元件接线定额库 (已升级为二次方案绑定，此处保留字段做向后兼容兜底)
        [JsonPropertyName("secondaryElements")]
        public List<SecondaryElementRule> SecondaryElements { get; set; } = new List<SecondaryElementRule>();

        // 固定长度元器件接线映射规则列表 (短跳线免计算箱体宽高，如接触器 300mm)
        [JsonPropertyName("fixedLengthRules")]
        public List<FixedLengthComponentRule> FixedLengthRules { get; set; } = new List<FixedLengthComponentRule>
        {
            new FixedLengthComponentRule { Keyword = "接触器", FixedLengthMm = 300, Description = "接触器短接跳线" }
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
        // 人工匹配名称 (默认 "人工费"，优先在计费区查找，找不到在元器件区查找) --硬编码--
        [JsonPropertyName("laborMatchName")]
        public string LaborMatchName { get; set; } = "人工费";

        // 壳体面积平铺制作工价系数 (单位: 元/分米² 即 元/0.01㎡，默认 2.95) --硬编码--
        [JsonPropertyName("areaBaseRate")]
        public double AreaBaseRate { get; set; } = 2.95;

        // 预留回路工价折减系数 (默认 0.4) --硬编码--
        [JsonPropertyName("reservedCircuitDiscount")]
        public double ReservedCircuitDiscount { get; set; } = 0.4;

        // 预留回路打折判定的整柜断路器最大台数门限 (单位: 台，默认 3) --硬编码--
        [JsonPropertyName("reservedMaxBreakersThreshold")]
        public int ReservedMaxBreakersThreshold { get; set; } = 3;
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

        // 第 32 列 (AF 列) 绑定的回路代号或图号 (如 "CA1B", "接触器变频器")
        public string BoundDwgCode { get; set; } = string.Empty;

        // 是否为二次元件组 (B列='元件组' 或以 '*' 开头，或第 32 列绑定了图号)
        public bool IsComponentGroup { get; set; }

        // 元器件绑定的图纸所属目录名称 (来自 Excel 对应配置列，默认 Y 列)
        public string DwgDir { get; set; } = string.Empty;

        // 元器件绑定的 DWG 图纸文件名称 (来自 Excel 对应配置列，默认 X 列)
        public string DwgName { get; set; } = string.Empty;

        // 真实 CAD 外形宽度 (单位: mm，0 表示未获取)
        public double RealWidth { get; set; } = 0.0;

        // 真实 CAD 外形高度 (单位: mm，0 表示未获取)
        public double RealHeight { get; set; } = 0.0;

        // 从图纸文字中解析出的真实安装进深/厚度 (单位: mm，0 表示未获取)
        public double RealDepth { get; set; } = 0.0;

        // 标记该元器件是否成功命中并采用了 CAD 真实外形尺寸
        public bool HasRealDimensions { get; set; } = false;
    }

    /// <summary>
    /// 单个箱柜命中的二次回路方案定额计算条目
    /// </summary>
    public class SecondarySchemeCalcItem
    {
        // 方案 ID
        [JsonPropertyName("schemeId")]
        public int SchemeId { get; set; }

        // 方案名称 (如 "双电源标准控制回路方案A")
        [JsonPropertyName("schemeName")]
        public string SchemeName { get; set; } = string.Empty;

        // 绑定的回路代号或图号 (如 "CA1B")
        [JsonPropertyName("circuitCode")]
        public string CircuitCode { get; set; } = string.Empty;

        // 该方案在箱柜中出现的数量 / 回路套数 (默认 1) --硬编码--
        [JsonPropertyName("quantity")]
        public int Quantity { get; set; } = 1;

        // 方案级二次线跨门根数 (单套)
        [JsonPropertyName("crossDoorCount")]
        public double CrossDoorCount { get; set; }

        // 单套回路二次导线推导长度 (单位: 米)
        [JsonPropertyName("singleWireLength")]
        public double SingleWireLength { get; set; }

        // 该方案二次导线累计推导总长度 (单位: 米)
        [JsonPropertyName("totalWireLength")]
        public double TotalWireLength { get; set; }

        // 二次配线辅材费用小计 (元)
        [JsonPropertyName("wireCost")]
        public double WireCost { get; set; }

        // 单套二次装配与接线人工工费 (元)
        [JsonPropertyName("unitLaborCost")]
        public double UnitLaborCost { get; set; }

        // 二次装配与接线人工工费小计 (元)
        [JsonPropertyName("laborCost")]
        public double LaborCost { get; set; }
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

        // 辅材回写目标位置说明 (如 "计费区域第 46 行" 或 "元器件区域第 30 行")
        [JsonPropertyName("auxTargetLocation")]
        public string AuxTargetLocation { get; set; } = string.Empty;

        // 人工回写目标位置说明 (如 "计费区域第 47 行" 或 "元器件区域第 31 行")
        [JsonPropertyName("laborTargetLocation")]
        public string LaborTargetLocation { get; set; } = string.Empty;

        // 铜排回写目标位置说明 (如 "元器件区域第 32 行")
        [JsonPropertyName("copperTargetLocation")]
        public string CopperTargetLocation { get; set; } = string.Empty;

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

        // 命中的二次回路方案计算明细列表
        [JsonPropertyName("secondarySchemeDetails")]
        public List<SecondarySchemeCalcItem> SecondarySchemeDetails { get; set; } = new List<SecondarySchemeCalcItem>();

        // 单根二次跨门线基准推导长度 (单位: 米)
        [JsonPropertyName("secondarySingleWireLength")]
        public double SecondarySingleWireLength { get; set; }

        // 箱柜二次跨门导线总根数
        [JsonPropertyName("secondaryTotalWireCount")]
        public double SecondaryTotalWireCount { get; set; }

        // 箱柜二次导线推导消耗总长度 (单位: 米)
        [JsonPropertyName("secondaryTotalWireLength")]
        public double SecondaryTotalWireLength { get; set; }

        // 铜排各分项算式明细列表 (展示主母排、各动态附件排、分支排的具体计算式与尺寸联动)
        [JsonPropertyName("copperFormulaDetails")]
        public List<string> CopperFormulaDetails { get; set; } = new List<string>();

        // 推导过程与说明明细
        [JsonPropertyName("description")]
        public string Description { get; set; } = string.Empty;

        // 纯高度推导的推荐箱柜深度 (单位: mm，与电流解耦)
        [JsonPropertyName("recommendedShellDepth")]
        public int RecommendedShellDepth { get; set; } = 200;

        // 包含深度的完整推荐壳体尺寸 (如 800*2200*1000)
        [JsonPropertyName("recommendedShellSizeFull")]
        public string RecommendedShellSizeFull { get; set; } = string.Empty;

        // 推荐匹配的壳体标准拼装型号 (如 "XM-800*2200*1000-2.0mm 镀锌板")
        [JsonPropertyName("recommendedShellModel")]
        public string RecommendedShellModel { get; set; } = string.Empty;

        // 自动计算出的壳体单价 (单位: 元)
        [JsonPropertyName("recommendedShellUnitPrice")]
        public double RecommendedShellUnitPrice { get; set; }

        // 自动计算出的壳体展开总面积 (单位: m²)
        [JsonPropertyName("recommendedShellExpandedArea")]
        public double RecommendedShellExpandedArea { get; set; }

        // 自动匹配的壳体材质
        [JsonPropertyName("recommendedShellMaterial")]
        public string RecommendedShellMaterial { get; set; } = string.Empty;

        // 自动匹配的壳体板厚 (单位: mm)
        [JsonPropertyName("recommendedShellThickness")]
        public double RecommendedShellThickness { get; set; }

        // 未填写电流等计算警告与提醒列表 (如提示某些断路器W列为空未计入计算)
        [JsonPropertyName("warnings")]
        public List<string> Warnings { get; set; } = new List<string>();

        // 整柜元器件中提取出的最大安装进深/厚度 (单位: mm)
        [JsonPropertyName("maxComponentDepth")]
        public double MaxComponentDepth { get; set; } = 0.0;

        // 根据元器件最大进深计算的安全深度门限要求 (MaxComponentDepth + 60mm)
        [JsonPropertyName("minRequiredDepth")]
        public double MinRequiredDepth { get; set; } = 0.0;

        // 整柜中成功命中并采用 CAD 真实外形尺寸的元器件项数
        [JsonPropertyName("realDimensionsCount")]
        public int RealDimensionsCount { get; set; } = 0;
    }

    /// <summary>
    /// 高度对应深度推导梯度条目模型 (纯高度驱动，无电流关联)
    /// </summary>
    public class HeightDepthGradientItem
    {
        // 高度上限门限 (单位: mm，当 Height <= MaxHeight 时命中)
        [JsonPropertyName("maxHeight")]
        public int MaxHeight { get; set; }

        // 对应推荐的深度数值 (单位: mm)
        [JsonPropertyName("depth")]
        public int Depth { get; set; }

        // 常用备选深度列表 (如 160, 180, 200)
        [JsonPropertyName("candidateDepths")]
        public List<int> CandidateDepths { get; set; } = new List<int>();

        // 描述或应用场景说明
        [JsonPropertyName("remark")]
        public string Remark { get; set; } = string.Empty;
    }

    /// <summary>
    /// 板材材质各厚度单价条目模型
    /// </summary>
    public class MaterialThicknessPriceItem
    {
        // 板材厚度 (单位: mm，如 1.0, 1.2, 1.5, 2.0)
        [JsonPropertyName("thickness")]
        public double Thickness { get; set; }

        // 对应平方单价 (单位: 元/m²)
        [JsonPropertyName("pricePerSqMeter")]
        public double PricePerSqMeter { get; set; }
    }

    /// <summary>
    /// 板材材质分类与厚度单价配置模型
    /// </summary>
    public class MaterialPriceItem
    {
        // 材质名称 (如 "冷轧板", "不锈钢201", "不锈钢304", "镀锌板")
        [JsonPropertyName("materialName")]
        public string MaterialName { get; set; } = string.Empty;

        // 材料密度 (g/cm³，默认 7.85)
        [JsonPropertyName("density")]
        public double Density { get; set; } = 7.85;

        // 各厚度规格单价列表
        [JsonPropertyName("thicknessPrices")]
        public List<MaterialThicknessPriceItem> ThicknessPrices { get; set; } = new List<MaterialThicknessPriceItem>();
    }

    /// <summary>
    /// 箱体钣金加工与折弯放量预留配置
    /// </summary>
    public class CabinetAllowanceConfig
    {
        // 宽度加工预留放量 (单位: mm，默认 50) --硬编码--
        [JsonPropertyName("widthAllowance")]
        public int WidthAllowance { get; set; } = 50;

        // 高度加工预留放量 (单位: mm，默认 50) --硬编码--
        [JsonPropertyName("heightAllowance")]
        public int HeightAllowance { get; set; } = 50;

        // 深度加工预留放量 (单位: mm，默认 20) --硬编码--
        [JsonPropertyName("depthAllowance")]
        public int DepthAllowance { get; set; } = 20;
    }

    /// <summary>
    /// 箱体单面结构计算模型 (如面门板、后背板、左右侧板、顶底板、二层板)
    /// </summary>
    public class CabinetStructureItem
    {
        // 结构部位名称 (如 "面门板", "后背板", "左右侧板", "顶板底板", "二层板")
        [JsonPropertyName("partName")]
        public string PartName { get; set; } = string.Empty;

        // 材质 (如 "冷轧板", "不锈钢201")
        [JsonPropertyName("material")]
        public string Material { get; set; } = "冷轧板";

        // 板厚 (单位: mm，如 1.5)
        [JsonPropertyName("thickness")]
        public double Thickness { get; set; } = 1.5;

        // 单价 (单位: 元/m²)
        [JsonPropertyName("unitPrice")]
        public double UnitPrice { get; set; } = 67.5;

        // 面积公式代码 ("HW", "HD", "WD")
        [JsonPropertyName("formulaCode")]
        public string FormulaCode { get; set; } = "HW";

        // 公式展示文本 (如 "*(H*W)", "*(H*D)", "*(W*D)")
        [JsonPropertyName("formulaText")]
        public string FormulaText { get; set; } = "*(H*W)";

        // 数量面数系数 (如 1, 2, 0)
        [JsonPropertyName("coef")]
        public int Coef { get; set; } = 1;
    }

    /// <summary>
    /// 回写算料壳体至 Excel 计费区域的请求载荷模型
    /// </summary>
    public class CabinetShellWritePayload
    {
        // 目标箱柜的 Det 定义名称 (如 "Cab_Det_1")
        [JsonPropertyName("cabDetName")]
        public string CabDetName { get; set; } = string.Empty;

        // 箱体物料名称 (默认 "柜体" 或 "箱体")
        [JsonPropertyName("itemName")]
        public string ItemName { get; set; } = "柜体";

        // 拼装好的完整规格型号 (如 "XM-800*2200*1000-1.5mm 冷轧板 户内明装")
        [JsonPropertyName("model")]
        public string Model { get; set; } = string.Empty;

        // 品牌
        [JsonPropertyName("brand")]
        public string Brand { get; set; } = string.Empty;

        // 单位 (默认 "台")
        [JsonPropertyName("unit")]
        public string Unit { get; set; } = "台";

        // 数量 (默认 1)
        [JsonPropertyName("quantity")]
        public double Quantity { get; set; } = 1.0;

        // 单价 (算出的总价格)
        [JsonPropertyName("unitPrice")]
        public double UnitPrice { get; set; }

        // 展开面积 (单位: m²)
        [JsonPropertyName("expandedArea")]
        public double ExpandedArea { get; set; }

        // 回写动作模式 ("replace" 替换现有壳体行，"insert" 插入新行，"auto" 优先替换未找到则插入)
        [JsonPropertyName("writeMode")]
        public string WriteMode { get; set; } = "auto";
    }
}
