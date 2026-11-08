using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 箱体一键改价配置实体，定义多模式改价规则及生效范围
    /// 遵循规范：新增代码每 3 行至少包含 1 行中文注释
    /// </summary>
    public class CabinetBatchPriceConfigDto
    {
        /// <summary>
        /// 调价模式：area=面积单方造价重算, percent=比例浮动, fixed=固定单价 --硬编码: 模式枚举键--
        /// </summary>
        public string Mode { get; set; } = "area";

        /// <summary>
        /// 钣金每平方米综合造价 (元/㎡，包含冷轧板/敷铝锌板、加工、喷塑、五金辅件) --硬编码: 默认 320 元/㎡--
        /// </summary>
        public double AreaUnitPrice { get; set; } = 320.0;

        /// <summary>
        /// 箱体展开系数 (针对二层板、中隔板、安装横梁等结构的展开面积倍率，默认 1.20) --硬编码: 默认展开系数--
        /// </summary>
        public double ExpandRatio { get; set; } = 1.20;

        /// <summary>
        /// 箱体最低保底单价 (元/台，防止特小尺寸箱体算出超低不合常理的价格) --硬编码: 保底 100 元--
        /// </summary>
        public double MinUnitPrice { get; set; } = 100.0;

        /// <summary>
        /// 尺寸缺失时的兜底基准单价 (元/台) --硬编码: 默认 200 元--
        /// </summary>
        public double MissingSizePrice { get; set; } = 200.0;

        /// <summary>
        /// 比例浮动百分比 (如 +5 表示上调 5%，-8 表示下调 8%)
        /// </summary>
        public double PercentRatio { get; set; } = 0.0;

        /// <summary>
        /// 固定统一单价 (元/台，一口价模式) --硬编码: 默认 350 元--
        /// </summary>
        public double FixedUnitPrice { get; set; } = 350.0;

        /// <summary>
        /// PZ30 模数箱每位基准造价 (元/位) --硬编码: 默认 7.5 元/位--
        /// </summary>
        public double Pz30PricePerDigit { get; set; } = 7.5;

        /// <summary>
        /// 是否采用辅材壳体价格定额矩阵 (按材质×板厚高度阶梯精算，默认开启) --硬编码: 默认精算--
        /// </summary>
        public bool UseShellPriceRules { get; set; } = true;

        /// <summary>
        /// 批量计算基准默认材质 (镀锌板/冷轧板/不锈钢201/不锈钢304) --硬编码: 默认镀锌板--
        /// </summary>
        public string DefaultMaterial { get; set; } = "镀锌板";

        /// <summary>
        /// 关联的壳体价格定额规则配置对象 (材质单价库、高度阶梯、面数)
        /// </summary>
        public ShellPriceConfig ShellPriceRules { get; set; } = new ShellPriceConfig();

        /// <summary>
        /// 生效范围模式：current=当前工作表, all=全项目所有分类表 --硬编码: 范围枚举键--
        /// </summary>
        public string ScopeMode { get; set; } = "current";

        /// <summary>
        /// 指定生效的分类工作表名称列表
        /// </summary>
        public List<string> TargetSheets { get; set; } = new List<string>();
    }

    /// <summary>
    /// 单个箱体/壳体物料明细数据模型
    /// </summary>
    public class CabinetBatchItemDto
    {
        /// <summary>
        /// 唯一键标识 (如 分类1_Row85)
        /// </summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>
        /// 所属分类工作表名称
        /// </summary>
        public string SheetName { get; set; } = string.Empty;

        /// <summary>
        /// 箱柜序号索引 (1-based)
        /// </summary>
        public int CabinetIndex { get; set; }

        /// <summary>
        /// 所属箱柜名称 (如 APY1 动力配电箱)
        /// </summary>
        public string CabinetName { get; set; } = string.Empty;

        /// <summary>
        /// 箱体物料在 Excel 工作表中的物理行号
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 物料名称 (B列，如 箱体、配电箱壳体、动力柜体)
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 规格型号文本 (C列，如 800*1200*300 或 18位)
        /// </summary>
        public string Model { get; set; } = string.Empty;

        /// <summary>
        /// 计量单位 (D列，如 台、套、个)
        /// </summary>
        public string Unit { get; set; } = "台";

        /// <summary>
        /// 数量 (E列)
        /// </summary>
        public double Qty { get; set; } = 1.0;

        /// <summary>
        /// 箱体宽度 (mm)
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// 箱体高度 (mm)
        /// </summary>
        public double Height { get; set; }

        /// <summary>
        /// 箱体深度/厚度 (mm)
        /// </summary>
        public double Depth { get; set; }

        /// <summary>
        /// 是否成功提取到有效外形尺寸
        /// </summary>
        public bool HasValidSize { get; set; }

        /// <summary>
        /// 是否为 PZ30 模数箱体 (按位数计费)
        /// </summary>
        public bool IsModular { get; set; }

        /// <summary>
        /// PZ30 模数位数 (如 12、15、18、20、24、36)
        /// </summary>
        public int ModularDigits { get; set; }

        /// <summary>
        /// 箱体板材材质 (如 镀锌板、冷轧板、不锈钢201、不锈钢304)
        /// </summary>
        public string Material { get; set; } = "镀锌板";

        /// <summary>
        /// 箱体推荐/设定板材厚度 (mm，如 1.2, 1.5, 2.0)
        /// </summary>
        public double Thickness { get; set; } = 1.5;

        /// <summary>
        /// 该材质该厚度的每平米核算单价 (元/㎡)
        /// </summary>
        public double PricePerSqMeter { get; set; }

        /// <summary>
        /// 是否包含二层板 (含二层板时自动叠加 1.0 面宽高面积)
        /// </summary>
        public bool HasSecondPlate { get; set; }

        /// <summary>
        /// 符合工业规范的标准壳体型号 (如 800*600*200-1.2mm 镀锌板)
        /// </summary>
        public string StandardModel { get; set; } = string.Empty;

        /// <summary>
        /// 六面体展开表面积 (㎡，单台)
        /// </summary>
        public double SurfaceArea { get; set; }

        /// <summary>
        /// 修改前原基准单价 (M列原始值)
        /// </summary>
        public double OldUnitPrice { get; set; }

        /// <summary>
        /// 修改前原基准合价 (原单价 * 数量)
        /// </summary>
        public double OldTotalPrice { get; set; }

        /// <summary>
        /// 测算出的新基准单价 (M列待回写值)
        /// </summary>
        public double NewUnitPrice { get; set; }

        /// <summary>
        /// 测算出的新基准合价 (新单价 * 数量)
        /// </summary>
        public double NewTotalPrice { get; set; }

        /// <summary>
        /// 调价差额 (新合价 - 原合价，正数为涨价，负数为降价)
        /// </summary>
        public double DiffAmount { get; set; }

        /// <summary>
        /// 前端勾选状态，默认勾选
        /// </summary>
        public bool Selected { get; set; } = true;

        /// <summary>
        /// 改价计算备注/推导依据说明
        /// </summary>
        public string Remark { get; set; } = string.Empty;
    }

    /// <summary>
    /// 全项目箱体扫描结果传输响应模型
    /// </summary>
    public class CabinetBatchPriceScanResultDto
    {
        /// <summary>
        /// 扫描操作是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 状态或提示文本
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// 当前激活的 Excel 工作表名称
        /// </summary>
        public string CurrentSheetName { get; set; } = string.Empty;

        /// <summary>
        /// 工作簿内所有可用的分类工作表名称集合
        /// </summary>
        public List<string> AvailableSheets { get; set; } = new List<string>();

        /// <summary>
        /// 检测提取到的箱体物料明细集合
        /// </summary>
        public List<CabinetBatchItemDto> Rows { get; set; } = new List<CabinetBatchItemDto>();

        /// <summary>
        /// 涉及的箱柜总台数
        /// </summary>
        public int TotalCabinetCount { get; set; }

        /// <summary>
        /// 当前生效的辅材壳体价格定额规则配置 (包含材质单价库与高度阶梯)
        /// </summary>
        public ShellPriceConfig ShellPriceRules { get; set; } = new ShellPriceConfig();

        /// <summary>
        /// 默认基准材质
        /// </summary>
        public string DefaultMaterial { get; set; } = "镀锌板";
    }

    /// <summary>
    /// 前端执行一键改价请求载荷模型
    /// </summary>
    public class CabinetApplyPricePayload
    {
        /// <summary>
        /// 改价规则配置实体
        /// </summary>
        public CabinetBatchPriceConfigDto Config { get; set; } = new CabinetBatchPriceConfigDto();

        /// <summary>
        /// 待回写更新的箱体明细行集合
        /// </summary>
        public List<CabinetBatchItemDto> Items { get; set; } = new List<CabinetBatchItemDto>();
    }
}

