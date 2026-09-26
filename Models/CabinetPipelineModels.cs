using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 管道单条判定规则实体模型
    /// 用于定义三箱型号在管道中的前置判定条件与目标输出型号
    /// </summary>
    public class CabinetPipelineRule
    {
        // 规则唯一标识符
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        // 规则显示名称 (例如: "双电源箱 (ATS)", "落地式动力柜 (XL-21)")
        public string Name { get; set; } = string.Empty;

        // 判定命中后输出的箱柜型号 (例如: "ATS", "XL-21", "PZ30", "JXF")
        public string TargetModel { get; set; } = string.Empty;

        // 是否启用当前规则
        public bool IsEnabled { get; set; } = true;

        // 管道中的执行次序 (数字越小越先执行)
        public int SortOrder { get; set; } = 0;

        // 箱体高度判断模式: "none"(不限), ">=", "<=", "between", "=="
        public string HeightCondition { get; set; } = "none"; // --硬编码: 默认无高度限制--

        // 箱体高度最小值或基准值(mm)
        public int HeightMin { get; set; } = 0;

        // 箱体高度最大值(mm，用于 between 模式)
        public int HeightMax { get; set; } = 0;

        // 箱体深度判断模式: "none"(不限), ">=", "<=", "between", "=="
        public string DepthCondition { get; set; } = "none"; // --硬编码: 默认无深度限制--

        // 箱体深度最小值或基准值(mm)
        public int DepthMin { get; set; } = 0;

        // 箱体深度最大值(mm，用于 between 模式)
        public int DepthMax { get; set; } = 0;

        // 箱体宽度判断模式: "none"(不限), ">=", "<=", "between"
        public string WidthCondition { get; set; } = "none"; // --硬编码: 默认无宽度限制--

        // 箱体宽度最小值(mm)
        public int WidthMin { get; set; } = 0;

        // 箱体宽度最大值(mm)
        public int WidthMax { get; set; } = 0;

        // 元器件包含关键词集合 (例如: ["双电源", "ATS", "自动转换开关"])
        public List<string> ComponentKeywords { get; set; } = new List<string>();

        // 元器件关键词匹配模式: "any"(包含任意一个), "all"(必须包含全部), "none"(不限)
        public string ComponentMatchMode { get; set; } = "any"; // --硬编码: 默认包含任意一个--

        // 排除的元器件关键词集合 (若包含则直接不匹配，例如 PZ30 排除包含塑壳断路器)
        public List<string> ExcludeComponentKeywords { get; set; } = new List<string>();

        // 箱柜名称/柜号辅助包含关键词 (例如名称包含 "双电源", "应急", "照明")
        public List<string> NameKeywords { get; set; } = new List<string>();

        // 规则业务说明或备注
        public string Description { get; set; } = string.Empty;
    }

    /// <summary>
    /// 三箱型号管道全局配置持久化实体
    /// </summary>
    public class CabinetPipelineConfig
    {
        // 管道规则列表 (按优先级顺序排序)
        public List<CabinetPipelineRule> Rules { get; set; } = new List<CabinetPipelineRule>();

        // 兜底默认型号策略: "original"(保持原表格型号), "custom"(自定义指定), "pending"(标记待确认)
        public string FallbackStrategy { get; set; } = "original"; // --硬编码: 默认保留原值--

        // 当策略为 custom 时的指定兜底型号
        public string FallbackCustomModel { get; set; } = "待确认"; // --硬编码: 默认待确认--

        // 是否回写顶部汇总行 (Cab_Sum) 的 N 列 (第 14 列，用户明确指定目标列)
        public bool WriteToSummaryN { get; set; } = true;

        // 是否回写顶部汇总行 (Cab_Sum) 的 D 列 (保持兼容)
        public bool WriteToSummaryD { get; set; } = false;

        // 是否回写底部明细箱柜信息行 (Cab_Det) 的 D 列
        public bool WriteToDetailD { get; set; } = false;
    }

    /// <summary>
    /// 箱柜扫描明细与管道判断结果展示 DTO
    /// </summary>
    public class CabinetPipelineItemDto
    {
        // 所属工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 箱柜序号数字 K (例如 1 代表 Cab_Sum_1)
        public int CabinetIndex { get; set; }

        // 柜号文本 (如 "1AP1", "箱柜1")
        public string CabinetNo { get; set; } = string.Empty;

        // 箱柜名称文本 (如 "成套配电箱", "动力配电箱")
        public string CabinetName { get; set; } = string.Empty;

        // 当前 Excel 表格中已填写的型号 (原型号)
        public string CurrentModel { get; set; } = string.Empty;

        // 箱体宽度 (mm)
        public int Width { get; set; }

        // 箱体高度 (mm)
        public int Height { get; set; }

        // 箱体深度 (mm)
        public int Depth { get; set; }

        // 提取到的原始尺寸字符串 (例如 "800*600*200")
        public string RawDimensionText { get; set; } = string.Empty;

        // 是否成功识别到尺寸
        public bool HasDimension => Width > 0 && Height > 0;

        // 箱柜内部元器件摘要文字 (例如 "小型断路器*8, 塑壳断路器*1, 浪涌保护器*1")
        public string ComponentsSummary { get; set; } = string.Empty;

        // 内部元器件名称集合列表 (供前端快速匹配查看)
        public List<string> ComponentNames { get; set; } = new List<string>();

        // 内部元器件型号规格集合列表
        public List<string> ComponentModels { get; set; } = new List<string>();

        // 管道命中的规则 ID
        public string HitRuleId { get; set; } = string.Empty;

        // 管道命中的规则显示名称
        public string HitRuleName { get; set; } = string.Empty;

        // 管道判定推导出的建议型号
        public string SuggestedModel { get; set; } = string.Empty;

        // 用户最终确认要回写的型号 (可在前端手动改写)
        public string FinalModel { get; set; } = string.Empty;

        // 是否勾选执行回写
        public bool IsSelected { get; set; } = true;

        // 顶部汇总行物理行号 (Cab_Sum_k)
        public int SumRow { get; set; }

        // 底部明细箱柜信息行物理行号 (Cab_Det_k)
        public int DetRow { get; set; }
    }

    /// <summary>
    /// 前端初始化数据包 DTO
    /// </summary>
    public class CabinetPipelineInitDataDto
    {
        // 当前工作簿中所有有效的分类表名称列表
        public List<string> SheetNames { get; set; } = new List<string>();

        // 当前默认选中的分类表名称
        public string SelectedSheetName { get; set; } = string.Empty;

        // 管道判定规则列表
        public List<CabinetPipelineRule> Rules { get; set; } = new List<CabinetPipelineRule>();

        // 全局配置实体
        public CabinetPipelineConfig Config { get; set; } = new CabinetPipelineConfig();

        // 当前分类表扫描出的箱柜明细与判定结果列表
        public List<CabinetPipelineItemDto> Cabinets { get; set; } = new List<CabinetPipelineItemDto>();
    }

    /// <summary>
    /// 前端提交单项回写请求实体
    /// </summary>
    public class CabinetPipelineApplyItem
    {
        // 目标工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 箱柜序号数字
        public int CabinetIndex { get; set; }

        // 顶部汇总行物理行号
        public int SumRow { get; set; }

        // 底部明细信息行物理行号
        public int DetRow { get; set; }

        // 最终要写入的型号字符串
        public string Model { get; set; } = string.Empty;
    }

    /// <summary>
    /// 批量回写执行结果 DTO
    /// </summary>
    public class CabinetPipelineApplyResult
    {
        // 是否执行成功
        public bool Success { get; set; }

        // 提示消息
        public string Message { get; set; } = string.Empty;

        // 成功回写的箱柜数量
        public int UpdatedCount { get; set; }
    }
}
