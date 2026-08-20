using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 流水线单项匹配规则模型
    /// </summary>
    public class PipelineRuleItem
    {
        // 规则唯一标识符
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        // 规则标题名称（如“寻找 P 前面的数字”）
        public string Title { get; set; } = string.Empty;

        // 是否启用当前规则
        public bool Enabled { get; set; } = true;

        // 匹配模式代号 (如 FindBeforeP, CodeMapping, FindBeforeJi 等)
        public string Mode { get; set; } = string.Empty;

        // 脱扣代号映射表 (针对 CodeMapping 模式，例如 {"/3300": "3P", "/4300": "4P"})
        public Dictionary<string, string> Mapping { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // 固定取值 (针对 FixedValue 模式，如固定为 3P)
        public string FixedValue { get; set; } = string.Empty;

        // 自定义正则表达式 (针对 CustomRegex 模式，如 \b(?<pole>[1-4]P(?:\+N)?)\b)
        public string CustomRegex { get; set; } = string.Empty;
    }

    /// <summary>
    /// 元器件型号参数提取完整配置模型
    /// </summary>
    public class ModelParserConfig
    {
        // 默认源数据型号所在列 (如 C 列)
        public string SourceColumn { get; set; } = "C"; // --硬编码-- 默认源型号列

        // 默认提取后电流输出的目标列 (如 S 列)
        public string CurrentColumn { get; set; } = "S"; // --硬编码-- 默认电流输出列

        // 默认提取后极数输出的目标列 (如 T 列)
        public string PoleColumn { get; set; } = "T"; // --硬编码-- 默认极数输出列

        // 默认处理的起始行号 (表头下方第 2 行开始)
        public int StartRow { get; set; } = 2; // --硬编码-- 默认起始数据行

        // 处理范围模式: "AllValid" (整表有效行) 或 "Selection" (仅选中区域)
        public string ProcessRangeMode { get; set; } = "AllValid";

        // ===================== 极数通道配置 =====================

        // 极数前置必去项/负向排除关键词列表 (如 IP 防护等级、DPN 前缀等)
        public List<string> PoleExcludeKeywords { get; set; } = new List<string>
        {
            "IP", "DPN", "PF", "PT", "PIN", "PAGE", "PRO" // --硬编码-- 默认极数排除词
        };

        // 极数后置标称白名单有效值集合
        public List<string> PoleAllowedValues { get; set; } = new List<string>
        {
            "1P", "2P", "3P", "4P", "1P+N", "3P+N", "1", "2", "3", "4" // --硬编码-- 默认极数白名单
        };

        // 是否开启极数严格白名单校验
        public bool EnableStrictPoleWhitelist { get; set; } = true;

        // 极数多级顺位匹配流水线规则列表
        public List<PipelineRuleItem> PolePipeline { get; set; } = new List<PipelineRuleItem>();

        // 极数输出格式: "WithP" (如 3P) 或 "NumberOnly" (如 3)
        public string PoleFormat { get; set; } = "WithP"; // --硬编码-- 默认极数输出格式

        // ===================== 电流通道配置 =====================

        // 电流前置必去项/负向排除关键词列表 (如 mA 漏电电流、kA 分断能力、电压等)
        public List<string> CurrentExcludeKeywords { get; set; } = new List<string>
        {
            "mA", "毫安", "kA", "千安", "VAC", "VDC", "V", "Hz", "ms" // --硬编码-- 默认电流排除词
        };

        // 电流后置标称白名单有效值集合 (工业常用标称电流)
        public List<string> CurrentAllowedValues { get; set; } = new List<string>
        {
            "0.5", "1", "1.6", "2", "2.5", "3", "4", "5", "6", "10", "16", "20", "25",
            "32", "40", "50", "63", "80", "100", "125", "140", "160", "180", "200",
            "225", "250", "315", "350", "400", "500", "630", "800", "1000", "1250",
            "1600", "2000", "2500", "3200", "4000", "6300" // --硬编码-- 默认标称电流白名单
        };

        // 是否开启电流严格白名单校验 (不在列表内如 99 将被拒绝)
        public bool EnableStrictCurrentWhitelist { get; set; } = true;

        // 电流多级顺位匹配流水线规则列表
        public List<PipelineRuleItem> CurrentPipeline { get; set; } = new List<PipelineRuleItem>();

        // 电流输出格式: "NumberOnly" (如 100) 或 "WithA" (如 100A)
        public string CurrentFormat { get; set; } = "NumberOnly"; // --硬编码-- 默认电流输出格式

        /// <summary>
        /// 创建带有出厂默认规则的配置对象
        /// </summary>
        public static ModelParserConfig CreateDefault()
        {
            // 初始化基础配置实例
            var cfg = new ModelParserConfig();

            // 1. 初始化极数多级顺位流水线
            cfg.PolePipeline.Add(new PipelineRuleItem
            {
                // 顺位 1：寻找 P 前面的数字 (如 3P, 4P, 1P+N)
                Title = "寻找 'P' 前面的数字 (如 3P, 4P)",
                Mode = "FindBeforeP",
                Enabled = true
            });

            // 构造塑壳脱扣代号默认对照表
            var mapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "/3300", "3P" }, // --硬编码-- 正泰/德力西塑壳 3极
                { "/4300", "4P" }, // --硬编码-- 4极
                { "/3320", "3P" }, // --硬编码-- 3极带辅助
                { "/4320", "4P" }, // --硬编码-- 4极带辅助
                { "/33002", "3P" }, // --硬编码-- 电操3极
                { "/43002", "4P" }, // --硬编码-- 电操4极
                { " 3P", "3P" },
                { " 4P", "4P" }
            };

            cfg.PolePipeline.Add(new PipelineRuleItem
            {
                // 顺位 2：脱扣代号与特殊特征映射 (如 /3300 -> 3P)
                Title = "脱扣代号/特征映射 (如 /3300->3P)",
                Mode = "CodeMapping",
                Enabled = true,
                Mapping = mapping
            });

            cfg.PolePipeline.Add(new PipelineRuleItem
            {
                // 顺位 3：寻找中文“极”前面的数字 (如 3极, 4极)
                Title = "寻找中文 '极' 前面的数字 (如 3极)",
                Mode = "FindBeforeJi",
                Enabled = true
            });

            cfg.PolePipeline.Add(new PipelineRuleItem
            {
                // 顺位 4：寻找斜杠后面的单数字极数 (如 /3, /4 框架开关)
                Title = "寻找斜杠后极数 (如 /3, /4)",
                Mode = "FrameCode",
                Enabled = true
            });

            // 2. 初始化电流多级顺位流水线
            cfg.CurrentPipeline.Add(new PipelineRuleItem
            {
                // 顺位 1：寻找带单位 A 的数字 (如 100A, 63A)
                Title = "寻找带单位 'A' 的数字 (如 100A)",
                Mode = "NumberWithA",
                Enabled = true
            });

            cfg.CurrentPipeline.Add(new PipelineRuleItem
            {
                // 顺位 2：寻找脱扣曲线字母 [C/D/K/Z] 后面的数字 (如 C32, D16)
                Title = "寻找脱扣曲线 [C/D] 后的数字 (如 C32)",
                Mode = "CurveLetterNumber",
                Enabled = true
            });

            cfg.CurrentPipeline.Add(new PipelineRuleItem
            {
                // 顺位 3：寻找整定电流可调区间 (如 2.5-4A, 0.63-1A 电机保护器)
                Title = "寻找整定电流区间 (如 2.5-4A)",
                Mode = "CurrentRange",
                Enabled = true
            });

            cfg.CurrentPipeline.Add(new PipelineRuleItem
            {
                // 顺位 4：寻找末尾纯数字 (如 NM1-125 100)
                Title = "寻找末尾纯数字 (如 NM1-125 100)",
                Mode = "TrailingNumber",
                Enabled = true
            });

            // 返回构造完成的默认配置对象
            return cfg;
        }
    }

    /// <summary>
    /// 单个型号解析结果数据传输对象
    /// </summary>
    public class ParseResultDto
    {
        // 原始输入型号
        public string RawModel { get; set; } = string.Empty;

        // 解析提取出的极数
        public string Pole { get; set; } = string.Empty;

        // 解析提取出的电流
        public string Current { get; set; } = string.Empty;

        // 命中的极数规则标题
        public string HitPoleRule { get; set; } = string.Empty;

        // 命中的电流规则标题
        public string HitCurrentRule { get; set; } = string.Empty;

        // 解析状态标识: Success (两者成功), Partial (部分成功), Failed (均失败)
        public string Status { get; set; } = "Failed";
    }

    /// <summary>
    /// Excel 批量解析与回填执行统计结果
    /// </summary>
    public class BatchParseResult
    {
        // 是否执行成功
        public bool Success { get; set; } = true;

        // 处理的总数据行数
        public int TotalRows { get; set; } = 0;

        // 完全解析成功的行数 (极数与电流均提取到)
        public int SuccessCount { get; set; } = 0;

        // 部分或完全未识别的行数
        public int FailedCount { get; set; } = 0;

        // 执行总耗时 (毫秒)
        public long ElapsedMilliseconds { get; set; } = 0;

        // 结果摘要消息
        public string Message { get; set; } = string.Empty;
    }
}
