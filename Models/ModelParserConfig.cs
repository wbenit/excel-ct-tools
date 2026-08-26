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

        // 脱扣代号映射表 (针对 CodeMapping 模式，例如 {"/3300": "3P", "/4300": "4P"})，区分大小写
        public Dictionary<string, string> Mapping { get; set; } = new Dictionary<string, string>();

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

        // 默认提取后最小电流输出的目标列 (如 S 列)
        public string MinCurrentColumn { get; set; } = "S"; // --硬编码-- 默认最小电流输出列

        // 默认提取后最大电流输出的目标列 (可自由配置如 S 或 T 列，留空则不写入)
        public string MaxCurrentColumn { get; set; } = string.Empty; // --硬编码-- 默认最大电流输出列

        // 兼容性字段：电流输出列 (自动重定向至 MinCurrentColumn)
        public string CurrentColumn
        {
            get => MinCurrentColumn;
            set => MinCurrentColumn = value;
        }

        // 默认提取后极数输出的目标列 (如 T 列)
        public string PoleColumn { get; set; } = "T"; // --硬编码-- 默认极数输出列

        // 默认提取后脱扣方式输出的目标列 (如 U 列)
        public string TripModeColumn { get; set; } = "U"; // --硬编码-- 默认脱扣方式输出列

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

        // 匹配到多个电流或区间电流时是否选择最大值 (true: 取最大值，false: 取最小值)
        public bool PreferMaxCurrent { get; set; } = true; // --硬编码-- 默认取最大值

        // ===================== 脱扣方式通道配置 =====================

        // 脱扣方式前置必去项/负向排除关键词列表
        public List<string> TripModeExcludeKeywords { get; set; } = new List<string>();

        // 脱扣方式后置标称白名单有效值集合 (输出简写代号)
        public List<string> TripModeAllowedValues { get; set; } = new List<string>
        {
            "TM", "TMD", "TMA", "MA", "Elec", "C", "D", "B", "K", "Z", "LE" // --硬编码-- 默认脱扣方式白名单
        };

        // 是否开启脱扣方式严格白名单校验
        public bool EnableStrictTripModeWhitelist { get; set; } = false;

        // 脱扣方式多级顺位匹配流水线规则列表
        public List<PipelineRuleItem> TripModePipeline { get; set; } = new List<PipelineRuleItem>();

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

            // 构造塑壳脱扣代号默认对照表 (区分大小写)
            var mapping = new Dictionary<string, string>
            {
                { "3300", "3P" }, // --硬编码-- 正泰/德力西塑壳 3极
                { "4300", "4P" }, // --硬编码-- 4极
                { "3320", "3P" }, // --硬编码-- 3极带辅助
                { "4320", "4P" }, // --硬编码-- 4极带辅助
                { "33002", "3P" }, // --硬编码-- 电操3极
                { "43002", "4P" }, // --硬编码-- 电操4极
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

            // 3. 初始化脱扣方式多级顺位流水线 (输出标准简写代号，区分大小写)
            var gbTripMapping = new Dictionary<string, string>
            {
                { "33002", "D" },  // --硬编码-- 3极电操热磁
                { "43002", "D" },  // --硬编码-- 4极电操热磁
                { "3200", "MA" },   // --硬编码-- 3极单电磁
                { "4200", "MA" },   // --硬编码-- 4极单电磁
                { "32002", "DMA" }   // --硬编码-- 3极单电磁带辅助
            };

            cfg.TripModePipeline.Add(new PipelineRuleItem
            {
                // 顺位 1：国标 4 位脱扣代号映射
                Title = "国标 4 位脱扣代号映射 (如 /3300->TM)",
                Mode = "CodeMapping",
                Enabled = true,
                Mapping = gbTripMapping
            });

            var brandTripMapping = new Dictionary<string, string>
            {
                { "MA", "DMA" } // --硬编码-- 可调热磁
            };

            cfg.TripModePipeline.Add(new PipelineRuleItem
            {
                // 顺位 2：合资/外资品牌脱扣器代号映射
                Title = "品牌脱扣器代号映射 (如 TMD->TMD, Ekip->Elec)",
                Mode = "CodeMapping",
                Enabled = true,
                Mapping = brandTripMapping
            });

            cfg.TripModePipeline.Add(new PipelineRuleItem
            {
                // 顺位 3：微断脱扣特性曲线提取 (C/D/B/K/Z)
                Title = "微断脱扣特性曲线提取 (如 C16->C, D20->D)",
                Mode = "CurveLetter",
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

        // 解析提取出的最小电流 (如 2.5 或 100)
        public string MinCurrent { get; set; } = string.Empty;

        // 解析提取出的最大电流 (如 4 或 100)
        public string MaxCurrent { get; set; } = string.Empty;

        // 兼容性字段：电流 (默认与 MinCurrent 一致)
        public string Current { get; set; } = string.Empty;

        // 解析提取出的脱扣方式简写代号 (如 TM, C, D, MA, Elec, LE)
        public string TripMode { get; set; } = string.Empty;

        // 命中的极数规则标题
        public string HitPoleRule { get; set; } = string.Empty;

        // 命中的电流规则标题
        public string HitCurrentRule { get; set; } = string.Empty;

        // 命中的脱扣方式规则标题
        public string HitTripModeRule { get; set; } = string.Empty;

        // 识别到的所有候选电流列表 (用于前端沙盒展示多匹配情况)
        public List<string> CandidateCurrents { get; set; } = new List<string>();

        // 解析状态标识: Success (均成功), Partial (部分成功), Failed (均失败)
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

        // 完全解析成功的行数 (极数、电流、脱扣方式均有效识别)
        public int SuccessCount { get; set; } = 0;

        // 部分或完全未识别的行数
        public int FailedCount { get; set; } = 0;

        // 执行总耗时 (毫秒)
        public long ElapsedMilliseconds { get; set; } = 0;

        // 结果摘要消息
        public string Message { get; set; } = string.Empty;
    }
}
