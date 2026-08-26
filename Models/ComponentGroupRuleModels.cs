using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 二次元件组数量计算策略枚举
    /// </summary>
    public enum QuantityPolicy
    {
        /// <summary>
        /// 固定套数 (默认为 1 套)
        /// </summary>
        Fixed = 0,

        /// <summary>
        /// 跟随主控元件数量 (# 修饰符驱动)
        /// </summary>
        FollowMainElement = 1,

        /// <summary>
        /// 自动按成组比例组数计算 (/ 比例联动驱动)
        /// </summary>
        AutoByRatio = 2
    }

    /// <summary>
    /// 条件逻辑运算符
    /// </summary>
    public enum LogicalOperator
    {
        /// <summary>
        /// 逻辑与 (所有条件同时满足)
        /// </summary>
        And = 0,

        /// <summary>
        /// 逻辑或 (满足任意一个条件即可)
        /// </summary>
        Or = 1
    }

    /// <summary>
    /// 元件匹配模式枚举
    /// </summary>
    public enum ElementMatchMode
    {
        /// <summary>
        /// 必须包含该元件 (普通包含)
        /// </summary>
        MustInclude = 0,

        /// <summary>
        /// 必须排除该元件 (不允许存在, 数量为 0)
        /// </summary>
        MustExclude = 1,

        /// <summary>
        /// 作为主控元件 (数量驱动 #)
        /// </summary>
        MainDriver = 2,

        /// <summary>
        /// 成组比例成员 (/ 分段配比)
        /// </summary>
        RatioMember = 3
    }

    /// <summary>
    /// 属性过滤单项配置
    /// </summary>
    public class PropertyFilterItem
    {
        /// <summary>
        /// 属性类型: Current(电流), Model(型号), Poles(极数), Appendix(附件)
        /// </summary>
        public string PropertyType { get; set; } = "Current";

        /// <summary>
        /// 比较运算符: >, >=, <, <=, ==, !=, contains, not_contains, empty, not_empty
        /// </summary>
        public string Operator { get; set; } = ">=";

        /// <summary>
        /// 目标比较值 (如: "1000", "WATSG", "3P", "1")
        /// </summary>
        public string Value { get; set; } = "";
    }

    /// <summary>
    /// 成组比例配比项 (如: 3 个 电流表)
    /// </summary>
    public class RatioItem
    {
        /// <summary>
        /// 比例基准数量 (如 3)
        /// </summary>
        public int Count { get; set; } = 1;

        /// <summary>
        /// 元件名称或匹配词 (如 "电流表")
        /// </summary>
        public string ElementName { get; set; } = "";

        /// <summary>
        /// 该比例项所附加的属性过滤列表
        /// </summary>
        public List<PropertyFilterItem> PropertyFilters { get; set; } = new List<PropertyFilterItem>();
    }

    /// <summary>
    /// 规则管道中的原子条件节点
    /// </summary>
    public class RuleConditionNode
    {
        /// <summary>
        /// 唯一标识符
        /// </summary>
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        /// <summary>
        /// 元件名称或通配符 (*)
        /// </summary>
        public string ElementName { get; set; } = "";

        /// <summary>
        /// 匹配模式: 包含 / 排除 / 主驱动 / 比例
        /// </summary>
        public ElementMatchMode Mode { get; set; } = ElementMatchMode.MustInclude;

        /// <summary>
        /// 成组比例集合 (当包含比例语法如 3电流表/1电压表 时使用)
        /// </summary>
        public List<RatioItem> RatioItems { get; set; } = new List<RatioItem>();

        /// <summary>
        /// 属性过滤列表 (电流、型号、极数、附件)
        /// </summary>
        public List<PropertyFilterItem> PropertyFilters { get; set; } = new List<PropertyFilterItem>();
    }

    /// <summary>
    /// 规则条件组 (支持 AND / OR 嵌套逻辑)
    /// </summary>
    public class RuleConditionGroup
    {
        /// <summary>
        /// 本组内的逻辑运算符 (默认 AND)
        /// </summary>
        public LogicalOperator Op { get; set; } = LogicalOperator.And;

        /// <summary>
        /// 条件节点列表
        /// </summary>
        public List<RuleConditionNode> Nodes { get; set; } = new List<RuleConditionNode>();

        /// <summary>
        /// 嵌套子条件组列表
        /// </summary>
        public List<RuleConditionGroup> SubGroups { get; set; } = new List<RuleConditionGroup>();
    }

    /// <summary>
    /// 前置守卫动作模式枚举 (满足条件时跳过 / 满足条件时才执行)
    /// </summary>
    public enum GuardActionMode
    {
        /// <summary>
        /// 满足条件时跳过本规则 (默认模式)
        /// </summary>
        SkipIfMatched = 0,

        /// <summary>
        /// 满足条件时才执行本规则
        /// </summary>
        ExecuteIfMatched = 1
    }

    /// <summary>
    /// 前置生效/跳过守卫条件模型 (基于箱柜原始数量比较两个元件数量)
    /// </summary>
    public class RuleGuardCondition
    {
        /// <summary>
        /// 唯一标识
        /// </summary>
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        /// <summary>
        /// 源元件名称 (如: "接触器")
        /// </summary>
        public string SourceElement { get; set; } = "";

        /// <summary>
        /// 数量比较运算符: ==, !=, >, >=, <, <=
        /// </summary>
        public string Operator { get; set; } = "==";

        /// <summary>
        /// 目标元件名称 (如: "热继电器")
        /// </summary>
        public string TargetElement { get; set; } = "";

        /// <summary>
        /// 守卫动作模式 (满足时跳过 / 满足时才执行)
        /// </summary>
        public GuardActionMode ActionMode { get; set; } = GuardActionMode.SkipIfMatched;

        /// <summary>
        /// 是否启用本守卫项
        /// </summary>
        public bool Enabled { get; set; } = true;
    }

    /// <summary>
    /// 单条二次元件组规则管道实体
    /// </summary>
    public class ComponentGroupRule
    {
        /// <summary>
        /// 规则唯一 ID
        /// </summary>
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        /// <summary>
        /// 规则友好名称 (如: "双电源报警回路")
        /// </summary>
        public string Name { get; set; } = "";

        /// <summary>
        /// 规则说明/备注
        /// </summary>
        public string Description { get; set; } = "";

        /// <summary>
        /// 是否启用该规则
        /// </summary>
        public bool Enabled { get; set; } = true;

        /// <summary>
        /// 匹配优先级序号 (数值越小越先执行)
        /// </summary>
        public int Priority { get; set; } = 1;

        /// <summary>
        /// 目标二次元件组名称 (写入 C 列，如: "*ATS+MXOF")
        /// </summary>
        public string TargetGroup { get; set; } = "";

        /// <summary>
        /// 数量生成策略
        /// </summary>
        public QuantityPolicy Policy { get; set; } = QuantityPolicy.FollowMainElement;

        /// <summary>
        /// 固定套数 (当 Policy 为 Fixed 时生效)
        /// </summary>
        public int FixedQuantity { get; set; } = 1;

        /// <summary>
        /// 主控元件名称 (当 Policy 为 FollowMainElement 时可指定，未指定自动取带 # 的元件)
        /// </summary>
        public string MainElementKey { get; set; } = "";

        /// <summary>
        /// 前置生效/跳过守卫条件列表 (基于箱柜原始元件总数对比)
        /// </summary>
        public List<RuleGuardCondition> Guards { get; set; } = new List<RuleGuardCondition>();

        /// <summary>
        /// 可视化条件树根节点
        /// </summary>
        public RuleConditionGroup ConditionTree { get; set; } = new RuleConditionGroup();

        /// <summary>
        /// 对应导出的经典文本表达式 (如: "双电源# 电流:>1000 型号:WATSG 附件:1")
        /// </summary>
        public string RawExpression { get; set; } = "";
    }

    /// <summary>
    /// Excel 列映射配置模型 (消除硬编码)
    /// </summary>
    public class ComponentGroupColumnMapping
    {
        /// <summary>
        /// 元件名称列索引 (默认 B 列 = 2)
        /// </summary>
        public int NameCol { get; set; } = 2;

        /// <summary>
        /// 元件型号规格列索引 (默认 C 列 = 3)
        /// </summary>
        public int NormsCol { get; set; } = 3;

        /// <summary>
        /// 元件数量/套数列索引 (默认 F 列 = 6)
        /// </summary>
        public int QuantityCol { get; set; } = 6;

        /// <summary>
        /// 单位列索引 (默认 E 列 = 5)
        /// </summary>
        public int UnitCol { get; set; } = 5;

        /// <summary>
        /// 电流参数列索引 (默认 V 列 = 22)
        /// </summary>
        public int CurrentCol { get; set; } = 22;

        /// <summary>
        /// 极数参数列索引 (默认 W 列 = 23)
        /// </summary>
        public int PolesCol { get; set; } = 23;

        /// <summary>
        /// 附件参数列索引 (默认 X 列 = 24)
        /// </summary>
        public int AppendixCol { get; set; } = 24;

        /// <summary>
        /// 类别标识列索引 (默认 B 列 = 2，写入 "元件组")
        /// </summary>
        public int CategoryCol { get; set; } = 2;
    }

    /// <summary>
    /// 二次元件组全局配置实体 (持久化至 ComponentGroupRules.json)
    /// </summary>
    public class ComponentGroupConfig
    {
        /// <summary>
        /// 列映射配置
        /// </summary>
        public ComponentGroupColumnMapping ColumnMapping { get; set; } = new ComponentGroupColumnMapping();

        /// <summary>
        /// 规则管道列表
        /// </summary>
        public List<ComponentGroupRule> Rules { get; set; } = new List<ComponentGroupRule>();

        /// <summary>
        /// 默认单位文本 (如 "套")
        /// </summary>
        public string DefaultUnitText { get; set; } = "套";

        /// <summary>
        /// 默认二次元件类别标识 (如 "元件组")
        /// </summary>
        public string DefaultCategoryText { get; set; } = "元件组";

        /// <summary>
        /// 是否在生成前检查去重 (避免同名元件组重复插入)
        /// </summary>
        public bool EnableDeduplication { get; set; } = true;

        /// <summary>
        /// 生成出厂默认规则库
        /// </summary>
        public static ComponentGroupConfig CreateDefault()
        {
            var config = new ComponentGroupConfig();
            int priority = 1;

            // 1. 凝露控制器
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "凝露/防潮控制回路",
                TargetGroup = "*凝露",
                Policy = QuantityPolicy.Fixed,
                FixedQuantity = 1,
                Priority = priority++,
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.Or,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode { ElementName = "凝露控制器", Mode = ElementMatchMode.MustInclude },
                        new RuleConditionNode { ElementName = "防潮控制器", Mode = ElementMatchMode.MustInclude }
                    }
                }
            });

            // 2. 双V3A电压电流表测量组
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "双V3A测量回路",
                TargetGroup = "*双V3A",
                Policy = QuantityPolicy.AutoByRatio,
                Priority = priority++,
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.And,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode
                        {
                            Mode = ElementMatchMode.RatioMember,
                            RatioItems = new List<RatioItem>
                            {
                                new RatioItem { ElementName = "电流表", Count = 3 },
                                new RatioItem { ElementName = "电压表", Count = 2 }
                            }
                        }
                    }
                }
            });

            // 3. 1V3A1Z 电流电压转换开关测量组
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "1V3A1Z转换开关测量",
                TargetGroup = "*1V3A1Z",
                Policy = QuantityPolicy.AutoByRatio,
                Priority = priority++,
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.And,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode
                        {
                            Mode = ElementMatchMode.RatioMember,
                            RatioItems = new List<RatioItem>
                            {
                                new RatioItem { ElementName = "电流表", Count = 3 },
                                new RatioItem { ElementName = "电压表", Count = 1 },
                                new RatioItem { ElementName = "转换开关", Count = 1 }
                            }
                        }
                    }
                }
            });

            // 4. 1V3A (无转换开关)
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "1V3A无转换测量",
                TargetGroup = "*1V3A",
                Policy = QuantityPolicy.AutoByRatio,
                Priority = priority++,
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.And,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode
                        {
                            Mode = ElementMatchMode.RatioMember,
                            RatioItems = new List<RatioItem>
                            {
                                new RatioItem { ElementName = "电流表", Count = 3 },
                                new RatioItem { ElementName = "电压表", Count = 1 }
                            }
                        },
                        new RuleConditionNode { ElementName = "转换开关", Mode = ElementMatchMode.MustExclude }
                    }
                }
            });

            // 5. 双电源大电流报警回路
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "双电源大电流报警",
                TargetGroup = "*ATS+MXOF",
                Policy = QuantityPolicy.FollowMainElement,
                Priority = priority++,
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.And,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode
                        {
                            ElementName = "双电源",
                            Mode = ElementMatchMode.MainDriver,
                            PropertyFilters = new List<PropertyFilterItem>
                            {
                                new PropertyFilterItem { PropertyType = "Appendix", Operator = "==", Value = "1" }
                            }
                        }
                    }
                }
            });

            // 6. 变频器控保组
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "变频器控保回路",
                TargetGroup = "*变频器控保",
                Policy = QuantityPolicy.AutoByRatio,
                Priority = priority++,
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.And,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode
                        {
                            Mode = ElementMatchMode.RatioMember,
                            RatioItems = new List<RatioItem>
                            {
                                new RatioItem { ElementName = "变频器", Count = 1 },
                                new RatioItem { ElementName = "控制保护开关", Count = 1 }
                            }
                        }
                    }
                }
            });

            // 7. KM变频器回路 (配置前置守卫: 当接触器数量与热继电器数量相等时跳过)
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "KM变频器联动",
                TargetGroup = "*KM变频器",
                Policy = QuantityPolicy.Fixed,
                Priority = priority++,
                Guards = new List<RuleGuardCondition>
                {
                    new RuleGuardCondition
                    {
                        SourceElement = "接触器",
                        Operator = "==",
                        TargetElement = "热继电器",
                        ActionMode = GuardActionMode.SkipIfMatched,
                        Enabled = true
                    }
                },
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.And,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode { ElementName = "变频器", Mode = ElementMatchMode.MustInclude },
                        new RuleConditionNode { ElementName = "接触器", Mode = ElementMatchMode.MustInclude }
                    }
                }
            });

            // 8. 单变频器回路 (排除控保与接触器)
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "单变频器独立回路",
                TargetGroup = "*单变频器",
                Policy = QuantityPolicy.Fixed,
                Priority = priority++,
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.And,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode { ElementName = "变频器", Mode = ElementMatchMode.MustInclude },
                        new RuleConditionNode { ElementName = "控制保护开关", Mode = ElementMatchMode.MustExclude },
                        new RuleConditionNode { ElementName = "接触器", Mode = ElementMatchMode.MustExclude }
                    }
                }
            });

            // 9. 软启动回路
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "软启动控制回路",
                TargetGroup = "*软启动",
                Policy = QuantityPolicy.FollowMainElement,
                Priority = priority++,
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.And,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode { ElementName = "软启动", Mode = ElementMatchMode.MainDriver }
                    }
                }
            });

            // 10. FA风机消防电源联动
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "FA风机消防电源监控",
                TargetGroup = "*FA风机",
                Policy = QuantityPolicy.FollowMainElement,
                Priority = priority++,
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.And,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode { ElementName = "消防电源监控", Mode = ElementMatchMode.MustInclude },
                        new RuleConditionNode { ElementName = "时控开关", Mode = ElementMatchMode.MustExclude }
                    },
                    SubGroups = new List<RuleConditionGroup>
                    {
                        new RuleConditionGroup
                        {
                            Op = LogicalOperator.Or,
                            Nodes = new List<RuleConditionNode>
                            {
                                new RuleConditionNode { ElementName = "热继电器", Mode = ElementMatchMode.MainDriver },
                                new RuleConditionNode { ElementName = "控制保护开关", Mode = ElementMatchMode.MainDriver }
                            }
                        }
                    }
                }
            });

            // 11. BA风机回路
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "BA风机楼控回路",
                TargetGroup = "*BA风机",
                Policy = QuantityPolicy.FollowMainElement,
                Priority = priority++,
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.And,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode { ElementName = "消防电源监控", Mode = ElementMatchMode.MustExclude },
                        new RuleConditionNode { ElementName = "变频器", Mode = ElementMatchMode.MustExclude },
                        new RuleConditionNode { ElementName = "时控开关", Mode = ElementMatchMode.MustExclude }
                    },
                    SubGroups = new List<RuleConditionGroup>
                    {
                        new RuleConditionGroup
                        {
                            Op = LogicalOperator.Or,
                            Nodes = new List<RuleConditionNode>
                            {
                                new RuleConditionNode { ElementName = "热继电器", Mode = ElementMatchMode.MainDriver },
                                new RuleConditionNode { ElementName = "控制保护开关", Mode = ElementMatchMode.MainDriver }
                            }
                        }
                    }
                }
            });

            // 12. 电能表回路
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "电能表计量回路",
                TargetGroup = "*电能表",
                Policy = QuantityPolicy.FollowMainElement,
                Priority = priority++,
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.And,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode { ElementName = "电能表", Mode = ElementMatchMode.MainDriver }
                    }
                }
            });

            // 13. 直接式表通讯
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "直接式表通讯",
                TargetGroup = "*表通讯",
                Policy = QuantityPolicy.FollowMainElement,
                Priority = priority++,
                ConditionTree = new RuleConditionGroup
                {
                    Op = LogicalOperator.And,
                    Nodes = new List<RuleConditionNode>
                    {
                        new RuleConditionNode
                        {
                            ElementName = "直接式表",
                            Mode = ElementMatchMode.MainDriver,
                            PropertyFilters = new List<PropertyFilterItem>
                            {
                                new PropertyFilterItem { PropertyType = "Model", Operator = "==", Value = "0" }
                            }
                        },
                        new RuleConditionNode { ElementName = "多功能表", Mode = ElementMatchMode.MustExclude }
                    }
                }
            });

            // 自动为所有规则生成只读摘要文本
            foreach (var rule in config.Rules)
            {
                rule.RawExpression = PipelineCompiler.BuildRuleSummary(rule);
            }

            return config;
        }
    }

    /// <summary>
    /// 一次元件数据传输对象 (用于内存评估与沙盒测试)
    /// </summary>
    public class EleComponentDto
    {
        /// <summary>
        /// Excel 中的行号 (便于追踪与定位)
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 元件名称 (B 列)
        /// </summary>
        public string EleName { get; set; } = "";

        /// <summary>
        /// 元件型号规格 (C 列)
        /// </summary>
        public string EleNorms { get; set; } = "";

        /// <summary>
        /// 元件数量 (F 列)
        /// </summary>
        public int EleNums { get; set; } = 1;

        /// <summary>
        /// 电流参数 (V 列)
        /// </summary>
        public string EleCurrent { get; set; } = "";

        /// <summary>
        /// 极数参数 (W 列)
        /// </summary>
        public string ElePoles { get; set; } = "";

        /// <summary>
        /// 附件参数 (X 列)
        /// </summary>
        public string EleAppendix { get; set; } = "";
    }

    /// <summary>
    /// 单个规则匹配命中结果
    /// </summary>
    public class RuleMatchResult
    {
        /// <summary>
        /// 规则 ID
        /// </summary>
        public string RuleId { get; set; } = "";

        /// <summary>
        /// 规则名称
        /// </summary>
        public string RuleName { get; set; } = "";

        /// <summary>
        /// 目标二次元件组名称 (C 列)
        /// </summary>
        public string TargetGroup { get; set; } = "";

        /// <summary>
        /// 计算得出的套数 (F 列)
        /// </summary>
        public int Quantity { get; set; } = 1;

        /// <summary>
        /// 匹配详情说明
        /// </summary>
        public string DetailInfo { get; set; } = "";

        /// <summary>
        /// 是否匹配成功
        /// </summary>
        public bool IsMatched { get; set; } = true;
    }

    /// <summary>
    /// 沙盒测试结果传输对象
    /// </summary>
    public class PipelineTestResultDto
    {
        /// <summary>
        /// 测试的元件总行数
        /// </summary>
        public int TotalComponents { get; set; }

        /// <summary>
        /// 命中的规则结果列表
        /// </summary>
        public List<RuleMatchResult> MatchedRules { get; set; } = new List<RuleMatchResult>();

        /// <summary>
        /// 诊断与执行过程日志
        /// </summary>
        public List<string> Logs { get; set; } = new List<string>();

        /// <summary>
        /// 规则流按序扣减后的最终剩余元件库存快照 (用于沙盒直观验证扣减)
        /// </summary>
        public List<EleComponentDto> RemainingComponents { get; set; } = new List<EleComponentDto>();
    }

    /// <summary>
    /// Excel 批量生成执行结果传输对象
    /// </summary>
    public class BatchGroupResultDto
    {
        /// <summary>
        /// 执行是否完全成功
        /// </summary>
        public bool Success { get; set; } = true;

        /// <summary>
        /// 处理的箱柜总数
        /// </summary>
        public int ProcessedCabinets { get; set; }

        /// <summary>
        /// 生成插入的二次元件组总行数
        /// </summary>
        public int InsertedGroupsCount { get; set; }

        /// <summary>
        /// 跳过的重复项数量
        /// </summary>
        public int SkippedDuplicateCount { get; set; }

        /// <summary>
        /// 耗时 (毫秒)
        /// </summary>
        public long ElapsedMilliseconds { get; set; }

        /// <summary>
        /// 详细提示或错误信息
        /// </summary>
        public string Message { get; set; } = "";

        /// <summary>
        /// 生成日志清单
        /// </summary>
        public List<string> Details { get; set; } = new List<string>();
    }

    #region 双向编译器与评估引擎

    /// <summary>
    /// 规则管道双向编译器: 实现可视化管道 ⇄ 经典文本表达式无损互转与解析
    /// </summary>
    /// <summary>
    /// 规则管道轻量摘要生成器: 纯结构化架构下仅用于生成人类可读的单行规则摘要 (供日志与界面概览展示)
    /// </summary>
    public static class PipelineCompiler
    {
        /// <summary>
        /// 根据规则实体生成人类可读的简要概括文本
        /// </summary>
        /// <param name="rule">二次元件组规则实体</param>
        /// <returns>简短可读的单行摘要字符串</returns>
        public static string BuildRuleSummary(ComponentGroupRule rule)
        {
            // 校验规则实体非空
            if (rule == null) return string.Empty;

            // 格式化条件树与前置守卫为单行概括文本
            return BuildExpressionFromTree(rule.ConditionTree, rule.Guards);
        }

        /// <summary>
        /// 将结构化条件树与前置守卫转换为简短文字标签
        /// </summary>
        /// <param name="group">条件树根节点</param>
        /// <param name="guards">前置守卫列表 (可选)</param>
        /// <returns>简要文字标签</returns>
        public static string BuildExpressionFromTree(RuleConditionGroup group, List<RuleGuardCondition>? guards = null)
        {
            // 校验条件树根节点有效性
            if (group == null) return string.Empty;

            var parts = new List<string>();

            // 1. 提取顶层普通条件节点名称与模式
            if (group.Nodes != null)
            {
                foreach (var node in group.Nodes)
                {
                    string nodeStr = FormatNodeSummary(node);
                    if (!string.IsNullOrEmpty(nodeStr)) parts.Add(nodeStr);
                }
            }

            // 2. 提取【或】关系子分支摘要
            if (group.SubGroups != null)
            {
                foreach (var sub in group.SubGroups)
                {
                    string subStr = BuildExpressionFromTree(sub);
                    if (!string.IsNullOrEmpty(subStr)) parts.Add($"({subStr})");
                }
            }

            // 拼接主体条件部分
            string opSymbol = group.Op == LogicalOperator.Or ? " | " : " & ";
            string bodyExpr = string.Join(opSymbol, parts);

            // 3. 拼接前置守卫概览前缀 (若配置了守卫)
            string guardPrefix = FormatGuardsSummary(guards);
            if (!string.IsNullOrEmpty(guardPrefix))
            {
                return string.IsNullOrEmpty(bodyExpr) ? guardPrefix : $"{guardPrefix} {bodyExpr}";
            }

            return bodyExpr;
        }

        /// <summary>
        /// 格式化前置守卫为简短前缀概括
        /// </summary>
        public static string FormatGuardsSummary(List<RuleGuardCondition>? guards)
        {
            if (guards == null || guards.Count == 0) return string.Empty;

            // 过滤有效启用的守卫条件项
            var valid = guards.Where(g => g.Enabled && !string.IsNullOrWhiteSpace(g.SourceElement) && !string.IsNullOrWhiteSpace(g.TargetElement)).ToList();
            if (valid.Count == 0) return string.Empty;

            // 分组提取跳过与生效守卫
            var skipGuards = valid.Where(g => g.ActionMode == GuardActionMode.SkipIfMatched).ToList();
            var onlyGuards = valid.Where(g => g.ActionMode == GuardActionMode.ExecuteIfMatched).ToList();

            var prefixList = new List<string>();
            if (skipGuards.Count > 0)
            {
                string inner = string.Join("; ", skipGuards.Select(g => $"{g.SourceElement} {g.Operator} {g.TargetElement}"));
                prefixList.Add($"[跳过: {inner}]");
            }
            if (onlyGuards.Count > 0)
            {
                string inner = string.Join("; ", onlyGuards.Select(g => $"{g.SourceElement} {g.Operator} {g.TargetElement}"));
                prefixList.Add($"[仅当: {inner}]");
            }

            return string.Join(" ", prefixList);
        }

        /// <summary>
        /// 格式化单个条件节点的简短概览文本 (供列表与摘要概览)
        /// </summary>
        private static string FormatNodeSummary(RuleConditionNode node)
        {
            // 校验节点非空
            if (node == null) return string.Empty;

            string baseText = "";
            // 成组比例模式: 拼接多个成员比例 (如 3电流表/1电压表)
            if (node.Mode == ElementMatchMode.RatioMember && node.RatioItems?.Count > 0)
            {
                // 拼接比例成员数量与名称
                baseText = string.Join("/", node.RatioItems.Select(r => $"{r.Count}{r.ElementName}"));
            }
            // 排除模式: 增加负号前缀 (如 -时控开关)
            else if (node.Mode == ElementMatchMode.MustExclude)
            {
                // 拼接排除前缀
                baseText = $"-{node.ElementName}";
            }
            // 主控数量驱动模式: 增加 # 后缀 (如 双电源#)
            else if (node.Mode == ElementMatchMode.MainDriver)
            {
                // 拼接主控驱动符号
                baseText = $"{node.ElementName}#";
            }
            else
            {
                // 普通包含模式
                baseText = node.ElementName ?? string.Empty;
            }

            // 拼接属性过滤器简要文本 (如 电流:>=1000 附件:1)
            if (node.Mode != ElementMatchMode.RatioMember && node.PropertyFilters != null && node.PropertyFilters.Count > 0)
            {
                // 收集非空属性过滤文本
                var filterStrings = new List<string>();
                // 遍历各个属性过滤器项
                foreach (var pf in node.PropertyFilters)
                {
                    // 跳过未设置比较值的过滤项
                    if (string.IsNullOrWhiteSpace(pf.Value)) continue;
                    // 等于运算符时省略 == 符号以保持简洁
                    string op = pf.Operator == "==" || pf.Operator == "=" ? "" : pf.Operator;
                    // 映射属性类型为中文标签
                    string propLabel = pf.PropertyType switch
                    {
                        "Current" => "电流",
                        "Model" => "型号",
                        "Poles" => "极数",
                        "Appendix" => "附件",
                        _ => pf.PropertyType
                    };
                    // 添加单项属性过滤字符串
                    filterStrings.Add($"{propLabel}:{op}{pf.Value}");
                }
                // 若存在有效属性过滤器则拼接至节点文本后
                if (filterStrings.Count > 0)
                {
                    // 拼接空格与属性清单
                    baseText += " " + string.Join(" ", filterStrings);
                }
            }

            // 返回格式化节点文本
            return baseText;
        }
    }

    /// <summary>
    /// 扣减执行计划项 (记录待扣减的过滤条件、基准单套扣减数量与最大可用套数)
    /// </summary>
    public class DeductionPlanItem
    {
        /// <summary>
        /// 匹配项描述 (供日志输出，如 "[双电源# (附件:1)]")
        /// </summary>
        public string Description { get; set; } = "";

        /// <summary>
        /// 每套二次元件组消耗该元件的基础数量 (主控 # 为 1，比例 / 为 baseRatio)
        /// </summary>
        public int BaseCountPerSet { get; set; } = 1;

        /// <summary>
        /// 该条件项单独支持的最大成套数
        /// </summary>
        public int MaxAvailableSets { get; set; } = int.MaxValue;

        /// <summary>
        /// 判定元件是否符合该扣减项的谓词委托
        /// </summary>
        public Func<EleComponentDto, bool> MatchPredicate { get; set; } = _ => false;
    }

    /// <summary>
    /// 规则管道高效内存执行评估器 (支持箱柜元件动态资源池扣减机制与前置守卫过滤)
    /// </summary>
    public static class PipelineEvaluator
    {
        /// <summary>
        /// 针对整个箱柜元件集合，按规则优先级依次进行带【前置守卫检查】与【动态资源池扣减】的完整评估
        /// </summary>
        /// <param name="rules">启用的规则管道列表 (已按优先级排序)</param>
        /// <param name="components">箱柜一次元件集合</param>
        /// <param name="executionLogs">执行与扣减过程诊断日志列表</param>
        /// <param name="remainingComponents">执行完毕后剩余元件库存列表</param>
        /// <returns>命中的规则匹配结果列表</returns>
        public static List<RuleMatchResult> EvaluateRulesWithResourcePool(
            List<ComponentGroupRule> rules,
            List<EleComponentDto> components,
            out List<string> executionLogs,
            out List<EleComponentDto> remainingComponents)
        {
            // 初始化命中结果集合
            var matchedResults = new List<RuleMatchResult>();
            // 初始化执行过程诊断日志集合
            executionLogs = new List<string>();

            // 基础前置参数校验
            if (rules == null || rules.Count == 0 || components == null || components.Count == 0)
            {
                // 记录空列表提示
                executionLogs.Add("[资源池] 元件列表或规则列表为空，无需评估。");
                // 兜底返回剩余库存
                remainingComponents = components ?? new List<EleComponentDto>();
                // 返回空匹配结果
                return matchedResults;
            }

            // 1. 克隆一份深拷贝作为当前箱柜的动态工作资源池 (按序消耗扣减)
            var workingPool = components.Select(c => new EleComponentDto
            {
                // 复制所在行号
                RowIndex = c.RowIndex,
                // 复制元件名称
                EleName = c.EleName,
                // 复制型号规格
                EleNorms = c.EleNorms,
                // 复制数量
                EleNums = c.EleNums,
                // 复制电流参数
                EleCurrent = c.EleCurrent,
                // 复制极数参数
                ElePoles = c.ElePoles,
                // 复制附件参数
                EleAppendix = c.EleAppendix
            }).ToList();

            // 2. 同时克隆一份原始库存镜像 (用于前置守卫比对两元件的原始数量，不受前面规则扣减影响)
            var originalPool = components.Select(c => new EleComponentDto
            {
                // 复制所在行号
                RowIndex = c.RowIndex,
                // 复制元件名称
                EleName = c.EleName,
                // 复制型号规格
                EleNorms = c.EleNorms,
                // 复制数量
                EleNums = c.EleNums,
                // 复制电流参数
                EleCurrent = c.EleCurrent,
                // 复制极数参数
                ElePoles = c.ElePoles,
                // 复制附件参数
                EleAppendix = c.EleAppendix
            }).ToList();

            // 打印初始元件库存明细
            var initialCompSummary = workingPool
                .Where(c => c.EleNums > 0)
                .Select(c => $"[{c.EleName}]: {c.EleNums}件");
            // 记录初始库存池快照至日志
            executionLogs.Add($"📦 【初始元件库存池】: {string.Join(", ", initialCompSummary)}");

            // 规则评估步骤序号计数器
            int step = 1;
            // 3. 依次按规则优先级从上到下遍历评估已启用的规则
            foreach (var rule in rules.Where(r => r.Enabled).OrderBy(r => r.Priority))
            {
                // 1. 前置生效与跳过守卫判断 (基于原始元件池 originalPool 对比两元件数量)
                var guards = rule.Guards;
                // 判断是否存在守卫配置
                if (guards != null && guards.Count > 0)
                {
                    // 标记是否被守卫拦截跳过
                    bool isGuardedOut = false;
                    // 记录守卫拦截原因
                    string guardLogReason = "";

                    // 遍历所有已启用的守卫项
                    foreach (var guard in guards.Where(g => g.Enabled && !string.IsNullOrWhiteSpace(g.SourceElement) && !string.IsNullOrWhiteSpace(g.TargetElement)))
                    {
                        // 统计源元件在原始箱柜中的总数
                        int srcCount = originalPool.Where(c => c.EleNums > 0 && MatchSingleComponentProperties(c, guard.SourceElement)).Sum(c => c.EleNums);
                        // 统计目标元件在原始箱柜中的总数
                        int tgtCount = originalPool.Where(c => c.EleNums > 0 && MatchSingleComponentProperties(c, guard.TargetElement)).Sum(c => c.EleNums);
                        // 进行数值关系比较
                        bool isConditionTrue = EvaluateNumberComparison(srcCount, guard.Operator, tgtCount);

                        // 满足时跳过模式
                        if (guard.ActionMode == GuardActionMode.SkipIfMatched && isConditionTrue)
                        {
                            // 标记拦截
                            isGuardedOut = true;
                            // 构造拦截日志
                            guardLogReason = $"原始[{guard.SourceElement}]数量({srcCount}件) {guard.Operator} 原始[{guard.TargetElement}]数量({tgtCount}件) -> 触发跳过";
                            // 跳出守卫循环
                            break;
                        }
                        // 满足时才生效模式
                        else if (guard.ActionMode == GuardActionMode.ExecuteIfMatched && !isConditionTrue)
                        {
                            // 标记拦截
                            isGuardedOut = true;
                            // 构造前提不成立日志
                            guardLogReason = $"原始[{guard.SourceElement}]数量({srcCount}件) 不满足 {guard.Operator} 原始[{guard.TargetElement}]数量({tgtCount}件) -> 前提不成立跳过";
                            // 跳出守卫循环
                            break;
                        }
                    }

                    // 若判定跳过，记录日志并直接处理下一规则
                    if (isGuardedOut)
                    {
                        // 输出守卫跳过日志
                        executionLogs.Add($"[步骤 {step}] ⏭️ 规则 [{rule.Name}] 被前置守卫跳过");
                        // 输出守卫决策细节
                        executionLogs.Add($"    └── 守卫决策: {guardLogReason}");
                        // 步进序号
                        step++;
                        // 继续下一规则
                        continue;
                    }
                }

                // 2. 基于强类型条件树直接在当前工作资源池中进行精准求值与扣减计划推导
                bool isMatched = EvaluateConditionGroup(rule.ConditionTree, workingPool, out int driverQuantity, out List<DeductionPlanItem> planItems);

                // 判断规则是否匹配成功且数量有效
                if (isMatched && driverQuantity > 0)
                {
                    // 确定最终生成套数 (依据套数策略确定)
                    int finalQty = driverQuantity;
                    // 若配置了固定套数策略且数值合法
                    if (rule.Policy == QuantityPolicy.Fixed && rule.FixedQuantity > 0)
                    {
                        // 采用固定套数
                        finalQty = rule.FixedQuantity;
                    }
                    // 兜底保证至少 1 套
                    if (finalQty <= 0) finalQty = 1;

                    // 提交扣减: 正式从工作资源池中扣除消耗的元件数量 (严格按 finalQty * BaseCountPerSet 扣减)
                    var deductSummary = new List<string>();
                    // 遍历所有的待扣减项
                    foreach (var plan in planItems)
                    {
                        // 计算该项需要扣减的总件数 (生成套数 * 每套消耗件数)
                        int requiredDeduct = finalQty * plan.BaseCountPerSet;
                        // 追踪剩余待扣减件数
                        int remainingToDeduct = requiredDeduct;
                        // 追踪本次实际扣减的件数
                        int totalDeductedForThis = 0;

                        // 遍历工作资源池中的元件行
                        foreach (var comp in workingPool)
                        {
                            // 跳过已耗尽库存的元件行
                            if (comp.EleNums <= 0) continue;
                            // 校验元件是否满足扣减匹配谓词
                            if (plan.MatchPredicate(comp))
                            {
                                // 计算当前行可扣减数量
                                int deductThis = Math.Min(comp.EleNums, remainingToDeduct);
                                // 从工作池扣除库存
                                comp.EleNums -= deductThis;
                                // 减少剩余待扣减额度
                                remainingToDeduct -= deductThis;
                                // 累加实际扣减数
                                totalDeductedForThis += deductThis;
                                // 满足扣减额度后提前退出遍历
                                if (remainingToDeduct <= 0) break;
                            }
                        }

                        // 统计扣减后该类元件在当前工作池中的剩余总数
                        int poolRemaining = workingPool
                            .Where(c => plan.MatchPredicate(c))
                            .Sum(c => c.EleNums);

                        // 记录单项扣减汇总摘要
                        deductSummary.Add($"{plan.Description} 消耗 {totalDeductedForThis}件 (池剩余 {poolRemaining}件)");
                    }

                    // 构造匹配详情说明
                    string detailStr = deductSummary.Count > 0
                        ? $"匹配成功! 生成 {finalQty} 套 [{rule.TargetGroup}] (消耗: {string.Join(", ", deductSummary)})"
                        : $"匹配成功! 生成 {finalQty} 套 [{rule.TargetGroup}] (关联条件满足，无扣减)";

                    // 统计当前规则匹配结果
                    var matchRes = new RuleMatchResult
                    {
                        // 记录规则 ID
                        RuleId = rule.Id,
                        // 记录规则名称
                        RuleName = rule.Name,
                        // 记录目标二次元件组名称
                        TargetGroup = rule.TargetGroup,
                        // 记录生成的套数
                        Quantity = finalQty,
                        // 记录匹配状态为成功
                        IsMatched = true,
                        // 记录详细信息
                        DetailInfo = detailStr
                    };
                    // 加入匹配结果集
                    matchedResults.Add(matchRes);

                    // 输出命中日志
                    executionLogs.Add($"[步骤 {step}] ✅ 规则 [{rule.Name}] 命中! 生成 {finalQty} 套 [{rule.TargetGroup}]");
                    // 输出扣减明细日志
                    if (deductSummary.Count > 0)
                    {
                        // 打印扣减项明细
                        executionLogs.Add($"    └── 🔻 主控/比例元件扣减: {string.Join("; ", deductSummary)}");
                    }
                    else
                    {
                        // 打印无扣减项提示
                        executionLogs.Add($"    └── 条件满足 (无主控/比例扣减项)");
                    }
                }
                else
                {
                    // 记录未命中日志
                    executionLogs.Add($"[步骤 {step}] ⚪ 规则 [{rule.Name}] 未命中 (条件不满足或资源池不足)");
                }

                // 步进序号
                step++;
            }

            // 4. 输出最终剩余元件库存快照
            var finalRemaining = workingPool.Select(c => $"[{c.EleName}]: {c.EleNums}件");
            // 记录最终剩余库存日志
            executionLogs.Add($"🏁 【最终剩余元件库存】: {string.Join(", ", finalRemaining)}");
            // 赋值输出剩余库存
            remainingComponents = workingPool;

            // 返回全部命中的规则列表
            return matchedResults;
        }

        /// <summary>
        /// 基于结构化条件组直接在元件池中求值，推导是否命中、驱动套数与待扣减计划 (支持 AND/OR 嵌套与主控#精准扣减)
        /// </summary>
        /// <param name="group">条件组节点</param>
        /// <param name="pool">当前工作资源池</param>
        /// <param name="driverQuantity">推导出的主控驱动套数 (若无主控则为 1)</param>
        /// <param name="planItems">待扣减计划项列表</param>
        /// <returns>条件组是否完全满足</returns>
        private static bool EvaluateConditionGroup(
            RuleConditionGroup group,
            List<EleComponentDto> pool,
            out int driverQuantity,
            out List<DeductionPlanItem> planItems)
        {
            // 初始化驱动套数默认为 1
            driverQuantity = 1;
            // 初始化待扣减计划清单
            planItems = new List<DeductionPlanItem>();

            // 校验条件组非空与有效性
            if (group == null || ((group.Nodes == null || group.Nodes.Count == 0) && (group.SubGroups == null || group.SubGroups.Count == 0)))
            {
                // 空条件组视为不满足
                return false;
            }

            // 分支 1: 【逻辑与 (AND)】模式 —— 所有条件节点与子分支必须全部满足
            if (group.Op == LogicalOperator.And)
            {
                // 追踪主控或比例驱动套数的最小值
                int minDriverSets = int.MaxValue;
                // 标记是否存在主控或比例驱动节点
                bool hasDriver = false;

                // 1. 遍历当前组内的所有原子条件节点
                if (group.Nodes != null)
                {
                    // 遍历每个条件节点
                    foreach (var node in group.Nodes)
                    {
                        // 模式 A: 必须排除模式 (MustExclude)
                        if (node.Mode == ElementMatchMode.MustExclude)
                        {
                            // 统计池中满足该排除条件的元件总数
                            int count = pool.Where(c => c.EleNums > 0 && MatchSingleNode(c, node)).Sum(c => c.EleNums);
                            // 若存在排除元件，则 AND 逻辑直接失败
                            if (count > 0) return false;
                        }
                        // 模式 B: 必须包含模式 (MustInclude)
                        else if (node.Mode == ElementMatchMode.MustInclude)
                        {
                            // 统计池中满足包含条件的元件总数
                            int count = pool.Where(c => c.EleNums > 0 && MatchSingleNode(c, node)).Sum(c => c.EleNums);
                            // 若不存在该元件，则 AND 逻辑直接失败
                            if (count == 0) return false;
                        }
                        // 模式 C: 主控数量驱动模式 (MainDriver #)
                        else if (node.Mode == ElementMatchMode.MainDriver)
                        {
                            // 统计池中满足主控条件的可用元件总数
                            int count = pool.Where(c => c.EleNums > 0 && MatchSingleNode(c, node)).Sum(c => c.EleNums);
                            // 若主控元件不足 1 件，则 AND 逻辑直接失败
                            if (count == 0) return false;

                            // 标记存在驱动节点
                            hasDriver = true;
                            // 记录可用套数
                            minDriverSets = Math.Min(minDriverSets, count);
                            // 加入待扣减计划项 (每套消耗 1 件该主控元件)
                            planItems.Add(new DeductionPlanItem
                            {
                                // 描述文本
                                Description = FormatNodeDescription(node),
                                // 每套消耗 1 件
                                BaseCountPerSet = 1,
                                // 最大支持套数
                                MaxAvailableSets = count,
                                // 匹配谓词委托
                                MatchPredicate = c => MatchSingleNode(c, node)
                            });
                        }
                        // 模式 D: 成组比例联动模式 (RatioMember /)
                        else if (node.Mode == ElementMatchMode.RatioMember)
                        {
                            // 校验比例子项有效性
                            if (node.RatioItems == null || node.RatioItems.Count == 0) continue;

                            // 遍历每个比例成员项 (如 3个 电流表)
                            foreach (var rItem in node.RatioItems)
                            {
                                // 统计当前池中匹配该比例成员的可用总数
                                int count = pool.Where(c => c.EleNums > 0 && MatchRatioItem(c, rItem)).Sum(c => c.EleNums);
                                // 若连基础 1 套配比都不足，则 AND 逻辑失败
                                if (count < rItem.Count) return false;

                                // 标记存在驱动节点
                                hasDriver = true;
                                // 计算该成员能单独支持的套数
                                int setsForThis = count / rItem.Count;
                                // 取所有成员套数的最小公约数
                                minDriverSets = Math.Min(minDriverSets, setsForThis);
                                // 加入待扣减计划项 (每套消耗 baseRatio 件)
                                planItems.Add(new DeductionPlanItem
                                {
                                    // 描述文本
                                    Description = FormatRatioDescription(rItem),
                                    // 每套消耗 baseRatio 件
                                    BaseCountPerSet = rItem.Count,
                                    // 最大支持套数
                                    MaxAvailableSets = setsForThis,
                                    // 匹配谓词委托
                                    MatchPredicate = c => MatchRatioItem(c, rItem)
                                });
                            }
                        }
                    }
                }

                // 2. 递归遍历当前组内的所有子条件组
                if (group.SubGroups != null)
                {
                    // 遍历每个子条件组
                    foreach (var subGroup in group.SubGroups)
                    {
                        // 递归求值子组
                        bool subOk = EvaluateConditionGroup(subGroup, pool, out int subDriverQty, out List<DeductionPlanItem> subPlans);
                        // 若子组不满足，则 AND 逻辑失败
                        if (!subOk) return false;

                        // 若子组产出了扣减计划
                        if (subPlans.Count > 0)
                        {
                            // 标记存在驱动节点
                            hasDriver = true;
                            // 综合子组套数
                            minDriverSets = Math.Min(minDriverSets, subDriverQty);
                            // 合并子组扣减清单
                            planItems.AddRange(subPlans);
                        }
                    }
                }

                // 计算最终驱动套数 (若存在主控/比例则取最小值，否则为 1)
                driverQuantity = hasDriver && minDriverSets != int.MaxValue ? minDriverSets : 1;
                // 返回 AND 条件完全满足
                return true;
            }
            // 分支 2: 【逻辑或 (OR)】模式 —— 按顺序试探各个候选分支，首个满足者互斥命中 (Short-Circuit OR)
            else
            {
                // 1. 试探顶层原子条件节点中的候选者
                if (group.Nodes != null)
                {
                    // 遍历候选节点
                    foreach (var node in group.Nodes)
                    {
                        // 主控数量驱动模式
                        if (node.Mode == ElementMatchMode.MainDriver)
                        {
                            // 统计池中满足该主控条件的可用总数
                            int count = pool.Where(c => c.EleNums > 0 && MatchSingleNode(c, node)).Sum(c => c.EleNums);
                            // 若主控数量满足
                            if (count > 0)
                            {
                                // 设置驱动套数为该主控数量
                                driverQuantity = count;
                                // 仅将该命中的分支加入扣减清单 (互斥锁定)
                                planItems = new List<DeductionPlanItem>
                                {
                                    new DeductionPlanItem
                                    {
                                        Description = FormatNodeDescription(node),
                                        BaseCountPerSet = 1,
                                        MaxAvailableSets = count,
                                        MatchPredicate = c => MatchSingleNode(c, node)
                                    }
                                };
                                // 首个满足即代表 OR 成功，直接返回
                                return true;
                            }
                        }
                        // 必须包含模式
                        else if (node.Mode == ElementMatchMode.MustInclude)
                        {
                            // 统计池中满足包含条件的元件总数
                            int count = pool.Where(c => c.EleNums > 0 && MatchSingleNode(c, node)).Sum(c => c.EleNums);
                            // 若存在该元件
                            if (count > 0)
                            {
                                // 驱动套数设为 1
                                driverQuantity = 1;
                                // 清空扣减项
                                planItems = new List<DeductionPlanItem>();
                                // OR 命中成功返回
                                return true;
                            }
                        }
                    }
                }

                // 2. 试探子条件组候选分支
                if (group.SubGroups != null)
                {
                    // 遍历每个子候选分支
                    foreach (var subGroup in group.SubGroups)
                    {
                        // 递归求值子候选分支
                        bool subOk = EvaluateConditionGroup(subGroup, pool, out int subDriverQty, out List<DeductionPlanItem> subPlans);
                        // 若该子分支满足
                        if (subOk)
                        {
                            // 锁定该分支的驱动套数
                            driverQuantity = subDriverQty;
                            // 锁定该分支的扣减项清单 (不触发后续分支扣减)
                            planItems = subPlans;
                            // OR 命中成功返回
                            return true;
                        }
                    }
                }

                // 若所有候选分支均未满足，则 OR 逻辑失败
                driverQuantity = 0;
                // 清空扣减项
                planItems = new List<DeductionPlanItem>();
                // 返回失败
                return false;
            }
        }

        /// <summary>
        /// 校验单个元件是否匹配指定的规则条件节点 (包含名称通配符与所有属性过滤)
        /// </summary>
        private static bool MatchSingleNode(EleComponentDto ec, RuleConditionNode node)
        {
            // 校验基础对象有效性
            if (ec == null || node == null) return false;

            // 1. 元件名称匹配 (支持通配符 *，为空或 * 表示匹配任意元件)
            if (!string.IsNullOrWhiteSpace(node.ElementName) && node.ElementName.Trim() != "*")
            {
                // 若元件名称不包含指定关键词则判定不匹配
                if (string.IsNullOrEmpty(ec.EleName) || !ec.EleName.Contains(node.ElementName.Trim()))
                {
                    // 返回不匹配
                    return false;
                }
            }

            // 2. 属性过滤器列表逐项校验 (电流、型号、极数、附件)
            if (node.PropertyFilters != null && node.PropertyFilters.Count > 0)
            {
                // 遍历每个属性过滤器
                foreach (var filter in node.PropertyFilters)
                {
                    // 只要有任意一个属性过滤不满足，则判定不匹配
                    if (!MatchPropertyFilter(ec, filter))
                    {
                        // 返回不匹配
                        return false;
                    }
                }
            }

            // 所有条件均满足
            return true;
        }

        /// <summary>
        /// 校验单个元件是否匹配指定的比例成员项
        /// </summary>
        private static bool MatchRatioItem(EleComponentDto ec, RatioItem item)
        {
            // 校验基础对象有效性
            if (ec == null || item == null) return false;

            // 1. 元件名称匹配
            if (!string.IsNullOrWhiteSpace(item.ElementName) && item.ElementName.Trim() != "*")
            {
                // 若元件名称不包含指定关键词则判定不匹配
                if (string.IsNullOrEmpty(ec.EleName) || !ec.EleName.Contains(item.ElementName.Trim()))
                {
                    // 返回不匹配
                    return false;
                }
            }

            // 2. 属性过滤匹配
            if (item.PropertyFilters != null && item.PropertyFilters.Count > 0)
            {
                // 遍历每个属性过滤器
                foreach (var filter in item.PropertyFilters)
                {
                    // 校验属性过滤项
                    if (!MatchPropertyFilter(ec, filter))
                    {
                        // 返回不匹配
                        return false;
                    }
                }
            }

            // 满足比例项匹配
            return true;
        }

        /// <summary>
        /// 校验单个元件是否满足指定的属性过滤条件 (电流、型号、极数、附件)
        /// </summary>
        private static bool MatchPropertyFilter(EleComponentDto ec, PropertyFilterItem filter)
        {
            // 空过滤器视为通过
            if (ec == null || filter == null) return true;

            // 提取属性类型、运算符与目标值
            string propType = filter.PropertyType?.Trim() ?? "";
            // 比较运算符
            string op = filter.Operator?.Trim() ?? "==";
            // 目标比较值
            string targetVal = filter.Value?.Trim() ?? "";

            // 根据属性类型分别进行精确评估
            switch (propType)
            {
                // 电流过滤 (V 列)
                case "Current":
                    // 评估电流数值条件比较
                    return EvaluateCurrentConditionWithOp(ec.EleCurrent, op, targetVal);

                // 型号规格过滤 (C 列)
                case "Model":
                    // 若目标值为空，要求元件型号也为空
                    if (string.IsNullOrEmpty(targetVal))
                    {
                        // 判断型号是否为空
                        return string.IsNullOrEmpty(ec.EleNorms);
                    }
                    // 否则判断型号是否包含目标文本
                    return ec.EleNorms != null && ec.EleNorms.Contains(targetVal);

                // 极数过滤 (W 列)
                case "Poles":
                    // 评估极数格式比较
                    return EvaluatePoleConditionWithOp(ec.ElePoles, op, targetVal);

                // 附件过滤 (X 列)
                case "Appendix":
                    // 若目标值为空，要求必须无附件标识
                    if (string.IsNullOrEmpty(targetVal))
                    {
                        // 判断附件是否为空
                        return string.IsNullOrWhiteSpace(ec.EleAppendix);
                    }
                    // 否则要求附件列包含指定标识
                    return ec.EleAppendix != null && ec.EleAppendix.Contains(targetVal);

                // 其他类型默认通过
                default:
                    return true;
            }
        }

        /// <summary>
        /// 评估电流数值条件比较 (支持 >, >=, <, <=, ==, !=)
        /// </summary>
        private static bool EvaluateCurrentConditionWithOp(string rawCurrent, string op, string targetVal)
        {
            // 校验输入字符串有效性
            if (string.IsNullOrWhiteSpace(rawCurrent) || string.IsNullOrWhiteSpace(targetVal)) return false;

            // 清洗电流中的单位符号 A
            string cleanCur = rawCurrent.Replace("A", "").Trim();
            // 清洗目标值中的单位符号 A
            string cleanTgt = targetVal.Replace("A", "").Trim();

            // 解析电流数值
            if (!double.TryParse(cleanCur, NumberStyles.Any, CultureInfo.InvariantCulture, out double curVal) ||
                !double.TryParse(cleanTgt, NumberStyles.Any, CultureInfo.InvariantCulture, out double tgtVal))
            {
                // 解析失败返回 false
                return false;
            }

            // 根据运算符执行纯数值比较 (微秒级高性能)
            return op switch
            {
                ">" => curVal > tgtVal,
                ">=" or "=>" => curVal >= tgtVal,
                "<" => curVal < tgtVal,
                "<=" or "=<" => curVal <= tgtVal,
                "==" or "=" => Math.Abs(curVal - tgtVal) < 0.0001,
                "!=" or "<>" => Math.Abs(curVal - tgtVal) >= 0.0001,
                _ => curVal >= tgtVal
            };
        }

        /// <summary>
        /// 评估极数条件比较 (如 3P, 4P, 1P+N)
        /// </summary>
        private static bool EvaluatePoleConditionWithOp(string rawPoles, string op, string targetVal)
        {
            // 校验输入字符串非空
            if (string.IsNullOrWhiteSpace(rawPoles) || string.IsNullOrWhiteSpace(targetVal)) return false;

            // 标准化极数字符串 (统一大写并去除 P)
            string p1 = rawPoles.Trim().ToUpper().Replace("P", "");
            // 标准化目标值
            string p2 = targetVal.Trim().ToUpper().Replace("P", "");

            // 不等关系判断
            if (op == "!=" || op == "<>")
            {
                // 返回不等比较结果
                return p1 != p2;
            }
            // 默认等于比较
            return p1 == p2;
        }

        /// <summary>
        /// 格式化单个条件节点的描述标签 (供日志输出)
        /// </summary>
        private static string FormatNodeDescription(RuleConditionNode node)
        {
            // 节点非空校验
            if (node == null) return "";
            // 提取元件名称
            string name = string.IsNullOrWhiteSpace(node.ElementName) ? "*" : node.ElementName;
            // 提取主控标识
            string modeTag = node.Mode == ElementMatchMode.MainDriver ? "#" : "";

            // 格式化属性过滤器
            var filterTags = new List<string>();
            // 遍历属性过滤器列表
            if (node.PropertyFilters != null)
            {
                // 逐项提取
                foreach (var pf in node.PropertyFilters)
                {
                    // 跳过空值
                    if (string.IsNullOrWhiteSpace(pf.Value)) continue;
                    // 映射属性中文名
                    string pName = pf.PropertyType switch
                    {
                        "Current" => "电流",
                        "Model" => "型号",
                        "Poles" => "极数",
                        "Appendix" => "附件",
                        _ => pf.PropertyType
                    };
                    // 添加属性标签
                    filterTags.Add($"{pName}:{pf.Operator}{pf.Value}");
                }
            }

            // 若存在属性过滤标签则输出完整描述
            if (filterTags.Count > 0)
            {
                // 返回带属性过滤的标签
                return $"[{name}{modeTag} ({string.Join(", ", filterTags)})]";
            }
            // 返回纯名称标签
            return $"[{name}{modeTag}]";
        }

        /// <summary>
        /// 格式化比例项描述标签 (供日志输出)
        /// </summary>
        private static string FormatRatioDescription(RatioItem item)
        {
            // 校验比例项非空
            if (item == null) return "";
            // 返回比例数量与名称标签
            return $"[{item.Count}个 {item.ElementName}]";
        }

        /// <summary>
        /// 执行两整数数值间的逻辑关系比较
        /// </summary>
        /// <param name="a">左操作数 (源元件数量)</param>
        /// <param name="op">比较运算符: ==, !=, >, >=, <, <=</param>
        /// <param name="b">右操作数 (目标元件数量)</param>
        /// <returns>比较结果布尔值</returns>
        private static bool EvaluateNumberComparison(int a, string op, int b)
        {
            // 标准化比较运算符并计算布尔值
            return op switch
            {
                "==" or "=" => a == b,
                "!=" or "<>" => a != b,
                ">" => a > b,
                ">=" or "=>" => a >= b,
                "<" => a < b,
                "<=" or "=<" => a <= b,
                _ => a == b
            };
        }

        /// <summary>
        /// 对单条规则管道在给定元件集合上进行执行评估 (兼容单规则沙盒测试)
        /// </summary>
        public static RuleMatchResult EvaluateRule(ComponentGroupRule rule, List<EleComponentDto> components)
        {
            // 执行单条规则评估
            var resList = EvaluateRulesWithResourcePool(new List<ComponentGroupRule> { rule }, components, out _, out _);
            // 若命中则返回第一项
            if (resList.Count > 0) return resList[0];

            // 未命中时返回默认失败结果
            return new RuleMatchResult
            {
                // 赋值规则 ID
                RuleId = rule.Id,
                // 赋值规则名称
                RuleName = rule.Name,
                // 赋值目标元件组
                TargetGroup = rule.TargetGroup,
                // 默认套数
                Quantity = 1,
                // 标记未匹配
                IsMatched = false,
                // 提示未匹配
                DetailInfo = "未匹配到任何元件"
            };
        }

        /// <summary>
        /// 校验单个元件是否满足属性过滤条件 (供守卫等外部调用兼容)
        /// </summary>
        private static bool MatchSingleComponentProperties(EleComponentDto ec, string propertyExpr)
        {
            // 校验元件与表达式非空
            if (ec == null || string.IsNullOrWhiteSpace(propertyExpr)) return false;

            // 按空格拆分表达式各项
            var parts = propertyExpr.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            // 拆分结果非空校验
            if (parts.Length == 0) return false;

            // 第 1 项: 元件名称匹配 (支持通配符 *)
            string targetName = parts[0].Trim();
            // 检查名称匹配
            if (targetName != "*" && !ec.EleName.Contains(targetName))
            {
                // 名称不包含则返回 false
                return false;
            }

            // 第 2 项及后续: 各个属性过滤器
            for (int i = 1; i < parts.Length; i++)
            {
                // 拆分属性键值对
                var kv = parts[i].Split(new[] { ':', '：' }, 2);
                // 校验键值对格式
                if (kv.Length < 2) continue;

                // 属性名称
                string propName = kv[0].Trim();
                // 属性比较条件
                string propCondition = kv[1].Trim();

                // 按属性名称分别校验
                switch (propName)
                {
                    case "电流":
                        // 电流比较
                        if (!EvaluateCurrentCondition(ec.EleCurrent, propCondition)) return false;
                        break;
                    case "型号":
                        // 型号包含校验
                        if (string.IsNullOrEmpty(propCondition))
                        {
                            if (!string.IsNullOrEmpty(ec.EleNorms)) return false;
                        }
                        else
                        {
                            if (!ec.EleNorms.Contains(propCondition)) return false;
                        }
                        break;
                    case "极数":
                        // 极数比较
                        if (!EvaluatePoleCondition(ec.ElePoles, propCondition)) return false;
                        break;
                    case "附件":
                        // 附件包含校验
                        if (string.IsNullOrEmpty(propCondition))
                        {
                            if (!string.IsNullOrWhiteSpace(ec.EleAppendix)) return false;
                        }
                        else
                        {
                            if (!ec.EleAppendix.Contains(propCondition)) return false;
                        }
                        break;
                }
            }

            // 属性均匹配
            return true;
        }

        /// <summary>
        /// 电流条件评估兼容方法 (如 >1000, <=400, 100)
        /// </summary>
        private static bool EvaluateCurrentCondition(string rawCurrent, string condition)
        {
            // 校验字符串非空
            if (string.IsNullOrWhiteSpace(rawCurrent) || string.IsNullOrWhiteSpace(condition)) return false;

            // 提取运算符与目标值
            string op = ">=";
            string valStr = condition;
            if (condition.StartsWith(">=") || condition.StartsWith("=>")) { op = ">="; valStr = condition.Substring(2); }
            else if (condition.StartsWith("<=") || condition.StartsWith("=<")) { op = "<="; valStr = condition.Substring(2); }
            else if (condition.StartsWith("==")) { op = "=="; valStr = condition.Substring(2); }
            else if (condition.StartsWith("!=")) { op = "!="; valStr = condition.Substring(2); }
            else if (condition.StartsWith(">")) { op = ">"; valStr = condition.Substring(1); }
            else if (condition.StartsWith("<")) { op = "<"; valStr = condition.Substring(1); }
            else if (condition.StartsWith("=")) { op = "=="; valStr = condition.Substring(1); }
            else { op = "=="; valStr = condition; }

            // 调用纯数值比较
            return EvaluateCurrentConditionWithOp(rawCurrent, op, valStr);
        }

        /// <summary>
        /// 极数条件评估兼容方法 (如 3P, 4P, 3)
        /// </summary>
        private static bool EvaluatePoleCondition(string rawPoles, string condition)
        {
            // 调用纯极数比较
            return EvaluatePoleConditionWithOp(rawPoles, "==", condition);
        }
    }

    #endregion
}

