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
        AutoByRatio = 2,

        /// <summary>
        /// 所有匹配元件数量之和
        /// </summary>
        SumOfMatches = 3
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
                RawExpression = "凝露控制器|防潮控制器"
            });

            // 2. 双V3A电压电流表测量组
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "双V3A测量回路",
                TargetGroup = "*双V3A",
                Policy = QuantityPolicy.AutoByRatio,
                Priority = priority++,
                RawExpression = "3电流表/2电压表"
            });

            // 3. 1V3A1Z 电流电压转换开关测量组
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "1V3A1Z转换开关测量",
                TargetGroup = "*1V3A1Z",
                Policy = QuantityPolicy.AutoByRatio,
                Priority = priority++,
                RawExpression = "3电流表/1电压表/1转换开关"
            });

            // 4. 1V3A (无转换开关)
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "1V3A无转换测量",
                TargetGroup = "*1V3A",
                Policy = QuantityPolicy.AutoByRatio,
                Priority = priority++,
                RawExpression = "(3电流表/1电压表) &-转换开关"
            });

            // 5. 双电源大电流报警回路
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "双电源大电流报警",
                TargetGroup = "*ATS+MXOF",
                Policy = QuantityPolicy.FollowMainElement,
                Priority = priority++,
                RawExpression = "双电源# 附件:1"
            });

            // 6. 变频器控保组
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "变频器控保回路",
                TargetGroup = "*变频器控保",
                Policy = QuantityPolicy.AutoByRatio,
                Priority = priority++,
                RawExpression = "1变频器/1控制保护开关"
            });

            // 7. KM变频器回路
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "KM变频器联动",
                TargetGroup = "*KM变频器",
                Policy = QuantityPolicy.Fixed,
                Priority = priority++,
                RawExpression = "变频器&接触器"
            });

            // 8. 单变频器回路 (排除控保与接触器)
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "单变频器独立回路",
                TargetGroup = "*单变频器",
                Policy = QuantityPolicy.Fixed,
                Priority = priority++,
                RawExpression = "变频器&-控制保护开关&-接触器"
            });

            // 9. 软启动回路
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "软启动控制回路",
                TargetGroup = "*软启动",
                Policy = QuantityPolicy.FollowMainElement,
                Priority = priority++,
                RawExpression = "软启动#"
            });

            // 10. FA风机消防电源联动
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "FA风机消防电源监控",
                TargetGroup = "*FA风机",
                Policy = QuantityPolicy.FollowMainElement,
                Priority = priority++,
                RawExpression = "(热继电器#|控制保护开关#) &消防电源监控&-时控开关"
            });

            // 11. BA风机回路
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "BA风机楼控回路",
                TargetGroup = "*BA风机",
                Policy = QuantityPolicy.FollowMainElement,
                Priority = priority++,
                RawExpression = "(热继电器#|控制保护开关#) &-消防电源监控&-变频器&-时控开关"
            });

            // 12. 电能表回路
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "电能表计量回路",
                TargetGroup = "*电能表",
                Policy = QuantityPolicy.FollowMainElement,
                Priority = priority++,
                RawExpression = "电能表#"
            });

            // 13. 直接式表通讯
            config.Rules.Add(new ComponentGroupRule
            {
                Name = "直接式表通讯",
                TargetGroup = "*表通讯",
                Policy = QuantityPolicy.FollowMainElement,
                Priority = priority++,
                RawExpression = "直接式表# 型号:0 &-多功能表"
            });

            // 将所有初始规则的 RawExpression 自动解析填充进 ConditionTree
            foreach (var rule in config.Rules)
            {
                if (!string.IsNullOrWhiteSpace(rule.RawExpression))
                {
                    rule.ConditionTree = PipelineCompiler.ParseExpressionToTree(rule.RawExpression, rule.Policy);
                }
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
    public static class PipelineCompiler
    {
        /// <summary>
        /// 将经典表达式字符串解析为可视化条件树
        /// </summary>
        public static RuleConditionGroup ParseExpressionToTree(string expression, QuantityPolicy policy)
        {
            var rootGroup = new RuleConditionGroup { Op = LogicalOperator.And };
            if (string.IsNullOrWhiteSpace(expression)) return rootGroup;

            // 预处理: 将中文括号替换为标准括号
            string expr = expression.Trim().Replace("（", "(").Replace("）", ")");

            // 处理顶层带有 OR 的逻辑分组
            if (expr.Contains("|") && !expr.StartsWith("(") && !expr.EndsWith(")"))
            {
                // 顶层若包含 | 且无外层整体括号包围，拆分为 OR 组
                var orGroup = new RuleConditionGroup { Op = LogicalOperator.Or };
                var segments = SplitByTopLevelOperator(expr, '|');
                foreach (var seg in segments)
                {
                    if (string.IsNullOrWhiteSpace(seg)) continue;
                    var subGroup = ParseSimpleExpressionBlock(seg);
                    orGroup.SubGroups.Add(subGroup);
                }
                return orGroup;
            }

            // 否则按标准复合块解析
            return ParseSimpleExpressionBlock(expr);
        }

        /// <summary>
        /// 解析一个复合条件块 (处理 & 逻辑与括号)
        /// </summary>
        private static RuleConditionGroup ParseSimpleExpressionBlock(string block)
        {
            var group = new RuleConditionGroup { Op = LogicalOperator.And };
            string text = block.Trim();

            // 若整体被括号包围，脱去外层括号
            if (text.StartsWith("(") && text.EndsWith(")"))
            {
                // 检查是否是一对闭合的整体括号
                if (IsBalancedParentheses(text.Substring(1, text.Length - 2)))
                {
                    text = text.Substring(1, text.Length - 2).Trim();
                }
            }

            // 按顶层 & 拆分
            var andSegments = SplitByTopLevelOperator(text, '&');
            foreach (var seg in andSegments)
            {
                string s = seg.Trim();
                if (string.IsNullOrWhiteSpace(s)) continue;

                // 若分段内部是带有 | 的括号子表达式
                if (s.StartsWith("(") && s.EndsWith(")"))
                {
                    string inner = s.Substring(1, s.Length - 2).Trim();
                    if (inner.Contains("|"))
                    {
                        var subOrGroup = new RuleConditionGroup { Op = LogicalOperator.Or };
                        var orParts = SplitByTopLevelOperator(inner, '|');
                        foreach (var part in orParts)
                        {
                            var node = ParseAtomNode(part.Trim());
                            if (node != null) subOrGroup.Nodes.Add(node);
                        }
                        group.SubGroups.Add(subOrGroup);
                        continue;
                    }
                }

                // 普通原子节点解析
                var atomNode = ParseAtomNode(s);
                if (atomNode != null)
                {
                    group.Nodes.Add(atomNode);
                }
            }

            return group;
        }

        /// <summary>
        /// 解析单个原子条件节点 (如 "双电源# 电流:>1000 型号:WATSG 附件:1 极数:3P" 或 "3电流表/1电压表")
        /// </summary>
        public static RuleConditionNode? ParseAtomNode(string rawAtom)
        {
            if (string.IsNullOrWhiteSpace(rawAtom)) return null;
            string atom = rawAtom.Trim();

            var node = new RuleConditionNode();

            // 1. 检查是否为成组比例语法 (包含 '/')
            if (atom.Contains("/"))
            {
                node.Mode = ElementMatchMode.RatioMember;
                var ratioSegments = atom.Split('/');
                var ratioList = new List<RatioItem>();

                foreach (var rSeg in ratioSegments)
                {
                    string segText = rSeg.Trim();
                    if (string.IsNullOrEmpty(segText)) continue;

                    var rItem = new RatioItem();
                    // 提取前缀数字比例 (如 3电流表)
                    var matchNum = Regex.Match(segText, @"^(\d+)(.*)$");
                    if (matchNum.Success)
                    {
                        rItem.Count = int.Parse(matchNum.Groups[1].Value);
                        segText = matchNum.Groups[2].Value.Trim();
                    }
                    else
                    {
                        rItem.Count = 1;
                    }

                    // 拆分元件名称与可能跟随的属性
                    var parts = segText.Split(' ');
                    rItem.ElementName = parts[0].Trim();
                    for (int p = 1; p < parts.Length; p++)
                    {
                        var filter = ParsePropertyFilter(parts[p]);
                        if (filter != null) rItem.PropertyFilters.Add(filter);
                    }

                    ratioList.Add(rItem);
                }

                node.RatioItems = ratioList;
                node.ElementName = string.Join("/", ratioList.Select(r => $"{r.Count}{r.ElementName}"));
                return node;
            }

            // 2. 检查模式 (排除 - 或 数量驱动 #)
            if (atom.StartsWith("-"))
            {
                node.Mode = ElementMatchMode.MustExclude;
                atom = atom.Substring(1).Trim();
            }
            else if (atom.Contains("#"))
            {
                node.Mode = ElementMatchMode.MainDriver;
                atom = atom.Replace("#", "").Trim();
            }
            else
            {
                node.Mode = ElementMatchMode.MustInclude;
            }

            // 3. 拆分元件名称与属性
            var spaceParts = atom.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (spaceParts.Length > 0)
            {
                node.ElementName = spaceParts[0].Trim();

                // 解析后续各个属性过滤项
                for (int i = 1; i < spaceParts.Length; i++)
                {
                    var filter = ParsePropertyFilter(spaceParts[i]);
                    if (filter != null)
                    {
                        node.PropertyFilters.Add(filter);
                    }
                }
            }

            return node;
        }

        /// <summary>
        /// 解析单个属性过滤器 (如 "电流:>1000", "型号:WATSG", "极数:3P", "附件:1")
        /// </summary>
        private static PropertyFilterItem? ParsePropertyFilter(string propText)
        {
            if (string.IsNullOrWhiteSpace(propText)) return null;

            var colonParts = propText.Split(new[] { ':', '：' }, 2);
            if (colonParts.Length == 0) return null;

            string propHeader = colonParts[0].Trim();
            string propVal = colonParts.Length > 1 ? colonParts[1].Trim() : "";

            var filter = new PropertyFilterItem();

            // 识别属性类型
            switch (propHeader)
            {
                case "电流":
                    filter.PropertyType = "Current";
                    break;
                case "型号":
                    filter.PropertyType = "Model";
                    break;
                case "极数":
                    filter.PropertyType = "Poles";
                    break;
                case "附件":
                    filter.PropertyType = "Appendix";
                    break;
                default:
                    filter.PropertyType = propHeader;
                    break;
            }

            // 解析操作符与值
            if (propVal.StartsWith(">="))
            {
                filter.Operator = ">=";
                filter.Value = propVal.Substring(2).Trim();
            }
            else if (propVal.StartsWith("<="))
            {
                filter.Operator = "<=";
                filter.Value = propVal.Substring(2).Trim();
            }
            else if (propVal.StartsWith(">"))
            {
                filter.Operator = ">";
                filter.Value = propVal.Substring(1).Trim();
            }
            else if (propVal.StartsWith("<"))
            {
                filter.Operator = "<";
                filter.Value = propVal.Substring(1).Trim();
            }
            else if (propVal.StartsWith("!="))
            {
                filter.Operator = "!=";
                filter.Value = propVal.Substring(2).Trim();
            }
            else if (propVal.StartsWith("=="))
            {
                filter.Operator = "==";
                filter.Value = propVal.Substring(2).Trim();
            }
            else
            {
                // 默认文本匹配采用 contains 或等于
                filter.Operator = (filter.PropertyType == "Current" || filter.PropertyType == "Poles") ? "==" : "contains";
                filter.Value = propVal;
            }

            return filter;
        }

        /// <summary>
        /// 将可视化条件树转换为经典文本表达式字符串
        /// </summary>
        public static string BuildExpressionFromTree(RuleConditionGroup group)
        {
            if (group == null) return "";

            var parts = new List<string>();

            // 序列化所有当前层节点
            if (group.Nodes != null)
            {
                foreach (var node in group.Nodes)
                {
                    string nodeStr = BuildAtomNodeString(node);
                    if (!string.IsNullOrEmpty(nodeStr)) parts.Add(nodeStr);
                }
            }

            // 序列化子条件组
            if (group.SubGroups != null)
            {
                foreach (var sub in group.SubGroups)
                {
                    string subStr = BuildExpressionFromTree(sub);
                    if (!string.IsNullOrEmpty(subStr))
                    {
                        // 若子组包含多个条件，用括号包裹
                        parts.Add($"({subStr})");
                    }
                }
            }

            string opSymbol = group.Op == LogicalOperator.Or ? " | " : " & ";
            return string.Join(opSymbol, parts);
        }

        /// <summary>
        /// 格式化单个条件节点为文本字符串
        /// </summary>
        private static string BuildAtomNodeString(RuleConditionNode node)
        {
            if (node == null) return "";

            // 成组比例模式
            if (node.Mode == ElementMatchMode.RatioMember && node.RatioItems != null && node.RatioItems.Count > 0)
            {
                var rList = new List<string>();
                foreach (var r in node.RatioItems)
                {
                    var sbItem = new StringBuilder();
                    sbItem.Append($"{r.Count}{r.ElementName}");
                    if (r.PropertyFilters != null)
                    {
                        foreach (var f in r.PropertyFilters)
                        {
                            sbItem.Append($" {FormatFilterForExpression(f)}");
                        }
                    }
                    rList.Add(sbItem.ToString());
                }
                return string.Join("/", rList);
            }

            var sb = new StringBuilder();
            if (node.Mode == ElementMatchMode.MustExclude)
            {
                sb.Append($"-{node.ElementName}");
            }
            else if (node.Mode == ElementMatchMode.MainDriver)
            {
                sb.Append($"{node.ElementName}#");
            }
            else
            {
                sb.Append(node.ElementName);
            }

            if (node.PropertyFilters != null)
            {
                foreach (var filter in node.PropertyFilters)
                {
                    sb.Append($" {FormatFilterForExpression(filter)}");
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// 格式化属性过滤项
        /// </summary>
        private static string FormatFilterForExpression(PropertyFilterItem filter)
        {
            string label = filter.PropertyType switch
            {
                "Current" => "电流",
                "Model" => "型号",
                "Poles" => "极数",
                "Appendix" => "附件",
                _ => filter.PropertyType
            };

            string opStr = filter.Operator switch
            {
                "==" or "contains" => "",
                _ => filter.Operator
            };

            return $"{label}:{opStr}{filter.Value}";
        }

        /// <summary>
        /// 顶层运算符安全拆分 (忽略括号内部的运算符)
        /// </summary>
        private static List<string> SplitByTopLevelOperator(string text, char op)
        {
            var list = new List<string>();
            int depth = 0;
            int lastIdx = 0;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '(' || c == '（') depth++;
                else if (c == ')' || c == '）') depth--;
                else if (c == op && depth == 0)
                {
                    list.Add(text.Substring(lastIdx, i - lastIdx).Trim());
                    lastIdx = i + 1;
                }
            }

            if (lastIdx < text.Length)
            {
                list.Add(text.Substring(lastIdx).Trim());
            }

            return list;
        }

        /// <summary>
        /// 校验括号是否对称匹配
        /// </summary>
        private static bool IsBalancedParentheses(string str)
        {
            int count = 0;
            foreach (char c in str)
            {
                if (c == '(' || c == '（') count++;
                if (c == ')' || c == '）') count--;
                if (count < 0) return false;
            }
            return count == 0;
        }
    }

    /// <summary>
    /// 规则管道高效内存执行评估器 (支持箱柜元件动态资源池扣减机制)
    /// </summary>
    public static class PipelineEvaluator
    {
        /// <summary>
        /// 针对整个箱柜元件集合，按规则优先级依次进行带【动态资源池扣减】的完整评估
        /// </summary>
        /// <param name="rules">启用的规则管道列表 (已按优先级排序)</param>
        /// <param name="components">箱柜一次元件集合</param>
        /// <param name="executionLogs">执行与扣减过程诊断日志列表</param>
        /// <returns>命中的规则匹配结果列表</returns>
        public static List<RuleMatchResult> EvaluateRulesWithResourcePool(
            List<ComponentGroupRule> rules,
            List<EleComponentDto> components,
            out List<string> executionLogs,
            out List<EleComponentDto> remainingComponents)
        {
            var matchedResults = new List<RuleMatchResult>();
            executionLogs = new List<string>();

            if (rules == null || rules.Count == 0 || components == null || components.Count == 0)
            {
                executionLogs.Add("[资源池] 元件列表或规则列表为空，无需评估。");
                remainingComponents = components ?? new List<EleComponentDto>();
                return matchedResults;
            }

            // 1. 克隆一份深拷贝作为当前箱柜的动态工作资源池
            var workingPool = components.Select(c => new EleComponentDto
            {
                RowIndex = c.RowIndex,
                EleName = c.EleName,
                EleNorms = c.EleNorms,
                EleNums = c.EleNums,
                EleCurrent = c.EleCurrent,
                ElePoles = c.ElePoles,
                EleAppendix = c.EleAppendix
            }).ToList();

            // 打印初始元件库存明细
            var initialCompSummary = workingPool
                .Where(c => c.EleNums > 0)
                .Select(c => $"[{c.EleName}]: {c.EleNums}件");
            executionLogs.Add($"📦 【初始元件库存池】: {string.Join(", ", initialCompSummary)}");

            int step = 1;
            // 2. 依次按规则优先级从上到下遍历评估
            foreach (var rule in rules.Where(r => r.Enabled).OrderBy(r => r.Priority))
            {
                string expr = rule.RawExpression;
                if (string.IsNullOrWhiteSpace(expr))
                {
                    expr = PipelineCompiler.BuildExpressionFromTree(rule.ConditionTree);
                }
                if (string.IsNullOrWhiteSpace(expr)) continue;

                // 尝试在当前剩余资源池中评估本规则，并准备扣减清单
                var pendingDeductions = new List<(string filterStr, int deductTotal)>();
                int computedQuantity = 0;

                bool isMatched = EvaluateExpressionWithPool(expr, workingPool, pendingDeductions, out computedQuantity);

                if (isMatched && computedQuantity > 0)
                {
                    // 确定生成套数
                    int finalQty = computedQuantity;
                    if (rule.Policy == QuantityPolicy.Fixed && rule.FixedQuantity > 0)
                    {
                        finalQty = rule.FixedQuantity;
                    }

                    // 提交扣减: 正式从工作资源池中扣除消耗的元件数量 (跨多行智能合并扣减)
                    var deductSummary = new List<string>();
                    foreach (var (filterStr, deductTotal) in pendingDeductions)
                    {
                        int remainingToDeduct = deductTotal;
                        int totalDeductedForThis = 0;
                        foreach (var comp in workingPool)
                        {
                            if (comp.EleNums <= 0) continue;
                            if (MatchSingleComponentProperties(comp, filterStr))
                            {
                                int deductThis = Math.Min(comp.EleNums, remainingToDeduct);
                                comp.EleNums -= deductThis;
                                remainingToDeduct -= deductThis;
                                totalDeductedForThis += deductThis;
                                if (remainingToDeduct <= 0) break;
                            }
                        }

                        // 统计扣减后该类元件在当前池中的剩余总数
                        int poolRemaining = workingPool
                            .Where(c => MatchSingleComponentProperties(c, filterStr))
                            .Sum(c => c.EleNums);

                        deductSummary.Add($"[{filterStr}] 消耗 {totalDeductedForThis}件 (池剩余 {poolRemaining}件)");
                    }

                    // 统计当前规则匹配结果
                    var matchRes = new RuleMatchResult
                    {
                        RuleId = rule.Id,
                        RuleName = rule.Name,
                        TargetGroup = rule.TargetGroup,
                        Quantity = finalQty,
                        IsMatched = true,
                        DetailInfo = $"匹配成功! 生成 {finalQty} 套 [{rule.TargetGroup}] (消耗: {string.Join(", ", deductSummary)})"
                    };
                    matchedResults.Add(matchRes);

                    executionLogs.Add($"[步骤 {step}] ✅ 规则 [{rule.Name}] 命中! 生成 {finalQty} 套 [{rule.TargetGroup}]");
                    executionLogs.Add($"    └── 消耗扣减: {string.Join("; ", deductSummary)}");
                }
                else
                {
                    executionLogs.Add($"[步骤 {step}] ⚪ 规则 [{rule.Name}] 未命中 (条件不满足或资源池不足)");
                }

                step++;
            }

            // 3. 输出最终剩余元件库存快照
            var finalRemaining = workingPool.Select(c => $"[{c.EleName}]: {c.EleNums}件");
            executionLogs.Add($"🏁 【最终剩余元件库存】: {string.Join(", ", finalRemaining)}");
            remainingComponents = workingPool;

            return matchedResults;
        }

        /// <summary>
        /// 对单条规则管道在给定元件集合上进行执行评估 (兼容单规则沙盒测试)
        /// </summary>
        public static RuleMatchResult EvaluateRule(ComponentGroupRule rule, List<EleComponentDto> components)
        {
            var resList = EvaluateRulesWithResourcePool(new List<ComponentGroupRule> { rule }, components, out _, out _);
            if (resList.Count > 0) return resList[0];

            return new RuleMatchResult
            {
                RuleId = rule.Id,
                RuleName = rule.Name,
                TargetGroup = rule.TargetGroup,
                Quantity = 1,
                IsMatched = false,
                DetailInfo = "未匹配到任何元件"
            };
        }

        /// <summary>
        /// 带资源池待扣减追踪的表达式核心求值算法
        /// </summary>
        private static bool EvaluateExpressionWithPool(
            string rawCode,
            List<EleComponentDto> pool,
            List<(string filterStr, int deductTotal)> pendingDeductions,
            out int totalQuantity)
        {
            totalQuantity = 0;
            if (string.IsNullOrWhiteSpace(rawCode)) return false;

            string compileCode = rawCode.Replace("（", "(").Replace("）", ")");
            var atoms = compileCode.Split(new[] { '&', '|', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);

            int accumulatedQty = 0;

            foreach (var rawAtom in atoms)
            {
                string atom = rawAtom.Trim();
                if (string.IsNullOrEmpty(atom)) continue;

                int singleQty = 0;
                bool isAtomTrue = false;

                if (atom.Contains("/"))
                {
                    // 比例模式求值并记录待扣减
                    isAtomTrue = EvaluateRatioAtomWithPool(atom, pool, pendingDeductions, out singleQty);
                }
                else
                {
                    // 普通包含/排除/主控求值并记录待扣减
                    isAtomTrue = EvaluateNormalAtomWithPool(atom, pool, pendingDeductions, out singleQty);
                }

                // 替换布尔值进入最终表达式
                compileCode = compileCode.Replace(rawAtom, isAtomTrue ? " true " : " false ");
                accumulatedQty += singleQty;
            }

            // 将 & 和 | 替换为 DataTable 能够识别的 and / or
            compileCode = compileCode.Replace("|", " or ").Replace("&", " and ");

            bool finalRes = false;
            try
            {
                object evalRes = new DataTable().Compute(compileCode, "");
                finalRes = Convert.ToBoolean(evalRes);
            }
            catch
            {
                finalRes = false;
            }

            totalQuantity = accumulatedQty > 0 ? accumulatedQty : 1;
            return finalRes;
        }

        /// <summary>
        /// 评估成组比例原子项 (如 "3互感器/1直接式表") 并根据最小公约计算成套数
        /// </summary>
        private static bool EvaluateRatioAtomWithPool(
            string ratioAtom,
            List<EleComponentDto> pool,
            List<(string filterStr, int deductTotal)> pendingDeductions,
            out int ratioGroups)
        {
            ratioGroups = 0;
            var segments = ratioAtom.Split('/');
            if (segments.Length == 0) return false;

            int minPossibleSets = int.MaxValue;
            var memberInfos = new List<(string filterStr, int baseRatio, int availableCount)>();

            // 1. 扫描各个成员在当前资源池中的可用数量与可组成套数
            foreach (var seg in segments)
            {
                string segText = seg.Trim();
                if (string.IsNullOrEmpty(segText)) return false;

                // 提取基准比例前缀 (如 3互感器 -> baseRatio=3, filterStr=互感器)
                var match = Regex.Match(segText, @"^(\d+)(.*)$");
                int baseRatio = 1;
                string filterStr = segText;
                if (match.Success)
                {
                    baseRatio = int.Parse(match.Groups[1].Value);
                    filterStr = match.Groups[2].Value.Trim();
                }

                // 统计当前池中匹配该条件的元件总可用数量
                int availableCount = pool
                    .Where(c => c.EleNums > 0 && MatchSingleComponentProperties(c, filterStr))
                    .Sum(c => c.EleNums);

                if (availableCount < baseRatio)
                {
                    // 只要有一个成员连基础 1 套都凑不够，则该比例不成立
                    return false;
                }

                // 计算该成员单独能支持的最大套数
                int setsForMember = availableCount / baseRatio;
                if (setsForMember < minPossibleSets)
                {
                    minPossibleSets = setsForMember;
                }

                memberInfos.Add((filterStr, baseRatio, availableCount));
            }

            if (minPossibleSets <= 0 || minPossibleSets == int.MaxValue)
            {
                return false;
            }

            // 2. 确认可成套数 minPossibleSets (如 1 套)
            ratioGroups = minPossibleSets;

            // 3. 记录待扣减额度: 各成员扣减 minPossibleSets * baseRatio
            foreach (var member in memberInfos)
            {
                int deductCount = minPossibleSets * member.baseRatio;
                pendingDeductions.Add((member.filterStr, deductCount));
            }

            return true;
        }

        /// <summary>
        /// 评估普通原子条件项 (如 "直接式表#", "消防电源监控", "-时控开关")
        /// </summary>
        private static bool EvaluateNormalAtomWithPool(
            string atomStr,
            List<EleComponentDto> pool,
            List<(string filterStr, int deductTotal)> pendingDeductions,
            out int matchQty)
        {
            matchQty = 0;
            string cleanStr = atomStr.Trim();
            bool isMustExclude = cleanStr.StartsWith("-") || cleanStr.Contains("-");
            bool isMainDriver = cleanStr.Contains("#");

            string filterStr = cleanStr.Replace("#", "").Replace("-", "").Trim();

            // 统计当前池中可用数量
            int availableSum = pool
                .Where(c => c.EleNums > 0 && MatchSingleComponentProperties(c, filterStr))
                .Sum(c => c.EleNums);

            if (isMustExclude)
            {
                // 排除模式: 池中数量必须为 0
                return availableSum == 0;
            }
            else if (isMainDriver)
            {
                // 主控模式: 数量大于 0 且将其作为输出套数，并加入待扣减
                if (availableSum > 0)
                {
                    matchQty = availableSum;
                    pendingDeductions.Add((filterStr, availableSum));
                    return true;
                }
                return false;
            }
            else
            {
                // 普通包含模式: 池中数量大于 0 即可 (作为关联条件存在)
                return availableSum > 0;
            }
        }

        /// <summary>
        /// 校验单个元件是否满足属性过滤条件 (名称、电流、型号、极数、附件)
        /// </summary>
        private static bool MatchSingleComponentProperties(EleComponentDto ec, string propertyExpr)
        {
            if (ec == null || string.IsNullOrWhiteSpace(propertyExpr)) return false;

            var parts = propertyExpr.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return false;

            // 第 1 项: 元件名称匹配 (支持通配符 *)
            string targetName = parts[0].Trim();
            if (targetName != "*" && !ec.EleName.Contains(targetName))
            {
                return false;
            }

            // 第 2 项及后续: 各个属性过滤器
            for (int i = 1; i < parts.Length; i++)
            {
                var kv = parts[i].Split(new[] { ':', '：' }, 2);
                if (kv.Length < 2) continue;

                string propName = kv[0].Trim();
                string propCondition = kv[1].Trim();

                switch (propName)
                {
                    case "电流":
                        if (!EvaluateCurrentCondition(ec.EleCurrent, propCondition)) return false;
                        break;
                    case "型号":
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
                        if (!EvaluatePoleCondition(ec.ElePoles, propCondition)) return false;
                        break;
                    case "附件":
                        if (string.IsNullOrEmpty(propCondition))
                        {
                            // 必须无附件
                            if (!string.IsNullOrWhiteSpace(ec.EleAppendix)) return false;
                        }
                        else
                        {
                            // 必须包含附件标识
                            if (!ec.EleAppendix.Contains(propCondition)) return false;
                        }
                        break;
                }
            }

            return true;
        }

        /// <summary>
        /// 电流条件评估 (如 >1000, <=400, 100)
        /// </summary>
        private static bool EvaluateCurrentCondition(string rawCurrent, string condition)
        {
            if (string.IsNullOrWhiteSpace(rawCurrent) || string.IsNullOrWhiteSpace(condition)) return false;

            string cleanCur = rawCurrent.Replace("A", "").Trim();
            if (!double.TryParse(cleanCur, NumberStyles.Any, CultureInfo.InvariantCulture, out double curVal))
            {
                return false;
            }

            try
            {
                // 拼接 DataTable 表达式如 "1200>1000"
                string expr = condition.StartsWith(">") || condition.StartsWith("<") || condition.StartsWith("=")
                    ? $"{curVal}{condition}"
                    : $"{curVal}=={condition}";

                object res = new DataTable().Compute(expr, "");
                return Convert.ToBoolean(res);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 极数条件评估 (如 3P, 4P, 3)
        /// </summary>
        private static bool EvaluatePoleCondition(string rawPoles, string condition)
        {
            if (string.IsNullOrWhiteSpace(rawPoles) || string.IsNullOrWhiteSpace(condition)) return false;

            string p1 = rawPoles.Trim().ToUpper().Replace("P", "");
            string p2 = condition.Trim().ToUpper().Replace("P", "");

            return p1 == p2;
        }
    }

    #endregion
}
