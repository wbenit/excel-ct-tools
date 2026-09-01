using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 元器件参数反查默认配置与常量定义
    /// </summary>
    public static class ComponentMatchDefaults
    {
        // 多结果占位提示词 (当数据库查出多个匹配时写入 D 列)
        public const string MultipleCandidatesText = "点击查询"; // --硬编码-- 多结果占位提示文本

        // 多结果高亮背景颜色: 淡黄色 RGB(255, 242, 204) 对应 OLE 颜色值 0xCCF2FF
        public const int LightYellowOleColor = 0xCCF2FF; // --硬编码-- 淡黄底色 OLE 颜色值 (RGB: 255, 242, 204)

        // 无背景填充常量 (xlNone = -4142)
        public const int XlNoneColorIndex = -4142; // --硬编码-- Excel 无填充颜色索引

        // 元件汇总表标准工作表名称
        public const string ComponentSummarySheetName = "元件汇总表"; // --硬编码-- 元件汇总表默认工作表名称
    }

    /// <summary>
    /// 单条必含字段约束规则实体模型 (例如: 型号必须包含 "NDX"、"12" 等)
    /// </summary>
    public class MustContainRule
    {
        // 规则唯一标识 ID
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        // 是否启用当前约束规则 (默认启用)
        public bool Enabled { get; set; } = true;

        // 必须包含的关键字文本 (如: "NDX", "12", "CVS")
        public string Keyword { get; set; } = string.Empty;
    }

    /// <summary>
    /// 元器件物料匹配多维高级过滤配置模型 (包含品牌偏好与动态必含字段约束)
    /// </summary>
    public class ComponentMatchFilterConfig
    {
        // 当前选中的目标品牌筛选 (为空表示匹配全部品牌)
        public string SelectedBrand { get; set; } = string.Empty;

        // 动态必含字段约束规则集合 (多条规则间为 AND 与关系)
        public List<MustContainRule> MustContainRules { get; set; } = new List<MustContainRule>();

        // 是否开启 D 列单元格点击自动弹起物料搜索框 (默认关闭，仅开启勾选时点击 D 列才弹起搜索框)
        public bool EnableSearchOverlay { get; set; } = false;

        // 列映射配置对象
        public ComponentMatchColumnConfig ColumnConfig { get; set; } = new ComponentMatchColumnConfig();

        /// <summary>
        /// 创建带有默认推荐规则的配置实例
        /// </summary>
        public static ComponentMatchFilterConfig CreateDefault()
        {
            var config = new ComponentMatchFilterConfig();
            // 默认无预置必含约束，由用户在界面动态添加
            return config;
        }
    }

    /// <summary>
    /// 元器件参数识别与物料反查列映射配置模型
    /// </summary>
    public class ComponentMatchColumnConfig
    {
        // 识别后电流输入列 (默认 S 列)
        public string CurrentColumn { get; set; } = "S"; // --硬编码-- 默认电流所在列

        // 识别后极数输入列 (默认 T 列)
        public string PoleColumn { get; set; } = "T"; // --硬编码-- 默认极数所在列

        // 识别后脱扣方式输入列 (默认 U 列)
        public string TripModeColumn { get; set; } = "U"; // --硬编码-- 默认脱扣所在列

        // 数据库反查后名称回填列 (默认 B 列)
        public string NameColumn { get; set; } = "B"; // --硬编码-- 默认元器件名称列

        // 数据库反查后型号回填列 (默认 D 列)
        public string ModelColumn { get; set; } = "D"; // --硬编码-- 默认规格型号列

        // 数据库反查后单价回填列 (默认 G 列)
        public string PriceColumn { get; set; } = "G"; // --硬编码-- 默认销售单价列

        // 数据库反查后备注回填列 (默认 I 列)
        public string RemarkColumn { get; set; } = "I"; // --硬编码-- 默认备注信息列

        // 数据库反查后扩展参数1回填列 (默认 V 列)
        public string Param1Column { get; set; } = "V"; // --硬编码-- 默认参数1输出列

        // 数据库反查后扩展参数2回填列 (默认 W 列)
        public string Param2Column { get; set; } = "W"; // --硬编码-- 默认参数2输出列
    }

    /// <summary>
    /// 品牌聚合统计信息 DTO
    /// </summary>
    public class BrandStatItemDto
    {
        // 品牌名称 (如: 施耐德, ABB, 正泰)
        [JsonPropertyName("brand")]
        public string Brand { get; set; } = string.Empty;

        // 该品牌下的元器件数量统计
        [JsonPropertyName("count")]
        public int Count { get; set; }
    }

    /// <summary>
    /// 远程商城 WebAPI 返回的元器件条目实体 DTO (对应 DrawMall.Ability.Docking.Dto.ComponentDto)
    /// </summary>
    public class ComponentApiDto
    {
        // 主键唯一标识 ID
        [JsonPropertyName("id")]
        public int Id { get; set; }

        // 品牌/厂商 (如: 施耐德, ABB, 常熟)
        [JsonPropertyName("brand")]
        public string? Brand { get; set; }

        // 元器件名称 (如: 微型断路器, 塑壳断路器, 双电源)
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        // 规格型号 (如: iC65N 3P C32A, OTM32F4C12D380C)
        [JsonPropertyName("model")]
        public string? Model { get; set; }

        // 销售单价/参考价格 (元)
        [JsonPropertyName("price")]
        public decimal Price { get; set; }

        // 备注信息
        [JsonPropertyName("remark")]
        public string? Remark { get; set; }

        // 框架/结构扩展参数1 (如: 6kA, OTM10D 250A 4P)
        [JsonPropertyName("param1")]
        public string? Param1 { get; set; }

        // 分类说明/扩展参数2 (如: 400V, 双电源)
        [JsonPropertyName("param2")]
        public string? Param2 { get; set; }

        // 额定电流数值 (如: 32, 100)
        [JsonPropertyName("current")]
        public int? Current { get; set; }

        // 极数 (如: "3", "4", "1+N")
        [JsonPropertyName("poles")]
        public string? Poles { get; set; }

        // 脱扣方式/脱扣器代号 (如: "C", "D", "TM", "MA")
        [JsonPropertyName("tripping")]
        public string? Tripping { get; set; }
    }

    /// <summary>
    /// 远程 WebAPI 分页响应通用包装容器
    /// </summary>
    /// <typeparam name="T">列表项实体类型</typeparam>
    public class ComponentPagedApiResponse<T>
    {
        // 符合查询条件的总记录条数
        [JsonPropertyName("totalCount")]
        public int TotalCount { get; set; }

        // 当前页返回的数据记录集合
        [JsonPropertyName("items")]
        public List<T> Items { get; set; } = new List<T>();

        // 当前分页页码
        [JsonPropertyName("pageIndex")]
        public int PageIndex { get; set; }

        // 每页记录条数
        [JsonPropertyName("pageSize")]
        public int PageSize { get; set; }

        // 总页数
        [JsonPropertyName("totalPages")]
        public int TotalPages { get; set; }
    }

    /// <summary>
    /// 批量反查与回填执行报告模型
    /// </summary>
    public class BatchMatchExecuteResult
    {
        // 执行是否成功
        public bool Success { get; set; } = true;

        // 处理的总行数
        public int TotalRows { get; set; }

        // 唯一匹配成功并完整回填的行数
        public int UniqueMatchCount { get; set; }

        // 多条候选并标注“点击查询”的行数
        public int MultipleMatchCount { get; set; }

        // 未匹配到物料的行数
        public int NoneMatchCount { get; set; }

        // 执行总耗时 (毫秒)
        public long ElapsedMilliseconds { get; set; }

        // 提示消息
        public string Message { get; set; } = string.Empty;
    }
}
