using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 在线查价数据源平台枚举
    /// </summary>
    public enum OnlinePricePlatform
    {
        // 电气天下 (dq123.com / 电小二)
        Dq123 = 0,

        // 天工矩阵 (titanmatrix.com / cpq.titanmatrix.com)
        TitanMatrix = 1
    }

    /// <summary>
    /// 在线查价用户偏好与回填规则配置实体
    /// </summary>
    public class OnlinePriceConfig
    {
        // 当前选中的查询平台 (默认电气天下)
        [JsonPropertyName("platform")]
        public OnlinePricePlatform Platform { get; set; } = OnlinePricePlatform.Dq123;

        // 当前选中的厂商品牌 (如 "全部", "正泰", "德力西", "常熟", "施耐德", "良信" 等)
        [JsonPropertyName("brand")]
        public string Brand { get; set; } = "全部";

        // 型号回填目标列名称 (默认 "C"，设为 "NONE" 或 "不回填" 则不覆盖型号)
        [JsonPropertyName("modelTargetCol")]
        public string ModelTargetCol { get; set; } = "C";

        // 价格回填目标列名称 (默认 "M" 面价/表价，支持 "G" 单价、"K" 成本单价等)
        [JsonPropertyName("priceTargetCol")]
        public string PriceTargetCol { get; set; } = "M";

        // 品牌/厂商回填目标列名称 (默认 "D"，支持自由填写如 "E" 或设为 "NONE" 不回写)
        [JsonPropertyName("brandTargetCol")]
        public string BrandTargetCol { get; set; } = "D";

        // 最大异步网络并发数 (默认 6，防止被远端平台限流) --硬编码: 并发阈值--
        [JsonPropertyName("maxConcurrency")]
        public int MaxConcurrency { get; set; } = 6;

        // 是否仅更新价格而不覆盖原有型号文本
        [JsonPropertyName("onlyUpdatePrice")]
        public bool OnlyUpdatePrice { get; set; } = false;

        // 是否自动去除型号中的括号参数（如从 "NM1-125S/3300 (100A)" 提取 "NM1-125S/3300"）
        [JsonPropertyName("trimParentheses")]
        public bool TrimParentheses { get; set; } = true;
    }

    /// <summary>
    /// 单个元器件查价与回填项传输模型
    /// </summary>
    public class OnlinePriceItemDto
    {
        // Excel 物理行号 (1-indexed)
        [JsonPropertyName("row")]
        public int Row { get; set; }

        // 原始单元格读取文本
        [JsonPropertyName("rawText")]
        public string RawText { get; set; } = string.Empty;

        // 经清洗后提取的检索规格型号
        [JsonPropertyName("searchModel")]
        public string SearchModel { get; set; } = string.Empty;

        // 远端接口匹配返回的官方标准型号
        [JsonPropertyName("matchedModel")]
        public string MatchedModel { get; set; } = string.Empty;

        // 远端接口获取到的官方单价 (面价/表价)
        [JsonPropertyName("price")]
        public double Price { get; set; } = 0.0;

        // 远端返回的所属品牌或生产厂家
        [JsonPropertyName("vendor")]
        public string Vendor { get; set; } = string.Empty;

        // 当前项查询与匹配状态 (如 "Success", "NotFound", "Error")
        [JsonPropertyName("status")]
        public string Status { get; set; } = "Pending";

        // 状态文字描述或异常原因
        [JsonPropertyName("message")]
        public string Message { get; set; } = string.Empty;
    }

    /// <summary>
    /// 查价进度通知数据包
    /// </summary>
    public class OnlinePriceProgressDto
    {
        // 当前已完成项计数
        [JsonPropertyName("current")]
        public int Current { get; set; }

        // 总待查项数量
        [JsonPropertyName("total")]
        public int Total { get; set; }

        // 完成百分比 (0-100)
        [JsonPropertyName("percent")]
        public int Percent { get; set; }

        // 最新处理项的简要描述或日志
        [JsonPropertyName("logText")]
        public string LogText { get; set; } = string.Empty;

        // 最新一项的详细处理数据
        [JsonPropertyName("item")]
        public OnlinePriceItemDto? Item { get; set; }
    }

    /// <summary>
    /// 选区探测概况模型
    /// </summary>
    public class SelectionDetectDto
    {
        // 是否包含有效选区
        [JsonPropertyName("hasSelection")]
        public bool HasSelection { get; set; }

        // 当前工作表名称
        [JsonPropertyName("sheetName")]
        public string SheetName { get; set; } = string.Empty;

        // 探测到的有效元器件行数
        [JsonPropertyName("validRowCount")]
        public int ValidRowCount { get; set; }

        // 选区起始行
        [JsonPropertyName("startRow")]
        public int StartRow { get; set; }

        // 选区结束行
        [JsonPropertyName("endRow")]
        public int EndRow { get; set; }

        // 前 5 个预览型号列表
        [JsonPropertyName("previewModels")]
        public List<string> PreviewModels { get; set; } = new List<string>();
    }
}
