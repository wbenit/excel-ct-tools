using System.Text.Json.Serialization;

namespace ExcelAddInDemo
{
    // 全局根配置对象模型类，对应 appsettings.json 根结构
    public class AppConfig
    {
        // API 接口相关的配置选项分节
        [JsonPropertyName("ApiSettings")]
        public ApiSettings Api { get; set; } = new ApiSettings();

        // Excel 业务处理与样式相关的配置选项分节
        [JsonPropertyName("ExcelSettings")]
        public ExcelSettings Excel { get; set; } = new ExcelSettings();
    }

    // 后端 API 接口连接配置数据模型
    public class ApiSettings
    {
        // 远程 API 服务的基准访问地址（带默认兜底值）
        public string BaseUrl { get; set; } = "https://api.example.com/v1";

        // 网络 API 请求的超时等待时间（单位：秒）
        public int TimeoutSeconds { get; set; } = 30;

        // 请求失败时的最大自动重试次数
        public int MaxRetryCount { get; set; } = 3;
    }

    // Excel 视图、工作表及生成相关的配置选项模型
    public class ExcelSettings
    {
        // 模板工作簿中默认引用的工作表名称
        public string DefaultTemplateSheet { get; set; } = "分类1";

        // 顶部汇总行使用的 Excel 定义名称前缀
        public string SumNamePrefix { get; set; } = "Cab_Sum_";

        // 底部明细块使用的 Excel 定义名称前缀
        public string DetNamePrefix { get; set; } = "Cab_Det_";

        // UI 界面及表头强调的默认主题颜色（主色调为 #009688）
        public string DefaultThemeColor { get; set; } = "#009688";

        // 导出的 Excel 表头文字默认字号
        public int HeaderFontSize { get; set; } = 11;

        // 模板用于特征识别匹配的基准行号
        public int FeatureRowIndex { get; set; } = 41;

        // 模板用于特征识别匹配的基准列号
        public int FeatureColumnIndex { get; set; } = 1;

        // 模板及工作表汇总行的基准行号
        public int TemplateSumRowIndex { get; set; } = 7;

        // 超链接跳转时视口首行滚动的行号修正值 (偏移量，默认为 0)
        public int ScrollRowOffset { get; set; } = 0;
    }
}
