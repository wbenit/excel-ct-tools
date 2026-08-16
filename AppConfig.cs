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

        // 底部明细小计行使用的 Excel 定义名称前缀
        public string SubsumNamePrefix { get; set; } = "Cab_Subsum_";

        // 底部明细总计行使用的 Excel 定义名称前缀
        public string TolsumNamePrefix { get; set; } = "Cab_Tolsum_";

        // UI 界面及表头强调的默认主题颜色（主色调为 #009688）
        public string DefaultThemeColor { get; set; } = "#009688";

        // 导出的 Excel 表头文字默认字号
        public int HeaderFontSize { get; set; } = 11;

        // 模板用于特征识别匹配的基准行号
        public int FeatureRowIndex { get; set; } = 41;

        // 模板用于特征识别匹配的基准列号
        public int FeatureColumnIndex { get; set; } = 1;

        // 初始明细中顶部汇总行 (Cab_Sum) 的物理行号
        public int CabSumRowIndex { get; set; } = 7;

        // 初始明细中箱柜信息行 (Cab_Det) 的物理行号
        public int CabDetRowIndex { get; set; } = 44;

        // 初始明细中总计行 (Cab_Tolsum) 的物理行号
        public int CabTolsumRowIndex { get; set; } = 71;

        // 模板箱柜明细包含总计行的总行数 A (通过 CabTolsumRowIndex - CabDetRowIndex + 1 动态计算得出)
        [JsonIgnore]
        public int TemplateDetailTotalRows => CabTolsumRowIndex >= CabDetRowIndex ? (CabTolsumRowIndex - CabDetRowIndex + 1) : 25;

        // 超链接跳转时视口首行滚动的行号修正值 (偏移量，默认为 0)
        public int ScrollRowOffset { get; set; } = 0;

        // 右键菜单“新建箱柜”按钮的显示文本 (默认: 新建箱柜)
        public string NewCabinetMenuCaption { get; set; } = "新建箱柜";

        // 右键菜单“新建箱柜”按钮的唯一 Tag 标识 (默认: CT_BTN_NEW_CABINET)
        public string NewCabinetMenuTag { get; set; } = "CT_BTN_NEW_CABINET";
    }
}
