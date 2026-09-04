using System.Text.Json.Serialization;
using ExcelAddInDemo.Models;

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

        // 二次图回路方案与 DWG 图纸本地目录配置选项分节
        [JsonPropertyName("SecondaryCircuitSettings")]
        public SecondaryCircuitSettings SecondaryCircuit { get; set; } = new SecondaryCircuitSettings();
    }

    // 后端 API 接口连接配置数据模型
    public class ApiSettings
    {
        // 远程商城 API 服务的基准访问地址 (默认: https://mall.xingren.online) --硬编码--
        public string BaseUrl { get; set; } = "https://mall.xingren.online";

        // 元器件分页检索接口相对路径 (默认: /api/api/Component/GetPagedList) --硬编码--
        public string ComponentGetPagedListEndpoint { get; set; } = "/api/api/Component/GetPagedList";

        // 网络 API 请求的超时等待时间（单位：秒）
        public int TimeoutSeconds { get; set; } = 15;

        // 请求失败时的最大自动重试次数
        public int MaxRetryCount { get; set; } = 2;
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

        // 获取当前配置实例对应的箱柜定义名称前缀值对象 (零堆分配)
        [JsonIgnore]
        public CabinetPrefixConfig Prefixes => new CabinetPrefixConfig(this);

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
        public int TemplateDetailTotalRows => CabTolsumRowIndex >= CabDetRowIndex ? (CabTolsumRowIndex - CabDetRowIndex + 1) : 28;

        // 模板箱柜明细整块（含大标题至总计行）的总行数 (通过 CabTolsumRowIndex - (CabDetRowIndex - 3) + 1 结合计算得出)
        [JsonIgnore]
        public int TemplateDetailBlockTotalRows => CabTolsumRowIndex >= (CabDetRowIndex - 3) ? (CabTolsumRowIndex - (CabDetRowIndex - 3) + 1) : 31;

        // 模板箱柜元器件预留行数 (结合 CabTolsumRowIndex 与 CabDetRowIndex，减去表头2行与默认计费6行得出)
        [JsonIgnore]
        public int DefaultComponentRowCount => Math.Max(1, (CabTolsumRowIndex - CabDetRowIndex) - 7);

        // 超链接跳转时视口首行滚动的行号修正值 (偏移量，默认为 0)
        public int ScrollRowOffset { get; set; } = 0;

        // 右键菜单“新建箱柜”按钮的显示文本 (默认: 新建箱柜)
        public string NewCabinetMenuCaption { get; set; } = "新建箱柜";

        // 右键菜单“新建箱柜”按钮的唯一 Tag 标识 (默认: CT_BTN_NEW_CABINET)
        public string NewCabinetMenuTag { get; set; } = "CT_BTN_NEW_CABINET";

        // 右键菜单“识别参数并匹配物料”按钮的显示文本 (默认: 识别参数并匹配物料) --硬编码--
        public string ParseMatchComponentsMenuCaption { get; set; } = "识别参数并匹配物料";

        // 右键菜单“识别参数并匹配物料”按钮的唯一 Tag 标识 (默认: CT_BTN_PARSE_MATCH_COMPONENTS) --硬编码--
        public string ParseMatchComponentsMenuTag { get; set; } = "CT_BTN_PARSE_MATCH_COMPONENTS";

        // 【项目信息】工作表中【分类汇总】区域的起始物理行号 (默认: 29)
        public int ProjectInfoCategorySummaryStartRow { get; set; } = 29;

        // 扫描【项目信息】工作表【分类汇总】区域时的最大扫描行数限制 (默认: 100)
        public int ProjectInfoCategorySummaryMaxScanRows { get; set; } = 100;

        // 是否启用基于 WebView2 + Vue 3 的自定义业务右键菜单 (默认: true，若为 false 则完全使用 Excel 原生右键菜单)
        public bool UseCustomContextMenu { get; set; } = true;
    }

    /// <summary>
    /// 二次图回路方案管理与 DWG 本地图纸目录配置实体
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class SecondaryCircuitSettings
    {
        // 二次排布图 DWG 本地图纸目录绝对路径
        public string LayoutDwgDirectory { get; set; } = string.Empty;

        // 回路代号原理图 DWG 本地图纸目录绝对路径
        public string CircuitDwgDirectory { get; set; } = string.Empty;
    }
}

