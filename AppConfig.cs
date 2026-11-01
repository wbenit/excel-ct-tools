using System.Text.Json.Serialization;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    // 全局根配置对象模型类，对应 appsettings.json 根结构
    public class AppConfig
    {
        // 便捷全局静态访问当前配置单例实例
        [JsonIgnore]
        public static AppConfig Instance => ConfigManager.Instance.Current;

        // API 接口相关的配置选项分节
        [JsonPropertyName("ApiSettings")]
        public ApiSettings Api { get; set; } = new ApiSettings();

        // Excel 业务处理与样式相关的配置选项分节
        [JsonPropertyName("ExcelSettings")]
        public ExcelSettings Excel { get; set; } = new ExcelSettings();

        // 二次图回路方案与 DWG 图纸本地目录配置选项分节
        [JsonPropertyName("SecondaryCircuitSettings")]
        public SecondaryCircuitSettings SecondaryCircuit { get; set; } = new SecondaryCircuitSettings();

        // 一次成套方案与 DWG 图纸本地目录配置选项分节
        [JsonPropertyName("PrimaryCircuitSettings")]
        public PrimaryCircuitSettings PrimaryCircuit { get; set; } = new PrimaryCircuitSettings();

        // 投标报表导出高级偏好配置选项分节 (自动持久化至 appsettings.json)
        [JsonPropertyName("TenderReportSettings")]
        public TenderReportSettings TenderReport { get; set; } = new TenderReportSettings();

        // 元器件图纸参数匹配侧边窗口 (200x800) 配置选项分节
        [JsonPropertyName("ComponentParamMatchSettings")]
        public ComponentParamMatchSettings ComponentParamMatch { get; set; } = new ComponentParamMatchSettings();

        // AutoCAD 协同通信与离线尺寸测算配置选项分节
        [JsonPropertyName("CadSyncSettings")]
        public CadSyncSettings CadSync { get; set; } = new CadSyncSettings();

        // 工程项目与单体数据管理配置选项分节
        [JsonPropertyName("ProjectSettings")]
        public ProjectSettings Project { get; set; } = new ProjectSettings();

        // 在线元器件查价与第三方平台签名凭据配置选项分节
        [JsonPropertyName("OnlinePriceSearchSettings")]
        public OnlinePriceSearchSettings OnlinePriceSearch { get; set; } = new OnlinePriceSearchSettings();

        // 有色金属大盘期货与现货实时行情配置选项分节
        [JsonPropertyName("MarketPriceSettings")]
        public MarketPriceSettings MarketPrice { get; set; } = new MarketPriceSettings();

        // 本地个人物料库与 SQLite 数据库配置选项分节
        [JsonPropertyName("PersonalComponentDbSettings")]
        public PersonalComponentDbSettings PersonalComponentDb { get; set; } = new PersonalComponentDbSettings();
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

        // 本地文件管理控制台 Web 页面访问基准地址 (默认嵌入独立路由) --硬编码--
        // 用于在 Excel 插件点击【本机项目】时唤起无外壳的纯净文件管理控制台
        public string LocalFileControlUrl { get; set; } = "https://code.xingren.online/#/embed/fileCtrol";

        // 远程 DrawCode 云端 WebAPI 服务的基准访问地址 (默认: https://code.xingren.online/api) --硬编码--
        // 用于登录认证、租户工作组获取及箱柜项目多租户数据同步
        public string DrawCodeBaseUrl { get; set; } = "https://code.xingren.online/api";
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

        // 箱柜明细块大标题与顶部预留行数 (默认: 3 行，即大标题、工程编号及隔断行)
        public int DetailHeaderRows { get; set; } = 3;

        // 箱柜明细块底部落款签名与商务附注预留行数 (默认: 3 行，即编制人、审核人/日期及落款签名)
        public int DetailFooterRows { get; set; } = 3;

        // 模板箱柜明细包含总计行的总行数 A (通过 CabTolsumRowIndex - CabDetRowIndex + 1 动态计算得出)
        [JsonIgnore]
        public int TemplateDetailTotalRows => CabTolsumRowIndex >= CabDetRowIndex ? (CabTolsumRowIndex - CabDetRowIndex + 1) : 28;

        // 模板箱柜明细整块（含大标题至总计行下方签名落款）的总行数 (动态结合总计行、箱柜行、顶部预留行与底部签名行)
        [JsonIgnore]
        public int TemplateDetailBlockTotalRows => (CabTolsumRowIndex >= CabDetRowIndex ? (CabTolsumRowIndex - CabDetRowIndex + 1) : 28) + DetailHeaderRows + DetailFooterRows;

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

    /// <summary>
    /// 一次图成套方案管理与 DWG 本地图纸目录配置实体
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class PrimaryCircuitSettings
    {
        // 一次成套方案 DWG 本地图纸目录绝对路径
        public string CircuitDwgDirectory { get; set; } = string.Empty;
    }

    /// <summary>
    /// 元器件图纸参数匹配窗口 (200x800) 配置实体
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class ComponentParamMatchSettings
    {
        // 元器件图纸库根物理目录路径 (默认留空，由系统智能自适应探测并可通过设置界面指定保存)
        public string BaseDirectory { get; set; } = string.Empty;

        // 目录名称写入的目标 Excel 列号名称 (普通分类表默认: AB 列即第 28 列，元件汇总表自动识别为 Y 列) --硬编码--
        public string TargetDirColumn { get; set; } = "AB";

        // 图纸名称写入的目标 Excel 列号名称 (普通分类表默认: AA 列即第 27 列，元件汇总表自动识别为 X 列) --硬编码--
        public string TargetDwgColumn { get; set; } = "AA";

        // 双击写入后是否自动跳转并选中下一行单元格 (默认: true)
        public bool AutoNextRow { get; set; } = true;

        // 双击写入图纸名时是否去除 .dwg 扩展名 (默认: true)
        public bool RemoveExtension { get; set; } = true;
    }

    /// <summary>
    /// AutoCAD 协同通信与离线图纸尺寸测算配置选项实体
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class CadSyncSettings
    {
        // 离线尺寸测算插件程序集相对或绝对路径 (默认: ClassLibrary3.dll)
        public string PluginDllPath { get; set; } = "ClassLibrary3.dll";

        // 自定义 accoreconsole.exe 绝对执行路径 (可选，留空时自适应探测系统安装目录)
        public string CustomAccoreConsolePath { get; set; } = string.Empty;
    }

    /// <summary>
    /// 工程项目与单体数据管理配置实体模型
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class ProjectSettings
    {
        // 候选的工程单体 JSON 搜索目录列表 (支持相对路径与绝对路径，可随时自定义增减)
        public System.Collections.Generic.List<string> ProjectJsonDirectories { get; set; } = new System.Collections.Generic.List<string>
        {
            @"E:\x项目文件\projectJsonFile"
        };
    }

    /// <summary>
    /// 在线元器件查价与多数据源平台配置实体模型
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class OnlinePriceSearchSettings
    {
        // 天工矩阵在线查价平台签名凭据与 API 端点分节
        public TitanMatrixSettings Titan { get; set; } = new TitanMatrixSettings();

        // 电气天下查价接口终结点完整 URL (默认: https://www.dq123.com/view/price/search/element)
        public string Dq123ApiUrl { get; set; } = "https://www.dq123.com/view/price/search/element";

        // 电气天下品牌与厂商 ID 动态映射字典 (支持用户在配置文件中无缝追加或覆盖新品牌)
        public System.Collections.Generic.Dictionary<string, int> Dq123BrandMap { get; set; } = new System.Collections.Generic.Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase)
        {
            // 正泰电器厂商 ID
            { "正泰", 1 },
            // 德力西电气厂商 ID
            { "德力西", 2 },
            // 常熟开关制造厂商 ID
            { "常熟", 3 },
            // 常熟开关别名
            { "常熟开关", 3 },
            // 施耐德电气厂商 ID
            { "施耐德", 4 },
            // ABB 电气厂商 ID
            { "ABB", 5 },
            // 西门子厂商 ID
            { "西门子", 6 },
            // 人民电器厂商 ID
            { "人民电器", 7 },
            // 天正电气厂商 ID
            { "天正", 8 },
            // 天正电气别名
            { "天正电气", 8 },
            // 良信电器厂商 ID
            { "良信", 138 }
        };
    }

    /// <summary>
    /// 天工矩阵 (titanmatrix.com) 平台 API 签名与接口端点配置实体
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class TitanMatrixSettings
    {
        // 天工矩阵 API 分配的签名应用 ID (默认: 201010)
        public string AppId { get; set; } = "201010";

        // 天工矩阵 API 分配的签名应用密钥
        public string AppKey { get; set; } = "72933362EAA649B893699E6191BC898F";

        // 天工矩阵查价接口终结点完整 URL
        public string ApiUrl { get; set; } = "https://cpq.titanmatrix.com/api/cpq/search/specsearch";

        // 防盗链与跨域请求来源 Referer (默认: http://www.titanmatrix.com/tgxx)
        public string Referer { get; set; } = "http://www.titanmatrix.com/tgxx";

        // 防盗链与跨域请求来源 Origin (默认: http://www.titanmatrix.com)
        public string Origin { get; set; } = "http://www.titanmatrix.com";
    }

    /// <summary>
    /// 有色金属实时大盘期货/现货行情与数据源配置实体
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class MarketPriceSettings
    {
        // 新浪等第三方行情 API 接口完整 URL
        public string ApiUrl { get; set; } = "https://hq.sinajs.cn/list=nf_CU0,nf_AL0";

        // 防盗链请求头 Referer 标头 (默认: https://finance.sina.com.cn)
        public string Referer { get; set; } = "https://finance.sina.com.cn";

        // 铜期货主力合约代码标识 (默认: nf_CU0)
        public string CopperContractCode { get; set; } = "nf_CU0";

        // 铝期货主力合约代码标识 (默认: nf_AL0)
        public string AluminumContractCode { get; set; } = "nf_AL0";

        // 铜期货转现货升水基差 (默认: 625.0 元/吨)
        public double CopperSpotPremium { get; set; } = 625.0;

        // 离线/断网状态下默认基准铜价 (默认: 110625.0 元/吨)
        public double DefaultCopperPrice { get; set; } = 110625.0;

        // 离线/断网状态下默认基准铝价 (默认: 23075.0 元/吨)
        public double DefaultAluminumPrice { get; set; } = 23075.0;

        // 行情内存缓存有效时间 (分钟，默认: 10 分钟)
        public int CacheDurationMinutes { get; set; } = 10;
    }

    /// <summary>
    /// 本地个人物料库 SQLite 数据库存储与连接池配置实体
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class PersonalComponentDbSettings
    {
        // 本地个人物料库 SQLite 数据库文件的自定义绝对路径 (若留空则由系统智能多级自愈探测)
        public string CustomDbFilePath { get; set; } = string.Empty;

        // 本地个人物料库数据库文件名称 (默认: personal_components.db)
        public string DbFileName { get; set; } = "personal_components.db";

        // SQLite 底层驱动是否启用连接池复用机制 (默认: true，显著提升高频读写吞吐量)
        public bool EnablePooling { get; set; } = true;

        // SQLite 连接池允许维持的最大物理连接数 (默认: 20)
        public int MaxPoolSize { get; set; } = 20;

        // SQLite 数据库执行命令与连接等待的超时时间 (单位：秒，默认: 30 秒)
        public int DefaultTimeoutSeconds { get; set; } = 30;
    }
}


