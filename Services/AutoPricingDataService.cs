using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// 自动组价方案数据服务 (纯云端 WebAPI 远程数据通道)
    /// </summary>
    public static class AutoPricingDataService
    {
        // 通用 HttpClient 实例 (全局复用单例以提高网络吞吐率)
        private static readonly HttpClient _httpClient = CreateHttpClient();

        // 挂载宽松反序列化转换器的 JSON 序列化选项实例
        private static readonly JsonSerializerOptions JsonOptions = CreateJsonOptions();

        // 创建带超时控制的 HTTP 客户端
        private static HttpClient CreateHttpClient()
        {
            // 实例化原生 HttpClient
            var client = new HttpClient();
            // 设置网络请求超时等待时间为 5 秒，平衡公网网络波动与离线快速降级
            client.Timeout = TimeSpan.FromSeconds(5);
            return client;
        }

        // 构建通用 JSON 序列化选项并挂载宽容字符串转换器
        private static JsonSerializerOptions CreateJsonOptions()
        {
            // 实例化基础序列化配置
            var opt = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                PropertyNameCaseInsensitive = true,
                WriteIndented = false
            };
            // 挂载宽容字符串转换器，彻底解决后端整型 ID 与前端字符串 DTO 反序列化冲突
            opt.Converters.Add(new FlexibleStringConverter());
            return opt;
        }

        /// <summary>
        /// 1. 获取完整的方案分类树 (纯云端 WebAPI 远程数据通道)
        /// </summary>
        public static List<AutoPricingCategoryDto> GetCategoryTree()
        {
            try
            {
                // 优先尝试从云端 WebAPI 拉取最新树形数据
                string url = BuildSchemeApiUrl("/api/Scheme/GetCategoryTree");
                var task = _httpClient.GetStringAsync(url);
                // 等待 4 秒超时判定
                if (task.Wait(4000))
                {
                    string json = task.Result;
                    // 使用挂载了 FlexibleStringConverter 的 JSON 选项解析
                    var tree = JsonSerializer.Deserialize<List<AutoPricingCategoryDto>>(json, JsonOptions);
                    if (tree != null && tree.Count > 0)
                    {
                        return tree;
                    }
                }
            }
            catch (Exception exTree)
            {
                // 网络或服务未就绪，记录日志后安全返回空列表
                LogHelper.WriteLog($"[AutoPricingDataService] 云端 GetCategoryTree 异常: {exTree.Message}");
            }

            // 云端请求失败或无数据时直接返回空列表 (不使用本地文件)
            return new List<AutoPricingCategoryDto>();
        }

        /// <summary>
        /// 1.1 按需懒加载获取子节点 (支持 el-tree lazy 模式，仅拉取当前层级，节省网络流量)
        /// </summary>
        public static List<AutoPricingCategoryDto> GetCategoryNodes(string? parentId)
        {
            // 根节点为空或 0 时规范化为 "0"
            string pId = string.IsNullOrWhiteSpace(parentId) ? "0" : parentId!.Trim();
            try
            {
                // 优先尝试从云端 WebAPI 按需拉取当前层级子节点
                string url = BuildSchemeApiUrl($"/api/Scheme/GetCategoryNodes?parentId={Uri.EscapeDataString(pId)}");
                var task = _httpClient.GetStringAsync(url);
                // 等待 4 秒超时判定
                if (task.Wait(4000))
                {
                    string json = task.Result;
                    // 使用宽松配置反序列化子节点集合
                    var nodes = JsonSerializer.Deserialize<List<AutoPricingCategoryDto>>(json, JsonOptions);
                    if (nodes != null && nodes.Count > 0)
                    {
                        return nodes;
                    }
                }
            }
            catch (Exception exNodes)
            {
                // 云端超时或未连接，记录日志并安全返回空列表
                LogHelper.WriteLog($"[AutoPricingDataService] 云端 GetCategoryNodes 异常: {exNodes.Message}");
            }

            // 云端请求失败时直接返回空列表 (不使用本地文件)
            return new List<AutoPricingCategoryDto>();
        }

        /// <summary>
        /// 2. 分页或按条件查询方案列表 (纯云端 WebAPI 远程数据通道)
        /// </summary>
        public static List<AutoPricingSchemeDto> GetSchemes(string? categoryId, string? keyword, string? cabModel, string? situation)
        {
            try
            {
                // 优先请求云端接口，组装 URL 参数并执行合法编码
                var queryParams = new List<string>();
                // 拼接分类 ID 过滤参数
                if (!string.IsNullOrWhiteSpace(categoryId)) queryParams.Add($"categoryId={Uri.EscapeDataString(categoryId!.Trim())}");
                // 拼接关键字过滤参数
                if (!string.IsNullOrWhiteSpace(keyword)) queryParams.Add($"keyword={Uri.EscapeDataString(keyword!.Trim())}");
                // 拼接柜型过滤参数
                if (!string.IsNullOrWhiteSpace(cabModel)) queryParams.Add($"cabModel={Uri.EscapeDataString(cabModel!.Trim())}");
                // 拼接适用场景过滤参数
                if (!string.IsNullOrWhiteSpace(situation)) queryParams.Add($"situation={Uri.EscapeDataString(situation!.Trim())}");
                queryParams.Add("pageSize=100");
                // 拼接查询字符串
                string qs = string.Join("&", queryParams);
                string url = BuildSchemeApiUrl($"/api/Scheme/GetPagedSchemes?{qs}");
                var task = _httpClient.GetStringAsync(url);
                // 等待 4 秒判定
                if (task.Wait(4000))
                {
                    string json = task.Result;
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("items", out var itemsEl))
                    {
                        // 提取方案列表数据
                        var list = JsonSerializer.Deserialize<List<AutoPricingSchemeDto>>(itemsEl.GetRawText(), JsonOptions);
                        if (list != null) return list;
                    }
                }
            }
            catch (Exception exSchemes)
            {
                // 记录云端方案列表请求异常日志
                LogHelper.WriteLog($"[AutoPricingDataService] 云端 GetSchemes 异常: {exSchemes.Message}");
            }

            // 云端未取到时安全返回空列表 (不使用本地文件)
            return new List<AutoPricingSchemeDto>();
        }

        /// <summary>
        /// 3. 获取单个方案详情及完整 BOM 清单 (纯云端 WebAPI 远程数据通道)
        /// </summary>
        public static AutoPricingSchemeDetailDto? GetSchemeDetail(string schemeId)
        {
            try
            {
                // 校验待查询的方案 ID
                if (!string.IsNullOrWhiteSpace(schemeId))
                {
                    // 优先云端查询方案与 BOM 明细
                    string url = BuildSchemeApiUrl($"/api/Scheme/GetSchemeDetail?schemeId={Uri.EscapeDataString(schemeId.Trim())}");
                    var task = _httpClient.GetStringAsync(url);
                    // 等待 4 秒判定
                    if (task.Wait(4000))
                    {
                        string json = task.Result;
                        // 宽松反序列化方案详情及关联 BOM
                        var detail = JsonSerializer.Deserialize<AutoPricingSchemeDetailDto>(json, JsonOptions);
                        if (detail != null && detail.BomItems != null && detail.BomItems.Count > 0)
                        {
                            return detail;
                        }
                    }
                }
            }
            catch (Exception exDetail)
            {
                // 记录云端详情请求异常日志
                LogHelper.WriteLog($"[AutoPricingDataService] 云端 GetSchemeDetail 异常: {exDetail.Message}");
            }

            // 云端请求失败或无数据时返回 null (不使用本地文件)
            return null;
        }

        /// <summary>
        /// 统一组装方案中心远程 WebAPI 请求完整绝对 URL
        /// 优先读取 ConfigManager 全局配置 BaseUrl，并自动适配公网 /api/api/Scheme 网关路由
        /// </summary>
        private static string BuildSchemeApiUrl(string relativePath)
        {
            // 通过 ComponentApiClient 统一自愈提取远程商城服务的 BaseUrl
            string baseUrl = ComponentApiClient.GetBaseUrl();
            // 规范化相对接口路径，统一剔除前导斜杠
            string path = (relativePath ?? string.Empty).TrimStart('/');

            // 若基础地址为公网生产环境 (包含 mall.xingren.online)，统一补齐网关必需的 /api 前缀
            if (baseUrl.IndexOf("mall.xingren.online", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // 确保对齐网关路径为 /api/api/Scheme/...
                if (path.StartsWith("api/Scheme", StringComparison.OrdinalIgnoreCase))
                {
                    // 补齐第一层网关 api 路由
                    path = "api/" + path;
                }
            }

            // 拼接返回完整可用的网络请求 URL
            return $"{baseUrl}/{path}";
        }
    }

    /// <summary>
    /// 自动组价分类树 DTO
    /// </summary>
    public class AutoPricingCategoryDto
    {
        // 节点唯一标识 ID (挂载宽容转换器兼容后端整型 ID)
        [System.Text.Json.Serialization.JsonConverter(typeof(FlexibleStringConverter))]
        public string Id { get; set; } = string.Empty;

        // 父级分类标识 ID (挂载宽容转换器兼容数字 0 与整型 ID)
        [System.Text.Json.Serialization.JsonConverter(typeof(FlexibleStringConverter))]
        public string ParentId { get; set; } = string.Empty;

        // 节点显示名称
        public string Name { get; set; } = string.Empty;

        // 别名：分类名称 (兼容前端 categoryName 取值)
        [System.Text.Json.Serialization.JsonPropertyName("categoryName")]
        public string CategoryName => Name;

        // 目录层级级别
        public int Level { get; set; }

        // 分类面包屑完整层级路径
        public string FullPath { get; set; } = string.Empty;

        // 当前分类下方案总数统计
        public int SchemeCount { get; set; }

        // 节点类型标识 (true: 分类目录, false: 方案叶子节点)
        public bool IsCategory { get; set; } = true;

        // 节点是否为叶子节点 (用于前端 el-tree lazy 懒加载标识，方案为 true，目录为 false)
        [System.Text.Json.Serialization.JsonPropertyName("isLeaf")]
        public bool IsLeaf { get; set; } = false;

        // 方案 ID (挂载宽容转换器兼容数字与空字符串)
        [System.Text.Json.Serialization.JsonConverter(typeof(FlexibleStringConverter))]
        public string SchemeId { get; set; } = string.Empty;

        // 方案对应柜型型号
        public string CabModel { get; set; } = string.Empty;

        // 方案图号代号
        public string Model { get; set; } = string.Empty;

        // 参考方案总价
        public decimal TotalPrice { get; set; }

        // BOM 元器件条目数
        public int BomCount { get; set; }

        // 下级子节点集合
        public List<AutoPricingCategoryDto> Children { get; set; } = new List<AutoPricingCategoryDto>();
    }

    /// <summary>
    /// 自动组价方案简要 DTO
    /// </summary>
    public class AutoPricingSchemeDto
    {
        // 方案唯一主键 ID (挂载宽容转换器兼容后端自增整型)
        [System.Text.Json.Serialization.JsonConverter(typeof(FlexibleStringConverter))]
        public string Id { get; set; } = string.Empty;

        // 所属分类主键 ID (挂载宽容转换器)
        [System.Text.Json.Serialization.JsonConverter(typeof(FlexibleStringConverter))]
        public string CategoryId { get; set; } = string.Empty;

        // 所属分类全路径
        public string CategoryPath { get; set; } = string.Empty;

        // 方案显示名称
        public string Name { get; set; } = string.Empty;

        // 别名：方案名称 (兼容前端 schemeName 取值)
        [System.Text.Json.Serialization.JsonPropertyName("schemeName")]
        public string SchemeName => Name;

        // 开关柜型型号
        public string CabModel { get; set; } = string.Empty;

        // 别名：开关柜型 (兼容前端 cabinetModel 取值)
        [System.Text.Json.Serialization.JsonPropertyName("cabinetModel")]
        public string CabinetModel => CabModel;

        // 方案代号编号
        public string Model { get; set; } = string.Empty;

        // 别名：方案代号/编号 (兼容前端 schemeCode 取值)
        [System.Text.Json.Serialization.JsonPropertyName("schemeCode")]
        public string SchemeCode => Model;

        // 柜体外形尺寸
        public string Dimensions { get; set; } = string.Empty;

        // 使用应用场景
        public string Situation { get; set; } = string.Empty;

        // 主回路额定规格参数
        public string MainSpec { get; set; } = string.Empty;

        // 别名：额定电流规格 (兼容前端 ratedCurrent 取值)
        [System.Text.Json.Serialization.JsonPropertyName("ratedCurrent")]
        public string RatedCurrent => MainSpec;

        // 修改人
        public string Modifier { get; set; } = string.Empty;

        // 修改时间
        public string ModifyTime { get; set; } = string.Empty;

        // 方案考量参考总价
        public decimal TotalPrice { get; set; }

        // BOM 元器件条目数
        public int BomCount { get; set; }
    }

    /// <summary>
    /// 自动组价方案详情 DTO
    /// </summary>
    public class AutoPricingSchemeDetailDto : AutoPricingSchemeDto
    {
        // 关联的全部 BOM 元器件明细清单
        public List<CloudSchemeBomItem> BomItems { get; set; } = new List<CloudSchemeBomItem>();
    }
}
