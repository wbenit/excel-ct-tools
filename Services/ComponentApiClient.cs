using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.Encodings.Web;
using System.Text.Json;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 商城电气元器件 WebAPI 客户端 (负责与 https://mall.xingren.online 真实接口通信)
    /// </summary>
    public static class ComponentApiClient
    {
        // 静态单例 HttpClient，避免套接字耗尽问题
        private static readonly HttpClient _httpClient;

        // JSON 反序列化全局配置选项
        private static readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions
        {
            // 忽略属性名称大小写差异
            PropertyNameCaseInsensitive = true,
            // 允许包含未转义字符
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        /// <summary>
        /// 静态构造函数：初始化网络通信协议与 HttpClient 实例
        /// </summary>
        static ComponentApiClient()
        {
            // 强制启用 TLS 1.2 与 TLS 1.3 现代安全加密传输协议
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

            // 初始化 HttpClientHandler 并配置连接生命周期
            var handler = new HttpClientHandler
            {
                // 自动解压缩 GZip / Deflate 响应流
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            // 创建全局复用的 HttpClient 实例
            _httpClient = new HttpClient(handler)
            {
                // 设置默认请求超时时间 (默认 15 秒)
                Timeout = TimeSpan.FromSeconds(15)
            };

            // 设置通用的标准浏览器 User-Agent 请求头，防止被网关防火墙拦截
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "ExcelAddInDemo/1.0 (Windows NT 10.0; Win64; x64)");
            // 设置接受的标准响应内容类型为 JSON
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
        }

        /// <summary>
        /// 从远程商城 WebAPI 获取所有品牌及其包含的元器件数量统计
        /// </summary>
        /// <returns>品牌及其数量统计列表</returns>
        public static List<BrandStatItemDto> GetBrandStats()
        {
            try
            {
                // 从配置读取 API 基准地址
                string baseUrl = ConfigManager.Instance.Current.Api.BaseUrl ?? "https://mall.xingren.online";
                string url = $"{baseUrl.TrimEnd('/')}/api/api/Component/GetBrandStats";

                // 同步发起 HTTP GET 请求
                var response = _httpClient.GetAsync(url).GetAwaiter().GetResult();
                if (!response.IsSuccessStatusCode)
                {
                    return new List<BrandStatItemDto>();
                }

                // 读取并反序列化响应内容
                string json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                var list = JsonSerializer.Deserialize<List<BrandStatItemDto>>(json, _jsonOptions);
                return list ?? new List<BrandStatItemDto>();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[ComponentApiClient] 获取品牌统计异常: {ex.Message}");
                return new List<BrandStatItemDto>();
            }
        }

        /// <summary>
        /// 综合检索与智能模糊查询核心方法 (支持根据已存在参数或模糊输入，严格受限于过滤管道)
        /// </summary>
        /// <param name="searchKeyword">用户即时输入的模糊搜索关键字 (如型号片段或名称)</param>
        /// <param name="name">元器件名称 (如: 微型断路器, 塑壳断路器, 双电源)</param>
        /// <param name="current">额定电流字符串 (如: 32, 100A)</param>
        /// <param name="pole">极数字符串 (如: 4P, 3, 1P+N)</param>
        /// <param name="tripMode">脱扣方式代号 (如: C, D, TM, MA)</param>
        /// <param name="brand">指定品牌筛选 (如: 施耐德, ABB，为空表示不限品牌)</param>
        /// <param name="mustContainRules">动态必含字段约束规则列表 (多条规则间为 AND 关系)</param>
        /// <param name="maxResults">最多返回结果条数限制 (默认 100)</param>
        /// <returns>经过全局过滤管道筛选后的匹配条目列表</returns>
        public static List<ComponentApiDto> SearchComponents(
            string? searchKeyword,
            string? name,
            string? current,
            string? pole,
            string? tripMode,
            string? brand = null,
            List<MustContainRule>? mustContainRules = null,
            int maxResults = 100)
        {
            try
            {
                // 从全局配置管理器读取 API 基准地址与接口端点
                string baseUrl = ConfigManager.Instance.Current.Api.BaseUrl ?? "https://mall.xingren.online";
                string endpoint = ConfigManager.Instance.Current.Api.ComponentGetPagedListEndpoint ?? "/api/api/Component/GetPagedList";

                // 构建查询参数 QueryString 列表
                var queryParams = new List<string>();

                // 1. 处理名称参数 (优先使用明确名称，若无则使用搜索关键字)
                string cleanName = name?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(cleanName))
                {
                    queryParams.Add($"Name={Uri.EscapeDataString(cleanName)}");
                }

                // 2. 处理电流参数 (提取出纯整型数字，如 "32A" ➔ 32)
                int? parsedCurrent = ExtractIntegerCurrent(current);
                if (parsedCurrent.HasValue)
                {
                    queryParams.Add($"Current={parsedCurrent.Value}");
                }

                // 3. 处理极数参数 (剥离末尾的 P 字符，如 "4P" ➔ "4")
                string cleanPoles = NormalizePolesParam(pole);
                if (!string.IsNullOrEmpty(cleanPoles))
                {
                    queryParams.Add($"Poles={Uri.EscapeDataString(cleanPoles)}");
                }

                // 4. 处理品牌筛选参数 (全局过滤管道第一层: 品牌)
                string cleanBrand = brand?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(cleanBrand) && !string.Equals(cleanBrand, "全部", StringComparison.OrdinalIgnoreCase) && !string.Equals(cleanBrand, "All", StringComparison.OrdinalIgnoreCase))
                {
                    queryParams.Add($"Brand={Uri.EscapeDataString(cleanBrand)}");
                }

                // 5. 设置大分页大小，获取足够多的候选项供客户端精细过滤
                queryParams.Add("PageIndex=1");
                queryParams.Add("PageSize=500");

                // 组合拼接完整的请求 URL 地址
                string queryString = string.Join("&", queryParams);
                string fullUrl = $"{baseUrl.TrimEnd('/')}/{endpoint.TrimStart('/')}?{queryString}";

                // 发送 HTTP GET 异步请求
                var response = _httpClient.GetAsync(fullUrl).GetAwaiter().GetResult();
                if (!response.IsSuccessStatusCode)
                {
                    LogHelper.WriteLog($"[ComponentApiClient] 请求返回状态码: {response.StatusCode}");
                    return new List<ComponentApiDto>();
                }

                // 读取响应内容
                string jsonString = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                if (string.IsNullOrWhiteSpace(jsonString))
                {
                    return new List<ComponentApiDto>();
                }

                // 反序列化为分页结果
                var pagedResult = JsonSerializer.Deserialize<ComponentPagedApiResponse<ComponentApiDto>>(jsonString, _jsonOptions);
                if (pagedResult == null || pagedResult.Items == null || pagedResult.Items.Count == 0)
                {
                    return new List<ComponentApiDto>();
                }

                var items = pagedResult.Items;

                // 6. 全局过滤管道第二层: 动态必含字段约束 (Must-Contain Rules, AND 关系)
                if (mustContainRules != null && mustContainRules.Count > 0)
                {
                    var activeRules = mustContainRules
                        .Where(r => r.Enabled && !string.IsNullOrWhiteSpace(r.Keyword))
                        .Select(r => r.Keyword.Trim())
                        .ToList();

                    if (activeRules.Count > 0)
                    {
                        items = items.Where(item =>
                        {
                            string fullModelText = (item.Model ?? string.Empty) + " " + (item.Param1 ?? string.Empty) + " " + (item.Param2 ?? string.Empty);
                            return activeRules.All(kw => fullModelText.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0);
                        }).ToList();
                    }
                }

                // 7. 用户输入的即时模糊搜索词过滤 (在已收敛的候选集中进一步模糊过滤)
                string kw = searchKeyword?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(kw))
                {
                    items = items.Where(it =>
                    {
                        string target = $"{it.Model} {it.Brand} {it.Name} {it.Remark} {it.Param1} {it.Param2}";
                        return target.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0;
                    }).ToList();
                }

                // 8. 若指定了脱扣特性，进一步精细筛选
                string normTrip = tripMode?.Trim().ToUpper() ?? string.Empty;
                if (!string.IsNullOrEmpty(normTrip) && items.Count > 1)
                {
                    var tripMatched = items.Where(x =>
                        string.Equals(x.Tripping?.Trim(), normTrip, StringComparison.OrdinalIgnoreCase) ||
                        (!string.IsNullOrEmpty(x.Model) && x.Model.ToUpper().Contains(normTrip))
                    ).ToList();

                    if (tripMatched.Count > 0)
                    {
                        items = tripMatched;
                    }
                }

                // 截取最大限制条数
                return items.Take(maxResults).ToList();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[ComponentApiClient] 搜索异常: {ex.Message}");
                return new List<ComponentApiDto>();
            }
        }

        /// <summary>
        /// 调用远程商城 WebAPI 根据【名称、电流、极数、脱扣、品牌、必含字段】多维组合检索元器件数据列表
        /// </summary>
        public static List<ComponentApiDto> QueryComponents(
            string name,
            string current,
            string pole,
            string tripMode,
            string? brand = null,
            List<MustContainRule>? mustContainRules = null)
        {
            return SearchComponents(null, name, current, pole, tripMode, brand, mustContainRules);
        }

        /// <summary>
        /// 兼容保留无品牌及必含规则的查询重载
        /// </summary>
        public static List<ComponentApiDto> QueryComponents(string name, string current, string pole, string tripMode)
        {
            return QueryComponents(name, current, pole, tripMode, null, null);
        }

        /// <summary>
        /// 从电流字符串中提取纯数字整型 (如 "32A", "100A", "32" ➔ 32)
        /// </summary>
        private static int? ExtractIntegerCurrent(string? currentStr)
        {
            if (string.IsNullOrWhiteSpace(currentStr)) return null;

            // 清洗字符串
            string clean = currentStr.Trim().ToUpper();
            if (clean.EndsWith("A"))
            {
                clean = clean.Substring(0, clean.Length - 1).Trim();
            }

            // 尝试直接解析整数
            if (int.TryParse(clean, out int curVal))
            {
                return curVal;
            }

            // 若包含浮点数则取整
            if (double.TryParse(clean, out double dVal))
            {
                return (int)Math.Round(dVal);
            }

            return null;
        }

        /// <summary>
        /// 规范化极数入参 (数据库中存储的是 "4", "3", "2", "1", "1+N" 等，剥离末尾的 P 字符)
        /// </summary>
        private static string NormalizePolesParam(string? poleStr)
        {
            if (string.IsNullOrWhiteSpace(poleStr)) return string.Empty;

            // 去除首尾空格并统一转为大写
            string clean = poleStr.Trim().ToUpper().Replace(" ", "");

            // 若末尾带 P (如 4P, 3P, 2P, 1P)，去除末尾的 P 适配数据库存储格式
            if (clean.EndsWith("P") && clean.Length > 1)
            {
                clean = clean.Substring(0, clean.Length - 1).Trim();
            }

            return clean;
        }
    }
}
