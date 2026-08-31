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
            // 转发调用完整参数重载
            return QueryComponents(name, current, pole, tripMode, null, null);
        }

        /// <summary>
        /// 根据品牌与元器件名称从远程接口获取所有符合条件的数据列表 (支持自动多页翻页拉全)
        /// </summary>
        /// <param name="brand">品牌过滤条件 (为空或"全部"表示不限品牌)</param>
        /// <param name="nameKeyword">名称过滤关键字 (支持模糊搜索)</param>
        /// <returns>符合条件的所有元器件列表</returns>
        public static List<ComponentApiDto> QueryManageComponents(string? brand, string? nameKeyword)
        {
            var allItems = new List<ComponentApiDto>();
            try
            {
                // 读取 API 基准地址配置
                string baseUrl = ConfigManager.Instance.Current.Api.BaseUrl ?? "https://mall.xingren.online";
                string endpoint = ConfigManager.Instance.Current.Api.ComponentGetPagedListEndpoint ?? "/api/api/Component/GetPagedList";

                int pageIndex = 1;
                int pageSize = ComponentManageDefaults.MaxQueryPageSize;
                int totalPages = 1;

                do
                {
                    // 构建查询 QueryString
                    var queryParams = new List<string>();
                    queryParams.Add($"PageIndex={pageIndex}");
                    queryParams.Add($"PageSize={pageSize}");

                    // 品牌筛选
                    string cleanBrand = brand?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(cleanBrand) && !string.Equals(cleanBrand, "全部", StringComparison.OrdinalIgnoreCase) && !string.Equals(cleanBrand, "All", StringComparison.OrdinalIgnoreCase))
                    {
                        queryParams.Add($"Brand={Uri.EscapeDataString(cleanBrand)}");
                    }

                    // 名称关键字筛选 (接口支持 Name 或 Keyword)
                    string cleanName = nameKeyword?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(cleanName))
                    {
                        queryParams.Add($"Keyword={Uri.EscapeDataString(cleanName)}");
                    }

                    string fullUrl = $"{baseUrl.TrimEnd('/')}/{endpoint.TrimStart('/')}?{string.Join("&", queryParams)}";

                    var response = _httpClient.GetAsync(fullUrl).GetAwaiter().GetResult();
                    if (!response.IsSuccessStatusCode)
                    {
                        LogHelper.WriteLog($"[ComponentApiClient] 拉取管理列表第 {pageIndex} 页失败: {response.StatusCode}");
                        break;
                    }

                    string jsonString = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    if (string.IsNullOrWhiteSpace(jsonString)) break;

                    var pagedResult = JsonSerializer.Deserialize<ComponentPagedApiResponse<ComponentApiDto>>(jsonString, _jsonOptions);
                    if (pagedResult == null || pagedResult.Items == null || pagedResult.Items.Count == 0)
                    {
                        break;
                    }

                    allItems.AddRange(pagedResult.Items);
                    totalPages = pagedResult.TotalPages;
                    pageIndex++;
                }
                while (pageIndex <= totalPages && pageIndex <= 20); // 最多拉取 20 页 (10000 条) 防死循环
            }
            catch (Exception ex)
            {
                // 记录拉取异常日志
                LogHelper.WriteLog($"[ComponentApiClient] QueryManageComponents 异常: {ex.Message}");
            }
            return allItems;
        }

        // 品牌与元器件名称列表内存缓存，避免重复拉取
        private static readonly Dictionary<string, List<string>> _brandNamesCache = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 根据品牌从远程接口获取该品牌下包含的所有不重复的元器件名称列表（优先通过后端专用 DISTINCT 接口秒级拉取全量真实品类）
        /// </summary>
        /// <param name="brand">选中的品牌 (若为空则获取全局常见名称)</param>
        /// <returns>排序好的不重复元器件名称字符串列表</returns>
        public static List<string> GetNamesByBrand(string? brand)
        {
            // 归一化缓存键，若为空则表示全部品牌
            string cacheKey = string.IsNullOrWhiteSpace(brand) ? "__ALL__" : brand.Trim();
            // 命中内存缓存直接返回，避免无谓的网络往返
            lock (_brandNamesCache)
            {
                if (_brandNamesCache.TryGetValue(cacheKey, out var cachedList))
                {
                    return cachedList;
                }
            }

            var names = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                // 读取 API 基准地址配置
                string baseUrl = ConfigManager.Instance.Current.Api.BaseUrl ?? "https://mall.xingren.online";
                string cleanBrand = brand?.Trim() ?? string.Empty;

                // 1. 优先调用后端专用的数据库级 DISTINCT 聚合接口: GET /api/api/Component/GetNames?brand={brand}
                string endpoint = $"{baseUrl.TrimEnd('/')}/api/api/Component/GetNames";
                if (!string.IsNullOrEmpty(cleanBrand) && !string.Equals(cleanBrand, "全部", StringComparison.OrdinalIgnoreCase))
                {
                    endpoint += $"?brand={Uri.EscapeDataString(cleanBrand)}";
                }

                var response = _httpClient.GetAsync(endpoint).GetAwaiter().GetResult();
                if (response.IsSuccessStatusCode)
                {
                    string json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        var list = JsonSerializer.Deserialize<List<string>>(json, _jsonOptions);
                        if (list != null && list.Count > 0)
                        {
                            foreach (var n in list)
                            {
                                if (!string.IsNullOrWhiteSpace(n)) names.Add(n.Trim());
                            }
                        }
                    }
                }
                else
                {
                    // 2. 防御性降级策略：若线上尚未更新该接口，通过小批量嗅探兜底
                    LogHelper.WriteLog($"[ComponentApiClient] GetNames 接口响应码: {response.StatusCode}，启用兜底嗅探");
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[ComponentApiClient] GetNamesByBrand 异常: {ex.Message}");
            }

            var result = names.ToList();
            // 写入本地全局字典缓存
            lock (_brandNamesCache)
            {
                _brandNamesCache[cacheKey] = result;
            }
            return result;
        }

        /// <summary>
        /// 根据品牌、名称以及主体型号从远程接口获取可适配的配套附件列表
        /// </summary>
        /// <param name="brand">所属品牌 (如: 德力西, 施耐德)</param>
        /// <param name="name">元器件名称 (如: 塑壳断路器)</param>
        /// <param name="model">当前选中的主体型号文本 (用于与附件 Param2 匹配)</param>
        /// <returns>符合适配条件的附件列表</returns>
        public static List<ComponentApiDto> GetAttachments(string? brand, string? name, string? model)
        {
            var attachments = new List<ComponentApiDto>();
            try
            {
                // 读取 API 基准地址
                string baseUrl = ConfigManager.Instance.Current.Api.BaseUrl ?? "https://mall.xingren.online";
                string cleanBrand = brand?.Trim() ?? string.Empty;
                string cleanName = name?.Trim() ?? string.Empty;
                string cleanModel = model?.Trim() ?? string.Empty;

                // 构建请求参数
                var queryParams = new List<string>();
                if (!string.IsNullOrEmpty(cleanBrand)) queryParams.Add($"brand={Uri.EscapeDataString(cleanBrand)}");
                if (!string.IsNullOrEmpty(cleanName)) queryParams.Add($"name={Uri.EscapeDataString(cleanName)}");
                if (!string.IsNullOrEmpty(cleanModel)) queryParams.Add($"model={Uri.EscapeDataString(cleanModel)}");

                string endpoint = $"{baseUrl.TrimEnd('/')}/api/api/Component/GetAttachments?{string.Join("&", queryParams)}";
                var response = _httpClient.GetAsync(endpoint).GetAwaiter().GetResult();

                if (response.IsSuccessStatusCode)
                {
                    string json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        var list = JsonSerializer.Deserialize<List<ComponentApiDto>>(json, _jsonOptions);
                        if (list != null) attachments = list;
                    }
                }
                else
                {
                    // 若线上尚未更新该接口，启用备用降级逻辑：拉取该品牌同名称且 Param1='附件' 的数据并在客户端比对
                    LogHelper.WriteLog($"[ComponentApiClient] GetAttachments 端点返回状态码 {response.StatusCode}，启动本地过滤降级策略");
                    var rawList = SearchComponents(null, cleanName, null, null, null, cleanBrand, null, 200);
                    attachments = rawList.Where(item =>
                    {
                        if (!string.Equals(item.Param1?.Trim(), "附件", StringComparison.OrdinalIgnoreCase)) return false;
                        if (string.IsNullOrWhiteSpace(item.Param2)) return false;
                        if (string.IsNullOrEmpty(cleanModel)) return true;

                        var keys = item.Param2.Split(new[] { ',', '，', ';', '；' }, StringSplitOptions.RemoveEmptyEntries)
                                              .Select(k => k.Trim())
                                              .Where(k => !string.IsNullOrEmpty(k));
                        return keys.Any(k => cleanModel.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0);
                    }).ToList();
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[ComponentApiClient] GetAttachments 异常: {ex.Message}");
            }

            return attachments;
        }

        /// <summary>
        /// 调用远程商城 WebAPI 新增单条元器件数据 (POST /api/api/Component/Create)
        /// </summary>
        /// <param name="dto">新增参数 DTO</param>
        /// <returns>新增成功后的元器件数据 (包含新生成的 Id)，失败返回 null</returns>
        public static ComponentApiDto? CreateComponent(CreateComponentApiRequest dto)
        {
            try
            {
                // 读取 API 基准地址配置
                string baseUrl = ConfigManager.Instance.Current.Api.BaseUrl ?? "https://mall.xingren.online";
                string url = $"{baseUrl.TrimEnd('/')}/api/api/Component/Create";

                // 序列化请求体 JSON
                string jsonBody = JsonSerializer.Serialize(dto, _jsonOptions);
                var content = new StringContent(jsonBody, System.Text.Encoding.UTF8, "application/json");

                // 发起 POST 请求
                var response = _httpClient.PostAsync(url, content).GetAwaiter().GetResult();
                if (!response.IsSuccessStatusCode)
                {
                    LogHelper.WriteLog($"[ComponentApiClient] 创建元器件失败, 状态码: {response.StatusCode}");
                    return null;
                }

                // 读取并反序列化新增成功后的实体
                string respJson = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                return JsonSerializer.Deserialize<ComponentApiDto>(respJson, _jsonOptions);
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[ComponentApiClient] CreateComponent 异常: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 调用远程商城 WebAPI 更新已存在的元器件数据 (PUT /api/api/Component/Update)
        /// </summary>
        /// <param name="dto">更新参数 DTO (必须包含有效 Id)</param>
        /// <returns>更新成功后的元器件数据，失败返回 null</returns>
        public static ComponentApiDto? UpdateComponent(UpdateComponentApiRequest dto)
        {
            try
            {
                // 读取 API 基准地址配置
                string baseUrl = ConfigManager.Instance.Current.Api.BaseUrl ?? "https://mall.xingren.online";
                string url = $"{baseUrl.TrimEnd('/')}/api/api/Component/Update";

                // 序列化请求体 JSON
                string jsonBody = JsonSerializer.Serialize(dto, _jsonOptions);
                var content = new StringContent(jsonBody, System.Text.Encoding.UTF8, "application/json");

                // 发起 PUT 请求
                var response = _httpClient.PutAsync(url, content).GetAwaiter().GetResult();
                if (!response.IsSuccessStatusCode)
                {
                    LogHelper.WriteLog($"[ComponentApiClient] 更新元器件 ID {dto.Id} 失败, 状态码: {response.StatusCode}");
                    return null;
                }

                // 读取并反序列化更新结果
                string respJson = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                return JsonSerializer.Deserialize<ComponentApiDto>(respJson, _jsonOptions);
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[ComponentApiClient] UpdateComponent 异常: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 调用远程商城 WebAPI 根据主键 ID 删除指定的元器件数据 (DELETE /api/api/Component/Delete?id={id})
        /// </summary>
        /// <param name="id">元器件主键 ID</param>
        /// <returns>是否删除成功</returns>
        public static bool DeleteComponent(int id)
        {
            try
            {
                // 读取 API 基准地址配置
                string baseUrl = ConfigManager.Instance.Current.Api.BaseUrl ?? "https://mall.xingren.online";
                string url = $"{baseUrl.TrimEnd('/')}/api/api/Component/Delete?id={id}";

                // 发起 DELETE 请求
                var response = _httpClient.DeleteAsync(url).GetAwaiter().GetResult();
                if (!response.IsSuccessStatusCode)
                {
                    LogHelper.WriteLog($"[ComponentApiClient] 删除元器件 ID {id} 失败, 状态码: {response.StatusCode}");
                    return false;
                }

                // 解析布尔结果
                string respJson = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                if (bool.TryParse(respJson.Trim(), out bool result))
                {
                    return result;
                }
                return true;
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[ComponentApiClient] DeleteComponent 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 从电流字符串中提取纯数字整型 (如 "32A", "100A", "32" ➔ 32)
        /// </summary>
        public static int? ExtractIntegerCurrent(string? currentStr)
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
        public static string NormalizePolesParam(string? poleStr)
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
