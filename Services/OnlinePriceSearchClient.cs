using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// 在线查价网络客户端服务，纯原生集成电气天下与天工矩阵 API
    /// </summary>
    public class OnlinePriceSearchClient
    {
        // 全局单例 HttpClient 实例，复用底层 TCP 连接池
        private static readonly HttpClient _httpClient;

        // JSON 反序列化宽松配置选项
        private static readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// 静态构造函数: 预置安全协议与 HttpClient 基础配置
        /// </summary>
        static OnlinePriceSearchClient()
        {
            // 确保启用现代 TLS 1.2 与 TLS 1.3 传输加密协议
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

            // 初始化套接字处理器
            var handler = new HttpClientHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            // 创建带 15 秒超时的客户端实例 --硬编码: 网络超时时间--
            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(15)
            };

            // 预装通用的浏览器 User-Agent 标头
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
        }

        #region 品牌厂商精准识别与归一化引擎

        // 电气天下品牌 ID 映射表 --硬编码: 厂商代码对照--
        private static readonly Dictionary<string, int> Dq123BrandMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            // 正泰厂家代码
            { "正泰", 1 },
            // 德力西厂家代码
            { "德力西", 2 },
            // 常熟开关厂家代码
            { "常熟", 3 },
            // 常熟开关别名
            { "常熟开关", 3 },
            // 施耐德厂家代码
            { "施耐德", 4 },
            // ABB 厂家代码
            { "ABB", 5 },
            // 西门子厂家代码
            { "西门子", 6 },
            // 人民电器厂家代码
            { "人民电器", 7 },
            // 天正电气厂家代码
            { "天正", 8 },
            // 天正电气别名
            { "天正电气", 8 },
            // 良信电器厂家代码
            { "良信", 138 }
        };

        // 电气天下已知厂商代码与实际品牌反向对照表 --硬编码: 厂商代码到实际品牌--
        private static readonly Dictionary<int, string> Dq123FactoryIdMap = new Dictionary<int, string>
        {
            // 正泰电器官方 ID 1
            { 1, "正泰" },
            // 正泰温州电器/系列 ID 113
            { 113, "正泰" },
            // 正泰微断系列 ID 375
            { 375, "正泰" },
            // 正泰 DZ47 系列 ID 521
            { 521, "正泰" },
            // 德力西官方 ID 2
            { 2, "德力西" },
            // 德力西电气系列 ID 51
            { 51, "德力西" },
            // 德力西配电系列 ID 536
            { 536, "德力西" },
            // 常熟开关官方 ID 3
            { 3, "常熟开关" },
            // 施耐德官方 ID 4
            { 4, "施耐德" },
            // ABB 官方 ID 5
            { 5, "ABB" },
            // 西门子官方 ID 6
            { 6, "西门子" },
            // 人民电器官方 ID 7
            { 7, "人民电器" },
            // 天正电气官方 ID 8
            { 8, "天正" },
            // 良信电器官方 ID 138
            { 138, "良信" }
        };

        // 知名电气品牌关键字列表 (按优先级降序排列) --硬编码: 知名电气品牌关键词--
        private static readonly (string keyword, string normalizedBrand)[] WellKnownBrandKeywords = new[]
        {
            // 正泰品牌
            ( "正泰", "正泰" ),
            // 德力西品牌
            ( "德力西", "德力西" ),
            // 常熟开关全称
            ( "常熟开关", "常熟开关" ),
            // 常熟简称
            ( "常熟", "常熟开关" ),
            // 施耐德品牌
            ( "施耐德", "施耐德" ),
            // 良信品牌
            ( "良信", "良信" ),
            // ABB 品牌
            ( "ABB", "ABB" ),
            // 西门子品牌
            ( "西门子", "西门子" ),
            // 西门子英文
            ( "SIEMENS", "西门子" ),
            // 施耐德英文
            ( "SCHNEIDER", "施耐德" ),
            // 德力西英文
            ( "DELIXI", "德力西" ),
            // 正泰英文
            ( "CHINT", "正泰" ),
            // 良信英文
            ( "NADER", "良信" ),
            // 人民电器品牌
            ( "人民电器", "人民电器" ),
            // 上海人民电气
            ( "上海人民", "上海人民" ),
            // 天正电气
            ( "天正", "天正" ),
            // 罗格朗电气
            ( "罗格朗", "罗格朗" ),
            // 伊顿电气
            ( "伊顿", "伊顿" ),
            // 伊顿英文
            ( "EATON", "伊顿" ),
            // 三菱电机
            ( "三菱", "三菱" ),
            // 三菱英文
            ( "MITSUBISHI", "三菱" ),
            // 富士电机
            ( "富士", "富士" ),
            // 富士英文
            ( "FUJI", "富士" ),
            // 海格电气
            ( "海格", "海格" ),
            // 海格英文
            ( "HAGER", "海格" ),
            // 天水二一三
            ( "天水二一三", "天水二一三" ),
            // 长城电器
            ( "长城电器", "长城电器" ),
            // 长城简称
            ( "长城", "长城电器" )
        };

        /// <summary>
        /// 从元器件型号特征前缀智能推断实际品牌 (符合电气成套行业通用命名惯例)
        /// </summary>
        /// <param name="model">元器件型号文本</param>
        /// <returns>推断得到的实际品牌名称</returns>
        public static string DeduceBrandFromModel(string model)
        {
            // 空字符串保护
            if (string.IsNullOrWhiteSpace(model)) return string.Empty;
            // 清除内部空白并转换为大写比对
            string clean = Regex.Replace(model, @"\s+", "").ToUpperInvariant();

            // 常熟开关专属型号前缀 (塑壳断路器CM系列/框架断路器CW系列/隔离开关CA系列/交流接触器CJ40)
            if (Regex.IsMatch(clean, @"^(CM[1-5]|CW[1-3]|CA[1-2]|CJ40|CH1)")) return "常熟开关";

            // 施耐德专属型号前缀 (微断iC65/C65/C120, 塑壳NSX/CVS/EZD, 框架MT/NW, 电动机保护GV2/GV3, 接触器LC1)
            if (Regex.IsMatch(clean, @"^(IC65|C65|NSX|C120|EZD|CVS|MT|NW|NS|GV2|GV3|LC1|TESYS|ACTI9)")) return "施耐德";

            // 良信专属型号前缀 (塑壳NDM系列/微断NDB系列/框架NDW系列)
            if (Regex.IsMatch(clean, @"^(NDM[1-5]|NDB[1-3]|NDW[1-3]|NDC[1-3]|NDG[1-3])")) return "良信";

            // ABB 专属型号前缀 (微断S200系列/塑壳XT系列/框架Emax系列/Tmax系列/接触器A9/AF系列)
            if (Regex.IsMatch(clean, @"^(S20|S28|XT[1-4]|EMAX|TMAX|A9|A16|A26|AF[0-9]|OT[0-9]|MS1[0-9])")) return "ABB";

            // 西门子专属型号前缀 (微断5SY/5SJ系列, 塑壳3VM/3VA系列, 框架3WL/3WT系列, 接触器3TF/3RT系列)
            if (Regex.IsMatch(clean, @"^(5SY|5SJ|5SP|3VM|3VA|3WL|3WT|3TF|3RT|3RV)")) return "西门子";

            // 天正电气专属型号前缀 (微断TGB1N/TGBG, 塑壳TGM系列, 框架TGW系列)
            if (Regex.IsMatch(clean, @"^(TGB|TGM|TGW|TGQ|CJX2S)")) return "天正";

            // 人民电器专属型号前缀 (微断RMC系列, 塑壳RDM系列, 框架RDW系列)
            if (Regex.IsMatch(clean, @"^(RMC|RDM|RDW|RDQ)")) return "人民电器";

            // 德力西专属型号前缀 (微断CDB6/CDB9, 塑壳CDM1/CDM3, 框架CDW9)
            if (Regex.IsMatch(clean, @"^(CDB[6-9]|CDM[1-3]|CDW[1-9]|CDZ[1-9]|CDC[1-9]|CDL[1-9]|CDQ[1-9])")) return "德力西";

            // 正泰专属型号前缀 (微断NB1, 塑壳NM1/NM8/NXM系列, 接触器CJX2, 热继NR2)
            if (Regex.IsMatch(clean, @"^(NB[1-9]|NM[1-8]|NXM|NXBLE|NM8N|DZ158|NR2|NVF|NQQ)")) return "正泰";

            // 未命中明确前缀返回空
            return string.Empty;
        }

        /// <summary>
        /// 从天工矩阵返回的工商公司全称中智能提炼核心实际品牌名称
        /// </summary>
        /// <param name="rawVendor">天工矩阵原始厂商名称</param>
        /// <param name="model">检索型号</param>
        /// <param name="selectedBrand">用户选定的品牌</param>
        /// <returns>归一化后的实际品牌名称</returns>
        public static string ResolveBrandForTitan(string rawVendor, string model, string selectedBrand)
        {
            // 若用户在界面指定了明确品牌(且不是全部)，优先校验
            if (!string.IsNullOrWhiteSpace(selectedBrand) && selectedBrand != "全部")
            {
                // 若原始厂商为空或包含所选品牌，直接返回用户选择
                if (string.IsNullOrWhiteSpace(rawVendor) || rawVendor.IndexOf(selectedBrand, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return selectedBrand;
                }
            }

            // 若接口返回了有效厂商公司名
            if (!string.IsNullOrWhiteSpace(rawVendor))
            {
                // 优先在知名品牌关键词表中匹配
                foreach (var (kw, norm) in WellKnownBrandKeywords)
                {
                    // 若公司名包含关键词，直接返回规范化简明品牌
                    if (rawVendor.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return norm;
                    }
                }

                // 若未在已知大厂表中，智能剥离工商后缀与行政前缀提取核心字号
                string cleaned = rawVendor.Trim();
                // 剔除括号附注内容
                cleaned = Regex.Replace(cleaned, @"(（.*?）|\(.*?\))", "");
                // 剔除公司组织形式后缀
                cleaned = Regex.Replace(cleaned, @"(股份有限公司|有限责任公司|有限公司|分公司|集团|制造厂|厂)", "");
                // 剔除末尾行业属性词
                cleaned = Regex.Replace(cleaned, @"(电气|电器|自动化|成套|技术|科技|设备|系统)$", "");
                // 剔除省份地市前缀
                cleaned = Regex.Replace(cleaned, @"^(浙江|江苏|上海|北京|广东|山东|温州|乐清|苏州|常熟|深圳)", "");
                cleaned = cleaned.Trim();

                // 长度合理则采纳为品牌
                if (cleaned.Length >= 2 && cleaned.Length <= 8)
                {
                    return cleaned;
                }
            }

            // 尝试根据型号前缀特征推导
            string deduced = DeduceBrandFromModel(model);
            if (!string.IsNullOrWhiteSpace(deduced)) return deduced;

            // 回退到用户选定的品牌(若非全部)
            if (!string.IsNullOrWhiteSpace(selectedBrand) && selectedBrand != "全部") return selectedBrand;

            // 最终回退
            return !string.IsNullOrWhiteSpace(rawVendor) ? rawVendor : "正泰";
        }

        /// <summary>
        /// 从电气天下接口返回的条目数据中解析准确的实际品牌 (杜绝将系列分类名误填为品牌)
        /// </summary>
        /// <param name="item">匹配到的 JSON 数据节点</param>
        /// <param name="model">匹配型号</param>
        /// <param name="selectedBrand">用户选定的品牌</param>
        /// <returns>准确的元器件实际品牌名称</returns>
        public static string ResolveBrandForDq123(JsonElement item, string model, string selectedBrand)
        {
            // 步骤 1: 提取 F_ManufactoryID 数值
            int fid = 0;
            if (item.TryGetProperty("F_ManufactoryID", out var fProp))
            {
                // 处理数值类型
                if (fProp.ValueKind == JsonValueKind.Number) fid = fProp.GetInt32();
                // 处理字符串数字类型
                else if (int.TryParse(fProp.GetString(), out int fv)) fid = fv;
            }

            // 若未取到，再尝试读取 manufacturerName
            if (fid == 0 && item.TryGetProperty("manufacturerName", out var mnProp))
            {
                // 处理数值类型
                if (mnProp.ValueKind == JsonValueKind.Number) fid = mnProp.GetInt32();
                // 处理字符串数字类型
                else if (int.TryParse(mnProp.GetString(), out int mv)) fid = mv;
            }

            // 步骤 2: 查电气天下厂家代码对照表精准命中品牌
            if (fid > 0 && Dq123FactoryIdMap.TryGetValue(fid, out var mappedBrand))
            {
                return mappedBrand;
            }

            // 步骤 3: 若用户在界面指定了明确品牌(且不是全部)，直接采纳
            if (!string.IsNullOrWhiteSpace(selectedBrand) && selectedBrand != "全部")
            {
                return selectedBrand;
            }

            // 步骤 4: 从元器件型号特征前缀智能推导实际品牌
            string deduced = DeduceBrandFromModel(model);
            if (!string.IsNullOrWhiteSpace(deduced))
            {
                return deduced;
            }

            // 步骤 5: 检查 className 中是否偶然包含大厂关键词
            string className = item.TryGetProperty("className", out var cn) ? (cn.GetString() ?? "") : "";
            if (!string.IsNullOrWhiteSpace(className))
            {
                foreach (var (kw, norm) in WellKnownBrandKeywords)
                {
                    // 若系列名中包含品牌词
                    if (className.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return norm;
                    }
                }
            }

            // 步骤 6: 默认国内通用元器件厂商兜底
            return "正泰";
        }

        #endregion

        #region 电气天下 (dq123.com) 查询引擎

        /// <summary>
        /// 异步向电气天下接口查询单个元器件价格
        /// </summary>
        /// <param name="model">检索型号</param>
        /// <param name="brand">指定品牌(为空或全部时不限)</param>
        /// <returns>查价结果元组 (状态, 匹配型号, 价格, 厂商, 描述)</returns>
        public static async Task<(bool success, string matchedModel, double price, string vendor, string message)> SearchDq123Async(string model, string brand)
        {
            try
            {
                // 解析厂商 ID
                int? factoryId = null;
                // 若指定了具体品牌且存在于映射表，则指定对应 ID
                if (!string.IsNullOrWhiteSpace(brand) && Dq123BrandMap.TryGetValue(brand.Trim(), out int fid))
                {
                    factoryId = fid;
                }

                // 组装请求 Payload 实体
                var requestPayload = new
                {
                    keyword = model,
                    currentPage = 1,
                    factoryId = factoryId,
                    seriesId = (string?)null
                };

                // 序列化为 JSON 字符串
                string jsonBody = JsonSerializer.Serialize(requestPayload);
                // 构造 UTF-8 JSON 格式的 HTTP 内容
                var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

                // 发起 POST 请求至电气天下查询终结点 --硬编码: 电气天下接口--
                const string apiUrl = "https://www.dq123.com/view/price/search/element";
                var response = await _httpClient.PostAsync(apiUrl, content).ConfigureAwait(false);

                // 校验 HTTP 响应状态码
                if (!response.IsSuccessStatusCode)
                {
                    return (false, string.Empty, 0.0, string.Empty, $"HTTP 响应异常: {response.StatusCode}");
                }

                // 读取响应文本流
                string responseJson = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                // 解析 JSON DOM 节点
                using var doc = JsonDocument.Parse(responseJson);
                var root = doc.RootElement;

                // 检查 data.data 数组节点
                if (!root.TryGetProperty("data", out var dataProp) || !dataProp.TryGetProperty("data", out var listProp) || listProp.ValueKind != JsonValueKind.Array)
                {
                    return (false, string.Empty, 0.0, string.Empty, "未查得该型号价格");
                }

                // 清洗检索词空白字符以便精准比对
                string cleanTarget = Regex.Replace(model, @"\s+", "");

                // 遍历结果集寻找最佳匹配项
                JsonElement bestMatch = default;
                // 标记是否找到
                bool found = false;

                // 优先寻找型号一致或包含的条目
                foreach (var item in listProp.EnumerateArray())
                {
                    // 提取条目中的 model 字段
                    string itemModel = item.TryGetProperty("model", out var mProp) ? (mProp.GetString() ?? "") : "";
                    // 清除空白符
                    string cleanItemModel = Regex.Replace(itemModel, @"\s+", "");

                    // 若无缝完全相等，直接命中最佳项
                    if (cleanItemModel.Equals(cleanTarget, StringComparison.OrdinalIgnoreCase))
                    {
                        bestMatch = item;
                        found = true;
                        break;
                    }
                }

                // 若未完全相等，回退取返回列表中的首项作为推选
                if (!found && listProp.GetArrayLength() > 0)
                {
                    bestMatch = listProp[0];
                    found = true;
                }

                // 若依然未找到任何条目
                if (!found)
                {
                    return (false, string.Empty, 0.0, string.Empty, "未查得该型号价格");
                }

                // 提取匹配的型号
                string matchedM = bestMatch.TryGetProperty("model", out var mp) ? (mp.GetString() ?? "") : model;
                // 提取单价数值
                double priceVal = 0.0;
                if (bestMatch.TryGetProperty("price", out var pp))
                {
                    // 兼容数值型或字符型单价
                    if (pp.ValueKind == JsonValueKind.Number) priceVal = pp.GetDouble();
                    else if (double.TryParse(pp.GetString(), out double pv)) priceVal = pv;
                }

                // 精准提炼电气天下实际品牌名称 (基于厂商代码映射、型号前缀推导与品类清洗，杜绝误填系列名)
                string vendorName = ResolveBrandForDq123(bestMatch, matchedM, brand);

                // 成功返回查价数据
                return (true, matchedM, priceVal, vendorName, "查询成功");
            }
            catch (Exception ex)
            {
                // 捕获异常并返回错误原因
                return (false, string.Empty, 0.0, string.Empty, $"电气天下查价异常: {ex.Message}");
            }
        }

        #endregion

        #region 天工矩阵 (titanmatrix.com) 查询引擎

        // 天工矩阵签名专用 AppID 与 AppKey --硬编码: 天工矩阵签名密钥--
        private const string TitanAppId = "201010";
        private const string TitanAppKey = "72933362EAA649B893699E6191BC898F";

        /// <summary>
        /// 纯原生计算 MD5 32位十六进制小写哈希
        /// </summary>
        private static string ComputeMd5(string input)
        {
            // 创建 MD5 散列计算器
            using var md5 = MD5.Create();
            // 获取 UTF-8 字节序列
            byte[] bytes = Encoding.UTF8.GetBytes(input);
            // 计算散列值
            byte[] hash = md5.ComputeHash(bytes);
            // 组装 32 位十六进制小写字串
            var sb = new StringBuilder(32);
            // 遍历字节逐个转换
            foreach (var b in hash)
            {
                sb.Append(b.ToString("x2"));
            }
            // 返回小写散列结果
            return sb.ToString();
        }

        /// <summary>
        /// 异步向天工矩阵接口查询单个元器件价格
        /// </summary>
        /// <param name="model">检索型号</param>
        /// <param name="brand">指定品牌过滤</param>
        /// <returns>查价结果元组 (状态, 匹配型号, 价格, 厂商, 描述)</returns>
        public static async Task<(bool success, string matchedModel, double price, string vendor, string message)> SearchTitanMatrixAsync(string model, string brand)
        {
            try
            {
                // 严格对齐 JavaScript1.js 中的系统环境 UA 字符串 --硬编码: 天工专用UA--
                const string systemInfo = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36";
                // 客户端 JSON 结构 (对齐 JavaScript1.js 中 f 对象的序列化格式)
                string clientJson = $"{{\"system\":\"{systemInfo}\",\"version\":\"1.0.0\"}}";
                // 参数 JSON 结构 (对齐 JavaScript1.js 中 s 对象的序列化格式)
                string paramJson = $"{{\"input\":[\"{model.Replace("\"", "\\\"")}\"],\"VendorIds\":[]}}";

                // 生成 16 位随机数字 rank 码 (对齐 JavaScript1.js 中 Ce(16))
                var rnd = new Random();
                // 组装 16 位随机数字
                var rankBuilder = new StringBuilder(16);
                for (int i = 0; i < 16; i++) rankBuilder.Append(rnd.Next(0, 10));
                string rank = rankBuilder.ToString();

                // 生成东八区格式化时间戳 yyyy/MM/dd HH:mm:ss (对齐 JavaScript1.js 中 Se.call(Ee(), "yyyy/MM/dd hh:mm:ss"))
                string timestamp = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                // 按照天工矩阵官方 JavaScript1.js 算法拼接待签名原始串
                string signRaw = $"appid={TitanAppId}&client={clientJson}&param={paramJson}&rank={rank}&timestamp={timestamp}&token=&key={TitanAppKey}";
                // 纯原生计算 MD5 散列签名 (与 JavaScript1.js 中 hash(p) 完全一致)
                string sign = ComputeMd5(signRaw);

                // 构造最终 POST JSON 请求报文 (与 JavaScript1.js 返回后序列化的 Rootobject 一致)
                string postBody = $"{{\"appid\":\"{TitanAppId}\",\"client\":{clientJson},\"param\":{paramJson},\"timestamp\":\"{timestamp}\",\"rank\":\"{rank}\",\"sign\":\"{sign}\"}}";

                // 构造针对天工矩阵的专门 HTTP 请求 --硬编码: 天工接口终结点--
                using var request = new HttpRequestMessage(HttpMethod.Post, "https://cpq.titanmatrix.com/api/cpq/search/specsearch");
                // 设置请求内容与编码 UTF-8
                request.Content = new StringContent(postBody, Encoding.UTF8, "application/json");

                // 严格配置天工矩阵跨域与防盗链标头
                request.Headers.Add("Referer", "http://www.titanmatrix.com/tgxx");
                request.Headers.Add("Origin", "http://www.titanmatrix.com");
                request.Headers.Add("apiVersion", "1.0");
                request.Headers.Add("Accept", "application/json, text/plain, */*");

                // 异步发送 HTTP 请求
                var response = await _httpClient.SendAsync(request).ConfigureAwait(false);
                // 校验返回状态码
                if (!response.IsSuccessStatusCode)
                {
                    return (false, string.Empty, 0.0, string.Empty, $"HTTP 响应异常: {(int)response.StatusCode} {response.ReasonPhrase}");
                }

                // 读取响应文本流
                string responseJson = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                // 解析外层 JSON 结构
                using var doc = JsonDocument.Parse(responseJson);
                var root = doc.RootElement;

                // 校验 entity 字段
                if (!root.TryGetProperty("entity", out var entityProp) || entityProp.ValueKind != JsonValueKind.Object)
                {
                    return (false, string.Empty, 0.0, string.Empty, "天工矩阵未返回结果");
                }

                // 提取内层内嵌的 resultData JSON 文本
                if (!entityProp.TryGetProperty("resultData", out var resDataProp) || resDataProp.ValueKind != JsonValueKind.String)
                {
                    return (false, string.Empty, 0.0, string.Empty, "天工矩阵结果集为空");
                }

                // 解析内层数据
                string innerJson = resDataProp.GetString() ?? "[]";
                using var innerDoc = JsonDocument.Parse(innerJson);
                var innerRoot = innerDoc.RootElement;

                // 校验数组有效性
                if (innerRoot.ValueKind != JsonValueKind.Array || innerRoot.GetArrayLength() == 0)
                {
                    return (false, string.Empty, 0.0, string.Empty, "未查得该型号价格");
                }

                // 提取第一组分类中的 list 列表
                var firstCategory = innerRoot[0];
                if (!firstCategory.TryGetProperty("list", out var listArray) || listArray.ValueKind != JsonValueKind.Array || listArray.GetArrayLength() == 0)
                {
                    return (false, string.Empty, 0.0, string.Empty, "未查得该型号价格");
                }

                // 遍历结果集，根据指定品牌进行筛选
                JsonElement? targetItem = null;
                // 是否需要按品牌筛选
                bool needFilterBrand = !string.IsNullOrWhiteSpace(brand) && brand != "全部";

                // 迭代列表项寻找匹配
                foreach (var item in listArray.EnumerateArray())
                {
                    // 若需要筛选品牌
                    if (needFilterBrand)
                    {
                        // 提取 vendor 厂家字段
                        string itemVendor = item.TryGetProperty("vendor", out var v) ? (v.GetString() ?? "") : "";
                        // 提炼实际品牌名称
                        string itemBrand = ResolveBrandForTitan(itemVendor, "", "");
                        // 包含品牌关键字或提炼的品牌相符即可命中
                        if (itemVendor.IndexOf(brand, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            brand.IndexOf(itemBrand, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            itemBrand.IndexOf(brand, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            targetItem = item;
                            break;
                        }
                    }
                    else
                    {
                        // 不限品牌，直接取首个命中项
                        targetItem = item;
                        break;
                    }
                }

                // 若筛选后无结果
                if (!targetItem.HasValue)
                {
                    return (false, string.Empty, 0.0, string.Empty, $"未查得该型号在【{brand}】下的价格");
                }

                // 提取最终匹配条目
                var chosen = targetItem.Value;
                // 获取匹配规格型号名称
                string matchedName = chosen.TryGetProperty("name", out var n) ? (n.GetString() ?? model) : model;
                // 获取原始厂家公司名称
                string rawVendor = chosen.TryGetProperty("vendor", out var vn) ? (vn.GetString() ?? "") : "";
                // 精准提炼为简洁标准的实际品牌名称 (剔除工商冗余词与噪音，如德力西/正泰/常熟开关)
                string finalVendor = ResolveBrandForTitan(rawVendor, matchedName, brand);
                // 提取单价
                double finalPrice = 0.0;
                if (chosen.TryGetProperty("price", out var p))
                {
                    // 转换数值
                    if (p.ValueKind == JsonValueKind.Number) finalPrice = p.GetDouble();
                    else if (double.TryParse(p.GetString(), out double pv)) finalPrice = pv;
                }

                // 返回成功结果
                return (true, matchedName, finalPrice, finalVendor, "查询成功");
            }
            catch (Exception ex)
            {
                // 捕获异常
                return (false, string.Empty, 0.0, string.Empty, $"天工矩阵查价异常: {ex.Message}");
            }
        }

        #endregion

        #region 多任务并发调度引擎

        /// <summary>
        /// 批量执行多任务查价调度，带并发控制与进度上报
        /// </summary>
        /// <param name="items">待查项列表</param>
        /// <param name="config">查价配置</param>
        /// <param name="progress">进度上报通知</param>
        /// <param name="ct">取消令牌</param>
        public static async Task ExecuteBatchSearchAsync(
            List<OnlinePriceItemDto> items,
            OnlinePriceConfig config,
            IProgress<OnlinePriceProgressDto>? progress,
            CancellationToken ct)
        {
            // 选区为空时直接返回
            if (items == null || items.Count == 0) return;

            // 限制并发数量 (最小 1，最大 12)
            int concurrency = Math.Max(1, Math.Min(12, config.MaxConcurrency));
            // 实例化信号量以控制并发池
            using var semaphore = new SemaphoreSlim(concurrency);

            // 总待处理计数
            int totalCount = items.Count;
            // 已完成计数
            int completedCount = 0;

            // 任务集合
            var tasks = new List<Task>();

            // 遍历每个条目分发并发任务
            foreach (var item in items)
            {
                // 检查取消请求
                if (ct.IsCancellationRequested) break;

                // 异步等待获取信号量配额
                await semaphore.WaitAsync(ct).ConfigureAwait(false);

                // 投递并发任务至线程池
                tasks.Add(Task.Run(async () =>
                {
                    try
                    {
                        // 再次检查取消状态
                        if (ct.IsCancellationRequested) return;

                        // 根据平台分流调用接口
                        (bool success, string matchedModel, double price, string vendor, string msg) result;
                        if (config.Platform == OnlinePricePlatform.TitanMatrix)
                        {
                            // 调用天工矩阵查询
                            result = await SearchTitanMatrixAsync(item.SearchModel, config.Brand).ConfigureAwait(false);
                        }
                        else
                        {
                            // 默认调用电气天下查询
                            result = await SearchDq123Async(item.SearchModel, config.Brand).ConfigureAwait(false);
                        }

                        // 回写至数据项传输实体
                        if (result.success)
                        {
                            item.Status = "Success";
                            item.MatchedModel = result.matchedModel;
                            item.Price = result.price;
                            item.Vendor = result.vendor;
                            item.Message = $"查得单价: {result.price:F2}";
                        }
                        else
                        {
                            item.Status = "NotFound";
                            item.Message = result.msg;
                        }

                        // 原子递增已完成数
                        int currentDone = Interlocked.Increment(ref completedCount);
                        // 计算百分比
                        int percent = (int)Math.Round((double)currentDone / totalCount * 100);

                        // 汇报实时进度
                        progress?.Report(new OnlinePriceProgressDto
                        {
                            Current = currentDone,
                            Total = totalCount,
                            Percent = percent,
                            LogText = $"[{currentDone}/{totalCount}] 行 {item.Row}: {item.SearchModel} -> {(result.success ? $"{item.Price:F2}元" : "未查得")}",
                            Item = item
                        });
                    }
                    finally
                    {
                        // 无论成功失败，释放信号量席位
                        semaphore.Release();
                    }
                }, ct));
            }

            // 等待全部并发任务执行收尾
            await Task.WhenAll(tasks).ConfigureAwait(false);
        }

        #endregion
    }
}
