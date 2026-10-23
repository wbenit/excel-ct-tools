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
        /// <returns>查价结果元组 (状态, 匹配型号, 价格, 厂商, 订货号, 描述)</returns>
        public static async Task<(bool success, string matchedModel, double price, string vendor, string orderNo, string message)> SearchDq123Async(string model, string brand)
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
                    return (false, string.Empty, 0.0, string.Empty, string.Empty, $"HTTP 响应异常: {response.StatusCode}");
                }

                // 读取响应文本流
                string responseJson = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                // 解析 JSON DOM 节点
                using var doc = JsonDocument.Parse(responseJson);
                var root = doc.RootElement;

                // 检查 data.data 数组节点
                if (!root.TryGetProperty("data", out var dataProp) || !dataProp.TryGetProperty("data", out var listProp) || listProp.ValueKind != JsonValueKind.Array)
                {
                    return (false, string.Empty, 0.0, string.Empty, string.Empty, "未查得该型号价格");
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
                    return (false, string.Empty, 0.0, string.Empty, string.Empty, "未查得该型号价格");
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

                // 提取商品订货号或物料编码 (优先读取 orderNum，备选兼容 prodcode)
                string orderNo = string.Empty;
                if (bestMatch.TryGetProperty("orderNum", out var onProp))
                {
                    // 获取订货号字符串
                    orderNo = onProp.GetString() ?? string.Empty;
                }
                else if (bestMatch.TryGetProperty("prodcode", out var pcProp))
                {
                    // 获取备用商品编码
                    orderNo = pcProp.GetString() ?? string.Empty;
                }

                // 成功返回查价数据与订货号
                return (true, matchedM, priceVal, vendorName, orderNo, "查询成功");
            }
            catch (Exception ex)
            {
                // 捕获异常并返回错误原因
                return (false, string.Empty, 0.0, string.Empty, string.Empty, $"电气天下查价异常: {ex.Message}");
            }
        }

        #endregion

        #region 电小二 (Dx2 智能查价引擎)

        /// <summary>
        /// 底层私有辅助方法: 向利驰电气数据库发送原始检索请求
        /// </summary>
        /// <param name="keyword">检索关键词</param>
        /// <param name="factoryId">厂商代号过滤</param>
        /// <returns>返回反序列化的条目集合</returns>
        private static async Task<List<JsonElement>> QueryDq123RawListAsync(string keyword, int? factoryId)
        {
            // 空关键词直接返回空列表
            if (string.IsNullOrWhiteSpace(keyword)) return new List<JsonElement>();

            // 构造请求载荷对象
            var payload = new
            {
                // 检索关键词
                keyword = keyword.Trim(),
                // 分页索引固定取第 1 页
                currentPage = 1,
                // 厂家代码
                factoryId = factoryId,
                // 系列代码为空
                seriesId = (string?)null
            };

            // 序列化请求参数为 JSON 文本
            string jsonBody = JsonSerializer.Serialize(payload);
            // 封装 UTF-8 编码的请求报文
            var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            // 远端接口终结点地址
            const string apiUrl = "https://www.dq123.com/view/price/search/element";
            // 发起异步 POST 请求
            var response = await _httpClient.PostAsync(apiUrl, content).ConfigureAwait(false);

            // 若响应不成功则返回空列表
            if (!response.IsSuccessStatusCode) return new List<JsonElement>();

            // 异步读取响应内容
            string responseJson = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            // 解析 JSON 根文档
            using var doc = JsonDocument.Parse(responseJson);
            // 获取根元素
            var root = doc.RootElement;

            // 提取数据数组节点
            if (!root.TryGetProperty("data", out var dataProp) ||
                !dataProp.TryGetProperty("data", out var listProp) ||
                listProp.ValueKind != JsonValueKind.Array)
            {
                // 节点不存在返回空
                return new List<JsonElement>();
            }

            // 克隆结果项以安全返回
            var resultList = new List<JsonElement>();
            // 遍历每个数据条目
            foreach (var item in listProp.EnumerateArray())
            {
                // 复制并加入返回列表
                resultList.Add(item.Clone());
            }

            // 返回提取的条目集合
            return resultList;
        }

        /// <summary>
        /// 电小二智能分词引擎: 生成针对工程图纸和复合型号的多级检索变体
        /// </summary>
        /// <param name="model">原始型号字符串</param>
        /// <returns>优先级排序的搜索关键词变体列表</returns>
        private static List<string> GenerateDx2SearchCandidates(string model)
        {
            // 初始化候选集并使用有序哈希去重
            var candidates = new List<string>();
            // 内部辅助添加方法
            void AddCandidate(string c)
            {
                // 清洗空白字符
                string trimmed = c?.Trim() ?? string.Empty;
                // 长度有效且不重复时收录
                if (!string.IsNullOrWhiteSpace(trimmed) && !candidates.Contains(trimmed, StringComparer.OrdinalIgnoreCase))
                {
                    // 加入候选清单
                    candidates.Add(trimmed);
                }
            }

            // 候选 1: 原始型号基础规范化
            string cleanRaw = Regex.Replace(model, @"\s+", " ").Trim();
            // 添加清洗后的原词
            AddCandidate(cleanRaw);

            // 候选 2: 施耐德/国产品牌复合漏电智能重组 (如 iC65N 3P D20A Pro+VEA 30mA A)
            if (Regex.IsMatch(cleanRaw, @"(VEA|VIGI|漏电|30mA)", RegexOptions.IgnoreCase))
            {
                // 变体 A: 剥离平台级 Pro+ 前缀并保留完整 VEA/漏电参数 (如: iC65N 3P D20A VEA 30mA A)
                string noPro = Regex.Replace(cleanRaw, @"\bPro\+\s*", "", RegexOptions.IgnoreCase).Trim();
                // 规范多余空格并添加
                noPro = Regex.Replace(noPro, @"\s+", " ").Trim();
                // 关键变体: 精准命中利驰库中 +VEA 整套型号
                AddCandidate(noPro);

                // 变体 B: 剥离尾部单字母分类后缀 (如从 VEA 30mA A 提取 VEA 30mA)
                string shortLeak = Regex.Replace(noPro, @"\s+[A-Z]$", "", RegexOptions.IgnoreCase).Trim();
                // 添加短漏电参数项
                AddCandidate(shortLeak);

                // 变体 C: 转换 VEA 为 ELE 电子式 (如 iC65N 3P D20A ELE 30mA A)
                string eleLeak = Regex.Replace(noPro, @"\bVEA\b", "ELE", RegexOptions.IgnoreCase);
                // 添加电子式脱扣变体
                AddCandidate(eleLeak);

                // 变体 D: 转换 VEA 为通用 Vigi 标识 (如 iC65N 3P D20A Vigi 30mA)
                string vigiLeak = Regex.Replace(shortLeak, @"\bVEA\b", "Vigi", RegexOptions.IgnoreCase);
                // 添加通用 Vigi 变体
                AddCandidate(vigiLeak);

                // 变体 E: 针对 2P 等规格的一体式 Vigi+ 标准命名格式 (如 iC65N Vigi+ 2P C 20A 30mA)
                var matchPI = Regex.Match(cleanRaw, @"(\d+P).*?(?:(?<=\s|/)|^)([CD])\s*(\d+)A?", RegexOptions.IgnoreCase);
                // 提取极数与安数成功
                if (matchPI.Success)
                {
                    // 构造 Vigi+ 序列关键词
                    string vigiCombined = $"iC65N Vigi+ {matchPI.Groups[1].Value} {matchPI.Groups[2].Value} {matchPI.Groups[3].Value}A 30mA";
                    // 格式化多余空格
                    vigiCombined = Regex.Replace(vigiCombined, @"\s+", " ").Trim();
                    // 加入候选集
                    AddCandidate(vigiCombined);
                }
            }

            // 候选 3: 参数粘连规范化处理 (如 2PC20A -> 2P C20A, C20A/2P -> 2P C20A)
            string spaceFixed = cleanRaw;
            // 处理斜杠格式 C20A/2P
            var slashMatch = Regex.Match(spaceFixed, @"([A-Z0-9\-]+)-([CD]\d+A?)/(\d+P)", RegexOptions.IgnoreCase);
            // 成功匹配斜杠形式
            if (slashMatch.Success)
            {
                // 转换为常规标准空格形式
                string converted = $"{slashMatch.Groups[1].Value} {slashMatch.Groups[3].Value} {slashMatch.Groups[2].Value}";
                // 添加转换项
                AddCandidate(converted);
            }

            // 候选 4: 剥离复合附件获取微断/塑壳主开关本体 (供纯断路器或拼装回退)
            string baseModel = Regex.Replace(cleanRaw, @"(\+?\s*(Pro\+)?VEA.*|\+?\s*Vigi.*|\+?\s*30mA.*|(?:\s*\(.*?\)))", "", RegexOptions.IgnoreCase).Trim();
            // 避免与已有重复且保证长度合理
            if (!string.IsNullOrWhiteSpace(baseModel) && baseModel.Length >= 4)
            {
                // 添加主开关本体
                AddCandidate(baseModel);
            }

            // 返回整理好的候选序列
            return candidates;
        }

        /// <summary>
        /// 电小二智能打分优选算法: 评估条目与目标规格的契合度
        /// </summary>
        /// <param name="item">候选 JSON 条目</param>
        /// <param name="targetModel">用户原始检索词</param>
        /// <param name="targetBrand">预期品牌</param>
        /// <returns>匹配得分数值 (分数越高契合度越佳)</returns>
        private static int ScoreDx2Candidate(JsonElement item, string targetModel, string targetBrand)
        {
            // 基础初始得分
            int score = 0;
            // 提取条目型号
            string itemModel = item.TryGetProperty("model", out var mp) ? (mp.GetString() ?? "") : "";
            // 提取条目分类/系列名称
            string className = item.TryGetProperty("className", out var cn) ? (cn.GetString() ?? "") : "";

            // 提取条目单价
            double priceVal = 0.0;
            if (item.TryGetProperty("price", out var pp))
            {
                // 读取数值价格
                if (pp.ValueKind == JsonValueKind.Number) priceVal = pp.GetDouble();
                // 兼容字符串价格
                else if (double.TryParse(pp.GetString(), out double pv)) priceVal = pv;
            }

            // 准则 1: 有效官方价格加权 (+100分)
            if (priceVal > 0.01) score += 100;
            // 无价格扣分
            else score -= 150;

            // 准则 2: 在产现行型号加权 (排除已停产条目)
            if (className.Contains("已停产") || itemModel.Contains("已停产")) score -= 80;
            // 在产正常加分
            else score += 30;

            // 准则 3: 极数强相关校验 (如 2P、3P、4P、1P)
            var poleMatch = Regex.Match(targetModel, @"\b(\d+P(\+N)?)\b", RegexOptions.IgnoreCase);
            // 若原始输入有明确极数
            if (poleMatch.Success)
            {
                // 获取目标极数字符串
                string targetPole = poleMatch.Groups[1].Value.ToUpperInvariant();
                // 条目包含相同极数
                if (itemModel.IndexOf(targetPole, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // 极数符合加 50 分
                    score += 50;
                }
                else
                {
                    // 极数错误重扣 120 分 (例如 2P 绝不能匹配为 3P)
                    score -= 120;
                }
            }

            // 准则 4: 额定电流强相关校验 (如 20A、D20、C20、100A 等)
            var currentMatch = Regex.Match(targetModel, @"(?:\b|[CD])(\d+)(?:A|\b)", RegexOptions.IgnoreCase);
            // 若存在电流参数
            if (currentMatch.Success)
            {
                // 获取安数
                string targetAmp = currentMatch.Groups[1].Value;
                // 条目是否包含对应安数
                if (Regex.IsMatch(itemModel, $@"\b[CD]?{targetAmp}A?\b", RegexOptions.IgnoreCase))
                {
                    // 安数匹配加 50 分
                    score += 50;
                }
                else
                {
                    // 安数不符扣 80 分
                    score -= 80;
                }
            }

            // 准则 5: 漏电特征保护校验 (若目标包含 VEA/Vigi/30mA，条目必须包含漏电标识)
            bool isTargetLeakage = Regex.IsMatch(targetModel, @"(VEA|Vigi|漏电|30mA)", RegexOptions.IgnoreCase);
            // 校验条目是否包含漏电标识
            bool isItemLeakage = Regex.IsMatch(itemModel + className, @"(Vigi|VEA|ELE|ELM|LE|漏电|剩余电流)", RegexOptions.IgnoreCase);
            // 漏电特征同时满足
            if (isTargetLeakage && isItemLeakage)
            {
                // 漏电属性一致重磅加 80 分
                score += 80;
            }
            // 目标明确要求漏电，但条目为纯断路器！
            else if (isTargetLeakage && !isItemLeakage)
            {
                // 必须重扣 250 分，坚决防止将纯断路器误判为漏电断路器返回！
                score -= 250;
            }

            // 准则 6: 文本去空格无缝相等超高加权
            string cTarget = Regex.Replace(targetModel, @"\s+", "");
            string cItem = Regex.Replace(itemModel, @"\s+", "");
            // 完全一致加 120 分
            if (cTarget.Equals(cItem, StringComparison.OrdinalIgnoreCase)) score += 120;

            // 返回综合评分
            return score;
        }

        /// <summary>
        /// 异步调用电小二智能查价引擎 (整合复合型号拆解、漏电特征识别与多级智能优选)
        /// </summary>
        /// <param name="model">检索型号</param>
        /// <param name="brand">指定品牌</param>
        /// <returns>查价结果元组 (状态, 匹配型号, 价格, 厂商, 订货号, 描述)</returns>
        public static async Task<(bool success, string matchedModel, double price, string vendor, string orderNo, string message)> SearchDx2Async(string model, string brand)
        {
            try
            {
                // 判空过滤
                if (string.IsNullOrWhiteSpace(model)) return (false, string.Empty, 0.0, string.Empty, string.Empty, "未输入检索型号");

                // 智能解析目标厂商代号
                int? factoryId = null;
                // 优先从界面用户选择解析
                if (!string.IsNullOrWhiteSpace(brand) && brand != "全部" && Dq123BrandMap.TryGetValue(brand.Trim(), out int fid))
                {
                    // 确定厂家代码
                    factoryId = fid;
                }
                // 若未指定则从型号特征自动推导
                else
                {
                    // 智能推导品牌
                    string deducedBrand = DeduceBrandFromModel(model);
                    // 匹配厂商表
                    if (!string.IsNullOrWhiteSpace(deducedBrand) && Dq123BrandMap.TryGetValue(deducedBrand, out int dfid))
                    {
                        // 赋予推导代码
                        factoryId = dfid;
                    }
                }

                // 阶段 1: 生成电小二智能分词候选序列
                var searchCandidates = GenerateDx2SearchCandidates(model);
                // 收集到的所有候选项与评分
                var candidatePool = new List<(JsonElement item, int score)>();

                // 依次检索各个候选序列
                foreach (var candidateWord in searchCandidates)
                {
                    // 发起远端检索
                    var rawList = await QueryDq123RawListAsync(candidateWord, factoryId).ConfigureAwait(false);
                    // 校验条目列表
                    if (rawList != null && rawList.Count > 0)
                    {
                        // 对条目执行智能打分
                        foreach (var elem in rawList)
                        {
                            // 计算契合度评分
                            int score = ScoreDx2Candidate(elem, model, brand);
                            // 记录有效候选
                            candidatePool.Add((elem, score));
                        }

                        // 若已命中高分项 (高于 160 分)，直接跳出后续候选检索
                        if (candidatePool.Any(c => c.score >= 160)) break;
                    }
                }

                // 阶段 2: 择优选取最佳评分条目
                if (candidatePool.Count > 0)
                {
                    // 按得分降序排序取最高分
                    var best = candidatePool.OrderByDescending(c => c.score).First();
                    // 提取最高分条目
                    var bestElem = best.item;

                    // 提取型号与价格
                    string bestModel = bestElem.TryGetProperty("model", out var bm) ? (bm.GetString() ?? "") : model;
                    double bestPrice = 0.0;
                    if (bestElem.TryGetProperty("price", out var bp))
                    {
                        // 解析数值单价
                        if (bp.ValueKind == JsonValueKind.Number) bestPrice = bp.GetDouble();
                        else if (double.TryParse(bp.GetString(), out double bVal)) bestPrice = bVal;
                    }

                    // 提取并解析品牌厂商
                    string vendorName = ResolveBrandForDq123(bestElem, bestModel, brand);

                    // 提取商品订货号或物料编码 (优先 orderNum，备选 prodcode)
                    string bestOrderNo = string.Empty;
                    if (bestElem.TryGetProperty("orderNum", out var onProp))
                    {
                        // 获取订货号
                        bestOrderNo = onProp.GetString() ?? string.Empty;
                    }
                    else if (bestElem.TryGetProperty("prodcode", out var pcProp))
                    {
                        // 获取商品编码
                        bestOrderNo = pcProp.GetString() ?? string.Empty;
                    }

                    // 检查漏电冲突 (若用户明确检索漏电，但最高分条目仅为纯断路器，则绝不能直接采纳，必须转入拼装求和阶段)
                    bool isTargetLeakage = Regex.IsMatch(model, @"(VEA|Vigi|漏电|30mA)", RegexOptions.IgnoreCase);
                    bool isBestItemLeakage = Regex.IsMatch(bestModel, @"(Vigi|VEA|ELE|ELM|LE|漏电|剩余电流)", RegexOptions.IgnoreCase);

                    // 评分合格 (得分大于等于 50 分) 且价格有效，且不存在漏电缺失冲突
                    if ((!isTargetLeakage || isBestItemLeakage) && best.score >= 50 && bestPrice > 0.01)
                    {
                        // 成功返回电小二智能优选结果与订货号
                        return (true, bestModel, bestPrice, vendorName, bestOrderNo, $"电小二查得: ¥{bestPrice:F2} (匹配度:{best.score})");
                    }
                }

                // 阶段 3: 拼装拆解多组件求和兜底 (主开关 + 独立漏电保护附件)
                if (Regex.IsMatch(model, @"(VEA|Vigi|漏电|30mA)", RegexOptions.IgnoreCase))
                {
                    // 提取主断路器型号 (如 iC65N 2P C20A)
                    string mainBreakerModel = Regex.Replace(model, @"(\+?\s*(Pro\+)?VEA.*|\+?\s*Vigi.*|\+?\s*30mA.*)", "", RegexOptions.IgnoreCase).Trim();
                    // 提取极数
                    var pMatch = Regex.Match(model, @"(\d+P)", RegexOptions.IgnoreCase);
                    string poles = pMatch.Success ? pMatch.Groups[1].Value : "2P";

                    // 异步检索主开关
                    var mainBreakerResult = await QueryDq123RawListAsync(mainBreakerModel, factoryId).ConfigureAwait(false);
                    // 查找合适的主开关
                    var mainMatch = mainBreakerResult?.FirstOrDefault(x =>
                    {
                        // 价格有效且极数相同
                        double p = x.TryGetProperty("price", out var pr) && pr.ValueKind == JsonValueKind.Number ? pr.GetDouble() : 0.0;
                        string m = x.TryGetProperty("model", out var mo) ? mo.GetString() ?? "" : "";
                        return p > 0 && m.IndexOf(poles, StringComparison.OrdinalIgnoreCase) >= 0;
                    });

                    // 异步检索对应的漏电附件 (如 Vigi iC65 ELE 2P 30mA A)
                    string vigiAccKeyword = $"Vigi iC65 ELE {poles} 30mA A";
                    var vigiAccResult = await QueryDq123RawListAsync(vigiAccKeyword, factoryId).ConfigureAwait(false);
                    // 寻找价格大于 0 的漏电附件
                    var vigiMatch = vigiAccResult?.FirstOrDefault(x =>
                    {
                        double p = x.TryGetProperty("price", out var pr) && pr.ValueKind == JsonValueKind.Number ? pr.GetDouble() : 0.0;
                        return p > 0;
                    });

                    // 若两件均查得
                    if (mainMatch.HasValue && vigiMatch.HasValue)
                    {
                        // 提取主开关单价
                        double mainP = mainMatch.Value.GetProperty("price").GetDouble();
                        string mainM = mainMatch.Value.GetProperty("model").GetString() ?? mainBreakerModel;
                        // 提取附件单价
                        double vigiP = vigiMatch.Value.GetProperty("price").GetDouble();
                        string vigiM = vigiMatch.Value.GetProperty("model").GetString() ?? vigiAccKeyword;

                        // 合并总面价
                        double totalP = mainP + vigiP;
                        // 组合官方型号描述
                        string combinedModel = $"{mainM} + {vigiM}";
                        // 解析厂商
                        string vendorName = ResolveBrandForDq123(mainMatch.Value, combinedModel, brand);

                        // 提取主开关与附件订货号
                        string mainOrder = mainMatch.Value.TryGetProperty("orderNum", out var mon) ? (mon.GetString() ?? "") : "";
                        string vigiOrder = vigiMatch.Value.TryGetProperty("orderNum", out var von) ? (von.GetString() ?? "") : "";
                        // 组合订货号文本
                        string combinedOrder = !string.IsNullOrWhiteSpace(mainOrder) && !string.IsNullOrWhiteSpace(vigiOrder) ? $"{mainOrder} + {vigiOrder}" : (!string.IsNullOrWhiteSpace(mainOrder) ? mainOrder : vigiOrder);

                        // 返回拼装求和结果
                        return (true, combinedModel, totalP, vendorName, combinedOrder, $"电小二拼装查得: 主体(¥{mainP:F2}) + 漏电模块(¥{vigiP:F2}) = ¥{totalP:F2}");
                    }
                }

                // 阶段 4: 若智能流程均未果，回退至基础查询
                return await SearchDq123Async(model, brand).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // 异常保护返回详细报错
                return (false, string.Empty, 0.0, string.Empty, string.Empty, $"电小二查价异常: {ex.Message}");
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
        /// <returns>查价结果元组 (状态, 匹配型号, 价格, 厂商, 订货号, 描述)</returns>
        public static async Task<(bool success, string matchedModel, double price, string vendor, string orderNo, string message)> SearchTitanMatrixAsync(string model, string brand)
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
                    return (false, string.Empty, 0.0, string.Empty, string.Empty, $"HTTP 响应异常: {(int)response.StatusCode} {response.ReasonPhrase}");
                }

                // 读取响应文本流
                string responseJson = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                // 解析外层 JSON 结构
                using var doc = JsonDocument.Parse(responseJson);
                var root = doc.RootElement;

                // 校验 entity 字段
                if (!root.TryGetProperty("entity", out var entityProp) || entityProp.ValueKind != JsonValueKind.Object)
                {
                    return (false, string.Empty, 0.0, string.Empty, string.Empty, "天工矩阵未返回结果");
                }

                // 提取内层内嵌的 resultData JSON 文本
                if (!entityProp.TryGetProperty("resultData", out var resDataProp) || resDataProp.ValueKind != JsonValueKind.String)
                {
                    return (false, string.Empty, 0.0, string.Empty, string.Empty, "天工矩阵结果集为空");
                }

                // 解析内层数据
                string innerJson = resDataProp.GetString() ?? "[]";
                using var innerDoc = JsonDocument.Parse(innerJson);
                var innerRoot = innerDoc.RootElement;

                // 校验数组有效性
                if (innerRoot.ValueKind != JsonValueKind.Array || innerRoot.GetArrayLength() == 0)
                {
                    return (false, string.Empty, 0.0, string.Empty, string.Empty, "未查得该型号价格");
                }

                // 提取第一组分类中的 list 列表
                var firstCategory = innerRoot[0];
                if (!firstCategory.TryGetProperty("list", out var listArray) || listArray.ValueKind != JsonValueKind.Array || listArray.GetArrayLength() == 0)
                {
                    return (false, string.Empty, 0.0, string.Empty, string.Empty, "未查得该型号价格");
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
                    return (false, string.Empty, 0.0, string.Empty, string.Empty, $"未查得该型号在【{brand}】下的价格");
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

                // 提取商品订货号或物料编码 (天工矩阵核心字段 prodcode，备选兼容 orderNo)
                string prodCode = chosen.TryGetProperty("prodcode", out var pcode) ? (pcode.GetString() ?? "") : "";
                if (string.IsNullOrWhiteSpace(prodCode) && chosen.TryGetProperty("orderNo", out var ono))
                {
                    // 获取备选订货号
                    prodCode = ono.GetString() ?? "";
                }

                // 返回成功结果与订货号
                return (true, matchedName, finalPrice, finalVendor, prodCode, "查询成功");
            }
            catch (Exception ex)
            {
                // 捕获异常
                return (false, string.Empty, 0.0, string.Empty, string.Empty, $"天工矩阵查价异常: {ex.Message}");
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
                        (bool success, string matchedModel, double price, string vendor, string orderNo, string msg) result;
                        if (config.Platform == OnlinePricePlatform.TitanMatrix)
                        {
                            // 调用天工矩阵查询
                            result = await SearchTitanMatrixAsync(item.SearchModel, config.Brand).ConfigureAwait(false);
                        }
                        else if (config.Platform == OnlinePricePlatform.Dx2)
                        {
                            // 调用电小二智能查价引擎
                            result = await SearchDx2Async(item.SearchModel, config.Brand).ConfigureAwait(false);
                        }
                        else
                        {
                            // 默认调用电气天下基础查询
                            result = await SearchDq123Async(item.SearchModel, config.Brand).ConfigureAwait(false);
                        }

                        // 回写至数据项传输实体
                        if (result.success)
                        {
                            item.Status = "Success";
                            item.MatchedModel = result.matchedModel;
                            item.Price = result.price;
                            item.Vendor = result.vendor;
                            item.OrderNo = result.orderNo;
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

                        // 组装订货号附注文字
                        string orderInfo = !string.IsNullOrWhiteSpace(item.OrderNo) ? $" [{item.OrderNo}]" : "";
                        // 汇报实时进度
                        progress?.Report(new OnlinePriceProgressDto
                        {
                            Current = currentDone,
                            Total = totalCount,
                            Percent = percent,
                            LogText = $"[{currentDone}/{totalCount}] 行 {item.Row}: {item.SearchModel} -> {(result.success ? $"{item.Price:F2}元{orderInfo}" : "未查得")}",
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
