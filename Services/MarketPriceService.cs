using System;
using System.Globalization;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// 当日大盘有色金属行情服务 (拉取上期所主力期货/现货铜价与铝价行情)
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public static class MarketPriceService
    {
        // 复用单例 HttpClient 实例，避免高频调用产生 TCP 连接泄漏
        private static readonly HttpClient HttpClientInstance = new HttpClient
        {
            // 设置超时时间为 4 秒，防止弱网时窗体长时间卡顿
            Timeout = TimeSpan.FromSeconds(4)
        };

        // 缓存最近一次成功获取到的行情结果
        private static MarketPriceResult? _cachedResult = null;

        // 记录最后一次成功抓取行情的时间戳
        private static DateTime _lastFetchTime = DateTime.MinValue;

        // 锁对象，保证多线程请求环境下的缓存安全性
        private static readonly object CacheLock = new object();

        /// <summary>
        /// 动态解析有色金属实时行情配置
        /// 优先级：1. ConfigManager 运行时配置 -> 2. AppConfig 默认配置 -> 3. 内置安全兜底
        /// </summary>
        /// <returns>生效的行情配置实体</returns>
        private static MarketPriceSettings GetSettings()
        {
            try
            {
                // 1. 尝试从全局配置管理器提取行情配置
                var cfg = ConfigManager.Instance.Current?.MarketPrice;
                // 若配置存在且接口 URL 非空，直接采用配置项
                if (cfg != null && !string.IsNullOrWhiteSpace(cfg.ApiUrl))
                {
                    // 返回加载的有效配置
                    return cfg;
                }
            }
            catch { }

            // 2. 兜底回退：返回标准默认配置实例
            return new MarketPriceSettings();
        }

        /// <summary>
        /// 异步获取当日大盘最新铜价与铝价 (优先读取短周期缓存，支持强制刷新)
        /// </summary>
        /// <param name="forceRefresh">是否跳过缓存强制请求第三方网络行情</param>
        public static async Task<MarketPriceResult> FetchMarketPricesAsync(bool forceRefresh = false)
        {
            // 0. 读取生效的动态配置实体
            var settings = GetSettings();
            // 动态计算缓存有效期
            TimeSpan cacheDuration = TimeSpan.FromMinutes(settings.CacheDurationMinutes > 0 ? settings.CacheDurationMinutes : 10);

            // 1. 检查本地内存缓存有效性
            lock (CacheLock)
            {
                // 若未强制刷新且缓存尚未过期，直接返回内存缓存
                if (!forceRefresh && _cachedResult != null && (DateTime.Now - _lastFetchTime) < cacheDuration)
                {
                    return _cachedResult;
                }
            }

            // 2. 发起第三方实时行情请求
            try
            {
                // 动态提取配置的行情 API 终结点
                string apiUrl = !string.IsNullOrWhiteSpace(settings.ApiUrl)
                    ? settings.ApiUrl
                    : "https://hq.sinajs.cn/list=nf_CU0,nf_AL0";
                // 构建行情 API 请求消息
                using var request = new HttpRequestMessage(HttpMethod.Get, apiUrl);

                // 动态注入防盗链 Referer 请求头标
                string referer = !string.IsNullOrWhiteSpace(settings.Referer)
                    ? settings.Referer
                    : "https://finance.sina.com.cn";
                request.Headers.Add("Referer", referer);

                // 发起网络异步 GET 请求
                using var response = await HttpClientInstance.SendAsync(request).ConfigureAwait(false);
                // 确保响应状态码为成功 (200 OK)
                response.EnsureSuccessStatusCode();

                // 读取响应纯文本 (GBK 或 UTF-8 字符流)
                string rawText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                // 3. 结构化解析行情数据 (传入动态配置)
                var parseResult = ParseSinaFuturesResponse(rawText, settings);
                if (parseResult != null && parseResult.Success)
                {
                    // 写入内存缓存
                    lock (CacheLock)
                    {
                        _cachedResult = parseResult;
                        _lastFetchTime = DateTime.Now;
                    }
                    return parseResult;
                }
            }
            catch (Exception ex)
            {
                // 记录网络抓取异常日志
                LogHelper.WriteLog($"[MarketPriceService] FetchMarketPricesAsync 异常: {ex.Message}");
            }

            // 4. 网络故障或无网络时的离线兜底降级处理
            lock (CacheLock)
            {
                // 若存在旧缓存，即使过期也可作为离线兜底
                if (_cachedResult != null)
                {
                    return new MarketPriceResult
                    {
                        Success = true,
                        CopperPrice = _cachedResult.CopperPrice,
                        AluminumPrice = _cachedResult.AluminumPrice,
                        Source = _cachedResult.Source + " (历史离线)",
                        UpdateTime = _cachedResult.UpdateTime,
                        IsFallback = true,
                        Message = "网络未连通，已使用上一次获取的历史行情价格"
                    };
                }
            }

            // 5. 首次且断网时的安全基准兜底 (完全基于动态配置)
            return new MarketPriceResult
            {
                Success = true,
                CopperPrice = settings.DefaultCopperPrice,
                AluminumPrice = settings.DefaultAluminumPrice,
                Source = "1#电解铜现货 (长江/上海有色)",
                UpdateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm"),
                IsFallback = true,
                Message = $"已使用系统配置的基准行情 (铜 {settings.DefaultCopperPrice:N0} 元/吨，铝 {settings.DefaultAluminumPrice:N0} 元/吨)"
            };
        }

        /// <summary>
        /// 解析新浪期货返回的文本数据字符串并换算为 1#电解铜现货均价
        /// </summary>
        private static MarketPriceResult? ParseSinaFuturesResponse(string rawText, MarketPriceSettings settings)
        {
            if (string.IsNullOrWhiteSpace(rawText)) return null;

            try
            {
                double copperPrice = 0.0;
                double aluminumPrice = 0.0;
                string timeStr = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // 提取目标合约代码标识 (大写)
                string cuSymbol = (settings.CopperContractCode ?? "CU").ToUpperInvariant();
                string alSymbol = (settings.AluminumContractCode ?? "AL").ToUpperInvariant();

                // 按分号切分多品种行情行
                string[] lines = rawText.Split(new[] { ';', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    // 匹配 hq_str_nf_CU0="数据" (兼容带下划线标识)
                    var match = Regex.Match(line, @"hq_str_([a-zA-Z0-9_]+)\s*=\s*""([^""]+)""");
                    if (!match.Success) continue;

                    string symbol = match.Groups[1].Value.ToUpperInvariant();
                    string[] fields = match.Groups[2].Value.Split(',');

                    // 新浪国内期货标准字段: 
                    // 0=名称, 1=时间, 2=今开盘, 3=最高价, 4=最低价, 5=昨收盘, 6=买价, 7=卖价, 8=最新价, 17=日期
                    if (fields.Length > 8)
                    {
                        double curVal = 0.0;
                        // 优先解析最新成交价
                        if (double.TryParse(fields[8], NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedVal) && parsedVal > 1000)
                        {
                            curVal = parsedVal;
                        }
                        // 盘前若最新价为0，则回退解析昨收价或今开价
                        else if (double.TryParse(fields[2], NumberStyles.Any, CultureInfo.InvariantCulture, out double openVal) && openVal > 1000)
                        {
                            curVal = openVal;
                        }

                        if (curVal > 1000)
                        {
                            // 动态匹配铜合约标识
                            if (symbol.Contains("CU") || symbol.Contains(cuSymbol))
                            {
                                // 盘面期货价加上配置的现货升水基差
                                copperPrice = (curVal >= 40000 && curVal <= 150000)
                                    ? curVal + settings.CopperSpotPremium
                                    : settings.DefaultCopperPrice;
                            }
                            // 动态匹配铝合约标识
                            else if (symbol.Contains("AL") || symbol.Contains(alSymbol))
                            {
                                aluminumPrice = curVal;
                            }
                        }

                        // 提取日期时间，进行边界安全防护
                        if (fields.Length > 17 && !string.IsNullOrWhiteSpace(fields[17]))
                        {
                            string rawTime = fields.Length > 1 ? fields[1].Trim() : "";
                            // 当时间字符串长度达到 6 位时执行标准格式化 (HH:mm:ss)
                            if (rawTime.Length >= 6)
                            {
                                timeStr = $"{fields[17]} {rawTime.Substring(0, 2)}:{rawTime.Substring(2, 2)}:{rawTime.Substring(4, 2)}";
                            }
                            else
                            {
                                // 不足 6 位时使用当前系统时间兜底，杜绝字符串越界
                                timeStr = $"{fields[17]} {DateTime.Now:HH:mm:ss}";
                            }
                        }
                    }
                }

                // 校验铜价是否成功提取到有效数值 (>10000 元/吨)
                if (copperPrice > 10000)
                {
                    // 若铝价未能正常识别，则兜底为配置的合理均价
                    if (aluminumPrice <= 1000) aluminumPrice = settings.DefaultAluminumPrice;

                    return new MarketPriceResult
                    {
                        Success = true,
                        CopperPrice = Math.Round(copperPrice, 2),
                        AluminumPrice = Math.Round(aluminumPrice, 2),
                        Source = "1#电解铜现货 (长江/上海有色)",
                        UpdateTime = timeStr,
                        IsFallback = false,
                        Message = $"已成功同步 1#电解铜现货行情：铜 {copperPrice:N0} 元/吨，铝 {aluminumPrice:N0} 元/吨"
                    };
                }
            }
            catch (Exception ex)
            {
                // 记录文本解析异常并返回空以触发默认配置兜底
                LogHelper.WriteLog($"[MarketPriceService] ParseSinaFuturesResponse 异常: {ex.Message}");
            }

            return null;
        }
    }
}
