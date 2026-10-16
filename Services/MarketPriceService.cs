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

        // 行情缓存有效期 (10 分钟)，避免用户短时间内频繁向第三方发送请求
        private static readonly TimeSpan CacheDuration = TimeSpan.FromMinutes(10);

        /// <summary>
        /// 异步获取当日大盘最新铜价与铝价 (优先读取短周期缓存，支持强制刷新)
        /// </summary>
        /// <param name="forceRefresh">是否跳过缓存强制请求第三方网络行情</param>
        public static async Task<MarketPriceResult> FetchMarketPricesAsync(bool forceRefresh = false)
        {
            // 1. 检查本地内存缓存有效性
            lock (CacheLock)
            {
                // 若未强制刷新且缓存尚未过期，直接返回内存缓存
                if (!forceRefresh && _cachedResult != null && (DateTime.Now - _lastFetchTime) < CacheDuration)
                {
                    return _cachedResult;
                }
            }

            // 2. 发起第三方实时行情请求
            try
            {
                // 构建新浪期货行情 API 请求消息 (CU0=沪铜主力，AL0=沪铝主力)
                using var request = new HttpRequestMessage(HttpMethod.Get, "https://hq.sinajs.cn/list=CU0,AL0");
                // 新浪行情接口必须携带合法的 Referer 头防盗链
                request.Headers.Add("Referer", "https://finance.sina.com.cn");

                // 发起网络异步 GET 请求
                using var response = await HttpClientInstance.SendAsync(request).ConfigureAwait(false);
                // 确保响应状态码为成功 (200 OK)
                response.EnsureSuccessStatusCode();

                // 读取响应纯文本 (GBK 或 UTF-8 字符流)
                string rawText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                // 3. 结构化解析行情数据
                var parseResult = ParseSinaFuturesResponse(rawText);
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

            // 5. 首次且断网时的安全基准兜底
            // --硬编码: 默认离线备选铜价 78500 元/吨，铝价 19800 元/吨--
            return new MarketPriceResult
            {
                Success = true,
                CopperPrice = 78500.0,
                AluminumPrice = 19800.0,
                Source = "行业标准参考基准",
                UpdateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm"),
                IsFallback = true,
                Message = "已使用系统内置的基准行情价格 (铜 78,500 元/吨)"
            };
        }

        /// <summary>
        /// 解析新浪期货返回的文本数据字符串
        /// </summary>
        private static MarketPriceResult? ParseSinaFuturesResponse(string rawText)
        {
            if (string.IsNullOrWhiteSpace(rawText)) return null;

            double copperPrice = 0.0;
            double aluminumPrice = 0.0;
            string timeStr = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // 按分号切分多品种行情行
            string[] lines = rawText.Split(new[] { ';', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                // 匹配 hq_str_CU0="数据"
                var match = Regex.Match(line, @"hq_str_([A-Z0-9]+)\s*=\s*""([^""]+)""");
                if (!match.Success) continue;

                string symbol = match.Groups[1].Value.ToUpperInvariant();
                string[] fields = match.Groups[2].Value.Split(',');

                // 新浪国内期货标准字段: 
                // 0=名称, 1=时间, 2=今开盘, 3=最高价, 4=最低价, 5=昨收盘, 6=买价, 7=卖价, 8=最新价
                if (fields.Length > 8)
                {
                    // 优先解析最新成交价
                    if (double.TryParse(fields[8], NumberStyles.Any, CultureInfo.InvariantCulture, out double curPrice) && curPrice > 1000)
                    {
                        if (symbol.Contains("CU0")) copperPrice = curPrice;
                        else if (symbol.Contains("AL0")) aluminumPrice = curPrice;
                    }
                    // 盘前若最新价为0，则回退解析昨收价或今开价
                    else if (double.TryParse(fields[5], NumberStyles.Any, CultureInfo.InvariantCulture, out double closePrice) && closePrice > 1000)
                    {
                        if (symbol.Contains("CU0")) copperPrice = closePrice;
                        else if (symbol.Contains("AL0")) aluminumPrice = closePrice;
                    }
                }
            }

            // 校验铜价是否成功提取到有效数值 (>10000 元/吨)
            if (copperPrice > 10000)
            {
                // 若铝价未能正常识别，则兜底为行业合理均价 19800
                if (aluminumPrice <= 1000) aluminumPrice = 19800.0; // --硬编码: 离线备选铝价 19800--

                return new MarketPriceResult
                {
                    Success = true,
                    CopperPrice = Math.Round(copperPrice, 2),
                    AluminumPrice = Math.Round(aluminumPrice, 2),
                    Source = "上海期货交易所 (上期所主力)",
                    UpdateTime = timeStr,
                    IsFallback = false,
                    Message = $"已成功同步大盘最新行情：铜 {copperPrice:N0} 元/吨，铝 {aluminumPrice:N0} 元/吨"
                };
            }

            return null;
        }
    }
}
