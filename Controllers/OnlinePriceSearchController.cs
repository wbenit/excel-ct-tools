using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 在线查价业务控制器，协调配置持久化、远端接口请求与 Excel 批量回写
    /// </summary>
    public class OnlinePriceSearchController
    {
        // 配置文件名称 --硬编码: 查价配置文件名--
        private const string ConfigFileName = "online_price_config.json";

        // JSON 格式化选项
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNameCaseInsensitive = true
        };

        // 内存中缓存的当前配置
        private OnlinePriceConfig _currentConfig;

        /// <summary>
        /// 构造函数: 初始化并加载用户历史配置
        /// </summary>
        public OnlinePriceSearchController()
        {
            // 加载磁盘配置或初始化默认配置
            _currentConfig = LoadConfig();
        }

        /// <summary>
        /// 获取配置文件的绝对存储路径
        /// </summary>
        private string GetConfigFilePath()
        {
            // 获取插件专属 AppData 目录
            string appDataDir = Tool.GetAppDataDirectory();
            // 组装完整路径
            return Path.Combine(appDataDir, ConfigFileName);
        }

        /// <summary>
        /// 从磁盘读取用户查价偏好配置
        /// </summary>
        public OnlinePriceConfig LoadConfig()
        {
            try
            {
                // 获取文件路径
                string path = GetConfigFilePath();
                // 若文件存在则读取解析
                if (File.Exists(path))
                {
                    string json = File.ReadAllText(path);
                    var cfg = JsonSerializer.Deserialize<OnlinePriceConfig>(json, JsonOptions);
                    if (cfg != null)
                    {
                        _currentConfig = cfg;
                        return cfg;
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"OnlinePriceSearchController LoadConfig 异常: {ex.Message}");
            }

            // 失败时回退默认配置
            _currentConfig = new OnlinePriceConfig();
            return _currentConfig;
        }

        /// <summary>
        /// 将当前查价偏好配置持久化存盘
        /// </summary>
        public bool SaveConfig(OnlinePriceConfig config)
        {
            try
            {
                // 更新内存对象
                _currentConfig = config ?? new OnlinePriceConfig();
                // 获取路径
                string path = GetConfigFilePath();
                // 序列化为规范 JSON
                string json = JsonSerializer.Serialize(_currentConfig, JsonOptions);
                // 写入磁盘文件
                File.WriteAllText(path, json);
                return true;
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"OnlinePriceSearchController SaveConfig 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 探测当前 Excel 选区概况
        /// </summary>
        public SelectionDetectDto DetectSelection()
        {
            // 调度服务层进行选区分析
            return ExcelServices.DetectSelectionForPriceSearch(_currentConfig.TrimParentheses);
        }

        /// <summary>
        /// 单项型号精确查价测试
        /// </summary>
        public async Task<OnlinePriceItemDto> SearchSingleAsync(string model, OnlinePricePlatform platform, string brand)
        {
            // 初始化返回项
            var item = new OnlinePriceItemDto
            {
                SearchModel = model,
                Status = "Pending"
            };

            // 判空过滤
            if (string.IsNullOrWhiteSpace(model))
            {
                item.Status = "NotFound";
                item.Message = "请输入要查询的型号";
                return item;
            }

            // 根据平台分发
            (bool success, string matchedModel, double price, string vendor, string msg) res;
            if (platform == OnlinePricePlatform.TitanMatrix)
            {
                res = await OnlinePriceSearchClient.SearchTitanMatrixAsync(model, brand).ConfigureAwait(false);
            }
            else
            {
                res = await OnlinePriceSearchClient.SearchDq123Async(model, brand).ConfigureAwait(false);
            }

            // 封装结果
            if (res.success)
            {
                item.Status = "Success";
                item.MatchedModel = res.matchedModel;
                item.Price = res.price;
                item.Vendor = res.vendor;
                item.Message = $"查得价格: {res.price:F2}";
            }
            else
            {
                item.Status = "NotFound";
                item.Message = res.msg;
            }

            return item;
        }

        /// <summary>
        /// 纯后台执行批量远端价格检索（不触碰任何 Excel COM 对象，彻底杜绝跨线程死锁）
        /// </summary>
        /// <param name="items">主线程预先提取好的元器件内存列表</param>
        /// <param name="config">查价配置</param>
        /// <param name="progressAction">进度通知委托</param>
        /// <param name="ct">取消令牌</param>
        public async Task ExecuteRemoteBatchSearchAsync(
            List<OnlinePriceItemDto> items,
            OnlinePriceConfig config,
            Action<OnlinePriceProgressDto>? progressAction,
            CancellationToken ct)
        {
            // 保存并记住用户配置
            SaveConfig(config);

            // 触发初始进度通知
            progressAction?.Invoke(new OnlinePriceProgressDto
            {
                Current = 0,
                Total = items.Count,
                Percent = 0,
                LogText = $"开始在线查价，共 {items.Count} 个元器件..."
            });

            // 构造异步进度汇报包装器
            var progress = new Progress<OnlinePriceProgressDto>(p => progressAction?.Invoke(p));

            // 执行并发网络抓取
            await OnlinePriceSearchClient.ExecuteBatchSearchAsync(items, config, progress, ct).ConfigureAwait(false);
        }

        /// <summary>
        /// 针对已检索完成的数据集执行 Excel 批量静默回写（由 UI/主线程安全调用）
        /// </summary>
        /// <param name="items">已填充价格的元器件列表</param>
        /// <param name="config">查价配置</param>
        /// <returns>回写操作状态元组</returns>
        public (bool success, int updatedCount, string message) WriteBackToExcel(List<OnlinePriceItemDto> items, OnlinePriceConfig config)
        {
            // 调度服务层进行二维矩阵批量安全回写
            return ExcelServices.BatchWriteBackPrices(items, config);
        }

        /// <summary>
        /// 执行框选一键静默批量查价并自动回写 Excel (兼容保留方法)
        /// </summary>
        /// <param name="config">前端确认的最新配置</param>
        /// <param name="progressAction">实时进度回调委托</param>
        /// <param name="ct">取消令牌</param>
        /// <returns>最终执行汇总结果</returns>
        public async Task<(bool success, int total, int successCount, string message, List<OnlinePriceItemDto> items)> ExecuteBatchProcessAsync(
            OnlinePriceConfig config,
            Action<OnlinePriceProgressDto>? progressAction,
            CancellationToken ct)
        {
            // 保存并记住用户配置
            SaveConfig(config);

            // 1. 提取当前选区条目
            var items = ExcelServices.GetSelectedItemsForSearch(config.TrimParentheses);
            // 判空保护
            if (items == null || items.Count == 0)
            {
                return (false, 0, 0, "当前选区内未检测到有效的元器件型号，请框选包含规格型号的单元格区域", new List<OnlinePriceItemDto>());
            }

            // 2. 执行远端价格抓取
            await ExecuteRemoteBatchSearchAsync(items, config, progressAction, ct).ConfigureAwait(false);

            // 若用户中途取消
            if (ct.IsCancellationRequested)
            {
                return (false, items.Count, 0, "操作已被用户中止", items);
            }

            // 3. 批量静默回写 Excel
            var writeResult = WriteBackToExcel(items, config);

            // 组装最终反馈提示
            string finalMsg = writeResult.success ?
                $"完成！{writeResult.message}" :
                $"查价完成，但回写 Excel 时提示: {writeResult.message}";

            // 打印结束日志
            progressAction?.Invoke(new OnlinePriceProgressDto
            {
                Current = items.Count,
                Total = items.Count,
                Percent = 100,
                LogText = finalMsg
            });

            return (writeResult.success, items.Count, writeResult.updatedCount, finalMsg, items);
        }
    }
}
