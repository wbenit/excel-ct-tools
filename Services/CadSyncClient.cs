using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// Excel 端向 AutoCAD 跨进程同步夹点选中的客户端服务
    /// 支持 50ms 自动防抖与异步非阻塞管道推送
    /// </summary>
    public static class CadSyncClient
    {
        // 目标命名管道名称（与 CAD 端保持一致）--硬编码--
        private const string PipeName = "CadExcelHandleSyncPipe";

        // 全局联动开关（默认开启）
        public static bool SyncToCadEnabled { get; set; } = true;

        // 全局自动缩放对焦开关（默认开启）
        public static bool AutoZoomEnabled { get; set; } = true;

        // 防抖计时器
        private static Timer? _debounceTimer;

        // 待发送的最新句柄集合暂存区
        private static List<string> _pendingHandles = new List<string>();

        // 待发送的是否缩放标志暂存
        private static bool _pendingAutoZoom = true;

        // 同步锁
        private static readonly object _timerLock = new object();

        /// <summary>
        /// 带有 50ms 防抖的句柄推送方法：用户在 Excel 快速切换行时仅发送停稳后的最后一项
        /// </summary>
        /// <param name="handles">CAD 句柄列表</param>
        /// <param name="autoZoom">是否开启视角自动缩放对焦，默认 true</param>
        /// <param name="delayMs">防抖延时毫秒数，默认 50ms</param>
        public static void SendHandlesDebounced(List<string>? handles, bool autoZoom = true, int delayMs = 50)
        {
            // 若联动开关未开启，直接忽略
            if (!SyncToCadEnabled) return;

            lock (_timerLock)
            {
                // 暂存最新的句柄数据副本与对焦标志
                _pendingHandles = handles != null ? new List<string>(handles) : new List<string>();
                _pendingAutoZoom = autoZoom && AutoZoomEnabled;

                // 若计时器已存在则重置触发时间，否则新建一次性计时器
                if (_debounceTimer == null)
                {
                    _debounceTimer = new Timer(OnDebounceTimerFired, null, delayMs, Timeout.Infinite);
                }
                else
                {
                    _debounceTimer.Change(delayMs, Timeout.Infinite);
                }
            }
        }

        /// <summary>
        /// 防抖计时器触发回调：提取最新暂存句柄并投递至异步管道发送任务
        /// </summary>
        private static void OnDebounceTimerFired(object? state)
        {
            List<string> handlesToSend;
            bool autoZoomToSend;
            lock (_timerLock)
            {
                // 复制出待发送的句柄集合与缩放标志
                handlesToSend = new List<string>(_pendingHandles);
                autoZoomToSend = _pendingAutoZoom;
            }

            // 启动后台异步任务发送数据至 CAD
            Task.Run(() => SendToPipeAsync(handlesToSend, autoZoomToSend));
        }

        /// <summary>
        /// 异步向命名管道发送句柄数据（非阻塞，超时 50ms 即焚，绝不卡死 Excel）
        /// </summary>
        /// <param name="handles">句柄列表</param>
        /// <param name="autoZoom">是否自动聚焦缩放</param>
        private static async Task SendToPipeAsync(List<string> handles, bool autoZoom)
        {
            try
            {
                // 构造入站管道客户端实例
                using (var pipeClient = new NamedPipeClientStream(".", PipeName, PipeDirection.Out))
                {
                    // 尝试连接 CAD 管道服务端，设置超时 50 毫秒
                    await pipeClient.ConnectAsync(50);

                    // 构造发送的载荷对象
                    var payload = new
                    {
                        action = "selectHandles",
                        handles = handles,
                        autoZoom = autoZoom
                    };

                    // 序列化为 JSON 字符串并在尾部添加换行符以契合服务端按行读取协议
                    string jsonStr = JsonSerializer.Serialize(payload) + "\n";
                    byte[] buffer = Encoding.UTF8.GetBytes(jsonStr);

                    // 写入管道并刷新缓冲区
                    await pipeClient.WriteAsync(buffer, 0, buffer.Length);
                    await pipeClient.FlushAsync();
                }
            }
            catch
            {
                // CAD 未运行或管道未就绪时静默忽略，保证 Excel 零感无缝运行
            }
        }

        /// <summary>
        /// 同步向 AutoCAD 发送缺失 DWG 图纸尺寸提取请求并等待直接回传结果
        /// 采用 100ms 极速探测（握手上限 80ms），CAD 未响应时零感降级返回空列表，绝不卡死 Excel
        /// </summary>
        /// <param name="dwgPaths">待测算的 DWG 磁盘物理路径列表</param>
        /// <param name="timeoutMs">总等待超时毫秒数，默认 100ms --硬编码: 极速管道探测超时--</param>
        /// <returns>CAD 端返回的尺寸实体列表</returns>
        public static List<DwgDimensionItem> RequestExtractDwgDimensions(List<string>? dwgPaths, int timeoutMs = 100)
        {
            // 校验路径集合是否为空
            if (dwgPaths == null || dwgPaths.Count == 0)
            {
                // 空路径集合直接返回空列表
                return new List<DwgDimensionItem>();
            }

            try
            {
                // 在后台线程运行异步请求并极速等待，超时或未响应立即安全返回
                return Task.Run(() => RequestExtractDwgDimensionsAsync(dwgPaths, timeoutMs)).GetAwaiter().GetResult();
            }
            catch
            {
                // 出现任何不可预见异常或超时降级返回空列表
                return new List<DwgDimensionItem>();
            }
        }

        /// <summary>
        /// 异步向 AutoCAD 命名管道发送 extractDimensions 请求并读取返回结果
        /// </summary>
        /// <param name="dwgPaths">图纸物理路径列表</param>
        /// <param name="timeoutMs">超时毫秒数</param>
        /// <returns>提取结果集合</returns>
        public static async Task<List<DwgDimensionItem>> RequestExtractDwgDimensionsAsync(List<string> dwgPaths, int timeoutMs = 100)
        {
            var resultList = new List<DwgDimensionItem>();
            // 校验图纸路径列表是否为空
            if (dwgPaths == null || dwgPaths.Count == 0) return resultList;

            using (var cts = new CancellationTokenSource(timeoutMs))
            {
                try
                {
                    // 构造双向异步命名管道客户端以支持双向收发
                    using (var pipeClient = new NamedPipeClientStream(".", PipeName, PipeDirection.InOut, PipeOptions.Asynchronous))
                    {
                        // 尝试连接 CAD 服务端（握手连接超时设为 80ms 或剩余时间，CAD 未打开时极速降级）
                        int connectTimeout = Math.Min(80, timeoutMs);
                        await pipeClient.ConnectAsync(connectTimeout, cts.Token);

                        // 创建流写入器并保留流开启状态
                        using (var writer = new StreamWriter(pipeClient, Encoding.UTF8, 4096, leaveOpen: true))
                        {
                            // 构造提取图纸尺寸的请求载荷
                            var reqPayload = new
                            {
                                action = "extractDimensions",
                                dwgPaths = dwgPaths
                            };

                            // 序列化为 JSON 字符串
                            string jsonStr = JsonSerializer.Serialize(reqPayload);
                            // 换行发送并立即刷新缓冲区，服务端采用 ReadLineAsync 读取
                            await writer.WriteLineAsync(jsonStr);
                            await writer.FlushAsync();
                        }

                        // 创建流读取器以读取 CAD 端直接回传的尺寸响应
                        using (var reader = new StreamReader(pipeClient, Encoding.UTF8, false, 4096, leaveOpen: true))
                        {
                            // 按行异步读取响应数据
                            string? responseJson = await reader.ReadLineAsync();
                            if (!string.IsNullOrWhiteSpace(responseJson))
                            {
                                // 配置 JSON 反序列化选项：开启大小写不敏感匹配以兼容 CAD 端返回的 PascalCase 格式
                                var jsonOptions = new JsonSerializerOptions
                                {
                                    // 允许属性名称大小写不敏感（匹配 "Width" 与 "width"）
                                    PropertyNameCaseInsensitive = true
                                };

                                // 反序列化 CAD 返回的尺寸数据载荷（采用大小写兼容配置）
                                var resp = JsonSerializer.Deserialize<CadExtractDimensionsResponse>(responseJson, jsonOptions);
                                if (resp != null && resp.Success && resp.Items != null)
                                {
                                    // 提取成功解析出的测算实体集合
                                    resultList = resp.Items;
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // 超时或管道中断时静默降级，确保 Excel 流程稳定
                }
            }

            return resultList;
        }
    }

    /// <summary>
    /// CAD 端图纸尺寸提取响应数据载荷模型
    /// </summary>
    public class CadExtractDimensionsResponse
    {
        // 动作指令标识
        [JsonPropertyName("action")]
        public string? Action { get; set; }

        // 是否执行成功
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        // 提取出的图纸尺寸实体集合
        [JsonPropertyName("items")]
        public List<DwgDimensionItem>? Items { get; set; }
    }
}
