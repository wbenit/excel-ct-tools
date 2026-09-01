using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

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

                    // 序列化为 JSON 字符串
                    string jsonStr = JsonSerializer.Serialize(payload);
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
    }
}
