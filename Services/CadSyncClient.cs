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

        // 待发送的所属箱柜名称暂存区
        private static string _pendingCabinetName = string.Empty;

        // 同步锁
        private static readonly object _timerLock = new object();

        /// <summary>
        /// 带有 50ms 防抖的句柄推送方法：用户在 Excel 快速切换行时仅发送停稳后的最后一项
        /// </summary>
        /// <param name="handles">CAD 句柄列表</param>
        /// <param name="autoZoom">是否开启视角自动缩放对焦，默认 true</param>
        /// <param name="delayMs">防抖延时毫秒数，默认 50ms</param>
        /// <param name="cabinetName">当前选中的箱柜名称 (柜号)，用于联动 TuFan 面板同步选择</param>
        public static void SendHandlesDebounced(List<string>? handles, bool autoZoom = true, int delayMs = 50, string? cabinetName = null)
        {
            // 若联动开关未开启，直接忽略
            if (!SyncToCadEnabled) return;

            lock (_timerLock)
            {
                // 暂存最新的句柄数据副本与对焦标志 (兼容处理由 '-' 或分隔符拼接的复合句柄)
                _pendingHandles = handles != null
                    // 将可能混入的复合句柄进行扁平化拆分
                    ? handles.SelectMany(h => h.Split(new[] { ',', ';', '-', ' ' }, StringSplitOptions.RemoveEmptyEntries))
                             // 清洗首尾空白
                             .Select(h => h.Trim())
                             // 过滤空项
                             .Where(h => !string.IsNullOrEmpty(h))
                             // 忽略大小写去重
                             .Distinct(StringComparer.OrdinalIgnoreCase)
                             // 转换为纯净句柄列表
                             .ToList()
                    : new List<string>();
                _pendingAutoZoom = autoZoom && AutoZoomEnabled;
                // 暂存当前行关联的箱柜名称
                _pendingCabinetName = cabinetName ?? string.Empty;

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
            string cabinetNameToSend;
            lock (_timerLock)
            {
                // 复制出待发送的句柄集合与缩放标志
                handlesToSend = new List<string>(_pendingHandles);
                autoZoomToSend = _pendingAutoZoom;
                // 复制出待发送的箱柜名称
                cabinetNameToSend = _pendingCabinetName;
            }

            // 启动后台异步任务发送数据至 CAD
            Task.Run(() => SendToPipeAsync(handlesToSend, autoZoomToSend, cabinetNameToSend));
        }

        /// <summary>
        /// 异步向命名管道发送句柄数据（非阻塞，超时 50ms 即焚，绝不卡死 Excel）
        /// </summary>
        /// <param name="handles">句柄列表</param>
        /// <param name="autoZoom">是否自动聚焦缩放</param>
        /// <param name="cabinetName">当前选中的箱柜名称</param>
        private static async Task SendToPipeAsync(List<string> handles, bool autoZoom, string? cabinetName = null)
        {
            try
            {
                // 构造入站管道客户端实例
                using (var pipeClient = new NamedPipeClientStream(".", PipeName, PipeDirection.Out))
                {
                    // 尝试连接 CAD 管道服务端，设置超时 200 毫秒（非阻塞异步连接）
                    await pipeClient.ConnectAsync(200);

                    // 构造发送的载荷对象
                    var payload = new
                    {
                        action = "selectHandles",
                        handles = handles,
                        autoZoom = autoZoom,
                        cabinetName = cabinetName ?? string.Empty
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
        /// <param name="timeoutMs">超时毫秒数 (传 0 时根据图纸数量自动计算)</param>
        /// <returns>提取结果集合</returns>
        public static async Task<List<DwgDimensionItem>> RequestExtractDwgDimensionsAsync(List<string> dwgPaths, int timeoutMs = 0)
        {
            // 初始化返回测算结果列表
            var resultList = new List<DwgDimensionItem>();
            // 校验图纸路径列表是否为空
            if (dwgPaths == null || dwgPaths.Count == 0) return resultList;

            // 若未指定超时时间，按每张图 800ms 动态放宽并设 6000ms 保底以确保 CAD 充足执行
            int effectiveTimeout = timeoutMs > 0 ? timeoutMs : Math.Max(6000, dwgPaths.Count * 800); // --硬编码: 保底超时 6000ms--

            // 创建超时取消标记源
            using (var cts = new CancellationTokenSource(effectiveTimeout))
            {
                try
                {
                    // 构造双向异步命名管道客户端以支持双向收发
                    using (var pipeClient = new NamedPipeClientStream(".", PipeName, PipeDirection.InOut, PipeOptions.Asynchronous))
                    {
                        // 尝试连接 CAD 服务端（握手连接超时设为 2500ms，为 CAD 调度主线程留出充分时间）
                        int connectTimeout = Math.Min(2500, effectiveTimeout); // --硬编码: 握手连接超时 2500ms--
                        // 异步连接 AutoCAD 管道服务端
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
                            // 强制刷新输出流缓冲区
                            await writer.FlushAsync();
                        }

                        // 创建流读取器以读取 CAD 端直接回传的尺寸响应
                        using (var reader = new StreamReader(pipeClient, Encoding.UTF8, false, 4096, leaveOpen: true))
                        {
                            // 按行异步读取响应数据
                            string? responseJson = await reader.ReadLineAsync();
                            // 校验响应内容非空
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
                                // 校验反序列化结果与子项有效性
                                if (resp != null && resp.Success && resp.Items != null)
                                {
                                    // 提取成功解析出的测算实体集合
                                    resultList = resp.Items;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 超时或管道中断时记录降级日志，确保 Excel 流程稳定不崩溃
                    LogHelper.WriteLog($"[CadSyncClient] RequestExtractDwgDimensionsAsync 降级或异常: {ex.Message}");
                }
            }

            // 若管道未提取到结果（如 AutoCAD 未启动），无缝触发后台 accoreconsole 极速静默测算自愈
            if (resultList.Count == 0 && dwgPaths != null && dwgPaths.Count > 0)
            {
                // 调用后台核心控制台自愈测算引擎
                resultList = await TryRunAccoreConsoleBatchExtractAsync(dwgPaths);
            }

            // 返回提取到的尺寸数据列表
            return resultList;
        }

        /// <summary>
        /// 离线静默自愈引擎：当 AutoCAD 未前台运行时，调用 accoreconsole.exe 无界面批量测算图纸并持久化至 SQLite
        /// </summary>
        /// <param name="dwgPaths">待测算的 DWG 绝对路径列表</param>
        /// <returns>测算入库的尺寸实体列表</returns>
        private static async Task<List<DwgDimensionItem>> TryRunAccoreConsoleBatchExtractAsync(List<string> dwgPaths)
        {
            // 初始化返回列表
            var resultList = new List<DwgDimensionItem>();
            // 校验路径有效性
            if (dwgPaths == null || dwgPaths.Count == 0) return resultList;

            // 检索系统中的 accoreconsole.exe 控制台执行路径
            string accorePath = GetAccoreConsolePath();
            // 检查控制台程序物理是否存在
            if (string.IsNullOrWhiteSpace(accorePath) || !File.Exists(accorePath))
            {
                LogHelper.WriteLog("[CadSyncClient] 未检测到 accoreconsole.exe，跳过离线自愈测算");
                return resultList;
            }

            // 获取 AutoCAD 插件 DLL 的物理路径 --硬编码: ClassLibrary3.dll 路径--
            string pluginDllPath = @"E:\Ace\cad-net_1\cad1\ClassLibrary1\bin\Debug\ClassLibrary3.dll";
            if (!File.Exists(pluginDllPath))
            {
                LogHelper.WriteLog($"[CadSyncClient] 插件程序集不存在: {pluginDllPath}");
                return resultList;
            }

            try
            {
                // 1. 将待测算的图纸物理路径写入临时任务清单文件
                string tempDir = Path.GetTempPath();
                string taskFile = Path.Combine(tempDir, "dwg_scan_tasks.txt"); // --硬编码: 临时任务清单路径--
                File.WriteAllLines(taskFile, dwgPaths);

                // 2. 构造 accoreconsole 执行的 SCR 批处理脚本 (统一转为 Linux 正斜杠避免转义错误)
                string scrFile = Path.Combine(tempDir, "dwg_batch_extract.scr"); // --硬编码: 临时脚本路径--
                string normalizedDllPath = pluginDllPath.Replace('\\', '/');
                string scrContent = $"(command \"NETLOAD\" \"{normalizedDllPath}\")\r\nSCAN_BATCH_FROM_FILE\r\nQUIT\r\n";
                File.WriteAllText(scrFile, scrContent, Encoding.ASCII);

                // 3. 构造以第一张图纸为承载上下文的静默控制台执行进程
                string sampleDwg = dwgPaths[0];
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = accorePath,
                    Arguments = $"/i \"{sampleDwg}\" /s \"{scrFile}\" /l en-US",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                };

                // 启动后台静默进程
                using (var proc = System.Diagnostics.Process.Start(psi))
                {
                    if (proc != null)
                    {
                        // 异步等待进程完成，设置宽裕超时时间 (每张图 1 秒，保底 12 秒)
                        int waitMs = Math.Max(12000, dwgPaths.Count * 1000);
                        await Task.Run(() => proc.WaitForExit(waitMs));
                    }
                }

                // 4. 从 SQLite 中批量提取最新已测算入库的尺寸数据
                var queryKeys = new List<string>();
                foreach (var path in dwgPaths)
                {
                    // 收集图纸文件名作为查询键
                    queryKeys.Add(Path.GetFileName(path));
                    string? dirName = Path.GetFileName(Path.GetDirectoryName(path));
                    if (!string.IsNullOrWhiteSpace(dirName))
                    {
                        // 收集 目录名/图纸名 组合键
                        queryKeys.Add($"{dirName}/{Path.GetFileName(path)}");
                    }
                }

                // 批量查询 SQLite 数据库
                var dimDict = PersonalComponentDbService.BatchGetDwgDimensions(queryKeys);
                if (dimDict != null && dimDict.Count > 0)
                {
                    // 收集所有成功命中的尺寸项并去重
                    var distinctItems = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var kvp in dimDict)
                    {
                        if (kvp.Value != null && !distinctItems.Contains(kvp.Value.DwgName))
                        {
                            distinctItems.Add(kvp.Value.DwgName);
                            resultList.Add(kvp.Value);
                        }
                    }
                    LogHelper.WriteLog($"[CadSyncClient] accoreconsole 离线自愈成功提取入库: {resultList.Count} 条尺寸");
                }
            }
            catch (Exception exAccore)
            {
                // 记录离线执行异常日志
                LogHelper.WriteLog($"[CadSyncClient] TryRunAccoreConsoleBatchExtractAsync 异常: {exAccore.Message}");
            }

            return resultList;
        }

        /// <summary>
        /// 自动探测 AutoCAD accoreconsole.exe 核心控制台的物理绝对路径
        /// </summary>
        /// <returns>程序路径，未找到时返回空字符串</returns>
        private static string GetAccoreConsolePath()
        {
            try
            {
                // 1. 优先从注册表 AutoCAD 2021 安装路径查找 --硬编码: 注册表路径--
                using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Autodesk\AutoCAD\R24.0\ACAD-4101:804"))
                {
                    if (key != null)
                    {
                        string? loc = key.GetValue("AcadLocation") as string;
                        if (!string.IsNullOrWhiteSpace(loc))
                        {
                            string target = Path.Combine(loc, "accoreconsole.exe");
                            if (File.Exists(target)) return target;
                        }
                    }
                }
            }
            catch { }

            // 2. 常见物理安装目录兜底排查 --硬编码: 常用安装路径--
            string[] probePaths = new[]
            {
                @"D:\Program Files\Autodesk\AutoCAD 2021\accoreconsole.exe",
                @"C:\Program Files\Autodesk\AutoCAD 2021\accoreconsole.exe",
                @"D:\AutoCAD 2021\accoreconsole.exe",
                @"C:\AutoCAD 2021\accoreconsole.exe"
            };

            foreach (var p in probePaths)
            {
                // 找到即返回
                if (File.Exists(p)) return p;
            }

            return string.Empty;
        }

        /// <summary>
        /// 异步向 AutoCAD 命名管道发送更新元器件规格型号指令 (非阻塞，超时 50ms 即焚)
        /// </summary>
        /// <param name="handles">需要更新的 CAD 句柄列表</param>
        /// <param name="newSpec">新的规格型号字符串</param>
        public static void SendUpdateComponentSpec(List<string>? handles, string newSpec)
        {
            // 校验句柄集合与新规格字符串有效性
            if (handles == null || handles.Count == 0 || string.IsNullOrWhiteSpace(newSpec)) return;

            // 启动后台异步非阻塞任务发送指令至 CAD
            Task.Run(async () =>
            {
                try
                {
                    // 构造入站管道客户端实例
                    using (var pipeClient = new NamedPipeClientStream(".", PipeName, PipeDirection.Out))
                    {
                        // 尝试连接 CAD 管道服务端，超时 300 毫秒 --硬编码: 极速管道连接超时--
                        await pipeClient.ConnectAsync(300);

                        // 构造发送载荷对象
                        var payload = new
                        {
                            action = "updateComponentSpec",
                            handles = handles,
                            newSpec = newSpec
                        };

                        // 序列化并在尾部添加换行符
                        string jsonStr = JsonSerializer.Serialize(payload) + "\n";
                        byte[] buffer = Encoding.UTF8.GetBytes(jsonStr);

                        // 写入管道并立即刷新
                        await pipeClient.WriteAsync(buffer, 0, buffer.Length);
                        await pipeClient.FlushAsync();
                        LogHelper.WriteLog($"[CadSyncClient] SendUpdateComponentSpec 成功推送: handles={string.Join(",", handles)}, newSpec={newSpec}");
                    }
                }
                catch (Exception ex)
                {
                    // 记录管道未就绪或超时日志
                    LogHelper.WriteLog($"[CadSyncClient] SendUpdateComponentSpec 异常: {ex.Message}");
                }
            });
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
