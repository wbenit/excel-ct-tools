using System;
using System.IO;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 系统公共日志记录工具类，提供控制台与本地文件双写日志功能
    /// </summary>
    public static class LogHelper
    {
        // 静态日志追加锁对象，防止并发写入发生文件占包冲突
        private static readonly object _logLock = new object();

        /// <summary>
        /// 写入运行日志到 IDE 调试控制台及 %AppData%/ExcelAddInDemo/debug.log 文本文件
        /// </summary>
        /// <param name="message">待记录的日志文本消息</param>
        public static void WriteLog(string message)
        {
            try
            {
                // 1. 同步输出到 IDE 控制台调试窗口
                System.Diagnostics.Debug.WriteLine(message);

                // 2. 调用公共 Tool 工具类获取程序集所在专属日志保存目录
                string logDir = Tool.GetAppDirectory();
                // 拼接程序集目录下 debug.log 全路径字符串
                string logFilePath = Path.Combine(logDir, "debug.log");

                // 获取 AppData 漫游目录
                string appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelAddInDemo");
                // 确保 AppData 目录存在
                if (!Directory.Exists(appDataDir)) Directory.CreateDirectory(appDataDir);
                // 拼接 AppData 目录下 debug.log 全路径
                string appDataLogPath = Path.Combine(appDataDir, "debug.log");

                // 格式化日志文本内容
                string logContent = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";

                // 使用线程并发锁保障写盘追加操作安全
                lock (_logLock)
                {
                    // 追加写入程序集目录日志
                    try { File.AppendAllText(logFilePath, logContent); } catch { }
                    // 追加写入 AppData 漫游目录日志
                    try { File.AppendAllText(appDataLogPath, logContent); } catch { }
                }
            }
            catch
            {
                // 忽略日志追加过程中的静默捕获异常
            }
        }
    }
}
