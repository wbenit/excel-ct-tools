using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 全局系统与运行环境设置实体数据契约
    /// </summary>
    public class AppSettingsData
    {
        // CAD 与 Excel 共享的数据与配置文件存储目录（如调价公式、企业设置、物料库等）
        public string CustomDataDirectory { get; set; } = string.Empty;

        // 新建工程项目时的默认保存路径（留空则回退至用户桌面）
        public string DefaultProjectDirectory { get; set; } = string.Empty;
    }

    /// <summary>
    /// 系统设置服务控制器，提供读写全局配置的异步服务
    /// </summary>
    public class AppSettingsController
    {
        // 声明全局序列化选项
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            // 启用驼峰命名
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            // 忽略属性大小写
            PropertyNameCaseInsensitive = true,
            // 格式化缩进
            WriteIndented = true
        };

        /// <summary>
        /// 异步加载当前全局生效的配置实体
        /// </summary>
        /// <returns>系统设置实体数据</returns>
        public async Task<AppSettingsData> LoadSettingsAsync()
        {
            // 在后台 Task 线程中读取配置，防止阻塞 UI 线程
            return await Task.Run(() =>
            {
                try
                {
                    // 1. 读取当前生效的数据共享目录
                    string customDataDir = Tool.GetCustomDataDirectoryFromGlobalConfig();
                    // 若未配置，兜底展示插件自身的默认 data 目录
                    if (string.IsNullOrWhiteSpace(customDataDir))
                    {
                        // 获取插件当前安装目录下的 data 物理路径
                        customDataDir = Path.Combine(Tool.GetAppDirectory(), "data");
                    }

                    // 2. 读取当前生效的新建项目默认保存目录
                    string defaultProjDir = Tool.GetDefaultProjectDirectoryFromGlobalConfig();
                    // 若未配置，兜底展示用户系统桌面路径
                    if (string.IsNullOrWhiteSpace(defaultProjDir))
                    {
                        // 提取桌面路径
                        defaultProjDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    }

                    // 封装并返回设置实体
                    return new AppSettingsData
                    {
                        CustomDataDirectory = customDataDir,
                        DefaultProjectDirectory = defaultProjDir
                    };
                }
                catch (Exception ex)
                {
                    // 记录读取异常日志
                    LogHelper.WriteLog($"[AppSettingsController] LoadSettingsAsync 异常: {ex.Message}");
                    // 异常时返回空实体兜底
                    return new AppSettingsData();
                }
            });
        }

        /// <summary>
        /// 异步原子持久化保存全局系统设置
        /// </summary>
        /// <param name="settings">前端提交的设置数据实体</param>
        /// <returns>保存是否成功</returns>
        public async Task<bool> SaveSettingsAsync(AppSettingsData settings)
        {
            // 校验入参对象非空
            if (settings == null) return false;

            // 在后台 Task 线程中完成持久化与文件同步
            return await Task.Run(() =>
            {
                try
                {
                    // 提取并清洗数据目录路径
                    string dataDir = settings.CustomDataDirectory?.Trim() ?? string.Empty;
                    // 提取并清洗项目保存目录路径
                    string projDir = settings.DefaultProjectDirectory?.Trim() ?? string.Empty;

                    // 调用底层 Tool 的原子配置持久化方法
                    return Tool.SaveGlobalSettings(dataDir, projDir);
                }
                catch (Exception ex)
                {
                    // 记录持久化异常日志
                    LogHelper.WriteLog($"[AppSettingsController] SaveSettingsAsync 异常: {ex.Message}");
                    return false;
                }
            });
        }
    }
}
