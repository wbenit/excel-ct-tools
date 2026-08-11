using System;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace ExcelAddInDemo
{
    // 配置管理器，采用单例模式统一提供全局配置的加载、保存及兜底处理
    public class ConfigManager
    {
        // 延迟初始化的线程安全静态单例句柄
        private static readonly Lazy<ConfigManager> _instance = new Lazy<ConfigManager>(() => new ConfigManager());

        // 获取全局配置管理器的单例访问点
        public static ConfigManager Instance => _instance.Value;

        // 当前加载生效的 AppConfig 配置实例对象
        public AppConfig Current { get; private set; }

        // 配置文件的实际物理存储路径
        private readonly string _configFilePath;

        // 私有构造函数，完成本地配置路径判定并读取配置文件
        private ConfigManager()
        {
            // 获取当前系统 AppData 目录下的插件专属配置目录
            string appDataDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ExcelAddInDemo"
            );

            // 检查配置存储文件夹是否存在，不存在则自动创建
            if (!Directory.Exists(appDataDir))
            {
                Directory.CreateDirectory(appDataDir);
            }

            // 组合生成完整的配置文件 appsettings.json 绝对路径
            _configFilePath = Path.Combine(appDataDir, "appsettings.json");

            // 执行配置加载初始化
            Current = LoadConfig();
        }

        // 从磁盘读取 JSON 配置，遇到异常或文件缺失时启用兜底机制
        private AppConfig LoadConfig()
        {
            try
            {
                // 首先判断本地 AppData 目录是否存在自定义配置文件
                if (File.Exists(_configFilePath))
                {
                    // 读取本地配置文件的文本内容
                    string json = File.ReadAllText(_configFilePath);

                    // 反序列化 JSON 文本为 AppConfig 强类型对象
                    var config = JsonSerializer.Deserialize<AppConfig>(json);

                    // 校验解析结果，不为空则作为当前配置使用
                    if (config != null)
                    {
                        return config;
                    }
                }
                
                // 若 AppData 中无文件，尝试读取程序同级目录打包的 appsettings.json 默认文件
                string baseConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
                if (File.Exists(baseConfigPath))
                {
                    // 读取打包在程序同级目录的默认配置文件内容
                    string baseJson = File.ReadAllText(baseConfigPath);

                    // 反序列化默认配置文件内容
                    var baseConfig = JsonSerializer.Deserialize<AppConfig>(baseJson);

                    // 解析成功时更新并写回本地 AppData 存储目录
                    if (baseConfig != null)
                    {
                        SaveConfig(baseConfig);
                        return baseConfig;
                    }
                }
            }
            catch (Exception ex)
            {
                // 捕获反序列化或文件 IO 过程中的异常并输出调试日志
                System.Diagnostics.Debug.WriteLine($"加载配置文件遇到异常，启用默认配置: {ex.Message}");
            }

            // 当上述方式均未成功加载时，生成全新的内置默认配置实例
            var defaultConfig = new AppConfig();

            // 将默认配置写回本地 AppData 存储路径
            SaveConfig(defaultConfig);

            // 返回默认配置对象
            return defaultConfig;
        }

        // 将修改后的配置序列化并保存至磁盘文件中
        public void SaveConfig(AppConfig config)
        {
            try
            {
                // 配置 System.Text.Json 的输出选项，开启换行缩进与中文字符友好编码
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                // 将配置模型对象序列化为格式化 JSON 字符串
                string json = JsonSerializer.Serialize(config, options);

                // 将 JSON 内容安全写入配置文件中
                File.WriteAllText(_configFilePath, json);

                // 同步更新内存中的当前配置引用
                Current = config;
            }
            catch (Exception ex)
            {
                // 记录配置文件写入时的异常错误信息
                System.Diagnostics.Debug.WriteLine($"保存配置文件发生异常: {ex.Message}");
            }
        }
    }
}
