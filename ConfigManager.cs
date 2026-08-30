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
            // 通过 Tool 工具类获取 AppData 插件专属目录
            string appDataDir = Tool.GetAppDataDirectory();

            // 组合生成 AppData 中的配置文件 appsettings.json 绝对路径
            _configFilePath = Path.Combine(appDataDir, "appsettings.json");

            // 执行配置加载初始化
            Current = LoadConfig();
        }

        // 从磁盘读取 JSON 配置，优先读取 DLL/XLL 所在物理目录 (如 publish 文件夹) 中的 appsettings.json
        private AppConfig LoadConfig()
        {
            try
            {
                // 通过公共工具类 Tool 安全获取当前插件 DLL/XLL 运行的物理目录
                string currentDir = Tool.GetAppDirectory();

                // 判断是否成功获取到有效目录
                if (!string.IsNullOrWhiteSpace(currentDir))
                {
                    // 拼接 DLL/XLL 同级目录下的 appsettings.json 全路径
                    string localConfigPath = Path.Combine(currentDir, "appsettings.json");

                    // 判断同级目录下是否存在 appsettings.json
                    if (File.Exists(localConfigPath))
                    {
                        try
                        {
                            // 读取同级目录下的配置 JSON 文本
                            string localJson = File.ReadAllText(localConfigPath);

                            // 反序列化为 AppConfig 强类型对象
                            var localConfig = JsonSerializer.Deserialize<AppConfig>(localJson);

                            // 若解析成功，优先使用同级目录下的配置对象
                            if (localConfig != null)
                            {
                                return localConfig;
                            }
                        }
                        catch { }
                    }
                }

                // 若 DLL 同级目录无配置文件，回退读取 AppData 用户专属配置目录
                if (File.Exists(_configFilePath))
                {
                    try
                    {
                        // 读取 AppData 目录下的配置 JSON 文本
                        string json = File.ReadAllText(_configFilePath);

                        // 反序列化 JSON 文本为 AppConfig 强类型对象
                        var config = JsonSerializer.Deserialize<AppConfig>(json);

                        // 校验解析结果，不为空则作为当前配置使用
                        if (config != null)
                        {
                            return config;
                        }
                    }
                    catch { }
                }
            }
            catch { }

            // 当上述方式均未成功加载时，生成全新的内置默认配置实例
            var defaultConfig = new AppConfig();
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

        /// <summary>
        /// 设置并持久化右键菜单模式 (自定义业务菜单 或 Excel 原生菜单)
        /// </summary>
        /// <param name="useCustom">true 为启用 WebView2 业务菜单，false 为完全放行原生右键菜单</param>
        public void SetCustomContextMenuMode(bool useCustom)
        {
            // 读取当前配置对象的引用
            var cfg = Current;
            // 判断当前配置是否有效
            if (cfg != null && cfg.Excel != null)
            {
                // 更新右键菜单模式布尔值
                cfg.Excel.UseCustomContextMenu = useCustom;
                // 执行持久化保存至配置文件
                SaveConfig(cfg);
            }
        }

        /// <summary>
        /// 切换右键菜单模式并返回切换后的最新模式状态
        /// </summary>
        /// <returns>切换后的 UseCustomContextMenu 状态</returns>
        public bool ToggleCustomContextMenuMode()
        {
            // 获取当前右键菜单是否启用自定义模式
            bool currentMode = Current?.Excel?.UseCustomContextMenu ?? true;
            // 计算取反后的目标模式
            bool targetMode = !currentMode;
            // 保存并应用最新的目标模式
            SetCustomContextMenuMode(targetMode);
            // 返回更新后的目标模式状态
            return targetMode;
        }
    }
}
