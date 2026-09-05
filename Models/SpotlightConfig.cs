using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 聚光灯高亮模式枚举
    /// </summary>
    public enum SpotlightMode
    {
        /// <summary>
        /// 十字模式 (同时高亮整行和整列)
        /// </summary>
        Crosshair = 0,

        /// <summary>
        /// 仅行模式 (仅高亮单元格所在整行)
        /// </summary>
        RowOnly = 1,

        /// <summary>
        /// 仅列模式 (仅高亮单元格所在整列)
        /// </summary>
        ColumnOnly = 2
    }

    /// <summary>
    /// 聚光灯功能运行与外观配置模型
    /// </summary>
    public class SpotlightConfig
    {
        /// <summary>
        /// 聚光灯是否开启
        /// </summary>
        [JsonPropertyName("isEnabled")]
        public bool IsEnabled { get; set; } = false;

        /// <summary>
        /// 高亮展示模式：0 十字，1 仅行，2 仅列
        /// </summary>
        [JsonPropertyName("mode")]
        public SpotlightMode Mode { get; set; } = SpotlightMode.Crosshair;

        /// <summary>
        /// 高亮区域十六进制主色调 (遵循系统主色调 #009688 绿蓝相间)
        /// </summary>
        [JsonPropertyName("colorHex")]
        public string ColorHex { get; set; } = "#009688"; // --硬编码: 默认主题色绿蓝相间--

        /// <summary>
        /// 半透明不透明度 (0.05 ~ 0.80，推荐 0.22 保持文字与网格清晰透光)
        /// </summary>
        [JsonPropertyName("opacity")]
        public double Opacity { get; set; } = 0.22; // --硬编码: 默认半透明度 22%--

        /// <summary>
        /// 是否将当前活动单元格镂空保持原样 (默认 false, 连同活动格一同柔和高亮)
        /// </summary>
        [JsonPropertyName("excludeActiveCell")]
        public bool ExcludeActiveCell { get; set; } = false;

        // 静态单例缓存当前运行配置
        private static SpotlightConfig? _current;

        // 线程安全互斥锁
        private static readonly object _lock = new object();

        /// <summary>
        /// 获取当前生效的聚光灯配置实例 (带自愈与持久化初始化)
        /// </summary>
        public static SpotlightConfig Current
        {
            get
            {
                // 双重校验锁单例加载
                if (_current == null)
                {
                    // 加锁保障多线程安全
                    lock (_lock)
                    {
                        // 再次判空防并发重复加载
                        if (_current == null)
                        {
                            // 执行本地磁盘反序列化读取或创建默认配置
                            _current = LoadFromDisk();
                        }
                    }
                }
                // 返回当前有效配置对象
                return _current;
            }
        }

        /// <summary>
        /// 获取配置文件持久化绝对物理路径
        /// </summary>
        private static string GetConfigFilePath()
        {
            // 获取本机 LocalAppData 目录
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            // 组装存储目录路径
            string dir = Path.Combine(localAppData, "ExcelCTTools", "config"); // --硬编码: 配置文件存储路径--
            // 若目录不存在则自动级联创建
            if (!Directory.Exists(dir))
            {
                // 递归建立目录
                Directory.CreateDirectory(dir);
            }
            // 返回 JSON 配置文件完整物理路径
            return Path.Combine(dir, "spotlight_config.json"); // --硬编码: 配置文件名称--
        }

        /// <summary>
        /// 从本地磁盘加载配置，若失败则安全回退到默认设置
        /// </summary>
        public static SpotlightConfig LoadFromDisk()
        {
            try
            {
                // 获取配置文件路径
                string path = GetConfigFilePath();
                // 检查本地文件是否存在
                if (File.Exists(path))
                {
                    // 读取文本内容
                    string json = File.ReadAllText(path);
                    // 校验内容有效性
                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        // 反序列化配置模型
                        var cfg = JsonSerializer.Deserialize<SpotlightConfig>(json);
                        // 若反序列化成功则直接返回
                        if (cfg != null) return cfg;
                    }
                }
            }
            catch
            {
                // 异常静默兜底，避免影响插件启动流程
            }
            // 兜底返回默认配置实例
            return new SpotlightConfig();
        }

        /// <summary>
        /// 将当前配置持久化写回本地磁盘 JSON 文件
        /// </summary>
        public void SaveToDisk()
        {
            try
            {
                // 取得配置文件路径
                string path = GetConfigFilePath();
                // 配置格式化美化选项
                var options = new JsonSerializerOptions { WriteIndented = true };
                // 序列化 JSON 文本
                string json = JsonSerializer.Serialize(this, options);
                // 写入磁盘文件保存
                File.WriteAllText(path, json);
            }
            catch
            {
                // 异常捕获保障健壮性
            }
        }
    }
}
