using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 企业设置数据实体模型，记录用户配置的所有文本与选项
    /// </summary>
    public class EnterpriseSettingsData
    {
        // 企业中文名称
        public string CompanyName { get; set; } = "扬州华科智能科技有限公司";

        // 企业英文名称
        public string EnglishName { get; set; } = "Yangzhou Huake Intelligent Technology Co., Ltd.";

        // 企业 Logo 图片数据 (Base64 编码字符串或文件路径)
        public string LogoBase64 { get; set; } = string.Empty;

        // 报价人姓名
        public string Quoter { get; set; } = "吴磊";

        // 联系人姓名
        public string ContactPerson { get; set; } = "";

        // 联系电话
        public string ContactPhone { get; set; } = "";

        // 销售地区
        public string SalesRegion { get; set; } = "";

        // 选中的报价说明模板分类 (如 "说明2")
        public string QuoteTemplate { get; set; } = "说明2";

        // 报价说明详细多行文本内容
        public string QuoteDescription { get; set; } = "123444";

        // 保存设置时是否同步更新当前打开的已选中项目
        public bool SyncOpenProject { get; set; } = false;

        // 自定义数据配置共享存储目录 (留空则默认使用插件目录下的 data 目录)
        public string CustomDataDirectory { get; set; } = string.Empty;
    }

    /// <summary>
    /// 企业设置 Backend WebAPI/本地服务控制器
    /// </summary>
    public class EnterpriseSettingsController
    {
        // 企业设置配置文件的默认保存文件名 --硬编码--
        private const string SettingsFileName = "EnterpriseSettings.json";

        // 定义全局 JSON 序列化选项，开启驼峰命名与忽略大小写匹配以保障前后端属性精准同步
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            // 启用驼峰命名转换
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            // 启用属性名称忽略大小写校验
            PropertyNameCaseInsensitive = true,
            // 启用 JSON 格式缩进输出
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数
        /// </summary>
        public EnterpriseSettingsController()
        {
        }

        /// <summary>
        /// 获取当前生效的企业设置文件物理存储绝对路径
        /// </summary>
        private string GetCurrentSettingsFilePath()
        {
            // 动态获取当前生效的数据目录 (优先取自定义目录，兜底取默认目录)
            string appDataDir = Tool.GetAppDataDirectory();
            // 拼接 EnterpriseSettings.json 的完整保存路径
            return Path.Combine(appDataDir, SettingsFileName);
        }

        /// <summary>
        /// 异步从本地磁盘加载企业设置数据，若不存在则返回默认初始值
        /// </summary>
        public async Task<EnterpriseSettingsData> LoadSettingsAsync()
        {
            // 在后台 Task 线程中异步读取文件，防止阻塞 UI 线程
            return await Task.Run(() =>
            {
                try
                {
                    // 获取当前配置文件绝对路径
                    string filePath = GetCurrentSettingsFilePath();

                    // 实例化默认实体
                    var settings = new EnterpriseSettingsData();

                    // 判断本地配置文件是否存在
                    if (File.Exists(filePath))
                    {
                        // 读取本地 JSON 文本内容
                        string jsonText = File.ReadAllText(filePath);
                        // 反序列化为 EnterpriseSettingsData 数据对象
                        var loaded = JsonSerializer.Deserialize<EnterpriseSettingsData>(jsonText, JsonOptions);
                        if (loaded != null)
                        {
                            settings = loaded;
                        }
                    }

                    // 若模型中尚未记录自定义目录，尝试读取全局引导配置中的目录进行反显
                    if (string.IsNullOrWhiteSpace(settings.CustomDataDirectory))
                    {
                        // 读取全局引导配置中配置的路径
                        string globalCustomDir = Tool.GetCustomDataDirectoryFromGlobalConfig();
                        // 赋值反显
                        settings.CustomDataDirectory = globalCustomDir;
                    }

                    // 返回最终设置对象
                    return settings;
                }
                catch
                {
                    // 发生读取或解析异常时容错返回默认数据
                    return new EnterpriseSettingsData();
                }
            });
        }

        /// <summary>
        /// 异步将最新的企业设置序列化保存至本地磁盘
        /// </summary>
        public async Task<bool> SaveSettingsAsync(EnterpriseSettingsData settings)
        {
            // 在 Task 后台线程完成序列化与磁盘写入
            return await Task.Run(() =>
            {
                try
                {
                    // 1. 同步将用户在界面中配置的自定义目录更新持久化至全局引导文件
                    Tool.SetCustomDataDirectory(settings.CustomDataDirectory);

                    // 2. 重新获取更新后生效的目标保存路径 (若切换了目录则写入新目录)
                    string filePath = GetCurrentSettingsFilePath();

                    // 3. 将设置对象序列化为 JSON 格式化字符串
                    string jsonText = JsonSerializer.Serialize(settings, JsonOptions);

                    // 4. 将最新的 JSON 文本覆盖写入文件
                    File.WriteAllText(filePath, jsonText);

                    // 写入成功返回 true
                    return true;
                }
                catch
                {
                    // 写入失败捕获异常并返回 false
                    return false;
                }
            });
        }
    }
}
