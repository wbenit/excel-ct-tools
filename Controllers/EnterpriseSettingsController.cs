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
    }

    /// <summary>
    /// 企业设置 Backend WebAPI/本地服务控制器
    /// </summary>
    public class EnterpriseSettingsController
    {
        // 存储企业配置 JSON 文件的本地绝对路径
        private readonly string _settingsFilePath;

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
        /// 构造函数：初始化本地存储目录与文件路径
        /// </summary>
        public EnterpriseSettingsController()
        {
            // 获取当前系统 AppData/Local 本地数据存放目录
            string appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAddInDemo");

            // 如果配置文件夹不存在，则自动创建该目录
            if (!Directory.Exists(appDataDir))
            {
                // 创建文件夹
                Directory.CreateDirectory(appDataDir);
            }

            // 拼接 EnterpriseSettings.json 的完整保存路径
            _settingsFilePath = Path.Combine(appDataDir, "EnterpriseSettings.json");
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
                    // 判断本地配置文件是否存在
                    if (!File.Exists(_settingsFilePath))
                    {
                        // 若不存在配置文件，则直接返回默认的企业配置示例数据
                        return new EnterpriseSettingsData();
                    }

                    // 读取本地 EnterpriseSettings.json 中的全部文本内容
                    string jsonText = File.ReadAllText(_settingsFilePath);

                    // 将读出的 JSON 字符串反序列化为 EnterpriseSettingsData 数据对象
                    var settings = JsonSerializer.Deserialize<EnterpriseSettingsData>(jsonText, JsonOptions);

                    // 若反序列化结果不为空则返回，否则返回默认新对象
                    return settings ?? new EnterpriseSettingsData();
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
                    // 将设置对象序列化为 JSON 格式的格式化字符串
                    string jsonText = JsonSerializer.Serialize(settings, JsonOptions);

                    // 将最新的 JSON 文本覆盖写入本地 EnterpriseSettings.json 文件
                    File.WriteAllText(_settingsFilePath, jsonText);

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
