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

        // 存储 说明1 ~ 说明5 的各自多行文本内容字典
        public System.Collections.Generic.Dictionary<string, string> QuoteTemplates { get; set; } = GetDefaultQuoteTemplates();

        /// <summary>
        /// 获取 5 套默认报价说明模板内容字典
        /// </summary>
        public static System.Collections.Generic.Dictionary<string, string> GetDefaultQuoteTemplates()
        {
            // 返回预设的 5 套成套设备专业报价说明范本
            return new System.Collections.Generic.Dictionary<string, string>
            {
                // 说明1：标准质保与交付
                { "说明1", "1. 本项目报价有效期为30天。\n2. 设备到货后提供免费指导安装调试与技术培训。\n3. 设备质保期为工程整体验收合格后12个月。" },
                // 说明2：台数与总价动态参数范本
                { "说明2", "1. 本项目报价包含[箱柜数量]台成套设备，项目总价：[项目总价]元（大写：[大写总价]）。\n2. 付款方式：预付30%，发货前付60%，质保金10%。\n3. 交货周期：合同签订并技术确认后25个工作日。" },
                // 说明3：元件与品牌技术规范
                { "说明3", "1. 报价依据设计图纸及技术规范配置核算，关键元器件：[器件名称]。\n2. 主要元器件采用[器件厂家]原厂正品标准件，随箱提供合格证与出厂检测报告。\n3. 现场接线与电缆敷设由采购方负责。" },
                // 说明4：柜体与制造工艺标准
                { "说明4", "1. 设备制造遵循 GB/T 7251 及国家低压成套开关设备现行强制性标准。\n2. 柜体采用优质敷铝锌板/冷轧钢板制造，防护等级满足设计要求。\n3. 本报价包含市内标准运输费用，不含现场卸车吊装费。" },
                // 说明5：特殊补充条款
                { "说明5", "1. 补充说明：报价未包含不可抗力因素造成的工期延误与费用增加。\n2. 如需加装智能监控仪表或通讯模块，按实际增加硬件及调试成本另行结算。" }
            };
        }

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
        /// 同步直接从本地磁盘加载企业设置数据，杜绝 Task 上下文死锁，供 Excel 主线程直接调用
        /// </summary>
        /// <returns>企业设置实体对象</returns>
        public static EnterpriseSettingsData LoadSettingsDirect()
        {
            try
            {
                // 获取数据目录物理路径
                string appDataDir = Tool.GetAppDataDirectory();
                // 拼接 EnterpriseSettings.json 物理路径
                string filePath = Path.Combine(appDataDir, SettingsFileName);

                // 实例化默认实体
                var settings = new EnterpriseSettingsData();

                // 判断配置文件物理文件是否存在
                if (File.Exists(filePath))
                {
                    // 读取磁盘 JSON 字符串
                    string jsonText = File.ReadAllText(filePath);
                    // 驼峰命名反序列化
                    var loaded = JsonSerializer.Deserialize<EnterpriseSettingsData>(jsonText, JsonOptions);
                    if (loaded != null) settings = loaded;
                }

                // 防御性校验报价模版
                if (settings.QuoteTemplates == null || settings.QuoteTemplates.Count == 0)
                {
                    // 注入默认模版字典
                    settings.QuoteTemplates = EnterpriseSettingsData.GetDefaultQuoteTemplates();
                }

                // 返回最终结果
                return settings;
            }
            catch
            {
                // 发生读取异常时返回默认设置实体
                return new EnterpriseSettingsData();
            }
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

                    // 防御性校验：若旧版本数据中未包含报价模版字典，则自动初始化注入 5 套专业范本
                    if (settings.QuoteTemplates == null || settings.QuoteTemplates.Count == 0)
                    {
                        // 实例化默认 5 套模版字典
                        settings.QuoteTemplates = EnterpriseSettingsData.GetDefaultQuoteTemplates();
                    }

                    // 校验当前选中模版并保持当前编辑描述内容双向同步对齐
                    if (!string.IsNullOrWhiteSpace(settings.QuoteTemplate) && settings.QuoteTemplates.ContainsKey(settings.QuoteTemplate))
                    {
                        // 若旧数据存有非空描述且模版字典中为空，则用历史描述回填该模版
                        if (!string.IsNullOrWhiteSpace(settings.QuoteDescription))
                        {
                            // 回填当前模版内容
                            settings.QuoteTemplates[settings.QuoteTemplate] = settings.QuoteDescription;
                        }
                        else
                        {
                            // 否则将模版预设内容同步回显至当前编辑描述文本中
                            settings.QuoteDescription = settings.QuoteTemplates[settings.QuoteTemplate];
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
