using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 元器件型号参数提取后端 WebAPI 控制器
    /// </summary>
    public class ModelParamParserController
    {
        // 存储型号识别配置文件的绝对路径
        private readonly string _configFilePath;

        // 全局 JSON 序列化选项 (启用驼峰命名与格式化输出)
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            // 启用驼峰命名转换以适配前端 Vue 3
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            // 启用属性名称忽略大小写
            PropertyNameCaseInsensitive = true,
            // 启用格式化缩进输出
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化配置文件路径 (存储在插件 AppData/data 目录中)
        /// </summary>
        public ModelParamParserController()
        {
            // 获取数据保存基目录
            string appDataDir = Tool.GetAppDataDirectory();
            // 拼接 ModelParserConfig.json 保存路径
            _configFilePath = Path.Combine(appDataDir, "ModelParserConfig.json");
        }

        /// <summary>
        /// 获取配置: 从磁盘加载或初始化默认配置
        /// </summary>
        public ModelParserConfig LoadConfig()
        {
            try
            {
                // 检查本地配置文件是否存在
                if (!File.Exists(_configFilePath))
                {
                    // 若不存在则生成出厂默认配置并持久化写入
                    var defaultConfig = ModelParserConfig.CreateDefault();
                    SaveConfig(defaultConfig);
                    return defaultConfig;
                }

                // 读取 JSON 文件内容
                string jsonText = File.ReadAllText(_configFilePath);
                var config = JsonSerializer.Deserialize<ModelParserConfig>(jsonText, JsonOptions);

                // 若反序列化有效则返回，否则返回默认配置
                return config ?? ModelParserConfig.CreateDefault();
            }
            catch
            {
                // 异常兜底返回默认配置
                return ModelParserConfig.CreateDefault();
            }
        }

        /// <summary>
        /// 保存配置至本地 JSON 文件
        /// </summary>
        public bool SaveConfig(ModelParserConfig config)
        {
            try
            {
                if (config == null) return false;
                // 序列化配置为格式化 JSON 字符串
                string jsonText = JsonSerializer.Serialize(config, JsonOptions);
                // 写入磁盘文件
                File.WriteAllText(_configFilePath, jsonText);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 重置为出厂默认配置
        /// </summary>
        public ModelParserConfig ResetToDefault()
        {
            // 构造出厂默认配置
            var defaultConfig = ModelParserConfig.CreateDefault();
            // 保存至本地磁盘
            SaveConfig(defaultConfig);
            // 返回默认对象
            return defaultConfig;
        }

        /// <summary>
        /// 实时测试单条型号识别效果 (沙盒测试)
        /// </summary>
        public ParseResultDto TestParse(string rawModel, ModelParserConfig config)
        {
            // 若未传入配置则使用当前最新已保存配置
            var activeConfig = config ?? LoadConfig();
            // 调用公共服务层单条解析逻辑
            return ExcelServices.ParseSingleModel(rawModel, activeConfig);
        }

        /// <summary>
        /// 触发 Excel 批量识别回填
        /// </summary>
        public BatchParseResult ExecuteBatch(ModelParserConfig config)
        {
            // 若传入配置，先保存配置
            if (config != null)
            {
                SaveConfig(config);
            }

            var activeConfig = config ?? LoadConfig();
            // 调用公共服务层执行二维数组内存批量回填 (已集成写入第5行表头)
            return ExcelServices.ExecuteBatchModelParse(activeConfig);
        }
    }
}
