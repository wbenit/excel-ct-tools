using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 二次元件组规则管道后端控制器 (负责配置持久化、沙盒调试与 Excel 批量生成)
    /// </summary>
    public class ComponentGroupBuilderController
    {
        // 存储二次元件组配置文件的绝对路径
        private readonly string _configFilePath;

        // 全局 JSON 序列化选项 (启用驼峰命名与格式化输出)
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化配置文件路径 (存储在 AppData 目录中)
        /// </summary>
        public ComponentGroupBuilderController()
        {
            // 获取数据保存基目录
            string appDataDir = Tool.GetAppDataDirectory();
            // 拼接 ComponentGroupRules.json 保存路径
            _configFilePath = Path.Combine(appDataDir, "ComponentGroupRules.json");
        }

        /// <summary>
        /// 加载配置: 从本地磁盘加载或初始化默认出厂规则库
        /// </summary>
        public ComponentGroupConfig LoadConfig()
        {
            try
            {
                // 检查本地配置文件是否存在
                if (!File.Exists(_configFilePath))
                {
                    // 若不存在则生成出厂默认配置并持久化写入
                    var defaultConfig = ComponentGroupConfig.CreateDefault();
                    SaveConfig(defaultConfig);
                    return defaultConfig;
                }

                // 读取 JSON 文件内容
                string jsonText = File.ReadAllText(_configFilePath);
                var config = JsonSerializer.Deserialize<ComponentGroupConfig>(jsonText, JsonOptions);

                // 若反序列化有效则返回，否则返回默认配置
                return config ?? ComponentGroupConfig.CreateDefault();
            }
            catch
            {
                // 异常兜底返回默认配置
                return ComponentGroupConfig.CreateDefault();
            }
        }

        /// <summary>
        /// 保存配置至本地 JSON 文件
        /// </summary>
        public bool SaveConfig(ComponentGroupConfig config)
        {
            try
            {
                if (config == null) return false;

                // 保存前保证每条规则的 RawExpression 包含最新只读文字摘要 (供列表概览)
                if (config.Rules != null)
                {
                    foreach (var rule in config.Rules)
                    {
                        if (rule.ConditionTree != null)
                        {
                            // 自动生成单行概括摘要
                            rule.RawExpression = PipelineCompiler.BuildRuleSummary(rule);
                        }
                    }
                }

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
        /// 恢复出厂默认规则库
        /// </summary>
        public ComponentGroupConfig ResetToDefault()
        {
            // 构造出厂默认配置
            var defaultConfig = ComponentGroupConfig.CreateDefault();
            // 保存至本地磁盘
            SaveConfig(defaultConfig);
            // 返回默认对象
            return defaultConfig;
        }

        /// <summary>
        /// 从 Excel 中抓取当前选中箱柜的真实元件数据 (供前端沙盒测试)
        /// </summary>
        public List<EleComponentDto> GetActiveCabinetComponents(ComponentGroupConfig config)
        {
            var activeConfig = config ?? LoadConfig();
            return ExcelServices.GetActiveCabinetComponentsFromExcel(activeConfig);
        }

        /// <summary>
        /// 沙盒测试规则管道匹配
        /// </summary>
        public PipelineTestResultDto RunSandboxTest(ComponentGroupConfig config, List<EleComponentDto> components)
        {
            var activeConfig = config ?? LoadConfig();
            return ExcelServices.RunSandboxPipelineTest(activeConfig, components);
        }

        /// <summary>
        /// 执行批量二次元件组生成至 Excel
        /// </summary>
        public BatchGroupResultDto ExecuteBatch(ComponentGroupConfig config, bool activeCabinetOnly)
        {
            // 若传入配置，先保存配置
            if (config != null)
            {
                SaveConfig(config);
            }

            var activeConfig = config ?? LoadConfig();
            return ExcelServices.ExecuteBatchComponentGroup(activeConfig, activeCabinetOnly);
        }
    }
}
