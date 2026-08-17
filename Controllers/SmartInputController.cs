using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 智能输入 WebAPI 风格控制器，负责元器件去重数据存取、配置持久化与调度
    /// </summary>
    public class SmartInputController
    {
        // 元器件分类缓存 JSON 文件物理路径 (data/smart_components.json)
        private readonly string _componentsFilePath;

        // 智能填写配置 JSON 文件物理路径 (data/smart_input_config.json)
        private readonly string _configFilePath;

        // JSON 序列化与反序列化选项配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            PropertyNameCaseInsensitive = true
        };

        // 文件读写互斥锁
        private static readonly object _fileLock = new object();

        /// <summary>
        /// 构造函数: 初始化存储路径并确保 data 目录就绪
        /// </summary>
        public SmartInputController()
        {
            // 获取插件 data 数据目录
            string appDataDir = Tool.GetAppDataDirectory();

            // 拼接元器件缓存文件完整物理路径
            _componentsFilePath = Path.Combine(appDataDir, "smart_components.json");

            // 拼接配置文件完整物理路径
            _configFilePath = Path.Combine(appDataDir, "smart_input_config.json");
        }

        /// <summary>
        /// 获取当前保存的智能输入配置信息
        /// </summary>
        /// <returns>智能输入配置实体对象</returns>
        public SmartInputConfigModel GetConfig()
        {
            lock (_fileLock)
            {
                try
                {
                    // 若本地磁盘存在配置文件则直接读取
                    if (File.Exists(_configFilePath))
                    {
                        string json = File.ReadAllText(_configFilePath);
                        var config = JsonSerializer.Deserialize<SmartInputConfigModel>(json, JsonOptions);
                        if (config != null)
                        {
                            return config;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 记录读取配置异常
                    LogHelper.WriteLog($"读取 smart_input_config.json 失败: {ex.Message}");
                }

                // 若不存在或读取失败，构建默认配置
                var defaultConfig = new SmartInputConfigModel
                {
                    SelectedSheets = new List<string>(),
                    FillName = true,
                    FillManufacturer = true,
                    FillUnit = true,
                    FillUnitPrice = true,
                    AutoDropdownEnabled = true
                };

                // 保存默认配置至本地文件
                SaveConfig(defaultConfig);

                return defaultConfig;
            }
        }

        /// <summary>
        /// 持久化保存智能输入配置至本地磁盘
        /// </summary>
        /// <param name="config">配置实体对象</param>
        /// <returns>保存是否成功</returns>
        public bool SaveConfig(SmartInputConfigModel config)
        {
            lock (_fileLock)
            {
                try
                {
                    // 校验输入对象有效性
                    if (config == null) return false;

                    // 序列化配置为格式化 JSON
                    string json = JsonSerializer.Serialize(config, JsonOptions);

                    // 确保目标文件夹已创建
                    string? dir = Path.GetDirectoryName(_configFilePath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    // 写入配置文件
                    File.WriteAllText(_configFilePath, json);

                    return true;
                }
                catch (Exception ex)
                {
                    // 记录保存配置异常
                    LogHelper.WriteLog($"保存 smart_input_config.json 失败: {ex.Message}");
                    return false;
                }
            }
        }

        /// <summary>
        /// 从本地磁盘加载已存储的各表去重元器件数据集
        /// </summary>
        /// <returns>元器件存储结构对象</returns>
        public SmartComponentsStorage GetStoredComponents()
        {
            lock (_fileLock)
            {
                try
                {
                    // 若存在缓存文件则反序列化加载
                    if (File.Exists(_componentsFilePath))
                    {
                        string json = File.ReadAllText(_componentsFilePath);
                        var storage = JsonSerializer.Deserialize<SmartComponentsStorage>(json, JsonOptions);
                        if (storage != null && storage.Sheets != null)
                        {
                            return storage;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 记录读取元器件数据异常
                    LogHelper.WriteLog($"读取 smart_components.json 失败: {ex.Message}");
                }

                // 若无缓存返回空结构
                return new SmartComponentsStorage();
            }
        }

        /// <summary>
        /// 保存提取去重后的元器件数据集到本地磁盘
        /// </summary>
        /// <param name="storage">元器件存储数据根对象</param>
        /// <returns>是否保存成功</returns>
        public bool SaveComponents(SmartComponentsStorage storage)
        {
            lock (_fileLock)
            {
                try
                {
                    // 校验入参
                    if (storage == null) return false;

                    // 更新最后刷新时间
                    storage.LastUpdatedTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    // 序列化为 JSON 字符串
                    string json = JsonSerializer.Serialize(storage, JsonOptions);

                    // 确保目标路径存在
                    string? dir = Path.GetDirectoryName(_componentsFilePath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    // 写入数据文件
                    File.WriteAllText(_componentsFilePath, json);

                    return true;
                }
                catch (Exception ex)
                {
                    // 记录写入元器件缓存异常
                    LogHelper.WriteLog($"保存 smart_components.json 失败: {ex.Message}");
                    return false;
                }
            }
        }

        /// <summary>
        /// 核心方法：刷新并提取当前工作簿所有表的元器件，去重并持久化存储
        /// </summary>
        /// <returns>刷新后的元器件存储对象</returns>
        public SmartComponentsStorage RefreshAndExtract()
        {
            try
            {
                // 调用 Excel 核心业务服务类从工作簿读取所有表格数据
                var storage = ExcelServices.ExtractComponentsFromAllSheets();

                // 若提取到有效数据，将其持久化写入本地 JSON 文件
                if (storage != null)
                {
                    SaveComponents(storage);
                    return storage;
                }
            }
            catch (Exception ex)
            {
                // 记录刷新提取元器件失败日志
                LogHelper.WriteLog($"刷新提取元器件异常: {ex.Message}");
            }

            // 返回当前磁盘已有缓存兜底
            return GetStoredComponents();
        }

        /// <summary>
        /// 依据选中的工作表为当前活动表注入 C 列数据有效性（下拉列表）
        /// </summary>
        /// <param name="config">智能输入配置</param>
        /// <returns>操作结果提示消息与状态</returns>
        public (bool success, string message) ApplyDropdown(SmartInputConfigModel config)
        {
            try
            {
                // 校验配置
                if (config == null || config.SelectedSheets == null || config.SelectedSheets.Count == 0)
                {
                    return (false, "请先在上方勾选至少一个数据源工作表！");
                }

                // 保存当前勾选配置
                SaveConfig(config);

                // 读取已存储的所有表数据
                var storage = GetStoredComponents();

                // 筛选出选中的工作表
                var matchedSheets = storage.Sheets
                    .Where(s => config.SelectedSheets.Contains(s.SheetName))
                    .ToList();

                // 汇总所有被选中表的去重 C 列型号
                var modelList = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var sheet in matchedSheets)
                {
                    if (sheet.Components != null)
                    {
                        foreach (var comp in sheet.Components)
                        {
                            if (!string.IsNullOrWhiteSpace(comp.Model))
                            {
                                modelList.Add(comp.Model.Trim());
                            }
                        }
                    }
                }

                // 判断型号列表是否为空
                if (modelList.Count == 0)
                {
                    return (false, "所选工作表中未提取到任何有效的规格型号！");
                }

                // 调用 Excel 服务层为当前活动工作表注入下拉数据验证
                bool result = ExcelServices.ApplySmartDropdownToActiveSheet(modelList.OrderBy(m => m).ToList());

                if (result)
                {
                    return (true, $"成功为当前工作表各箱柜 C 列注入 {modelList.Count} 个规格型号下拉项！");
                }
                else
                {
                    return (false, "注入下拉列表失败，请确保当前处于有效的箱柜工作表中！");
                }
            }
            catch (Exception ex)
            {
                // 记录注入下拉异常
                LogHelper.WriteLog($"注入下拉列表异常: {ex.Message}");
                return (false, $"注入下拉列表发生异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 将选中的元器件条目回填至 Excel 当前选中活动行
        /// </summary>
        /// <param name="item">选中的元器件实体</param>
        /// <param name="config">回填字段配置</param>
        /// <returns>回填操作是否成功</returns>
        public bool FillToActiveRow(SmartComponentItem item, SmartInputConfigModel config)
        {
            try
            {
                // 校验元器件对象
                if (item == null) return false;

                // 若配置为空则读取默认配置
                if (config == null) config = GetConfig();

                // 调用 Excel 服务层执行回填
                return ExcelServices.FillComponentToActiveRow(item, config);
            }
            catch (Exception ex)
            {
                // 记录回填异常
                LogHelper.WriteLog($"回填元器件至活动行异常: {ex.Message}");
                return false;
            }
        }
    }
}
