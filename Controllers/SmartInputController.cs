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

        // 静态内存配置缓存，避免高频 SelectionChange 频繁读取磁盘文件
        private static SmartInputConfigModel? _cachedConfig = null;

        // 静态元器件数据集内存缓存，加速物料候选检索
        private static SmartComponentsStorage? _cachedStorage = null;

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
        /// 主动使内存缓存失效，强制下次重新读取磁盘
        /// </summary>
        public static void InvalidateCache()
        {
            // 清空静态配置缓存
            _cachedConfig = null;
            // 清空静态物料数据缓存
            _cachedStorage = null;
        }

        /// <summary>
        /// 获取当前保存的智能输入配置信息 (优先内存直出，耗时 0ms)
        /// </summary>
        /// <param name="forceReload">是否强制从磁盘重载</param>
        /// <returns>智能输入配置实体对象</returns>
        public SmartInputConfigModel GetConfig(bool forceReload = false)
        {
            // 若内存缓存有效且非强制重载，直接返回内存对象
            if (!forceReload && _cachedConfig != null)
            {
                return _cachedConfig;
            }

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
                            // 写入内存缓存加速后续读取
                            _cachedConfig = config;
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
                    AutoDropdownEnabled = true,
                    AutoPopupFloatWindow = true,
                    AutoLearnNewComponents = true
                };

                // 保存默认配置至本地文件
                SaveConfig(defaultConfig);

                // 更新内存缓存
                _cachedConfig = defaultConfig;
                return defaultConfig;
            }
        }

        /// <summary>
        /// 持久化保存智能输入配置至本地磁盘并同步更新内存缓存
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

                    // 同步刷新静态内存缓存，保障后续零延迟读取
                    _cachedConfig = config;

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
        /// 从本地磁盘加载已存储的各表去重元器件数据集 (优先内存直出)
        /// </summary>
        /// <param name="forceReload">是否强制从磁盘重载</param>
        /// <returns>元器件存储结构对象</returns>
        public SmartComponentsStorage GetStoredComponents(bool forceReload = false)
        {
            // 若内存缓存有效且非强制重载，直接返回内存对象
            if (!forceReload && _cachedStorage != null)
            {
                return _cachedStorage;
            }

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
                            // 写入内存缓存
                            _cachedStorage = storage;
                            return storage;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 记录读取元器件数据异常
                    LogHelper.WriteLog($"读取 smart_components.json 失败: {ex.Message}");
                }

                // 若无缓存返回空结构并缓存
                var emptyStorage = new SmartComponentsStorage();
                _cachedStorage = emptyStorage;
                return emptyStorage;
            }
        }

        /// <summary>
        /// 保存提取去重后的元器件数据集到本地磁盘并更新内存缓存
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

                    // 同步更新内存缓存
                    _cachedStorage = storage;

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

        // 后台异步防抖持久化定时器句柄，避免用户连续击键引起高频磁盘 IO
        private static System.Threading.Timer? _debounceSaveTimer = null;

        // 默认防抖延迟时间 (毫秒: 2500ms) --硬编码: 防抖持久化延迟--
        private const int DebounceDelayMilliseconds = 2500;

        /// <summary>
        /// 尝试将新录入的元器件条目增量学习并合并加入候选词库 (质量门控 + 内存秒级同步 + 异步落盘)
        /// </summary>
        /// <param name="newItem">新录入或改动的元器件数据项</param>
        /// <returns>是否成功学习或补充完善已有元器件</returns>
        public bool TryLearnComponent(SmartComponentItem newItem)
        {
            // 校验待学习项非空
            if (newItem == null) return false;

            // 1. 检查配置是否开启了自动学习功能
            var config = GetConfig();
            // 若用户在配置中关闭了自动学习，直接退出
            if (config == null || !config.AutoLearnNewComponents) return false;

            // 2. 质量门控 (Gate 1): 规格型号非空且长度至少 2 个有效字符
            string model = (newItem.Model ?? string.Empty).Trim();
            // 过滤过短字符
            if (model.Length < 2) return false;

            // 质量门控 (Gate 2): 排除临时占位符、小计、合计及特殊字符
            string upperModel = model.ToUpperInvariant();
            // 排除关键词数组 --硬编码: 排除无效型号关键字--
            string[] forbiddenKeywords = new[] { "小计", "合计", "总计", "汇总", "待定", "暂估", "点击查询", "--", "==", "..." };
            // 遍历检查是否包含非法关键字
            foreach (var kw in forbiddenKeywords)
            {
                // 命中非法关键字直接拦截
                if (upperModel.Contains(kw.ToUpperInvariant())) return false;
            }

            // 3. 提取同行辅助属性 (若有则记录，无名称或单价也完全允许入库，后续可增量自愈补齐)
            string name = (newItem.Name ?? string.Empty).Trim();
            // 提取生产厂家
            string mfr = (newItem.Manufacturer ?? string.Empty).Trim();
            // 提取计量单位
            string unit = (newItem.Unit ?? string.Empty).Trim();
            // 提取销售单价
            decimal price = newItem.UnitPrice;

            // 4. 内存候选集自愈与合并
            lock (_fileLock)
            {
                try
                {
                    // 获取当前已有的全部元器件存储集合 (优先内存)
                    var storage = GetStoredComponents();
                    // 兜底防御空对象
                    if (storage == null) storage = new SmartComponentsStorage();
                    // 兜底防御工作表列表
                    if (storage.Sheets == null) storage.Sheets = new List<SheetComponentData>();

                    // 确定目标工作表名称
                    string targetSheetName = string.IsNullOrWhiteSpace(newItem.SheetName) ? "常用物料" : newItem.SheetName;

                    // 寻找或新建目标工作表数据包
                    var targetSheetData = storage.Sheets.FirstOrDefault(s => string.Equals(s.SheetName, targetSheetName, StringComparison.OrdinalIgnoreCase));
                    if (targetSheetData == null)
                    {
                        // 若不存在则新增该表数据包
                        targetSheetData = new SheetComponentData
                        {
                            SheetName = targetSheetName,
                            Components = new List<SmartComponentItem>()
                        };
                        // 追加至全局数据集
                        storage.Sheets.Add(targetSheetData);
                    }

                    // 兜底组件列表非空
                    if (targetSheetData.Components == null)
                    {
                        targetSheetData.Components = new List<SmartComponentItem>();
                    }

                    // 检查该表中是否已有同型号的元器件
                    var existingItem = targetSheetData.Components.FirstOrDefault(c => string.Equals(c.Model, model, StringComparison.OrdinalIgnoreCase));

                    // 标记是否发生了数据变化
                    bool hasChanged = false;
                    if (existingItem != null)
                    {
                        // 若已有条目，执行属性增量自愈补齐（若新输入提供了更全的属性则补充）
                        if (string.IsNullOrEmpty(existingItem.Name) && !string.IsNullOrEmpty(name))
                        {
                            existingItem.Name = name;
                            hasChanged = true;
                        }
                        // 补充生产厂家
                        if (string.IsNullOrEmpty(existingItem.Manufacturer) && !string.IsNullOrEmpty(mfr))
                        {
                            existingItem.Manufacturer = mfr;
                            hasChanged = true;
                        }
                        // 补充计量单位
                        if (string.IsNullOrEmpty(existingItem.Unit) && !string.IsNullOrEmpty(unit))
                        {
                            existingItem.Unit = unit;
                            hasChanged = true;
                        }
                        // 补充单价
                        if (existingItem.UnitPrice <= 0 && price > 0)
                        {
                            existingItem.UnitPrice = price;
                            hasChanged = true;
                        }
                    }
                    else
                    {
                        // 若为全新型号，构建新实体并加入集合
                        var itemToAdd = new SmartComponentItem
                        {
                            Model = model,
                            Name = name,
                            Manufacturer = mfr,
                            Unit = unit,
                            UnitPrice = price,
                            Category = newItem.Category ?? string.Empty,
                            SheetName = targetSheetName,
                            CabinetNo = newItem.CabinetNo ?? "新增学习"
                        };
                        // 加入列表
                        targetSheetData.Components.Add(itemToAdd);
                        // 更新数量统计
                        targetSheetData.UniqueCount = targetSheetData.Components.Count;
                        targetSheetData.TotalCount++;
                        hasChanged = true;
                    }

                    // 4. 若发生了数据变更，即时刷新内存缓存并安排后台防抖落盘
                    if (hasChanged)
                    {
                        // 保证该工作表包含在配置的已勾选数据源中，确保立刻可被后续单元格联想检索
                        if (config.SelectedSheets != null && !config.SelectedSheets.Contains(targetSheetName))
                        {
                            config.SelectedSheets.Add(targetSheetName);
                            // 异步保存配置
                            SaveConfig(config);
                        }

                        // 立即更新静态内存缓存（后续单元格联想 0ms 瞬间命中）
                        _cachedStorage = storage;

                        // 触发后台异步防抖持久化
                        ScheduleDebouncedSave(storage);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    // 记录学习异常日志
                    LogHelper.WriteLog($"TryLearnComponent 增量学习异常: {ex.Message}");
                }
            }

            return false;
        }

        /// <summary>
        /// 安排后台异步防抖落盘保存，避免频繁打字引发高频磁盘 IO
        /// </summary>
        /// <param name="storage">待持久化的元器件存储数据对象</param>
        private void ScheduleDebouncedSave(SmartComponentsStorage storage)
        {
            try
            {
                // 若定时器未初始化则新建
                if (_debounceSaveTimer == null)
                {
                    // 初始化定时器并在 2.5 秒后触发一次
                    _debounceSaveTimer = new System.Threading.Timer(OnDebounceTimerFired, storage, DebounceDelayMilliseconds, System.Threading.Timeout.Infinite);
                }
                else
                {
                    // 刷新定时器，重新延后 2.5 秒执行
                    _debounceSaveTimer.Change(DebounceDelayMilliseconds, System.Threading.Timeout.Infinite);
                }
            }
            catch (Exception ex)
            {
                // 记录调度异常
                LogHelper.WriteLog($"安排防抖持久化异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 防抖定时器触发回调: 在后台线程安全将元器件缓存写入本地磁盘
        /// </summary>
        /// <param name="state">状态对象 (SmartComponentsStorage)</param>
        private void OnDebounceTimerFired(object? state)
        {
            try
            {
                // 类型安全转换为数据存储对象
                if (state is SmartComponentsStorage storage)
                {
                    // 异步持久化写入本地磁盘 smart_components.json
                    SaveComponents(storage);
                }
            }
            catch (Exception ex)
            {
                // 记录落盘异常日志
                LogHelper.WriteLog($"防抖定时器执行保存异常: {ex.Message}");
            }
        }
    }
}
