using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 公式组数据模型结构定义
    /// </summary>
    public class FormulaGroupModel
    {
        // 公式组唯一标识 ID
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        // 公式组名称 (例如: 简易费用公式、多费用公式、其它样式费用公式等)
        public string Name { get; set; } = string.Empty;

        // 是否为系统默认公式组 (图标显示黄金钥匙标志，系统默认不可删除)
        public bool IsSystemDefault { get; set; } = false;

        // 是否被用户设为当前激活的默认公式组
        public bool IsDefault { get; set; } = false;

        // 公式组包含的具体表格明细行数据列表
        public List<FormulaItemModel> Details { get; set; } = new List<FormulaItemModel>();
    }

    /// <summary>
    /// 公式组包含的具体表格明细行数据模型
    /// </summary>
    public class FormulaItemModel
    {
        // 行序号 (如 1, 2, [序号], 总计)
        public string No { get; set; } = string.Empty;

        // 元件/项目名称 (如 小计, 管理费, 利润, 税金, 单台合计)
        public string Name { get; set; } = string.Empty;

        // 型号规格
        public string Model { get; set; } = string.Empty;

        // 生产厂家
        public string Manufacturer { get; set; } = string.Empty;

        // 单位 (如 台, 套)
        public string Unit { get; set; } = string.Empty;

        // 数量计算表达式或数值
        public string Quantity { get; set; } = string.Empty;

        // 单价计算表达式或数值
        public string Price { get; set; } = string.Empty;

        // 总价计算公式 (例如 =ROUND(H2*0.12, 2) 或 =ROUND(SUM(H2:H3)*0.15, 2))
        public string TotalPriceFormula { get; set; } = string.Empty;

        // 成本单价公式或数值
        public string CostPrice { get; set; } = string.Empty;

        // 成本总价公式 (例如 =ROUND(SUM(K2:K5), 2))
        public string CostTotalPriceFormula { get; set; } = string.Empty;

        // 项目类别 (如 费用)
        public string Category { get; set; } = string.Empty;
    }

    /// <summary>
    /// 调费请求参数实体
    /// </summary>
    public class ApplyFormulaRequest
    {
        // 目标调费范围: "currentCabinet"(当前箱柜), "currentCategory"(当前分类), "allCabinets"(所有箱柜), "selectedCabinet"(选择箱柜)
        public string TargetScope { get; set; } = "currentCabinet";

        // 当前选择的公式组名称
        public string GroupName { get; set; } = string.Empty;
    }

    /// <summary>
    /// 公式法调费全局 JSON 配置文件数据根结构
    /// </summary>
    public class FormulaFeeConfig
    {
        // 所有的公式组集合列表
        public List<FormulaGroupModel> Groups { get; set; } = new List<FormulaGroupModel>();
    }

    /// <summary>
    /// 公式法调费 WebAPI 风格控制器，负责公式模板获取、调度与 JSON 存盘处理
    /// </summary>
    public class FormulaAdjustFeeController
    {
        // 配置文件物理保存路径
        private readonly string _configFilePath;

        // JSON 序列化与反序列化选项缓存对象
        private static JsonSerializerOptions? _jsonOptions;
        // 属性获取器：支持懒加载与依赖异常降级兜底，避免静态字段初始化引发 TypeInitializationException
        private static JsonSerializerOptions JsonOptions
        {
            get
            {
                if (_jsonOptions == null)
                {
                    try
                    {
                        // 优先使用带非转义字符编码器的配置
                        _jsonOptions = new JsonSerializerOptions
                        {
                            WriteIndented = true,
                            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                            PropertyNameCaseInsensitive = true
                        };
                    }
                    catch
                    {
                        // 降级兜底使用基础配置
                        _jsonOptions = new JsonSerializerOptions
                        {
                            WriteIndented = true,
                            PropertyNameCaseInsensitive = true
                        };
                    }
                }
                return _jsonOptions;
            }
        }

        // 线程安全互斥锁
        private static readonly object _fileLock = new object();

        // 内存中缓存的公式组列表
        private List<FormulaGroupModel> _formulaGroups;

        /// <summary>
        /// 构造函数：初始化配置文件存储路径并从 JSON 文件加载
        /// </summary>
        public FormulaAdjustFeeController()
        {
            // 通过 Tool 工具类获取 AppData 插件专属目录
            string appDataDir = Tool.GetAppDataDirectory();
            // 拼接得到 formula_fee_settings.json 存储文件路径
            _configFilePath = Path.Combine(appDataDir, "formula_fee_settings.json");

            // 从本地磁盘加载配置数据
            _formulaGroups = LoadConfigFromDisk();
        }

        /// <summary>
        /// 从本地磁盘加载公式组配置文件
        /// </summary>
        private List<FormulaGroupModel> LoadConfigFromDisk()
        {
            lock (_fileLock)
            {
                try
                {
                    // 若 AppData 中存在配置文件则优先读取
                    if (File.Exists(_configFilePath))
                    {
                        // 读取磁盘文件中的 JSON 文本
                        string json = File.ReadAllText(_configFilePath);
                        // 反序列化为配置对象
                        var config = JsonSerializer.Deserialize<FormulaFeeConfig>(json, JsonOptions);
                        // 若反序列化成功且包含有效分组
                        if (config != null && config.Groups != null && config.Groups.Count > 0)
                        {
                            // 返回加载的公式组集合
                            return config.Groups;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 记录加载失败异常日志
                    LogHelper.WriteLog($"加载 formula_fee_settings.json 失败: {ex.Message}");
                }

                // 若磁盘无文件或解析失败，生成默认内置公式组配置
                var defaultGroups = CreateDefaultFormulaGroups();
                // 立即持久化保存至本地磁盘
                SaveConfigToDisk(defaultGroups);
                // 返回默认配置项
                return defaultGroups;
            }
        }

        /// <summary>
        /// 将当前所有公式组持久化写入本地 JSON 文件
        /// </summary>
        public bool SaveConfigToDisk(List<FormulaGroupModel> groups)
        {
            lock (_fileLock)
            {
                try
                {
                    // 更新内部内存缓存引用
                    _formulaGroups = groups;
                    // 构造待序列化的配置根对象
                    var config = new FormulaFeeConfig { Groups = groups };
                    // 序列化为格式化 JSON 字符串
                    string json = JsonSerializer.Serialize(config, JsonOptions);
                    // 确保目标目录存在
                    string? dir = Path.GetDirectoryName(_configFilePath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    {
                        // 递归创建所在文件夹
                        Directory.CreateDirectory(dir);
                    }
                    // 将 JSON 内容安全写入物理文件
                    File.WriteAllText(_configFilePath, json);
                    // 返回写入成功
                    return true;
                }
                catch (Exception ex)
                {
                    // 记录写入错误日志
                    LogHelper.WriteLog($"保存 formula_fee_settings.json 失败: {ex.Message}");
                    // 返回写入失败
                    return false;
                }
            }
        }

        /// <summary>
        /// 后端 WebAPI 接口: 获取所有公式组列表 (包含其绑定的明细项)
        /// </summary>
        /// <returns>公式组集合</returns>
        public List<FormulaGroupModel> GetFormulaGroups()
        {
            // 返回当前加载的所有公式组列表
            return _formulaGroups;
        }

        /// <summary>
        /// 后端 WebAPI 接口: 获取当前激活的默认公式组
        /// </summary>
        /// <returns>默认公式组对象</returns>
        public FormulaGroupModel GetDefaultGroup()
        {
            // 查找标记为默认的公式组
            var def = _formulaGroups?.FirstOrDefault(g => g.IsDefault);
            // 兜底返回第一个公式组或内置默认组
            return def ?? _formulaGroups?.FirstOrDefault() ?? CreateDefaultFormulaGroups().First();
        }

        /// <summary>
        /// 后端 WebAPI 接口: 保存全部公式组数据 (由前端编辑后整体提交)
        /// </summary>
        public bool SaveAllFormulaGroups(List<FormulaGroupModel> groups)
        {
            // 校验入参合法性
            if (groups == null || groups.Count == 0) return false;
            // 写入本地磁盘持久化
            return SaveConfigToDisk(groups);
        }

        /// <summary>
        /// 后端 WebAPI 接口: 设置指定公式组为默认公式组
        /// </summary>
        /// <param name="groupId">公式组 ID</param>
        /// <returns>操作是否成功</returns>
        public bool SetDefaultFormulaGroup(string groupId)
        {
            // 遍历所有公式组，重置 IsDefault 状态
            foreach (var g in _formulaGroups)
            {
                // 若 ID 匹配则设为默认，否则置为 false
                g.IsDefault = (g.Id == groupId);
            }
            // 立即持久化保存更新
            return SaveConfigToDisk(_formulaGroups);
        }

        /// <summary>
        /// 后端 WebAPI 接口: 复制指定的公式组及其明细列表
        /// </summary>
        /// <param name="groupId">源公式组 ID</param>
        /// <returns>新建的公式组对象</returns>
        public FormulaGroupModel CopyFormulaGroup(string groupId)
        {
            // 查找对应的源公式组
            var source = _formulaGroups.FirstOrDefault(g => g.Id == groupId);
            // 若源不存在则返回空
            if (source == null) return null!;

            // 深度克隆其包含的明细行集合
            var clonedDetails = new List<FormulaItemModel>();
            if (source.Details != null)
            {
                // 循环克隆每一个公式明细行
                foreach (var item in source.Details)
                {
                    clonedDetails.Add(new FormulaItemModel
                    {
                        No = item.No,
                        Name = item.Name,
                        Model = item.Model,
                        Manufacturer = item.Manufacturer,
                        Unit = item.Unit,
                        Quantity = item.Quantity,
                        Price = item.Price,
                        TotalPriceFormula = item.TotalPriceFormula,
                        CostPrice = item.CostPrice,
                        CostTotalPriceFormula = item.CostTotalPriceFormula,
                        Category = item.Category
                    });
                }
            }

            // 构造新的复制实体 --名称加上副本后缀--
            var newGroup = new FormulaGroupModel
            {
                // 生成新唯一标识
                Id = Guid.NewGuid().ToString("N"),
                // 名称加上副本后缀
                Name = $"{source.Name}_副本",
                // 新副本非系统默认
                IsSystemDefault = false,
                // 非默认激活
                IsDefault = false,
                // 赋予深克隆后的明细项
                Details = clonedDetails
            };

            // 追加至全局集合中
            _formulaGroups.Add(newGroup);
            // 持久化保存
            SaveConfigToDisk(_formulaGroups);
            // 返回新创建的公式组对象
            return newGroup;
        }

        /// <summary>
        /// 后端 WebAPI 接口: 删除指定的公式组
        /// </summary>
        /// <param name="groupId">要删除的公式组 ID</param>
        /// <returns>操作结果</returns>
        public bool DeleteFormulaGroup(string groupId)
        {
            // 查找对应公式组
            var target = _formulaGroups.FirstOrDefault(g => g.Id == groupId);
            // 系统默认公式组不允许删除
            if (target == null || target.IsSystemDefault)
            {
                // 返回删除失败
                return false;
            }

            // 从列表中移除目标公式组
            _formulaGroups.Remove(target);
            // 持久化写入磁盘
            SaveConfigToDisk(_formulaGroups);
            // 返回删除成功标志
            return true;
        }

        /// <summary>
        /// 后端 WebAPI 接口: 根据公式组名称获取对应的明细计算公式表
        /// </summary>
        /// <param name="groupName">公式名称</param>
        /// <returns>明细列表行数据集合</returns>
        public List<FormulaItemModel> GetFormulaDetails(string groupName)
        {
            // 根据名称查找对应的公式组
            var group = _formulaGroups.FirstOrDefault(g => g.Name == groupName);
            // 若命中且明细不为空
            if (group != null && group.Details != null && group.Details.Count > 0)
            {
                // 返回该公式组配置的明细表
                return group.Details;
            }

            // 若未找到，回退返回默认的“多费用公式”明细模板
            return CreateStandardMultiFeeDetails();
        }

        /// <summary>
        /// 创建系统内置的默认公式组集合
        /// </summary>
        private List<FormulaGroupModel> CreateDefaultFormulaGroups()
        {
            // 构造默认公式组列表
            return new List<FormulaGroupModel>
            {
                // 1. 简易费用公式 (系统默认，带钥匙图标)
                new FormulaGroupModel
                {
                    Id = "1",
                    Name = "简易费用公式",
                    IsSystemDefault = true,
                    IsDefault = false,
                    Details = CreateStandardSimpleFeeDetails()
                },
                // 2. 多费用公式 (默认激活)
                new FormulaGroupModel
                {
                    Id = "2",
                    Name = "多费用公式",
                    IsSystemDefault = false,
                    IsDefault = true,
                    Details = CreateStandardMultiFeeDetails()
                },
                // 3. 其它样式费用公式
                new FormulaGroupModel
                {
                    Id = "3",
                    Name = "其它样式费用公式",
                    IsSystemDefault = false,
                    IsDefault = false,
                    Details = CreateStandardOtherFeeDetails()
                },
                // 4. 国网报价费用公式
                new FormulaGroupModel
                {
                    Id = "4",
                    Name = "国网报价费用公式",
                    IsSystemDefault = false,
                    IsDefault = false,
                    Details = CreateStandardStateGridFeeDetails()
                },
                // 5. 人工辅料定额公式
                new FormulaGroupModel
                {
                    Id = "5",
                    Name = "人工辅料定额公式",
                    IsSystemDefault = false,
                    IsDefault = false,
                    Details = CreateStandardLaborMaterialFeeDetails()
                }
            };
        }

        /// <summary>
        /// 创建标准“多费用公式”明细列表 (总计行在 A 列为总计)
        /// </summary>
        private static List<FormulaItemModel> CreateStandardMultiFeeDetails()
        {
            return new List<FormulaItemModel>
            {
                // 1. 小计行
                new FormulaItemModel
                {
                    No = "[序号]",
                    Name = "小计",
                    Model = "",
                    Manufacturer = "",
                    Unit = "",
                    Quantity = "",
                    Price = "",
                    TotalPriceFormula = "[总价小计]",
                    CostPrice = "",
                    CostTotalPriceFormula = "[成本总价小计]",
                    Category = ""
                },
                // 2. 管理费
                new FormulaItemModel
                {
                    No = "[序号]",
                    Name = "管理费",
                    Model = "",
                    Manufacturer = "",
                    Unit = "",
                    Quantity = "",
                    Price = "",
                    TotalPriceFormula = "=ROUND(H1*0.12, 2)",
                    CostPrice = "",
                    CostTotalPriceFormula = "",
                    Category = "费用"
                },
                // 3. 利润
                new FormulaItemModel
                {
                    No = "[序号]",
                    Name = "利润",
                    Model = "",
                    Manufacturer = "",
                    Unit = "",
                    Quantity = "",
                    Price = "",
                    TotalPriceFormula = "=ROUND(SUM(H1:H2)*0.15, 2)",
                    CostPrice = "",
                    CostTotalPriceFormula = "",
                    Category = "费用"
                },
                // 4. 税金
                new FormulaItemModel
                {
                    No = "[序号]",
                    Name = "税金",
                    Model = "",
                    Manufacturer = "",
                    Unit = "",
                    Quantity = "",
                    Price = "",
                    TotalPriceFormula = "=ROUND(SUM(H1:H3)*0.13, 2)",
                    CostPrice = "",
                    CostTotalPriceFormula = "",
                    Category = "费用"
                },
                // 5. 单台合计
                new FormulaItemModel
                {
                    No = "[序号]",
                    Name = "单台合计",
                    Model = "",
                    Manufacturer = "",
                    Unit = "",
                    Quantity = "",
                    Price = "",
                    TotalPriceFormula = "=ROUND(SUM(H1:H4), 2)",
                    CostPrice = "",
                    CostTotalPriceFormula = "=ROUND(SUM(K1:K4), 2)",
                    Category = ""
                },
                // 6. 总计 (A 列写“总计”，B 列为空)
                new FormulaItemModel
                {
                    No = "总计",
                    Name = "",
                    Model = "",
                    Manufacturer = "",
                    Unit = "台",
                    Price = "=ROUND(H5, 2)",
                    TotalPriceFormula = "=ROUND(F6*G6, 2)",
                    CostPrice = "",
                    CostTotalPriceFormula = "=ROUND(K5*F6, 2)",
                    Category = ""
                }
            };
        }

        /// <summary>
        /// 创建“简易费用公式”明细列表
        /// </summary>
        private static List<FormulaItemModel> CreateStandardSimpleFeeDetails()
        {
            return new List<FormulaItemModel>
            {
                // 小计
                new FormulaItemModel { No = "[序号]", Name = "小计", TotalPriceFormula = "[总价小计]", CostTotalPriceFormula = "[成本总价小计]" },
                // 综合费率
                new FormulaItemModel { No = "[序号]", Name = "综合成套费", TotalPriceFormula = "=ROUND(H1*0.25, 2)", Category = "费用" },
                // 单台合计
                new FormulaItemModel { No = "[序号]", Name = "单台合计", TotalPriceFormula = "=ROUND(SUM(H1:H2), 2)", CostTotalPriceFormula = "=ROUND(SUM(K1:K2), 2)" },
                // 总计行
                new FormulaItemModel { No = "总计", Name = "", Unit = "台", Price = "=ROUND(H3, 2)", TotalPriceFormula = "=ROUND(F4*G4, 2)", CostTotalPriceFormula = "=ROUND(K3*F4, 2)" }
            };
        }

        /// <summary>
        /// 创建“其它样式费用公式”明细列表
        /// </summary>
        private static List<FormulaItemModel> CreateStandardOtherFeeDetails()
        {
            return CreateStandardMultiFeeDetails();
        }

        /// <summary>
        /// 创建“国网报价费用公式”明细列表
        /// </summary>
        private static List<FormulaItemModel> CreateStandardStateGridFeeDetails()
        {
            return CreateStandardMultiFeeDetails();
        }

        /// <summary>
        /// 创建“人工辅料定额公式”明细列表
        /// </summary>
        private static List<FormulaItemModel> CreateStandardLaborMaterialFeeDetails()
        {
            return new List<FormulaItemModel>
            {
                // 小计
                new FormulaItemModel { No = "[序号]", Name = "小计", TotalPriceFormula = "[总价小计]", CostTotalPriceFormula = "[成本总价小计]" },
                // 人工定额
                new FormulaItemModel { No = "[序号]", Name = "人工费", TotalPriceFormula = "=ROUND(H2*0.08, 2)", Category = "费用" },
                // 辅料定额
                new FormulaItemModel { No = "[序号]", Name = "辅料费", TotalPriceFormula = "=ROUND(H2*0.05, 2)", Category = "费用" },
                // 利润
                new FormulaItemModel { No = "[序号]", Name = "利润", TotalPriceFormula = "=ROUND(SUM(H2:H4)*0.10, 2)", Category = "费用" },
                // 税金
                new FormulaItemModel { No = "[序号]", Name = "税金", TotalPriceFormula = "=ROUND(SUM(H2:H5)*0.13, 2)", Category = "费用" },
                // 单台合计
                new FormulaItemModel { No = "[序号]", Name = "单台合计", TotalPriceFormula = "=ROUND(SUM(H2:H6), 2)", CostTotalPriceFormula = "=ROUND(SUM(K2:K6), 2)" },
                // 总计行
                new FormulaItemModel { No = "总计", Name = "", Unit = "台", Quantity = "=ROUND(H7, 2)", Price = "=ROUND(F8*G8, 2)", CostTotalPriceFormula = "=ROUND(K7*F8, 2)" }
            };
        }
    }
}
