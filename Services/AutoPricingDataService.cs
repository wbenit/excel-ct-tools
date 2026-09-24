using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// 自动组价方案数据服务 (支持云端 WebAPI 优先拉取与本地 SQLite 离线自愈双通道)
    /// </summary>
    public static class AutoPricingDataService
    {
        // 远程云端 WebAPI 基础地址 (可外部配置，默认本地或云端测试地址 --硬编码--)
        private const string DefaultApiBaseUrl = "http://localhost:5219"; // --硬编码-- 云端服务备用地址

        // 本地离线 SQLite 数据库文件相对或绝对路径 --硬编码--
        private const string LocalSqlitePath = @"d:\code\cad-net_1\ExWinner_Schemes.db"; // --硬编码-- 本地方案缓存库

        // 通用 HttpClient 实例
        private static readonly HttpClient _httpClient = CreateHttpClient();

        // 通用 JSON 反序列化设置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        // 创建带超时控制的 HTTP 客户端
        private static HttpClient CreateHttpClient()
        {
            var client = new HttpClient();
            // 设置超时时间为 2 秒，超时即快速降级至本地 SQLite，杜绝界面假死
            client.Timeout = TimeSpan.FromSeconds(2);
            return client;
        }

        /// <summary>
        /// 1. 获取完整的方案分类树 (双通道保障)
        /// </summary>
        public static List<AutoPricingCategoryDto> GetCategoryTree()
        {
            try
            {
                // 通道 A: 优先尝试从云端 WebAPI 拉取最新树形数据
                string url = $"{GetApiBaseUrl()}/api/Scheme/GetCategoryTree";
                var task = _httpClient.GetStringAsync(url);
                if (task.Wait(1500))
                {
                    string json = task.Result;
                    var tree = JsonSerializer.Deserialize<List<AutoPricingCategoryDto>>(json, JsonOptions);
                    if (tree != null && tree.Count > 0)
                    {
                        return tree;
                    }
                }
            }
            catch
            {
                // 网络或服务未就绪，安全转入通道 B
            }

            // 通道 B: 自动降级从本地 SQLite 数据库读取离线分类树
            return GetCategoryTreeFromLocalSqlite();
        }

        /// <summary>
        /// 1.1 按需懒加载获取子节点 (支持 el-tree lazy 模式，仅拉取当前层级，节省网络流量)
        /// </summary>
        public static List<AutoPricingCategoryDto> GetCategoryNodes(string? parentId)
        {
            string pId = string.IsNullOrWhiteSpace(parentId) ? "0" : parentId.Trim();
            try
            {
                // 通道 A: 优先尝试从云端 WebAPI 按需拉取
                string url = $"{GetApiBaseUrl()}/api/Scheme/GetCategoryNodes?parentId={pId}";
                var task = _httpClient.GetStringAsync(url);
                if (task.Wait(1500))
                {
                    string json = task.Result;
                    var nodes = JsonSerializer.Deserialize<List<AutoPricingCategoryDto>>(json, JsonOptions);
                    if (nodes != null && nodes.Count > 0)
                    {
                        return nodes;
                    }
                }
            }
            catch
            {
                // 云端超时或未连接，转入本地 SQLite 离线通道
            }

            // 通道 B: 从本地 SQLite 离线查询按需节点
            return GetCategoryNodesFromLocalSqlite(pId);
        }

        /// <summary>
        /// 2. 分页或按条件查询方案列表 (双通道保障)
        /// </summary>
        public static List<AutoPricingSchemeDto> GetSchemes(string? categoryId, string? keyword, string? cabModel, string? situation)
        {
            try
            {
                // 通道 A: 优先请求云端接口
                string qs = $"categoryId={categoryId}&keyword={keyword}&cabModel={cabModel}&situation={situation}&pageSize=100";
                string url = $"{GetApiBaseUrl()}/api/Scheme/GetPagedSchemes?{qs}";
                var task = _httpClient.GetStringAsync(url);
                if (task.Wait(1500))
                {
                    string json = task.Result;
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("items", out var itemsEl))
                    {
                        var list = JsonSerializer.Deserialize<List<AutoPricingSchemeDto>>(itemsEl.GetRawText(), JsonOptions);
                        if (list != null) return list;
                    }
                }
            }
            catch
            {
                // 网络超时转本地
            }

            // 通道 B: 从本地 SQLite 数据库查询方案
            return GetSchemesFromLocalSqlite(categoryId, keyword, cabModel, situation);
        }

        /// <summary>
        /// 3. 获取单个方案详情及完整 BOM 清单 (双通道保障)
        /// </summary>
        public static AutoPricingSchemeDetailDto? GetSchemeDetail(string schemeId)
        {
            try
            {
                // 通道 A: 优先云端查询
                string url = $"{GetApiBaseUrl()}/api/Scheme/GetSchemeDetail?schemeId={schemeId}";
                var task = _httpClient.GetStringAsync(url);
                if (task.Wait(1500))
                {
                    string json = task.Result;
                    var detail = JsonSerializer.Deserialize<AutoPricingSchemeDetailDto>(json, JsonOptions);
                    if (detail != null && detail.BomItems != null && detail.BomItems.Count > 0)
                    {
                        return detail;
                    }
                }
            }
            catch
            {
                // 云端未取到则降级
            }

            // 通道 B: 本地 SQLite 加载
            return GetSchemeDetailFromLocalSqlite(schemeId);
        }

        #region 本地 SQLite 离线加载逻辑

        // 从本地 SQLite 查询组装树形分类
        private static List<AutoPricingCategoryDto> GetCategoryTreeFromLocalSqlite()
        {
            var result = new List<AutoPricingCategoryDto>();
            string dbPath = GetLocalDbPath();
            if (!File.Exists(dbPath)) return result;

            try
            {
                var allCats = new List<AutoPricingCategoryDto>();
                using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();
                    // 统计各分类方案数
                    var countMap = new Dictionary<string, int>();
                    using (var cmdCount = conn.CreateCommand())
                    {
                        cmdCount.CommandText = "SELECT category_id, COUNT(*) FROM schemes GROUP BY category_id;";
                        using var reader = cmdCount.ExecuteReader();
                        while (reader.Read())
                        {
                            string cid = reader.GetString(0);
                            int c = reader.GetInt32(1);
                            countMap[cid] = c;
                        }
                    }

                    // 读取所有分类
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT category_id, parent_id, name, level, full_path, sort_order FROM categories ORDER BY level ASC, sort_order ASC;";
                        using var reader = cmd.ExecuteReader();
                        while (reader.Read())
                        {
                            string cid = reader.GetString(0);
                            string pid = reader.IsDBNull(1) ? "" : reader.GetString(1);
                            string name = reader.GetString(2);
                            int lvl = reader.GetInt32(3);
                            string path = reader.IsDBNull(4) ? "" : reader.GetString(4);
                            int cnt = countMap.TryGetValue(cid, out int sc) ? sc : 0;

                            allCats.Add(new AutoPricingCategoryDto
                            {
                                Id = cid,
                                ParentId = pid,
                                Name = name,
                                Level = lvl,
                                FullPath = path,
                                SchemeCount = cnt,
                                IsCategory = true,
                                Children = new List<AutoPricingCategoryDto>()
                            });
                        }
                    }

                    // 构建父子树
                    var dict = new Dictionary<string, AutoPricingCategoryDto>();
                    // 索引分类 ID 映射
                    foreach (var c in allCats) dict[c.Id] = c;

                    // 组织树级目录结构
                    foreach (var c in allCats)
                    {
                        // 根节点处理
                        if (string.IsNullOrEmpty(c.ParentId) || c.ParentId == "0" || !dict.ContainsKey(c.ParentId))
                        {
                            result.Add(c);
                        }
                        else
                        {
                            // 挂入父级节点子集
                            dict[c.ParentId].Children.Add(c);
                        }
                    }

                    // 递归累计方案总数
                    CalculateTreeCounts(result);

                    // 像利驰一样将具体方案挂载为最末级叶子节点
                    using (var cmdScheme = conn.CreateCommand())
                    {
                        // 查询 schemes 表所有方案概要信息
                        cmdScheme.CommandText = "SELECT scheme_id, category_id, name, cab_model, model, total_price, bom_count FROM schemes ORDER BY name ASC;";
                        // 执行 SQLite 查询
                        using var sReader = cmdScheme.ExecuteReader();
                        // 循环读取方案记录
                        while (sReader.Read())
                        {
                            // 方案主键 ID
                            string sid = sReader.GetString(0);
                            // 所属分类 ID
                            string scid = sReader.GetString(1);
                            // 方案显示名称
                            string sname = sReader.GetString(2);
                            // 开关柜型
                            string cabModel = sReader.IsDBNull(3) ? "" : sReader.GetString(3);
                            // 方案代号
                            string model = sReader.IsDBNull(4) ? "" : sReader.GetString(4);
                            // 参考总价
                            decimal totalPrice = Convert.ToDecimal(sReader.GetDouble(5));
                            // BOM 项数
                            int bCount = sReader.GetInt32(6);

                            // 如果字典中存在该分类节点
                            if (dict.TryGetValue(scid, out var catNode))
                            {
                                // 挂载方案叶子节点
                                catNode.Children.Add(new AutoPricingCategoryDto
                                {
                                    // 节点唯一标识
                                    Id = sid,
                                    // 父级分类标识
                                    ParentId = scid,
                                    // 方案名称
                                    Name = sname,
                                    // 标识该节点为具体方案节点 (非分类目录)
                                    IsCategory = false,
                                    // 方案业务标识
                                    SchemeId = sid,
                                    // 柜型型号
                                    CabModel = cabModel,
                                    // 方案代号
                                    Model = model,
                                    // 参考总价
                                    TotalPrice = totalPrice,
                                    // BOM 元器件项数
                                    BomCount = bCount,
                                    // 叶子节点无子级
                                    Children = new List<AutoPricingCategoryDto>()
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[AutoPricingDataService] 读取本地分类树异常: {ex.Message}");
            }

            return result;
        }

        // 本地 SQLite 离线按需获取分类节点或末级方案 (支持 lazy 懒加载省流)
        private static List<AutoPricingCategoryDto> GetCategoryNodesFromLocalSqlite(string parentId)
        {
            // 初始化返回集合
            var result = new List<AutoPricingCategoryDto>();
            // 获取本地数据库文件路径
            string dbPath = GetLocalDbPath();
            // 文件不存在则安全返回空列表
            if (!File.Exists(dbPath)) return result;

            try
            {
                // 打开 SQLite 离线连接
                using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();

                    // 1. 先检索当前 parentId 下是否含有直接子分类
                    using (var cmd = conn.CreateCommand())
                    {
                        // 若 parentId 为 0 或空，表示获取顶层根分类 --硬编码: 根节点标记--
                        if (string.IsNullOrEmpty(parentId) || parentId == "0")
                        {
                            // 查询无父级或父级为 0 的顶层分类
                            cmd.CommandText = "SELECT category_id, parent_id, name, level, full_path, sort_order FROM categories WHERE parent_id IS NULL OR parent_id = '' OR parent_id = '0' ORDER BY sort_order ASC;";
                        }
                        else
                        {
                            // 查询指定父分类下的直接子分类
                            cmd.CommandText = "SELECT category_id, parent_id, name, level, full_path, sort_order FROM categories WHERE parent_id = @pid ORDER BY sort_order ASC;";
                            // 绑定参数
                            cmd.Parameters.AddWithValue("@pid", parentId);
                        }

                        // 执行查询
                        using var reader = cmd.ExecuteReader();
                        // 循环读取数据
                        while (reader.Read())
                        {
                            // 读取分类 ID
                            string cid = reader.GetString(0);
                            // 读取父分类 ID
                            string pid = reader.IsDBNull(1) ? "" : reader.GetString(1);
                            // 读取名称
                            string name = reader.GetString(2);
                            // 读取层级
                            int lvl = reader.GetInt32(3);
                            // 读取路径
                            string path = reader.IsDBNull(4) ? "" : reader.GetString(4);

                            // 装配分类节点
                            result.Add(new AutoPricingCategoryDto
                            {
                                Id = cid,
                                ParentId = pid,
                                Name = name,
                                Level = lvl,
                                FullPath = path,
                                IsCategory = true,
                                IsLeaf = false, // 分类目录节点允许继续展开
                                Children = new List<AutoPricingCategoryDto>()
                            });
                        }
                    }

                    // 2. 如果存在子分类，统计各子分类方案数量
                    if (result.Count > 0)
                    {
                        // 统计各自分类的方案总数
                        using (var cmdCount = conn.CreateCommand())
                        {
                            // 分组统计命令
                            cmdCount.CommandText = "SELECT category_id, COUNT(*) FROM schemes GROUP BY category_id;";
                            // 执行统计查询
                            using var cReader = cmdCount.ExecuteReader();
                            // 计数映射表
                            var countMap = new Dictionary<string, int>();
                            // 循环填充映射表
                            while (cReader.Read())
                            {
                                countMap[cReader.GetString(0)] = cReader.GetInt32(1);
                            }
                            // 赋值方案统计数字
                            foreach (var item in result)
                            {
                                if (countMap.TryGetValue(item.Id, out int cnt))
                                {
                                    item.SchemeCount = cnt;
                                }
                            }
                        }
                    }

                    // 3. 同时查询直属于当前分类的方案（解决 MNS-GCS-GCK 既有子分类又有直属抽屉方案的展示问题）
                    if (!string.IsNullOrEmpty(parentId) && parentId != "0")
                    {
                        using (var cmdScheme = conn.CreateCommand())
                        {
                            // 查询直属方案记录
                            cmdScheme.CommandText = "SELECT scheme_id, category_id, name, cab_model, model, total_price, bom_count FROM schemes WHERE category_id = @pid ORDER BY name ASC;";
                            // 绑定当前末级分类 ID
                            cmdScheme.Parameters.AddWithValue("@pid", parentId);
                            // 执行查询
                            using var sReader = cmdScheme.ExecuteReader();
                            // 循环构建方案叶子
                            while (sReader.Read())
                            {
                                string sid = sReader.GetString(0);
                                string scid = sReader.GetString(1);
                                string sname = sReader.GetString(2);
                                string cabModel = sReader.IsDBNull(3) ? "" : sReader.GetString(3);
                                string model = sReader.IsDBNull(4) ? "" : sReader.GetString(4);
                                decimal totalPrice = Convert.ToDecimal(sReader.GetDouble(5));
                                int bCount = sReader.GetInt32(6);

                                // 组装方案叶子节点
                                result.Add(new AutoPricingCategoryDto
                                {
                                    Id = sid,
                                    ParentId = scid,
                                    Name = sname,
                                    IsCategory = false,
                                    IsLeaf = true, // 方案节点标记为末级叶子，不可再展开
                                    SchemeId = sid,
                                    CabModel = cabModel,
                                    Model = model,
                                    TotalPrice = totalPrice,
                                    BomCount = bCount,
                                    Children = new List<AutoPricingCategoryDto>()
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 异常日志记录
                LogHelper.WriteLog($"[AutoPricingDataService] 读取本地按需节点异常: {ex.Message}");
            }

            return result;
        }

        // 递归汇总树节点计数
        private static int CalculateTreeCounts(List<AutoPricingCategoryDto> nodes)
        {
            int total = 0;
            foreach (var n in nodes)
            {
                int childSum = CalculateTreeCounts(n.Children);
                n.SchemeCount += childSum;
                total += n.SchemeCount;
            }
            return total;
        }

        // 从本地 SQLite 查询方案列表
        private static List<AutoPricingSchemeDto> GetSchemesFromLocalSqlite(string? categoryId, string? keyword, string? cabModel, string? situation)
        {
            var list = new List<AutoPricingSchemeDto>();
            string dbPath = GetLocalDbPath();
            if (!File.Exists(dbPath)) return list;

            try
            {
                using var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;");
                conn.Open();
                using var cmd = conn.CreateCommand();

                string sql = "SELECT scheme_id, category_id, category_path, name, cab_model, model, dimensions, situation, main_spec, modifier, modify_time, total_price, bom_count FROM schemes WHERE 1=1";

                if (!string.IsNullOrWhiteSpace(categoryId))
                {
                    sql += " AND (category_id = @cid OR category_path LIKE @cpath)";
                    cmd.Parameters.AddWithValue("@cid", categoryId);
                    cmd.Parameters.AddWithValue("@cpath", $"%{categoryId}%");
                }
                if (!string.IsNullOrWhiteSpace(keyword))
                {
                    sql += " AND (name LIKE @kw OR model LIKE @kw OR cab_model LIKE @kw)";
                    cmd.Parameters.AddWithValue("@kw", $"%{keyword.Trim()}%");
                }
                if (!string.IsNullOrWhiteSpace(cabModel))
                {
                    sql += " AND cab_model = @cm";
                    cmd.Parameters.AddWithValue("@cm", cabModel.Trim());
                }
                if (!string.IsNullOrWhiteSpace(situation))
                {
                    sql += " AND situation = @sit";
                    cmd.Parameters.AddWithValue("@sit", situation.Trim());
                }

                sql += " ORDER BY name ASC LIMIT 200;";
                cmd.CommandText = sql;

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    list.Add(new AutoPricingSchemeDto
                    {
                        Id = reader.GetString(0),
                        CategoryId = reader.GetString(1),
                        CategoryPath = reader.IsDBNull(2) ? "" : reader.GetString(2),
                        Name = reader.GetString(3),
                        CabModel = reader.IsDBNull(4) ? "" : reader.GetString(4),
                        Model = reader.IsDBNull(5) ? "" : reader.GetString(5),
                        Dimensions = reader.IsDBNull(6) ? "" : reader.GetString(6),
                        Situation = reader.IsDBNull(7) ? "" : reader.GetString(7),
                        MainSpec = reader.IsDBNull(8) ? "" : reader.GetString(8),
                        Modifier = reader.IsDBNull(9) ? "" : reader.GetString(9),
                        ModifyTime = reader.IsDBNull(10) ? "" : reader.GetString(10),
                        TotalPrice = Convert.ToDecimal(reader.GetDouble(11)),
                        BomCount = reader.GetInt32(12)
                    });
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[AutoPricingDataService] 读取本地方案列表异常: {ex.Message}");
            }
            return list;
        }

        // 从本地 SQLite 查询方案详情及 BOM 明细
        private static AutoPricingSchemeDetailDto? GetSchemeDetailFromLocalSqlite(string schemeId)
        {
            string dbPath = GetLocalDbPath();
            if (!File.Exists(dbPath) || string.IsNullOrWhiteSpace(schemeId)) return null;

            try
            {
                using var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;");
                conn.Open();

                AutoPricingSchemeDetailDto? detail = null;
                // 读取方案主信息
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT scheme_id, category_id, category_path, name, cab_model, model, dimensions,
                               situation, main_spec, modifier, modify_time, total_price, bom_count
                        FROM schemes
                        WHERE scheme_id = @id OR name = @id
                        LIMIT 1;
                    ";
                    cmd.Parameters.AddWithValue("@id", schemeId);
                    using var reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        detail = new AutoPricingSchemeDetailDto
                        {
                            Id = reader.GetString(0),
                            CategoryId = reader.GetString(1),
                            CategoryPath = reader.IsDBNull(2) ? "" : reader.GetString(2),
                            Name = reader.GetString(3),
                            CabModel = reader.IsDBNull(4) ? "" : reader.GetString(4),
                            Model = reader.IsDBNull(5) ? "" : reader.GetString(5),
                            Dimensions = reader.IsDBNull(6) ? "" : reader.GetString(6),
                            Situation = reader.IsDBNull(7) ? "" : reader.GetString(7),
                            MainSpec = reader.IsDBNull(8) ? "" : reader.GetString(8),
                            Modifier = reader.IsDBNull(9) ? "" : reader.GetString(9),
                            ModifyTime = reader.IsDBNull(10) ? "" : reader.GetString(10),
                            TotalPrice = Convert.ToDecimal(reader.GetDouble(11)),
                            BomCount = reader.GetInt32(12),
                            BomItems = new List<CloudSchemeBomItem>()
                        };
                    }
                }

                if (detail == null) return null;

                // 读取 BOM 明细
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT bom_id, seq_no, name, model, brand, unit, quantity,
                               marked_price, discount, quotation_factor, cost_price, subtotal, description
                        FROM bom_items
                        WHERE scheme_id = @id
                        ORDER BY seq_no ASC;
                    ";
                    cmd.Parameters.AddWithValue("@id", detail.Id);
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        double markedPrice = reader.GetDouble(7);
                        double discount = reader.GetDouble(8);
                        double quotation = reader.GetDouble(9);
                        double cost = reader.GetDouble(10);
                        double subtotal = reader.GetDouble(11);

                        detail.BomItems.Add(new CloudSchemeBomItem
                        {
                            Id = reader.GetString(0),
                            SchemeId = detail.Id,
                            SortOrder = reader.GetInt32(1),
                            Name = reader.GetString(2),
                            Model = reader.IsDBNull(3) ? "" : reader.GetString(3),
                            Brand = reader.IsDBNull(4) ? "" : reader.GetString(4),
                            Unit = reader.IsDBNull(5) ? "台" : reader.GetString(5),
                            Quantity = reader.GetDouble(6),
                            CatalogPrice = Convert.ToDecimal(markedPrice),
                            MarkedPrice = Convert.ToDecimal(markedPrice),
                            Discount = Convert.ToDecimal(discount),
                            QuoteDiscount = Convert.ToDecimal(quotation),
                            QuotationFactor = Convert.ToDecimal(quotation),
                            CostPrice = Convert.ToDecimal(cost),
                            QuotePrice = Convert.ToDecimal(Math.Round(markedPrice * discount * quotation, 2)),
                            TotalPrice = Convert.ToDecimal(subtotal),
                            Remark = reader.IsDBNull(12) ? "" : reader.GetString(12),
                            Selected = true
                        });
                    }
                }

                return detail;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[AutoPricingDataService] 读取本地 BOM 详情异常: {ex.Message}");
                return null;
            }
        }

        #endregion

        // 获取云端 API 地址
        private static string GetApiBaseUrl()
        {
            return DefaultApiBaseUrl;
        }

        // 获取本地 SQLite 物理文件路径
        private static string GetLocalDbPath()
        {
            if (File.Exists(LocalSqlitePath)) return LocalSqlitePath;
            string appData = Tool.GetAppDataDirectory();
            string candidate = Path.Combine(appData, "ExWinner_Schemes.db");
            if (File.Exists(candidate)) return candidate;
            return LocalSqlitePath;
        }
    }

    /// <summary>
    /// 自动组价分类树 DTO
    /// </summary>
    public class AutoPricingCategoryDto
    {
        public string Id { get; set; } = string.Empty;
        public string ParentId { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        // 别名：分类名称 (兼容前端 categoryName 取值)
        [System.Text.Json.Serialization.JsonPropertyName("categoryName")]
        public string CategoryName => Name;

        public int Level { get; set; }
        public string FullPath { get; set; } = string.Empty;
        public int SchemeCount { get; set; }

        // 节点类型标识 (true: 分类目录, false: 方案叶子节点)
        public bool IsCategory { get; set; } = true;
        // 节点是否为叶子节点 (用于前端 el-tree lazy 懒加载标识，方案为 true，目录为 false)
        [System.Text.Json.Serialization.JsonPropertyName("isLeaf")]
        public bool IsLeaf { get; set; } = false;
        // 方案 ID
        public string SchemeId { get; set; } = string.Empty;
        // 方案柜型
        public string CabModel { get; set; } = string.Empty;
        // 方案代号
        public string Model { get; set; } = string.Empty;
        // 参考总价
        public decimal TotalPrice { get; set; }
        // BOM 数量
        public int BomCount { get; set; }

        public List<AutoPricingCategoryDto> Children { get; set; } = new List<AutoPricingCategoryDto>();
    }

    /// <summary>
    /// 自动组价方案简要 DTO
    /// </summary>
    public class AutoPricingSchemeDto
    {
        public string Id { get; set; } = string.Empty;
        public string CategoryId { get; set; } = string.Empty;
        public string CategoryPath { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        // 别名：方案名称 (兼容前端 schemeName 取值)
        [System.Text.Json.Serialization.JsonPropertyName("schemeName")]
        public string SchemeName => Name;

        public string CabModel { get; set; } = string.Empty;
        // 别名：开关柜型 (兼容前端 cabinetModel 取值)
        [System.Text.Json.Serialization.JsonPropertyName("cabinetModel")]
        public string CabinetModel => CabModel;

        public string Model { get; set; } = string.Empty;
        // 别名：方案代号/编号 (兼容前端 schemeCode 取值)
        [System.Text.Json.Serialization.JsonPropertyName("schemeCode")]
        public string SchemeCode => Model;

        public string Dimensions { get; set; } = string.Empty;
        public string Situation { get; set; } = string.Empty;
        public string MainSpec { get; set; } = string.Empty;
        // 别名：额定电流规格 (兼容前端 ratedCurrent 取值)
        [System.Text.Json.Serialization.JsonPropertyName("ratedCurrent")]
        public string RatedCurrent => MainSpec;

        public string Modifier { get; set; } = string.Empty;
        public string ModifyTime { get; set; } = string.Empty;
        public decimal TotalPrice { get; set; }
        public int BomCount { get; set; }
    }

    /// <summary>
    /// 自动组价方案详情 DTO
    /// </summary>
    public class AutoPricingSchemeDetailDto : AutoPricingSchemeDto
    {
        public List<CloudSchemeBomItem> BomItems { get; set; } = new List<CloudSchemeBomItem>();
    }
}
