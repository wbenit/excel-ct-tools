using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Text.Json;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// 本地 SQLite 个人数据库服务分部类：二次图控制回路方案与 BOM 数据访问
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，配置与硬编码显式标明
    /// </summary>
    public static partial class PersonalComponentDbService
    {
        // 声明通用的 JSON 序列化选项 (兼容驼峰与忽略大小写，支持中文字符直出)
        private static readonly JsonSerializerOptions SecJsonOptions = new JsonSerializerOptions
        {
            // 属性名忽略大小写匹配
            PropertyNameCaseInsensitive = true,
            // 支持中文等宽字符直出，避免转义为 \uXXXX
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            // 紧凑输出节省存储
            WriteIndented = false
        };

        /// <summary>
        /// 安全序列化 BOM 物料列表为 JSON 字符串
        /// </summary>
        private static string SafeSerializeBomList(List<SecondaryBomItem>? items)
        {
            if (items == null || items.Count == 0) return "[]";
            try
            {
                return JsonSerializer.Serialize(items, SecJsonOptions);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[SecondaryDb] SafeSerializeBomList 异常: {ex.Message}");
                return "[]";
            }
        }

        /// <summary>
        /// 安全反序列化 JSON 字符串为 BOM 物料列表
        /// </summary>
        private static List<SecondaryBomItem> SafeDeserializeBomList(string? json)
        {
            if (string.IsNullOrWhiteSpace(json)) return new List<SecondaryBomItem>();
            try
            {
                return JsonSerializer.Deserialize<List<SecondaryBomItem>>(json, SecJsonOptions) ?? new List<SecondaryBomItem>();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[SecondaryDb] SafeDeserializeBomList 异常: {ex.Message}");
                return new List<SecondaryBomItem>();
            }
        }

        /// <summary>
        /// 查询所有二次方案配置列表 (支持按关键字和二次组模糊检索，并动态补全最新物料单价)
        /// </summary>
        /// <param name="keyword">方案名、适用回路或 CAD 图名检索关键字</param>
        /// <param name="groupName">指定二次组名称 (为空表示全部)</param>
        /// <returns>二次方案配置实体集合</returns>
        public static List<SecondarySchemeEntity> GetAllSecondarySchemes(string? keyword = null, string? groupName = null)
        {
            // 初始化返回列表
            var list = new List<SecondarySchemeEntity>();

            try
            {
                // 获取数据库连接字符串
                string connStr = GetConnectionString();

                // 加锁执行线程安全查询
                lock (_dbLock)
                {
                    // 创建并打开数据库连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 构建参数化查询 SQL 语句
                    var sb = new StringBuilder(@"
                        SELECT id, group_name, scheme_name, applicable_codes, cad_drawing_name, 
                               cross_door_count, hole_spec, labor_cost, bom_json, remark, created_at, updated_at
                        FROM secondary_circuit_schemes
                        WHERE 1=1 
                    ");

                    // 创建命令对象
                    using var cmd = new SQLiteCommand(conn);

                    // 1. 若指定了二次组则按二次组精确筛选
                    if (!string.IsNullOrWhiteSpace(groupName) && !string.Equals(groupName, "全部", StringComparison.OrdinalIgnoreCase))
                    {
                        // 拼接组别过滤条件
                        sb.Append("AND group_name = @group ");
                        // 绑定组别参数
                        cmd.Parameters.AddWithValue("@group", groupName.Trim());
                    }

                    // 2. 若指定了关键字，则在方案名、适用回路、CAD图名、备注中模糊检索
                    if (!string.IsNullOrWhiteSpace(keyword))
                    {
                        // 拼接多字段模糊查询条件
                        sb.Append(@"AND (scheme_name LIKE @kw 
                                      OR applicable_codes LIKE @kw 
                                      OR cad_drawing_name LIKE @kw 
                                      OR remark LIKE @kw) ");
                        // 绑定模糊参数
                        cmd.Parameters.AddWithValue("@kw", $"%{keyword.Trim()}%");
                    }

                    // 按更新时间倒序排序
                    sb.Append("ORDER BY updated_at DESC, id DESC;");
                    // 赋值最终执行 SQL
                    cmd.CommandText = sb.ToString();

                    // 执行读取器查询
                    using var reader = cmd.ExecuteReader();
                    // 循环读取每一行方案数据
                    while (reader.Read())
                    {
                        // 解析基础方案实体
                        var scheme = ReadSchemeFromReader(reader);
                        // 添加到结果集合中
                        list.Add(scheme);
                    }
                }

                // 批量从物料库补齐子 BOM 的最新单价，保障价格绝对一致
                SyncBomPricesFromMaterialDb(list);
            }
            catch (Exception ex)
            {
                // 记录检索异常日志
                LogHelper.WriteLog($"[SecondaryDb] GetAllSecondarySchemes 查询异常: {ex.Message}");
            }

            // 返回方案集合
            return list;
        }

        /// <summary>
        /// 根据方案 ID 获取单个二次回路方案详情
        /// </summary>
        /// <param name="id">方案主键 ID</param>
        /// <returns>方案实体 (不存在返回 null)</returns>
        public static SecondarySchemeEntity? GetSecondarySchemeById(int id)
        {
            // ID 非法直接返回空
            if (id <= 0) return null;

            try
            {
                // 获取连接字符串
                string connStr = GetConnectionString();

                // 互斥安全查询
                lock (_dbLock)
                {
                    // 创建连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 查询单条方案 SQL
                    string sql = @"
                        SELECT id, group_name, scheme_name, applicable_codes, cad_drawing_name, 
                               cross_door_count, hole_spec, labor_cost, bom_json, remark, created_at, updated_at
                        FROM secondary_circuit_schemes
                        WHERE id = @id LIMIT 1;
                    ";

                    // 创建命令
                    using var cmd = new SQLiteCommand(sql, conn);
                    // 绑定 ID 参数
                    cmd.Parameters.AddWithValue("@id", id);

                    // 执行查询
                    using var reader = cmd.ExecuteReader();
                    // 若找到记录
                    if (reader.Read())
                    {
                        // 解析实体
                        var scheme = ReadSchemeFromReader(reader);
                        // 补全最新单价
                        SyncBomPricesFromMaterialDb(new List<SecondarySchemeEntity> { scheme });
                        // 返回方案
                        return scheme;
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[SecondaryDb] GetSecondarySchemeById 异常: {ex.Message}");
            }

            // 未找到返回 null
            return null;
        }

        /// <summary>
        /// 根据回路代号 (如 "双电源1" 或 "CA1B") 智能检索命中的二次回路方案
        /// </summary>
        /// <param name="circuitCode">图纸回路代号或型号</param>
        /// <returns>匹配到的方案实体 (未命中返回 null)</returns>
        public static SecondarySchemeEntity? FindSchemeByCircuitCode(string circuitCode)
        {
            // 参数为空直接返回
            if (string.IsNullOrWhiteSpace(circuitCode)) return null;

            // 清洗输入回路代号
            string cleanCode = circuitCode.Trim();
            // 获取全量方案库进行内存快速多回路匹配
            var allSchemes = GetAllSecondarySchemes();

            // 遍历所有方案
            foreach (var scheme in allSchemes)
            {
                // 检查适用回路代号列表中是否包含该代号 (忽略大小写)
                bool matchCode = scheme.ApplicableCodes.Any(code => string.Equals(code.Trim(), cleanCode, StringComparison.OrdinalIgnoreCase));
                // 若方案主名称相同或适用回路匹配
                if (matchCode || string.Equals(scheme.SchemeName.Trim(), cleanCode, StringComparison.OrdinalIgnoreCase))
                {
                    // 命中则返回该方案
                    return scheme;
                }
            }

            // 未命中返回 null
            return null;
        }

        /// <summary>
        /// 根据 CAD 图名 (如 "双电源二次图.dwg") 智能检索匹配的二次回路方案
        /// </summary>
        /// <param name="cadName">CAD 图名或图块名称</param>
        /// <returns>匹配到的方案实体</returns>
        public static SecondarySchemeEntity? FindSchemeByCadDrawing(string cadName)
        {
            // 参数为空直接返回
            if (string.IsNullOrWhiteSpace(cadName)) return null;

            // 清洗图名
            string cleanCad = cadName.Trim();
            // 读取所有方案
            var allSchemes = GetAllSecondarySchemes();

            // 优先查找 CAD 图名完全匹配或包含关系的方案
            return allSchemes.FirstOrDefault(s =>
                !string.IsNullOrWhiteSpace(s.CadDrawingName) &&
                (string.Equals(s.CadDrawingName.Trim(), cleanCad, StringComparison.OrdinalIgnoreCase) ||
                 cleanCad.IndexOf(s.CadDrawingName.Trim(), StringComparison.OrdinalIgnoreCase) >= 0 ||
                 s.CadDrawingName.Trim().IndexOf(cleanCad, StringComparison.OrdinalIgnoreCase) >= 0));
        }

        /// <summary>
        /// 保存或更新单个二次回路方案
        /// </summary>
        /// <param name="scheme">二次方案实体</param>
        /// <returns>保存成功返回自增 ID，失败返回 -1</returns>
        public static int SaveSecondaryScheme(SecondarySchemeEntity scheme)
        {
            // 实体为空直接返回错误
            if (scheme == null) return -1;

            try
            {
                // 获取连接字符串
                string connStr = GetConnectionString();
                // 序列化子 BOM 为 JSON 文本
                string bomJson = SafeSerializeBomList(scheme.BomItems);
                // 序列化适用回路代号为逗号分隔文本 (如: "双电源1, CA1B")
                string applicableCodesStr = scheme.ApplicableCodes == null ? string.Empty : string.Join(", ", scheme.ApplicableCodes.Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => c.Trim()));
                // 当前时间戳
                string nowTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // 加锁执行数据库写入
                lock (_dbLock)
                {
                    // 创建连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 判断新增还是更新
                    if (scheme.Id <= 0)
                    {
                        // 插入新方案 SQL
                        string insertSql = @"
                            INSERT INTO secondary_circuit_schemes 
                            (group_name, scheme_name, applicable_codes, cad_drawing_name, cross_door_count, hole_spec, labor_cost, bom_json, remark, created_at, updated_at)
                            VALUES 
                            (@group, @name, @codes, @cad, @cross, @hole, @labor, @bom, @remark, @created, @updated);
                            SELECT last_insert_rowid();
                        ";

                        // 创建插入命令
                        using var cmd = new SQLiteCommand(insertSql, conn);
                        // 绑定字段参数
                        cmd.Parameters.AddWithValue("@group", scheme.GroupName?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@name", scheme.SchemeName?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@codes", applicableCodesStr);
                        cmd.Parameters.AddWithValue("@cad", scheme.CadDrawingName?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@cross", scheme.CrossDoorCount);
                        cmd.Parameters.AddWithValue("@hole", scheme.HoleSpec?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@labor", scheme.LaborCost);
                        cmd.Parameters.AddWithValue("@bom", bomJson);
                        cmd.Parameters.AddWithValue("@remark", scheme.Remark?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@created", nowTime);
                        cmd.Parameters.AddWithValue("@updated", nowTime);

                        // 执行并获取新生成自增 ID
                        long newId = Convert.ToInt64(cmd.ExecuteScalar());
                        // 更新实体 ID 并返回
                        scheme.Id = (int)newId;
                        return scheme.Id;
                    }
                    else
                    {
                        // 更新既有方案 SQL
                        string updateSql = @"
                            UPDATE secondary_circuit_schemes 
                            SET group_name = @group,
                                scheme_name = @name,
                                applicable_codes = @codes,
                                cad_drawing_name = @cad,
                                cross_door_count = @cross,
                                hole_spec = @hole,
                                labor_cost = @labor,
                                bom_json = @bom,
                                remark = @remark,
                                updated_at = @updated
                            WHERE id = @id;
                        ";

                        // 创建更新命令
                        using var cmd = new SQLiteCommand(updateSql, conn);
                        // 绑定参数
                        cmd.Parameters.AddWithValue("@id", scheme.Id);
                        cmd.Parameters.AddWithValue("@group", scheme.GroupName?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@name", scheme.SchemeName?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@codes", applicableCodesStr);
                        cmd.Parameters.AddWithValue("@cad", scheme.CadDrawingName?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@cross", scheme.CrossDoorCount);
                        cmd.Parameters.AddWithValue("@hole", scheme.HoleSpec?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@labor", scheme.LaborCost);
                        cmd.Parameters.AddWithValue("@bom", bomJson);
                        cmd.Parameters.AddWithValue("@remark", scheme.Remark?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@updated", nowTime);

                        // 执行更新
                        cmd.ExecuteNonQuery();
                        // 返回当前 ID
                        return scheme.Id;
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录保存异常日志
                LogHelper.WriteLog($"[SecondaryDb] SaveSecondaryScheme 保存异常: {ex.Message}");
                return -1;
            }
        }

        /// <summary>
        /// 删除指定 ID 的二次回路方案
        /// </summary>
        /// <param name="id">方案主键 ID</param>
        /// <returns>是否删除成功</returns>
        public static bool DeleteSecondaryScheme(int id)
        {
            // ID 非法直接返回
            if (id <= 0) return false;

            try
            {
                // 获取连接字符串
                string connStr = GetConnectionString();

                // 互斥操作
                lock (_dbLock)
                {
                    // 创建连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 删除 SQL
                    string sql = "DELETE FROM secondary_circuit_schemes WHERE id = @id;";
                    // 创建命令
                    using var cmd = new SQLiteCommand(sql, conn);
                    // 绑定 ID 参数
                    cmd.Parameters.AddWithValue("@id", id);

                    // 执行删除并判断受影响行数
                    int rows = cmd.ExecuteNonQuery();
                    return rows > 0;
                }
            }
            catch (Exception ex)
            {
                // 记录删除异常
                LogHelper.WriteLog($"[SecondaryDb] DeleteSecondaryScheme 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 批量保存或导入二次回路方案 (带数据库事务保护)
        /// </summary>
        /// <param name="schemes">待导入方案列表</param>
        /// <returns>成功导入的方案总数</returns>
        public static int BatchSaveSecondarySchemes(List<SecondarySchemeEntity> schemes)
        {
            // 列表为空直接返回 0
            if (schemes == null || schemes.Count == 0) return 0;

            // 成功计数器
            int successCount = 0;

            try
            {
                // 获取连接字符串
                string connStr = GetConnectionString();
                // 当前时间
                string nowTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // 加锁执行事务批量写入
                lock (_dbLock)
                {
                    // 创建连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 开启显式 SQLite 事务提升百万级写入性能
                    using var trans = conn.BeginTransaction();

                    // 预编译插入 SQL 语句
                    string sql = @"
                        INSERT INTO secondary_circuit_schemes 
                        (group_name, scheme_name, applicable_codes, cad_drawing_name, cross_door_count, hole_spec, labor_cost, bom_json, remark, created_at, updated_at)
                        VALUES 
                        (@group, @name, @codes, @cad, @cross, @hole, @labor, @bom, @remark, @created, @updated);
                    ";

                    // 循环保存每一个方案
                    foreach (var scheme in schemes)
                    {
                        // 序列化 BOM 清单
                        string bomJson = SafeSerializeBomList(scheme.BomItems);
                        // 序列化适用回路代号
                        string codesStr = scheme.ApplicableCodes == null ? string.Empty : string.Join(", ", scheme.ApplicableCodes.Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => c.Trim()));

                        // 创建命令对象
                        using var cmd = new SQLiteCommand(sql, conn, trans);
                        // 参数绑定
                        cmd.Parameters.AddWithValue("@group", scheme.GroupName?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@name", scheme.SchemeName?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@codes", codesStr);
                        cmd.Parameters.AddWithValue("@cad", scheme.CadDrawingName?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@cross", scheme.CrossDoorCount);
                        cmd.Parameters.AddWithValue("@hole", scheme.HoleSpec?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@labor", scheme.LaborCost);
                        cmd.Parameters.AddWithValue("@bom", bomJson);
                        cmd.Parameters.AddWithValue("@remark", scheme.Remark?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@created", nowTime);
                        cmd.Parameters.AddWithValue("@updated", nowTime);

                        // 执行单条插入
                        cmd.ExecuteNonQuery();
                        // 计数自增
                        successCount++;
                    }

                    // 提交事务
                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                // 记录批量保存异常
                LogHelper.WriteLog($"[SecondaryDb] BatchSaveSecondarySchemes 异常: {ex.Message}");
            }

            // 返回成功数量
            return successCount;
        }

        /// <summary>
        /// 获取当前系统中已存在的所有不重复二次组分类名称列表
        /// </summary>
        /// <returns>二次组名称列表</returns>
        public static List<string> GetSecondaryGroups()
        {
            // 初始化集合
            var groups = new List<string>();

            try
            {
                // 获取连接字符串
                string connStr = GetConnectionString();

                // 加锁查询
                lock (_dbLock)
                {
                    // 创建连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 查询去重组名 SQL
                    string sql = "SELECT DISTINCT group_name FROM secondary_circuit_schemes WHERE group_name IS NOT NULL AND TRIM(group_name) != '' ORDER BY group_name ASC;";
                    // 创建命令
                    using var cmd = new SQLiteCommand(sql, conn);
                    // 读取数据
                    using var reader = cmd.ExecuteReader();
                    // 遍历结果
                    while (reader.Read())
                    {
                        // 提取组名并添加
                        groups.Add(reader.GetString(0));
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[SecondaryDb] GetSecondaryGroups 异常: {ex.Message}");
            }

            // 返回组名列表
            return groups;
        }

        #region 私有辅助方法

        /// <summary>
        /// 从 DataReader 游标中反序列化方案实体基础数据
        /// </summary>
        private static SecondarySchemeEntity ReadSchemeFromReader(SQLiteDataReader reader)
        {
            // 实例化实体对象
            var scheme = new SecondarySchemeEntity
            {
                // 读取主键 ID
                Id = reader.GetInt32(reader.GetOrdinal("id")),
                // 读取二次组名
                GroupName = reader.IsDBNull(reader.GetOrdinal("group_name")) ? string.Empty : reader.GetString(reader.GetOrdinal("group_name")),
                // 读取方案主名
                SchemeName = reader.IsDBNull(reader.GetOrdinal("scheme_name")) ? string.Empty : reader.GetString(reader.GetOrdinal("scheme_name")),
                // 读取 CAD 图名
                CadDrawingName = reader.IsDBNull(reader.GetOrdinal("cad_drawing_name")) ? string.Empty : reader.GetString(reader.GetOrdinal("cad_drawing_name")),
                // 读取二次跨门线
                CrossDoorCount = reader.IsDBNull(reader.GetOrdinal("cross_door_count")) ? 0.0 : reader.GetDouble(reader.GetOrdinal("cross_door_count")),
                // 读取开孔规范
                HoleSpec = reader.IsDBNull(reader.GetOrdinal("hole_spec")) ? string.Empty : reader.GetString(reader.GetOrdinal("hole_spec")),
                // 读取人工工费
                LaborCost = reader.IsDBNull(reader.GetOrdinal("labor_cost")) ? 0.0 : reader.GetDouble(reader.GetOrdinal("labor_cost")),
                // 读取备注
                Remark = reader.IsDBNull(reader.GetOrdinal("remark")) ? string.Empty : reader.GetString(reader.GetOrdinal("remark")),
                // 读取创建时间
                CreatedAt = reader.IsDBNull(reader.GetOrdinal("created_at")) ? string.Empty : reader.GetString(reader.GetOrdinal("created_at")),
                // 读取更新时间
                UpdatedAt = reader.IsDBNull(reader.GetOrdinal("updated_at")) ? string.Empty : reader.GetString(reader.GetOrdinal("updated_at"))
            };

            // 解析适用回路代号 (逗号分隔文本解析为 List<string>)
            string codesRaw = reader.IsDBNull(reader.GetOrdinal("applicable_codes")) ? string.Empty : reader.GetString(reader.GetOrdinal("applicable_codes"));
            if (!string.IsNullOrWhiteSpace(codesRaw))
            {
                // 按逗号拆分并清洗空格
                scheme.ApplicableCodes = codesRaw.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                                                 .Select(c => c.Trim())
                                                 .Where(c => !string.IsNullOrEmpty(c))
                                                 .ToList();
            }

            // 解析子 BOM 明细 JSON 文本
            string bomRaw = reader.IsDBNull(reader.GetOrdinal("bom_json")) ? string.Empty : reader.GetString(reader.GetOrdinal("bom_json"));
            scheme.BomItems = SafeDeserializeBomList(bomRaw);

            // 返回构造完成的方案实体
            return scheme;
        }

        /// <summary>
        /// 从本地物料库 components 表中批量同步匹配子 BOM 物料的最新单价，保持全局一致性
        /// </summary>
        private static void SyncBomPricesFromMaterialDb(List<SecondarySchemeEntity> schemes)
        {
            // 列表为空直接返回
            if (schemes == null || schemes.Count == 0) return;

            // 收集所有子 BOM 中存在的物料 ID 集合
            var componentIds = new HashSet<int>();
            // 收集未关联 ID 但有型号名称的待匹配物料
            var pendingMatches = new List<SecondaryBomItem>();

            // 遍历所有方案
            foreach (var scheme in schemes)
            {
                // 遍历子项
                foreach (var item in scheme.BomItems)
                {
                    // 若有明确 ID
                    if (item.ComponentId > 0)
                    {
                        // 收集 ID
                        componentIds.Add(item.ComponentId);
                    }
                    else if (!string.IsNullOrWhiteSpace(item.Model))
                    {
                        // 收集待匹配项
                        pendingMatches.Add(item);
                    }
                }
            }

            // 字典缓存物料库 ID 到最新单价的映射
            var priceMap = new Dictionary<int, double>();
            // 字典缓存型号到物料实体信息的映射
            var modelMap = new Dictionary<string, (int id, double price, string name, string brand)>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // 获取连接字符串
                string connStr = GetConnectionString();

                // 互斥安全只读
                lock (_dbLock)
                {
                    // 创建连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 1. 若存在特定 ID 则批量查询 ID 对应的最新价格
                    if (componentIds.Count > 0)
                    {
                        // 构造 IN 查询语句
                        string inClause = string.Join(",", componentIds);
                        // 查询 SQL
                        string idSql = $"SELECT id, price FROM components WHERE id IN ({inClause});";
                        // 创建命令
                        using var idCmd = new SQLiteCommand(idSql, conn);
                        // 执行读取
                        using var idReader = idCmd.ExecuteReader();
                        // 填充映射表
                        while (idReader.Read())
                        {
                            int cid = idReader.GetInt32(0);
                            double cPrice = idReader.IsDBNull(1) ? 0.0 : idReader.GetDouble(1);
                            priceMap[cid] = cPrice;
                        }
                    }

                    // 2. 对未绑定 ID 的项，按型号从物料库模糊匹配
                    if (pendingMatches.Count > 0)
                    {
                        // 提取不重复型号集合
                        var models = pendingMatches.Select(m => m.Model.Trim()).Distinct().ToList();
                        // 遍历型号查询物料库
                        foreach (var m in models)
                        {
                            // 查询单条最新物料
                            string mSql = "SELECT id, price, name, brand FROM components WHERE model = @model LIMIT 1;";
                            using var mCmd = new SQLiteCommand(mSql, conn);
                            mCmd.Parameters.AddWithValue("@model", m);
                            using var mReader = mCmd.ExecuteReader();
                            if (mReader.Read())
                            {
                                int mid = mReader.GetInt32(0);
                                double mPrice = mReader.IsDBNull(1) ? 0.0 : mReader.GetDouble(1);
                                string mName = mReader.IsDBNull(2) ? string.Empty : mReader.GetString(2);
                                string mBrand = mReader.IsDBNull(3) ? string.Empty : mReader.GetString(3);
                                modelMap[m] = (mid, mPrice, mName, mBrand);
                            }
                        }
                    }
                }

                // 回填最新单价到方案实体中
                foreach (var scheme in schemes)
                {
                    foreach (var item in scheme.BomItems)
                    {
                        // 若 ID 在价格映射中存在
                        if (item.ComponentId > 0 && priceMap.TryGetValue(item.ComponentId, out double latestPrice))
                        {
                            // 刷新为物料库最新价格
                            item.UnitPrice = latestPrice;
                        }
                        // 若通过型号命中了物料库
                        else if (!string.IsNullOrWhiteSpace(item.Model) && modelMap.TryGetValue(item.Model.Trim(), out var hit))
                        {
                            // 自动补齐物料库 ID 与最新价格
                            item.ComponentId = hit.id;
                            item.UnitPrice = hit.price;
                            if (string.IsNullOrWhiteSpace(item.Name)) item.Name = hit.name;
                            if (string.IsNullOrWhiteSpace(item.Brand)) item.Brand = hit.brand;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录价格补全异常
                LogHelper.WriteLog($"[SecondaryDb] SyncBomPricesFromMaterialDb 异常: {ex.Message}");
            }
        }

        #endregion

        #region 回路代号全局唯一性防重检测

        /// <summary>
        /// 回路代号全局唯一性冲突检测结果结构
        /// </summary>
        public class CodeConflictResult
        {
            // 发生冲突的代号文本
            public string Code { get; set; } = string.Empty;
            // 占用该代号的已有方案主名称
            public string SchemeName { get; set; } = string.Empty;
            // 占用该代号的已有方案主键 ID
            public int SchemeId { get; set; }
        }

        /// <summary>
        /// 检查一组回路代号是否已被其他方案绑定占用 (全局防重门禁校验)
        /// </summary>
        /// <param name="codes">待检查的回路代号列表</param>
        /// <param name="excludeSchemeId">排除当前正在编辑的方案 ID (新建为 0)</param>
        /// <returns>若存在冲突则返回首个冲突详情，完全无冲突则返回 null</returns>
        public static CodeConflictResult? CheckApplicableCodeConflict(IEnumerable<string>? codes, int excludeSchemeId = 0)
        {
            // 若列表为空则直接放行
            if (codes == null) return null;
            // 清洗并提取有效非空代号列表 (去重且去除首尾空白字符)
            var codeList = codes.Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => c.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            // 若有效代号数量为 0 则无需检查
            if (codeList.Count == 0) return null;

            try
            {
                // 获取当前本地 SQLite 方案库中的全量方案数据
                var allSchemes = GetAllSecondarySchemes();
                // 排除当前正在编辑的方案自身
                var otherSchemes = allSchemes.Where(s => s.Id != excludeSchemeId).ToList();

                // 逐个遍历待检查的代号
                foreach (var code in codeList)
                {
                    // 遍历其他既有方案
                    foreach (var other in otherSchemes)
                    {
                        // 检查其他方案的适用回路代号列表中是否包含该代号 (大小写不敏感匹配)
                        if (other.ApplicableCodes != null && other.ApplicableCodes.Any(c => string.Equals(c?.Trim(), code, StringComparison.OrdinalIgnoreCase)))
                        {
                            // 发现全局同名冲突，封装冲突详情返回
                            return new CodeConflictResult
                            {
                                Code = code,
                                SchemeName = other.SchemeName,
                                SchemeId = other.Id
                            };
                        }
                    }
                }

                // 所有代号均具备全局唯一性，校验通过
                return null;
            }
            catch (Exception ex)
            {
                // 记录防重检测异常日志
                LogHelper.WriteLog($"[SecondaryDb] CheckApplicableCodeConflict 异常: {ex.Message}");
                // 异常情况下安全放行
                return null;
            }
        }

        #endregion
    }
}
