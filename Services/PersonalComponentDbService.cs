using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// 本地个人物料库 SQLite 数据库访问服务 (结构与云端 components 表 1:1 镜像)
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，配置与硬编码显式标明
    /// </summary>
    public static class PersonalComponentDbService
    {
        // 线程安全互斥锁，保障多线程与并发访问 SQLite 安全
        private static readonly object _dbLock = new object();

        // 默认数据库文件名称
        private const string DefaultDbFileName = "personal_components.db"; // --硬编码-- 默认个人物料库文件名

        // 缓存当前生效的 SQLite 连接字符串
        private static string? _connectionString;

        /// <summary>
        /// 获取或初始化本地个人物料库 SQLite 数据库文件的绝对物理路径
        /// </summary>
        /// <returns>数据库文件全路径</returns>
        public static string GetDatabaseFilePath()
        {
            try
            {
                // 获取 Windows LocalAppData 标准目录
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                // 拼接 ExcelCTTools 专属数据目录
                string dataFolder = Path.Combine(localAppData, "ExcelCTTools", "data"); // --硬编码-- 默认应用数据子目录

                // 校验目录是否存在，不存在则自动递归创建
                if (!Directory.Exists(dataFolder))
                {
                    // 递归创建 data 数据保存目录
                    Directory.CreateDirectory(dataFolder);
                }

                // 组合生成个人物料库数据库文件全路径
                return Path.Combine(dataFolder, DefaultDbFileName);
            }
            catch (Exception ex)
            {
                // 记录日志并在异常时使用插件目录降级兜底
                LogHelper.WriteLog($"[PersonalDb] 获取数据库路径异常: {ex.Message}，降级至应用目录");
                // 获取插件 data 目录
                string fallbackDir = Tool.GetAppDataDirectory();
                // 拼接返回降级路径
                return Path.Combine(fallbackDir, DefaultDbFileName);
            }
        }

        /// <summary>
        /// 获取 SQLite 标准连接字符串并执行数据库自愈建表
        /// </summary>
        /// <returns>可用的 SQLite 连接字符串</returns>
        public static string GetConnectionString()
        {
            // 如果连接字符串已初始化则快速返回
            if (!string.IsNullOrEmpty(_connectionString))
            {
                return _connectionString!;
            }

            // 加锁保障多线程安全初始化
            lock (_dbLock)
            {
                // 双重校验锁
                if (!string.IsNullOrEmpty(_connectionString))
                {
                    return _connectionString!;
                }

                // 获取数据库文件的物理全路径
                string dbPath = GetDatabaseFilePath();
                // 构建标准 SQLite 连接字符串 (启用 UTF-8 与连接池)
                _connectionString = $"Data Source={dbPath};Version=3;Pooling=True;Max Pool Size=20;"; // --硬编码-- 连接字符串配置

                // 确保数据库文件及数据表、索引已初始化
                EnsureDatabaseCreated(dbPath);

                // 返回构建好的连接字符串
                return _connectionString;
            }
        }

        /// <summary>
        /// 确保数据库文件和 1:1 镜像的 components 表结构与索引自愈创建
        /// </summary>
        /// <param name="dbPath">数据库物理路径</param>
        public static void EnsureDatabaseCreated(string? dbPath = null)
        {
            // 路径为空时自动获取默认路径
            string targetPath = dbPath ?? GetDatabaseFilePath();

            // 进入互斥锁，避免多并发同时执行 DDL 建表
            lock (_dbLock)
            {
                try
                {
                    // 若数据库文件不存在，创建新 SQLite 数据库
                    if (!File.Exists(targetPath))
                    {
                        // 创建空白数据库文件
                        SQLiteConnection.CreateFile(targetPath);
                    }

                    // 创建并打开数据库连接
                    using var conn = new SQLiteConnection($"Data Source={targetPath};Version=3;");
                    // 打开底层物理连接
                    conn.Open();

                    // 构建建表 DDL 语句 (与服务器 DrawMall.Domain.Models.Component 字段 1:1 对齐)
                    string ddl = @"
                        CREATE TABLE IF NOT EXISTS components (
                            id          INTEGER PRIMARY KEY AUTOINCREMENT,
                            brand       TEXT NOT NULL,
                            original_id INTEGER,
                            name        TEXT,
                            model       TEXT NOT NULL,
                            price       REAL NOT NULL DEFAULT 0.0,
                            remark      TEXT,
                            param1      TEXT,
                            param2      TEXT,
                            current     INTEGER,
                            poles       TEXT,
                            tripping    TEXT,
                            created_at  TEXT,
                            updated_at  TEXT
                        );
                        CREATE INDEX IF NOT EXISTS idx_comp_brand_name ON components(brand, name);
                        CREATE INDEX IF NOT EXISTS idx_comp_model ON components(model);
                        CREATE INDEX IF NOT EXISTS idx_comp_brand ON components(brand);
                    "; // --硬编码-- SQLite 建表 DDL 脚本

                    // 执行 DDL 创建数据表和索引
                    using var cmd = new SQLiteCommand(ddl, conn);
                    // 执行非查询 SQL
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    // 记录建表异常日志
                    LogHelper.WriteLog($"[PersonalDb] 初始化数据库表异常: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 从本地个人物料库中统计所有品牌及其物料数量 (与云端 GetBrandStats 结构对齐)
        /// </summary>
        /// <returns>品牌及其数量统计列表</returns>
        public static List<BrandStatItemDto> GetBrandStats()
        {
            // 初始化返回集合
            var result = new List<BrandStatItemDto>();

            try
            {
                // 获取有效连接字符串
                string connStr = GetConnectionString();

                // 加锁执行只读查询
                lock (_dbLock)
                {
                    // 创建数据库连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 构建品牌聚合统计 SQL 语句 (按数量降序、品牌名升序排列)
                    string sql = @"
                        SELECT brand, COUNT(*) AS cnt 
                        FROM components 
                        WHERE brand IS NOT NULL AND TRIM(brand) != '' 
                        GROUP BY brand 
                        ORDER BY cnt DESC, brand ASC;
                    "; // --硬编码-- 品牌统计 SQL

                    // 创建执行命令
                    using var cmd = new SQLiteCommand(sql, conn);
                    // 执行数据读取
                    using var reader = cmd.ExecuteReader();

                    // 遍历数据行
                    while (reader.Read())
                    {
                        // 提取品牌名称
                        string brandName = reader.IsDBNull(0) ? string.Empty : reader.GetString(0);
                        // 提取元器件数量
                        int count = reader.IsDBNull(1) ? 0 : Convert.ToInt32(reader.GetInt64(1));

                        // 添加到返回结果集
                        result.Add(new BrandStatItemDto
                        {
                            Brand = brandName,
                            Count = count
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录统计查询异常
                LogHelper.WriteLog($"[PersonalDb] GetBrandStats 异常: {ex.Message}");
            }

            // 返回统计结果集合
            return result;
        }

        /// <summary>
        /// 根据品牌获取该品牌在个人库中的所有去重元器件名称列表
        /// </summary>
        /// <param name="brand">指定品牌 (为空表示不限品牌)</param>
        /// <returns>元器件名称列表</returns>
        public static List<string> GetNamesByBrand(string? brand)
        {
            // 初始化名称集合
            var list = new List<string>();

            try
            {
                // 获取连接字符串
                string connStr = GetConnectionString();

                // 加锁执行查询
                lock (_dbLock)
                {
                    // 创建连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 动态构建 SQL 语句
                    var sb = new StringBuilder("SELECT DISTINCT name FROM components WHERE name IS NOT NULL AND TRIM(name) != '' ");
                    // 若指定了品牌则添加品牌过滤条件
                    if (!string.IsNullOrWhiteSpace(brand))
                    {
                        // 拼接品牌相等条件
                        sb.Append("AND brand = @brand ");
                    }
                    // 按名称正序排序
                    sb.Append("ORDER BY name ASC;");

                    // 创建命令
                    using var cmd = new SQLiteCommand(sb.ToString(), conn);
                    // 若指定了品牌则绑定参数
                    if (!string.IsNullOrWhiteSpace(brand))
                    {
                        // 添加品牌参数
                        cmd.Parameters.AddWithValue("@brand", brand!.Trim());
                    }

                    // 执行读取
                    using var reader = cmd.ExecuteReader();
                    // 循环读取数据
                    while (reader.Read())
                    {
                        // 读取名称文本
                        if (!reader.IsDBNull(0))
                        {
                            // 加入列表
                            list.Add(reader.GetString(0));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录日志
                LogHelper.WriteLog($"[PersonalDb] GetNamesByBrand 异常: {ex.Message}");
            }

            // 返回名称列表
            return list;
        }

        /// <summary>
        /// 供元器件数据管理窗口全量/条件拉取物料明细数据
        /// </summary>
        /// <param name="brand">品牌过滤条件</param>
        /// <param name="nameKeyword">名称过滤条件</param>
        /// <returns>符合条件的个人库物料实体列表</returns>
        public static List<ComponentApiDto> QueryManageComponents(string? brand, string? nameKeyword)
        {
            // 初始化返回列表
            var result = new List<ComponentApiDto>();

            try
            {
                // 获取连接字符串
                string connStr = GetConnectionString();

                // 加锁执行
                lock (_dbLock)
                {
                    // 创建连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 构建 SQL 基础查询
                    var sb = new StringBuilder(@"
                        SELECT id, brand, name, model, price, remark, param1, param2, current, poles, tripping 
                        FROM components 
                        WHERE 1=1 
                    "); // --硬编码-- 查询字段列表

                    // 品牌筛选
                    if (!string.IsNullOrWhiteSpace(brand))
                    {
                        sb.Append("AND brand = @brand ");
                    }
                    // 名称筛选
                    if (!string.IsNullOrWhiteSpace(nameKeyword))
                    {
                        sb.Append("AND name LIKE @name ");
                    }
                    // 按主键降序排列
                    sb.Append("ORDER BY id DESC;");

                    // 创建 SQL 命令
                    using var cmd = new SQLiteCommand(sb.ToString(), conn);
                    // 绑定品牌参数
                    if (!string.IsNullOrWhiteSpace(brand))
                    {
                        cmd.Parameters.AddWithValue("@brand", brand!.Trim());
                    }
                    // 绑定名称模糊参数
                    if (!string.IsNullOrWhiteSpace(nameKeyword))
                    {
                        cmd.Parameters.AddWithValue("@name", $"%{nameKeyword!.Trim()}%");
                    }

                    // 执行查询
                    using var reader = cmd.ExecuteReader();
                    // 遍历提取记录
                    while (reader.Read())
                    {
                        // 映射 DTO 实体
                        result.Add(MapReaderToDto(reader));
                    }
                }
            }
            catch (Exception ex)
            {
                // 异常日志
                LogHelper.WriteLog($"[PersonalDb] QueryManageComponents 异常: {ex.Message}");
            }

            // 返回结果集合
            return result;
        }

        /// <summary>
        /// 综合检索与智能模糊查询核心方法 (支持本地 SQLite 个人物料库多维检索与智能降级)
        /// </summary>
        public static List<ComponentApiDto> SearchComponents(
            string? searchKeyword,
            string? name,
            string? current,
            string? pole,
            string? tripMode,
            string? brand,
            List<MustContainRule>? mustContainRules,
            int maxResults = 100)
        {
            // 初始化返回值列表
            var result = new List<ComponentApiDto>();

            try
            {
                // 获取连接字符串
                string connStr = GetConnectionString();

                // 加锁线程安全读取
                lock (_dbLock)
                {
                    // 创建数据库连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 构建查询
                    var sb = new StringBuilder(@"
                        SELECT id, brand, name, model, price, remark, param1, param2, current, poles, tripping 
                        FROM components 
                        WHERE 1=1 
                    ");

                    // 创建命令
                    using var cmd = new SQLiteCommand(conn);

                    // 1. 品牌精确过滤 (排除“全部”、“全部品牌”等非真实品牌)
                    string cleanBrand = brand?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(cleanBrand) &&
                        !string.Equals(cleanBrand, "全部", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(cleanBrand, "全部品牌", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(cleanBrand, "All", StringComparison.OrdinalIgnoreCase))
                    {
                        sb.Append("AND brand = @brand ");
                        cmd.Parameters.AddWithValue("@brand", cleanBrand);
                    }

                    // 2. 名称模糊过滤
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        sb.Append("AND name LIKE @name ");
                        cmd.Parameters.AddWithValue("@name", $"%{name!.Trim()}%");
                    }

                    // 3. 额定电流过滤 (兼容字段存储及型号中内嵌如 16A 的情况)
                    int? parsedCurrent = ExtractIntegerCurrent(current);
                    if (parsedCurrent.HasValue)
                    {
                        int cVal = parsedCurrent.Value;
                        sb.Append("AND (current = @current OR (current IS NULL AND (model LIKE @curLike1 OR model LIKE @curLike2 OR model LIKE @curLike3))) ");
                        cmd.Parameters.AddWithValue("@current", cVal);
                        cmd.Parameters.AddWithValue("@curLike1", $"%{cVal}A%");
                        cmd.Parameters.AddWithValue("@curLike2", $"%/{cVal}%");
                        cmd.Parameters.AddWithValue("@curLike3", $"%{cVal}%");
                    }

                    // 4. 极数过滤 (兼容 3 与 3P 以及型号中内嵌)
                    string cleanPoles = NormalizePolesParam(pole);
                    if (!string.IsNullOrEmpty(cleanPoles))
                    {
                        sb.Append("AND (poles = @pole OR poles = @poleP OR (poles IS NULL AND (model LIKE @poleLike1 OR model LIKE @poleLike2))) ");
                        cmd.Parameters.AddWithValue("@pole", cleanPoles);
                        cmd.Parameters.AddWithValue("@poleP", $"{cleanPoles}P");
                        cmd.Parameters.AddWithValue("@poleLike1", $"%{cleanPoles}P%");
                        cmd.Parameters.AddWithValue("@poleLike2", $"%/{cleanPoles}%");
                    }

                    // 5. 脱扣方式过滤
                    if (!string.IsNullOrWhiteSpace(tripMode))
                    {
                        sb.Append("AND tripping LIKE @trip ");
                        cmd.Parameters.AddWithValue("@trip", $"%{tripMode!.Trim()}%");
                    }

                    // 6. 用户关键字模糊匹配 (同时命中 model 或 name 或 param1 或 remark)
                    if (!string.IsNullOrWhiteSpace(searchKeyword))
                    {
                        string kw = searchKeyword!.Trim();
                        sb.Append("AND (model LIKE @kw OR name LIKE @kw OR param1 LIKE @kw OR remark LIKE @kw) ");
                        cmd.Parameters.AddWithValue("@kw", $"%{kw}%");
                    }

                    // 7. 必含约束规则 (多条规则间为 AND 关系，要求 model 必须包含关键字)
                    if (mustContainRules != null && mustContainRules.Count > 0)
                    {
                        int ruleIndex = 0;
                        foreach (var rule in mustContainRules)
                        {
                            if (rule.Enabled && !string.IsNullOrWhiteSpace(rule.Keyword))
                            {
                                string paramName = $"@rule_{ruleIndex}";
                                sb.Append($"AND model LIKE {paramName} ");
                                cmd.Parameters.AddWithValue(paramName, $"%{rule.Keyword.Trim()}%");
                                ruleIndex++;
                            }
                        }
                    }

                    // 限制最大返回行数
                    sb.Append($"ORDER BY id DESC LIMIT {maxResults};");
                    cmd.CommandText = sb.ToString();

                    // 执行主要条件查询
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(MapReaderToDto(reader));
                        }
                    }

                    // 8. 智能降级机制：当用户主动在搜索框输入了关键字，但由于单元格带入的名称等其他约束导致 0 命中时，自动按用户输入的关键字在当前品牌下检索
                    if (result.Count == 0 && !string.IsNullOrWhiteSpace(searchKeyword))
                    {
                        var fallbackSb = new StringBuilder(@"
                            SELECT id, brand, name, model, price, remark, param1, param2, current, poles, tripping 
                            FROM components 
                            WHERE 1=1 
                        ");
                        using var fallbackCmd = new SQLiteCommand(conn);
                        if (!string.IsNullOrEmpty(cleanBrand) &&
                            !string.Equals(cleanBrand, "全部", StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(cleanBrand, "全部品牌", StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(cleanBrand, "All", StringComparison.OrdinalIgnoreCase))
                        {
                            fallbackSb.Append("AND brand = @fbBrand ");
                            fallbackCmd.Parameters.AddWithValue("@fbBrand", cleanBrand);
                        }
                        string kw = searchKeyword!.Trim();
                        fallbackSb.Append("AND (model LIKE @fbKw OR name LIKE @fbKw OR param1 LIKE @fbKw OR remark LIKE @fbKw) ");
                        fallbackCmd.Parameters.AddWithValue("@fbKw", $"%{kw}%");

                        // 必含规则约束
                        if (mustContainRules != null && mustContainRules.Count > 0)
                        {
                            int ruleIdx = 0;
                            foreach (var rule in mustContainRules)
                            {
                                if (rule.Enabled && !string.IsNullOrWhiteSpace(rule.Keyword))
                                {
                                    string paramName = $"@fbRule_{ruleIdx}";
                                    fallbackSb.Append($"AND model LIKE {paramName} ");
                                    fallbackCmd.Parameters.AddWithValue(paramName, $"%{rule.Keyword.Trim()}%");
                                    ruleIdx++;
                                }
                            }
                        }

                        fallbackSb.Append($"ORDER BY id DESC LIMIT {maxResults};");
                        fallbackCmd.CommandText = fallbackSb.ToString();

                        using var fbReader = fallbackCmd.ExecuteReader();
                        while (fbReader.Read())
                        {
                            result.Add(MapReaderToDto(fbReader));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[PersonalDb] SearchComponents 异常: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 批量新增元器件至本地个人物料库 (采用事务高速插入)
        /// </summary>
        /// <param name="requests">新增请求列表</param>
        /// <returns>新增成功并填充了自增 ID 的实体列表</returns>
        public static List<ComponentApiDto> CreateComponents(List<CreateComponentApiRequest> requests)
        {
            var createdList = new List<ComponentApiDto>();
            if (requests == null || requests.Count == 0) return createdList;

            try
            {
                string connStr = GetConnectionString();
                lock (_dbLock)
                {
                    using var conn = new SQLiteConnection(connStr);
                    conn.Open();
                    // 开启事务极速批处理
                    using var trans = conn.BeginTransaction();

                    string insertSql = @"
                        INSERT INTO components (brand, name, model, price, remark, param1, param2, current, poles, tripping, created_at, updated_at) 
                        VALUES (@brand, @name, @model, @price, @remark, @param1, @param2, @current, @poles, @tripping, @created_at, @updated_at);
                        SELECT last_insert_rowid();
                    ";

                    string nowIso = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    foreach (var req in requests)
                    {
                        using var cmd = new SQLiteCommand(insertSql, conn, trans);
                        cmd.Parameters.AddWithValue("@brand", req.Brand?.Trim() ?? "通用"); // --硬编码-- 默认品牌
                        cmd.Parameters.AddWithValue("@name", req.Name?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@model", req.Model?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@price", Convert.ToDouble(req.Price));
                        cmd.Parameters.AddWithValue("@remark", req.Remark ?? string.Empty);
                        cmd.Parameters.AddWithValue("@param1", req.Param1 ?? string.Empty);
                        cmd.Parameters.AddWithValue("@param2", req.Param2 ?? string.Empty);
                        cmd.Parameters.AddWithValue("@current", (object?)req.Current ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@poles", req.Poles ?? string.Empty);
                        cmd.Parameters.AddWithValue("@tripping", req.Tripping ?? string.Empty);
                        cmd.Parameters.AddWithValue("@created_at", nowIso);
                        cmd.Parameters.AddWithValue("@updated_at", nowIso);

                        // 获取刚插入的主键自增 ID
                        long newId = (long)cmd.ExecuteScalar();

                        createdList.Add(new ComponentApiDto
                        {
                            Id = (int)newId,
                            Brand = req.Brand,
                            Name = req.Name,
                            Model = req.Model,
                            Price = req.Price,
                            Remark = req.Remark,
                            Param1 = req.Param1,
                            Param2 = req.Param2,
                            Current = req.Current,
                            Poles = req.Poles,
                            Tripping = req.Tripping
                        });
                    }

                    // 提交事务
                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[PersonalDb] CreateComponents 异常: {ex.Message}");
            }

            return createdList;
        }

        /// <summary>
        /// 批量更新本地个人物料库元器件 (采用事务批量执行)
        /// </summary>
        /// <param name="requests">待更新列表</param>
        /// <returns>成功更新的行数</returns>
        public static int UpdateComponents(List<UpdateComponentApiRequest> requests)
        {
            if (requests == null || requests.Count == 0) return 0;
            int updatedCount = 0;

            try
            {
                string connStr = GetConnectionString();
                lock (_dbLock)
                {
                    using var conn = new SQLiteConnection(connStr);
                    conn.Open();
                    using var trans = conn.BeginTransaction();

                    string updateSql = @"
                        UPDATE components 
                        SET brand = @brand,
                            name = @name,
                            model = @model,
                            price = @price,
                            remark = @remark,
                            param1 = @param1,
                            param2 = @param2,
                            current = @current,
                            poles = @poles,
                            tripping = @tripping,
                            updated_at = @updated_at 
                        WHERE id = @id;
                    ";

                    string nowIso = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    foreach (var req in requests)
                    {
                        using var cmd = new SQLiteCommand(updateSql, conn, trans);
                        cmd.Parameters.AddWithValue("@id", req.Id);
                        cmd.Parameters.AddWithValue("@brand", req.Brand?.Trim() ?? "通用");
                        cmd.Parameters.AddWithValue("@name", req.Name?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@model", req.Model?.Trim() ?? string.Empty);
                        cmd.Parameters.AddWithValue("@price", Convert.ToDouble(req.Price ?? 0m));
                        cmd.Parameters.AddWithValue("@remark", req.Remark ?? string.Empty);
                        cmd.Parameters.AddWithValue("@param1", req.Param1 ?? string.Empty);
                        cmd.Parameters.AddWithValue("@param2", req.Param2 ?? string.Empty);
                        cmd.Parameters.AddWithValue("@current", (object?)req.Current ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@poles", req.Poles ?? string.Empty);
                        cmd.Parameters.AddWithValue("@tripping", req.Tripping ?? string.Empty);
                        cmd.Parameters.AddWithValue("@updated_at", nowIso);

                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0) updatedCount++;
                    }

                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[PersonalDb] UpdateComponents 异常: {ex.Message}");
            }

            return updatedCount;
        }

        /// <summary>
        /// 根据主键 ID 列表从本地个人物料库批量物理删除
        /// </summary>
        /// <param name="ids">待删除 ID 列表</param>
        /// <returns>成功删除的行数</returns>
        public static int DeleteComponents(List<int> ids)
        {
            if (ids == null || ids.Count == 0) return 0;
            int deletedCount = 0;

            try
            {
                string connStr = GetConnectionString();
                lock (_dbLock)
                {
                    using var conn = new SQLiteConnection(connStr);
                    conn.Open();
                    using var trans = conn.BeginTransaction();

                    // 分块执行删除 (防范 IN 条件参数过长)
                    const int batchSize = 100;
                    for (int i = 0; i < ids.Count; i += batchSize)
                    {
                        var chunk = ids.Skip(i).Take(batchSize).ToList();
                        string inClause = string.Join(",", chunk);
                        string sql = $"DELETE FROM components WHERE id IN ({inClause});";

                        using var cmd = new SQLiteCommand(sql, conn, trans);
                        deletedCount += cmd.ExecuteNonQuery();
                    }

                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[PersonalDb] DeleteComponents 异常: {ex.Message}");
            }

            return deletedCount;
        }

        /// <summary>
        /// 查询配套附件 (根据品牌、名称及 param1='附件' 与 param2 型号适配查询)
        /// </summary>
        public static List<ComponentApiDto> GetAttachments(string? brand, string? name, string? model)
        {
            var result = new List<ComponentApiDto>();
            try
            {
                string connStr = GetConnectionString();
                lock (_dbLock)
                {
                    using var conn = new SQLiteConnection(connStr);
                    conn.Open();

                    string sql = @"
                        SELECT id, brand, name, model, price, remark, param1, param2, current, poles, tripping 
                        FROM components 
                        WHERE brand = @brand AND param1 = '附件' 
                        ORDER BY id DESC;
                    "; // --硬编码-- 附件筛选条件

                    using var cmd = new SQLiteCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@brand", brand?.Trim() ?? string.Empty);

                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var item = MapReaderToDto(reader);
                        // 若指定了本体型号且附件 param2 声明了适用型号，则进行内存包含比对
                        if (!string.IsNullOrWhiteSpace(model) && !string.IsNullOrWhiteSpace(item.Param2))
                        {
                            var allowedModels = item.Param2!.Split(new[] { ',', ';', '，', '；' }, StringSplitOptions.RemoveEmptyEntries);
                            if (allowedModels.Any(m => model!.IndexOf(m.Trim(), StringComparison.OrdinalIgnoreCase) >= 0))
                            {
                                result.Add(item);
                            }
                        }
                        else
                        {
                            result.Add(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[PersonalDb] GetAttachments 异常: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 从 IDataReader 安全映射为 ComponentApiDto 实体
        /// </summary>
        private static ComponentApiDto MapReaderToDto(IDataReader reader)
        {
            return new ComponentApiDto
            {
                Id = reader.IsDBNull(0) ? 0 : Convert.ToInt32(reader.GetInt64(0)),
                Brand = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                Name = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                Model = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                Price = reader.IsDBNull(4) ? 0.00m : Convert.ToDecimal(reader.GetDouble(4)),
                Remark = reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                Param1 = reader.IsDBNull(6) ? string.Empty : reader.GetString(6),
                Param2 = reader.IsDBNull(7) ? string.Empty : reader.GetString(7),
                Current = reader.IsDBNull(8) ? null : (int?)reader.GetInt32(8),
                Poles = reader.IsDBNull(9) ? string.Empty : reader.GetString(9),
                Tripping = reader.IsDBNull(10) ? string.Empty : reader.GetString(10)
            };
        }

        /// <summary>
        /// 从电流字符串中提取纯数字整型 (如 "32A", "100A", "32" ➔ 32)
        /// </summary>
        private static int? ExtractIntegerCurrent(string? currentStr)
        {
            if (string.IsNullOrWhiteSpace(currentStr)) return null;
            string clean = currentStr.Trim().ToUpper();
            if (clean.EndsWith("A"))
            {
                clean = clean.Substring(0, clean.Length - 1).Trim();
            }
            if (int.TryParse(clean, out int curVal)) return curVal;
            if (double.TryParse(clean, out double dVal)) return (int)Math.Round(dVal);
            return null;
        }

        /// <summary>
        /// 规范化极数入参 (数据库中存储的是 "4", "3", "2", "1" 等，剥离末尾的 P 字符)
        /// </summary>
        private static string NormalizePolesParam(string? poleStr)
        {
            if (string.IsNullOrWhiteSpace(poleStr)) return string.Empty;
            string clean = poleStr.Trim().ToUpper();
            if (clean.EndsWith("P") && clean.Length > 1)
            {
                clean = clean.Substring(0, clean.Length - 1).Trim();
            }
            return clean;
        }
    }
}
