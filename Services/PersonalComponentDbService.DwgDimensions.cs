using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// 本地 SQLite 个人数据库服务分部类：DWG 元器件三维物理尺寸与指纹缓存数据访问
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，配置与硬编码显式标明
    /// </summary>
    public static partial class PersonalComponentDbService
    {
        /// <summary>
        /// 格式化/规范化相对路径索引键 (统一正斜杠、全小写，避免大小写或反斜杠差异)
        /// </summary>
        /// <param name="rawPath">原始相对路径或图纸名</param>
        /// <returns>标准化 Key</returns>
        public static string NormalizeDwgKey(string? rawPath)
        {
            // 校验入参有效性
            if (string.IsNullOrWhiteSpace(rawPath)) return string.Empty;
            // 统一替换 Windows 反斜杠为正斜杠并去除首尾空白
            string clean = rawPath.Trim().Replace('\\', '/');
            // 去除开头的斜杠
            if (clean.StartsWith("/")) clean = clean.Substring(1);
            // 统一小写以支持跨系统与环境无感匹配
            return clean.ToLowerInvariant();
        }

        /// <summary>
        /// 保存或更新单个 DWG 元器件三维尺寸与版本指纹
        /// </summary>
        /// <param name="item">尺寸数据实体</param>
        /// <returns>操作是否成功</returns>
        public static bool SaveOrUpdateDwgDimension(DwgDimensionItem item)
        {
            // 校验实体对象有效性
            if (item == null || string.IsNullOrWhiteSpace(item.RelPath)) return false;

            // 加锁保障多线程写入安全
            lock (_dbLock)
            {
                try
                {
                    // 获取数据库连接字符串
                    string connStr = GetConnectionString();
                    // 建立物理数据库连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 规范化唯一键
                    string normKey = NormalizeDwgKey(item.RelPath);
                    // 获取当前更新时间
                    string nowStr = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    // 编写插入或更新 UPSERT 语句 (基于 SQLite 的 INSERT OR REPLACE 语法)
                    string sql = @"
                        INSERT OR REPLACE INTO dwg_component_dimensions (
                            dir_name, dwg_name, rel_path, width, height, depth,
                            has_text_depth, matched_text, last_modified_ticks, file_size_bytes, updated_at
                        ) VALUES (
                            @dirName, @dwgName, @relPath, @width, @height, @depth,
                            @hasTextDepth, @matchedText, @lastMod, @fileSize, @updatedAt
                        );"; // --硬编码-- SQL 插入或替换语句

                    // 创建 SQL 参数化命令
                    using var cmd = new SQLiteCommand(sql, conn);
                    // 绑定各字段入参
                    cmd.Parameters.AddWithValue("@dirName", item.DirName ?? string.Empty);
                    cmd.Parameters.AddWithValue("@dwgName", item.DwgName ?? string.Empty);
                    cmd.Parameters.AddWithValue("@relPath", normKey);
                    cmd.Parameters.AddWithValue("@width", Math.Round(item.Width, 1));
                    cmd.Parameters.AddWithValue("@height", Math.Round(item.Height, 1));
                    cmd.Parameters.AddWithValue("@depth", Math.Round(item.Depth, 1));
                    cmd.Parameters.AddWithValue("@hasTextDepth", item.HasTextDepth ? 1 : 0);
                    cmd.Parameters.AddWithValue("@matchedText", item.MatchedText ?? string.Empty);
                    cmd.Parameters.AddWithValue("@lastMod", item.LastModifiedTicks);
                    cmd.Parameters.AddWithValue("@fileSize", item.FileSizeBytes);
                    cmd.Parameters.AddWithValue("@updatedAt", nowStr);

                    // 执行语句
                    cmd.ExecuteNonQuery();
                    return true;
                }
                catch (Exception ex)
                {
                    // 记录保存异常日志
                    LogHelper.WriteLog($"[PersonalDb] SaveOrUpdateDwgDimension 异常 {item.RelPath}: {ex.Message}");
                    return false;
                }
            }
        }

        /// <summary>
        /// 批量保存或更新 DWG 元器件三维尺寸 (启用事务，千条数据毫秒级入库)
        /// </summary>
        /// <param name="items">尺寸数据列表</param>
        /// <returns>成功写入的条数</returns>
        public static int BatchSaveOrUpdateDwgDimensions(List<DwgDimensionItem> items)
        {
            // 校验列表有效性
            if (items == null || items.Count == 0) return 0;

            // 加锁保障多线程写入安全
            lock (_dbLock)
            {
                try
                {
                    // 获取数据库连接字符串
                    string connStr = GetConnectionString();
                    // 建立物理连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 开启事务极速批量入库
                    using var trans = conn.BeginTransaction();
                    // 当前更新时间
                    string nowStr = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    // 编写批量 UPSERT 语句
                    string sql = @"
                        INSERT OR REPLACE INTO dwg_component_dimensions (
                            dir_name, dwg_name, rel_path, width, height, depth,
                            has_text_depth, matched_text, last_modified_ticks, file_size_bytes, updated_at
                        ) VALUES (
                            @dirName, @dwgName, @relPath, @width, @height, @depth,
                            @hasTextDepth, @matchedText, @lastMod, @fileSize, @updatedAt
                        );"; // --硬编码-- SQL 批量插入或替换

                    // 创建命令
                    using var cmd = new SQLiteCommand(sql, conn, trans);
                    // 预编译参数
                    var pDir = cmd.Parameters.Add("@dirName", System.Data.DbType.String);
                    var pDwg = cmd.Parameters.Add("@dwgName", System.Data.DbType.String);
                    var pRel = cmd.Parameters.Add("@relPath", System.Data.DbType.String);
                    var pWidth = cmd.Parameters.Add("@width", System.Data.DbType.Double);
                    var pHeight = cmd.Parameters.Add("@height", System.Data.DbType.Double);
                    var pDepth = cmd.Parameters.Add("@depth", System.Data.DbType.Double);
                    var pHasDepth = cmd.Parameters.Add("@hasTextDepth", System.Data.DbType.Int32);
                    var pText = cmd.Parameters.Add("@matchedText", System.Data.DbType.String);
                    var pMod = cmd.Parameters.Add("@lastMod", System.Data.DbType.Int64);
                    var pSize = cmd.Parameters.Add("@fileSize", System.Data.DbType.Int64);
                    var pUpd = cmd.Parameters.Add("@updatedAt", System.Data.DbType.String);

                    int count = 0;
                    // 遍历所有待入库条目
                    foreach (var item in items)
                    {
                        // 忽略无效相对路径条目
                        if (string.IsNullOrWhiteSpace(item.RelPath)) continue;
                        // 赋值各参数
                        pDir.Value = item.DirName ?? string.Empty;
                        pDwg.Value = item.DwgName ?? string.Empty;
                        pRel.Value = NormalizeDwgKey(item.RelPath);
                        pWidth.Value = Math.Round(item.Width, 1);
                        pHeight.Value = Math.Round(item.Height, 1);
                        pDepth.Value = Math.Round(item.Depth, 1);
                        pHasDepth.Value = item.HasTextDepth ? 1 : 0;
                        pText.Value = item.MatchedText ?? string.Empty;
                        pMod.Value = item.LastModifiedTicks;
                        pSize.Value = item.FileSizeBytes;
                        pUpd.Value = nowStr;

                        // 执行写入
                        cmd.ExecuteNonQuery();
                        count++;
                    }

                    // 提交事务
                    trans.Commit();
                    return count;
                }
                catch (Exception ex)
                {
                    // 记录批量入库异常
                    LogHelper.WriteLog($"[PersonalDb] BatchSaveOrUpdateDwgDimensions 异常: {ex.Message}");
                    return 0;
                }
            }
        }

        /// <summary>
        /// 根据相对路径或图纸文件名检索已收录的 DWG 尺寸与指纹信息
        /// </summary>
        /// <param name="relPathOrDwgName">相对路径或图纸名称</param>
        /// <returns>DwgDimensionItem 实体或 null</returns>
        public static DwgDimensionItem? GetDwgDimension(string? relPathOrDwgName)
        {
            // 校验入参
            if (string.IsNullOrWhiteSpace(relPathOrDwgName)) return null;

            // 加锁安全查库
            lock (_dbLock)
            {
                try
                {
                    // 获取数据库连接字符串
                    string connStr = GetConnectionString();
                    // 建立物理连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 规范化 Key
                    string normKey = NormalizeDwgKey(relPathOrDwgName);
                    // 剥离出纯文件名 (用于双通道兼容匹配)
                    string pureName = Path.GetFileName(normKey);
                    // 剥离扩展名
                    string pureNameWithoutExt = pureName.EndsWith(".dwg") ? pureName.Substring(0, pureName.Length - 4) : pureName;

                    // 编写优先匹配相对路径、次级匹配文件名的高效查询 SQL
                    string sql = @"
                        SELECT id, dir_name, dwg_name, rel_path, width, height, depth,
                               has_text_depth, matched_text, last_modified_ticks, file_size_bytes, updated_at
                        FROM dwg_component_dimensions
                        WHERE rel_path = @relKey
                           OR rel_path = @pureNameDwg
                           OR dwg_name = @pureName
                           OR dwg_name = @pureNameDwg
                        ORDER BY (CASE WHEN rel_path = @relKey THEN 0 ELSE 1 END)
                        LIMIT 1;"; // --硬编码-- SQL 查询单条

                    // 创建查询命令
                    using var cmd = new SQLiteCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@relKey", normKey);
                    cmd.Parameters.AddWithValue("@pureName", pureNameWithoutExt);
                    cmd.Parameters.AddWithValue("@pureNameDwg", pureName.EndsWith(".dwg") ? pureName : pureName + ".dwg");

                    // 执行读取
                    using var reader = cmd.ExecuteReader();
                    // 若存在匹配记录
                    if (reader.Read())
                    {
                        return MapDwgDimensionFromReader(reader);
                    }
                }
                catch (Exception ex)
                {
                    // 记录查询异常日志
                    LogHelper.WriteLog($"[PersonalDb] GetDwgDimension 异常 {relPathOrDwgName}: {ex.Message}");
                }
                return null;
            }
        }

        /// <summary>
        /// 批量查询一组 DWG 图纸的尺寸信息并以字典形式返回 (Key: 规范化小写键)
        /// </summary>
        /// <param name="keys">相对路径或图纸名集合</param>
        /// <returns>规范化Key -> 实体模型字典</returns>
        public static Dictionary<string, DwgDimensionItem> BatchGetDwgDimensions(IEnumerable<string>? keys)
        {
            var result = new Dictionary<string, DwgDimensionItem>(StringComparer.OrdinalIgnoreCase);
            // 校验集合有效性
            if (keys == null) return result;

            // 过滤并规范化候选键列表
            var keyList = keys.Where(k => !string.IsNullOrWhiteSpace(k))
                              .Select(k => NormalizeDwgKey(k))
                              .Distinct(StringComparer.OrdinalIgnoreCase)
                              .ToList();

            if (keyList.Count == 0) return result;

            // 加锁安全查库
            lock (_dbLock)
            {
                try
                {
                    // 获取连接字符串
                    string connStr = GetConnectionString();
                    // 建立物理连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 构建 IN 参数化查询 (分批处理以防参数超标，500为一批) --硬编码--
                    int batchSize = 500;
                    for (int i = 0; i < keyList.Count; i += batchSize)
                    {
                        var chunk = keyList.Skip(i).Take(batchSize).ToList();
                        var paramNames = new List<string>();
                        using var cmd = new SQLiteCommand(conn);

                        // 填充参数
                        for (int j = 0; j < chunk.Count; j++)
                        {
                            string pName = $"@p{j}";
                            paramNames.Add(pName);
                            cmd.Parameters.AddWithValue(pName, chunk[j]);
                        }

                        // 构造 IN 查询 SQL
                        cmd.CommandText = $@"
                            SELECT id, dir_name, dwg_name, rel_path, width, height, depth,
                                   has_text_depth, matched_text, last_modified_ticks, file_size_bytes, updated_at
                            FROM dwg_component_dimensions
                            WHERE rel_path IN ({string.Join(",", paramNames)})
                               OR dwg_name IN ({string.Join(",", paramNames)});"; // --硬编码-- SQL 批量查询

                        // 执行批量读取
                        using var reader = cmd.ExecuteReader();
                        while (reader.Read())
                        {
                            var item = MapDwgDimensionFromReader(reader);
                            if (item != null)
                            {
                                // 以规范化 relPath 登记
                                if (!string.IsNullOrEmpty(item.RelPath))
                                {
                                    result[item.RelPath] = item;
                                }
                                // 以 dwgName 登记双向别名
                                if (!string.IsNullOrEmpty(item.DwgName) && !result.ContainsKey(item.DwgName))
                                {
                                    result[item.DwgName] = item;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 记录批量查询异常
                    LogHelper.WriteLog($"[PersonalDb] BatchGetDwgDimensions 异常: {ex.Message}");
                }
            }
            return result;
        }

        /// <summary>
        /// 从 SQLite 数据读取器中安全映射 DwgDimensionItem 实体
        /// </summary>
        private static DwgDimensionItem MapDwgDimensionFromReader(SQLiteDataReader reader)
        {
            return new DwgDimensionItem
            {
                Id = reader.GetInt64(0),
                DirName = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                DwgName = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                RelPath = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                Width = reader.IsDBNull(4) ? 0.0 : reader.GetDouble(4),
                Height = reader.IsDBNull(5) ? 0.0 : reader.GetDouble(5),
                Depth = reader.IsDBNull(6) ? 0.0 : reader.GetDouble(6),
                HasTextDepth = !reader.IsDBNull(7) && reader.GetInt32(7) == 1,
                MatchedText = reader.IsDBNull(8) ? string.Empty : reader.GetString(8),
                LastModifiedTicks = reader.IsDBNull(9) ? 0L : reader.GetInt64(9),
                FileSizeBytes = reader.IsDBNull(10) ? 0L : reader.GetInt64(10),
                UpdatedAt = reader.IsDBNull(11) ? string.Empty : reader.GetString(11)
            };
        }
    }
}
