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
    /// 本地 SQLite 个人数据库服务分部类：一次成套方案与 BOM 数据访问
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，配置与硬编码显式标明
    /// </summary>
    public static partial class PersonalComponentDbService
    {
        // 声明一次方案通用的 JSON 序列化选项
        private static readonly JsonSerializerOptions PrimJsonOptions = new JsonSerializerOptions
        {
            // 属性名忽略大小写
            PropertyNameCaseInsensitive = true,
            // 支持中文直出
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            // 紧凑输出
            WriteIndented = false
        };

        /// <summary>
        /// 安全序列化一次方案 BOM 列表为 JSON 字符串
        /// </summary>
        private static string SafeSerializePrimaryBomList(List<CloudSchemeBomItem>? items)
        {
            if (items == null || items.Count == 0) return "[]";
            try
            {
                // 序列化为 JSON
                return JsonSerializer.Serialize(items, PrimJsonOptions);
            }
            catch (Exception ex)
            {
                // 记录异常日志并降级返回空数组
                LogHelper.WriteLog($"[PrimaryDb] SafeSerializePrimaryBomList 异常: {ex.Message}");
                return "[]";
            }
        }

        /// <summary>
        /// 安全反序列化 JSON 字符串为一次方案 BOM 列表
        /// </summary>
        private static List<CloudSchemeBomItem> SafeDeserializePrimaryBomList(string? json)
        {
            if (string.IsNullOrWhiteSpace(json)) return new List<CloudSchemeBomItem>();
            try
            {
                // 反序列化为 BOM 列表
                return JsonSerializer.Deserialize<List<CloudSchemeBomItem>>(json, PrimJsonOptions) ?? new List<CloudSchemeBomItem>();
            }
            catch (Exception ex)
            {
                // 记录异常日志并降级返回空列表
                LogHelper.WriteLog($"[PrimaryDb] SafeDeserializePrimaryBomList 异常: {ex.Message}");
                return new List<CloudSchemeBomItem>();
            }
        }

        /// <summary>
        /// 查询所有一次成套方案配置列表 (支持按关键字、分类目录和柜型多维模糊检索)
        /// </summary>
        public static List<PrimarySchemeEntity> GetAllPrimarySchemes(string? keyword = null, string? groupName = null, string? cabinetModel = null)
        {
            var list = new List<PrimarySchemeEntity>();

            try
            {
                // 获取数据库连接字符串
                string connStr = GetConnectionString();

                lock (_dbLock)
                {
                    // 创建并打开数据库物理连接
                    using var conn = new SQLiteConnection(connStr);
                    conn.Open();

                    // 构建参数化查询 SQL 语句
                    var sb = new StringBuilder(@"
                        SELECT id, group_name, scheme_name, applicable_codes, cad_drawing_name, 
                               cabinet_model, dimensions, rated_current, busbar_spec, labor_cost, copper_cost, 
                               brand, description, bom_json, created_at, updated_at
                        FROM primary_circuit_schemes
                        WHERE 1=1 
                    ");

                    using var cmd = new SQLiteCommand(conn);

                    // 1. 若指定了分类目录则筛选
                    if (!string.IsNullOrWhiteSpace(groupName) && !string.Equals(groupName, "全部", StringComparison.OrdinalIgnoreCase))
                    {
                        sb.Append(" AND group_name = @group_name ");
                        cmd.Parameters.AddWithValue("@group_name", groupName.Trim());
                    }

                    // 2. 若指定了柜型则筛选
                    if (!string.IsNullOrWhiteSpace(cabinetModel) && !string.Equals(cabinetModel, "全部", StringComparison.OrdinalIgnoreCase))
                    {
                        sb.Append(" AND cabinet_model LIKE @cabinet_model ");
                        cmd.Parameters.AddWithValue("@cabinet_model", $"%{cabinetModel.Trim()}%");
                    }

                    // 3. 若指定了关键字则模糊匹配
                    if (!string.IsNullOrWhiteSpace(keyword))
                    {
                        string kw = keyword.Trim();
                        sb.Append(@" AND (
                            scheme_name LIKE @kw 
                            OR applicable_codes LIKE @kw 
                            OR cad_drawing_name LIKE @kw 
                            OR cabinet_model LIKE @kw 
                            OR brand LIKE @kw 
                            OR busbar_spec LIKE @kw 
                            OR description LIKE @kw
                        ) ");
                        cmd.Parameters.AddWithValue("@kw", $"%{kw}%");
                    }

                    // 按更新时间降序排列
                    sb.Append(" ORDER BY updated_at DESC, id DESC; ");
                    cmd.CommandText = sb.ToString();

                    // 执行查询读取实体数据
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var entity = ReadPrimarySchemeFromReader(reader);
                        if (entity != null)
                        {
                            list.Add(entity);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录检索异常日志
                LogHelper.WriteLog($"[PrimaryDb] GetAllPrimarySchemes 异常: {ex.Message}");
            }

            return list;
        }

        /// <summary>
        /// 根据主键 ID 获取单个一次方案完整实体
        /// </summary>
        public static PrimarySchemeEntity? GetPrimarySchemeById(int id)
        {
            if (id <= 0) return null;

            try
            {
                string connStr = GetConnectionString();
                lock (_dbLock)
                {
                    using var conn = new SQLiteConnection(connStr);
                    conn.Open();

                    string sql = @"
                        SELECT id, group_name, scheme_name, applicable_codes, cad_drawing_name, 
                               cabinet_model, dimensions, rated_current, busbar_spec, labor_cost, copper_cost, 
                               brand, description, bom_json, created_at, updated_at
                        FROM primary_circuit_schemes
                        WHERE id = @id LIMIT 1;
                    ";

                    using var cmd = new SQLiteCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@id", id);

                    using var reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        return ReadPrimarySchemeFromReader(reader);
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[PrimaryDb] GetPrimarySchemeById 异常: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// 根据图纸名称 (去后缀) 精准匹配一次成套方案实体
        /// </summary>
        public static PrimarySchemeEntity? FindPrimarySchemeByDwgName(string dwgName)
        {
            if (string.IsNullOrWhiteSpace(dwgName)) return null;

            string cleanName = dwgName.Trim();
            // 读取全量一次方案列表
            var allSchemes = GetAllPrimarySchemes();
            if (allSchemes.Count == 0) return null;

            // 1. 优先在 applicable_codes 逗号列表中精确匹配
            foreach (var s in allSchemes)
            {
                if (s.ApplicableCodes != null && s.ApplicableCodes.Any(code => string.Equals(code.Trim(), cleanName, StringComparison.OrdinalIgnoreCase)))
                {
                    return s;
                }
            }

            // 2. 次选 CAD 图纸名称精确匹配
            var matchCad = allSchemes.FirstOrDefault(s => string.Equals(s.CadDrawingName?.Trim(), cleanName, StringComparison.OrdinalIgnoreCase));
            if (matchCad != null) return matchCad;

            // 3. 最后按方案名称精确匹配
            var matchName = allSchemes.FirstOrDefault(s => string.Equals(s.SchemeName?.Trim(), cleanName, StringComparison.OrdinalIgnoreCase));
            if (matchName != null) return matchName;

            return null;
        }

        /// <summary>
        /// 保存或更新一次成套方案 (新增返回自增 ID，更新返回已有 ID)
        /// </summary>
        public static int SavePrimaryScheme(PrimarySchemeEntity scheme)
        {
            if (scheme == null || string.IsNullOrWhiteSpace(scheme.SchemeName)) return 0;

            try
            {
                string connStr = GetConnectionString();
                lock (_dbLock)
                {
                    using var conn = new SQLiteConnection(connStr);
                    conn.Open();

                    string nowText = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string bomJson = SafeSerializePrimaryBomList(scheme.BomItems);
                    string codesText = scheme.ApplicableCodes != null && scheme.ApplicableCodes.Count > 0
                        ? string.Join(",", scheme.ApplicableCodes.Select(c => c.Trim()).Where(c => !string.IsNullOrEmpty(c)))
                        : scheme.SchemeName.Trim();

                    // 1. 若 ID > 0，执行 UPDATE
                    if (scheme.Id > 0)
                    {
                        string updateSql = @"
                            UPDATE primary_circuit_schemes SET
                                group_name = @group_name,
                                scheme_name = @scheme_name,
                                applicable_codes = @applicable_codes,
                                cad_drawing_name = @cad_drawing_name,
                                cabinet_model = @cabinet_model,
                                dimensions = @dimensions,
                                rated_current = @rated_current,
                                busbar_spec = @busbar_spec,
                                labor_cost = @labor_cost,
                                copper_cost = @copper_cost,
                                brand = @brand,
                                description = @description,
                                bom_json = @bom_json,
                                updated_at = @updated_at
                            WHERE id = @id;
                        ";

                        using var cmd = new SQLiteCommand(updateSql, conn);
                        cmd.Parameters.AddWithValue("@group_name", scheme.GroupName ?? string.Empty);
                        cmd.Parameters.AddWithValue("@scheme_name", scheme.SchemeName.Trim());
                        cmd.Parameters.AddWithValue("@applicable_codes", codesText);
                        cmd.Parameters.AddWithValue("@cad_drawing_name", scheme.CadDrawingName ?? string.Empty);
                        cmd.Parameters.AddWithValue("@cabinet_model", scheme.CabinetModel ?? string.Empty);
                        cmd.Parameters.AddWithValue("@dimensions", scheme.Dimensions ?? string.Empty);
                        cmd.Parameters.AddWithValue("@rated_current", scheme.RatedCurrent);
                        cmd.Parameters.AddWithValue("@busbar_spec", scheme.BusbarSpec ?? string.Empty);
                        cmd.Parameters.AddWithValue("@labor_cost", scheme.LaborCost);
                        cmd.Parameters.AddWithValue("@copper_cost", scheme.CopperCost);
                        cmd.Parameters.AddWithValue("@brand", scheme.Brand ?? string.Empty);
                        cmd.Parameters.AddWithValue("@description", scheme.Description ?? string.Empty);
                        cmd.Parameters.AddWithValue("@bom_json", bomJson);
                        cmd.Parameters.AddWithValue("@updated_at", nowText);
                        cmd.Parameters.AddWithValue("@id", scheme.Id);

                        int rows = cmd.ExecuteNonQuery();
                        return rows > 0 ? scheme.Id : 0;
                    }
                    // 2. 若 ID <= 0，执行 INSERT
                    else
                    {
                        string insertSql = @"
                            INSERT INTO primary_circuit_schemes (
                                group_name, scheme_name, applicable_codes, cad_drawing_name,
                                cabinet_model, dimensions, rated_current, busbar_spec, labor_cost, copper_cost,
                                brand, description, bom_json, created_at, updated_at
                            ) VALUES (
                                @group_name, @scheme_name, @applicable_codes, @cad_drawing_name,
                                @cabinet_model, @dimensions, @rated_current, @busbar_spec, @labor_cost, @copper_cost,
                                @brand, @description, @bom_json, @created_at, @updated_at
                            );
                            SELECT last_insert_rowid();
                        ";

                        using var cmd = new SQLiteCommand(insertSql, conn);
                        cmd.Parameters.AddWithValue("@group_name", scheme.GroupName ?? string.Empty);
                        cmd.Parameters.AddWithValue("@scheme_name", scheme.SchemeName.Trim());
                        cmd.Parameters.AddWithValue("@applicable_codes", codesText);
                        cmd.Parameters.AddWithValue("@cad_drawing_name", scheme.CadDrawingName ?? string.Empty);
                        cmd.Parameters.AddWithValue("@cabinet_model", scheme.CabinetModel ?? string.Empty);
                        cmd.Parameters.AddWithValue("@dimensions", scheme.Dimensions ?? string.Empty);
                        cmd.Parameters.AddWithValue("@rated_current", scheme.RatedCurrent);
                        cmd.Parameters.AddWithValue("@busbar_spec", scheme.BusbarSpec ?? string.Empty);
                        cmd.Parameters.AddWithValue("@labor_cost", scheme.LaborCost);
                        cmd.Parameters.AddWithValue("@copper_cost", scheme.CopperCost);
                        cmd.Parameters.AddWithValue("@brand", scheme.Brand ?? string.Empty);
                        cmd.Parameters.AddWithValue("@description", scheme.Description ?? string.Empty);
                        cmd.Parameters.AddWithValue("@bom_json", bomJson);
                        cmd.Parameters.AddWithValue("@created_at", nowText);
                        cmd.Parameters.AddWithValue("@updated_at", nowText);

                        object? newIdObj = cmd.ExecuteScalar();
                        if (newIdObj != null && long.TryParse(newIdObj.ToString(), out long newIdLong))
                        {
                            return (int)newIdLong;
                        }
                        return 0;
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[PrimaryDb] SavePrimaryScheme 异常: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 根据主键 ID 删除指定的一次成套方案
        /// </summary>
        public static bool DeletePrimaryScheme(int id)
        {
            if (id <= 0) return false;

            try
            {
                string connStr = GetConnectionString();
                lock (_dbLock)
                {
                    using var conn = new SQLiteConnection(connStr);
                    conn.Open();

                    string sql = "DELETE FROM primary_circuit_schemes WHERE id = @id;";
                    using var cmd = new SQLiteCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@id", id);

                    return cmd.ExecuteNonQuery() > 0;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[PrimaryDb] DeletePrimaryScheme 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 从 SQLiteDataReader 中反序列化构建一次方案实体
        /// </summary>
        private static PrimarySchemeEntity? ReadPrimarySchemeFromReader(SQLiteDataReader reader)
        {
            try
            {
                int id = reader.GetInt32(reader.GetOrdinal("id"));
                string groupName = reader["group_name"]?.ToString() ?? string.Empty;
                string schemeName = reader["scheme_name"]?.ToString() ?? string.Empty;
                string codesText = reader["applicable_codes"]?.ToString() ?? string.Empty;
                string cadDwg = reader["cad_drawing_name"]?.ToString() ?? string.Empty;
                string cabModel = reader["cabinet_model"]?.ToString() ?? string.Empty;
                string dims = reader["dimensions"]?.ToString() ?? string.Empty;
                double cur = Convert.ToDouble(reader["rated_current"] != DBNull.Value ? reader["rated_current"] : 0.0);
                string busbar = reader["busbar_spec"]?.ToString() ?? string.Empty;
                double labor = Convert.ToDouble(reader["labor_cost"] != DBNull.Value ? reader["labor_cost"] : 0.0);
                double copper = Convert.ToDouble(reader["copper_cost"] != DBNull.Value ? reader["copper_cost"] : 0.0);
                string brand = reader["brand"]?.ToString() ?? string.Empty;
                string desc = reader["description"]?.ToString() ?? string.Empty;
                string bomJson = reader["bom_json"]?.ToString() ?? "[]";
                string created = reader["created_at"]?.ToString() ?? string.Empty;
                string updated = reader["updated_at"]?.ToString() ?? string.Empty;

                var codesList = new List<string>();
                if (!string.IsNullOrWhiteSpace(codesText))
                {
                    codesList = codesText.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                                         .Select(c => c.Trim())
                                         .Where(c => !string.IsNullOrEmpty(c))
                                         .ToList();
                }

                var bomItems = SafeDeserializePrimaryBomList(bomJson);

                return new PrimarySchemeEntity
                {
                    Id = id,
                    GroupName = groupName,
                    SchemeName = schemeName,
                    ApplicableCodes = codesList,
                    CadDrawingName = cadDwg,
                    CabinetModel = cabModel,
                    Dimensions = dims,
                    RatedCurrent = cur,
                    BusbarSpec = busbar,
                    LaborCost = labor,
                    CopperCost = copper,
                    Brand = brand,
                    Description = desc,
                    BomItems = bomItems,
                    CreatedAt = created,
                    UpdatedAt = updated
                };
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[PrimaryDb] ReadPrimarySchemeFromReader 异常: {ex.Message}");
                return null;
            }
        }
    }
}
