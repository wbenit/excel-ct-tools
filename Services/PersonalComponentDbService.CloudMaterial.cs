using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Text;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// 本地个人物料库 SQLite 数据库服务分部类：支持云物料工作台四分类复合检索与载流量矩阵
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public static partial class PersonalComponentDbService
    {
        /// <summary>
        /// 多条件复合分页查询本地物料库（支持按元件、母排、柜体、辅料四分类）
        /// </summary>
        /// <param name="queryParams">查询参数模型</param>
        /// <returns>分页查询结果集</returns>
        public static HubMaterialQueryResult QueryHubMaterials(HubMaterialQueryParams queryParams)
        {
            // 初始化返回容器
            var result = new HubMaterialQueryResult();
            // 确保数据库与表结构健全
            string connStr = GetConnectionString();

            try
            {
                // 数据库加锁保障并发读写安全
                lock (_dbLock)
                {
                    // 创建并打开数据库连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 构建基础过滤条件
                    var whereBuilder = new StringBuilder(" WHERE 1=1 ");
                    // 子库类别精准筛选 (支持元件、母排、柜体、辅料)
                    if (!string.IsNullOrWhiteSpace(queryParams.Category) && queryParams.Category != "全部")
                    {
                        if (queryParams.Category == "母排")
                        {
                            // 匹配母排类物料
                            whereBuilder.Append(" AND (param2 = '母排' OR name LIKE '%铜排%' OR name LIKE '%母线%' OR name LIKE '%母排%' OR name LIKE '%铝排%' OR model LIKE 'TMY%' OR model LIKE 'LMY%') ");
                        }
                        else if (queryParams.Category == "柜体")
                        {
                            // 匹配柜体壳体类物料
                            whereBuilder.Append(" AND (param2 = '柜体' OR name LIKE '%柜体%' OR name LIKE '%箱体%' OR name LIKE '%壳体%') ");
                        }
                        else if (queryParams.Category == "辅料")
                        {
                            // 匹配辅料类物料
                            whereBuilder.Append(" AND (param2 = '辅料' OR name LIKE '%端子%' OR name LIKE '%绝缘%' OR name LIKE '%导轨%' OR name LIKE '%扎带%' OR name LIKE '%标牌%') ");
                        }
                        else if (queryParams.Category == "元件")
                        {
                            // 匹配常规元器件（排除母排、柜体和辅料）
                            whereBuilder.Append(" AND (param2 = '元件' OR (ifnull(param2, '') NOT IN ('母排', '柜体', '辅料') AND name NOT LIKE '%铜排%' AND name NOT LIKE '%母线%' AND name NOT LIKE '%柜体%' AND name NOT LIKE '%箱体%' AND name NOT LIKE '%壳体%')) ");
                        }
                    }

                    // 型号或编码模糊检索
                    if (!string.IsNullOrWhiteSpace(queryParams.ModelKeyword))
                    {
                        // 检索型号列
                        whereBuilder.Append(" AND (model LIKE @modelKw OR remark LIKE @modelKw) ");
                    }

                    // 名称模糊检索
                    if (!string.IsNullOrWhiteSpace(queryParams.NameKeyword))
                    {
                        // 检索名称列
                        whereBuilder.Append(" AND name LIKE @nameKw ");
                    }

                    // 品牌精准筛选
                    if (!string.IsNullOrWhiteSpace(queryParams.Brand) && queryParams.Brand != "不限")
                    {
                        // 检索品牌列
                        whereBuilder.Append(" AND brand = @brand ");
                    }

                    // 1. 统计符合条件的总记录条数
                    string countSql = "SELECT COUNT(*) FROM components " + whereBuilder.ToString();
                    // 创建统计 SQL 命令
                    using (var countCmd = new SQLiteCommand(countSql, conn))
                    {
                        // 绑定参数
                        BindHubQueryParams(countCmd, queryParams);
                        // 获取总数
                        result.TotalCount = Convert.ToInt32(countCmd.ExecuteScalar());
                    }

                    // 如果总记录数为 0，直接返回空结果集
                    if (result.TotalCount == 0)
                    {
                        return result;
                    }

                    // 2. 执行分页数据检索
                    int offset = Math.Max(0, (queryParams.PageIndex - 1) * queryParams.PageSize);
                    // 构建分页查询语句
                    string dataSql = $@"
                        SELECT id, brand, original_id, name, model, price, remark, param1, param2, current, poles, tripping 
                        FROM components 
                        {whereBuilder} 
                        ORDER BY id DESC 
                        LIMIT @limit OFFSET @offset;
                    "; // --硬编码: 分页查询SQL--

                    // 创建数据提取命令
                    using var dataCmd = new SQLiteCommand(dataSql, conn);
                    // 绑定过滤参数
                    BindHubQueryParams(dataCmd, queryParams);
                    // 绑定分页限制参数
                    dataCmd.Parameters.AddWithValue("@limit", queryParams.PageSize);
                    // 绑定分页偏移量参数
                    dataCmd.Parameters.AddWithValue("@offset", offset);

                    // 流式读取记录
                    using var reader = dataCmd.ExecuteReader();
                    while (reader.Read())
                    {
                        // 转换为 HubMaterialDto
                        var item = new HubMaterialDto
                        {
                            Id = reader.GetInt32(0),
                            Brand = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                            Name = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                            Model = reader.IsDBNull(4) ? string.Empty : reader.GetString(4),
                            Price = reader.IsDBNull(5) ? 0.0m : Convert.ToDecimal(reader.GetDouble(5)),
                            Remark = reader.IsDBNull(6) ? string.Empty : reader.GetString(6),
                            Param1 = reader.IsDBNull(7) ? string.Empty : reader.GetString(7),
                            Param2 = reader.IsDBNull(8) ? string.Empty : reader.GetString(8),
                            Current = reader.IsDBNull(9) ? (int?)null : reader.GetInt32(9),
                            Poles = reader.IsDBNull(10) ? string.Empty : reader.GetString(10),
                            // 本地库物料标记为已认证或未认证 (original_id > 0 视为云端认证下发)
                            IsCertified = !reader.IsDBNull(2) && reader.GetInt32(2) > 0,
                            // 根据名称与分类推导计量单位
                            Unit = DeriveMaterialUnit(reader.IsDBNull(3) ? "" : reader.GetString(3), reader.IsDBNull(8) ? "" : reader.GetString(8)),
                            // 确定归属分类
                            Category = DeriveMaterialCategory(reader.IsDBNull(3) ? "" : reader.GetString(3), reader.IsDBNull(4) ? "" : reader.GetString(4), reader.IsDBNull(8) ? "" : reader.GetString(8))
                        };
                        result.Items.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录检索异常日志
                LogHelper.WriteLog($"[PersonalDb] QueryHubMaterials 检索异常: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 绑定查询参数到 SQLiteCommand
        /// </summary>
        private static void BindHubQueryParams(SQLiteCommand cmd, HubMaterialQueryParams p)
        {
            // 绑定分类参数
            if (!string.IsNullOrWhiteSpace(p.Category) && p.Category != "全部")
            {
                cmd.Parameters.AddWithValue("@cat", p.Category);
                cmd.Parameters.AddWithValue("@catName", $"%{p.Category}%");
                cmd.Parameters.AddWithValue("@catModel", $"%{p.Category}%");
            }
            // 绑定型号模糊参数
            if (!string.IsNullOrWhiteSpace(p.ModelKeyword))
            {
                cmd.Parameters.AddWithValue("@modelKw", $"%{p.ModelKeyword.Trim()}%");
            }
            // 绑定名称模糊参数
            if (!string.IsNullOrWhiteSpace(p.NameKeyword))
            {
                cmd.Parameters.AddWithValue("@nameKw", $"%{p.NameKeyword.Trim()}%");
            }
            // 绑定品牌参数
            if (!string.IsNullOrWhiteSpace(p.Brand) && p.Brand != "不限")
            {
                cmd.Parameters.AddWithValue("@brand", p.Brand.Trim());
            }
        }

        /// <summary>
        /// 保存从工作台（柜体计算器/母排计算器）新增或导入的物料到本地 SQLite
        /// </summary>
        /// <param name="item">待入库物料实体</param>
        /// <returns>新增或更新后的物料自增 ID</returns>
        public static int SaveHubMaterial(HubMaterialDto item)
        {
            // 确保数据库连接就绪
            string connStr = GetConnectionString();

            try
            {
                lock (_dbLock)
                {
                    // 创建数据库连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();

                    // 查询同品牌下同型号是否已存在
                    string checkSql = "SELECT id FROM components WHERE UPPER(TRIM(brand)) = UPPER(TRIM(@brand)) AND UPPER(TRIM(model)) = UPPER(TRIM(@model)) LIMIT 1;";
                    using var checkCmd = new SQLiteCommand(checkSql, conn);
                    // 绑定品牌
                    checkCmd.Parameters.AddWithValue("@brand", item.Brand.Trim());
                    // 绑定型号
                    checkCmd.Parameters.AddWithValue("@model", item.Model.Trim());

                    // 执行查验
                    object? existingIdObj = checkCmd.ExecuteScalar();
                    if (existingIdObj != null && existingIdObj != DBNull.Value)
                    {
                        int existingId = Convert.ToInt32(existingIdObj);
                        // 已存在则更新单价与扩展属性
                        string updateSql = @"
                            UPDATE components 
                            SET name = @name, price = @price, remark = @remark, param1 = @param1, param2 = @param2, updated_at = @now 
                            WHERE id = @id;
                        ";
                        using var upCmd = new SQLiteCommand(updateSql, conn);
                        upCmd.Parameters.AddWithValue("@name", item.Name.Trim());
                        upCmd.Parameters.AddWithValue("@price", Convert.ToDouble(item.Price));
                        upCmd.Parameters.AddWithValue("@remark", item.Remark ?? string.Empty);
                        upCmd.Parameters.AddWithValue("@param1", item.Param1 ?? string.Empty);
                        upCmd.Parameters.AddWithValue("@param2", item.Category ?? string.Empty);
                        upCmd.Parameters.AddWithValue("@now", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        upCmd.Parameters.AddWithValue("@id", existingId);
                        upCmd.ExecuteNonQuery();
                        return existingId;
                    }

                    // 不存在则执行新增插入
                    string insertSql = @"
                        INSERT INTO components (brand, name, model, price, remark, param1, param2, current, poles, tripping, created_at, updated_at) 
                        VALUES (@brand, @name, @model, @price, @remark, @param1, @param2, @current, @poles, '', @now, @now);
                        SELECT last_insert_rowid();
                    ";
                    using var inCmd = new SQLiteCommand(insertSql, conn);
                    inCmd.Parameters.AddWithValue("@brand", string.IsNullOrWhiteSpace(item.Brand) ? "通用" : item.Brand.Trim());
                    inCmd.Parameters.AddWithValue("@name", string.IsNullOrWhiteSpace(item.Name) ? item.Category : item.Name.Trim());
                    inCmd.Parameters.AddWithValue("@model", item.Model.Trim());
                    inCmd.Parameters.AddWithValue("@price", Convert.ToDouble(item.Price));
                    inCmd.Parameters.AddWithValue("@remark", item.Remark ?? string.Empty);
                    inCmd.Parameters.AddWithValue("@param1", item.Param1 ?? string.Empty);
                    inCmd.Parameters.AddWithValue("@param2", item.Category ?? string.Empty);
                    inCmd.Parameters.AddWithValue("@current", item.Current.HasValue ? (object)item.Current.Value : DBNull.Value);
                    inCmd.Parameters.AddWithValue("@poles", item.Poles ?? string.Empty);
                    inCmd.Parameters.AddWithValue("@now", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                    // 执行并返回新生成的自增主键 ID
                    object newId = inCmd.ExecuteScalar();
                    return Convert.ToInt32(newId);
                }
            }
            catch (Exception ex)
            {
                // 记录保存异常日志
                LogHelper.WriteLog($"[PersonalDb] SaveHubMaterial 异常: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 根据物料名称与分类自动推导计量单位
        /// </summary>
        private static string DeriveMaterialUnit(string name, string param2)
        {
            if (param2 == "母排" || name.Contains("铜排") || name.Contains("母线") || name.Contains("铝排"))
                return "米";
            if (param2 == "柜体" || name.Contains("柜体") || name.Contains("箱体") || name.Contains("壳体"))
                return "台";
            if (name.Contains("线") || name.Contains("软管"))
                return "米";
            return "个";
        }

        /// <summary>
        /// 自动推导物料归属子分类 (元件/母排/柜体/辅料)
        /// </summary>
        private static string DeriveMaterialCategory(string name, string model, string param2)
        {
            if (!string.IsNullOrWhiteSpace(param2) && (param2 == "元件" || param2 == "母排" || param2 == "柜体" || param2 == "辅料"))
            {
                return param2;
            }
            if (name.Contains("柜体") || name.Contains("箱体") || name.Contains("壳体") || model.Contains("*") && model.Split('*').Length >= 3)
            {
                return "柜体";
            }
            if (name.Contains("铜排") || name.Contains("母线") || name.Contains("母排") || model.StartsWith("TMY") || model.StartsWith("LMY"))
            {
                return "母排";
            }
            if (name.Contains("端子") || name.Contains("绝缘") || name.Contains("导轨") || name.Contains("扎带") || name.Contains("标牌"))
            {
                return "辅料";
            }
            return "元件";
        }

        /// <summary>
        /// 获取国标母线完整载流量速查数据集 (GB/T 5585 标准规格，依据当前铜价实时换算单价)
        /// </summary>
        /// <param name="copperPricePerTon">当前大盘行情铜价 (元/吨，默认 111465)</param>
        /// <param name="isCopper">true=铜排 TMY，false=铝排 LMY</param>
        /// <returns>国标载流量矩阵明细列表</returns>
        public static List<BusbarAmpacityItem> GetBusbarAmpacityTable(decimal copperPricePerTon, bool isCopper = true)
        {
            // 初始化载流量基础数据集合
            var list = new List<BusbarAmpacityItem>();
            // 基础换算价格: 元/kg
            decimal basePerKg = copperPricePerTon / 1000.0m;

            // 原始国标标准母排截面与载流量基准表 (包含序号、宽度、厚度、25/35/40度平放与竖放电流)
            var rawData = new (int idx, double w, double t, double a25f, double a25v, double a35f, double a35v, double a40f, double a40v)[]
            {
                (1,  15, 3,  200, 210, 176, 185, 162, 171),
                (2,  20, 3,  261, 275, 233, 245, 214, 225),
                (3,  25, 3,  323, 340, 285, 300, 271, 286),
                (4,  30, 3,  366, 385, 322, 339, 296, 312),
                (5,  30, 4,  451, 475, 394, 415, 366, 385),
                (6,  40, 4,  593, 625, 522, 550, 484, 510),
                (7,  40, 5,  665, 700, 584, 615, 551, 580),
                (8,  40, 6,  765, 805, 673, 708, 619, 652),
                (9,  40, 8,  912, 960, 803, 845, 739, 778),
                (10, 40, 10, 1064, 1120, 936, 986, 862, 907),
                (11, 50, 5,  816, 860, 721, 760, 669, 705),
                (12, 50, 6,  906, 955, 797, 840, 735, 775),
                (13, 50, 8,  1078, 1135, 949, 999, 873, 919),
                (14, 50, 10, 1254, 1320, 1104, 1162, 1016, 1069),
                (15, 60, 6,  1069, 1125, 940, 990, 873, 920),
                (16, 60, 8,  1251, 1320, 1101, 1160, 1016, 1070),
                (17, 60, 10, 1395, 1475, 1230, 1295, 1133, 1195),
                (18, 80, 6,  1360, 1480, 1195, 1300, 1110, 1205),
                (19, 80, 8,  1553, 1690, 1361, 1480, 1260, 1370),
                (20, 80, 10, 1747, 1900, 1531, 1665, 1417, 1540),
                (21, 100, 6, 1665, 1810, 1557, 1592, 1356, 1475),
                (22, 100, 8, 1911, 2080, 1674, 1820, 1546, 1685),
                (23, 100, 10, 2121, 2310, 1885, 2025, 1720, 1870),
                (24, 120, 10, 2470, 2690, 2170, 2360, 2000, 2180)
            }; // --硬编码: GB/T 5585 标准规格载流量矩阵--

            // 循环遍历组装条目
            foreach (var r in rawData)
            {
                // 计算理论米重 (紫铜 8.9，铝排 2.7)
                double density = isCopper ? 0.0089 : 0.0027;
                double meterWeight = Math.Round(r.w * r.t * density, 4);
                // 动态换算单价 (元/米)
                decimal unitPrice = Math.Round((decimal)meterWeight * basePerKg, 6);

                // 添加到载流量清单
                list.Add(new BusbarAmpacityItem
                {
                    Index = r.idx,
                    Spec = $"{r.w:0}*{r.t:0}",
                    Width = r.w,
                    Thickness = r.t,
                    Amp25Flat = isCopper ? r.a25f : Math.Round(r.a25f * 0.78, 0),
                    Amp25Vertical = isCopper ? r.a25v : Math.Round(r.a25v * 0.78, 0),
                    Amp35Flat = isCopper ? r.a35f : Math.Round(r.a35f * 0.78, 0),
                    Amp35Vertical = isCopper ? r.a35v : Math.Round(r.a35v * 0.78, 0),
                    Amp40Flat = isCopper ? r.a40f : Math.Round(r.a40f * 0.78, 0),
                    Amp40Vertical = isCopper ? r.a40v : Math.Round(r.a40v * 0.78, 0),
                    UnitPrice = unitPrice
                });
            }

            return list;
        }
    }
}
