using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// 本地个人物料库 SQLite 数据库服务分部类：专门处理分类明细元器件批量排重核对与一键入库/覆盖更新
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，硬编码显式标明
    /// </summary>
    public static partial class PersonalComponentDbService
    {
        /// <summary>
        /// 批量核对候选元器件在本地个人库中的重名与价格冲突状态
        /// </summary>
        /// <param name="candidates">待核对的元器件候选列表</param>
        public static void CheckAndMatchCandidates(List<AddToPersonalDbCandidateItem> candidates)
        {
            // 校验入参列表有效性
            if (candidates == null || candidates.Count == 0) return;

            try
            {
                // 获取数据库连接字符串并确保表结构已初始化
                string connStr = GetConnectionString();

                // 提取候选项中所有涉及的非空品牌列表 (用于分块精准加载加速，避免全表盲扫)
                var brands = candidates
                    .Select(c => (c.Brand?.Trim() ?? string.Empty).ToLowerInvariant())
                    .Where(b => !string.IsNullOrEmpty(b))
                    .Distinct()
                    .ToList();

                // 内存字典映射：Key为 "brand##model" (小写标准键)，Value为现有元器件实体
                var existingCompMap = new Dictionary<string, (int id, string brand, string name, string model, decimal price)>(StringComparer.OrdinalIgnoreCase);

                // 加数据库安全互斥锁保障并发一致性
                lock (_dbLock)
                {
                    // 创建数据库物理连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开底层物理连接
                    conn.Open();

                    // 构建查询 SQL：若品牌数较少直接按品牌过滤，否则查询全部物料精简字段
                    string querySql = @"
                        SELECT id, brand, name, model, price 
                        FROM components;
                    "; // --硬编码: 基础检索SQL--

                    // 创建命令执行对象
                    using var cmd = new SQLiteCommand(querySql, conn);
                    // 执行数据读取流
                    using var reader = cmd.ExecuteReader();

                    // 循环读取每一条现有物料记录
                    while (reader.Read())
                    {
                        // 读取主键 ID
                        int id = Convert.ToInt32(reader.GetInt64(0));
                        // 读取品牌
                        string b = reader.IsDBNull(1) ? string.Empty : reader.GetString(1).Trim();
                        // 读取名称
                        string n = reader.IsDBNull(2) ? string.Empty : reader.GetString(2).Trim();
                        // 读取规格型号
                        string m = reader.IsDBNull(3) ? string.Empty : reader.GetString(3).Trim();
                        // 读取单价
                        decimal p = reader.IsDBNull(4) ? 0.0m : Convert.ToDecimal(reader.GetDouble(4));

                        // 仅当型号有效非空时纳入字典
                        if (!string.IsNullOrEmpty(m))
                        {
                            // 构建唯一复合键
                            string key = BuildCompMatchKey(b, m);
                            // 若字典中尚未收录则记录
                            if (!existingCompMap.ContainsKey(key))
                            {
                                existingCompMap[key] = (id, b, n, m, p);
                            }
                        }
                    }
                }

                // 遍历每一个候选条目，进行状态核对与标记
                foreach (var item in candidates)
                {
                    // 获取当前条目的品牌与型号
                    string brand = item.Brand?.Trim() ?? string.Empty;
                    string model = item.Model?.Trim() ?? string.Empty;

                    // 若型号为空，标记为非法不可用
                    if (string.IsNullOrEmpty(model))
                    {
                        item.Status = "Invalid"; // --硬编码: 无效状态--
                        item.StatusText = "缺少型号(无法入库)"; // --硬编码: 状态文案--
                        item.IsSelected = false;
                        continue;
                    }

                    // 构建匹配键
                    string matchKey = BuildCompMatchKey(brand, model);

                    // 在现有数据库物料映射表中检索
                    if (existingCompMap.TryGetValue(matchKey, out var matched))
                    {
                        // 记录库中原数据的主键 ID 与单价
                        item.OldId = matched.id;
                        item.OldPrice = matched.price;

                        // 比对当前表格单价与库中原单价差异 (绝对值大于 0.001 判定为价格变动)
                        if (Math.Abs(item.Price - matched.price) > 0.001m)
                        {
                            item.Status = "PriceChanged"; // --硬编码: 价格变动状态--
                            item.StatusText = $"价格变动 (原价: {matched.price:0.##})"; // --硬编码: 状态文案--
                        }
                        else
                        {
                            item.Status = "ExactMatch"; // --硬编码: 完全一致状态--
                            item.StatusText = "完全一致 (库中已存在)"; // --硬编码: 状态文案--
                        }
                    }
                    else
                    {
                        // 库中不存在同品牌同型号，判定为全新物料
                        item.Status = "New"; // --硬编码: 全新状态--
                        item.StatusText = "全新物料 (新增入库)"; // --硬编码: 状态文案--
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录排重检测异常日志
                LogHelper.WriteLog($"[PersonalDb] CheckAndMatchCandidates 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 批量执行元器件入库或新价格覆盖更新 (采用 SQLite 极速事务批处理)
        /// </summary>
        /// <param name="items">待入库的元器件条目列表</param>
        /// <param name="overwritePrice">遇到同品牌同型号已存在时，是否用表格中的新单价覆盖更新</param>
        /// <returns>执行结果汇总统计</returns>
        public static AddToPersonalDbResult BatchSaveOrUpdatePersonalComponents(List<AddToPersonalDbCandidateItem> items, bool overwritePrice = true)
        {
            // 初始化返回结果对象
            var result = new AddToPersonalDbResult();
            if (items == null || items.Count == 0)
            {
                result.Success = false;
                result.Message = "没有可供入库的元器件条目！"; // --硬编码: 提示语--
                return result;
            }

            // 过滤出用户勾选并且包含有效型号的条目
            var validItems = items
                .Where(x => x.IsSelected && !string.IsNullOrWhiteSpace(x.Model))
                .ToList();

            result.TotalCount = validItems.Count;
            if (validItems.Count == 0)
            {
                result.Success = false;
                result.Message = "未选中任何有效的元器件条目！"; // --硬编码: 提示语--
                return result;
            }

            try
            {
                // 获取 SQLite 连接字符串
                string connStr = GetConnectionString();

                // 获取当前标准 ISO 时间戳
                string nowIso = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); // --硬编码: 时间格式--

                // 加数据库互斥锁
                lock (_dbLock)
                {
                    // 创建物理连接
                    using var conn = new SQLiteConnection(connStr);
                    // 打开连接
                    conn.Open();
                    // 开启事务保障批量原子性提交
                    using var trans = conn.BeginTransaction();

                    // 预编译插入 SQL 语句
                    string insertSql = @"
                        INSERT INTO components (
                            brand, name, model, price, remark, param1, current, poles, tripping, created_at, updated_at
                        ) VALUES (
                            @brand, @name, @model, @price, @remark, @param1, @current, @poles, @tripping, @created_at, @updated_at
                        );
                    "; // --硬编码: 插入SQL--

                    // 预编译更新 SQL 语句 (覆盖价格与补充信息)
                    string updateSql = @"
                        UPDATE components 
                        SET price = @price,
                            name = CASE WHEN @name != '' THEN @name ELSE name END,
                            remark = CASE WHEN @remark != '' THEN @remark ELSE remark END,
                            param1 = CASE WHEN @param1 != '' THEN @param1 ELSE param1 END,
                            current = CASE WHEN @current IS NOT NULL THEN @current ELSE current END,
                            poles = CASE WHEN @poles != '' THEN @poles ELSE poles END,
                            tripping = CASE WHEN @tripping != '' THEN @tripping ELSE tripping END,
                            updated_at = @updated_at 
                        WHERE id = @id;
                    "; // --硬编码: 更新SQL--

                    // 循环处理每一个有效条目
                    foreach (var item in validItems)
                    {
                        // 清洗品牌字段 (为空时赋予默认品牌“通用”)
                        string brand = string.IsNullOrWhiteSpace(item.Brand) ? "通用" : item.Brand.Trim(); // --硬编码: 兜底品牌--
                        string name = item.Name?.Trim() ?? string.Empty;
                        string model = item.Model?.Trim() ?? string.Empty;
                        double price = Convert.ToDouble(item.Price);
                        string remark = item.Remark?.Trim() ?? string.Empty;
                        string category = item.Category?.Trim() ?? string.Empty;

                        // 若已存在主键 ID (库中已存在)
                        if (item.OldId.HasValue && item.OldId.Value > 0)
                        {
                            // 若开启新价格覆盖更新
                            if (overwritePrice)
                            {
                                // 执行更新语句
                                using var updateCmd = new SQLiteCommand(updateSql, conn, trans);
                                updateCmd.Parameters.AddWithValue("@id", item.OldId.Value);
                                updateCmd.Parameters.AddWithValue("@price", price);
                                updateCmd.Parameters.AddWithValue("@name", name);
                                updateCmd.Parameters.AddWithValue("@remark", remark);
                                updateCmd.Parameters.AddWithValue("@param1", category);
                                updateCmd.Parameters.AddWithValue("@current", (object?)item.Current ?? DBNull.Value);
                                updateCmd.Parameters.AddWithValue("@poles", item.Poles ?? string.Empty);
                                updateCmd.Parameters.AddWithValue("@tripping", item.Tripping ?? string.Empty);
                                updateCmd.Parameters.AddWithValue("@updated_at", nowIso);

                                // 执行非查询指令并累加更新计数
                                int rows = updateCmd.ExecuteNonQuery();
                                if (rows > 0) result.UpdatedCount++;
                            }
                            else
                            {
                                // 未开启覆盖则跳过
                                result.SkippedCount++;
                            }
                        }
                        else
                        {
                            // 全新物料：执行插入语句
                            using var insertCmd = new SQLiteCommand(insertSql, conn, trans);
                            insertCmd.Parameters.AddWithValue("@brand", brand);
                            insertCmd.Parameters.AddWithValue("@name", name);
                            insertCmd.Parameters.AddWithValue("@model", model);
                            insertCmd.Parameters.AddWithValue("@price", price);
                            insertCmd.Parameters.AddWithValue("@remark", remark);
                            insertCmd.Parameters.AddWithValue("@param1", category);
                            insertCmd.Parameters.AddWithValue("@current", (object?)item.Current ?? DBNull.Value);
                            insertCmd.Parameters.AddWithValue("@poles", item.Poles ?? string.Empty);
                            insertCmd.Parameters.AddWithValue("@tripping", item.Tripping ?? string.Empty);
                            insertCmd.Parameters.AddWithValue("@created_at", nowIso);
                            insertCmd.Parameters.AddWithValue("@updated_at", nowIso);

                            // 执行插入指令并累加新增计数
                            int rows = insertCmd.ExecuteNonQuery();
                            if (rows > 0) result.InsertedCount++;
                        }
                    }

                    // 提交事务持久化所有修改
                    trans.Commit();
                }

                // 组装成功提示信息
                result.Success = true;
                result.Message = $"入库成功！共新增 {result.InsertedCount} 项全新物料，覆盖更新 {result.UpdatedCount} 项已有物料" +
                                 (result.SkippedCount > 0 ? $"，跳过 {result.SkippedCount} 项" : "。"); // --硬编码: 反馈信息模板--
            }
            catch (Exception ex)
            {
                // 记录批量入库异常日志
                LogHelper.WriteLog($"[PersonalDb] BatchSaveOrUpdatePersonalComponents 异常: {ex.Message}");
                result.Success = false;
                result.Message = $"入库失败: {ex.Message}"; // --硬编码: 异常提示--
            }

            // 返回最终执行统计结果
            return result;
        }

        /// <summary>
        /// 构建品牌与型号的匹配复合键
        /// </summary>
        private static string BuildCompMatchKey(string? brand, string? model)
        {
            // 清洗并规范化品牌与型号
            string b = (brand ?? string.Empty).Trim().ToLowerInvariant();
            string m = (model ?? string.Empty).Trim().ToLowerInvariant();
            // 采用 ## 分隔符拼接作为字典键 --硬编码: 键分隔符--
            return $"{b}##{m}";
        }
    }
}
