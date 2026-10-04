using System;
using System.Collections.Generic;
using System.Data.SQLite;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// 本地个人物料库 SQLite 数据库服务分部类：物料型号快速联想与查价
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public static partial class PersonalComponentDbService
    {
        /// <summary>
        /// 联想物料搜索结果项
        /// </summary>
        public class ComponentSuggestionItem
        {
            // 品牌
            public string Brand { get; set; } = string.Empty;
            // 名称
            public string Name { get; set; } = string.Empty;
            // 规格型号
            public string Model { get; set; } = string.Empty;
            // 单价
            public decimal Price { get; set; }
        }

        /// <summary>
        /// 根据型号或名称关键字模糊检索候选元器件，供拆分改型时智能补全
        /// </summary>
        /// <param name="keyword">搜索关键字</param>
        /// <param name="limit">最大返回数量，默认 15 条</param>
        /// <returns>匹配的物料条目列表</returns>
        public static List<ComponentSuggestionItem> QueryComponentSuggestions(string keyword, int limit = 15)
        {
            var results = new List<ComponentSuggestionItem>();
            // 关键字为空或空格时直接返回空列表
            if (string.IsNullOrWhiteSpace(keyword)) return results;

            try
            {
                // 获取数据库连接字符串并保证数据库文件与表结构已存在
                string connStr = GetConnectionString();
                string cleanKw = keyword.Trim();

                // 数据库加锁保障并发安全性
                lock (_dbLock)
                {
                    // 创建并打开数据库连接
                    using var conn = new SQLiteConnection(connStr);
                    conn.Open();

                    // 查询 SQL：优先匹配型号完全一致项，其次按模糊匹配排序输出
                    string sql = @"
                        SELECT brand, name, model, price 
                        FROM components 
                        WHERE model LIKE @kw OR name LIKE @kw 
                        ORDER BY (CASE WHEN model = @exact THEN 0 ELSE 1 END), id DESC 
                        LIMIT @limit;
                    "; // --硬编码: 快速联想SQL--

                    // 创建 SQL 参数化命令对象
                    using var cmd = new SQLiteCommand(sql, conn);
                    // 绑定模糊查询参数
                    cmd.Parameters.AddWithValue("@kw", $"%{cleanKw}%");
                    // 绑定精确比对参数
                    cmd.Parameters.AddWithValue("@exact", cleanKw);
                    // 绑定限制数量
                    cmd.Parameters.AddWithValue("@limit", limit);

                    // 执行流式读取
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        // 提取品牌
                        string b = reader.IsDBNull(0) ? string.Empty : reader.GetString(0).Trim();
                        // 提取名称
                        string n = reader.IsDBNull(1) ? string.Empty : reader.GetString(1).Trim();
                        // 提取规格型号
                        string m = reader.IsDBNull(2) ? string.Empty : reader.GetString(2).Trim();
                        // 提取单价
                        decimal p = reader.IsDBNull(3) ? 0.0m : Convert.ToDecimal(reader.GetDouble(3));

                        // 填充并记录条目
                        results.Add(new ComponentSuggestionItem
                        {
                            Brand = b,
                            Name = n,
                            Model = m,
                            Price = p
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录检索异常日志
                LogHelper.WriteLog($"[PersonalDb] QueryComponentSuggestions 异常: {ex.Message}");
            }

            return results;
        }
    }
}
