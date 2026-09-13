using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelAddInDemo.Controllers
{
    // 工作表分类项及箱柜台数数据模型
    public class DistributionCategoryDto
    {
        // 分类工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 该分类包含的箱柜总台数
        public int CabinetCount { get; set; } = 0;

        // 是否默认选中
        public bool IsSelected { get; set; } = true;
    }

    // 元件合并条件配置数据模型
    public class DistributionMergeConditionsDto
    {
        // 是否按名称合并（固定必选）
        public bool ByName { get; set; } = true;

        // 是否按型号合并（固定必选）
        public bool ByModel { get; set; } = true;

        // 是否按厂家合并
        public bool ByManufacturer { get; set; } = true;

        // 是否合并无厂家项
        public bool IncludeNoManufacturer { get; set; } = false;

        // 是否按价格合并
        public bool ByPrice { get; set; } = true;

        // 是否按备注合并
        public bool ByRemark { get; set; } = false;

        // 是否按方案名称合并 (AC列)
        public bool ByTemplateName { get; set; } = false;
    }

    // 汇总表排序设置数据模型
    public class DistributionSortSettingsDto
    {
        // 排序规则类型: type_location_name（类别+位置+名称型号）、mfg_name_model（厂家名称型号）、custom（自定义）
        public string SortType { get; set; } = "type_location_name";

        // 是否按类别排序
        public bool ByCategory { get; set; } = false;

        // 是否按位置排序
        public bool ByLocation { get; set; } = true;

        // 是否按上次顺序汇总
        public bool ByLastOrder { get; set; } = false;
    }

    // 生成材料分布表请求参数实体
    public class GenerateDistributionRequest
    {
        // 用户选中的要汇总的分类工作表名称列表
        public List<string> SelectedSheets { get; set; } = new List<string>();

        // 合并条件配置
        public DistributionMergeConditionsDto MergeConditions { get; set; } = new DistributionMergeConditionsDto();

        // 排序设置配置
        public DistributionSortSettingsDto SortSettings { get; set; } = new DistributionSortSettingsDto();
    }

    // 分布调价反向同步更新选项
    public class DistributionUpdateOptions
    {
        // 更新时是否调整元件排序（按当前分布表中的顺序重排各箱柜内元器件）
        public bool UpdateBomOrder { get; set; } = true;

        // 更新时是否不合并相同元件（保持柜内原有相同元件不合并）
        public bool NotMergeSameBom { get; set; } = false;

        // 是否记住当前元件顺序供下次复用
        public bool SaveCurrentOrder { get; set; } = false;
    }

    // 分布调价更新执行结果实体
    public class DistributionUpdateResult
    {
        // 执行是否成功
        public bool Success { get; set; } = false;

        // 返回消息提示
        public string Message { get; set; } = string.Empty;

        // 更新的分类工作表数量
        public int UpdatedSheetCount { get; set; } = 0;

        // 更新的箱柜总台数
        public int UpdatedCabinetCount { get; set; } = 0;

        // 更新的元器件总项数
        public int UpdatedComponentCount { get; set; } = 0;
    }

    /// <summary>
    /// WebAPI 风格的分布调价控制器，提供分类列表查询、分布表生成与一键反向更新接口
    /// </summary>
    public class DistributedAdjustPriceController
    {
        // 全局驼峰命名 JSON 序列化配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            // 属性名转为小驼峰式命名规则
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            // 启用反序列化属性名大小写忽略
            PropertyNameCaseInsensitive = true,
            // 输出格式化缩进
            WriteIndented = true
        };

        /// <summary>
        /// 获取当前工作簿中所有可选的分类工作表及其箱柜台数列表 (JSON 格式)
        /// </summary>
        public string GetCategorySheetsJson()
        {
            try
            {
                // 调用公共服务层方法读取各分类表及其箱柜统计
                var categories = ExcelServices.GetCategorySheetsForDistribution();

                // 组织标准成功响应报文
                var response = new
                {
                    // 状态标记为成功
                    success = true,
                    // 提示消息
                    message = "获取分类列表成功",
                    // 返回分类工作表数组
                    data = categories
                };

                // 序列化为 JSON 字符串返回
                return JsonSerializer.Serialize(response, JsonOptions);
            }
            catch (Exception ex)
            {
                // 组织异常失败响应报文
                var errorResponse = new
                {
                    // 状态标记为失败
                    success = false,
                    // 捕获并返回异常详细信息
                    message = $"获取分类列表失败: {ex.Message}",
                    // 返回空数组兜底
                    data = new List<DistributionCategoryDto>()
                };

                // 序列化并返回错误 JSON
                return JsonSerializer.Serialize(errorResponse, JsonOptions);
            }
        }

        /// <summary>
        /// 根据前端提交的配置参数，生成【材料分布表】(二维矩阵交叉表)
        /// </summary>
        public string GenerateDistributionSheetJson(string requestJson)
        {
            try
            {
                // 校验传入的请求 JSON 是否合法
                if (string.IsNullOrWhiteSpace(requestJson))
                {
                    // 抛出参数空异常
                    throw new ArgumentException("请求参数不能为空");
                }

                // 反序列化前端传来的生成分布表请求参数实体
                var request = JsonSerializer.Deserialize<GenerateDistributionRequest>(requestJson, JsonOptions);

                // 校验反序列化后的请求对象
                if (request == null || request.SelectedSheets == null || request.SelectedSheets.Count == 0)
                {
                    // 抛出未选择分类工作表异常
                    throw new ArgumentException("请至少选择一个要生成分布表的分类工作表");
                }

                // 调用 ExcelServices 业务层核心方法生成【元件汇总分布表】
                bool isSuccess = ExcelServices.GenerateComponentDistributionSheet(request);

                // 组织返回响应对象
                var response = new
                {
                    // 设置成功状态
                    success = isSuccess,
                    // 设置提示消息
                    message = isSuccess ? "【元件汇总分布表】生成成功！" : "生成【元件汇总分布表】失败，请检查数据格式",
                    // 携带生成的工作表名称 --硬编码--
                    data = new { targetSheetName = "元件汇总分布表" }
                };

                // 返回 JSON 序列化字符串
                return JsonSerializer.Serialize(response, JsonOptions);
            }
            catch (Exception ex)
            {
                // 组织错误异常响应
                var errorResponse = new
                {
                    // 标记失败
                    success = false,
                    // 记录异常描述信息
                    message = $"生成元件汇总分布表失败: {ex.Message}",
                    // 空数据返回
                    data = (object?)null
                };

                // 序列化并输出错误 JSON
                return JsonSerializer.Serialize(errorResponse, JsonOptions);
            }
        }

        /// <summary>
        /// 从当前【元件汇总分布表】中反向同步更新调价结果到所有分类工作表的各个箱柜明细中
        /// </summary>
        public string UpdateFromDistributionSheetJson(string optionsJson)
        {
            try
            {
                // 默认初始化更新选项配置
                var options = new DistributionUpdateOptions();

                // 若前端传入了具体选项 JSON，则进行反序列化
                if (!string.IsNullOrWhiteSpace(optionsJson))
                {
                    // 反序列化更新选项
                    var parsed = JsonSerializer.Deserialize<DistributionUpdateOptions>(optionsJson, JsonOptions);
                    // 覆盖有效选项
                    if (parsed != null) options = parsed;
                }

                // 调用核心服务层执行批量反向同步调价逻辑
                var result = ExcelServices.UpdateFromComponentDistributionSheet(options);

                // 组织响应报文实体
                var response = new
                {
                    // 操作成功状态
                    success = result.Success,
                    // 执行提示信息
                    message = result.Message,
                    // 返回详细的更新统计指标
                    data = new
                    {
                        // 更新的分类表数
                        updatedSheetCount = result.UpdatedSheetCount,
                        // 更新的箱柜台数
                        updatedCabinetCount = result.UpdatedCabinetCount,
                        // 更新的元件项数
                        updatedComponentCount = result.UpdatedComponentCount
                    }
                };

                // 序列化并返回 JSON
                return JsonSerializer.Serialize(response, JsonOptions);
            }
            catch (Exception ex)
            {
                // 组织执行失败的响应
                var errorResponse = new
                {
                    // 失败标记
                    success = false,
                    // 异常提示文本
                    message = $"反向同步调价数据失败: {ex.Message}",
                    // 空数据返回
                    data = (object?)null
                };

                // 序列化并返回错误 JSON
                return JsonSerializer.Serialize(errorResponse, JsonOptions);
            }
        }

        /// <summary>
        /// 检查当前工作簿中是否已存在【元件汇总分布表】
        /// </summary>
        public string CheckDistributionSheetExistsJson()
        {
            try
            {
                // 调用服务层检测分布表是否存在
                bool exists = ExcelServices.CheckDistributionSheetExists();

                // 构建响应报文
                var response = new
                {
                    // 请求成功标记
                    success = true,
                    // 携带存在性布尔值
                    data = new { exists = exists }
                };

                // 序列化返回 JSON
                return JsonSerializer.Serialize(response, JsonOptions);
            }
            catch (Exception ex)
            {
                // 异常响应构造
                var errorResponse = new
                {
                    // 标记失败
                    success = false,
                    // 错误信息
                    message = $"检查元件汇总分布表存在状态失败: {ex.Message}",
                    // 默认不存在
                    data = new { exists = false }
                };

                // 序列化返回 JSON
                return JsonSerializer.Serialize(errorResponse, JsonOptions);
            }
        }
    }
}
