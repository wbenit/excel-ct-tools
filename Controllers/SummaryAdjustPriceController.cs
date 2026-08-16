using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 工作表分类项及箱柜台数数据模型
    /// </summary>
    public class SummaryCategoryDto
    {
        // 分类工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 该分类包含的箱柜总台数
        public int CabinetCount { get; set; } = 0;

        // 是否默认选中
        public bool IsSelected { get; set; } = true;
    }

    /// <summary>
    /// 元件合并条件配置数据模型
    /// </summary>
    public class MergeConditionsDto
    {
        // 是否按名称合并（固定必选）
        public bool ByName { get; set; } = true;

        // 是否按型号合并（固定必选）
        public bool ByModel { get; set; } = true;

        // 是否按厂家合并
        public bool ByManufacturer { get; set; } = true;

        // 是否包含无厂家项
        public bool IncludeNoManufacturer { get; set; } = false;

        // 是否按价格合并
        public bool ByPrice { get; set; } = false;

        // 是否按备注合并
        public bool ByRemark { get; set; } = false;

        // 是否按原图型号合并
        public bool ByOriginalModel { get; set; } = false;
    }

    /// <summary>
    /// 汇总表显示设置数据模型
    /// </summary>
    public class DisplaySettingsDto
    {
        // 附件表价、折扣是否分列显示
        public bool SplitAccessoryAndDiscount { get; set; } = true;

        // 是否冻结单价列
        public bool FreezeUnitPrice { get; set; } = false;
    }

    /// <summary>
    /// 汇总表排序设置数据模型
    /// </summary>
    public class SortSettingsDto
    {
        // 排序类型：mfg_name_model（厂家、名称、型号）、category_pos_mfg（类别+位置、名称等）、custom（自定义）
        public string SortType { get; set; } = "mfg_name_model";

        // 是否按类别排序
        public bool ByCategory { get; set; } = false;

        // 是否按位置、名称、厂家、型号排序
        public bool ByPosition { get; set; } = true;
    }

    /// <summary>
    /// 生成元件汇总表请求参数实体
    /// </summary>
    public class GenerateSummaryRequest
    {
        // 用户选中的要汇总的分类工作表名称列表
        public List<string> SelectedSheets { get; set; } = new List<string>();

        // 合并条件配置
        public MergeConditionsDto MergeConditions { get; set; } = new MergeConditionsDto();

        // 显示设置配置
        public DisplaySettingsDto DisplaySettings { get; set; } = new DisplaySettingsDto();

        // 排序设置配置
        public SortSettingsDto SortSettings { get; set; } = new SortSettingsDto();
    }

    /// <summary>
    /// WebAPI 风格的汇总调价控制器，提供分类列表查询与汇总表生成接口
    /// </summary>
    public class SummaryAdjustPriceController
    {
        // 全局驼峰命名 JSON 序列化配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        /// <summary>
        /// 获取当前工作簿中所有分类工作表及其箱柜数量
        /// </summary>
        public string GetCategories()
        {
            try
            {
                // 调用 ExcelServices 底层服务，读取分类表和箱柜台数（必须在 STA 线程执行）
                var categories = ExcelServices.GetCategorySheetsWithCabinetCount();
                // 封装标准化返回结果
                var response = new
                {
                    action = "onCategoriesLoaded",
                    success = true,
                    message = "分类列表加载成功",
                    data = categories
                };
                // 序列化为 JSON 字符串 (采用驼峰命名)
                return JsonSerializer.Serialize(response, JsonOptions);
            }
            catch (Exception ex)
            {
                // 记录异常日志并封装失败消息
                LogHelper.WriteLog($"GetCategories 执行异常: {ex.Message}");
                var errorResponse = new
                {
                    action = "onCategoriesLoaded",
                    success = false,
                    message = $"读取分类列表失败: {ex.Message}",
                    data = new List<SummaryCategoryDto>()
                };
                // 返回错误 JSON
                return JsonSerializer.Serialize(errorResponse, JsonOptions);
            }
        }

        /// <summary>
        /// 执行元件数据提取、内存合并聚合与生成“元件汇总表”
        /// </summary>
        public string GenerateSummary(GenerateSummaryRequest request)
        {
            try
            {
                // 验证请求参数
                if (request == null || request.SelectedSheets == null || request.SelectedSheets.Count == 0)
                {
                    return JsonSerializer.Serialize(new
                    {
                        action = "onSummaryGenerated",
                        success = false,
                        message = "未选择任何分类工作表"
                    }, JsonOptions);
                }

                // 调用 ExcelServices 业务方法生成“元件汇总表”
                bool success = ExcelServices.GenerateComponentSummarySheet(request);
                // 封装成功响应
                var response = new
                {
                    action = "onSummaryGenerated",
                    success = success,
                    message = success ? "元件汇总表生成成功" : "生成元件汇总表时遇到异常，请检查工作表结构"
                };
                // 返回序列化 JSON 文本
                return JsonSerializer.Serialize(response, JsonOptions);
            }
            catch (Exception ex)
            {
                // 封装异常响应
                LogHelper.WriteLog($"GenerateSummary 执行异常: {ex.Message}");
                var errorResponse = new
                {
                    action = "onSummaryGenerated",
                    success = false,
                    message = $"生成汇总表失败: {ex.Message}"
                };
                // 返回失败结果
                return JsonSerializer.Serialize(errorResponse, JsonOptions);
            }
        }
    }
}
