using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 新建分类请求数据传输模型
    /// </summary>
    public class CreateCategoryRequest
    {
        // 分类工作表名称 (如: "分类2", "高压配电房")
        public string CategoryName { get; set; } = string.Empty;

        // 该分类下首台初始箱柜名称 (如: "箱柜1")
        public string InitialCabinetName { get; set; } = "箱柜1";

        // 选中的计费调费公式组 ID
        public string FormulaGroupId { get; set; } = string.Empty;
    }

    /// <summary>
    /// 分类初始化建议信息数据模型
    /// </summary>
    public class CategorySuggestInfo
    {
        // 系统推荐的下一个有效分类名称 (如 "分类2")
        public string SuggestedName { get; set; } = "分类1";

        // 当前工作簿中已存在的分类工作表名称列表 (用于前端查重校验)
        public List<string> ExistingCategories { get; set; } = new List<string>();

        // 可供选择的调费公式组选项集合
        public List<CategoryFormulaGroupOption> FormulaGroups { get; set; } = new List<CategoryFormulaGroupOption>();
    }

    /// <summary>
    /// 调费公式组下拉选项数据模型
    /// </summary>
    public class CategoryFormulaGroupOption
    {
        // 公式组唯一标识 ID
        public string Id { get; set; } = string.Empty;

        // 公式组显示名称 (如 "高压柜标准公式组")
        public string Name { get; set; } = string.Empty;

        // 是否为默认选中的公式组
        public bool IsDefault { get; set; } = false;
    }

    /// <summary>
    /// 新建分类操作响应模型
    /// </summary>
    public class CategoryOperationResult
    {
        // 操作执行成功标识
        public bool Success { get; set; }

        // 提示反馈消息文本
        public string Message { get; set; } = string.Empty;

        // 成功创建的分类工作表名称
        public string CategoryName { get; set; } = string.Empty;
    }
}
