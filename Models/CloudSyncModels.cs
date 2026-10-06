using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 云端项目概要模型（对应后端 Project 实体精简字段）
    /// 用于前端项目绑定选择下拉与匹配展示
    /// </summary>
    public class ProjectSummaryDto
    {
        // 云端项目唯一主键 ID (ProjectId)
        [JsonPropertyName("id")]
        public int Id { get; set; }

        // 项目名称（工程项目全称）
        [JsonPropertyName("projectName")]
        public string ProjectName { get; set; } = string.Empty;

        // 项目详细描述或工程概况
        [JsonPropertyName("projectDesc")]
        public string? ProjectDesc { get; set; }

        // 工程类型（如：高压成套、低压成套、箱变等）
        [JsonPropertyName("projectType")]
        public string? ProjectType { get; set; }

        // 项目创建时间文本
        [JsonPropertyName("createTime")]
        public DateTime? CreateTime { get; set; }
    }

    /// <summary>
    /// 云端箱柜推送导入项模型（对应云端 EBoxAddDto）
    /// </summary>
    public class CloudEBoxItemDto
    {
        // 关键序号：按“分类序号-汇总序号”格式组织，例如 2-1
        [JsonPropertyName("order")]
        public string Order { get; set; } = string.Empty;

        // 图纸代号或柜号（对应云端 SystemDrawName）
        [JsonPropertyName("systemDrawName")]
        public string SystemDrawName { get; set; } = string.Empty;

        // 柜体台数/数量（默认 1）
        [JsonPropertyName("count")]
        public int Count { get; set; } = 1;

        // 箱柜所属的项目 ID
        [JsonPropertyName("projectId")]
        public int ProjectId { get; set; }

        // 箱柜安装位置
        [JsonPropertyName("eboxPositon")]
        public string? EboxPositon { get; set; }

        // 箱柜型号规格文本（用于前端展示或自定义属性）
        [JsonPropertyName("model")]
        public string? Model { get; set; }

        // 所属分类名称（如：低压出线柜）
        [JsonPropertyName("categoryName")]
        public string? CategoryName { get; set; }

        // 所属分类序号（如：2）
        [JsonPropertyName("categoryIndex")]
        public int CategoryIndex { get; set; }

        // 分类明细顶部的箱柜汇总序号（如：1）
        [JsonPropertyName("cabinetIndex")]
        public int CabinetIndex { get; set; }

        // 自定义扩展属性 JSON 字典字符串
        [JsonPropertyName("customFields")]
        public string? CustomFields { get; set; }
    }

    /// <summary>
    /// 云端批量导入请求数据传输包（对应后端 ExcelImportDto）
    /// </summary>
    public class CloudImportRequestDto
    {
        // 待同步入库的箱柜列表集合
        [JsonPropertyName("eBoxs")]
        public List<CloudEBoxItemDto> EBoxs { get; set; } = new List<CloudEBoxItemDto>();

        // 二次图列表（插件端暂留空）
        [JsonPropertyName("eTwos")]
        public List<object> ETwos { get; set; } = new List<object>();

        // 核心指令：是否允许覆盖云端已有相同 Order 序号的箱柜
        [JsonPropertyName("isCover")]
        public bool IsCover { get; set; }

        // 生产批次编号（必须大于等于 1，默认 1）
        [JsonPropertyName("produceOrder")]
        public int ProduceOrder { get; set; } = 1;
    }

    /// <summary>
    /// 云端导入执行结果响应数据包
    /// </summary>
    public class CloudImportResultDto
    {
        // 操作是否成功标记
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        // 成功导入的箱柜总台套数
        [JsonPropertyName("eBoxCount")]
        public int EBoxCount { get; set; }

        // 新增的二次图条数
        [JsonPropertyName("eTwoCount")]
        public int ETwoCount { get; set; }

        // 后端返回的提示文本
        [JsonPropertyName("message")]
        public string Message { get; set; } = string.Empty;
    }
}
