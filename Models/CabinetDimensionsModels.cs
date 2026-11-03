using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 单台箱柜【尺寸安装方式】明细数据项实体
    /// </summary>
    public class CabinetDimensionsItem
    {
        // 箱柜在当前导出任务中的全局显示序号 (1, 2, 3...)
        public int Index { get; set; }

        // 箱柜名称 (默认“配电箱”或提取自汇总行)
        public string Name { get; set; } = "配电箱";

        // 柜号 (如 AP1、AL2，对应模板中的[柜号]，展示在“型号”列)
        public string CabinetCode { get; set; } = string.Empty;

        // 箱柜数量 (台数)
        public decimal Quantity { get; set; } = 1;

        // 计量单位 (固定为“台”) --硬编码: 默认计量单位--
        public string Unit { get; set; } = "台";

        // 外形尺寸: 宽高深 (如 600*800*200，对应模板[箱柜尺寸])
        public string Dimensions { get; set; } = string.Empty;

        // 安装方式 (对应模板[箱柜型号]，用户明确指令直接取汇总行 D 列)
        public string InstallMode { get; set; } = string.Empty;

        // 防护等级 (用户明确指令统一默认填 IP30) --硬编码: 默认防护等级--
        public string ProtectionLevel { get; set; } = "IP30";

        // 备注信息 (提取自汇总行 I 列)
        public string Remark { get; set; } = string.Empty;

        // 箱柜所属分类明细工作表名称 (供超链接构建)
        public string SheetName { get; set; } = string.Empty;

        // 箱柜在分类表中明细开始行物理行号 (Cab_Det_k)
        public int DetailStartRow { get; set; }

        // 箱柜在分类表中明细结束行物理行号 (Cab_Tolsum_k)
        public int DetailEndRow { get; set; }
    }

    /// <summary>
    /// 尺寸安装方式分类工作表选择项 DTO
    /// </summary>
    public class CabinetDimensionsCategoryDto
    {
        // 分类工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 该分类表下包含的有效箱柜数量 (> 0)
        public int CabinetCount { get; set; }

        // 前端是否默认勾选
        public bool IsSelected { get; set; } = true;
    }

    /// <summary>
    /// 尺寸安装方式初始化数据 DTO 模型
    /// </summary>
    public class CabinetDimensionsInitDataDto
    {
        // 活动工作簿中所有可选的分类工作表列表
        public List<CabinetDimensionsCategoryDto> Categories { get; set; } = new List<CabinetDimensionsCategoryDto>();

        // 项目名称 (提取自项目信息 B5 或工作簿文件名)
        public string ProjectName { get; set; } = string.Empty;

        // 客户单位名称 (提取自项目信息 B14)
        public string CustomerName { get; set; } = string.Empty;

        // 本企业供货单位名称 (提取自项目信息 B22 或 B1)
        public string CompanyName { get; set; } = string.Empty;

        // 当前工作簿是否已存在【尺寸安装方式】工作表
        public bool HasExistingDimensionsSheet { get; set; } = false;
    }

    /// <summary>
    /// 尺寸安装方式导出请求参数模型
    /// </summary>
    public class CabinetDimensionsExportRequest
    {
        // 用户选中的需要导出尺寸安装方式的分类工作表名称集合
        public List<string> SelectedSheets { get; set; } = new List<string>();
    }

    /// <summary>
    /// 尺寸安装方式导出执行结果模型
    /// </summary>
    public class CabinetDimensionsExportResult
    {
        // 导出操作是否成功完成
        public bool Success { get; set; }

        // 导出提示或异常信息文本
        public string Message { get; set; } = string.Empty;

        // 成功导出的箱柜总数量
        public int CabinetCount { get; set; }
    }
}
