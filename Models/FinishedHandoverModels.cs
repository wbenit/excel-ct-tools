using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 成品交接单初始化数据传输模型 (用于前端向导展示分类表列表与工程信息)
    /// </summary>
    public class FinishedHandoverInitDataDto
    {
        // 工程项目名称 (优先从【项目信息】提取，默认取工作簿文件名)
        public string ProjectName { get; set; } = string.Empty;

        // 客户单位名称 (优先从【项目信息】提取)
        public string CustomerName { get; set; } = string.Empty;

        // 本企业单位名称 (优先从【项目信息】提取)
        public string CompanyName { get; set; } = string.Empty;

        // 当前工作簿是否已存在【成品交接单】工作表
        public bool HasExistingFinishedHandover { get; set; } = false;

        // 有效分类工作表列表
        public List<FinishedHandoverCategoryDto> Categories { get; set; } = new List<FinishedHandoverCategoryDto>();
    }

    /// <summary>
    /// 分类工作表基本信息实体
    /// </summary>
    public class FinishedHandoverCategoryDto
    {
        // 分类工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 该分类表包含的有效箱柜数量
        public int CabinetCount { get; set; } = 0;

        // 向导勾选状态 (默认全选)
        public bool IsSelected { get; set; } = true;
    }

    /// <summary>
    /// 成品交接单明细行实体 (母版分列式)
    /// </summary>
    public class FinishedHandoverItem
    {
        // A 列：序号 (1, 2, 3...)
        public int Index { get; set; } = 0;

        // B 列：名称 (固定为“配电箱”或箱柜名称)
        public string Name { get; set; } = "配电箱"; // --硬编码: 默认箱柜名称--

        // C 列：型 号 (工程柜号，如 2-ALZ11、1APL1 3)
        public string CabinetCode { get; set; } = string.Empty;

        // D 列：数量 (箱柜台数)
        public decimal Quantity { get; set; } = 1;

        // E 列：单位 (固定为“台”)
        public string Unit { get; set; } = "台"; // --硬编码: 默认计量单位--

        // F 列：箱柜型号 (取汇总行 N 列三箱型号，若为空则由管道决策流判定兜底)
        public string CabinetModel { get; set; } = string.Empty;

        // G 列：电流 (抓取明细表中主进线开关整定电流，如 100A, 63A)
        public string Current { get; set; } = string.Empty;

        // H 列：防护等级 (默认标准 IP30，若有特殊标注则提取)
        public string ProtectionLevel { get; set; } = "IP30"; // --硬编码: 默认防护等级 IP30--

        // I 列：备注 (工程楼栋编号/图号，或分类表名称)
        public string Remark { get; set; } = string.Empty;

        // 来源工作表名称 (用于构造回跳超链接)
        public string SheetName { get; set; } = string.Empty;

        // 明细起始行号 (Cab_Det 行号)
        public int DetailStartRow { get; set; } = 0;

        // 明细终止行号 (Cab_Tolsum 行号)
        public int DetailEndRow { get; set; } = 0;
    }

    /// <summary>
    /// 成品交接单导出执行结果模型
    /// </summary>
    public class FinishedHandoverExportResult
    {
        // 导出是否成功
        public bool Success { get; set; } = false;

        // 提示或错误文本
        public string Message { get; set; } = string.Empty;

        // 生成的箱柜总台数
        public int CabinetCount { get; set; } = 0;

        // 最终生成的目标工作表名称
        public string SheetName { get; set; } = "成品交接单"; // --硬编码: 工作表名称--
    }
}
