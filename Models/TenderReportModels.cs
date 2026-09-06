using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 投标报表工程基础信息模型
    /// </summary>
    public class TenderReportProjectInfo
    {
        // 工程项目名称 (如: "某某变电站成套工程")
        public string ProjectName { get; set; } = string.Empty;

        // 报价文件编号 / 报价单号 (如: "BJ-2026-001")
        public string QuotationNumber { get; set; } = string.Empty;

        // 报价编制日期 (如: "2026-09-05")
        public string QuotationDate { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");

        // 采购单位 / 客户甲方单位名称
        public string CustomerName { get; set; } = string.Empty;

        // 客户联系人姓名
        public string CustomerContact { get; set; } = string.Empty;

        // 客户联系电话 / 传真
        public string CustomerPhone { get; set; } = string.Empty;

        // 供货单位 / 本单位名称
        public string FactoryName { get; set; } = string.Empty;

        // 供货单位业务联系人
        public string FactoryContact { get; set; } = string.Empty;

        // 供货单位联系电话
        public string FactoryPhone { get; set; } = string.Empty;

        // 报价编制人 / 业务经理
        public string Bidder { get; set; } = string.Empty;

        // 报价备注说明文本
        public string ProjectRemark { get; set; } = string.Empty;

        // 币种与金额单位说明 (默认: "人民币元") --硬编码--
        public string AmountUnit { get; set; } = "人民币元";
    }

    /// <summary>
    /// 投标报表明细元器件项目数据模型
    /// </summary>
    public class TenderReportComponentItem
    {
        // 元件行序号 (如: 1, 2, 3...)
        public int Index { get; set; }

        // 元件通用名称 (如: "真空断路器", "微机保护测控装置")
        public string Name { get; set; } = string.Empty;

        // 规格型号 (如: "VS1-12/1250-31.5")
        public string Model { get; set; } = string.Empty;

        // 计量单位 (如: "台", "只", "套")
        public string Unit { get; set; } = "台";

        // 单台箱柜使用数量
        public decimal Quantity { get; set; } = 0;

        // 销售报出单价
        public decimal UnitPrice { get; set; } = 0;

        // 报出总价合价 (Quantity * UnitPrice)
        public decimal TotalPrice { get; set; } = 0;

        // 生产制造厂家 / 品牌 (如: "ABB", "施耐德")
        public string Manufacturer { get; set; } = string.Empty;

        // 元器件备注说明
        public string Remark { get; set; } = string.Empty;
    }

    /// <summary>
    /// 投标报表箱柜设备项模型 (对应一台或一组同型号箱柜)
    /// </summary>
    public class TenderReportCabinetItem
    {
        // 箱柜序号 (如: 1, 2, 3...)
        public int Index { get; set; }

        // 箱柜柜号编号 (如: "1AA1", "2AA3")
        public string CabinetNo { get; set; } = string.Empty;

        // 箱柜标准名称 (如: "进线柜", "变压器出线柜")
        public string CabinetName { get; set; } = string.Empty;

        // 箱柜型号规格 (如: "KYN28A-12")
        public string CabinetModel { get; set; } = string.Empty;

        // 计量单位 (默认: "台") --硬编码--
        public string Unit { get; set; } = "台";

        // 箱柜台数数量
        public int Quantity { get; set; } = 1;

        // 单台箱柜报出单价 (含元件小计与成套计费)
        public decimal UnitPrice { get; set; } = 0;

        // 箱柜报出合价 (Quantity * UnitPrice)
        public decimal TotalPrice { get; set; } = 0;

        // 箱柜外形尺寸 (如: "800*1500*2200")
        public string Size { get; set; } = string.Empty;

        // 关联 CAD 设计图纸图号
        public string CadDrawingNo { get; set; } = string.Empty;

        // 箱柜备注说明
        public string Remark { get; set; } = string.Empty;

        // 箱柜内部包含的所有元器件清单列表
        public List<TenderReportComponentItem> Components { get; set; } = new List<TenderReportComponentItem>();
    }

    /// <summary>
    /// 投标报表分类工作表分组模型
    /// </summary>
    public class TenderReportCategoryGroup
    {
        // 所属分类工作表 Sheet 名称 (如: "10kV高压柜", "0.4kV低压柜")
        public string CategoryName { get; set; } = string.Empty;

        // 该分类下的箱柜总台数
        public int TotalCabinetCount { get; set; } = 0;

        // 该分类下的报出销售总金额
        public decimal TotalAmount { get; set; } = 0;

        // 是否被勾选参与报表导出
        public bool IsSelected { get; set; } = true;

        // 该分类包含的所有箱柜列表
        public List<TenderReportCabinetItem> Cabinets { get; set; } = new List<TenderReportCabinetItem>();
    }

    /// <summary>
    /// 前端提交的常规报表导出参数配置模型
    /// </summary>
    public class TenderReportExportConfig
    {
        // 覆盖更新的工程基本信息
        public TenderReportProjectInfo ProjectInfo { get; set; } = new TenderReportProjectInfo();

        // 选中的分类工作表名称列表
        public List<string> SelectedCategories { get; set; } = new List<string>();

        // 是否包含导出《封面》工作表
        public bool IncludeCover { get; set; } = true;

        // 是否包含导出《屏柜汇总表》工作表
        public bool IncludeSummary { get; set; } = true;

        // 是否包含导出《屏柜分项表》工作表
        public bool IncludeDetail { get; set; } = true;

        // 导出完成后是否自动在 Excel 中激活展示新工作簿
        public bool AutoOpen { get; set; } = true;

        // 用户指定的目标另存为文件路径 (若为空则自动保存至工程同级目录)
        public string TargetFilePath { get; set; } = string.Empty;
    }

    /// <summary>
    /// 报表导出执行结果响应模型
    /// </summary>
    public class TenderReportExportResult
    {
        // 导出是否成功标识
        public bool Success { get; set; } = false;

        // 结果状态提示文本
        public string Message { get; set; } = string.Empty;

        // 导出的实际目标 Excel 文件全路径
        public string OutputFilePath { get; set; } = string.Empty;

        // 导出的总分类数量
        public int ExportedCategoryCount { get; set; } = 0;

        // 导出的总箱柜台数
        public int ExportedCabinetCount { get; set; } = 0;

        // 导出的报表总金额
        public decimal ExportedTotalAmount { get; set; } = 0;
    }
}
