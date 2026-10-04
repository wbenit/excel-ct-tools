using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 元器件拆分与改型待处理数据上下文模型 (传输至前端展示)
    /// </summary>
    public class ComponentSplitCandidateDto
    {
        // 当前工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 当前选中的 Excel 物理行号 (1-based)
        public int RowIndex { get; set; }

        // 所属箱柜在分类表中的序号 K (对应 Cab_Det_K)
        public int CabinetK { get; set; }

        // 所属箱柜名称 (如: "1#进线柜", "AA1")
        public string CabinetName { get; set; } = string.Empty;

        // 该箱柜元器件有效起始行号 (Cab_Det_K.Row + 2)
        public int CompStartRow { get; set; }

        // 该箱柜元器件有效结束行号 (Cab_Subsum_K.Row - 1)
        public int CompEndRow { get; set; }

        // 该箱柜小计行号 (Cab_Subsum_K.Row)
        public int SubsumRow { get; set; }

        // 原序号 (A 列)
        public string Index { get; set; } = string.Empty;

        // 原元器件名称 (B 列)
        public string Name { get; set; } = string.Empty;

        // 原规格型号 (C 列)
        public string Model { get; set; } = string.Empty;

        // 原品牌厂家 (D 列)
        public string Brand { get; set; } = string.Empty;

        // 原计量单位 (E 列, 默认 "只") --硬编码: 默认单位--
        public string Unit { get; set; } = "只";

        // 原总数量 (F 列)
        public double Quantity { get; set; }

        // 原单价/面价 (G 列)
        public decimal Price { get; set; }

        // 原成本单价 (J 列)
        public decimal CostPrice { get; set; }

        // AD 列原始 CAD 句柄文本 (逗号/分号/空格分隔)
        public string RawHandles { get; set; } = string.Empty;

        // 解析后提取出的独立 CAD 句柄列表 (如: ["2A1", "2A2", "2A3", "2A4", "2A5"])
        public List<string> Handles { get; set; } = new List<string>();
    }

    /// <summary>
    /// 前端提交的拆分改型执行请求实体
    /// </summary>
    public class ComponentSplitSubmitRequest
    {
        // 目标工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 目标物理行号
        public int RowIndex { get; set; }

        // 所属箱柜序号 K
        public int CabinetK { get; set; }

        // 原行保留的数量 (如: 3)
        public double RemainQuantity { get; set; }

        // 原行保留的 CAD 句柄列表 (如: ["2A1", "2A2", "2A3"])
        public List<string> RemainHandles { get; set; } = new List<string>();

        // 拆分出的新行数量 (如: 2)
        public double SplitQuantity { get; set; }

        // 拆分出的新行 CAD 句柄列表 (如: ["2A4", "2A5"])
        public List<string> SplitHandles { get; set; } = new List<string>();

        // 新元器件名称
        public string NewName { get; set; } = string.Empty;

        // 新规格型号
        public string NewModel { get; set; } = string.Empty;

        // 新品牌厂家
        public string NewBrand { get; set; } = string.Empty;

        // 新计量单位 (默认 "只") --硬编码: 默认单位--
        public string NewUnit { get; set; } = "只";

        // 新单价/面价
        public decimal NewPrice { get; set; }

        // 新成本单价
        public decimal NewCostPrice { get; set; }

        // 是否同步更新 CAD 图元文字/属性
        public bool SyncToCad { get; set; } = true;
    }

    /// <summary>
    /// 拆分改型操作结果响应模型
    /// </summary>
    public class ComponentSplitResult
    {
        // 是否执行成功
        public bool Success { get; set; }

        // 执行结果消息或错误原因
        public string Message { get; set; } = string.Empty;

        // 物理插入的新行行号
        public int InsertedRowIndex { get; set; }
    }
}
