using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 母排一键改价配置传输对象
    /// </summary>
    public class BusbarBatchPriceConfigDto
    {
        // 基准铜价 (单位: 元/kg)
        public double CuBasePrice { get; set; } = 72.0;

        // 铜排加工费 (单位: 元/kg)
        public double CuProcessPrice { get; set; } = 8.5;

        // 基准铝价 (单位: 元/kg)
        public double AlBasePrice { get; set; } = 20.0;

        // 铝排加工费 (单位: 元/kg)
        public double AlProcessPrice { get; set; } = 4.0;

        // 综合铜排单价 (只读计算属性: 基准铜价 + 铜加工费)
        public double CuTotalUnitPrice => Math.Round(CuBasePrice + CuProcessPrice, 2);

        // 综合铝排单价 (只读计算属性: 基准铝价 + 铝加工费)
        public double AlTotalUnitPrice => Math.Round(AlBasePrice + AlProcessPrice, 2);

        // 改价生效范围模式 ("current": 仅当前分类工作表; "all": 全项目所有分类表)
        public string ScopeMode { get; set; } = "current";

        // 目标参与改价的工作表名称集合
        public List<string> TargetSheets { get; set; } = new List<string>();

        // 是否对“米/m”单位的母排自动依据规格截面折算每米理论单价
        public bool AutoConvertMeter { get; set; } = true;
    }

    /// <summary>
    /// 单行母排物料数据传输模型
    /// </summary>
    public class BusbarMaterialRowDto
    {
        // 唯一定位标识键
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        // 所属 Excel 工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 所在 Excel 工作表的物理行号 (1-indexed)
        public int RowIndex { get; set; }

        // 所属箱柜序号 (k)
        public int CabinetIndex { get; set; }

        // 所属箱柜名称 / 柜号 (如: 1#进线柜 GGD)
        public string CabinetName { get; set; } = string.Empty;

        // 材料名称 (B 列文本，如: 矩形铜母线、TMY铜排)
        public string Name { get; set; } = string.Empty;

        // 规格型号 (C 列文本，如: TMY-3*(80*8)+1*(80*8))
        public string Model { get; set; } = string.Empty;

        // 材料材质归属 ("copper": 铜排; "aluminum": 铝排)
        public string MaterialType { get; set; } = "copper";

        // 计量单位 (E 列文本，如: kg、公斤、米、m)
        public string Unit { get; set; } = "kg";

        // 工程用量数量 (F 列数值)
        public double Qty { get; set; }

        // 原基准单价 (M 列当前数值)
        public double OldUnitPrice { get; set; }

        // 原合价 (原单价 * 数量)
        public double OldTotalPrice { get; set; }

        // 测算出的新基准单价
        public double NewUnitPrice { get; set; }

        // 测算出的新合价 (新单价 * 数量)
        public double NewTotalPrice { get; set; }

        // 调价前后差额 (NewTotalPrice - OldTotalPrice)
        public double DiffAmount { get; set; }

        // 是否为按米计价模式
        public bool IsMeterMode { get; set; }

        // 理论每米单重 (单位: kg/m)
        public double TheoreticalKgPerMeter { get; set; }

        // 前端勾选状态 (默认全部参与改价)
        public bool Selected { get; set; } = true;

        // 备注信息 (I 列说明)
        public string Remark { get; set; } = string.Empty;
    }

    /// <summary>
    /// 母排全项目扫描统计与响应数据包
    /// </summary>
    public class BusbarBatchPriceScanResultDto
    {
        // 扫描执行是否成功
        public bool Success { get; set; } = true;

        // 状态消息提示
        public string Message { get; set; } = string.Empty;

        // 当前工程所有有效分类工作表列表
        public List<string> AvailableSheets { get; set; } = new List<string>();

        // 当前活动工作表名称
        public string CurrentSheetName { get; set; } = string.Empty;

        // 当前应用的改价基础配置
        public BusbarBatchPriceConfigDto Config { get; set; } = new BusbarBatchPriceConfigDto();

        // 扫描出的母排物料明细集合
        public List<BusbarMaterialRowDto> Rows { get; set; } = new List<BusbarMaterialRowDto>();

        // 涉及箱柜总数
        public int TotalCabinetCount { get; set; }

        // 母排物料条目总行数
        public int TotalRowCount { get; set; }

        // 铜排总重量估算 (kg)
        public double TotalCopperWeight { get; set; }

        // 铝排总重量估算 (kg)
        public double TotalAluminumWeight { get; set; }

        // 改价前总金额 (元)
        public double OldTotalAmount { get; set; }

        // 改价后总金额 (元)
        public double NewTotalAmount { get; set; }

        // 调价差额总和 (元)
        public double TotalDiffAmount { get; set; }
    }

    /// <summary>
    /// 前端向后端提交的一键改价执行载荷
    /// </summary>
    public class BusbarApplyPricePayload
    {
        // 改价全局配置
        public BusbarBatchPriceConfigDto Config { get; set; } = new BusbarBatchPriceConfigDto();

        // 用户确认应用的母排条目列表
        public List<BusbarMaterialRowDto> Items { get; set; } = new List<BusbarMaterialRowDto>();
    }
}

