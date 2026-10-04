using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// CAD 提取元器件直接投送至 Excel 的载荷模型
    /// </summary>
    public class CadImportComponentPayload
    {
        // 操作动作指令: ImportToActiveCabinet (直投当前柜) | CreateNewCabinet (新建柜)
        public string Action { get; set; } = "ImportToActiveCabinet";

        // 写入模式: Append (追加，默认) | Overwrite (覆盖)
        public string WriteMode { get; set; } = "Append";

        // 标记本次操作是否为 CAD 端复选取消选择 (反选删除/数量减1)
        public bool IsDelete { get; set; } = false;

        // 箱柜基础信息 (供新建箱柜或校验匹配使用)
        public string CabinetName { get; set; } = string.Empty;

        // 箱柜安装方式 (如 明装、暗装)
        public string InstallMode { get; set; } = string.Empty;

        // 箱柜布置位置
        public string BuildPosition { get; set; } = string.Empty;

        // 箱柜台数 (默认 1)
        public int CabinetQuantity { get; set; } = 1;

        // 箱柜在 CAD 图纸中标记的对角线 2 个 DBPoint 句柄 (用于全屏无留白视口定位)
        public List<string> CabinetMinMaxPoints { get; set; } = new List<string>();

        // 待写入或待扣减的元器件明细项清单
        public List<CadExtractedComponentItem> Components { get; set; } = new List<CadExtractedComponentItem>();
    }

    /// <summary>
    /// 从 CAD 图纸中单次提取的元器件明细项
    /// </summary>
    public class CadExtractedComponentItem
    {
        // 元器件名称 (保持 TuFan 识别的原样，如 "断路器")
        public string Name { get; set; } = string.Empty;

        // 元器件规格型号 (如 "NXM-125S/3300 100A", "C16/1P")
        public string Specification { get; set; } = string.Empty;

        // 数量 (默认 1)
        public int Quantity { get; set; } = 1;

        // 计量单位 (成套电气元器件默认为 "只" --硬编码: 元件默认单位--)
        public string Unit { get; set; } = "只";

        // 对应的 CAD 图元句柄 (多个图元以减号或逗号连接)
        public string Handle { get; set; } = string.Empty;

        // 备注信息
        public string Remark { get; set; } = string.Empty;
    }

    /// <summary>
    /// CAD 导入操作执行结果实体
    /// </summary>
    public class CadImportResult
    {
        // 是否执行成功
        public bool Success { get; set; }

        // 业务状态码: SUCCESS, NOT_IN_CABINET, NO_ACTIVE_SHEET, ERROR
        public string Code { get; set; } = "SUCCESS";

        // 详细反馈消息文本 (用于在 CAD 命令行或日志中打印)
        public string Message { get; set; } = string.Empty;

        // 本次操作受影响的元器件行数
        public int AffectedCount { get; set; }

        // 目标箱柜名称或编号标识
        public string TargetCabinetName { get; set; } = string.Empty;

        /// <summary>
        /// 便捷构造成功结果
        /// </summary>
        public static CadImportResult Ok(string message, int affectedCount = 0, string cabName = "")
        {
            // 初始化成功状态实体
            return new CadImportResult
            {
                // 标记为成功
                Success = true,
                // 设置标准成功状态码
                Code = "SUCCESS",
                // 填充反馈说明文本
                Message = message,
                // 记录受影响行数
                AffectedCount = affectedCount,
                // 记录目标箱柜
                TargetCabinetName = cabName
            };
        }

        /// <summary>
        /// 便捷构造失败或拦截结果
        /// </summary>
        public static CadImportResult Fail(string code, string message)
        {
            // 初始化失败状态实体
            return new CadImportResult
            {
                // 标记为失败
                Success = false,
                // 设置具体业务错误码
                Code = code,
                // 填充失败或警告原因说明
                Message = message
            };
        }
    }
}
