using System;
using ExcelDna.Integration;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：箱柜表头同步与编辑
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 将修改后的箱柜 Header 属性同步写回 Excel 工作表中
        /// </summary>
        public static bool WriteCabinetHeaderToSheet(dynamic sheet, Models.CabinetObject cabinet)
        {
            if (sheet == null || cabinet == null || cabinet.DetAnchorRow <= 0) return false;

            try
            {
                int headerRow = cabinet.DetAnchorRow;

                // 写回 B 列柜号
                sheet.Cells[headerRow, 2].Value = cabinet.Header.CabinetNo;

                // 写回 D 列型号
                sheet.Cells[headerRow, 4].Value = cabinet.Header.Model;

                // 写回 F 列名称
                sheet.Cells[headerRow, 6].Value = cabinet.Header.Name;

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"同步写回箱柜表头失败: {ex.Message}");
                return false;
            }
        }
    }
}
