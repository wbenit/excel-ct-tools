using System;
using System.IO;
using ExcelAddInDemo.Forms;
using ExcelDna.Integration;

namespace ExcelAddInDemo
{
    // ExcelServices 公共服务分部类：元器件图纸参数匹配核心服务
    public static partial class ExcelServices
    {
        /// <summary>
        /// 调起并激活“元器件图纸参数匹配 (200x800)”伴随式侧边浮窗
        /// </summary>
        public static void ShowComponentParamMatchDialog()
        {
            try
            {
                // 通过 WinForms 单例模式调起并展示窗口
                ComponentParamMatchForm.ShowForm();
            }
            catch (Exception ex)
            {
                // 记录打开窗口时的异常日志
                LogHelper.WriteLog($"[ExcelServices] 打开元器件图纸参数匹配窗口异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 将指定目录名称与 DWG 图纸名称即时写入当前活动工作表的当前行：
        /// 规则：
        /// 1. 若当前在【元件汇总表】：写入 X 列 (图纸名) 与 Y 列 (目录名)
        /// 2. 若当前在普通【分类表】：写入 AA 列 (图纸名) 与 AB 列 (目录名)
        /// </summary>
        /// <param name="dirName">当前目录名称</param>
        /// <param name="dwgName">DWG 图纸文件名称</param>
        /// <param name="autoNextRow">是否在写入后自动下移一行</param>
        /// <param name="colDir">自定义目录写入列标 (默认 null，自动按工作表类型路由)</param>
        /// <param name="colDwg">自定义图纸写入列标 (默认 null，自动按工作表类型路由)</param>
        /// <param name="removeExt">是否去除 .dwg 扩展名</param>
        /// <returns>写入结果状态、反馈消息、最新激活的行号、生效图纸列、生效目录列及工作表名</returns>
        public static (bool Success, string Message, int NextRow, string ColDwg, string ColDir, string SheetName) BindDwgParamToActiveRow(
            string dirName,
            string dwgName,
            bool autoNextRow = true,
            string? colDir = null,
            string? colDwg = null,
            bool removeExt = true)
        {
            try
            {
                // 获取 Excel 宿主全局 Application 对象句柄
                dynamic? app = ExcelDnaUtil.Application;
                // 校验 Excel 宿主对象有效性
                if (app == null)
                {
                    return (false, "Excel Application 未就绪", 0, "AA", "AB", "");
                }

                // 获取当前活动的工作表对象
                dynamic? sheet = app.ActiveSheet;
                // 校验工作表有效性
                if (sheet == null)
                {
                    return (false, "当前无有效活动工作表", 0, "AA", "AB", "");
                }

                // 获取当前活动工作表名称并判断是否为“元件汇总表”
                string sheetName = sheet.Name != null ? Convert.ToString(sheet.Name).Trim() : string.Empty;
                bool isSummarySheet = string.Equals(sheetName, "元件汇总表", StringComparison.OrdinalIgnoreCase);

                // 获取当前聚焦的活动单元格对象
                dynamic? activeCell = app.ActiveCell;
                // 校验活动单元格有效性
                if (activeCell == null)
                {
                    return (false, "请先在工作表中点击选中目标元器件行", 0, "AA", "AB", sheetName);
                }

                // 读取当前活动单元格的物理行号与列号
                int activeRow = Convert.ToInt32(activeCell.Row);
                int activeCol = Convert.ToInt32(activeCell.Column);

                // 处理图纸名称：若要求去除扩展名则提取无后缀纯文件名
                string cleanDwgName = removeExt ? Path.GetFileNameWithoutExtension(dwgName) : (dwgName ?? string.Empty);
                string cleanDirName = dirName ?? string.Empty;

                // 核心规则：根据当前工作表智能自适应决定回写的目标列
                // 1. 若当前在【元件汇总表】：固定使用 X 列 (参数1 图纸) 与 Y 列 (参数2 目录)
                // 2. 若当前在普通【分类表】：默认使用 AA 列 (图块名称 图纸) 与 AB 列 (图块类别 目录)
                string targetColDwg;
                string targetColDir;

                if (isSummarySheet)
                {
                    // 在元件汇总表中：图纸写入 X 列，目录写入 Y 列 --硬编码: 元件汇总表标准列位--
                    targetColDwg = "X";
                    targetColDir = "Y";
                }
                else
                {
                    // 在普通分类表中：优先使用 AA/AB 列，避免将极数(X)和脱扣方式(Y)冲掉
                    bool hasCustomDwg = !string.IsNullOrWhiteSpace(colDwg) && !string.Equals(colDwg, "X", StringComparison.OrdinalIgnoreCase);
                    bool hasCustomDir = !string.IsNullOrWhiteSpace(colDir) && !string.Equals(colDir, "Y", StringComparison.OrdinalIgnoreCase);
                    targetColDwg = hasCustomDwg ? colDwg!.Trim().ToUpper() : "AA";
                    targetColDir = hasCustomDir ? colDir!.Trim().ToUpper() : "AB";
                }

                // 将图纸名称写入当前行的目标图纸列
                sheet.Range[$"{targetColDwg}{activeRow}"].Value2 = cleanDwgName;

                // 将目录名称写入当前行的目标目录列
                sheet.Range[$"{targetColDir}{activeRow}"].Value2 = cleanDirName;

                // 记录下一行行号初始值
                int nextRow = activeRow;

                // 若开启自动下移一行特性，则驱动 Excel 焦点跳转
                if (autoNextRow)
                {
                    nextRow = activeRow + 1;
                    try
                    {
                        // 激活并选中下一行的相同列单元格
                        sheet.Cells[nextRow, activeCol].Select();
                    }
                    catch
                    {
                        // 若跳转失败尝试通过 Range 选中
                        try { sheet.Range[$"{targetColDwg}{nextRow}"].Select(); } catch { }
                    }
                }

                // 组装友好的回写模式提示信息
                string tableTag = isSummarySheet ? "【元件汇总表】" : (!string.IsNullOrWhiteSpace(sheetName) ? $"【{sheetName}】" : "【分类表】");
                return (true, $"已写入{tableTag}第 {activeRow} 行: [{targetColDwg}]={cleanDwgName}, [{targetColDir}]={cleanDirName}", nextRow, targetColDwg, targetColDir, sheetName);
            }
            catch (Exception ex)
            {
                // 记录写入异常日志
                LogHelper.WriteLog($"[ExcelServices] BindDwgParamToActiveRow 异常: {ex.Message}");
                // 返回失败状态与异常信息
                return (false, $"写入失败: {ex.Message}", 0, "AA", "AB", "");
            }
        }
    }
}
