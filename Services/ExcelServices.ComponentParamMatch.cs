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
        /// 将指定目录名称与 DWG 图纸名称即时写入当前活动工作表的当前行指定列，并可自动跳转到下一行
        /// </summary>
        /// <param name="dirName">当前目录名称</param>
        /// <param name="dwgName">DWG 图纸文件名称</param>
        /// <param name="autoNextRow">是否在写入后自动下移一行</param>
        /// <param name="colDir">目录名称写入的列标 (默认 Y)</param>
        /// <param name="colDwg">图纸名称写入的列标 (默认 X)</param>
        /// <param name="removeExt">是否去除 .dwg 扩展名</param>
        /// <returns>写入结果状态、反馈消息与最新激活的行号</returns>
        public static (bool Success, string Message, int NextRow) BindDwgParamToActiveRow(
            string dirName,
            string dwgName,
            bool autoNextRow = true,
            string colDir = "Y",
            string colDwg = "X",
            bool removeExt = true)
        {
            try
            {
                // 获取 Excel 宿主全局 Application 对象句柄
                dynamic? app = ExcelDnaUtil.Application;
                // 校验 Excel 宿主对象有效性
                if (app == null)
                {
                    return (false, "Excel Application 未就绪", 0);
                }

                // 获取当前活动的工作表对象
                dynamic? sheet = app.ActiveSheet;
                // 校验工作表有效性
                if (sheet == null)
                {
                    return (false, "当前无有效活动工作表", 0);
                }

                // 获取当前聚焦的活动单元格对象
                dynamic? activeCell = app.ActiveCell;
                // 校验活动单元格有效性
                if (activeCell == null)
                {
                    return (false, "请先在工作表中点击选中目标元器件行", 0);
                }

                // 读取当前活动单元格的物理行号与列号
                int activeRow = Convert.ToInt32(activeCell.Row);
                int activeCol = Convert.ToInt32(activeCell.Column);

                // 处理图纸名称：若要求去除扩展名则提取无后缀纯文件名
                string cleanDwgName = removeExt ? Path.GetFileNameWithoutExtension(dwgName) : (dwgName ?? string.Empty);
                string cleanDirName = dirName ?? string.Empty;

                // 标准化目标写入列标字母（图纸列默认 X，目录列默认 Y）
                string targetColDwg = string.IsNullOrWhiteSpace(colDwg) ? "X" : colDwg.Trim().ToUpper();
                string targetColDir = string.IsNullOrWhiteSpace(colDir) ? "Y" : colDir.Trim().ToUpper();

                // 将图纸名称写入当前行的目标图纸列 (默认 X 列)
                sheet.Range[$"{targetColDwg}{activeRow}"].Value2 = cleanDwgName;

                // 将目录名称写入当前行的目标目录列 (默认 Y 列)
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

                // 返回成功响应及写入详情
                return (true, $"已写入第 {activeRow} 行: [{targetColDwg}]={cleanDwgName}, [{targetColDir}]={cleanDirName}", nextRow);
            }
            catch (Exception ex)
            {
                // 记录写入异常日志
                LogHelper.WriteLog($"[ExcelServices] BindDwgParamToActiveRow 异常: {ex.Message}");
                // 返回失败状态与异常信息
                return (false, $"写入失败: {ex.Message}", 0);
            }
        }
    }
}
