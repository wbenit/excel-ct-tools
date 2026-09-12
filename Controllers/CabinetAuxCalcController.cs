using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelAddInDemo.Models;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 工作表及箱柜选择条目 DTO
    /// </summary>
    public class CabinetChoiceDto
    {
        // 分类工作表名称
        [JsonPropertyName("sheetName")]
        public string SheetName { get; set; } = string.Empty;

        // 箱柜定义名称 (如 Cab_Det_1)
        [JsonPropertyName("detName")]
        public string DetName { get; set; } = string.Empty;

        // 箱柜名称或编号 (如 AP1)
        [JsonPropertyName("cabinetName")]
        public string CabinetName { get; set; } = string.Empty;

        // 箱柜物理行号 (Cab_Det 行)
        [JsonPropertyName("detRow")]
        public int DetRow { get; set; }
    }

    /// <summary>
    /// 前端交互上下文初始数据 DTO
    /// </summary>
    public class CabinetAuxCalcContextDto
    {
        // 当前活动工作表名称
        [JsonPropertyName("activeSheetName")]
        public string ActiveSheetName { get; set; } = string.Empty;

        // 当前工作表下的箱柜列表
        [JsonPropertyName("cabinets")]
        public List<CabinetChoiceDto> Cabinets { get; set; } = new List<CabinetChoiceDto>();

        // 所有分类工作表名称列表
        [JsonPropertyName("sheetNames")]
        public List<string> SheetNames { get; set; } = new List<string>();

        // 当前加载的规则配置
        [JsonPropertyName("rules")]
        public QuotationRules Rules { get; set; } = new QuotationRules();
    }

    /// <summary>
    /// 辅材壳体计算控制器，负责处理与前端 WebView2 之间的消息分发与业务调用
    /// </summary>
    public class CabinetAuxCalcController
    {
        /// <summary>
        /// 获取前端初始化所需的上下文数据（工作表列表、箱柜列表、当前规则配置）
        /// </summary>
        /// <returns>CabinetAuxCalcContextDto 数据实体</returns>
        public CabinetAuxCalcContextDto GetInitialContext()
        {
            var context = new CabinetAuxCalcContextDto();
            try
            {
                // 获取 Excel Application 实例
                var app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return context;

                // 获取活动工作簿
                Workbook wb = app.ActiveWorkbook;
                if (wb == null) return context;

                // 获取活动工作表
                Worksheet activeWs = app.ActiveSheet as Worksheet;
                if (activeWs != null)
                {
                    context.ActiveSheetName = activeWs.Name;
                }

                // 遍历工作簿中所有有效分类表
                foreach (Worksheet ws in wb.Worksheets)
                {
                    if (ws.Visible != XlSheetVisibility.xlSheetVisible) continue;
                    if (ws.Name == "项目信息" || ws.Name == "元件汇总表" || ws.Name == "元器件数据管理") continue;

                    // 加入分类工作表名称列表
                    context.SheetNames.Add(ws.Name);

                    // 若为当前活动表，提取该表下的箱柜列表
                    if (activeWs != null && ws.Name == activeWs.Name)
                    {
                        var validCabinets = Tool.GetSheetValidCabinets(ws, wb);
                        foreach (var kvp in validCabinets)
                        {
                            int cabIndex = kvp.Key;
                            var anchor = kvp.Value;
                            if (anchor?.Det == null) continue;
                            int detRow = Convert.ToInt32(anchor.Det.Row);
                            string cabTitle = ws.Range[$"A{detRow}"].Value?.ToString() ?? $"箱柜{cabIndex}";
                            context.Cabinets.Add(new CabinetChoiceDto
                            {
                                SheetName = ws.Name,
                                DetName = $"Cab_Det_{cabIndex}",
                                CabinetName = cabTitle,
                                DetRow = detRow
                            });
                        }
                    }
                }

                // 加载当前规则配置
                context.Rules = ExcelServices.LoadQuotationRules();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取辅材计算初始上下文异常: {ex.Message}");
            }
            return context;
        }

        /// <summary>
        /// 切换工作表时获取该工作表下的箱柜列表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <returns>箱柜列表</returns>
        public List<CabinetChoiceDto> GetCabinetsBySheet(string sheetName)
        {
            var list = new List<CabinetChoiceDto>();
            try
            {
                var app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return list;
                Workbook wb = app.ActiveWorkbook;
                if (wb == null) return list;

                Worksheet ws = wb.Worksheets[sheetName] as Worksheet;
                if (ws == null) return list;

                var validCabinets = Tool.GetSheetValidCabinets(ws, wb);
                foreach (var kvp in validCabinets)
                {
                    int cabIndex = kvp.Key;
                    var anchor = kvp.Value;
                    if (anchor?.Det == null) continue;
                    int detRow = Convert.ToInt32(anchor.Det.Row);
                    string cabTitle = ws.Range[$"A{detRow}"].Value?.ToString() ?? $"箱柜{cabIndex}";
                    list.Add(new CabinetChoiceDto
                    {
                        SheetName = ws.Name,
                        DetName = $"Cab_Det_{cabIndex}",
                        CabinetName = cabTitle,
                        DetRow = detRow
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取指定表箱柜列表异常: {ex.Message}");
            }
            return list;
        }

        /// <summary>
        /// 保存并更新规则配置
        /// </summary>
        /// <param name="rules">规则实体</param>
        /// <returns>是否保存成功</returns>
        public bool SaveRules(QuotationRules rules)
        {
            if (rules == null) return false;
            try
            {
                ExcelServices.SaveQuotationRules(rules);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"保存规则配置异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 对指定箱柜进行智能推导分析
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="detName">箱柜Det定义名称</param>
        /// <param name="rules">计算规则</param>
        /// <returns>CabinetCalcResult 结果实体</returns>
        public CabinetCalcResult? AnalyzeSingleCabinet(string sheetName, string detName, QuotationRules rules)
        {
            try
            {
                var app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return null;
                Workbook wb = app.ActiveWorkbook;
                if (wb == null) return null;

                Worksheet ws = wb.Worksheets[sheetName] as Worksheet;
                if (ws == null) return null;

                var scanData = ExcelServices.ScanCabinetData(ws, detName);
                if (scanData == null) return null;

                var result = ExcelServices.CalculateCabinetAuxAndShell(scanData, rules);
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"推导单个箱柜异常: {ex.Message}");
                return null;
            }
        }


        /// <summary>
        /// 写入单个选中的箱柜数据与公式至工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="detName">箱柜 Det 名称 (如 Cab_Det_1)</param>
        /// <param name="rules">定额与规则</param>
        /// <returns>操作结果元组</returns>
        public (bool Success, string CabinetName, string Message) WriteSingleCabinet(
            string sheetName,
            string detName,
            QuotationRules rules)
        {
            try
            {
                // 获取 Excel COM 实例
                var app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return (false, string.Empty, "未能连接到 Excel 应用程序。");

                // 获取活动工作簿
                Workbook wb = app.ActiveWorkbook;
                if (wb == null) return (false, string.Empty, "当前没有打开的工作簿。");

                // 获取指定工作表
                Worksheet ws = wb.Worksheets[sheetName] as Worksheet;
                if (ws == null) return (false, string.Empty, $"未找到工作表【{sheetName}】。");

                // 调用服务层写入单台箱柜方法
                return ExcelServices.WriteSingleCabinetAuxAndShell(ws, detName, rules);
            }
            catch (Exception ex)
            {
                // 捕获并记录异常
                System.Diagnostics.Debug.WriteLine($"写入单台箱柜异常: {ex.Message}");
                return (false, string.Empty, $"写入过程发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 更新指定分类工作表中的所有箱柜
        /// </summary>
        /// <param name="sheetName">分类工作表名称</param>
        /// <param name="rules">定额与规则</param>
        /// <returns>更新结果元组</returns>
        public (bool Success, int UpdatedCabinets, string Message) UpdateCurrentCategory(
            string sheetName,
            QuotationRules rules)
        {
            try
            {
                // 获取 Excel COM 实例
                var app = ExcelDnaSafeAccessor.GetApplication();
                if (app == null) return (false, 0, "未能连接到 Excel 应用程序。");

                // 获取活动工作簿
                Workbook wb = app.ActiveWorkbook;
                if (wb == null) return (false, 0, "当前没有打开的工作簿。");

                // 获取目标工作表
                Worksheet ws = wb.Worksheets[sheetName] as Worksheet;
                if (ws == null) return (false, 0, $"未找到工作表【{sheetName}】。");

                // 调用服务层更新当前分类表方法
                return ExcelServices.UpdateCurrentCategoryAuxAndShell(ws, rules);
            }
            catch (Exception ex)
            {
                // 捕获并记录异常
                System.Diagnostics.Debug.WriteLine($"更新当前分类异常: {ex.Message}");
                return (false, 0, $"更新分类发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 更新当前工作簿中所有分类表的全部箱柜
        /// </summary>
        /// <param name="rules">定额与规则</param>
        /// <returns>全工作簿更新结果元组</returns>
        public (bool Success, int UpdatedSheets, int UpdatedCabinets, string Message) UpdateAllCategories(
            QuotationRules rules)
        {
            try
            {
                // 调用服务层更新全工作簿所有分类表方法
                return ExcelServices.UpdateAllCategoriesAuxAndShell(rules);
            }
            catch (Exception ex)
            {
                // 捕获并记录异常
                System.Diagnostics.Debug.WriteLine($"更新所有分类异常: {ex.Message}");
                return (false, 0, 0, $"更新所有分类发生错误: {ex.Message}");
            }
        }
    }
}
