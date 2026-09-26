using System;
using System.Collections.Generic;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 三箱型号管道业务控制器
    /// 充当前端 Vue 3 界面与 ExcelServices 业务逻辑层之间的轻量解耦中介
    /// </summary>
    public class CabinetModelPipelineController
    {
        // 声明向导窗口弱引用
        private readonly Forms.CabinetModelPipelineForm _form;

        /// <summary>
        /// 构造函数: 注入窗体宿主引用
        /// </summary>
        /// <param name="form">三箱型号管道向导窗体实例</param>
        public CabinetModelPipelineController(Forms.CabinetModelPipelineForm form)
        {
            // 保存窗体宿主引用
            _form = form;
        }

        /// <summary>
        /// 获取三箱型号添写初始化数据 (包含分类表列表、管道规则、当前表箱柜判定数据)
        /// 遵循规则 3: 具体 Excel 扫描逻辑委托公共服务 ExcelServices 处理
        /// </summary>
        /// <param name="targetSheetName">可选指定的工作表名称</param>
        /// <returns>初始化数据模型</returns>
        public CabinetPipelineInitDataDto GetInitData(string? targetSheetName = null)
        {
            // 调用服务层扫描工作簿
            return ExcelServices.GetCabinetPipelineInitData(targetSheetName);
        }

        /// <summary>
        /// 切换分类工作表并重新扫描该表箱柜数据
        /// </summary>
        /// <param name="sheetName">选中的工作表名称</param>
        /// <param name="config">当前管道规则配置</param>
        /// <returns>该分类表下的箱柜判定明细</returns>
        public List<CabinetPipelineItemDto> SwitchCategorySheet(string sheetName, CabinetPipelineConfig config)
        {
            try
            {
                // 获取当前活动 Excel 上下文
                var context = Tool.GetActiveExcelContext();
                if (context == null || string.IsNullOrWhiteSpace(sheetName)) return new List<CabinetPipelineItemDto>();

                // 获取工作簿与目标工作表 dynamic 引用 (容错避免 COM 转换异常)
                dynamic wb = context.Wb;
                dynamic ws = wb.Worksheets[sheetName];
                if (ws == null) return new List<CabinetPipelineItemDto>();

                // 委托服务层执行扫描与管道推导
                return ExcelServices.ScanAndEvaluateCabinetsForSheet(ws, wb, config);
            }
            catch (Exception ex)
            {
                // 记录切换分类表异常日志
                LogHelper.WriteLog($"[三箱管道控制器] 切换分类表异常: {ex.Message}");
                return new List<CabinetPipelineItemDto>();
            }
        }

        /// <summary>
        /// 保存用户配置的管道规则与策略至本地 JSON 文件
        /// </summary>
        /// <param name="config">管道配置实体</param>
        /// <returns>保存是否成功</returns>
        public bool SaveConfig(CabinetPipelineConfig config)
        {
            // 调用服务层持久化配置
            return ExcelServices.SaveCabinetPipelineConfig(config);
        }

        /// <summary>
        /// 执行三箱型号批量回写至 Excel 工作表
        /// </summary>
        /// <param name="items">待回写的条目集合</param>
        /// <param name="writeSummaryN">是否回填至汇总行 N 列 (默认 true)</param>
        /// <param name="writeSummaryD">是否写入汇总行 D 列</param>
        /// <param name="writeDetailD">是否写入明细行 D 列</param>
        /// <returns>回写执行结果</returns>
        public CabinetPipelineApplyResult BatchApplyModels(
            List<CabinetPipelineApplyItem> items,
            bool writeSummaryN = true,
            bool writeSummaryD = false,
            bool writeDetailD = false)
        {
            // 委托 ExcelServices 执行定点批量回写
            return ExcelServices.BatchApplyCabinetModelsToExcel(items, writeSummaryN, writeSummaryD, writeDetailD);
        }
    }
}
