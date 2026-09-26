using System;
using System.Collections.Generic;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 人工清单业务控制器
    /// 充当前端 Vue 3 界面与 ExcelServices 业务逻辑层之间的轻量解耦中介
    /// </summary>
    public class LaborListController
    {
        // 声明向导窗口弱引用
        private readonly Forms.LaborListForm _form;

        /// <summary>
        /// 构造函数: 注入窗体宿主引用
        /// </summary>
        /// <param name="form">人工清单向导窗体实例</param>
        public LaborListController(Forms.LaborListForm form)
        {
            // 保存窗体宿主引用
            _form = form;
        }

        /// <summary>
        /// 获取人工清单初始化数据 (包含分类列表、项目名称及已有表状态)
        /// 遵循规则 3: 具体 Excel 扫描逻辑委托公共服务 ExcelServices 处理
        /// </summary>
        /// <returns>人工清单初始化数据模型</returns>
        public LaborListInitDataDto GetInitData()
        {
            // 调用服务层扫描工作簿
            return ExcelServices.GetLaborListInitData();
        }

        /// <summary>
        /// 执行人工清单导出业务
        /// </summary>
        /// <param name="selectedCategorySheets">前端选中的分类工作表列表</param>
        /// <returns>导出执行结果</returns>
        public LaborListExportResult ExportLaborList(List<string> selectedCategorySheets)
        {
            // 委托 ExcelServices 执行真正的模板克隆、数据汇总与矩阵回写
            return ExcelServices.ExportLaborListToCurrentWorkbook(selectedCategorySheets);
        }
    }
}
