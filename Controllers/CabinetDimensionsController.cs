using System;
using System.Collections.Generic;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 尺寸安装方式业务控制器
    /// 充当前端 Vue 3 界面与 ExcelServices 业务逻辑层之间的轻量解耦中介
    /// </summary>
    public class CabinetDimensionsController
    {
        // 声明向导窗口引用
        private readonly Forms.CabinetDimensionsForm _form;

        /// <summary>
        /// 构造函数: 注入窗体宿主引用
        /// </summary>
        /// <param name="form">尺寸安装方式向导窗体实例</param>
        public CabinetDimensionsController(Forms.CabinetDimensionsForm form)
        {
            // 保存窗体宿主引用
            _form = form;
        }

        /// <summary>
        /// 获取尺寸安装方式初始化数据 (包含分类列表、项目名称及已有表状态)
        /// 遵循规则 3: 具体 Excel 扫描逻辑委托公共服务 ExcelServices 处理
        /// </summary>
        /// <returns>尺寸安装方式初始化数据模型</returns>
        public CabinetDimensionsInitDataDto GetInitData()
        {
            // 调用服务层扫描工作簿
            return ExcelServices.GetCabinetDimensionsInitData();
        }

        /// <summary>
        /// 执行尺寸安装方式导出业务
        /// </summary>
        /// <param name="selectedCategorySheets">前端选中的分类工作表列表</param>
        /// <returns>导出执行结果</returns>
        public CabinetDimensionsExportResult ExportCabinetDimensions(List<string> selectedCategorySheets)
        {
            // 委托 ExcelServices 执行真正的模板克隆、数据汇总与矩阵回写
            return ExcelServices.ExportCabinetDimensionsToCurrentWorkbook(selectedCategorySheets);
        }
    }
}
