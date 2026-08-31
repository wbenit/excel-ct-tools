using System;
using System.Collections.Generic;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 元器件数据管理控制器 (供 Vue 3 前端窗口进行双向通信交互)
    /// </summary>
    public class ComponentManageController
    {
        /// <summary>
        /// 从远程商城 WebAPI 获取所有电气元器件品牌统计列表
        /// </summary>
        /// <returns>品牌列表及其包含数量</returns>
        public List<BrandStatItemDto> GetBrandStats()
        {
            // 调用 API 客户端拉取品牌聚合数据
            return ComponentApiClient.GetBrandStats();
        }

        /// <summary>
        /// 根据选中的品牌获取该品牌下的所有元器件名称列表
        /// </summary>
        /// <param name="brand">选中的品牌</param>
        /// <returns>元器件名称列表</returns>
        public List<string> GetNamesByBrand(string? brand)
        {
            // 调用 API 客户端拉取该品牌的元器件名称
            return ComponentApiClient.GetNamesByBrand(brand);
        }

        /// <summary>
        /// 探测当前 Excel 选区中覆盖的行数与行号信息
        /// </summary>
        /// <returns>选区检测结果对象</returns>
        public SelectionDetectResult DetectSelection()
        {
            // 调度服务层执行选区探测
            return ExcelServices.DetectCurrentSelection();
        }

        /// <summary>
        /// 根据品牌与名称关键字筛选并全量拉取数据灌入 Excel 表格
        /// </summary>
        /// <param name="brand">选中的品牌</param>
        /// <param name="nameKeyword">名称筛选关键字</param>
        /// <returns>拉取并写入结果</returns>
        public ComponentManageActionResult LoadComponents(string? brand, string? nameKeyword)
        {
            // 调用业务服务层执行数据拉取与表格写入
            return ExcelServices.LoadComponentsToSheet(brand, nameKeyword);
        }

        /// <summary>
        /// 对当前选中的 1 行或多行执行定向【更新】提交
        /// </summary>
        /// <returns>更新执行结果报告</returns>
        public ComponentManageActionResult UpdateSelected()
        {
            // 调度服务层执行选中行更新
            return ExcelServices.UpdateSelectedComponents();
        }

        /// <summary>
        /// 对当前选中的 1 行或多行执行定向【新增】提交
        /// </summary>
        /// <returns>新增执行结果报告</returns>
        public ComponentManageActionResult CreateSelected()
        {
            // 调度服务层执行选中行新增
            return ExcelServices.CreateSelectedComponents();
        }

        /// <summary>
        /// 对当前选中的 1 行或多行执行定向【删除】
        /// </summary>
        /// <returns>删除执行结果报告</returns>
        public ComponentManageActionResult DeleteSelected()
        {
            // 调度服务层执行选中行删除
            return ExcelServices.DeleteSelectedComponents();
        }
    }
}
