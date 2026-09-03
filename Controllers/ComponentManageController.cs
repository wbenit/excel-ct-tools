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
        /// 获取电气元器件品牌统计列表 (支持根据数据源切换云端或本地 SQLite 个人库)
        /// </summary>
        /// <param name="dataSource">物料数据源: "cloud" 或 "personal"</param>
        /// <returns>品牌列表及其包含数量</returns>
        public List<BrandStatItemDto> GetBrandStats(string dataSource = "cloud")
        {
            // 判断是否为本地个人库
            if (string.Equals(dataSource, "personal", StringComparison.OrdinalIgnoreCase))
            {
                // 调用本地 SQLite 个人物料库服务
                return Services.PersonalComponentDbService.GetBrandStats();
            }

            // 默认调用远程 API 客户端拉取云端品牌数据
            return ComponentApiClient.GetBrandStats();
        }

        /// <summary>
        /// 根据选中的品牌获取该品牌下的所有元器件名称列表 (支持数据源切换)
        /// </summary>
        /// <param name="brand">选中的品牌</param>
        /// <param name="dataSource">物料数据源: "cloud" 或 "personal"</param>
        /// <returns>元器件名称列表</returns>
        public List<string> GetNamesByBrand(string? brand, string dataSource = "cloud")
        {
            // 判断是否为本地个人库
            if (string.Equals(dataSource, "personal", StringComparison.OrdinalIgnoreCase))
            {
                // 调用本地 SQLite 个人物料库去重名称服务
                return Services.PersonalComponentDbService.GetNamesByBrand(brand);
            }

            // 默认调用 API 客户端拉取云端品牌的元器件名称
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
        /// 根据品牌与名称关键字筛选并全量拉取数据灌入 Excel 表格 (支持数据源切换)
        /// </summary>
        /// <param name="brand">选中的品牌</param>
        /// <param name="nameKeyword">名称筛选关键字</param>
        /// <param name="dataSource">物料数据源: "cloud" 或 "personal"</param>
        /// <returns>拉取并写入结果</returns>
        public ComponentManageActionResult LoadComponents(string? brand, string? nameKeyword, string dataSource = "cloud")
        {
            // 调用业务服务层执行数据拉取与表格写入
            return ExcelServices.LoadComponentsToSheet(brand, nameKeyword, dataSource);
        }

        /// <summary>
        /// 对当前选中的 1 行或多行执行定向【更新】提交 (支持数据源切换)
        /// </summary>
        /// <param name="dataSource">物料数据源: "cloud" 或 "personal"</param>
        /// <returns>更新执行结果报告</returns>
        public ComponentManageActionResult UpdateSelected(string dataSource = "cloud")
        {
            // 调度服务层执行选中行更新
            return ExcelServices.UpdateSelectedComponents(dataSource);
        }

        /// <summary>
        /// 对当前选中的 1 行或多行执行定向【新增】提交 (支持数据源切换)
        /// </summary>
        /// <param name="dataSource">物料数据源: "cloud" 或 "personal"</param>
        /// <returns>新增执行结果报告</returns>
        public ComponentManageActionResult CreateSelected(string dataSource = "cloud")
        {
            // 调度服务层执行选中行新增
            return ExcelServices.CreateSelectedComponents(dataSource);
        }

        /// <summary>
        /// 对当前选中的 1 行或多行执行定向【删除】 (支持数据源切换)
        /// </summary>
        /// <param name="dataSource">物料数据源: "cloud" 或 "personal"</param>
        /// <returns>删除执行结果报告</returns>
        public ComponentManageActionResult DeleteSelected(string dataSource = "cloud")
        {
            // 调度服务层执行选中行删除
            return ExcelServices.DeleteSelectedComponents(dataSource);
        }
    }
}
