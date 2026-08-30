using System;
using System.Collections.Generic;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 元器件物料匹配与高级过滤规则控制器 (供 Vue 3 前端窗口交互调用)
    /// </summary>
    public class ComponentMatchController
    {
        /// <summary>
        /// 从远程商城 WebAPI 获取所有电气元器件品牌及其包含的数据条数统计
        /// </summary>
        /// <returns>品牌及其数量统计列表</returns>
        public List<BrandStatItemDto> GetBrandStats()
        {
            // 调用 API 客户端获取品牌统计
            return ComponentApiClient.GetBrandStats();
        }

        /// <summary>
        /// 从本地持久化存储加载物料匹配过滤规则配置
        /// </summary>
        /// <returns>物料匹配过滤配置实例</returns>
        public ComponentMatchFilterConfig LoadConfig()
        {
            // 调用服务层读取本地配置
            return ExcelServices.LoadComponentMatchFilterConfig();
        }

        /// <summary>
        /// 将物料匹配过滤规则配置保存至本地磁盘
        /// </summary>
        /// <param name="config">物料匹配过滤配置对象</param>
        public bool SaveConfig(ComponentMatchFilterConfig config)
        {
            if (config == null) return false;
            // 调用服务层保存配置
            ExcelServices.SaveComponentMatchFilterConfig(config);
            return true;
        }

        /// <summary>
        /// 单条参数实时模拟匹配测试 (用于在前端界面即时预览过滤效果)
        /// </summary>
        /// <param name="name">元器件名称</param>
        /// <param name="current">额定电流</param>
        /// <param name="pole">极数</param>
        /// <param name="tripMode">脱扣方式</param>
        /// <param name="brand">指定品牌</param>
        /// <param name="rules">动态必含字段规则列表</param>
        /// <returns>匹配到的物料列表</returns>
        public List<ComponentApiDto> TestMatch(
            string name,
            string current,
            string pole,
            string tripMode,
            string? brand,
            List<MustContainRule>? rules)
        {
            // 调用 API 客户端执行多维检索与管道过滤
            return ComponentApiClient.QueryComponents(name, current, pole, tripMode, brand, rules);
        }

        /// <summary>
        /// 触发 Excel 当前选区行批量执行参数识别、物料库反查与多列自动回填
        /// </summary>
        /// <param name="config">当前界面生效的过滤配置</param>
        /// <returns>批量执行结果报告</returns>
        public BatchMatchExecuteResult ExecuteBatch(ComponentMatchFilterConfig? config)
        {
            // 若传入了配置则同步保存
            if (config != null)
            {
                SaveConfig(config);
            }

            // 调度业务服务层执行 Excel 选区批量处理
            return ExcelServices.ExecuteBatchMatchWithDb(config);
        }
    }
}
