using System;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：AutoCAD 5:5 均等分屏协同调度接口
    /// 满足规则 3：所有窗口交互调度由公共业务层统一暴露
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 切换 Excel 与 AutoCAD 5:5 均等智能并排分屏
        /// 彻底去除右侧任务窗格面板，点击瞬间 5:5 铺满屏幕，再点一次还原 Excel 全屏最大化
        /// </summary>
        public static void ToggleCadSideBySide()
        {
            // 调度 CadEmbedManager 执行 5:5 均等分屏切换
            CadEmbedManager.ToggleSideBySide(0.5);
        }

        /// <summary>
        /// 兼容接口：触发 5:5 均等分屏切换
        /// </summary>
        public static void ToggleCadTaskPane()
        {
            // 执行 5:5 均等分屏切换
            ToggleCadSideBySide();
        }
    }
}
