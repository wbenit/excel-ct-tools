using System;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：AutoCAD 协同任务窗格宿主与窗口联动接口
    /// 满足规则 3：所有窗口交互调度由公共业务层统一暴露
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 切换 AutoCAD 协同画图任务窗格的显示与隐藏
        /// </summary>
        public static void ToggleCadTaskPane()
        {
            // 调度 CadEmbedManager 执行右侧任务窗格切换与自动嵌入
            CadEmbedManager.ToggleTaskPane();
        }

        /// <summary>
        /// 显式展开 AutoCAD 协同画图任务窗格并嵌入当前图纸
        /// </summary>
        public static void ShowCadTaskPane()
        {
            // 显式展开任务窗格
            CadEmbedManager.ShowTaskPane();
        }

        /// <summary>
        /// 将 AutoCAD 从任务窗格还原至桌面独立窗口
        /// </summary>
        public static void DetachCadFromTaskPane()
        {
            // 释放并还原 AutoCAD 窗口至桌面
            CadEmbedManager.DetachCad();
        }
    }
}
