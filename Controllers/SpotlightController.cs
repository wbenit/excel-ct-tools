using System;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 聚光灯 (Spotlight) 业务控制器，负责外观设置的读取、实时预览、保存持久化与默认重置
    /// </summary>
    public class SpotlightController
    {
        /// <summary>
        /// 获取当前生效的聚光灯全域配置模型
        /// </summary>
        /// <returns>聚光灯配置对象实例</returns>
        public SpotlightConfig GetConfig()
        {
            // 从全局单例中读取最新的聚光灯配置
            return SpotlightConfig.Current;
        }

        /// <summary>
        /// 保存并持久化用户在设置面板调整后的聚光灯参数
        /// </summary>
        /// <param name="config">前端提交的新配置模型</param>
        public void SaveConfig(SpotlightConfig config)
        {
            // 若传入配置为空则直接返回
            if (config == null) return;

            // 全局更新当前运行的单例配置并写回本地磁盘
            SpotlightConfig.UpdateCurrent(config);

            // 触发 Excel 业务层即时重绘应用新配置
            ExcelServices.ApplySpotlightConfig(config);
        }

        /// <summary>
        /// 将聚光灯外观与模式参数重置为系统出厂推荐默认值
        /// </summary>
        /// <returns>重置后的全新配置对象</returns>
        public SpotlightConfig ResetDefault()
        {
            // 调用模型静态方法恢复默认参数并持久化
            var defaultConfig = SpotlightConfig.ResetToDefault();

            // 触发 Excel 聚光灯图层立即应用默认样式
            ExcelServices.ApplySpotlightConfig(defaultConfig);

            // 将重置后的配置返回给前端更新 Vue 响应式状态
            return defaultConfig;
        }

        /// <summary>
        /// 实时联动预览：在不破坏持久化存储的前提下，在 Excel 中实时应用临时颜色与透明度
        /// </summary>
        /// <param name="previewConfig">用户正在调整的临时外观配置</param>
        public void ApplyPreview(SpotlightConfig previewConfig)
        {
            // 校验临时配置对象有效性
            if (previewConfig == null) return;

            // 调用 Excel 业务层提供的轻量级临时样式渲染通道
            ExcelServices.ApplyTemporarySpotlightStyle(previewConfig);
        }

        /// <summary>
        /// 快捷切换聚光灯高亮模式 (十字 / 仅行 / 仅列)
        /// </summary>
        /// <param name="mode">目标高亮模式</param>
        public void SetMode(SpotlightMode mode)
        {
            // 读取当前配置
            var current = SpotlightConfig.Current;
            // 更新高亮模式字段
            current.Mode = mode;
            // 持久化保存
            current.SaveToDisk();
            // 立即刷新 Excel 高亮区域裁剪与重绘
            ExcelServices.ApplySpotlightConfig(current);
        }
    }
}
