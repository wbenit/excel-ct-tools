using ExcelDna.Integration;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 插件主入口类，实现 IExcelAddIn 接口响应加载与卸载事件
    /// </summary>
    public class AddInMain : IExcelAddIn
    {
        /// <summary>
        /// 当 Excel 加载本插件 XLL 时触发此方法
        /// </summary>
        public void AutoOpen()
        {
            // 执行插件启动初始化，例如加载本地加密配置文件或初始化全局服务
            // 核心初始化业务逻辑均可完美支持商业加壳防逆向保护
        }

        /// <summary>
        /// 当 Excel 卸载本插件时触发此方法
        /// </summary>
        public void AutoClose()
        {
            // 执行插件注销或关闭时的资源释放操作
        }
    }
}
