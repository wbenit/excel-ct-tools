using System;
using System.IO;
using System.Reflection;
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
            // 注册全局程序集动态解析句柄，防止单文件加载项找不到依赖 DLL
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
        }

        /// <summary>
        /// 当 Excel 卸载本插件时触发此方法
        /// </summary>
        public void AutoClose()
        {
            // 注销全局程序集动态解析句柄
            AppDomain.CurrentDomain.AssemblyResolve -= OnAssemblyResolve;
        }

        /// <summary>
        /// 程序集缺失时的动态寻找与加载回调
        /// </summary>
        private static Assembly? OnAssemblyResolve(object? sender, ResolveEventArgs args)
        {
            try
            {
                // 获取缺少程序集的简单名称
                string assemblyName = new AssemblyName(args.Name).Name + ".dll";

                // 计算当前 BaseDirectory 绝对路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // 拼接程序集目标路径
                string targetPath = Path.Combine(baseDir, assemblyName);

                // 判断目标路径下的 DLL 是否存在
                if (File.Exists(targetPath))
                {
                    // 动态从硬盘加载所需的程序集
                    return Assembly.LoadFrom(targetPath);
                }
            }
            catch
            {
                // 忽略解析过程中的捕获异常
            }
            return null;
        }
    }
}
