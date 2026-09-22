using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 元器件整行数据交换实体模型 (对标电小二 ElementExchangeMode)
    /// 用于剪切、复制、跨箱柜/同箱柜插入时携带完整业务属性与隐藏列参数
    /// </summary>
    public class ComponentRowExchangeDto
    {
        // 标记当前是否为剪切模式 (true 为剪切，粘贴插入后需删除原行；false 为复制)
        public bool IsCutMode { get; set; } = false;

        // 来源工作簿名称，支持跨工作簿环境比对
        public string SourceWorkbookName { get; set; } = string.Empty;

        // 来源工作表名称 (例如 "动力"、"照明")
        public string SourceSheetName { get; set; } = string.Empty;

        // 来源箱柜序号 K (例如 Cab_Det_1 中的 1，用于判断是否跨箱柜)
        public int SourceCabinetK { get; set; } = 0;

        // 来源所在物理行号 (从 1 开始的 Excel 绝对行号，剪切模式删除原行时使用)
        public int SourceRowIndex { get; set; } = 0;

        // 整行读取的总列数 (默认覆盖前台可见列与后台隐藏列)
        public int ColumnCount { get; set; } = 0;

        // 整行完整单元格数据矩阵 (1 行 N 列二维数组，遵循规则 7 一次性读写)
        public object[,] FullRowValues { get; set; } = new object[0, 0];

        // 来源行中包含公式的单元格集合 (Key 为 1-based 物理列号，Value 为原公式字符串)
        public Dictionary<int, string> CellFormulas { get; set; } = new Dictionary<int, string>();

        // 来源元器件绑定的 AutoCAD 图纸图元句柄 (第 30 列，AD 列)
        public string CadHandle { get; set; } = string.Empty;

        // 来源元器件辅助图元句柄 (第 31 列，AE 列)
        public string HandleB { get; set; } = string.Empty;

        // 原始元器件名称 (B 列文本，供界面提示展示)
        public string ComponentName { get; set; } = string.Empty;

        // 原始规格型号 (C 列文本，供界面提示展示)
        public string ComponentModel { get; set; } = string.Empty;
    }

    /// <summary>
    /// 全局静态元器件内存剪贴板管理器
    /// 集中管理元器件整行数据的复制、剪切与暂存状态
    /// </summary>
    public static class ComponentClipboardManager
    {
        // 内存中缓存的元器件数据包实例 (线程安全静态单例)
        private static ComponentRowExchangeDto? _currentClipboardItem;

        // 对象锁，防止并发读写引发竞态
        private static readonly object _syncLock = new object();

        /// <summary>
        /// 将元器件整行数据包暂存入剪贴板
        /// </summary>
        /// <param name="data">元器件交换实体对象</param>
        public static void Set(ComponentRowExchangeDto data)
        {
            // 加锁保障原子写入
            lock (_syncLock)
            {
                // 缓存数据包
                _currentClipboardItem = data;
            }
        }

        /// <summary>
        /// 从剪贴板中获取当前暂存的元器件数据包 (浅拷贝引用)
        /// </summary>
        /// <returns>当前元器件数据包，若无则返回 null</returns>
        public static ComponentRowExchangeDto? Get()
        {
            // 加锁安全读取
            lock (_syncLock)
            {
                // 返回当前缓存引用
                return _currentClipboardItem;
            }
        }

        /// <summary>
        /// 检查当前剪贴板中是否存在有效的元器件数据
        /// </summary>
        public static bool HasData
        {
            get
            {
                // 加锁检测是否存在数据
                lock (_syncLock)
                {
                    // 数据包非空且包含有效数据行
                    return _currentClipboardItem != null && _currentClipboardItem.FullRowValues != null && _currentClipboardItem.FullRowValues.Length > 0;
                }
            }
        }

        /// <summary>
        /// 清空当前剪贴板数据
        /// </summary>
        public static void Clear()
        {
            // 加锁安全置空
            lock (_syncLock)
            {
                // 置空缓存对象
                _currentClipboardItem = null;
            }
        }
    }
}
