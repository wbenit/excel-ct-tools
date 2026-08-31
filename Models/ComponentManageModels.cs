using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 元器件数据管理工作表默认配置与常量定义
    /// </summary>
    public static class ComponentManageDefaults
    {
        // 默认元器件管理专用工作表名称
        public const string DefaultSheetName = "元器件数据管理"; // --硬编码-- 默认工作表名称

        // 表头背景主题颜色: 绿蓝主色调 #009688 对应 OLE 颜色值 0x889600 (RGB: 0, 150, 136)
        public const int ThemeHeaderOleColor = 0x889600; // --硬编码-- 表头主题 OLE 颜色值 (RGB: 0, 150, 136)

        // 系统 ID 列浅灰只读底色 (RGB: 242, 244, 245)
        public const int IdColumnLightGrayOleColor = 0xF5F4F2; // --硬编码-- ID 列浅灰底色

        // 单页最大拉取限制数量
        public const int MaxQueryPageSize = 500; // --硬编码-- API 单页最大条数
    }

    /// <summary>
    /// 元器件数据管理工作表列字段索引映射配置模型
    /// </summary>
    public class ComponentManageColumnConfig
    {
        // 系统 ID 所在列 (默认第 1 列 A)
        public int IdCol { get; set; } = 1; // --硬编码-- 系统ID所在列索引

        // 品牌所在列 (默认第 2 列 B)
        public int BrandCol { get; set; } = 2; // --硬编码-- 品牌所在列索引

        // 元器件名称所在列 (默认第 3 列 C)
        public int NameCol { get; set; } = 3; // --硬编码-- 元器件名称所在列索引

        // 规格型号所在列 (默认第 4 列 D)
        public int ModelCol { get; set; } = 4; // --硬编码-- 规格型号所在列索引

        // 参考单价所在列 (默认第 5 列 E)
        public int PriceCol { get; set; } = 5; // --硬编码-- 参考单价所在列索引

        // 额定电流所在列 (默认第 6 列 F)
        public int CurrentCol { get; set; } = 6; // --硬编码-- 额定电流所在列索引

        // 极数所在列 (默认第 7 列 G)
        public int PolesCol { get; set; } = 7; // --硬编码-- 极数所在列索引

        // 脱扣方式所在列 (默认第 8 列 H)
        public int TrippingCol { get; set; } = 8; // --硬编码-- 脱扣方式所在列索引

        // 扩展参数 1 所在列 (默认第 9 列 I)
        public int Param1Col { get; set; } = 9; // --硬编码-- 扩展参数1所在列索引

        // 扩展参数 2 所在列 (默认第 10 列 J)
        public int Param2Col { get; set; } = 10; // --硬编码-- 扩展参数2所在列索引

        // 备注说明所在列 (默认第 11 列 K)
        public int RemarkCol { get; set; } = 11; // --硬编码-- 备注所在列索引

        // 操作与同步状态提示所在列 (默认第 12 列 L)
        public int StatusCol { get; set; } = 12; // --硬编码-- 状态提示所在列索引

        // 总列数
        public int TotalColumns => 12;
    }

    /// <summary>
    /// 元器件新增参数传输对象模型
    /// </summary>
    public class CreateComponentApiRequest
    {
        // 品牌/厂商 (必填)
        [JsonPropertyName("brand")]
        public string Brand { get; set; } = string.Empty;

        // 元器件名称 (必填)
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        // 规格型号 (必填)
        [JsonPropertyName("model")]
        public string? Model { get; set; }

        // 参考价格 (元)
        [JsonPropertyName("price")]
        public decimal Price { get; set; } = 0.00m;

        // 备注信息
        [JsonPropertyName("remark")]
        public string? Remark { get; set; }

        // 扩展参数 1
        [JsonPropertyName("param1")]
        public string? Param1 { get; set; }

        // 扩展参数 2
        [JsonPropertyName("param2")]
        public string? Param2 { get; set; }

        // 额定电流 (A)
        [JsonPropertyName("current")]
        public int? Current { get; set; }

        // 极数 (如: 3, 4)
        [JsonPropertyName("poles")]
        public string? Poles { get; set; }

        // 脱扣方式 (如: C, D, TM)
        [JsonPropertyName("tripping")]
        public string? Tripping { get; set; }
    }

    /// <summary>
    /// 元器件更新参数传输对象模型
    /// </summary>
    public class UpdateComponentApiRequest
    {
        // 待更新的元器件主键 ID (必填)
        [JsonPropertyName("id")]
        public int Id { get; set; }

        // 品牌/厂商
        [JsonPropertyName("brand")]
        public string? Brand { get; set; }

        // 元器件名称
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        // 规格型号
        [JsonPropertyName("model")]
        public string? Model { get; set; }

        // 参考价格 (元)
        [JsonPropertyName("price")]
        public decimal? Price { get; set; }

        // 备注信息
        [JsonPropertyName("remark")]
        public string? Remark { get; set; }

        // 扩展参数 1
        [JsonPropertyName("param1")]
        public string? Param1 { get; set; }

        // 扩展参数 2
        [JsonPropertyName("param2")]
        public string? Param2 { get; set; }

        // 额定电流 (A)
        [JsonPropertyName("current")]
        public int? Current { get; set; }

        // 极数
        [JsonPropertyName("poles")]
        public string? Poles { get; set; }

        // 脱扣方式
        [JsonPropertyName("tripping")]
        public string? Tripping { get; set; }
    }

    /// <summary>
    /// Excel 选区行信息探测响应模型
    /// </summary>
    public class SelectionDetectResult
    {
        // 当前选区是否在元器件管理工作表中
        public bool IsInManageSheet { get; set; }

        // 当前工作表名称
        public string SheetName { get; set; } = string.Empty;

        // 选中的有效数据行数 (已排除表头第 1 行)
        public int SelectedRowCount { get; set; }

        // 选中的物理行号列表
        public List<int> RowIndices { get; set; } = new List<int>();

        // 提示消息
        public string Message { get; set; } = string.Empty;
    }

    /// <summary>
    /// 元器件批量操作执行响应模型
    /// </summary>
    public class ComponentManageActionResult
    {
        // 执行是否成功
        public bool Success { get; set; } = true;

        // 成功处理的记录条数
        public int SuccessCount { get; set; }

        // 失败的记录条数
        public int FailCount { get; set; }

        // 详细结果提示文本
        public string Message { get; set; } = string.Empty;
    }
}
