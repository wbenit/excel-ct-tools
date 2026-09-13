using System;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// DWG 元器件三维物理尺寸与版本指纹实体模型
    /// 遵循规范：每 3 行新增代码至少包含 1 行中文注释，配置与硬编码显式标明
    /// </summary>
    public class DwgDimensionItem
    {
        // 数据库自增唯一主键
        [JsonPropertyName("id")]
        public long Id { get; set; }

        // 元器件所属分类目录名称 (如 "塑壳断路器")
        [JsonPropertyName("dirName")]
        public string DirName { get; set; } = string.Empty;

        // DWG 图纸文件名称 (如 "NM1-125S.dwg" 或 "NM1-125S")
        [JsonPropertyName("dwgName")]
        public string DwgName { get; set; } = string.Empty;

        // 规范化相对路径或全局唯一检索 Key (如 "塑壳断路器/NM1-125S.dwg")
        [JsonPropertyName("relPath")]
        public string RelPath { get; set; } = string.Empty;

        // 平面图形外形包围盒宽度 W (单位: mm)
        [JsonPropertyName("width")]
        public double Width { get; set; } = 0.0;

        // 平面图形外形包围盒高度 H (单位: mm)
        [JsonPropertyName("height")]
        public double Height { get; set; } = 0.0;

        // 从图纸文字中解析出的元器件安装进深/厚度/高度 Depth (单位: mm)
        [JsonPropertyName("depth")]
        public double Depth { get; set; } = 0.0;

        // 标记是否成功从文字中正则匹配提取到深度数值 (1:是, 0:否)
        [JsonPropertyName("hasTextDepth")]
        public bool HasTextDepth { get; set; } = false;

        // 匹配到的原始文字片段内容 (如 "高度: 85mm")
        [JsonPropertyName("matchedText")]
        public string MatchedText { get; set; } = string.Empty;

        // 磁盘物理文件的最后修改时间戳 (Utc Ticks, 用于毫秒级双重指纹热更新)
        [JsonPropertyName("lastModifiedTicks")]
        public long LastModifiedTicks { get; set; } = 0;

        // 磁盘物理文件字节大小 (Bytes, 用于双重指纹碰撞防护)
        [JsonPropertyName("fileSizeBytes")]
        public long FileSizeBytes { get; set; } = 0;

        // 记录最后更新写入时间字符串
        [JsonPropertyName("updatedAt")]
        public string UpdatedAt { get; set; } = string.Empty;
    }
}
