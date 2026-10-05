using System;
using System.Collections.Generic;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 动态属性依附锚点模型
    /// 遵循双基线动态依附法则：在表头上方则依附表头，在结束行下方则依附结束行
    /// </summary>
    public class DynamicFieldAnchor
    {
        // 目标所在的列号 (1-based，例如 B列为2，E列为5)
        public int TargetCol { get; set; } = 1;

        // 是否依附于表头行 (true:依附表头行; false:依附元器件结束行)
        public bool AttachToHeader { get; set; } = true;

        // 相对位移偏移量 (依附表头时通常<=0，依附结束行时通常>=0)
        public int RowOffset { get; set; } = 0;

        // 绑定的原始文本前缀或标签 (如"柜号："，用于前缀清洗剥离)
        public string PrefixKeyword { get; set; } = string.Empty;

        /// <summary>
        /// 根据用户在第一柜点选的坐标与双基准行，自动建立依附关系
        /// </summary>
        /// <param name="clickedRow">用户点击的物理行号</param>
        /// <param name="clickedCol">用户点击的物理列号</param>
        /// <param name="headerRow">用户确认的表头行号</param>
        /// <param name="endRow">用户确认的元器件结束行号</param>
        /// <param name="rawText">单元格原始文本</param>
        public void Bind(int clickedRow, int clickedCol, int headerRow, int endRow, string rawText = "")
        {
            // 记录目标列索引
            TargetCol = clickedCol;
            // 判断点击行与表头行的相对位置
            if (clickedRow <= headerRow)
            {
                // 点击位置在表头或表头上方：依附于表头行
                AttachToHeader = true;
                // 计算相对位移值 (通常为负数或0)
                RowOffset = clickedRow - headerRow;
            }
            else
            {
                // 点击位置在表头下方(或结束行及下方)：依附于元器件结束行
                AttachToHeader = false;
                // 计算相对位移值 (通常为正数或0)
                RowOffset = clickedRow - endRow;
            }

            // 若包含冒号，自动提取冒号前缀作为清洗关键词
            if (!string.IsNullOrEmpty(rawText) && (rawText.Contains("：") || rawText.Contains(":")))
            {
                // 拆分冒号获取前缀
                var parts = rawText.Split(new char[] { '：', ':' }, 2);
                // 记录前缀标签
                if (parts.Length > 0) PrefixKeyword = parts[0].Trim();
            }
        }

        /// <summary>
        /// 动态计算任意新箱柜的目标绝对物理行号
        /// </summary>
        /// <param name="newHeaderRow">新箱柜的表头行号</param>
        /// <param name="newEndRow">新箱柜的元器件结束行号</param>
        /// <returns>计算得出的目标物理行号</returns>
        public int CalculateRow(int newHeaderRow, int newEndRow)
        {
            // 根据依附关系分别叠加偏移量
            return AttachToHeader ? (newHeaderRow + RowOffset) : (newEndRow + RowOffset);
        }
    }

    /// <summary>
    /// 扩展列映射模型：支持原表任意列映射写入导入表任意列
    /// </summary>
    public class ExtraColumnMapping
    {
        // 原表源列号 (1-based，0表示未启用不导入)
        public int SourceCol { get; set; } = 0;
        // 导入表目标列号 (1-based，例如 5 对应 E 列单位，9 对应 I 列备注，15 对应 O 列等)
        public int TargetCol { get; set; } = 0;
        // 扩展列标签/自定义名称 (如 "单位", "备注", "技术参数")
        public string CustomLabel { get; set; } = string.Empty;
        // 自动识别关键字 (逗号分隔，用于表头启发式嗅探)
        public string Keywords { get; set; } = string.Empty;
    }

    /// <summary>
    /// 智能导入列映射配置字典
    /// </summary>
    public class SmartImportColumnMapping
    {
        // 序号列号 (1-based，0表示未指定)
        public int IndexCol { get; set; } = 0;
        // 元器件名称列号 (必选)
        public int ItemNameCol { get; set; } = 2;
        // 规格型号列号 (必选)
        public int ItemSpecCol { get; set; } = 3;
        // 数量列号 (必选)
        public int QuantityCol { get; set; } = 5;
        // 单位列号 (可选)
        public int UnitCol { get; set; } = 4;
        // 单价/表价格列号 (可选，写入表价)
        public int PriceCol { get; set; } = 6;
        // 生产厂家/品牌列号 (可选)
        public int BrandCol { get; set; } = 8;
        // 备注列号 (可选)
        public int RemarkCol { get; set; } = 9;

        // 5 列自定义匹配关键字配置 (逗号分隔，可由用户自定义修改并持久化保存)
        // 元件名称匹配关键字词典
        public string ItemNameKeywords { get; set; } = "名称,品名";
        // 规格型号匹配关键字词典
        public string ItemSpecKeywords { get; set; } = "型号,规格";
        // 数量匹配关键字词典
        public string QuantityKeywords { get; set; } = "数量";
        // 价格/单价匹配关键字词典
        public string PriceKeywords { get; set; } = "单价,面价";
        // 品牌/厂家匹配关键字词典
        public string BrandKeywords { get; set; } = "品牌,厂家";

        // 方案 B：通用任意扩展列 1 (默认目标列为 5，即 E 列计量单位)
        public ExtraColumnMapping ExtraCol1 { get; set; } = new ExtraColumnMapping
        {
            // 默认源列未指定
            SourceCol = 0,
            // 目标列预设为第 5 列 (E 列计量单位)
            TargetCol = 5,
            // 自定义标签名称
            CustomLabel = "扩展列1",
            // 预设嗅探关键字
            Keywords = "单位,计量单位"
        };

        // 方案 B：通用任意扩展列 2 (默认目标列为 9，即 I 列备注说明)
        public ExtraColumnMapping ExtraCol2 { get; set; } = new ExtraColumnMapping
        {
            // 默认源列未指定
            SourceCol = 0,
            // 目标列预设为第 9 列 (I 列备注)
            TargetCol = 9,
            // 自定义标签名称
            CustomLabel = "扩展列2",
            // 预设嗅探关键字
            Keywords = "备注,说明,参数,要求,位号"
        };
    }

    /// <summary>
    /// 智能导入规则模板配置
    /// 记录用户在样本柜上点选确立的完整规则
    /// </summary>
    public class SmartImportTemplateConfig
    {
        // 模板名称 (如"天正成套报价清单模板")
        public string TemplateName { get; set; } = string.Empty;
        // 样本表头行号 (1-based)
        public int SampleHeaderRow { get; set; } = 2;
        // 样本元器件结束行号 (1-based，如小计所在行)
        public int SampleEndRow { get; set; } = 10;
        // 结束行标志关键词 (如"小计"、"合计")
        public string EndKeyword { get; set; } = "小计";
        // 柜号依附锚点
        public DynamicFieldAnchor CabinetAnchor { get; set; } = new DynamicFieldAnchor();
        // 台数依附锚点
        public DynamicFieldAnchor QuantityAnchor { get; set; } = new DynamicFieldAnchor();
        // 列映射关系
        public SmartImportColumnMapping ColumnMapping { get; set; } = new SmartImportColumnMapping();
    }

    /// <summary>
    /// 解析出的元器件条目模型
    /// </summary>
    public class ParsedComponentModel
    {
        // 序号
        public int Index { get; set; } = 0;
        // 元器件名称 (品名)
        public string ItemName { get; set; } = string.Empty;
        // 规格型号
        public string ItemSpec { get; set; } = string.Empty;
        // 数量
        public decimal Quantity { get; set; } = 1;
        // 单位 (台/只/套)
        public string Unit { get; set; } = "台";
        // 表价/面价 (来自外部单价)
        public decimal MarkedPrice { get; set; } = 0;
        // 生产厂家/推荐品牌
        public string Brand { get; set; } = string.Empty;
        // 备注说明
        public string Remark { get; set; } = string.Empty;
        // 方案 B：扩展列 1 提取的文本值
        public string ExtraValue1 { get; set; } = string.Empty;
        // 方案 B：扩展列 2 提取的文本值
        public string ExtraValue2 { get; set; } = string.Empty;
    }

    /// <summary>
    /// 解析出的单个箱柜模型
    /// </summary>
    public class ParsedCabinetModel
    {
        // 箱柜在源表中的物理起始行
        public int SourceStartRow { get; set; } = 0;
        // 箱柜在源表中的物理结束行
        public int SourceEndRow { get; set; } = 0;
        // 箱柜编号/名称 (如 QSAC1)
        public string CabinetName { get; set; } = string.Empty;
        // 箱柜台数 (如 2台)
        public decimal CabinetQuantity { get; set; } = 1;
        // 该箱柜包含的元器件明细集合
        public List<ParsedComponentModel> Components { get; set; } = new List<ParsedComponentModel>();
        // 表价总计预估 (便于前端卡片核对展示)
        public decimal TotalMarkedPrice { get; set; } = 0;
    }

    /// <summary>
    /// 外部 Excel 前端预览网格数据包
    /// </summary>
    public class SmartImportPreviewResult
    {
        // 是否读取成功
        public bool Success { get; set; } = true;
        // 错误提示信息
        public string Message { get; set; } = string.Empty;
        // 文件内所有工作表名称集合
        public List<string> SheetNames { get; set; } = new List<string>();
        // 当前选中的工作表名称
        public string CurrentSheet { get; set; } = string.Empty;
        // 网格行数
        public int RowCount { get; set; } = 0;
        // 网格列数
        public int ColCount { get; set; } = 0;
        // 二维单元格文本矩阵 (行优先，[row][col])
        public List<List<string>> GridData { get; set; } = new List<List<string>>();

        // 内部全量二维原始矩阵缓存 (不参与 JSON 序列化传输)
        [System.Text.Json.Serialization.JsonIgnore]
        public object[,]? FullMatrix { get; set; } = null;
        // 内部全量行数缓存
        [System.Text.Json.Serialization.JsonIgnore]
        public int FullRowCount { get; set; } = 0;
        // 内部全量列数缓存
        [System.Text.Json.Serialization.JsonIgnore]
        public int FullColCount { get; set; } = 0;
    }

    /// <summary>
    /// 执行智能导入的统一请求参数
    /// </summary>
    public class SmartImportExecuteRequest
    {
        // 外部 Excel 文件物理完整路径
        public string FilePath { get; set; } = string.Empty;
        // 选中的工作表名称
        public string SheetName { get; set; } = string.Empty;
        // 新建的分类表工作表名称
        public string NewCategoryName { get; set; } = "导入_箱柜清单";
        // 解析规则配置
        public SmartImportTemplateConfig Config { get; set; } = new SmartImportTemplateConfig();
        // 用户选中的待导入箱柜列表 (如果为空则全部导入)
        public List<ParsedCabinetModel> Cabinets { get; set; } = new List<ParsedCabinetModel>();
    }

    /// <summary>
    /// 智能导入执行结果实体
    /// </summary>
    public class SmartImportExecuteResult
    {
        // 是否执行成功
        public bool Success { get; set; } = false;
        // 状态消息提示
        public string Message { get; set; } = string.Empty;
        // 新建的分类表名称
        public string CreatedCategoryName { get; set; } = string.Empty;
        // 导入成功的箱柜总数
        public int ImportedCabinetCount { get; set; } = 0;
        // 导入成功的元器件总数
        public int ImportedComponentCount { get; set; } = 0;
    }
}
