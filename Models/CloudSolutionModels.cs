using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 云方案中心常量与默认配置
    /// </summary>
    public static class CloudSchemeDefaults
    {
        // 方案默认每页展示记录数 (对应 4 列 * 3 行)
        public const int DefaultPageSize = 12; // --硬编码-- 每页卡片数量

        // 默认主分类：低压进线方案
        public const string DefaultCategoryName = "低压成套方案"; // --硬编码-- 默认分类名称

        // 默认单价与合价小数位数精度
        public const int PriceDecimalPlaces = 2; // --硬编码-- 金额小数精度
    }

    /// <summary>
    /// 方案归属范围枚举常量 (行业方案 / 企业方案)
    /// </summary>
    public static class CloudSchemeScope
    {
        // 行业方案：系统管理员统管，只读标准库
        public const string Industry = "industry"; // --硬编码-- 行业方案标识

        // 企业方案：企业组织及用户自建维护
        public const string Enterprise = "enterprise"; // --硬编码-- 企业方案标识
    }

    /// <summary>
    /// 方案业务分类枚举常量 (一次方案 / 二次方案 / 我的收藏)
    /// </summary>
    public static class CloudSchemeType
    {
        // 一次方案：成套主回路柜体与配电箱
        public const string Primary = "primary"; // --硬编码-- 一次方案标识

        // 二次方案：二次控制回路与测控系统
        public const string Secondary = "secondary"; // --硬编码-- 二次方案标识

        // 我的收藏：用户星标标记的方案集
        public const string Favorite = "favorite"; // --硬编码-- 收藏标识
    }

    /// <summary>
    /// 云方案基础卡片摘要数据实体模型
    /// </summary>
    public class CloudSchemeItem
    {
        // 方案唯一主键标识 ID
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        // 方案归属范围 ("industry": 行业方案, "enterprise": 企业方案)
        public string Scope { get; set; } = CloudSchemeScope.Industry;

        // 方案业务分类 ("primary": 一次方案, "secondary": 二次方案)
        public string SchemeType { get; set; } = CloudSchemeType.Primary;

        // 方案完整显示名称 (如: "KYN28-12 PT柜(630A)")
        public string SchemeName { get; set; } = string.Empty;

        // 柜型型号 (如: "KYN28-12", "MNS", "GGD")
        public string CabinetModel { get; set; } = string.Empty;

        // 外形尺寸: 宽*深*高 mm (如: "800*1500*2200")
        public string Dimensions { get; set; } = string.Empty;

        // 封面缩略图路径 (空则默认提取 DrawingUrls[0])
        public string ThumbnailUrl { get; set; } = string.Empty;

        // 关联多张图纸与高清渲染图片集合 (支持 DWG 矢量图、PNG、JPG、PDF)
        public List<string> DrawingUrls { get; set; } = new List<string>();

        // 方案详细技术描述与使用工况说明
        public string Description { get; set; } = string.Empty;

        // 所属文件夹/分类节点 ID
        public string CategoryId { get; set; } = string.Empty;

        // 所属文件夹/分类节点名称
        public string CategoryName { get; set; } = string.Empty;

        // 方案浏览量统计数值
        public int ViewCount { get; set; } = 0;

        // 方案收藏量统计数值
        public int FavCount { get; set; } = 0;

        // 当前登录用户是否已收藏该方案
        public bool IsFavorite { get; set; } = false;

        // 方案作者或维护人名称
        public string CreatorName { get; set; } = string.Empty;

        // 方案最新更新时间 (yyyy-MM-dd HH:mm)
        public string UpdateTime { get; set; } = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

        // 特殊推荐标签 (如: "顶", "精", "企业")
        public string Tag { get; set; } = string.Empty;
    }

    /// <summary>
    /// 方案 BOM 元器件清单明细项实体
    /// </summary>
    public class CloudSchemeBomItem
    {
        // BOM 条目唯一标识 ID
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        // 关联的方案主键 ID
        public string SchemeId { get; set; } = string.Empty;

        // 行排列展示序号
        public int SortOrder { get; set; } = 1;

        // 元器件名称 (如: "PT手车", "电压互感器", "避雷器")
        public string ComponentName { get; set; } = string.Empty;

        // 元器件规格型号 (如: "JDZX10-10", "XRNP-10/0.5A")
        public string ModelSpec { get; set; } = string.Empty;

        // 品牌厂家 (如: "大连第一互感器", "上海一开")
        public string Brand { get; set; } = string.Empty;

        // 计量单位 (如: "台", "套", "只", "米")
        public string Unit { get; set; } = "台"; // --硬编码-- 默认计量单位

        // 基础单回路额定用量
        public double Quantity { get; set; } = 1.0;

        // 是否属于随回路数翻倍元器件 (WL 标志: 勾选则回路数倍增，不勾选保持基准数量)
        public bool IsWlDoubled { get; set; } = false;

        // 目录官方表价 (元)
        public decimal CatalogPrice { get; set; } = 0m;

        // 进货/成本折扣率 (默认 1.0)
        public decimal Discount { get; set; } = 1.0m;

        // 报出折扣率 (默认 1.0)
        public decimal QuoteDiscount { get; set; } = 1.0m;

        // 最终报价单价 (元)
        public decimal QuotePrice { get; set; } = 0m;

        // 该行元器件总价 (数量 * 单价)
        public decimal TotalPrice { get; set; } = 0m;

        // 备注说明文本
        public string Remark { get; set; } = string.Empty;

        // ERP/PLM 物料编码
        public string MaterialCode { get; set; } = string.Empty;

        // 生产产地
        public string Origin { get; set; } = string.Empty;

        // 元器件所属分类或元器件类别
        public string Category { get; set; } = string.Empty;

        // 前端表格勾选选中状态 (是否选中导出或插入)
        public bool Selected { get; set; } = true;
    }

    /// <summary>
    /// 方案完整详情实体 (包含主信息与元器件 BOM 清单)
    /// </summary>
    public class CloudSchemeDetail : CloudSchemeItem
    {
        // 方案包含的全部元器件 BOM 集合
        public List<CloudSchemeBomItem> BomItems { get; set; } = new List<CloudSchemeBomItem>();

        // 方案标准总报价金额 (自动根据 BomItems 合价求和)
        public decimal TotalAmount { get; set; } = 0m;
    }

    /// <summary>
    /// 企业方案左侧自定义树形分类节点实体
    /// </summary>
    public class CloudSchemeCategoryNode
    {
        // 分类节点主键 ID
        public string Id { get; set; } = string.Empty;

        // 分类名称 (如: "高压进线方案", "电容补偿柜")
        public string Name { get; set; } = string.Empty;

        // 父级分类 ID (顶级为空)
        public string ParentId { get; set; } = string.Empty;

        // 图标标识名称
        public string Icon { get; set; } = "Folder"; // --硬编码-- 默认文件夹图标

        // 子分类节点集合
        public List<CloudSchemeCategoryNode> Children { get; set; } = new List<CloudSchemeCategoryNode>();
    }

    /// <summary>
    /// 云方案分页多维检索查询过滤参数载体 DTO
    /// </summary>
    public class CloudSchemeQueryDto
    {
        // 方案范围 ("industry": 行业方案, "enterprise": 企业方案)
        public string Scope { get; set; } = CloudSchemeScope.Industry;

        // 方案类型 ("primary": 一次方案, "secondary": 二次方案, "favorite": 我的收藏)
        public string SchemeType { get; set; } = CloudSchemeType.Primary;

        // 选中的分类目录 ID (为空表示全部)
        public string CategoryId { get; set; } = string.Empty;

        // 检索关键字 (方案名、柜型、标签或作者)
        public string Keyword { get; set; } = string.Empty;

        // 柜型过滤 (如: "KYN28-12", "MNS")
        public string CabinetModel { get; set; } = string.Empty;

        // 排序规则 ("hot": 按浏览收藏热度, "latest": 按最新更新时间)
        public string SortBy { get; set; } = "hot"; // --硬编码-- 默认热门排序

        // 分页页码 (从 1 开始)
        public int PageIndex { get; set; } = 1;

        // 每页大小 (默认 12 项)
        public int PageSize { get; set; } = CloudSchemeDefaults.DefaultPageSize;
    }

    /// <summary>
    /// 云方案分页查询返回结果载体
    /// </summary>
    public class CloudSchemePageResult
    {
        // 当前页方案卡片数据列表
        public List<CloudSchemeItem> Items { get; set; } = new List<CloudSchemeItem>();

        // 满足过滤条件的总记录数
        public int TotalCount { get; set; } = 0;

        // 当前页码
        public int PageIndex { get; set; } = 1;

        // 每页大小
        public int PageSize { get; set; } = CloudSchemeDefaults.DefaultPageSize;
    }

    /// <summary>
    /// 方案一键插入到 Excel 的请求参数载体 DTO
    /// </summary>
    public class SchemeInsertToExcelDto
    {
        // 方案主键 ID
        public string SchemeId { get; set; } = string.Empty;

        // 回路数倍增器数值 (默认 1 个回路)
        public int LoopMultiplier { get; set; } = 1;

        // 插入目标方式 ("currentCabinet": 写入当前选中的箱柜元器件区, "newCabinet": 自动新建并追加箱柜)
        public string InsertMode { get; set; } = "newCabinet"; // --硬编码-- 默认插入为新箱柜

        // 用户勾选确认的 BOM 明细列表
        public List<CloudSchemeBomItem> SelectedBomItems { get; set; } = new List<CloudSchemeBomItem>();
    }

    /// <summary>
    /// 二次方案 DWG 根目录下的直接子文件夹项 DTO
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class SecondaryFolderItemDto
    {
        // 子文件夹显示名称 (如: "双电源", "时控开关")
        public string Name { get; set; } = string.Empty;

        // 子文件夹物理绝对路径
        public string FullPath { get; set; } = string.Empty;

        // 该子文件夹下包含的 DWG 图纸数量
        public int DwgCount { get; set; } = 0;
    }

    /// <summary>
    /// 云方案二次回路 DWG 图纸与数据库参数一体化卡片展示 DTO
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class SecondaryFolderDwgCardDto
    {
        // 完整文件名称 (带后缀，如: "WATSG.dwg")
        public string FileName { get; set; } = string.Empty;

        // 去扩展名的图名代号 (如: "WATSG")
        public string DwgName { get; set; } = string.Empty;

        // 方案大标题显示名称 (如: "双电源自动切换控制原理图")
        public string SchemeName { get; set; } = string.Empty;

        // 所属方案目录名称 (左侧选中的子文件夹名，如: "双电源")
        public string FolderName { get; set; } = string.Empty;

        // 所属子文件夹物理全路径
        public string FolderPath { get; set; } = string.Empty;

        // DWG 文件物理绝对全路径
        public string FullPath { get; set; } = string.Empty;

        // 方案品牌 (如: "施耐德")
        public string Brand { get; set; } = string.Empty;

        // 方案工艺与技术描述
        public string Description { get; set; } = string.Empty;

        // 二次线跨门根数
        public double CrossDoorCount { get; set; } = 0.0;

        // 开孔要求 (如: "圆孔 2个")
        public string HoleSpec { get; set; } = string.Empty;

        // 装配人工费用 (单位: 元)
        public double LaborCost { get; set; } = 0.0;

        // 二次材料费用小计 (实时从子 BOM 汇总，单位: 元)
        public double MaterialCost { get; set; } = 0.0;

        // 综合总成本 (人工 + 材料)
        public double TotalCost { get; set; } = 0.0;

        // 二次排布图 DWG (对应 cad_drawing_name，如: "FA")
        public string LayoutDwgName { get; set; } = string.Empty;

        // DWG 缩略图 Base64 图像流
        public string PreviewBase64 { get; set; } = string.Empty;

        // 是否在 personal_components.db 中精准匹配命中
        public bool IsMatched { get; set; } = false;

        // 命中的数据库主键 ID (未命中为 0)
        public int SchemeId { get; set; } = 0;

        // 方案最近更新时间文本
        public string UpdateTime { get; set; } = string.Empty;

        // 关联的完整二次方案实体 (供编辑弹窗直接反显)
        public SecondarySchemeEntity? SchemeData { get; set; }
    }
}
