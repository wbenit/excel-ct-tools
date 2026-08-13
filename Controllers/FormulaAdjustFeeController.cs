using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 公式组数据模型结构定义
    /// </summary>
    public class FormulaGroupModel
    {
        // 公式组唯一标识 ID
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        // 公式组名称 (例如: 简易费用公式、多费用公式、其它样式费用公式等)
        public string Name { get; set; } = string.Empty;

        // 是否为系统默认公式组 (图标显示钥匙标志)
        public bool IsSystemDefault { get; set; } = false;

        // 是否被用户设为当前激活的默认公式组
        public bool IsDefault { get; set; } = false;
    }

    /// <summary>
    /// 公式组包含的具体表格明细行数据模型
    /// </summary>
    public class FormulaItemModel
    {
        // 行序号 (如 1, 2, [序号])
        public string No { get; set; } = string.Empty;

        // 元件/项目名称 (如 小计, 管理费, 利润, 税金, 单台合计, 总计)
        public string Name { get; set; } = string.Empty;

        // 型号规格
        public string Model { get; set; } = string.Empty;

        // 生产厂家
        public string Manufacturer { get; set; } = string.Empty;

        // 单位 (如 台, 套)
        public string Unit { get; set; } = string.Empty;

        // 数量计算表达式或数值
        public string Quantity { get; set; } = string.Empty;

        // 单价计算表达式或数值
        public string Price { get; set; } = string.Empty;

        // 总价计算公式 (例如 =ROUND(H2*0.12, 2) 或 =ROUND(SUM(H2:H3)*0.15, 2))
        public string TotalPriceFormula { get; set; } = string.Empty;

        // 成本单价公式或数值
        public string CostPrice { get; set; } = string.Empty;

        // 成本总价公式 (例如 =ROUND(SUM(K2:K5), 2))
        public string CostTotalPriceFormula { get; set; } = string.Empty;

        // 项目类别 (如 费用)
        public string Category { get; set; } = string.Empty;
    }

    /// <summary>
    /// 调费请求参数实体
    /// </summary>
    public class ApplyFormulaRequest
    {
        // 目标调费范围: "currentCabinet"(当前箱柜), "currentCategory"(当前分类), "allCabinets"(所有箱柜), "selectedCabinet"(选择箱柜)
        public string TargetScope { get; set; } = "currentCabinet";

        // 当前选择的公式组名称
        public string GroupName { get; set; } = string.Empty;
    }

    /// <summary>
    /// 公式法调费 WebAPI 风格控制器，负责公式模板获取与调度处理
    /// </summary>
    public class FormulaAdjustFeeController
    {
        // 内存中保存的默认公式组列表 --硬编码初始数据，后续可从配置文件扩展--
        private static readonly List<FormulaGroupModel> _formulaGroups = new List<FormulaGroupModel>
        {
            // 预置简易费用公式组
            new FormulaGroupModel { Id = "1", Name = "简易费用公式", IsSystemDefault = true, IsDefault = false },
            // 预置多费用公式组 (图一默认选中)
            new FormulaGroupModel { Id = "2", Name = "多费用公式", IsSystemDefault = false, IsDefault = true },
            // 预置其它样式费用公式组
            new FormulaGroupModel { Id = "3", Name = "其它样式费用公式", IsSystemDefault = false, IsDefault = false },
            // 预置国网报价费用公式组
            new FormulaGroupModel { Id = "4", Name = "国网报价费用公式", IsSystemDefault = false, IsDefault = false },
            // 预置人工辅料定额公式组
            new FormulaGroupModel { Id = "5", Name = "人工辅料定额公式", IsSystemDefault = false, IsDefault = false }
        };

        /// <summary>
        /// 后端 WebAPI 接口: 获取所有公式组列表
        /// </summary>
        /// <returns>公式组集合</returns>
        public List<FormulaGroupModel> GetFormulaGroups()
        {
            // 直接返回已配置的公式组数据列表
            return _formulaGroups;
        }

        /// <summary>
        /// 后端 WebAPI 接口: 设置指定公式组为默认公式组
        /// </summary>
        /// <param name="groupId">公式组 ID</param>
        /// <returns>操作是否成功</returns>
        public bool SetDefaultFormulaGroup(string groupId)
        {
            // 遍历所有公式组，重置 IsDefault 状态
            foreach (var g in _formulaGroups)
            {
                // 若 ID 匹配则设为默认，否则置为 false
                g.IsDefault = (g.Id == groupId);
            }
            // 返回设为默认成功的状态
            return true;
        }

        /// <summary>
        /// 后端 WebAPI 接口: 复制指定的公式组
        /// </summary>
        /// <param name="groupId">源公式组 ID</param>
        /// <returns>新建的公式组对象</returns>
        public FormulaGroupModel CopyFormulaGroup(string groupId)
        {
            // 查找对应的源公式组
            var source = _formulaGroups.FirstOrDefault(g => g.Id == groupId);
            // 若源不存在则返回空
            if (source == null) return null!;

            // 构造新的复制实体 --硬编码名称后缀--
            var newGroup = new FormulaGroupModel
            {
                // 生成新唯一标识
                Id = Guid.NewGuid().ToString("N"),
                // 名称加上副本后缀
                Name = $"{source.Name}_副本",
                // 新副本非系统默认
                IsSystemDefault = false,
                // 非默认激活
                IsDefault = false
            };

            // 追加至全局内存集合中
            _formulaGroups.Add(newGroup);
            // 返回新创建的公式组对象
            return newGroup;
        }

        /// <summary>
        /// 后端 WebAPI 接口: 删除指定的公式组
        /// </summary>
        /// <param name="groupId">要删除的公式组 ID</param>
        /// <returns>操作结果</returns>
        public bool DeleteFormulaGroup(string groupId)
        {
            // 查找对应公式组
            var target = _formulaGroups.FirstOrDefault(g => g.Id == groupId);
            // 系统默认公式组不允许删除
            if (target == null || target.IsSystemDefault)
            {
                // 返回删除失败
                return false;
            }

            // 从列表中移除目标公式组
            _formulaGroups.Remove(target);
            // 返回删除成功标志
            return true;
        }

        /// <summary>
        /// 后端 WebAPI 接口: 根据公式组名称获取对应的明细计算公式表
        /// </summary>
        /// <param name="groupName">公式组名称</param>
        /// <returns>明细列表行数据集合</returns>
        public List<FormulaItemModel> GetFormulaDetails(string groupName)
        {
            // 初始化返回的明细行列表
            var items = new List<FormulaItemModel>();

            // 图一中默认显示的“多费用公式”明细表 --硬编码示例公式明细，符合标准规则--
            items.Add(new FormulaItemModel
            {
                No = "[序号]",
                Name = "小计",
                Model = "",
                Manufacturer = "",
                Unit = "",
                Quantity = "",
                Price = "",
                TotalPriceFormula = "[总价小计]",
                CostPrice = "",
                CostTotalPriceFormula = "[成本总价小计]",
                Category = ""
            });

            // 添加管理费计算公式行
            items.Add(new FormulaItemModel
            {
                No = "[序号]",
                Name = "管理费",
                Model = "",
                Manufacturer = "",
                Unit = "",
                Quantity = "",
                Price = "",
                TotalPriceFormula = "=ROUND(H2*0.12, 2)",
                CostPrice = "",
                CostTotalPriceFormula = "",
                Category = "费用"
            });

            // 添加利润计算公式行
            items.Add(new FormulaItemModel
            {
                No = "[序号]",
                Name = "利润",
                Model = "",
                Manufacturer = "",
                Unit = "",
                Quantity = "",
                Price = "",
                TotalPriceFormula = "=ROUND(SUM(H2:H3)*0.15, 2)",
                CostPrice = "",
                CostTotalPriceFormula = "",
                Category = "费用"
            });

            // 添加税金计算公式行
            items.Add(new FormulaItemModel
            {
                No = "[序号]",
                Name = "税金",
                Model = "",
                Manufacturer = "",
                Unit = "",
                Quantity = "",
                Price = "",
                TotalPriceFormula = "=ROUND(SUM(H2:H4)*0.13, 2)",
                CostPrice = "",
                CostTotalPriceFormula = "",
                Category = "费用"
            });

            // 添加单台合计计算公式行
            items.Add(new FormulaItemModel
            {
                No = "[序号]",
                Name = "单台合计",
                Model = "",
                Manufacturer = "",
                Unit = "",
                Quantity = "",
                Price = "",
                TotalPriceFormula = "=ROUND(SUM(H2:H5), 2)",
                CostPrice = "",
                CostTotalPriceFormula = "=ROUND(SUM(K2:K5), 2)",
                Category = ""
            });

            // 添加总计行公式
            items.Add(new FormulaItemModel
            {
                No = "7",
                Name = "总计",
                Model = "",
                Manufacturer = "",
                Unit = "台",
                Quantity = "=ROUND(H6, 2)",
                Price = "=ROUND(F7*G7, 2)",
                TotalPriceFormula = "",
                CostPrice = "",
                CostTotalPriceFormula = "=ROUND(K6*F7, 2)",
                Category = ""
            });

            // 返回公式明细表集合
            return items;
        }
    }
}
