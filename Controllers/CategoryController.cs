using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// WebAPI 风格的分类管理控制器：提供分类建议探测、公式模板加载与分类创建接口
    /// </summary>
    public class CategoryController
    {
        // 引用公式调费控制器实例以复用公式组加载能力
        private readonly FormulaAdjustFeeController _feeController;

        /// <summary>
        /// 构造函数：初始化相关依赖控制器
        /// </summary>
        public CategoryController()
        {
            // 实例化公式调费控制器
            _feeController = new FormulaAdjustFeeController();
        }

        /// <summary>
        /// WebAPI 接口: 获取分类新建/编辑/插入复制引导与初始化配置数据
        /// </summary>
        /// <param name="mode">操作模式: create / edit / insertCopied</param>
        /// <returns>分类建议与公式组集合数据模型</returns>
        public CategorySuggestInfo GetCategorySuggestInfo(string mode = "create")
        {
            // 初始化返回实体
            var result = new CategorySuggestInfo();

            try
            {
                // 1. 调用 Excel 底层服务探测当前工作簿已存在的分类列表、激活分类与下一个推荐名称
                var excelSuggest = ExcelServices.GetSuggestedCategoryInfo(mode);
                if (excelSuggest != null)
                {
                    // 设置推荐分类名
                    result.SuggestedName = excelSuggest.SuggestedName;
                    // 设置已有分类表名称列表
                    result.ExistingCategories = excelSuggest.ExistingCategories ?? new List<string>();
                    // 设置当前处于激活聚焦状态的分类表
                    result.ActiveCategoryName = excelSuggest.ActiveCategoryName;
                    // 设置暂存的被复制分类表
                    result.CopiedCategoryName = excelSuggest.CopiedCategoryName;
                    // 设置当前模式
                    result.Mode = excelSuggest.Mode;
                }

                // 2. 加载系统所有调费公式组选项列表
                var groups = _feeController.GetFormulaGroups();
                if (groups != null && groups.Count > 0)
                {
                    // 将底层公式组模型映射为下拉选项传输对象
                    foreach (var g in groups)
                    {
                        result.FormulaGroups.Add(new CategoryFormulaGroupOption
                        {
                            Id = g.Id,
                            Name = g.Name,
                            IsDefault = g.IsDefault
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"获取分类建议与公式组信息异常: {ex.Message}");
            }

            // 返回组装完成的数据模型
            return result;
        }

        /// <summary>
        /// WebAPI 接口: 执行新建分类工作表核心业务
        /// </summary>
        /// <param name="request">新建分类前端表单提交请求</param>
        /// <returns>操作执行结果对象</returns>
        public CategoryOperationResult CreateCategory(CreateCategoryRequest request)
        {
            // 校验请求参数有效性
            if (request == null || string.IsNullOrWhiteSpace(request.CategoryName))
            {
                return new CategoryOperationResult
                {
                    Success = false,
                    Message = "分类名称不能为空"
                };
            }

            try
            {
                // 调度 Excel 业务服务层执行分类工作表创建与公式/定义名称绑定
                return ExcelServices.CreateNewCategory(request);
            }
            catch (Exception ex)
            {
                // 记录执行异常日志
                LogHelper.WriteLog($"创建分类工作表异常: {ex.Message}");
                return new CategoryOperationResult
                {
                    Success = false,
                    Message = $"创建分类发生错误: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// WebAPI 接口: 执行编辑/重命名分类工作表
        /// </summary>
        /// <param name="request">编辑分类表单提交请求</param>
        /// <returns>操作执行结果对象</returns>
        public CategoryOperationResult EditCategory(EditCategoryRequest request)
        {
            // 校验请求有效性
            if (request == null || string.IsNullOrWhiteSpace(request.OldCategoryName) || string.IsNullOrWhiteSpace(request.NewCategoryName))
            {
                return new CategoryOperationResult { Success = false, Message = "原分类名称与新分类名称均不能为空！" };
            }

            try
            {
                // 调度业务服务层重命名工作表并同步项目信息表
                return ExcelServices.EditCategoryName(request);
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"控制器编辑分类异常: {ex.Message}");
                return new CategoryOperationResult { Success = false, Message = $"编辑分类失败: {ex.Message}" };
            }
        }

        /// <summary>
        /// WebAPI 接口: 执行插入复制的分类工作表
        /// </summary>
        /// <param name="request">插入复制分类请求</param>
        /// <returns>操作执行结果对象</returns>
        public CategoryOperationResult InsertCopiedCategory(InsertCopiedCategoryRequest request)
        {
            // 校验入参
            if (request == null || string.IsNullOrWhiteSpace(request.TargetCategoryName))
            {
                return new CategoryOperationResult { Success = false, Message = "目标分类名称不能为空！" };
            }

            try
            {
                // 调度业务服务层克隆工作表并重新绑定定义名称与项目信息汇总
                return ExcelServices.InsertCopiedCategory(request);
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"控制器插入复制分类异常: {ex.Message}");
                return new CategoryOperationResult { Success = false, Message = $"插入复制分类失败: {ex.Message}" };
            }
        }

        /// <summary>
        /// WebAPI 接口: 执行删除指定分类工作表
        /// </summary>
        /// <param name="categoryName">分类工作表名称</param>
        /// <returns>操作执行结果对象</returns>
        public CategoryOperationResult DeleteCategory(string categoryName)
        {
            try
            {
                // 调度业务服务层安全删除分类及联动项目信息表行
                return ExcelServices.DeleteCategory(categoryName);
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"控制器删除分类异常: {ex.Message}");
                return new CategoryOperationResult { Success = false, Message = $"删除分类失败: {ex.Message}" };
            }
        }
    }
}
