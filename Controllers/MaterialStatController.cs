using System;
using System.Collections.Generic;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// WebAPI 风格的材料统计业务控制器
    /// 承接前端 WebView2 消息调度，提供分类列表拉取与采购清单导出接口
    /// </summary>
    public class MaterialStatController
    {
        // 宿主窗体引用 (用于调度关闭窗口)
        private readonly System.Windows.Forms.Form? _ownerForm;

        /// <summary>
        /// 构造函数：注入宿主窗口引用
        /// </summary>
        /// <param name="ownerForm">宿主 WinForms 窗口对象</param>
        public MaterialStatController(System.Windows.Forms.Form? ownerForm = null)
        {
            // 保存宿主窗口引用
            _ownerForm = ownerForm;
        }

        /// <summary>
        /// WebAPI 接口: 获取材料统计初始化数据 (分类列表与已存在采购清单状态)
        /// </summary>
        /// <returns>材料统计初始化数据模型</returns>
        public MaterialStatInitDataDto GetInitData()
        {
            // 调用底层业务服务获取分类与已有采购清单状态
            return ExcelServices.GetMaterialStatInitData();
        }

        /// <summary>
        /// WebAPI 接口: 获取当前工程所有有效分类明细表及箱柜统计列表 (兼容旧接口)
        /// </summary>
        /// <returns>分类工作表信息列表</returns>
        public List<MaterialStatCategoryDto> GetCategories()
        {
            // 直接调用底层 ExcelServices 业务服务扫描分类表
            return ExcelServices.GetMaterialStatCategories();
        }

        /// <summary>
        /// WebAPI 接口: 根据用户勾选的分类工作表列表与排序规则，自动聚合采购总数量并生成采购清单
        /// </summary>
        /// <param name="selectedCategories">用户勾选的分类工作表名列表</param>
        /// <param name="sortBy">排序规则: brand_name_model (默认) 或 name_model</param>
        /// <returns>导出操作结果对象</returns>
        public PurchaseListExportResult ExportPurchaseList(List<string> selectedCategories, string sortBy = "brand_name_model")
        {
            // 采用 100% 静默模式执行导出 (绝不在底层调用模态 MessageBox.Show 阻塞 IPC)
            var result = ExcelServices.ExportPurchaseListToCurrentWorkbook(selectedCategories, sortBy, isSilent: true);

            // 导出结果直接原样返回，由前端展示成功 Toast 后再安全调度关闭窗口
            return result;
        }

        /// <summary>
        /// WebAPI 接口: 关闭当前向导窗口
        /// </summary>
        public void CloseWindow()
        {
            // 校验窗体有效性并执行安全关闭
            if (_ownerForm != null && !_ownerForm.IsDisposed)
            {
                _ownerForm.BeginInvoke(new Action(() =>
                {
                    _ownerForm.Close();
                }));
            }
        }
    }
}
