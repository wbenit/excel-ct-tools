using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：新建项目与模板初始化 (对应 create_project.html)
    /// </summary>
    public static partial class ExcelServices
    {
        // 新建项目窗口静态单例引用 (可空)
        private static CreateProjectForm? _createProjectForm;

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“新建项目”窗口 (非模态，可编辑 Excel)
        /// </summary>
        public static void ShowCreateProjectDialog()
        {
            try
            {
                // 重置上一次创建的目标工作簿路径缓存
                Controllers.ProjectController.LastCreatedTargetFilePath = string.Empty;

                // 以非模态方式展示新建项目窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _createProjectForm, () => new CreateProjectForm());
            }
            catch (Exception ex)
            {
                // 全局捕获异常防止程序闪退
                System.Windows.Forms.MessageBox.Show($"弹出新建项目窗口失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 新建项目初始化工作簿：完整回填【项目信息】与【分类1】工作表的数据、公式联动与定义名称锚点
        /// </summary>
        /// <param name="newWb">新创建的目标工作簿 COM 对象</param>
        /// <param name="model">前端提交的新建项目表单数据模型</param>
        public static void InitializeCreatedProjectWorkbook(dynamic newWb, Controllers.CreateProjectModel model)
        {
            if (newWb == null || model == null) return;

            try
            {
                // 1. 读取配置文件中的基准行号与前缀定义
                int cabSumRow = ConfigManager.Instance.Current.Excel.CabSumRowIndex;
                int cabDetRow = ConfigManager.Instance.Current.Excel.CabDetRowIndex;
                int cabTolsumRow = ConfigManager.Instance.Current.Excel.CabTolsumRowIndex;

                // 读取 4 种名称前缀配置
                string sumPrefix = ConfigManager.Instance.Current.Excel.SumNamePrefix ?? "Cab_Sum_";
                string detPrefix = ConfigManager.Instance.Current.Excel.DetNamePrefix ?? "Cab_Det_";
                string subsumPrefix = ConfigManager.Instance.Current.Excel.SubsumNamePrefix ?? "Cab_Subsum_";
                string tolsumPrefix = ConfigManager.Instance.Current.Excel.TolsumNamePrefix ?? "Cab_Tolsum_";
                string defaultSheetName = ConfigManager.Instance.Current.Excel.DefaultTemplateSheet ?? "分类1";

                // 2. 填写【项目信息】工作表
                try
                {
                    dynamic infoSheet = newWb.Sheets["项目信息"];
                    if (infoSheet != null)
                    {
                        // 顶栏单位名称 (Row 1)
                        infoSheet.Range["B1"].Value = model.CompanyName;

                        // 【工程信息】区域填值 (Row 5 - Row 12)
                        infoSheet.Range["B5"].Value = model.ProjectName;      // 项目名称 (Cell B5)
                        infoSheet.Range["B6"].Value = model.ProjectRemark;    // 描述 (Cell B6)
                        infoSheet.Range["B7"].Value = model.QuoteNumber;      // 报价单号 (Cell B7)
                        infoSheet.Range["B8"].Value = model.Quoter;           // 报价人 (Cell B8)
                        infoSheet.Range["B9"].Value = model.ProjectDate;      // 创建日期 (Cell B9)
                        infoSheet.Range["B12"].Value = model.ProjectRemark;   // 项目备注 (Cell B12)

                        // 【客户信息】区域填值 (Row 14 - Row 17)
                        infoSheet.Range["B14"].Value = model.CustomerName;    // 客户名称 (Cell B14)
                        infoSheet.Range["B15"].Value = model.CustomerContact; // 联系人 (Cell B15)
                        infoSheet.Range["B16"].Value = model.CustomerPhone;   // 联系电话 (Cell B16)
                        infoSheet.Range["B17"].Value = model.CustomerAddress; // 客户地址 (Cell B17)

                        // 【本企业信息】区域填值 (Row 22 - Row 25)
                        infoSheet.Range["B22"].Value = model.CompanyName;    // 单位名称 (Cell B22)
                        infoSheet.Range["B23"].Value = model.EnglishName;     // 英文名称 (Cell B23)
                        infoSheet.Range["B24"].Value = model.CompanyContact;  // 联系人 (Cell B24)
                        infoSheet.Range["B25"].Value = model.CompanyPhone;    // 联系电话 (Cell B25)
                    }
                }
                catch (Exception exInfo)
                {
                    // 记录填写项目信息表的异常
                    LogHelper.WriteLog($"填写项目信息表异常: {exInfo.Message}");
                }

                // 3. 填写与初始化【分类1】工作表 (直接复用公共分类初始化逻辑)
                try
                {
                    dynamic catSheet = null;
                    try { catSheet = newWb.Sheets[defaultSheetName]; } catch { }
                    if (catSheet == null)
                    {
                        try { catSheet = newWb.Sheets["分类1"]; } catch { }
                    }

                    if (catSheet != null)
                    {
                        // 提取实际工作表名称
                        string actualCategoryName = Convert.ToString(catSheet.Name) ?? defaultSheetName;
                        // 先在【项目信息】表中登记分类汇总行 (Row 29)
                        UpdateProjectInfoCategorySummary(newWb, actualCategoryName);
                        // 调用公共通用分类初始化方法 (新建项目与新建分类共用)
                        InitializeCategorySheet(newWb, catSheet, 1, "箱柜1", "");
                    }
                }
                catch (Exception exCat)
                {
                    // 记录填写分类表的异常日志
                    LogHelper.WriteLog($"初始化分类1工作表异常: {exCat.Message}");
                }
            }
            catch (Exception ex)
            {
                // 记录全局初始化异常
                LogHelper.WriteLog($"初始化新项目工作簿异常: {ex.Message}");
            }
        }
    }
}
