using System;
using System.Collections.Generic;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 报价智能核验业务控制器
    /// 负责衔接前端 Vue 3 界面与 ExcelServices 核心服务层之间的通信调度
    /// 严格遵循规则 3: 具体 Excel 读写操作统一交由 ExcelServices 实现
    /// </summary>
    public class QuotationCheckController
    {
        // 声明窗体宿主弱引用
        private readonly Forms.QuotationCheckForm _form;

        /// <summary>
        /// 构造函数: 注入窗体宿主对象
        /// </summary>
        /// <param name="form">报价核验窗体实例</param>
        public QuotationCheckController(Forms.QuotationCheckForm form)
        {
            // 保存窗体宿主弱引用
            _form = form;
        }

        /// <summary>
        /// 执行报价智能体检，全盘扫描各分类表中的价格、折扣、公式损毁及 AA/AB 列参数
        /// </summary>
        /// <returns>全量体检诊断结果</returns>
        public QuotationCheckResult RunAudit()
        {
            // 委托 ExcelServices 执行全量报价体检 (规则 3)
            return ExcelServices.ExecuteQuotationAudit();
        }

        /// <summary>
        /// 一键批量修复选中的或全部被破坏的计算公式
        /// </summary>
        /// <param name="issueIds">需修复的问题 ID 清单</param>
        /// <returns>修复操作结果</returns>
        public (bool Success, string Message, int FixedCount) FixFormulas(List<string>? issueIds = null)
        {
            // 委托 ExcelServices 执行公式修复
            return ExcelServices.FixQuotationFormulas(issueIds);
        }

        /// <summary>
        /// 同型号价格批量对齐广播服务
        /// </summary>
        /// <param name="req">对齐配置参数</param>
        /// <returns>对齐结果与受影响行数</returns>
        public (bool Success, string Message, int UpdatedCount) AlignModelPrice(AlignModelPriceRequest req)
        {
            // 委托 ExcelServices 执行同型号全表对齐
            return ExcelServices.AlignModelPriceAcrossSheets(req);
        }

        /// <summary>
        /// 一键清洗文本型数字
        /// </summary>
        /// <returns>清洗结果与数量</returns>
        public (bool Success, string Message, int CleanedCount) CleanTextNumbers()
        {
            // 委托 ExcelServices 执行数字格式清洗
            return ExcelServices.CleanQuotationTextNumbers();
        }

        /// <summary>
        /// 在 Excel 中高亮定位并滚动显示指定单元格
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="row">物理行号</param>
        /// <param name="colLetter">列标字符</param>
        /// <returns>是否成功激活定位</returns>
        public bool NavigateToCell(string sheetName, int row, string colLetter)
        {
            // 委托 ExcelServices 执行视口定位
            return ExcelServices.NavigateToQuotationCell(sheetName, row, colLetter);
        }
    }
}
