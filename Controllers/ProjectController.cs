using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 新建项目初始化请求数据实体
    /// </summary>
    public class CreateProjectModel
    {
        // 项目名称
        public string ProjectName { get; set; } = "新建项目";

        // 报价单号 (例如 WB202608090130)
        public string QuoteNumber { get; set; } = string.Empty;

        // 项目备注
        public string ProjectRemark { get; set; } = string.Empty;

        // 项目编号
        public string ProjectCode { get; set; } = string.Empty;

        // 文件名称
        public string FileName { get; set; } = string.Empty;

        // 保存路径 (例如 C:\Users\xxx\Desktop)
        public string SavePath { get; set; } = string.Empty;

        // 项目类型 (常规项目 / 国网项目)
        public string ProjectType { get; set; } = "常规项目";

        // 是否批建箱柜
        public bool BatchCabinet { get; set; } = false;

        // 企业单位名称
        public string CompanyName { get; set; } = string.Empty;

        // 企业英文名称
        public string EnglishName { get; set; } = string.Empty;

        // 报价人
        public string Quoter { get; set; } = string.Empty;

        // 企业联系人
        public string CompanyContact { get; set; } = string.Empty;

        // 企业联系电话
        public string CompanyPhone { get; set; } = string.Empty;

        // 客户名称
        public string CustomerName { get; set; } = string.Empty;

        // 客户地址
        public string CustomerAddress { get; set; } = string.Empty;

        // 客户联系人
        public string CustomerContact { get; set; } = string.Empty;

        // 客户联系电话
        public string CustomerPhone { get; set; } = string.Empty;

        // 项目日期
        public string ProjectDate { get; set; } = DateTime.Now.ToString("yyyy年MM月dd日");
    }

    /// <summary>
    /// 新建项目后端业务逻辑控制器
    /// </summary>
    public class ProjectController
    {
        // 记录最近一次成功创建的项目工作簿绝对路径，用于弹窗关闭后强力激活焦点
        public static string LastCreatedTargetFilePath { get; set; } = string.Empty;

        /// <summary>
        /// 自动生成包含时间戳与随机序号的报价单号
        /// </summary>
        public string GenerateQuoteNumber()
        {
            // 获取当前系统时间的年月日字符串
            string dateStr = DateTime.Now.ToString("yyyyMMdd");

            // 生成 4 位随机流水序号
            int seq = new Random().Next(1000, 9999);

            // 拼接 WB + 年月日 + 4位序号
            return $"WB{dateStr}{seq}";
        }

        /// <summary>
        /// 获取当前用户桌面的默认物理路径
        /// </summary>
        public string GetDefaultDesktopPath()
        {
            // 获取系统 SpecialFolder.Desktop 桌面绝对路径
            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }

        /// <summary>
        /// 异步初始化并创建新项目文件，复制 CabinetTemplate.xlsx 及其工作表并填入项目信息
        /// </summary>
        public async Task<bool> CreateProjectAsync(CreateProjectModel model)
        {
            // 在 Task 异步线程中完成 Workbook 创建与单元格数据填入
            return await Task.Run(() =>
            {
                try
                {
                    // 1. 校验保存路径有效性，若不存在则自动创建
                    if (!Directory.Exists(model.SavePath))
                    {
                        // 自动创建保存目录
                        Directory.CreateDirectory(model.SavePath);
                    }

                    // 2. 确定目标工作簿文件名与路径
                    string fileName = string.IsNullOrWhiteSpace(model.FileName) ? "新建项目" : model.FileName;
                    // 若文件名未包含 .xlsx 后缀则自动补全
                    if (!fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        // 补全 .xlsx 后缀
                        fileName += ".xlsx";
                    }
                    // 拼接得到目标保存路径
                    string targetFilePath = Path.Combine(model.SavePath, fileName);

                    // 3. 获取 Excel Application COM 对象 (安全调用)
                    dynamic? app = ExcelDnaSafeAccessor.GetApplication();
                    // 校验 Application 是否获取成功
                    if (app == null) return false;

                    // 临时关闭屏幕刷新与警告弹出以提升渲染性能
                    app.ScreenUpdating = false;
                    // 关闭 Excel 操作对话框提示
                    app.DisplayAlerts = false;

                    try
                    {
                        // 4. 获取或创建 CabinetTemplate.xlsx 模板物理路径
                        string templatePath = EnsureCabinetTemplate(app);

                        // 5. 直接只读打开 CabinetTemplate.xlsx 模板并另存为目标物理路径
                        // 使用 SaveAs 方式可完全保留模板原汁原味的无外部链接本地公式结构
                        dynamic newWb = app.Workbooks.Open(templatePath, ReadOnly: true);

                        // 6. 另存为新的目标项目工作簿物理文件
                        newWb.SaveAs(targetFilePath);

                        // 7. 移除目标工作簿中非 "项目信息" 与非 "分类1" 的多余工作表
                        for (int i = newWb.Sheets.Count; i >= 1; i--)
                        {
                            // 获取对应位置工作表对象
                            dynamic item = newWb.Sheets[i];
                            // 获取工作表名称字符串
                            string name = (string)item.Name;
                            // 若非目标保留工作表则自动清理删除
                            if (name != "项目信息" && name != "分类1")
                            {
                                // 删掉多余工作表
                                item.Delete();
                            }
                        }

                        // 8. 调用 ExcelServices 统一完成【项目信息】与【分类1】工作表的完整回填、公式联动与定义名称锚点绑定
                        ExcelServices.InitializeCreatedProjectWorkbook(newWb, model);

                        // 9. 保存修改后的新工作簿
                        newWb.Save();

                        // 记录最近一次成功创建的目标文件路径
                        LastCreatedTargetFilePath = targetFilePath;

                        // 恢复 Excel 屏幕刷新与警告属性
                        app.ScreenUpdating = true;
                        app.DisplayAlerts = true;

                        // 返回成功结果
                        return true;
                    }
                    finally
                    {
                        // 兜底恢复 Excel 屏幕刷新与提示
                        app.ScreenUpdating = true;
                        app.DisplayAlerts = true;
                    }
                }
                catch (Exception ex)
                {
                    // 弹出捕获的异常信息窗口
                    MessageBox.Show($"创建项目工作簿失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // 返回创建失败标识
                    return false;
                }
            });
        }



        /// <summary>
        /// 检索 CabinetTemplate.xlsx 模板文件路径，若磁盘不存在则动态构建标准模板
        /// </summary>
        public static string EnsureCabinetTemplate(dynamic app)
        {
            // 获取基准运行目录
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            // 配置多重备选路径列表
            string[] candidates = new string[]
            {
                Path.Combine(baseDir, "Resources", "CabinetTemplate.xlsx"),
                Path.Combine(baseDir, "CabinetTemplate.xlsx"),
                Path.Combine(Directory.GetCurrentDirectory(), "Resources", "CabinetTemplate.xlsx"),
                Path.Combine(Directory.GetCurrentDirectory(), "CabinetTemplate.xlsx")
            };

            // 循环检索是否存在模板物理文件
            foreach (string candidate in candidates)
            {
                // 若命中了存在的模板文件
                if (File.Exists(candidate))
                {
                    // 立即返回存在的模板文件路径
                    return candidate;
                }
            }

            // 若均不存在，自动在 baseDir/Resources/ 下生成默认 CabinetTemplate.xlsx
            string targetDir = Path.Combine(baseDir, "Resources");
            // 校验目录是否存在
            Directory.CreateDirectory(targetDir);
            // 拼接目标模板路径
            string newTemplatePath = Path.Combine(targetDir, "CabinetTemplate.xlsx");

            // 动态新建模板工作簿
            dynamic wb = app.Workbooks.Add();

            // 重命名第一个工作表为 "项目信息"
            dynamic sheet1 = wb.Sheets[1];
            sheet1.Name = "项目信息";

            // 新增第二个工作表并重命名为 "分类1"
            dynamic sheet2 = wb.Sheets.Add(After: sheet1);
            sheet2.Name = "分类1";

            // 初始化 "项目信息" 模版样式与行标题
            sheet1.Range["B1"].Value = "扬州华科智能科技有限公司";

            // 写入【工程信息】各行 Label
            sheet1.Range["A4"].Value = "工程信息";
            sheet1.Range["A5"].Value = "项目名称";
            sheet1.Range["A6"].Value = "描述";
            sheet1.Range["A7"].Value = "报价单号";
            sheet1.Range["A8"].Value = "报价人";
            sheet1.Range["A9"].Value = "创建日期";
            sheet1.Range["A10"].Value = "报价审核人";
            sheet1.Range["A11"].Value = "项目负责人";
            sheet1.Range["A12"].Value = "项目备注";

            // 写入【客户信息】各行 Label
            sheet1.Range["A13"].Value = "客户信息";
            sheet1.Range["A14"].Value = "客户名称";
            sheet1.Range["A15"].Value = "联系人";
            sheet1.Range["A16"].Value = "联系电话";
            sheet1.Range["A17"].Value = "客户地址";
            sheet1.Range["A18"].Value = "客户邮编";
            sheet1.Range["A19"].Value = "客户网址";
            sheet1.Range["A20"].Value = "客户email";

            // 写入【本企业信息】各行 Label
            sheet1.Range["A21"].Value = "本企业信息";
            sheet1.Range["A22"].Value = "单位名称";
            sheet1.Range["A23"].Value = "英文名称";
            sheet1.Range["A24"].Value = "联系人";
            sheet1.Range["A25"].Value = "联系电话";
            sheet1.Range["A26"].Value = "销售地区";

            // 写入【分类汇总】各行 Header
            sheet1.Range["A27"].Value = "分类汇总";
            sheet1.Range["C27"].Value = "金额单位：人民币元";
            sheet1.Range["A28"].Value = "序号";
            sheet1.Range["B28"].Value = "分类名称";
            sheet1.Range["C28"].Value = "箱柜数量";
            sheet1.Range["D28"].Value = "总价";

            // 保存生成的模板工作簿
            wb.SaveAs(newTemplatePath);
            // 关闭工作簿句柄
            wb.Close(false);

            // 返回生成的新模板文件路径
            return newTemplatePath;
        }
    }
}
