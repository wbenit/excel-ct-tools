using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 核心业务服务分部类：云端项目绑定探测、箱柜结构化提取与直推
    /// 遵循 Convergence Cognitive Architecture 规范与每 3 行至少 1 行中文注释
    /// </summary>
    public static partial class ExcelServices
    {
        // 绑定项目窗体单例引用 (可空)
        private static Forms.ProjectBindForm? _projectBindFormInstance;

        // 箱柜同步直推窗体单例引用 (可空)
        private static Forms.CloudEBoxSyncForm? _cloudEBoxSyncFormInstance;

        /// <summary>
        /// 从指定或当前活动工作簿中读取已绑定的云端项目元数据 (双保险: Names + CustomDocumentProperties)
        /// </summary>
        /// <param name="wb">目标工作簿 COM 对象，若为空则自动取 ActiveWorkbook</param>
        /// <returns>绑定的项目ID、项目名称与工作组ID</returns>
        /// <summary>
        /// 从指定或当前活动工作簿中读取已绑定的云端项目元数据 (双保险: Names + CustomDocumentProperties)
        /// </summary>
        /// <param name="wb">目标工作簿 COM 对象，若为空则自动取 ActiveWorkbook</param>
        /// <returns>绑定的项目ID、项目名称与工作组ID</returns>
        public static (int ProjectId, string ProjectName, int GroupId) GetBoundProject(object? wb = null)
        {
            try
            {
                // 若未传入工作簿则尝试获取活动工作簿 (安全调用)
                dynamic? activeWb = wb ?? ExcelDnaSafeAccessor.GetApplication()?.ActiveWorkbook;
                if (activeWb == null) return (0, string.Empty, 0);

                int projectId = 0;
                string projectName = string.Empty;
                int groupId = 0;

                // 优先路径 1：通过工作簿定义名称 (Names) 极速读取
                try
                {
                    foreach (dynamic nameObj in activeWb.Names)
                    {
                        string n = Convert.ToString(nameObj.Name) ?? "";
                        string refers = Convert.ToString(nameObj.RefersTo) ?? "";
                        refers = refers.TrimStart('=', '\"').TrimEnd('\"');

                        if (string.Equals(n, "DrawCode_ProjectId", StringComparison.OrdinalIgnoreCase))
                        {
                            int.TryParse(refers, out projectId);
                        }
                        else if (string.Equals(n, "DrawCode_ProjectName", StringComparison.OrdinalIgnoreCase))
                        {
                            projectName = refers;
                        }
                        else if (string.Equals(n, "DrawCode_GroupId", StringComparison.OrdinalIgnoreCase))
                        {
                            int.TryParse(refers, out groupId);
                        }
                    }
                }
                catch { }

                // 容错路径 2：若定义名称未找到，从 CustomDocumentProperties 深度补查
                if (projectId <= 0)
                {
                    try
                    {
                        dynamic props = activeWb.CustomDocumentProperties;
                        foreach (dynamic p in props)
                        {
                            string pName = Convert.ToString(p.Name) ?? "";
                            string pVal = Convert.ToString(p.Value) ?? "";
                            if (string.Equals(pName, "DrawCode_ProjectId", StringComparison.OrdinalIgnoreCase))
                            {
                                int.TryParse(pVal, out projectId);
                            }
                            else if (string.Equals(pName, "DrawCode_ProjectName", StringComparison.OrdinalIgnoreCase))
                            {
                                projectName = pVal;
                            }
                            else if (string.Equals(pName, "DrawCode_GroupId", StringComparison.OrdinalIgnoreCase))
                            {
                                int.TryParse(pVal, out groupId);
                            }
                        }
                    }
                    catch { }
                }

                return (projectId, projectName, groupId);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CloudProject] GetBoundProject 异常: {ex.Message}");
                return (0, string.Empty, 0);
            }
        }

        /// <summary>
        /// 将选定的云端项目元数据持久化写入 Excel 工作簿中 (一次绑定，终生免弹)
        /// </summary>
        /// <param name="projectId">云端项目 ID</param>
        /// <param name="projectName">云端项目名称</param>
        /// <param name="groupId">所属工作组 ID</param>
        /// <param name="wb">目标工作簿，若空取 ActiveWorkbook</param>
        public static bool BindProjectToWorkbook(int projectId, string projectName, int groupId, object? wb = null)
        {
            try
            {
                // 获取当前活动工作簿
                dynamic? activeWb = wb ?? ExcelDnaSafeAccessor.GetApplication()?.ActiveWorkbook;
                if (activeWb == null) return false;

                // 1. 写入工作簿定义名称 (跨版本原生支持，零类型限制)
                SetOrAddDefinedName(activeWb, "DrawCode_ProjectId", $"=\"{projectId}\"");
                SetOrAddDefinedName(activeWb, "DrawCode_ProjectName", $"=\"{projectName}\"");
                SetOrAddDefinedName(activeWb, "DrawCode_GroupId", $"=\"{groupId}\"");

                // 2. 写入 CustomDocumentProperties (双重持久化)
                try
                {
                    dynamic props = activeWb.CustomDocumentProperties;
                    SetOrAddDocumentProperty(props, "DrawCode_ProjectId", projectId.ToString());
                    SetOrAddDocumentProperty(props, "DrawCode_ProjectName", projectName);
                    SetOrAddDocumentProperty(props, "DrawCode_GroupId", groupId.ToString());
                }
                catch { }

                LogHelper.WriteLog($"[CloudProject] 工作簿成功绑定云端项目: {projectName} (ID={projectId}, Group={groupId})");
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CloudProject] BindProjectToWorkbook 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 智能探测提取当前 Excel 工作簿对应的候选项目名称
        /// 遵循项目信息 B5 单元格优先，表名/文件名容错的管道机制
        /// </summary>
        public static string DetectCandidateProjectName(object? wb = null)
        {
            try
            {
                dynamic? activeWb = wb ?? ExcelDnaSafeAccessor.GetApplication()?.ActiveWorkbook;
                if (activeWb == null) return string.Empty;

                // 1. 优先尝试从【项目信息】工作表 B5 单元格提取工程名称 (业务标准位置)
                try
                {
                    dynamic infoSheet = activeWb.Sheets["项目信息"];
                    if (infoSheet != null)
                    {
                        string nameB5 = Convert.ToString(infoSheet.Range["B5"].Value2)?.Trim() ?? "";
                        if (!string.IsNullOrWhiteSpace(nameB5) && !nameB5.Contains("#REF") && !nameB5.Contains("项目名称"))
                        {
                            return nameB5;
                        }
                    }
                }
                catch { }

                // 2. 尝试从活动工作簿的文件名提取有效工程名称
                try
                {
                    string fileName = Convert.ToString(activeWb.Name) ?? "";
                    if (!string.IsNullOrWhiteSpace(fileName))
                    {
                        // 去除文件扩展名 (.xlsx, .xlsm, .xls)
                        string cleanName = Path.GetFileNameWithoutExtension(fileName).Trim();
                        // 过滤常见的后缀修饰（如：-终版、(1)、_v2、日期等）
                        cleanName = Regex.Replace(cleanName, @"[\(\（][^\)\）]*[\)\）]", "");
                        cleanName = Regex.Replace(cleanName, @"\d{8}|\d{4}-\d{2}-\d{2}", "");
                        cleanName = cleanName.Trim('-', '_', ' ');
                        if (!string.IsNullOrWhiteSpace(cleanName))
                        {
                            return cleanName;
                        }
                    }
                }
                catch { }

                return "新建工程项目";
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CloudProject] DetectCandidateProjectName 异常: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// 从【项目信息】工作表中解析所有分类明细表的分类序号映射
        /// </summary>
        /// <param name="wb">活动工作簿</param>
        /// <returns>字典：工作表名称 -> 分类序号 (例如 分类1 -> 1, 分类2 -> 2)</returns>
        public static Dictionary<string, int> GetProjectCategoryOrderMap(dynamic wb)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (wb == null) return map;
                dynamic? infoSheet = null;
                try { infoSheet = wb.Sheets["项目信息"]; } catch { }
                if (infoSheet == null) return map;

                var cfg = ConfigManager.Instance.Current.Excel;
                int startRow = cfg.ProjectInfoCategorySummaryStartRow; // 默认 29
                int maxScan = cfg.ProjectInfoCategorySummaryMaxScanRows; // 默认 50
                int endRow = startRow + maxScan - 1;

                object[,]? matrix = null;
                try
                {
                    dynamic rng = infoSheet.Range[$"A{startRow}:B{endRow}"];
                    matrix = rng?.Value2 as object[,];
                }
                catch { }

                if (matrix != null)
                {
                    int rows = matrix.GetLength(0);
                    int fallbackIndex = 1;
                    for (int i = 1; i <= rows; i++)
                    {
                        int physicalRow = startRow + i - 1;
                        string cellA = Convert.ToString(matrix[i, 1])?.Trim() ?? "";
                        string cellB = Convert.ToString(matrix[i, 2])?.Trim() ?? "";

                        if (cellB.Contains("小计") || cellB.Contains("合计")) break;

                        // 尝试解析 A 列中的数值作为分类序号
                        int catOrder = fallbackIndex;
                        if (int.TryParse(cellA, out var parsedOrder) && parsedOrder > 0)
                        {
                            catOrder = parsedOrder;
                        }

                        // 提取工作表名称（优先从超链接提取）
                        string targetSheetName = "";
                        try
                        {
                            dynamic cA = infoSheet.Cells[physicalRow, 1];
                            if (cA.Hyperlinks != null && cA.Hyperlinks.Count > 0)
                            {
                                string subAddr = Convert.ToString(cA.Hyperlinks[1].SubAddress)?.Trim() ?? "";
                                if (!string.IsNullOrEmpty(subAddr))
                                {
                                    int exclamIdx = subAddr.IndexOf('!');
                                    targetSheetName = (exclamIdx > 0 ? subAddr.Substring(0, exclamIdx) : subAddr).Trim('\'', ' ', '\"');
                                }
                            }
                        }
                        catch { }

                        if (string.IsNullOrEmpty(targetSheetName) && !string.IsNullOrWhiteSpace(cellB) &&
                            !cellB.Contains("#REF") && !cellB.StartsWith("#"))
                        {
                            targetSheetName = cellB;
                        }

                        if (!string.IsNullOrWhiteSpace(targetSheetName) && !map.ContainsKey(targetSheetName))
                        {
                            map[targetSheetName] = catOrder;
                            fallbackIndex++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CloudProject] GetProjectCategoryOrderMap 异常: {ex.Message}");
            }
            return map;
        }

        /// <summary>
        /// 从当前工作簿结构化提取全部箱柜数据以供直推云端系统图库
        /// 遵循用户指定规则：序号为“分类序号-分类明细顶部的汇总序号”，例如 2-1
        /// </summary>
        /// <param name="wb">活动工作簿</param>
        /// <returns>待推送至云端的箱柜数据集合</returns>
        public static List<CloudEBoxItemDto> ExtractAllCabinetsForCloudSync(object? wb = null)
        {
            var resultList = new List<CloudEBoxItemDto>();
            try
            {
                // 获取当前活动工作簿
                dynamic? activeWb = wb ?? ExcelDnaSafeAccessor.GetApplication()?.ActiveWorkbook;
                if (activeWb == null) return resultList;

                // 1. 获取已绑定的项目 ID
                var boundProj = GetBoundProject((object?)activeWb);
                int boundProjId = boundProj.ProjectId;

                // 2. 获取分类序号映射字典 (表名 -> 分类序号)
                var categoryOrderMap = GetProjectCategoryOrderMap(activeWb);

                // 3. 扫描当前工作簿登记在册的所有分类明细工作表
                var validSheetNames = Tool.GetProjectCategorySheetNames(activeWb);
                if (validSheetNames == null || validSheetNames.Count == 0)
                {
                    // 若白名单未初始化，尝试直接使用分类映射的所有键
                    validSheetNames = new HashSet<string>(categoryOrderMap.Keys, StringComparer.OrdinalIgnoreCase);
                }

                int defaultCatOrder = 1;

                // 4. 遍历每个分类工作表提取箱柜
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    string sName = Convert.ToString(sheet.Name) ?? "";
                    if (!validSheetNames.Contains(sName)) continue;

                    // 获取当前分类表的分类序号 (如 1, 2, 3...)
                    int categoryIndex = defaultCatOrder;
                    if (categoryOrderMap.TryGetValue(sName, out int foundOrder))
                    {
                        categoryIndex = foundOrder;
                    }
                    else
                    {
                        defaultCatOrder++;
                    }

                    // 获取该分类表所有有效箱柜锚点 (Zero COM 缓存机制)
                    var validCabinets = Tool.GetSheetValidCabinets((object)sheet, (object)activeWb);
                    if (validCabinets == null || validCabinets.Count == 0) continue;

                    // 5. 遍历每个箱柜提取核心属性与序号
                    foreach (var cabPair in validCabinets)
                    {
                        int cabKey = cabPair.Key; // 汇总序号 K
                        var anchor = cabPair.Value;

                        // 提取柜号与属性 (优先从顶部汇总行 Cab_Sum)
                        string cabinetNo = "";
                        string model = "";
                        int count = 1;

                        try
                        {
                            int sumRow = anchor.Sum != null ? Convert.ToInt32(anchor.Sum.Row) : (anchor.SumRow > 0 ? anchor.SumRow : 0);
                            if (sumRow > 0)
                            {
                                cabinetNo = Convert.ToString(sheet.Cells[sumRow, 2].Value2)?.Trim() ?? "";
                                string cName = Convert.ToString(sheet.Cells[sumRow, 3].Value2)?.Trim() ?? "";
                                model = Convert.ToString(sheet.Cells[sumRow, 4].Value2)?.Trim() ?? "";
                                try
                                {
                                    double qVal = Convert.ToDouble(sheet.Cells[sumRow, 6].Value2);
                                    if (qVal > 0) count = (int)Math.Round(qVal);
                                }
                                catch { }

                                // 若柜号为空则使用箱柜名称兜底
                                if (string.IsNullOrWhiteSpace(cabinetNo)) cabinetNo = cName;
                            }
                        }
                        catch { }

                        // 兜底容错：若未提取到柜号，采用规范格式
                        if (string.IsNullOrWhiteSpace(cabinetNo))
                        {
                            cabinetNo = $"{sName}-{cabKey}#柜";
                        }

                        // 💡 核心序号算法：分类序号 - 汇总序号（例如：2-1）
                        string finalOrder = $"{categoryIndex}-{cabKey}";

                        // 组装云端箱柜实体
                        var cloudItem = new CloudEBoxItemDto
                        {
                            Order = finalOrder,
                            SystemDrawName = cabinetNo,
                            Count = count,
                            ProjectId = boundProjId,
                            Model = model,
                            CategoryName = sName,
                            CategoryIndex = categoryIndex,
                            CabinetIndex = cabKey,
                            EboxPositon = sName
                        };

                        resultList.Add(cloudItem);
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CloudProject] ExtractAllCabinetsForCloudSync 异常: {ex.Message}");
            }
            return resultList;
        }

        /// <summary>
        /// 唤起基于 WebView2 + Vue 3 的“绑定云端项目”面板窗口 (非模态)
        /// </summary>
        public static void ShowProjectBindDialog()
        {
            try
            {
                // 以非模态方式展示项目绑定窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _projectBindFormInstance, () => new Forms.ProjectBindForm());
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"弹出绑定项目窗口失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 唤起基于 WebView2 + Vue 3 的“推送箱柜至云端”同步面板窗口 (非模态，包含覆盖确认)
        /// </summary>
        public static void ShowCloudEBoxSyncDialog()
        {
            try
            {
                // 先校验是否登录
                if (string.IsNullOrWhiteSpace(CurrentToken) || CurrentGroupId <= 0)
                {
                    // 弹出提示并唤起登录
                    var res = System.Windows.Forms.MessageBox.Show("尚未登录云端或未选择工作组，是否立即登录？", "身份认证提醒", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                    if (res == System.Windows.Forms.DialogResult.Yes)
                    {
                        ShowLoginDialog();
                    }
                    return;
                }

                // 以非模态方式展示箱柜推送同步窗口
                ShowModelessForm(ref _cloudEBoxSyncFormInstance, () => new Forms.CloudEBoxSyncForm());
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"弹出同步箱柜窗口失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        #region 私有辅助工具方法

        // 在工作簿中安全设置或更新定义名称
        private static void SetOrAddDefinedName(dynamic wb, string name, string refersTo)
        {
            try
            {
                // 优先尝试删除可能存在的旧名称
                try { wb.Names[name]?.Delete(); } catch { }
                // 新建并指向常量字符串公式
                wb.Names.Add(name, refersTo);
            }
            catch { }
        }

        // 在 CustomDocumentProperties 中安全设置或更新属性
        private static void SetOrAddDocumentProperty(dynamic props, string propName, string propValue)
        {
            try
            {
                bool exists = false;
                foreach (dynamic p in props)
                {
                    if (string.Equals(Convert.ToString(p.Name), propName, StringComparison.OrdinalIgnoreCase))
                    {
                        p.Value = propValue;
                        exists = true;
                        break;
                    }
                }
                if (!exists)
                {
                    // msoPropertyTypeString = 4
                    props.Add(propName, false, 4, propValue);
                }
            }
            catch { }
        }

        #endregion
    }
}
