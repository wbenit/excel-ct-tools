using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;

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
                        // 容错处理：剔除可能存在的工作表前缀限定（例如 "'Sheet1'!DrawCode_ProjectId" -> "DrawCode_ProjectId"）
                        string cleanName = n.Contains("!") ? n.Substring(n.LastIndexOf('!') + 1) : n;
                        string refers = Convert.ToString(nameObj.RefersTo) ?? "";
                        // 彻底清洗等于号与各种单双引号
                        string cleanVal = refers.Replace("=", "").Replace("\"", "").Trim();

                        if (string.Equals(cleanName, "DrawCode_ProjectId", StringComparison.OrdinalIgnoreCase))
                        {
                            int.TryParse(cleanVal, out projectId);
                        }
                        else if (string.Equals(cleanName, "DrawCode_ProjectName", StringComparison.OrdinalIgnoreCase))
                        {
                            projectName = cleanVal;
                        }
                        else if (string.Equals(cleanName, "DrawCode_GroupId", StringComparison.OrdinalIgnoreCase))
                        {
                            int.TryParse(cleanVal, out groupId);
                        }
                    }
                }
                catch { }

                // 容错路径 2：从 CustomDocumentProperties 深度互补读取
                if (projectId <= 0 || string.IsNullOrEmpty(projectName))
                {
                    try
                    {
                        dynamic props = activeWb.CustomDocumentProperties;
                        foreach (dynamic p in props)
                        {
                            string pName = Convert.ToString(p.Name) ?? "";
                            string pVal = Convert.ToString(p.Value)?.Trim() ?? "";
                            if (string.Equals(pName, "DrawCode_ProjectId", StringComparison.OrdinalIgnoreCase) && projectId <= 0)
                            {
                                int.TryParse(pVal, out projectId);
                            }
                            else if (string.Equals(pName, "DrawCode_ProjectName", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(projectName))
                            {
                                projectName = pVal;
                            }
                            else if (string.Equals(pName, "DrawCode_GroupId", StringComparison.OrdinalIgnoreCase) && groupId <= 0)
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
        /// <summary>
        /// 云端箱柜属性独立协同工作表固定名称
        /// </summary>
        public const string CloudPropsSheetName = "云端箱柜属性";

        /// <summary>
        /// <summary>
        /// 获取当前活动工作簿中所有有效的分类明细工作表名称列表
        /// 遵循规范：每 3 行至少 1 行中文注释
        /// </summary>
        /// <param name="wb">活动工作簿对象</param>
        /// <returns>分类明细表名称集合列表</returns>
        public static List<string> GetProjectCategorySheetNamesList(object? wb = null)
        {
            var result = new List<string>();
            try
            {
                dynamic? activeWb = wb ?? ExcelDnaSafeAccessor.GetApplication()?.ActiveWorkbook;
                if (activeWb == null) return result;

                // 读取分类映射与分类明细登记表名白名单
                var categoryOrderMap = GetProjectCategoryOrderMap(activeWb);
                var validSheetNames = Tool.GetProjectCategorySheetNames(activeWb);
                if (validSheetNames == null || validSheetNames.Count == 0)
                {
                    validSheetNames = new HashSet<string>(categoryOrderMap.Keys, StringComparer.OrdinalIgnoreCase);
                }

                // 遍历当前活动工作簿的实际工作表
                foreach (dynamic sheet in activeWb.Worksheets)
                {
                    string sName = Convert.ToString(sheet.Name) ?? "";
                    if (validSheetNames.Contains(sName) && !result.Contains(sName, StringComparer.OrdinalIgnoreCase))
                    {
                        result.Add(sName);
                    }
                }

                // 若包含分类映射顺序，按分类序号升序排列
                if (categoryOrderMap != null && categoryOrderMap.Count > 0)
                {
                    result = result.OrderBy(name => categoryOrderMap.TryGetValue(name, out int o) ? o : 999).ToList();
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CloudProject] GetProjectCategorySheetNamesList 异常: {ex.Message}");
            }
            return result;
        }

        /// <summary>
        /// 从当前工作簿结构化提取箱柜数据以供直推云端系统图库 (支持按选中的分类明细进行过滤)
        /// 遵循用户指定规则：序号为“分类序号-分类明细顶部的汇总序号”，例如 2-1
        /// </summary>
        /// <param name="wb">活动工作簿</param>
        /// <param name="readCustomPropsSheet">是否同时关联提取【云端箱柜属性】工作表中的扩展属性</param>
        /// <param name="selectedCategories">可选指定需要提取的分类明细工作表名称集合，若为空则默认提取全部</param>
        /// <returns>待推送至云端的箱柜数据集合</returns>
        public static List<CloudEBoxItemDto> ExtractAllCabinetsForCloudSync(object? wb = null, bool readCustomPropsSheet = true, List<string>? selectedCategories = null)
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

                    // 遵循用户明确指示：若指定了被选中的分类明细，未勾选的分类工作表直接跳过
                    if (selectedCategories != null && selectedCategories.Count > 0 && !selectedCategories.Contains(sName, StringComparer.OrdinalIgnoreCase))
                    {
                        continue;
                    }

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

                // 6. 核心业务逻辑：若工作簿中存在【云端箱柜属性】协同表，以“箱柜序号”为 Key 全量同步 Excel 中的任何修改，包括删除！
                if (readCustomPropsSheet)
                {
                    try
                    {
                        dynamic? propsSheet = null;
                        foreach (dynamic s in activeWb.Worksheets)
                        {
                            if (string.Equals(Convert.ToString(s.Name), CloudPropsSheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                propsSheet = s;
                                break;
                            }
                        }

                        if (propsSheet != null)
                        {
                            int pRows = propsSheet.UsedRange?.Rows?.Count ?? 0;
                            int pCols = propsSheet.UsedRange?.Columns?.Count ?? 0;

                            if (pRows > 1 && pCols >= 1)
                            {
                                // 规则 7：二维数组一次性读取到内存
                                dynamic pRng = propsSheet.Range[propsSheet.Cells[1, 1], propsSheet.Cells[pRows, pCols]];
                                object[,] pMatrix = pRng?.Value2 as object[,];

                                if (pMatrix != null)
                                {
                                    int mRows = pMatrix.GetLength(0);
                                    int mCols = pMatrix.GetLength(1);

                                    // 解析第 6 列及以后的自定义列名映射
                                    var customCols = new List<(int ColIndex, string Name)>();
                                    for (int c = 6; c <= mCols; c++)
                                    {
                                        string colName = Convert.ToString(pMatrix[1, c])?.Trim() ?? "";
                                        if (!string.IsNullOrEmpty(colName))
                                        {
                                            customCols.Add((c, colName));
                                        }
                                    }

                                    // 读取属性表中以“箱柜序号”为 Key 的全量最新数据
                                    var propsMap = new Dictionary<string, (string Position, string DrawName, int Count, string Model, string CustomFields)>(StringComparer.OrdinalIgnoreCase);
                                    for (int r = 2; r <= mRows; r++)
                                    {
                                        string order = Convert.ToString(pMatrix[r, 1])?.Trim() ?? "";
                                        if (string.IsNullOrEmpty(order)) continue;

                                        // 容错：若不幸仍有被 Excel 自动转换为日期的，兼容反解为 "1-1"
                                        if (DateTime.TryParse(order, out var dt))
                                        {
                                            order = $"{dt.Month}-{dt.Day}";
                                        }

                                        // 提取位置 (第 2 列)
                                        string pos = mCols >= 2 ? (Convert.ToString(pMatrix[r, 2])?.Trim() ?? "") : "";
                                        // 提取箱柜代号/名称 (第 3 列)
                                        string dName = mCols >= 3 ? (Convert.ToString(pMatrix[r, 3])?.Trim() ?? "") : "";
                                        // 提取数量 (第 4 列)
                                        int qCount = 1;
                                        if (mCols >= 4)
                                        {
                                            try
                                            {
                                                double q = Convert.ToDouble(pMatrix[r, 4]);
                                                if (q > 0) qCount = (int)Math.Round(q);
                                            }
                                            catch { }
                                        }
                                        // 提取型号规格 (第 5 列)
                                        string mModel = mCols >= 5 ? (Convert.ToString(pMatrix[r, 5])?.Trim() ?? "") : "";

                                        // 提取自定义扩展列 (第 6 列及以后)
                                        var rowDict = new Dictionary<string, object>();
                                        foreach (var col in customCols)
                                        {
                                            var val = pMatrix[r, col.ColIndex];
                                            if (val != null)
                                            {
                                                string strVal = Convert.ToString(val)?.Trim() ?? "";
                                                if (!string.IsNullOrEmpty(strVal))
                                                {
                                                    rowDict[col.Name] = strVal;
                                                }
                                            }
                                        }
                                        string cFieldsJson = rowDict.Count > 0 ? JsonSerializer.Serialize(rowDict) : "";

                                        propsMap[order] = (pos, dName, qCount, mModel, cFieldsJson);
                                    }

                                    // 遵循用户明确指示：以“箱柜序号”为 Key，把 Excel 中的任何修改全量同步，包括删除！
                                    var syncedList = new List<CloudEBoxItemDto>();

                                    // 遍历从分类表提取出的箱柜列表
                                    foreach (var item in resultList)
                                    {
                                        if (string.IsNullOrEmpty(item.Order)) continue;

                                        // 若某箱柜在【云端箱柜属性】表中依然存在，执行全字段修改覆盖并保留
                                        if (propsMap.TryGetValue(item.Order, out var rowData))
                                        {
                                            // 1. 同步修改箱柜代号/名称
                                            if (!string.IsNullOrEmpty(rowData.DrawName))
                                            {
                                                item.SystemDrawName = rowData.DrawName;
                                            }
                                            // 2. 同步修改位置 (原所属分类)
                                            if (!string.IsNullOrEmpty(rowData.Position))
                                            {
                                                item.EboxPositon = rowData.Position;
                                                item.CategoryName = rowData.Position;
                                            }
                                            // 3. 同步修改数量
                                            if (rowData.Count > 0)
                                            {
                                                item.Count = rowData.Count;
                                            }
                                            // 4. 同步修改型号规格
                                            item.Model = rowData.Model;
                                            // 5. 同步修改自定义扩展字段 JSON
                                            item.CustomFields = rowData.CustomFields;

                                            syncedList.Add(item);
                                            // 标记已处理该箱柜
                                            propsMap.Remove(item.Order);
                                        }
                                        else
                                        {
                                            // 核心满足用户需求：若在 Excel 属性表中被删除，则从推送列表中同步彻底删除！
                                            LogHelper.WriteLog($"[CloudProject] 箱柜 {item.Order} 在属性表中已删除，已同步剔除推送列表");
                                        }
                                    }

                                    // 若用户在属性表中手动新增了箱柜行，也作为有效箱柜一并纳入推送列表
                                    foreach (var kvp in propsMap)
                                    {
                                        string newOrder = kvp.Key;
                                        var rowData = kvp.Value;

                                        // 若指定了选中分类，非选中分类的箱柜行直接忽略
                                        if (selectedCategories != null && selectedCategories.Count > 0)
                                        {
                                            if (string.IsNullOrEmpty(rowData.Position) || !selectedCategories.Contains(rowData.Position, StringComparer.OrdinalIgnoreCase))
                                            {
                                                continue;
                                            }
                                        }

                                        int catIdx = 1;
                                        int cabIdx = 1;
                                        var parts = newOrder.Split('-');
                                        if (parts.Length >= 2 && int.TryParse(parts[0], out int p1) && int.TryParse(parts[1], out int p2))
                                        {
                                            catIdx = p1;
                                            cabIdx = p2;
                                        }

                                        var newItem = new CloudEBoxItemDto
                                        {
                                            Order = newOrder,
                                            SystemDrawName = !string.IsNullOrEmpty(rowData.DrawName) ? rowData.DrawName : $"{newOrder}#柜",
                                            EboxPositon = !string.IsNullOrEmpty(rowData.Position) ? rowData.Position : "默认位置",
                                            CategoryName = !string.IsNullOrEmpty(rowData.Position) ? rowData.Position : "默认位置",
                                            Count = rowData.Count > 0 ? rowData.Count : 1,
                                            Model = rowData.Model,
                                            ProjectId = boundProjId,
                                            CategoryIndex = catIdx,
                                            CabinetIndex = cabIdx,
                                            CustomFields = rowData.CustomFields
                                        };
                                        syncedList.Add(newItem);
                                    }

                                    // 最终替换为按箱柜序号同步修改与删除后的全新列表
                                    resultList = syncedList;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogHelper.WriteLog($"[CloudProject] 关联云端箱柜属性表数据容错: {ex.Message}");
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
        /// 从云端动态拉取当前绑定项目的自定义字段规范，并自动生成或增量更新【云端箱柜属性】协同工作表
        /// 遵循规范：字段非写死每次实时拉取；遵守规则 7 采用二维数组一次性读写内存；保留用户已填写数据；仅包含被选中的分类明细
        /// </summary>
        /// <param name="targetProjectId">可选指定目标项目 ID，若未提供则自动读取工作簿绑定项目</param>
        /// <param name="wb">可选活动工作簿对象</param>
        /// <param name="selectedCategories">可选指定需要包含的分类明细工作表名称集合，若为空则默认全部</param>
        /// <returns>(是否成功, 提示消息)</returns>
        public static async Task<(bool Success, string Message)> SyncCloudCustomPropsSheetAsync(int? targetProjectId = null, object? wb = null, List<string>? selectedCategories = null)
        {
            try
            {
                // 1. 获取活动工作簿对象
                dynamic? activeWb = wb ?? ExcelDnaSafeAccessor.GetApplication()?.ActiveWorkbook;
                if (activeWb == null)
                {
                    return (false, "未检测到活动的 Excel 工作簿");
                }

                // 2. 获取目标项目 ID
                int projectId = targetProjectId ?? 0;
                if (projectId <= 0)
                {
                    var bound = GetBoundProject((object?)activeWb);
                    projectId = bound.ProjectId;
                }

                if (projectId <= 0)
                {
                    return (false, "当前工作簿尚未绑定云端项目，请先绑定项目！");
                }

                // 3. 实时从云端 WebAPI 动态拉取当前项目模板的自定义字段列表 (绝不硬编码，每次实时拉取)
                var (fetchOk, fetchMsg, customLabels) = await DrawCodeApiClient.GetProjectCustomFieldLabelsAsync(projectId);
                if (!fetchOk)
                {
                    return (false, $"拉取云端自定义字段失败: {fetchMsg}");
                }

                // 4. 从当前工作簿中提取所有现存箱柜的基础信息集合 (遵循用户指示：仅提取被选中的分类明细)
                var cabinets = ExtractAllCabinetsForCloudSync((object?)activeWb, readCustomPropsSheet: false, selectedCategories: selectedCategories);
                if (cabinets == null || cabinets.Count == 0)
                {
                    return (false, "当前工作簿中未检测到选定分类的有效箱柜，请先完善分类明细表！");
                }

                // 5. 检查工作簿中是否已存在【云端箱柜属性】工作表，若存在则预先备份用户已填数据
                dynamic? propsSheet = null;
                var existingValues = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

                foreach (dynamic s in activeWb.Worksheets)
                {
                    if (string.Equals(Convert.ToString(s.Name), CloudPropsSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        propsSheet = s;
                        break;
                    }
                }

                // 若已存在该工作表，安全读取已有自定义列的历史填报数据以防被覆盖
                if (propsSheet != null)
                {
                    try
                    {
                        // 获取旧表已用区域的最大行数和最大列数
                        int usedRows = propsSheet.UsedRange?.Rows?.Count ?? 0;
                        int usedCols = propsSheet.UsedRange?.Columns?.Count ?? 0;

                        if (usedRows > 1 && usedCols >= 6)
                        {
                            // 规则 7：二维数组一次性读取内存
                            dynamic oldRng = propsSheet.Range[propsSheet.Cells[1, 1], propsSheet.Cells[usedRows, usedCols]];
                            object[,] oldMatrix = oldRng?.Value2 as object[,];

                            if (oldMatrix != null)
                            {
                                int oRows = oldMatrix.GetLength(0);
                                int oCols = oldMatrix.GetLength(1);

                                // 解析旧表头中第 6 列起的自定义列名映射 (列索引 -> 列名)
                                var oldColNames = new Dictionary<int, string>();
                                for (int c = 6; c <= oCols; c++)
                                {
                                    string colName = Convert.ToString(oldMatrix[1, c])?.Trim() ?? "";
                                    if (!string.IsNullOrEmpty(colName)) oldColNames[c] = colName;
                                }

                                // 遍历数据行暂存旧数据: key 为箱柜序号 (如 "1-1")
                                for (int r = 2; r <= oRows; r++)
                                {
                                    string orderKey = Convert.ToString(oldMatrix[r, 1])?.Trim() ?? "";
                                    if (string.IsNullOrEmpty(orderKey)) continue;

                                    // 容错：若旧表中之前不幸已被 Excel 自动识别为日期（如 "1月1日" 或 "2026/1/1"），智能还原为 "1-1"
                                    if (DateTime.TryParse(orderKey, out var dt))
                                    {
                                        orderKey = $"{dt.Month}-{dt.Day}";
                                    }

                                    if (!existingValues.TryGetValue(orderKey, out var fieldMap))
                                    {
                                        fieldMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                        existingValues[orderKey] = fieldMap;
                                    }

                                    foreach (var kvp in oldColNames)
                                    {
                                        string cellVal = Convert.ToString(oldMatrix[r, kvp.Key])?.Trim() ?? "";
                                        if (!string.IsNullOrEmpty(cellVal))
                                        {
                                            fieldMap[kvp.Value] = cellVal;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogHelper.WriteLog($"[CloudProject] 读取已有自定义属性表历史数据容错: {ex.Message}");
                    }
                }
                else
                {
                    // 若不存在，则新建该工作表并置于末尾
                    propsSheet = activeWb.Worksheets.Add(After: activeWb.Worksheets[activeWb.Worksheets.Count]);
                    propsSheet.Name = CloudPropsSheetName;
                }

                // 6. 构造全新的二维网格数据 (表头 + 数据行)
                int fixedColCount = 5; // A~E 为固定列
                int customColCount = customLabels.Count;
                int totalCols = fixedColCount + customColCount;
                int totalRows = cabinets.Count + 1; // 包含 1 行表头

                // 规则 7：创建二维数组，准备一次性写入 Excel (索引从 1 开始对应 Excel)
                object[,] newGrid = new object[totalRows, totalCols];

                // 填充第 1 行表头 (数组下界为 0，遵循用户要求：所属分类修改为位置)
                newGrid[0, 0] = "箱柜序号";
                newGrid[0, 1] = "位置";
                newGrid[0, 2] = "箱柜代号/名称";
                newGrid[0, 3] = "数量";
                newGrid[0, 4] = "";

                // 填充动态自定义表头
                for (int ci = 0; ci < customColCount; ci++)
                {
                    newGrid[0, fixedColCount + ci] = customLabels[ci];
                }

                // 填充各个箱柜行数据
                for (int i = 0; i < cabinets.Count; i++)
                {
                    var cab = cabinets[i];
                    int rIdx = i + 1; // 对应第 rIdx 行

                    // 填充箱柜序号与位置
                    newGrid[rIdx, 0] = cab.Order ?? "";
                    newGrid[rIdx, 1] = !string.IsNullOrEmpty(cab.EboxPositon) ? cab.EboxPositon : (cab.CategoryName ?? "");
                    newGrid[rIdx, 2] = cab.SystemDrawName ?? "";
                    newGrid[rIdx, 3] = cab.Count;
                    newGrid[rIdx, 4] = cab.Model ?? "";

                    // 检索该箱柜的历史已填自定义属性进行还原匹配
                    if (!string.IsNullOrEmpty(cab.Order) && existingValues.TryGetValue(cab.Order, out var oldFields))
                    {
                        for (int ci = 0; ci < customColCount; ci++)
                        {
                            string fieldName = customLabels[ci];
                            if (oldFields.TryGetValue(fieldName, out var val))
                            {
                                newGrid[rIdx, fixedColCount + ci] = val;
                            }
                            else
                            {
                                newGrid[rIdx, fixedColCount + ci] = "";
                            }
                        }
                    }
                    else
                    {
                        for (int ci = 0; ci < customColCount; ci++)
                        {
                            newGrid[rIdx, fixedColCount + ci] = "";
                        }
                    }
                }

                // 7. 规则 7：一次性写回工作表内存
                propsSheet.Cells.Clear();

                // 遵循用户明确指示：A 列箱柜序号必须预先强制设为纯文本格式 "@"，防止类似 "1-1" 被 Excel 误转换为日期 "1月1日"
                try
                {
                    propsSheet.Columns[1].NumberFormatLocal = "@";
                }
                catch { }

                dynamic writeRange = propsSheet.Range[propsSheet.Cells[1, 1], propsSheet.Cells[totalRows, totalCols]];
                writeRange.Value2 = newGrid;

                // 再次确保已写入区域的 A 列依然保持纯文本格式 "@"
                try
                {
                    dynamic orderColFormat = propsSheet.Range[propsSheet.Cells[2, 1], propsSheet.Cells[totalRows, 1]];
                    orderColFormat.NumberFormatLocal = "@";
                }
                catch { }

                // 8. 样式美化与视觉引导定制 (主色调 #009688 绿蓝相间视觉)
                try
                {
                    // 设置整体字体与行高
                    propsSheet.Cells.Font.Name = "微软雅黑";
                    propsSheet.Cells.Font.Size = 10;
                    propsSheet.Rows[1].RowHeight = 28;

                    // 设置固定信息列表头样式 (浅灰底 #F1F5F9，深灰文字)
                    dynamic fixedHeader = propsSheet.Range[propsSheet.Cells[1, 1], propsSheet.Cells[1, fixedColCount]];
                    fixedHeader.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(241, 245, 249));
                    fixedHeader.Font.Bold = true;
                    fixedHeader.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(51, 65, 85));
                    fixedHeader.HorizontalAlignment = -4108; // xlCenter

                    // 设置自定义属性列表头样式 (翠青底 #009688，白色粗体文字)
                    if (customColCount > 0)
                    {
                        dynamic customHeader = propsSheet.Range[propsSheet.Cells[1, fixedColCount + 1], propsSheet.Cells[1, totalCols]];
                        customHeader.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 150, 136));
                        customHeader.Font.Bold = true;
                        customHeader.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        customHeader.HorizontalAlignment = -4108; // xlCenter
                    }

                    // A 列箱柜序号加粗居中
                    dynamic orderCol = propsSheet.Range[propsSheet.Cells[2, 1], propsSheet.Cells[totalRows, 1]];
                    orderCol.Font.Bold = true;
                    orderCol.HorizontalAlignment = -4108; // xlCenter

                    // 自动适应列宽
                    writeRange.Borders.LineStyle = 1; // 实线细边框
                    writeRange.Columns.AutoFit();
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[CloudProject] 设置样式容错: {ex.Message}");
                }

                // 9. 自动激活该属性表，方便用户直接查看并批量录入
                try { propsSheet.Activate(); } catch { }

                string resultMsg = customColCount > 0
                    ? $"成功生成【{CloudPropsSheetName}】表！已动态拉取 {customColCount} 个云端自定义列，共加载 {cabinets.Count} 台箱柜。"
                    : $"成功生成【{CloudPropsSheetName}】表！该云端项目未配置自定义列，已列出 {cabinets.Count} 台箱柜。";

                return (true, resultMsg);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CloudProject] SyncCloudCustomPropsSheetAsync 异常: {ex}");
                return (false, $"同步云端属性表异常: {ex.Message}");
            }
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
