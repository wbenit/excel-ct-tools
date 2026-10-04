using System;
using System.Collections.Generic;
using System.IO;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;
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

        // 本地文件管理控制台窗口静态单例引用 (可空)
        private static Forms.LocalProjectForm? _localProjectForm;

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“本地文件管理控制台”窗口 (非模态，可与 Excel 并行操作)
        /// </summary>
        public static void ShowLocalProjectDialog()
        {
            try
            {
                // 1. 获取当前活动 Excel 工作簿的物理路径与项目名称
                var (currentPath, currentProj) = Forms.LocalProjectForm.GetActiveWorkbookInfo();

                // 2. 若窗体已存在且未释放，向其通知切换聚焦到当前最新工作簿与项目
                if (_localProjectForm != null && !_localProjectForm.IsDisposed)
                {
                    // 通知 WebView2 页面切换项目
                    _localProjectForm.SwitchToProject(currentPath, currentProj);
                }

                // 3. 以非模态方式展示本地文件管理控制台窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _localProjectForm, () => new Forms.LocalProjectForm());
            }
            catch (Exception ex)
            {
                // 记录异常日志信息
                LogHelper.WriteLog($"[ExcelServices] 弹出本地文件管理窗口失败: {ex.Message}");
                // 弹出异常提示对话框
                System.Windows.Forms.MessageBox.Show($"弹出本地文件管理窗口失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        #region 本地文件管理控制台数据比对与 Ribbon 状态判定

        // 缓存最近一次检测的工作簿全路径
        private static string _lastCheckedLocalProjectPath = string.Empty;

        // 缓存最近一次判定的结果 (true: 已在表中, false: 未在表中)
        private static bool _lastCheckedLocalProjectResult = false;

        // 缓存最近一次检测的时间戳
        private static DateTime _lastCheckedLocalProjectTime = DateTime.MinValue;

        // 本地项目路径检测同步锁对象
        private static readonly object _localProjectCheckLock = new object();

        /// <summary>
        /// 主动使本地项目路径比对缓存失效
        /// </summary>
        public static void InvalidateLocalProjectCache()
        {
            // 加锁保障线程安全
            lock (_localProjectCheckLock)
            {
                // 清空路径缓存字符串
                _lastCheckedLocalProjectPath = string.Empty;
                // 重置时间戳为最小值强制下次重新扫描
                _lastCheckedLocalProjectTime = DateTime.MinValue;
            }
        }

        /// <summary>
        /// 判断当前 Excel 活动工作簿物理全路径是否已收录在本地文件管理控制台表格中
        /// </summary>
        /// <returns>若已在控制台表格中返回 true，否则返回 false</returns>
        public static bool IsActiveWorkbookInLocalProjectTable()
        {
            try
            {
                // 1. 获取当前活动工作簿的物理全路径与工程名
                var (activeFilePath, activeProject) = Forms.LocalProjectForm.GetActiveWorkbookInfo();

                // 若当前没有打开的工作簿或文件未保存到磁盘根路径，直接判定为未在表中
                if (string.IsNullOrWhiteSpace(activeFilePath) || !Path.IsPathRooted(activeFilePath))
                {
                    // 未存盘文件直接返回 false
                    return false;
                }

                // 2. 检查 3 秒内存短效缓存，消除高频 Ribbon 渲染开销
                lock (_localProjectCheckLock)
                {
                    // 若路径相同且在 3 秒有效期内，直接返回上次判定的纯内存缓存结果
                    if (string.Equals(_lastCheckedLocalProjectPath, activeFilePath, StringComparison.OrdinalIgnoreCase) &&
                        (DateTime.Now - _lastCheckedLocalProjectTime).TotalSeconds < 3.0)
                    {
                        // 命中短效纯内存缓存
                        return _lastCheckedLocalProjectResult;
                    }
                }

                // 3. 执行物理文件与 JSON 扫描比对
                bool isInTable = CheckIfFilePathInLocalProjectJson(activeFilePath, activeProject);

                // 4. 更新短效缓存
                lock (_localProjectCheckLock)
                {
                    // 记录最新检测的物理文件路径
                    _lastCheckedLocalProjectPath = activeFilePath;
                    // 记录判定结果
                    _lastCheckedLocalProjectResult = isInTable;
                    // 记录检测时间戳
                    _lastCheckedLocalProjectTime = DateTime.Now;
                }

                // 返回最终判定结果
                return isInTable;
            }
            catch (Exception ex)
            {
                // 记录检测异常日志
                LogHelper.WriteLog($"[ExcelServices] 校验当前工作簿是否在本地项目表中异常: {ex.Message}");
                // 异常情况下安全返回 false
                return false;
            }
        }

        /// <summary>
        /// 核心比对方法：扫描探测到的 projectJsonFile 目录下的工程 JSON，比对目标文件是否存在
        /// </summary>
        private static bool CheckIfFilePathInLocalProjectJson(string activeFilePath, string activeProject)
        {
            try
            {
                // 路径标准化清洗：转小写并统一反斜杠为正斜杠
                string normHostPath = activeFilePath.Replace('\\', '/').Trim().ToLowerInvariant();
                // 提取活动工作簿的纯文件名小写
                string hostFileName = Path.GetFileName(activeFilePath).ToLowerInvariant();

                // 收集所有候选的 projectJsonFile 目录路径集合
                var candidateDirs = new List<string>();

                // 策略 A：从当前工作簿所在目录逐级向上探测是否存在 projectJsonFile 目录
                try
                {
                    // 提取当前文件目录
                    string currentDir = Path.GetDirectoryName(activeFilePath) ?? string.Empty;
                    // 构建目录对象
                    DirectoryInfo? cur = new DirectoryInfo(currentDir);
                    // 逐层向上攀爬探测
                    while (cur != null)
                    {
                        // 拼接候选子目录
                        string candidate = Path.Combine(cur.FullName, "projectJsonFile");
                        // 校验该工程 JSON 目录物理是否存在
                        if (Directory.Exists(candidate))
                        {
                            // 探测命中并加入候选列表
                            candidateDirs.Add(candidate);
                            // 终止向上探测
                            break;
                        }
                        // 向上跳至父目录
                        cur = cur.Parent;
                    }
                }
                catch { }

                // 策略 B：预设常用/已知授权工程目录候选池 --硬编码: 本地预设工程搜索目录--
                string[] defaultDirs = new[]
                {
                    @"E:\BaiduNetdiskWorkspace\hj\projectJsonFile",
                    @"E:\2026\projectJsonFile",
                    @"D:\xingrenfile\projectJsonFile",
                    @"D:\xingrenfile"
                };
                // 遍历检查系统默认工程目录池
                foreach (var dir in defaultDirs)
                {
                    // 若物理目录存在且未包含，加入候选池
                    if (Directory.Exists(dir) && !candidateDirs.Contains(dir, StringComparer.OrdinalIgnoreCase))
                    {
                        // 添加候选目录
                        candidateDirs.Add(dir);
                    }
                }

                // 若未探测到任何有效的 projectJsonFile 目录，直接返回未命中
                if (candidateDirs.Count == 0)
                {
                    // 无有效工程目录返回 false
                    return false;
                }

                // 策略 C：优先识别与比对当前工程名对应的单体 JSON 文件
                string targetProjName = !string.IsNullOrWhiteSpace(activeProject)
                    ? activeProject
                    : Forms.LocalProjectForm.ExtractProjectNameFromPath(activeFilePath);

                // 遍历每个候选目录执行比对
                foreach (var dir in candidateDirs)
                {
                    // 1. 优先尝试直接加载对应的 {工程名}.json
                    if (!string.IsNullOrWhiteSpace(targetProjName))
                    {
                        // 拼接特定工程 JSON 路径
                        string specificJsonPath = Path.Combine(dir, $"{targetProjName}.json");
                        // 校验目标工程 JSON 物理存在
                        if (File.Exists(specificJsonPath))
                        {
                            // 单文件比对命中直接返回 true
                            if (CheckSingleProjectJson(specificJsonPath, normHostPath, hostFileName, targetProjName))
                            {
                                // 命中返回 true
                                return true;
                            }
                        }
                    }

                    // 2. 遍历该目录下所有的 .json 文件进行全面排查
                    string[] jsonFiles = Directory.GetFiles(dir, "*.json", SearchOption.TopDirectoryOnly);
                    // 逐个比对目录下全部工程 JSON
                    foreach (var jsonFile in jsonFiles)
                    {
                        // 排除已比对过的特定工程 JSON
                        if (!string.IsNullOrWhiteSpace(targetProjName) &&
                            string.Equals(Path.GetFileNameWithoutExtension(jsonFile), targetProjName, StringComparison.OrdinalIgnoreCase))
                        {
                            // 跳过已检测文件
                            continue;
                        }

                        // 单文件比对命中直接返回 true
                        if (CheckSingleProjectJson(jsonFile, normHostPath, hostFileName, targetProjName))
                        {
                            // 命中返回 true
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录扫描异常日志
                LogHelper.WriteLog($"[ExcelServices] 扫描 projectJsonFile 异常: {ex.Message}");
            }

            // 全部未匹配返回 false
            return false;
        }

        /// <summary>
        /// 读取单个工程 JSON 文件并校验是否包含目标文件路径或文件名
        /// </summary>
        private static bool CheckSingleProjectJson(string jsonPath, string normHostPath, string hostFileName, string targetProjName)
        {
            try
            {
                // 校验文件是否存在
                if (!File.Exists(jsonPath)) return false;

                // 快速粗筛：直接读取全部文本进行轻量包含检测，若不含文件名则直接排除
                string text = File.ReadAllText(jsonPath, System.Text.Encoding.UTF8);
                // 判断文本中是否包含文件名 (使用 IndexOf 保证兼容 .NET Framework 4.8)
                if (string.IsNullOrWhiteSpace(text) || text.IndexOf(hostFileName, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    // 纯文本未包含目标文件名，瞬间短路排除
                    return false;
                }

                // 使用 System.Text.Json 进行精准树解析
                using var doc = System.Text.Json.JsonDocument.Parse(text);
                // 获取 JSON 根节点
                var root = doc.RootElement;

                // 提取 JSON 中的工程名称 (兼容大小写)
                string projName = string.Empty;
                // 读取 ProjectName 或 projectName
                if (root.TryGetProperty("ProjectName", out var p1) || root.TryGetProperty("projectName", out p1))
                {
                    // 获取工程名文本
                    projName = p1.GetString() ?? string.Empty;
                }

                // 提取 Files 列表 (兼容大小写或直接为根数组)
                System.Text.Json.JsonElement filesElement = default;
                // 判定根是否为数组
                if (root.ValueKind == System.Text.Json.JsonValueKind.Array)
                {
                    // 赋予根数组
                    filesElement = root;
                }
                // 判定是否有名为 Files 或 files 的属性
                else if (root.TryGetProperty("Files", out var f1) || root.TryGetProperty("files", out f1))
                {
                    // 赋予 Files 节点
                    filesElement = f1;
                }

                // 若非有效数组则跳过
                if (filesElement.ValueKind != System.Text.Json.JsonValueKind.Array) return false;

                // 遍历条目比对
                foreach (var item in filesElement.EnumerateArray())
                {
                    // 提取 filePath 字段 (兼容大小写)
                    string itemFilePath = string.Empty;
                    // 读取 filePath、FilePath 或 path
                    if (item.TryGetProperty("filePath", out var fp) || item.TryGetProperty("FilePath", out fp) || item.TryGetProperty("path", out fp))
                    {
                        // 提取物理路径字符串
                        itemFilePath = fp.GetString() ?? string.Empty;
                    }

                    // 提取 fileName 字段 (兼容大小写)
                    string itemFileName = string.Empty;
                    // 读取 fileName、FileName 或 name
                    if (item.TryGetProperty("fileName", out var fn) || item.TryGetProperty("FileName", out fn) || item.TryGetProperty("name", out fn))
                    {
                        // 提取文件名字符串
                        itemFileName = fn.GetString() ?? string.Empty;
                    }

                    // 标准化 JSON 内部条目的路径格式
                    string normItemPath = itemFilePath.Replace('\\', '/').Trim().ToLowerInvariant();
                    // 标准化文件名格式
                    string normItemName = itemFileName.ToLowerInvariant();

                    // 规则 1：物理路径完全一致
                    if (!string.IsNullOrWhiteSpace(normItemPath) && normHostPath == normItemPath)
                    {
                        // 路径完全匹配
                        return true;
                    }

                    // 规则 2：相对路径与绝对路径相互包含匹配 (兼容相对路径存储机制)
                    if (!string.IsNullOrWhiteSpace(normItemPath) &&
                        (normHostPath.EndsWith("/" + normItemPath) || normHostPath.EndsWith(normItemPath) || normItemPath.EndsWith(normHostPath)))
                    {
                        // 相对路径后缀匹配命中
                        return true;
                    }

                    // 规则 3：文件名一致且同属于当前工程名称
                    if (!string.IsNullOrWhiteSpace(normItemName) && normItemName == hostFileName)
                    {
                        // 若同属于相同工程或路径中包含工程代号
                        if (!string.IsNullOrWhiteSpace(targetProjName) &&
                            (string.Equals(projName, targetProjName, StringComparison.OrdinalIgnoreCase) || normHostPath.Contains("/" + targetProjName.ToLowerInvariant() + "/")))
                        {
                            // 工程归属命中
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 单个文件读取异常记录日志并容错继续
                LogHelper.WriteLog($"[ExcelServices] 解析工程 JSON [{jsonPath}] 异常: {ex.Message}");
            }

            // 未命中返回 false
            return false;
        }

        #endregion

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

                // 读取 4 种名称前缀配置 (零堆分配)
                var (sumPrefix, detPrefix, subsumPrefix, tolsumPrefix) = CabinetPrefixConfig.Current;
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
