using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 本地文件管理控制台窗体 (基于 WebView2 + Vue 3 嵌入式视图)
    /// 用于在 Excel 内部无缝浏览、打开、定位本机项目图纸与清单
    /// </summary>
    public class LocalProjectForm : Form
    {
        // 内部嵌入的 Chromium WebView2 浏览器控件实例
        private readonly WebView2 _webView;

        /// <summary>
        /// 构造函数：初始化窗体几何属性与 WebView2 控件
        /// </summary>
        public LocalProjectForm()
        {
            // 实例化 WebView2 浏览器控件
            _webView = new WebView2();

            // 配置窗体标题、尺寸、边框样式及居中策略
            InitializeFormProperties();

            // 初始化 WebView2 控件布局并挂载至窗体控件集合
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基础属性 (支持用户拖拽缩放与最大化)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题文本
            this.Text = "本地文件管理控制台 - 扬州华科智能";

            // 初始视口尺寸设为 1280x800 像素，确保表格与筛选栏从容呈现
            this.ClientSize = new Size(1280, 800);

            // 限制最小窗体尺寸，防止过度缩小导致表格变形
            this.MinimumSize = new Size(960, 600);

            // 设置窗体初始在屏幕中央居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 设置标准可调整大小的窗体边框样式
            this.FormBorderStyle = FormBorderStyle.Sizable;

            // 开启最大化与最小化控制按钮
            this.MaximizeBox = true;
            this.MinimizeBox = true;

            // 任务栏显示窗体图标
            this.ShowInTaskbar = true;

            // 窗体背景设置为明亮白色
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件挂载与窗体加载事件
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 控件完全填充填充当前窗体客户区
            _webView.Dock = DockStyle.Fill;

            // 将浏览器控件添加至窗体控件树中
            this.Controls.Add(_webView);

            // 注册窗体 Load 生命周期异步初始化事件
            this.Load += OnFormLoadAsync;

            // 注册窗体关闭事件，安全释放浏览器上下文
            this.FormClosing += OnFormClosing;
        }

        /// <summary>
        /// 窗体加载异步事件：初始化 CoreWebView2 并导航至本地文件管理网页
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 校验窗体有效性，若已处于释放流程则提前退出
                if (this.IsDisposed || this.Disposing) return;

                // 计算 WebView2 用户数据缓存专用子目录
                string userDataFolder = Path.Combine(Tool.GetAppDataDirectory(), "WebView2_LocalProject");

                // 若缓存文件夹不存在则自动创建
                Directory.CreateDirectory(userDataFolder);

                // 创建独立的 CoreWebView2 环境句柄
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 二次校验窗体有效性
                if (this.IsDisposed || this.Disposing) return;

                // 初始化 CoreWebView2 核心引擎
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 禁用底部默认状态栏文本提示
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

                // 挂载前端 postMessage 消息接收监听器
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 从全局配置管理器中获取目标 Web 页面基准 URL
                string rawUrl = ConfigManager.Instance.Current.Api.LocalFileControlUrl;

                // 若配置为空则回退到官方线上嵌入式独立路由
                if (string.IsNullOrWhiteSpace(rawUrl))
                {
                    // 默认官方线上独立嵌入地址
                    rawUrl = "https://code.xingren.online/#/embed/fileCtrol";
                }

                // 读取当前 Excel 插件中已登录认证的 Token 凭证
                string token = ExcelServices.CurrentToken;

                // 💡 提取当前 Excel 活动工作簿的物理全路径与工程名称
                var (activeFilePath, activeProject) = GetActiveWorkbookInfo();

                // 构造带有 embed=excel 标识、免登 Token、物理路径与项目名称的目标访问 URL
                string targetUrl = BuildTargetUrl(rawUrl, token, activeFilePath, activeProject);

                // 导航加载目标前端单页应用
                _webView.Source = new Uri(targetUrl);
            }
            catch (Exception ex)
            {
                // 全局记录异常日志
                LogHelper.WriteLog($"[LocalProjectForm] 初始化加载 WebView2 异常: {ex.Message}");
                // 弹出异常提示对话框
                MessageBox.Show($"加载文件管理视图失败: {ex.Message}", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 组装带有 embed 参数、免登 token、物理文件全路径以及工程名称的完整 URL
        /// </summary>
        private string BuildTargetUrl(string baseUrl, string token, string filePath = "", string projectName = "")
        {
            // 若 URL 中未包含 embed 参数，智能拼接 embed=excel 参数
            string separator = baseUrl.Contains("?") ? "&" : "?";

            // 拼装 embed 嵌入标识参数
            string resultUrl = baseUrl.Contains("embed=") ? baseUrl : $"{baseUrl}{separator}embed=excel";

            // 若存在已登录的用户 Token，拼装至 URL 参数以实现跨端静默免登
            if (!string.IsNullOrWhiteSpace(token))
            {
                // 再次判断连接符
                separator = resultUrl.Contains("?") ? "&" : "?";
                // 附加安全编码后的 Token 凭证参数
                resultUrl = $"{resultUrl}{separator}token={Uri.EscapeDataString(token)}";
            }

            // 💡 优先拼装物理文件全路径参数 filePath (用于前端按基准目录+层级剥离匹配)
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                // 再次判断连接符
                separator = resultUrl.Contains("?") ? "&" : "?";
                // 附加安全编码后的物理文件路径参数
                resultUrl = $"{resultUrl}{separator}filePath={Uri.EscapeDataString(filePath)}";
            }

            // 💡 若当前活动工作簿具备明确项目名称，拼装 projectName 参数作为备用
            if (!string.IsNullOrWhiteSpace(projectName))
            {
                // 再次判断连接符
                separator = resultUrl.Contains("?") ? "&" : "?";
                // 附加安全编码后的项目工程名称参数
                resultUrl = $"{resultUrl}{separator}projectName={Uri.EscapeDataString(projectName)}";
            }

            // 返回最终的导航 URL 字符串
            return resultUrl;
        }

        /// <summary>
        /// 提取当前 Excel 活动工作簿的物理完整路径与项目名称
        /// </summary>
        /// <returns>包含 (filePath, projectName) 的元组</returns>
        public static (string filePath, string projectName) GetActiveWorkbookInfo()
        {
            try
            {
                // 获取当前 Excel 宿主应用程序实例
                dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;
                if (app == null) return (string.Empty, string.Empty);

                // 获取当前活动工作簿句柄
                dynamic? activeWb = null;
                try
                {
                    // 获取当前激活工作簿
                    activeWb = app.ActiveWorkbook;
                }
                catch
                {
                    // 忽略获取活动工作簿异常
                }

                // 若无有效打开的工作簿则返回空
                if (activeWb == null) return (string.Empty, string.Empty);

                // 提取活动工作簿的物理全路径
                string fullPath = string.Empty;
                try
                {
                    // 读取 FullName 属性
                    string rawPath = Convert.ToString(activeWb.FullName)?.Trim() ?? string.Empty;
                    // 确保是磁盘根路径（排除尚未存盘的“工作簿1”）
                    if (!string.IsNullOrWhiteSpace(rawPath) && Path.IsPathRooted(rawPath))
                    {
                        // 赋予有效路径
                        fullPath = rawPath;
                    }
                }
                catch
                {
                    // 忽略读取路径异常
                }

                // 💡 项目名称完全且唯一从物理路径中解析提取，彻底废除不可靠的 B5 单元格关联
                string projName = ExtractProjectNameFromPath(fullPath);

                // 记录提取调试日志
                LogHelper.WriteLog($"[LocalProjectForm] 提取当前工作簿结果: 路径=[{fullPath}], 工程=[{projName}]");

                // 返回提取到的物理全路径与工程名
                return (fullPath, projName);
            }
            catch (Exception ex)
            {
                // 记录提取异常日志
                LogHelper.WriteLog($"[LocalProjectForm] 提取活动工作簿信息异常: {ex.Message}");
            }

            // 兜底返回空值
            return (string.Empty, string.Empty);
        }

        /// <summary>
        /// 从文件物理全路径中智能提取真实工程代号 (优先按projectJsonFile实际文件比对，次选年份子目录层级提取)
        /// </summary>
        /// <param name="fullPath">物理文件绝对路径</param>
        /// <returns>提取出的工程代号 (如 GZ366)</returns>
        public static string ExtractProjectNameFromPath(string fullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath)) return string.Empty;

            try
            {
                // 拆分路径各级目录节点
                string[] rawParts = fullPath.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
                if (rawParts.Length <= 1) return string.Empty;

                // 转换为可变目录列表
                var dirParts = new List<string>(rawParts);
                // 移除末尾的文件名节点
                if (dirParts[dirParts.Count - 1].Contains("."))
                {
                    dirParts.RemoveAt(dirParts.Count - 1);
                }

                // 1. 优先策略：从当前路径逐级向上探测是否存在 projectJsonFile 工程目录
                string currentDir = Path.GetDirectoryName(fullPath) ?? string.Empty;
                string? projectJsonDir = null;
                DirectoryInfo? cur = new DirectoryInfo(currentDir);
                // 向上遍历至根驱动器
                while (cur != null)
                {
                    // 拼接探测 projectJsonFile
                    string candidate = Path.Combine(cur.FullName, "projectJsonFile");
                    if (Directory.Exists(candidate))
                    {
                        projectJsonDir = candidate;
                        break;
                    }
                    cur = cur.Parent;
                }

                // 若探测到 projectJsonFile 目录，从路径节点倒序比对是否存在对应的 {工程名}.json
                if (!string.IsNullOrWhiteSpace(projectJsonDir) && Directory.Exists(projectJsonDir))
                {
                    // 倒序比对各级目录名称
                    for (int i = dirParts.Count - 1; i >= 0; i--)
                    {
                        string seg = dirParts[i];
                        // 校验磁盘上是否存在对应的工程 JSON
                        if (File.Exists(Path.Combine(projectJsonDir, $"{seg}.json")))
                        {
                            // 100% 精确命中已存在的工程代号 (如 GZ366)
                            return seg;
                        }
                    }
                }

                // 2. 次选策略：识别标准成套结构 (形如: ...\2026\GZ366\...)，取 4 位年份后紧跟的目录
                for (int i = 0; i < dirParts.Count - 1; i++)
                {
                    // 匹配 4 位年份目录
                    if (System.Text.RegularExpressions.Regex.IsMatch(dirParts[i], @"^\d{4}$"))
                    {
                        string nextSeg = dirParts[i + 1];
                        // 排除纯数字日期目录
                        if (!System.Text.RegularExpressions.Regex.IsMatch(nextSeg, @"^\d{6,8}$"))
                        {
                            return nextSeg;
                        }
                    }
                }

                // 3. 兜底策略：倒序排除纯日期与纯盘符，返回最近一级的有效项目目录
                for (int i = dirParts.Count - 1; i >= 0; i--)
                {
                    string seg = dirParts[i];
                    // 排除盘符、4位年份与6-8位纯数字日期
                    if (seg.Contains(":") || System.Text.RegularExpressions.Regex.IsMatch(seg, @"^\d{4}$") || System.Text.RegularExpressions.Regex.IsMatch(seg, @"^\d{6,8}$"))
                        continue;
                    return seg;
                }
            }
            catch (Exception ex)
            {
                // 记录解析异常日志
                LogHelper.WriteLog($"[LocalProjectForm] 路径解析工程名异常: {ex.Message}");
            }

            // 未能提取返回空
            return string.Empty;
        }

        /// <summary>
        /// 智能提取当前 Excel 活动工作簿所对应的项目工程名称 (兼容旧接口)
        /// </summary>
        public static string GetActiveProjectName()
        {
            var (_, projName) = GetActiveWorkbookInfo();
            return projName;
        }

        /// <summary>
        /// 切换聚焦到指定工程项目 (通过 postMessage 通知内嵌 WebView2 页面)
        /// </summary>
        /// <param name="filePath">活动工作簿物理全路径</param>
        /// <param name="projectName">目标项目工程名称(可选)</param>
        public void SwitchToProject(string filePath, string projectName = "")
        {
            try
            {
                // 校验 CoreWebView2 是否已准备就绪且有路径或工程名
                if (_webView?.CoreWebView2 != null && (!string.IsNullOrWhiteSpace(filePath) || !string.IsNullOrWhiteSpace(projectName)))
                {
                    // 构造切换项目的消息体
                    var msg = new
                    {
                        type = "SELECT_PROJECT",
                        filePath = filePath,
                        projectName = projectName
                    };
                    // 序列化为 JSON 字符串
                    string json = JsonSerializer.Serialize(msg);
                    // 向前端页面发送 JSON 消息
                    _webView.CoreWebView2.PostWebMessageAsJson(json);
                    // 记录切换日志
                    LogHelper.WriteLog($"[LocalProjectForm] 已向前端派发聚焦项目消息: path={filePath}, proj={projectName}");
                }
            }
            catch (Exception ex)
            {
                // 记录派发消息异常
                LogHelper.WriteLog($"[LocalProjectForm] 向前端发送切换项目消息异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 前端 postMessage 消息接收与宿主原生动作分发处理
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 读取来自前端的原始 JSON 消息字符串
                string json = e.WebMessageAsJson;

                // 校验空字符串
                if (string.IsNullOrWhiteSpace(json)) return;

                // 解析 JSON 数据根元素
                using var doc = JsonDocument.Parse(json);
                // 获取根 JSON 元素
                var root = doc.RootElement;

                // 提取消息类型字段 (如 OPEN_FILE, LOCATE_FILE)
                string actionType = root.TryGetProperty("type", out var typeProp) ? typeProp.GetString() ?? "" : "";

                // 处理打开本地物理文件指令
                if (actionType == "OPEN_FILE")
                {
                    // 提取目标文件的绝对物理路径
                    string filePath = root.TryGetProperty("filePath", out var pathProp) ? pathProp.GetString() ?? "" : "";

                    // 若路径有效，调用文件打开处理方法
                    if (!string.IsNullOrWhiteSpace(filePath))
                    {
                        // 调度打开物理文件
                        HandleOpenFile(filePath);
                    }
                }
                // 处理在 Windows 资源管理器中定位文件/目录指令
                else if (actionType == "LOCATE_FILE")
                {
                    // 优先提取具体文件物理全路径
                    string filePath = root.TryGetProperty("filePath", out var pathProp) ? pathProp.GetString() ?? "" : "";

                    // 次选提取文件夹全路径
                    string folderPath = root.TryGetProperty("folderPath", out var folderProp) ? folderProp.GetString() ?? "" : "";

                    // 调用系统资源管理器进行文件高亮定位
                    HandleLocateFile(filePath, folderPath);
                }
            }
            catch (Exception ex)
            {
                // 捕获并记录处理 WebMessage 异常
                LogHelper.WriteLog($"[LocalProjectForm] 处理 WebMessageReceived 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 调度打开物理文件：若为 Excel 工作簿则优先在当前 Excel 进程中打开并激活，其余文件调用系统默认程序
        /// </summary>
        private void HandleOpenFile(string filePath)
        {
            try
            {
                // 校验物理文件是否存在于磁盘上
                if (!File.Exists(filePath))
                {
                    // 文件不存在时弹窗警示
                    MessageBox.Show($"目标物理文件不存在或已被移动：\n{filePath}", "文件未找到", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 获取文件小写扩展名
                string ext = Path.GetExtension(filePath).ToLowerInvariant();

                // 判断是否属于 Excel 关联工作簿类型
                bool isExcelFile = ext == ".xlsx" || ext == ".xls" || ext == ".xlsm" || ext == ".csv";

                // 若为 Excel 文件，在 Excel 主线程安全打开
                if (isExcelFile)
                {
                    // 调度 Excel 主线程宏安全打开该工作簿
                    OpenExcelWorkbookSafe(filePath);
                }
                else
                {
                    // 其他类型文件 (如 .dwg 图纸、.pdf 文档、.docx 文件) 调起 Windows 关联程序打开
                    Process.Start(new ProcessStartInfo
                    {
                        // 目标文件全路径
                        FileName = filePath,
                        // 启用 Windows Shell 执行引擎
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                // 记录打开文件失败日志
                LogHelper.WriteLog($"[LocalProjectForm] 打开物理文件失败 [{filePath}]: {ex.Message}");
                // 弹窗提示用户
                MessageBox.Show($"打开文件发生异常: {ex.Message}", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 在 Excel 主线程中安全打开工作簿并置顶激活窗口
        /// </summary>
        private void OpenExcelWorkbookSafe(string filePath)
        {
            try
            {
                // 使用 ExcelDna 的宏队列异步分发，避免跨 COM 线程调用导致的 RPC 死锁
                ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        // 获取当前 Excel 宿主 Application 对象
                        dynamic? app = ExcelDna.Integration.ExcelDnaUtil.Application;

                        // 若 Application 实例有效
                        if (app != null)
                        {
                            // 在当前 Excel 进程中打开该工作簿
                            dynamic wb = app.Workbooks.Open(filePath);

                            // 激活该打开的工作簿
                            wb?.Activate();

                            // 确保 Excel 主界面处于可见状态
                            app.Visible = true;
                        }
                    }
                    catch (Exception macroEx)
                    {
                        // 记录 COM 打开异常并自动回退为系统进程打开
                        LogHelper.WriteLog($"[LocalProjectForm] QueueAsMacro 打开工作簿异常，回退进程调起: {macroEx.Message}");
                        // 启动独立进程打开工作簿
                        Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
                    }
                });
            }
            catch (Exception ex)
            {
                // 记录排队异常并直接回退为系统进程打开
                LogHelper.WriteLog($"[LocalProjectForm] 派发 QueueAsMacro 失败，回退进程打开: {ex.Message}");
                // 进程方式打开
                Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
            }
        }

        /// <summary>
        /// 在 Windows 资源管理器中定位并选中文件
        /// </summary>
        private void HandleLocateFile(string filePath, string folderPath)
        {
            try
            {
                // 1. 若目标文件物理存在，直接定位并高亮选中该文件
                if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                {
                    // 启动 explorer.exe 携带 /select 参数
                    Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                    return;
                }

                // 2. 若目标为有效目录，直接打开该目录窗口
                string targetDir = !string.IsNullOrWhiteSpace(folderPath) ? folderPath : Path.GetDirectoryName(filePath) ?? "";

                // 校验目录物理是否存在
                if (!string.IsNullOrWhiteSpace(targetDir) && Directory.Exists(targetDir))
                {
                    // 打开目标目录窗口
                    Process.Start("explorer.exe", $"\"{targetDir}\"");
                }
                else
                {
                    // 目录不存在提示
                    MessageBox.Show($"目标路径不存在：\n{filePath}", "定位失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                // 记录定位异常日志
                LogHelper.WriteLog($"[LocalProjectForm] 定位文件异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 窗体关闭事件：安全释放 WebView2 核心对象
        /// </summary>
        private void OnFormClosing(object? sender, FormClosingEventArgs e)
        {
            try
            {
                // 注销消息接收监听器
                if (_webView.CoreWebView2 != null)
                {
                    // 移除 WebMessageReceived 委托
                    _webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
                }

                // 销毁并释放 WebView2 控件资源
                _webView.Dispose();
            }
            catch (Exception ex)
            {
                // 记录释放异常
                LogHelper.WriteLog($"[LocalProjectForm] 窗体关闭释放资源异常: {ex.Message}");
            }
        }
    }
}
