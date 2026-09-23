using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Controllers;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“云方案中心”无边框置顶宿主窗体
    /// 支持 4 列响应式网格布局、多图纸预览画廊与 BOM 动态回路倍增
    /// </summary>
    public class CloudSolutionForm : Form
    {
        // 声明 WebView2 浏览器控件
        private readonly WebView2 _webView;

        // 声明云方案业务控制器
        private readonly CloudSolutionController _controller;

        // 导入 Windows 原生 user32.dll 接口以支持无边框窗体拖拽
        [System.Runtime.InteropServices.DllImport("user32.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Winapi)]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 消息接口
        [System.Runtime.InteropServices.DllImport("user32.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Winapi)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 导入 GetAsyncKeyState 检测鼠标物理按键状态，防止幽灵拖拽死锁
        [System.Runtime.InteropServices.DllImport("user32.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Winapi)]
        private static extern short GetAsyncKeyState(int vKey);

        // Win32 常量: 鼠标左键虚拟键码
        private const int VK_LBUTTON = 0x01;

        // Win32 常量: 标题栏拖拽消息标识
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 通用 JSON 序列化设置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        /// <summary>
        /// 构造函数: 初始化控制器与 WebView2 控件属性
        /// </summary>
        public CloudSolutionForm()
        {
            // 实例化云方案业务控制器
            _controller = new CloudSolutionController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 设置 Form 基本尺寸与几何样式 (1180x820 像素，从容容纳 4 列卡片与图3详情)
            InitializeFormProperties();

            // 初始化并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 初始化窗体外观、尺寸与窗口属性
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题
            this.Text = "鑫壬云方案中心";

            // 获取用户当前主显示器的工作区有效尺寸 (去除任务栏)
            var workArea = Screen.PrimaryScreen.WorkingArea;
            // 默认加宽至 1460 像素，若屏幕较小则自适应取工作区 92% 宽度 --硬编码--
            int targetWidth = Math.Max(1200, Math.Min(1460, (int)(workArea.Width * 0.92)));
            // 高度适配为 880 像素，不超过工作区 90% 高度 --硬编码--
            int targetHeight = Math.Max(780, Math.Min(880, (int)(workArea.Height * 0.90)));
            // 应用加宽后的宽屏视口尺寸
            this.ClientSize = new Size(targetWidth, targetHeight);

            // 屏幕居中弹出展示
            this.StartPosition = FormStartPosition.CenterScreen;

            // 无边框工业现代样式
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用系统原生最大/最小化按钮 (由前端自定义胶囊控制)
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 背景色设置为白色
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件及核心事件回调
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 设置 WebView2 填满整个窗体
            _webView.Dock = DockStyle.Fill;

            // 添加到窗体控件树中
            this.Controls.Add(_webView);

            // 绑定窗口加载完成后初始化 CoreWebView2 事件
            this.Load += async (s, e) =>
            {
                try
                {
                    // 设置独立的本地用户缓存数据目录
                    string userDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "ExcelAddInDemo",
                        "WebView2_CloudSolution"
                    );

                    // 创建 CoreWebView2 运行环境
                    var env = await CoreWebView2Environment.CreateAsync(null, userDir);

                    // 异步初始化核心引擎
                    await _webView.EnsureCoreWebView2Async(env);

                    // 禁用内置右键菜单，保障界面整洁纯净
                    _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;

                    // 启用原生开发者工具（便于排查前端渲染）
                    _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                    // 绑定接收前端发来的 WebMessage 消息路由
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                    // 加载前端 HTML 页面
                    LoadHtmlPage();
                }
                catch (Exception ex)
                {
                    // 记录 WebView2 初始化异常
                    LogHelper.WriteLog($"[CloudSolutionForm] 初始化 WebView2 异常: {ex.Message}");
                    MessageBox.Show($"初始化方案中心界面失败: {ex.Message}", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
        }

        /// <summary>
        /// 寻址并导航加载 cloud_solution.html 前端页面
        /// 支持多重路径降级检测，避免在调试或独立部署路径下白屏
        /// </summary>
        private void LoadHtmlPage()
        {
            try
            {
                // 获取插件运行根目录
                string appDir = Tool.GetAppDirectory();
                // 获取当前应用域基准目录
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // 构造多级备选查找路径集合
                string[] candidatePaths = new[]
                {
                    Path.Combine(appDir, "Resources", "cloud_solution.html"),
                    Path.Combine(appDir, "publish", "Resources", "cloud_solution.html"),
                    Path.Combine(baseDir, "Resources", "cloud_solution.html"),
                    Path.Combine(baseDir, "publish", "Resources", "cloud_solution.html"),
                    Path.Combine(baseDir, "..", "..", "Resources", "cloud_solution.html"),
                    Path.Combine(baseDir, "..", "..", "..", "Resources", "cloud_solution.html"),
                    @"d:\code\excel-ct-tools\Resources\cloud_solution.html" // --硬编码: 本地开发源码目录兜底--
                };

                // 遍历寻找首个物理存在的 HTML 页面
                string targetHtml = string.Empty;
                foreach (var candidate in candidatePaths)
                {
                    if (File.Exists(candidate))
                    {
                        targetHtml = Path.GetFullPath(candidate);
                        break;
                    }
                }

                // 若找到有效 HTML 物理文件则执行虚拟主机映射并导航
                if (!string.IsNullOrEmpty(targetHtml) && File.Exists(targetHtml))
                {
                    // 提取资源根目录物理路径
                    string resDir = Path.GetDirectoryName(targetHtml)!;
                    // 将本地资源目录安全映射为 https://appassets.local (确保 WebWorker 与 WebAssembly 零跨域阻断)
                    _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        "appassets.local",
                        resDir,
                        Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.Allow);

                    // 导航至虚拟主机安全页面 (附加动态时间戳防缓存，确保每次打开均加载最新前端代码)
                    _webView.Source = new Uri($"https://appassets.local/cloud_solution.html?_t={DateTime.UtcNow.Ticks}");
                }
                else
                {
                    // 未找到物理文件时，在 WebView2 内部呈现友好错误页，彻底杜绝静默空白白板
                    string notFoundHtml = "<div style='font-family:Segoe UI,sans-serif;padding:30px;color:#dc2626;'>" +
                                          "<h2>⚠️ 未找到云方案页面文件 (cloud_solution.html)</h2>" +
                                          "<p>请确保 <code>Resources/cloud_solution.html</code> 存在并已复制至输出目录。</p>" +
                                          "<p>检索路径列表:<ul>" +
                                          string.Join("", Array.ConvertAll(candidatePaths, p => $"<li>{p}</li>")) +
                                          "</ul></p></div>";
                    _webView.CoreWebView2.NavigateToString(notFoundHtml);
                    // 记录未找到页面日志
                    LogHelper.WriteLog($"[CloudSolutionForm] 未找到 cloud_solution.html 页面文件，已呈现错误提示。");
                }
            }
            catch (Exception ex)
            {
                // 记录加载失败异常
                LogHelper.WriteLog($"[CloudSolutionForm] 加载 HTML 文件失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 前端 IPC 消息接收分发中心
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 读取收到的原始 JSON 文本
                string json = e.WebMessageAsJson;
                if (string.IsNullOrWhiteSpace(json)) return;

                // 解析为 JsonDocument 提取 action 字段
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;

                string action = actionProp.GetString() ?? "";

                // 兼容提取负载对象 (优先取 data 内部对象，若无则取 root 根节点本身)
                JsonElement dataEl = root.TryGetProperty("data", out var tmpData) && tmpData.ValueKind == JsonValueKind.Object
                    ? tmpData
                    : root;

                // 辅助安全读取字符串参数函数 (双重保障，杜绝因 payload 嵌套层级导致提取失败)
                string GetStringProp(string propName)
                {
                    if (dataEl.TryGetProperty(propName, out var p1) && p1.ValueKind == JsonValueKind.String)
                        return p1.GetString() ?? string.Empty;
                    if (root.TryGetProperty(propName, out var p2) && p2.ValueKind == JsonValueKind.String)
                        return p2.GetString() ?? string.Empty;
                    return string.Empty;
                }

                switch (action)
                {
                    // 1. 无边框拖拽移动窗体 (带物理鼠标状态检测，杜绝幽灵捕获死锁)
                    case "dragWindow":
                        if (this.WindowState == FormWindowState.Normal)
                        {
                            // 检测物理按键是否仍在按下状态
                            if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0)
                            {
                                ReleaseCapture();
                                SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                            }
                        }
                        break;

                    // 2. 关闭窗体
                    case "close":
                        this.Invoke(new Action(() => this.Close()));
                        break;

                    // 3. 最小化窗体
                    case "minimize":
                        this.Invoke(new Action(() => this.WindowState = FormWindowState.Minimized));
                        break;

                    // 4. 最大化与还原切换
                    case "toggleMaximize":
                        this.Invoke(new Action(() =>
                        {
                            this.WindowState = (this.WindowState == FormWindowState.Maximized)
                                ? FormWindowState.Normal
                                : FormWindowState.Maximized;
                        }));
                        break;

                    // 5. 分页多维检索方案列表
                    case "querySchemes":
                        string queryPayload = root.TryGetProperty("data", out var qProp) ? qProp.GetRawText() : "{}";
                        var result = _controller.QuerySchemes(queryPayload);
                        PostMessageSafe("querySchemesResult", result);
                        break;

                    // 6. 获取方案详情 (多图纸与 BOM)
                    case "getSchemeDetail":
                        string schemeId = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
                        var detail = _controller.GetSchemeDetail(schemeId);
                        PostMessageSafe("getSchemeDetailResult", detail);
                        break;

                    // 7. 切换收藏状态
                    case "toggleFavorite":
                        string favId = root.TryGetProperty("id", out var fIdProp) ? fIdProp.GetString() ?? "" : "";
                        bool isFav = _controller.ToggleFavorite(favId);
                        PostMessageSafe("toggleFavoriteResult", new { id = favId, isFavorite = isFav });
                        break;

                    // 8. 保存企业方案
                    case "saveEnterpriseScheme":
                        string schemePayload = root.TryGetProperty("data", out var sProp) ? sProp.GetRawText() : "{}";
                        bool saveOk = _controller.SaveEnterpriseScheme(schemePayload);
                        PostMessageSafe("saveEnterpriseSchemeResult", new { success = saveOk });
                        break;

                    // 9. 删除企业方案
                    case "deleteEnterpriseScheme":
                        string delId = root.TryGetProperty("id", out var dIdProp) ? dIdProp.GetString() ?? "" : "";
                        bool delOk = _controller.DeleteEnterpriseScheme(delId);
                        PostMessageSafe("deleteEnterpriseSchemeResult", new { success = delOk, id = delId });
                        break;

                    // 10. 插入方案到 Excel 活动工作表
                    case "insertSchemeToExcel":
                        string insertPayload = root.TryGetProperty("data", out var iProp) ? iProp.GetRawText() : "{}";
                        var (insertOk, insertMsg) = _controller.InsertSchemeToExcel(insertPayload);
                        PostMessageSafe("insertSchemeToExcelResult", new { success = insertOk, message = insertMsg });
                        break;

                    // 10.1 获取自动组价方案分类树
                    case "getAutoPricingCategoryTree":
                        var catTree = _controller.GetAutoPricingCategoryTree();
                        PostMessageSafe("getAutoPricingCategoryTreeResult", catTree);
                        break;

                    // 10.1.1 按需懒加载获取方案分类树子节点 (首屏省流，支持 el-tree lazy)
                    case "getAutoPricingCategoryNodes":
                        // 提取前端请求传递的父级节点 ID (0 为根目录)
                        string pId = GetStringProp("parentId");
                        // 调度控制器按需拉取当前层级的子分类或末级方案
                        var catNodes = _controller.GetAutoPricingCategoryNodes(pId);
                        // 将结果附带 parentId 回传给前端 WebView2 页面
                        PostMessageSafe("getAutoPricingCategoryNodesResult", new { parentId = pId, nodes = catNodes });
                        break;

                    // 10.2 检索自动组价方案列表
                    case "getAutoPricingSchemes":
                        {
                            string apCatId = GetStringProp("categoryId");
                            string apKw = GetStringProp("keyword");
                            string apCabMdl = GetStringProp("cabModel");
                            string apSit = GetStringProp("situation");
                            var schemes = _controller.GetAutoPricingSchemes(apCatId, apKw, apCabMdl, apSit);
                            PostMessageSafe("getAutoPricingSchemesResult", schemes);
                            break;
                        }

                    // 10.3 获取单个自动组价方案详情与 BOM 清单
                    case "getAutoPricingSchemeDetail":
                        string targetSchemeId = GetStringProp("schemeId");
                        if (string.IsNullOrWhiteSpace(targetSchemeId))
                        {
                            targetSchemeId = root.TryGetProperty("id", out var sidProp) ? sidProp.GetString() ?? "" : "";
                        }
                        var schDetail = _controller.GetAutoPricingSchemeDetail(targetSchemeId);
                        PostMessageSafe("getAutoPricingSchemeDetailResult", schDetail);
                        break;

                    // 11. 获取当前已保存的二次方案图纸根目录
                    case "getSecondaryCircuitConfig":
                        string currentSecDir = _controller.GetSecondaryCircuitDwgDir();
                        PostMessageSafe("getSecondaryCircuitConfigResult", new { rootDir = currentSecDir });
                        break;

                    // 12. 弹窗选择二次方案图纸根目录 (独立后台 STA 线程解耦，杜绝 Chromium IPC 模态死锁)
                    case "selectSecondaryCircuitDir":
                        string lastSecDir = _controller.GetSecondaryCircuitDwgDir();
                        var dialogThread = new System.Threading.Thread(() =>
                        {
                            try
                            {
                                // 创建目录浏览对话框 (注意: .NET Framework 4.8 中无 AutoUpgradeEnabled)
                                using var dialog = new FolderBrowserDialog
                                {
                                    Description = "请选择二次回路 DWG 图纸所在根目录",
                                    ShowNewFolderButton = true
                                };

                                // 恢复上一次记录的路径
                                if (!string.IsNullOrWhiteSpace(lastSecDir) && Directory.Exists(lastSecDir))
                                {
                                    dialog.SelectedPath = lastSecDir;
                                }

                                // 在独立 STA 模态消息循环中弹出
                                if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                                {
                                    string chosenDir = dialog.SelectedPath;
                                    // 保存并持久化配置
                                    _controller.SetSecondaryCircuitDwgDir(chosenDir);
                                    // 线程安全通知前端更新
                                    PostMessageSafe("selectSecondaryCircuitDirResult", new { success = true, path = chosenDir });
                                }
                            }
                            catch (Exception exDialog)
                            {
                                // 记录弹窗选择异常
                                LogHelper.WriteLog($"[CloudSolutionForm] selectSecondaryCircuitDir 线程异常: {exDialog.Message}");
                            }
                        });
                        dialogThread.SetApartmentState(System.Threading.ApartmentState.STA);
                        dialogThread.IsBackground = true;
                        dialogThread.Start();
                        break;

                    // 13. 手动粘贴或输入二次方案图纸根目录
                    case "setSecondaryCircuitDirManual":
                        string manualPath = GetStringProp("path");
                        if (!string.IsNullOrWhiteSpace(manualPath) && Directory.Exists(manualPath))
                        {
                            _controller.SetSecondaryCircuitDwgDir(manualPath);
                            PostMessageSafe("selectSecondaryCircuitDirResult", new { success = true, path = manualPath });
                        }
                        else
                        {
                            PostMessageSafe("selectSecondaryCircuitDirResult", new { success = false, message = "指定的文件夹路径在本地不存在，请检查后重试！" });
                        }
                        break;

                    // 14. 扫描二次方案根目录下的子文件夹列表
                    case "scanSecondaryFolders":
                        string scanRootDir = GetStringProp("rootDir");
                        var folders = _controller.ScanSecondaryFolders(string.IsNullOrEmpty(scanRootDir) ? null : scanRootDir);
                        PostMessageSafe("scanSecondaryFoldersResult", folders);
                        break;

                    // 15. 获取指定子文件夹下的 DWG 卡片列表 (含数据库参数、缩略图与强制刷新控制)
                    case "getFolderDwgCards":
                        string fPath = GetStringProp("folderPath");
                        string fName = GetStringProp("folderName");
                        string kw = GetStringProp("keyword");
                        bool forceRefresh = false;
                        // 尝试从顶层或 data 对象中提取 forceRefresh 布尔标志
                        if (dataEl.TryGetProperty("forceRefresh", out var frEl) ||
                            (dataEl.TryGetProperty("data", out var dInner) && dInner.TryGetProperty("forceRefresh", out frEl)))
                        {
                            try { forceRefresh = frEl.GetBoolean(); } catch { }
                        }
                        var cards = _controller.GetFolderDwgCards(fPath, fName, string.IsNullOrEmpty(kw) ? null : kw, forceRefresh);
                        PostMessageSafe("getFolderDwgCardsResult", cards);
                        break;

                    // 16. 调用默认关联的 AutoCAD 打开该 DWG 文件
                    case "openDwgInCad":
                        string dwgPath = GetStringProp("fullPath");
                        var (cadOk, cadMsg) = _controller.OpenDwgInCad(dwgPath);
                        PostMessageSafe("openDwgInCadResult", new { success = cadOk, message = cadMsg });
                        break;

                    // 16.1 激活当前运行的 AutoCAD 并提示指定位置以 1:1 比例插入图块
                    case "insertDwgToCad":
                        {
                            // 获取前端传入的 DWG 物理图纸全路径
                            string insertDwgPath = GetStringProp("fullPath");
                            // 调用控制器执行激活 CAD 并调度插入
                            var (insertCadOk, insertCadMsg) = _controller.InsertDwgToActiveCad(insertDwgPath);
                            // 向前端回传处理结果与提示消息
                            PostMessageSafe("insertDwgToCadResult", new { success = insertCadOk, message = insertCadMsg });
                            break;
                        }

                    // 17. 保存/更新二次回路方案
                    case "saveSecondaryScheme":
                        string secSchemePayload = "{}";
                        if (dataEl.TryGetProperty("data", out var innerPayload))
                        {
                            secSchemePayload = innerPayload.GetRawText();
                        }
                        else if (root.TryGetProperty("data", out var rootPayload))
                        {
                            secSchemePayload = rootPayload.GetRawText();
                        }
                        var (secOk, secId, secMsg) = _controller.SaveSecondaryScheme(secSchemePayload);
                        PostMessageSafe("saveSecondarySchemeResult", new { success = secOk, schemeId = secId, message = secMsg });
                        break;

                    // 18. 搜索物料库 (供编辑弹窗中的物料选择器使用)
                    case "searchMaterialComponents":
                        string mKw = GetStringProp("keyword");
                        var compList = _controller.SearchMaterialComponents(mKw);
                        PostMessageSafe("searchMaterialComponentsResult", compList);
                        break;

                    // 19. 获取所有二次回路方案列表 (供“复制其他方案”选择使用)
                    case "getSecondarySchemesForCopy":
                        string copyKw = GetStringProp("keyword");
                        var schemesForCopy = _controller.GetSecondarySchemesForCopy(copyKw);
                        PostMessageSafe("getSecondarySchemesForCopyResult", schemesForCopy);
                        break;

                    // 20. 提取 DWG 原始二进制 Base64 数据供前端 cad-view 矢量视口渲染
                    case "getDwgBinary":
                        string binPath = GetStringProp("filePath");
                        if (!string.IsNullOrWhiteSpace(binPath))
                        {
                            // 调度控制器读取二进制流
                            var binData = _controller.GetDwgBinary(binPath);
                            // 将数据回传前端 WebView2
                            PostMessageSafe("dwgBinaryLoaded", binData);
                        }
                        break;

                    // 21. 获取当前已保存的一次方案图纸根目录
                    case "getPrimaryCircuitConfig":
                        string currentPrimDir = _controller.GetPrimaryCircuitDwgDir();
                        PostMessageSafe("getPrimaryCircuitConfigResult", new { rootDir = currentPrimDir });
                        break;

                    // 22. 弹窗选择一次方案图纸根目录 (独立后台 STA 线程解耦，杜绝 Chromium IPC 模态死锁)
                    case "selectPrimaryCircuitDir":
                        string lastPrimDir = _controller.GetPrimaryCircuitDwgDir();
                        var primDialogThread = new System.Threading.Thread(() =>
                        {
                            try
                            {
                                using var dialog = new FolderBrowserDialog
                                {
                                    Description = "请选择一次方案 DWG 图纸所在根目录",
                                    ShowNewFolderButton = true
                                };

                                if (!string.IsNullOrWhiteSpace(lastPrimDir) && Directory.Exists(lastPrimDir))
                                {
                                    dialog.SelectedPath = lastPrimDir;
                                }

                                if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                                {
                                    string chosenDir = dialog.SelectedPath;
                                    _controller.SetPrimaryCircuitDwgDir(chosenDir);
                                    PostMessageSafe("selectPrimaryCircuitDirResult", new { success = true, path = chosenDir });
                                }
                            }
                            catch (Exception exDialog)
                            {
                                LogHelper.WriteLog($"[CloudSolutionForm] selectPrimaryCircuitDir 线程异常: {exDialog.Message}");
                            }
                        });
                        primDialogThread.SetApartmentState(System.Threading.ApartmentState.STA);
                        primDialogThread.IsBackground = true;
                        primDialogThread.Start();
                        break;

                    // 23. 手动粘贴或输入一次方案图纸根目录
                    case "setPrimaryCircuitDirManual":
                        string manualPrimPath = GetStringProp("path");
                        if (!string.IsNullOrWhiteSpace(manualPrimPath) && Directory.Exists(manualPrimPath))
                        {
                            _controller.SetPrimaryCircuitDwgDir(manualPrimPath);
                            PostMessageSafe("selectPrimaryCircuitDirResult", new { success = true, path = manualPrimPath });
                        }
                        else
                        {
                            PostMessageSafe("selectPrimaryCircuitDirResult", new { success = false, message = "指定的一次图纸目录在本地不存在，请检查后重试！" });
                        }
                        break;

                    // 24. 扫描一次方案根目录下的子文件夹列表
                    case "scanPrimaryFolders":
                        string scanPrimRootDir = GetStringProp("rootDir");
                        var primFolders = _controller.ScanPrimaryFolders(string.IsNullOrEmpty(scanPrimRootDir) ? null : scanPrimRootDir);
                        PostMessageSafe("scanPrimaryFoldersResult", primFolders);
                        break;

                    // 25. 获取指定子文件夹下的一次 DWG 卡片列表
                    case "getPrimaryFolderDwgCards":
                        string pPath = GetStringProp("folderPath");
                        string pName = GetStringProp("folderName");
                        string pKw = GetStringProp("keyword");
                        string pModel = GetStringProp("cabinetModel");
                        bool pForceRefresh = false;
                        if (dataEl.TryGetProperty("forceRefresh", out var pfrEl) ||
                            (dataEl.TryGetProperty("data", out var pdInner) && pdInner.TryGetProperty("forceRefresh", out pfrEl)))
                        {
                            try { pForceRefresh = pfrEl.GetBoolean(); } catch { }
                        }
                        var primCards = _controller.GetPrimaryFolderDwgCards(pPath, pName, string.IsNullOrEmpty(pKw) ? null : pKw, string.IsNullOrEmpty(pModel) ? null : pModel, pForceRefresh);
                        PostMessageSafe("getPrimaryFolderDwgCardsResult", primCards);
                        break;

                    // 26. 保存/更新一次成套方案
                    case "savePrimaryScheme":
                        string primSchemePayload = "{}";
                        if (dataEl.TryGetProperty("data", out var innerPrimPayload))
                        {
                            primSchemePayload = innerPrimPayload.GetRawText();
                        }
                        else if (root.TryGetProperty("data", out var rootPrimPayload))
                        {
                            primSchemePayload = rootPrimPayload.GetRawText();
                        }
                        var (primOk, primId, primMsg) = _controller.SavePrimaryScheme(primSchemePayload);
                        PostMessageSafe("savePrimarySchemeResult", new { success = primOk, schemeId = primId, message = primMsg });
                        break;

                    // 27. 获取所有一次方案列表 (供“复制其他方案”选择使用)
                    case "getPrimarySchemesForCopy":
                        string pCopyKw = GetStringProp("keyword");
                        var primSchemesForCopy = _controller.GetPrimarySchemesForCopy(pCopyKw);
                        PostMessageSafe("getPrimarySchemesForCopyResult", primSchemesForCopy);
                        break;

                    // 28. 抓取活动 Excel 当前光标箱柜为一次企业方案
                    case "captureActiveCabinetToPrimaryScheme":
                        string capFolder = GetStringProp("folderName");
                        string capSchemeName = GetStringProp("schemeName");
                        string capCabModel = GetStringProp("cabinetModel");
                        var (capOk, capMsg, capScheme) = _controller.CaptureActiveCabinetToPrimaryScheme(capFolder, capSchemeName, capCabModel);
                        PostMessageSafe("captureActiveCabinetToPrimarySchemeResult", new { success = capOk, message = capMsg, scheme = capScheme });
                        break;
                }
            }
            catch (Exception ex)
            {
                // 记录消息分发错误
                LogHelper.WriteLog($"[CloudSolutionForm] 消息处理异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 线程安全地在 UI 主线程执行委托 (严格遵循 Local Heuristics 规范 10)
        /// </summary>
        private void SafeInvoke(Action action)
        {
            // 防御性校验当前窗体是否已被释放或句柄尚未创建
            if (this.IsDisposed || !this.IsHandleCreated) return;

            // 检测当前调用是否需要跨线程切回 UI 调度
            if (this.InvokeRequired)
            {
                // 异步将委托推入 UI 主线程消息循环执行
                this.BeginInvoke(action);
            }
            else
            {
                // 若已处于主线程直接同步调用
                action();
            }
        }

        /// <summary>
        /// 线程安全地向前端 WebView2 发送 JSON 格式的回调消息
        /// 严格在 UI 主线程中校验与调用 CoreWebView2，杜绝 InvalidOperationException
        /// </summary>
        private void PostMessageSafe(string action, object? data)
        {
            // 切回 UI 主线程后再执行组件检查与消息发送
            SafeInvoke(() =>
            {
                try
                {
                    // 在主线程中安全校验 WebView2 控件及其内核实例生命周期
                    if (this.IsDisposed || _webView.IsDisposed || _webView.CoreWebView2 == null) return;

                    // 封装统一的回调信封实体
                    var envelope = new
                    {
                        action = action,
                        data = data
                    };

                    // 按照指定选项序列化为 JSON 字符串
                    string json = JsonSerializer.Serialize(envelope, JsonOptions);
                    // 安全投递给前端 JavaScript 监听器
                    _webView.CoreWebView2.PostWebMessageAsJson(json);
                }
                catch (Exception ex)
                {
                    // 记录消息投递异常日志
                    LogHelper.WriteLog($"[CloudSolutionForm] 投递消息失败: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 窗体关闭时解绑 WebMessage 事件并释放 WebView2 控件资源
        /// 杜绝后台 Chromium 子进程僵死残留
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 执行基类关闭生命周期逻辑
            base.OnFormClosing(e);

            try
            {
                // 安全解绑前后端消息接收事件处理器
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
                }

                // 显式释放并销毁 WebView2 控件
                _webView.Dispose();
            }
            catch (Exception ex)
            {
                // 记录释放异常日志
                LogHelper.WriteLog($"[CloudSolutionForm] 释放 WebView2 异常: {ex.Message}");
            }
        }
    }
}
