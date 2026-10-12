using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Services;
using ExcelAddInDemo.Models;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的元器件图纸参数匹配侧边浮窗 (200x800)
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，配置与硬编码显式标明
    /// </summary>
    public class ComponentParamMatchForm : Form
    {
        // 全局单例实例句柄
        private static ComponentParamMatchForm? _instance;

        // WebView2 浏览器控件实例
        private readonly WebView2 _webView;

        // JSON 序列化选项 (驼峰命名与宽松解析)
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        #region Windows API 窗口无边框原生拖拽与调整大小支持

        // 释放当前线程鼠标捕获状态
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 向指定窗口句柄发送 Win32 消息
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 异步获取物理按键按下状态，防止消息延迟导致模态死锁
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        // 物理鼠标左键虚拟键码
        private const int VK_LBUTTON = 0x01;
        // 非客户区鼠标左键按下消息常量
        private const int WM_NCLBUTTONDOWN = 0xA1;
        // 标题栏命中代码常量
        private const int HTCAPTION = 0x2;
        // 系统命令消息常量
        private const int WM_SYSCOMMAND = 0x0112;
        // 系统级调整窗口大小基准命令常量
        private const int SC_SIZE = 0xF000;

        #endregion

        /// <summary>
        /// 私有构造函数：配置 200x800 几何尺寸、置顶与无边框属性
        /// </summary>
        private ComponentParamMatchForm()
        {
            // 初始化 WebView2 控件实例
            _webView = new WebView2();

            // 配置窗体几何外观与尺寸
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体无边框、200x800 像素尺寸与屏幕右侧靠边定位
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗口标题栏文字
            this.Text = "元器件图纸参数匹配";

            // 设定初始窗口规格 (200x800) --硬编码: 窗口尺寸 200x800--
            this.ClientSize = new Size(200, 800);

            // 设定最小窗口尺寸限制，防止过度压缩破坏布局
            this.MinimumSize = new Size(180, 300); // --硬编码: 最小尺寸 180x300--

            // 彻底去除原生系统边框，由 Vue3 前端绘制自研精致顶栏
            this.FormBorderStyle = FormBorderStyle.None;

            // 彻底去除任何边距，确保前端顶栏紧密贴合窗口顶部
            this.Padding = new Padding(0);

            // 背景色设为纯白
            this.BackColor = Color.White;

            // 始终置顶显示，保障用户在 Excel 单元格切换时窗口不被遮挡
            this.TopMost = true;

            // 不在任务栏显示独立图标
            this.ShowInTaskbar = false;

            // 启用手动绝对坐标定位
            this.StartPosition = FormStartPosition.Manual;

            // 获取主屏幕工作区矩形，默认贴靠主屏幕最右侧
            var workingArea = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
            int posX = Math.Max(0, workingArea.Right - 205);
            int posY = Math.Max(20, (workingArea.Height - 800) / 2);
            this.Location = new Point(posX, posY);
        }

        /// <summary>
        /// 初始化 WebView2 控件并绑定加载事件
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 控件完全填充满窗体
            _webView.Dock = DockStyle.Fill;

            // 添加到 WinForms 控件集合
            this.Controls.Add(_webView);

            // 订阅窗体加载异步事件
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载事件：初始化 WebView2 环境并导航至 component_param_match.html
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 1. 获取系统本地 LocalApplicationData 专有缓存目录，避免污染用户业务配置或网盘同步
                string userDataFolder = Tool.GetWebView2UserDataFolder("WebView2_ParamMatch");
                // 2. 异步创建 WebView2 运行时环境
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                // 3. 确保控件初始化就绪
                await _webView.EnsureCoreWebView2Async(env);

                // 禁用默认右键菜单防止出现 Chromium 默认菜单
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                // 禁用状态栏显示
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

                // 注册 Web 消息接收事件监听
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 获取 Resources 资源物理文件夹 (优先检查输出目录，若无文件则回退基准目录)
                string resDir = Path.Combine(Tool.GetAppDirectory(), "Resources");
                if (!Directory.Exists(resDir) || !File.Exists(Path.Combine(resDir, "component_param_match.html")))
                {
                    resDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources");
                }

                // 映射本地虚拟 HTTPS 域名，解除本地静态资源安全限制
                _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "appassets.local",
                    resDir,
                    CoreWebView2HostResourceAccessKind.Allow
                );

                // 安全平滑导航至图纸匹配页面
                _webView.CoreWebView2.Navigate("https://appassets.local/component_param_match.html");
            }
            catch (Exception ex)
            {
                // 记录初始化异常日志
                LogHelper.WriteLog($"[ComponentParamMatchForm] 初始化异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 接收并路由处理来自前端 Vue 3 的 IPC 交互消息
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 读取前端传递的消息字符串
                string rawJson = "";
                try { rawJson = e.TryGetWebMessageAsString(); } catch { }
                if (string.IsNullOrWhiteSpace(rawJson))
                {
                    try { rawJson = e.WebMessageAsJson; } catch { }
                }
                if (string.IsNullOrWhiteSpace(rawJson)) return;

                // 解析为 JsonDocument 结构
                using var doc = JsonDocument.Parse(rawJson);
                var root = doc.RootElement;
                string action = root.TryGetProperty("action", out var actProp) ? (actProp.GetString() ?? "") : "";

                // 1. 初始化数据请求
                if (action == "getInitialData")
                {
                    var cfg = ConfigManager.Instance.Current?.ComponentParamMatch ?? new ComponentParamMatchSettings();
                    string baseDir = cfg.BaseDirectory;
                    // 扫描当前根目录层级
                    var scanResult = DwgPreviewService.ScanDirectoryHierarchy(baseDir);

                    // 获取当前 Excel 活动工作表名称并自适应决定目标列
                    string activeSheetName = string.Empty;
                    string effectiveColDwg = "AA";
                    string effectiveColDir = "AB";
                    try
                    {
                        dynamic? app = ExcelDnaUtil.Application;
                        if (app?.ActiveSheet != null)
                        {
                            activeSheetName = Convert.ToString(app.ActiveSheet.Name)?.Trim() ?? string.Empty;
                            if (string.Equals(activeSheetName, "元件汇总表", StringComparison.OrdinalIgnoreCase))
                            {
                                effectiveColDwg = "X";
                                effectiveColDir = "Y";
                            }
                            else
                            {
                                effectiveColDwg = cfg.TargetDwgColumn == "X" ? "AA" : cfg.TargetDwgColumn;
                                effectiveColDir = cfg.TargetDirColumn == "Y" ? "AB" : cfg.TargetDirColumn;
                            }
                        }
                    }
                    catch { }

                    // 回发初始配置与根目录扫描数据
                    PostWebMessageSafe(JsonSerializer.Serialize(new
                    {
                        action = "initialDataLoaded",
                        baseDirectory = baseDir,
                        activeSheetName = activeSheetName,
                        targetDirColumn = effectiveColDir,
                        targetDwgColumn = effectiveColDwg,
                        autoNextRow = cfg.AutoNextRow,
                        removeExtension = cfg.RemoveExtension,
                        scanData = scanResult
                    }, JsonOptions));

                    // 异步触发未录入尺寸图纸的自动测算补全
                    TriggerAsyncDimensionAutoFill(scanResult);
                }
                // 2. 点击设置图标选择根物理目录
                else if (action == "selectBaseDirectory")
                {
                    // 在专用 STA 线程中弹出文件夹选择器对话框
                    var dialogThread = new System.Threading.Thread(() =>
                    {
                        try
                        {
                            var cfg = ConfigManager.Instance.Current?.ComponentParamMatch ?? new ComponentParamMatchSettings();
                            string curDir = cfg.BaseDirectory;
                            using var fbd = new FolderBrowserDialog
                            {
                                Description = "请选择元器件图纸库根物理目录",
                                ShowNewFolderButton = false,
                                SelectedPath = Directory.Exists(curDir) ? curDir : string.Empty
                            };

                            if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                            {
                                string chosenPath = fbd.SelectedPath;
                                // 异步更新配置并持久化至磁盘
                                ConfigManager.Instance.UpdateComponentParamMatchBaseDirectory(chosenPath);
                                // 扫描新选中的目录层级
                                var newScanResult = DwgPreviewService.ScanDirectoryHierarchy(chosenPath);

                                // 回发目录变更成功通知
                                PostWebMessageSafe(JsonSerializer.Serialize(new
                                {
                                    action = "baseDirectoryChanged",
                                    baseDirectory = chosenPath,
                                    scanData = newScanResult
                                }, JsonOptions));

                                // 异步触发未录入尺寸图纸的自动测算补全
                                TriggerAsyncDimensionAutoFill(newScanResult);
                            }
                        }
                        catch (Exception ex)
                        {
                            LogHelper.WriteLog($"[ComponentParamMatchForm] selectBaseDirectory 异常: {ex.Message}");
                        }
                    });
                    dialogThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    dialogThread.IsBackground = true;
                    dialogThread.Start();
                }
                // 3. 扫描指定子目录 (下钻或返回上级)
                else if (action == "scanDirectory")
                {
                    string targetPath = root.TryGetProperty("path", out var pProp) ? (pProp.GetString() ?? "") : "";
                    if (!string.IsNullOrWhiteSpace(targetPath) && Directory.Exists(targetPath))
                    {
                        // 执行层级扫描
                        var scanResult = DwgPreviewService.ScanDirectoryHierarchy(targetPath);
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "directoryScanned",
                            scanData = scanResult
                        }, JsonOptions));

                        // 异步触发未录入尺寸图纸的自动测算补全
                        TriggerAsyncDimensionAutoFill(scanResult);
                    }
                }
                // 3.1 智能目录匹配：根据当前活动单元格所在行元器件名称自动进入最匹配目录
                else if (action == "autoMatchDirectory")
                {
                    // 在 Excel 主线程宏队列中安全读取当前行元器件名称并执行目录匹配
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            var cfg = ConfigManager.Instance.Current?.ComponentParamMatch ?? new ComponentParamMatchSettings();
                            string baseDir = cfg.BaseDirectory;
                            // 安全提取当前活动行元器件名称 (优先 B 列)
                            string compName = ExcelServices.GetActiveRowComponentName();

                            // 调度算法计算最匹配的一级分类目录
                            var matchDir = DwgPreviewService.FindBestMatchingDirectory(baseDir, compName);
                            string targetPath;
                            bool isMatched = false;
                            string matchedDirName = string.Empty;

                            if (matchDir != null && !string.IsNullOrWhiteSpace(matchDir.FullPath) && Directory.Exists(matchDir.FullPath))
                            {
                                // 成功命中目标分类子目录
                                targetPath = matchDir.FullPath;
                                isMatched = true;
                                matchedDirName = matchDir.Name;
                            }
                            else
                            {
                                // 用户决策指示：若未匹配到目录则自动返回图纸库根物理目录
                                targetPath = Directory.Exists(baseDir) ? baseDir : string.Empty;
                                isMatched = false;
                            }

                            // 扫描该目标目录下的子文件夹与图纸文件
                            var scanResult = DwgPreviewService.ScanDirectoryHierarchy(targetPath);

                            // 回发智能匹配执行结果通知前端 Vue3
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "autoMatchResult",
                                success = isMatched,
                                componentName = compName,
                                matchedDirName = matchedDirName,
                                scanData = scanResult
                            }, JsonOptions));

                            // 异步触发未录入尺寸图纸的自动测算补全
                            TriggerAsyncDimensionAutoFill(scanResult);
                        }
                        catch (Exception exMatch)
                        {
                            // 记录智能匹配异常日志
                            LogHelper.WriteLog($"[ComponentParamMatchForm] autoMatchDirectory 异常: {exMatch.Message}");
                        }
                    });
                }
                // 3.2 一键主动补测当前目录下图纸尺寸并存入物料库
                else if (action == "forceExtractDimensions")
                {
                    // 提取目标目录路径
                    string targetPath = root.TryGetProperty("path", out var pProp) ? (pProp.GetString() ?? "") : "";
                    // 校验目标目录物理有效性
                    if (!string.IsNullOrWhiteSpace(targetPath) && Directory.Exists(targetPath))
                    {
                        // 扫描获取当前目录下的文件列表
                        var scanResult = DwgPreviewService.ScanDirectoryHierarchy(targetPath);
                        if (scanResult?.DwgFiles != null && scanResult.DwgFiles.Count > 0)
                        {
                            // 优先筛选未录入尺寸且物理存在的图纸，若都已有尺寸则全量重新校验测算
                            var missingDwgs = scanResult.DwgFiles.FindAll(d => !d.HasDimension && !string.IsNullOrWhiteSpace(d.FullPath) && File.Exists(d.FullPath));
                            var targetDwgs = missingDwgs.Count > 0 ? missingDwgs : scanResult.DwgFiles.FindAll(d => !string.IsNullOrWhiteSpace(d.FullPath) && File.Exists(d.FullPath));
                            // 获取目标图纸物理路径列表
                            var targetPaths = targetDwgs.ConvertAll(d => d.FullPath);

                            // 启动非阻塞后台线程池任务发起测算
                            System.Threading.Tasks.Task.Run(async () =>
                            {
                                try
                                {
                                    // 异步请求 AutoCAD 管道服务端批量测算图纸尺寸并持久化至 SQLite
                                    var extractedDims = await CadSyncClient.RequestExtractDwgDimensionsAsync(targetPaths);
                                    // 构造响应结果回传前端 Vue 3 界面
                                    PostWebMessageSafe(JsonSerializer.Serialize(new
                                    {
                                        action = "forceExtractDimensionsResult",
                                        success = extractedDims != null && extractedDims.Count > 0,
                                        currentPath = targetPath,
                                        count = extractedDims?.Count ?? 0,
                                        items = extractedDims ?? new List<DwgDimensionItem>(),
                                        message = (extractedDims != null && extractedDims.Count > 0) ? $"成功测算 {extractedDims.Count} 张图纸尺寸" : "CAD 未能提取到尺寸，请检查图纸或确认 AutoCAD 是否已就绪"
                                    }, JsonOptions));
                                }
                                catch (Exception exExtract)
                                {
                                    // 记录测算异常并通知前端
                                    LogHelper.WriteLog($"[ComponentParamMatchForm] forceExtractDimensions 异常: {exExtract.Message}");
                                    PostWebMessageSafe(JsonSerializer.Serialize(new
                                    {
                                        action = "forceExtractDimensionsResult",
                                        success = false,
                                        message = $"测算异常: {exExtract.Message}"
                                    }, JsonOptions));
                                }
                            });
                        }
                        else
                        {
                            // 提示当前目录下无可测算的 DWG 文件
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "forceExtractDimensionsResult",
                                success = false,
                                message = "当前目录下未找到任何 DWG 图纸文件"
                            }, JsonOptions));
                        }
                    }
                }
                // 4. 双击 DWG 图纸文件：回写 Excel X/Y 列并自动跳向下一行
                else if (action == "bindDwg")
                {
                    string dirName = root.TryGetProperty("dirName", out var dnProp) ? (dnProp.GetString() ?? "") : "";
                    string dwgName = root.TryGetProperty("dwgName", out var dwgProp) ? (dwgProp.GetString() ?? "") : "";

                    // 脱离 WebView2 回调栈，派发至 Excel 纯净宏队列中执行 COM 单元格写入
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        var cfg = ConfigManager.Instance.Current?.ComponentParamMatch ?? new ComponentParamMatchSettings();
                        var result = ExcelServices.BindDwgParamToActiveRow(
                            dirName,
                            dwgName,
                            cfg.AutoNextRow,
                            cfg.TargetDirColumn,
                            cfg.TargetDwgColumn,
                            cfg.RemoveExtension
                        );

                        // 尝试从 SQLite 检索当前图纸的三维外形与进深尺寸
                        string dimStr = string.Empty;
                        try
                        {
                            // 优先通过 目录名/图纸名 组合检索
                            string lookupKey = string.IsNullOrWhiteSpace(dirName) ? dwgName : $"{dirName}/{dwgName}";
                            var dimItem = PersonalComponentDbService.GetDwgDimension(lookupKey);
                            if (dimItem != null && dimItem.Width > 0 && dimItem.Height > 0)
                            {
                                // 拼接友好展示文本
                                dimStr = $"{dimItem.Width}×{dimItem.Height}" + (dimItem.Depth > 0 ? $"×{dimItem.Depth}" : "");
                            }
                        }
                        catch { }

                        // 向前端回发写入结果反馈 (包含生效的列标与工作表名称)
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "bindDwgResult",
                            success = result.Success,
                            message = result.Message,
                            nextRow = result.NextRow,
                            colDwg = result.ColDwg,
                            colDir = result.ColDir,
                            sheetName = result.SheetName,
                            dirName = dirName,
                            dwgName = dwgName,
                            dimStr = dimStr
                        }, JsonOptions));
                    });
                }
                // 5. 鼠标按住顶栏系统原生拖拽移动窗口
                else if (action == "dragWindow")
                {
                    SafeInvoke(() =>
                    {
                        // 校验鼠标左键物理状态，若已松开则不触发系统拖动
                        if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) == 0) return;
                        // 释放当前鼠标捕获
                        ReleaseCapture();
                        // 触发系统原生标题栏平滑拖动
                        SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                    });
                }
                // 5.1 兼容旧版基于坐标差值的窗口位移指令
                else if (action == "moveWindow")
                {
                    int deltaX = root.TryGetProperty("deltaX", out var dxP) ? dxP.GetInt32() : 0;
                    int deltaY = root.TryGetProperty("deltaY", out var dyP) ? dyP.GetInt32() : 0;
                    SafeInvoke(() =>
                    {
                        this.Location = new Point(this.Location.X + deltaX, this.Location.Y + deltaY);
                    });
                }
                // 5.2 边缘透明感应区触发系统原生平滑调整窗口大小 (零坐标计算，系统级 60/120 帧丝滑响应)
                else if (action == "startResize")
                {
                    // 提取拖拽边缘方向标识
                    string dirStr = root.TryGetProperty("direction", out var dP) ? (dP.GetString() ?? "") : "right";
                    // 映射为 Win32 SC_SIZE 对应的方向代码 (1:左, 2:右, 3:顶, 4:左上, 5:右上, 6:底, 7:左下, 8:右下)
                    int dirCode = dirStr switch
                    {
                        "left" => 1,
                        "right" => 2,
                        "top" => 3,
                        "topLeft" => 4,
                        "topRight" => 5,
                        "bottom" => 6,
                        "bottomLeft" => 7,
                        "bottomRight" => 8,
                        _ => 2
                    };

                    SafeInvoke(() =>
                    {
                        // 校验物理左键状态防误触
                        if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) == 0) return;
                        // 释放鼠标捕获
                        ReleaseCapture();
                        // 发送系统命令，交由 Windows 内核直接执行原生平滑调整窗口尺寸
                        SendMessage(this.Handle, WM_SYSCOMMAND, (IntPtr)(SC_SIZE + dirCode), IntPtr.Zero);
                    });
                }
                // 6. 关闭窗口
                else if (action == "closeWindow")
                {
                    SafeInvoke(this.Hide);
                }
                // 7. 最小化窗口
                else if (action == "minimizeWindow")
                {
                    SafeInvoke(() => { this.WindowState = FormWindowState.Minimized; });
                }
            }
            catch (Exception ex)
            {
                // 记录消息处理异常日志
                LogHelper.WriteLog($"[ComponentParamMatchForm] 接收消息异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 安全向前端 WebView2 投递 WebMessage 文本消息
        /// </summary>
        private void PostWebMessageSafe(string jsonMessage)
        {
            SafeInvoke(() =>
            {
                try
                {
                    if (_webView?.CoreWebView2 != null)
                    {
                        _webView.CoreWebView2.PostWebMessageAsString(jsonMessage);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[ComponentParamMatchForm] 推送消息异常: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 跨线程安全调度执行 WinForms UI 委托
        /// </summary>
        private void SafeInvoke(Action action)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            if (this.InvokeRequired)
            {
                this.BeginInvoke(action);
            }
            else
            {
                action();
            }
        }

        /// <summary>
        /// 单例调起展示窗口入口
        /// </summary>
        public static void ShowForm()
        {
            try
            {
                // 若实例不存在或已释放，重新实例化
                if (_instance == null || _instance.IsDisposed)
                {
                    _instance = new ComponentParamMatchForm();
                }

                // 若窗口未呈现则展示
                if (!_instance.Visible)
                {
                    _instance.Show();
                }

                // 置顶并激活窗口焦点
                _instance.WindowState = FormWindowState.Normal;
                _instance.BringToFront();
                _instance.Activate();
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[ComponentParamMatchForm] ShowForm 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 异步自愈补测：检测当前目录下是否存在尚未录入尺寸的 DWG 图纸，自动向 CAD 管道发起提取并通知前端增量回填
        /// 遵循规范：每 3 行代码至少包含 1 行中文注释，配置与硬编码显式标明
        /// </summary>
        /// <param name="scanResult">当前目录层级扫描模型</param>
        private void TriggerAsyncDimensionAutoFill(DirectoryHierarchyResult? scanResult)
        {
            // 校验扫描模型及其包含的图纸列表有效性
            if (scanResult == null || scanResult.DwgFiles == null || scanResult.DwgFiles.Count == 0) return;

            // 筛选出未录入尺寸且文件实际存在的图纸项集合
            var missingDwgs = scanResult.DwgFiles.FindAll(d => !d.HasDimension && !string.IsNullOrWhiteSpace(d.FullPath) && File.Exists(d.FullPath));
            // 若全部图纸均已有尺寸记录则直接退出
            if (missingDwgs.Count == 0) return;

            // 收集缺失尺寸图纸的绝对物理路径
            var missingPaths = missingDwgs.ConvertAll(d => d.FullPath);
            string currentPath = scanResult.CurrentPath;

            // 启动非阻塞后台线程池任务发起管道提取，绝不阻塞 Excel 主线程
            System.Threading.Tasks.Task.Run(async () =>
            {
                try
                {
                    // 异步请求 AutoCAD 管道服务端批量测算图纸尺寸并持久化至 SQLite
                    var extractedDims = await CadSyncClient.RequestExtractDwgDimensionsAsync(missingPaths);
                    // 校验回传的尺寸实体结果集非空
                    if (extractedDims != null && extractedDims.Count > 0)
                    {
                        // 回发增量更新指令通知前端 Vue 3 就地响应式刷新尺寸标签
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "dwgDimensionsUpdated",
                            currentPath = currentPath,
                            items = extractedDims
                        }, JsonOptions));
                    }
                }
                catch (Exception exFill)
                {
                    // 记录自动测算补全异常日志，确保宿主健壮性
                    LogHelper.WriteLog($"[ComponentParamMatchForm] 自动补测尺寸异常: {exFill.Message}");
                }
            });
        }

        /// <summary>
        /// 重写 CreateParams 样式，赋予无边框窗体系统级调整大小的厚边框特性
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                // 获取基类创建参数
                var cp = base.CreateParams;
                // 追加 WS_THICKFRAME 样式标记，开启 Windows 原生拖拽边框能力
                cp.Style |= 0x00040000;
                return cp;
            }
        }

        // Win32 非客户区命中测试与非客户区尺寸计算消息常量
        private const int WM_NCHITTEST = 0x84;
        private const int WM_NCCALCSIZE = 0x83;
        private const int HTLEFT = 10;
        private const int HTRIGHT = 11;
        private const int HTTOP = 12;
        private const int HTTOPLEFT = 13;
        private const int HTTOPRIGHT = 14;
        private const int HTBOTTOM = 15;
        private const int HTBOTTOMLEFT = 16;
        private const int HTBOTTOMRIGHT = 17;

        /// <summary>
        /// 拦截 Windows 消息以原生支持无边框窗体拖拽调整大小并彻底消除顶部 8px 系统非客户区空白
        /// 遵循规范：每 3 行代码至少包含 1 行中文注释
        /// </summary>
        protected override void WndProc(ref Message m)
        {
            // 拦截 WM_NCCALCSIZE 消息，将非客户区尺寸全部清零，彻底抹除 WS_THICKFRAME 在顶部留下的 8px 空白
            if (m.Msg == WM_NCCALCSIZE && m.WParam != IntPtr.Zero)
            {
                m.Result = IntPtr.Zero;
                return;
            }

            // 调用基类默认消息循环处理
            base.WndProc(ref m);

            // 仅在收到命中测试消息时进行边缘坐标换算
            if (m.Msg == WM_NCHITTEST)
            {
                // 将鼠标当前屏幕物理坐标转换为窗体客户区相对坐标
                Point pos = this.PointToClient(Cursor.Position);
                int w = this.ClientSize.Width;
                int h = this.ClientSize.Height;
                const int grip = 6; // --硬编码: 边缘拖拽感应像素阈值--

                // 判断鼠标落在四角或四边，直接返回对应的 Win32 命中标识由系统原生接管拖动
                if (pos.X <= grip && pos.Y <= grip) m.Result = (IntPtr)HTTOPLEFT;
                else if (pos.X >= w - grip && pos.Y <= grip) m.Result = (IntPtr)HTTOPRIGHT;
                else if (pos.X <= grip && pos.Y >= h - grip) m.Result = (IntPtr)HTBOTTOMLEFT;
                else if (pos.X >= w - grip && pos.Y >= h - grip) m.Result = (IntPtr)HTBOTTOMRIGHT;
                else if (pos.X <= grip) m.Result = (IntPtr)HTLEFT;
                else if (pos.X >= w - grip) m.Result = (IntPtr)HTRIGHT;
                else if (pos.Y <= grip) m.Result = (IntPtr)HTTOP;
                else if (pos.Y >= h - grip) m.Result = (IntPtr)HTBOTTOM;
            }
        }
    }
}
