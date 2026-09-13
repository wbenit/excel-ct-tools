using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Services;
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

            // 设定严格的 200 × 800 像素规格
            this.ClientSize = new Size(200, 800); // --硬编码: 窗口尺寸 200x800--

            // 彻底去除原生系统边框，由 Vue3 前端绘制自研精致顶栏
            this.FormBorderStyle = FormBorderStyle.None;

            // 始终置顶显示，保障用户在 Excel 单元格切换时窗口不被遮挡
            this.TopMost = true;

            // 不在任务栏显示独立图标
            this.ShowInTaskbar = false;

            // 背景色设为纯白
            this.BackColor = Color.White;

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
                // 获取专属用户缓存数据目录
                string userDataFolder = Path.Combine(Tool.GetAppDataDirectory(), "WebView2_ParamMatch");
                // 异步创建 WebView2 运行时环境
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                // 确保控件初始化就绪
                await _webView.EnsureCoreWebView2Async(env);

                // 禁用默认右键菜单防止出现 Chromium 默认菜单
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                // 禁用状态栏显示
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

                // 注册 Web 消息接收事件监听
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 获取 Resources 资源物理文件夹
                string resDir = Path.Combine(Tool.GetAppDirectory(), "Resources");
                if (!Directory.Exists(resDir))
                {
                    resDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources");
                }

                // 映射本地虚拟 HTTPS 域名，解除本地静态资源安全限制
                _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "appassets.local",
                    resDir,
                    CoreWebView2HostResourceAccessKind.Allow
                );

                // 导航至图纸匹配页面
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

                    // 回发初始配置与根目录扫描数据
                    PostWebMessageSafe(JsonSerializer.Serialize(new
                    {
                        action = "initialDataLoaded",
                        baseDirectory = baseDir,
                        targetDirColumn = cfg.TargetDirColumn,
                        targetDwgColumn = cfg.TargetDwgColumn,
                        autoNextRow = cfg.AutoNextRow,
                        removeExtension = cfg.RemoveExtension,
                        scanData = scanResult
                    }, JsonOptions));
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

                        // 向前端回发写入结果反馈
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "bindDwgResult",
                            success = result.Success,
                            message = result.Message,
                            nextRow = result.NextRow,
                            dirName = dirName,
                            dwgName = dwgName
                        }, JsonOptions));
                    });
                }
                // 5. 鼠标按住顶栏物理拖拽移动无边框窗口
                else if (action == "moveWindow")
                {
                    int deltaX = root.TryGetProperty("deltaX", out var dxP) ? dxP.GetInt32() : 0;
                    int deltaY = root.TryGetProperty("deltaY", out var dyP) ? dyP.GetInt32() : 0;
                    SafeInvoke(() =>
                    {
                        this.Location = new Point(this.Location.X + deltaX, this.Location.Y + deltaY);
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
    }
}
