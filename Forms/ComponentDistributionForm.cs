using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“元器件分布穿透看板”无边框宿主窗口
    /// 遵循规则 3: 窗口类只负责 WebView2 通信与生命周期，绝不包含任何 Excel 内容操作
    /// </summary>
    public class ComponentDistributionForm : Form
    {
        // WebView2 浏览器主控件
        private readonly WebView2 _webView;

        // 当前正在查看的元器件特征入参
        private ComponentDistributionQuery _currentQuery;

        // WebView2 是否就绪
        private bool _isWebReady = false;

        // Windows 原生 API 导入用于拖动无边框窗口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 原生 SendMessage 接口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 标题栏按下拖动常量 --硬编码: Win32 API 常量--
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 通用 JSON 序列化配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        /// <summary>
        /// 构造函数: 传入元器件查询入参并初始化窗口
        /// </summary>
        public ComponentDistributionForm(ComponentDistributionQuery query)
        {
            _currentQuery = query;

            // 实例化 WebView2
            _webView = new WebView2();

            // 设置窗体基本几何尺寸与样式
            InitializeFormProperties();

            // 挂载 WebView2 并绑定事件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置无边框现代化窗体属性 (860x600 像素，弹性自适应)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗口标题
            this.Text = "元器件分布穿透看板";

            // 设置设计标准尺寸 860x600 --硬编码: 窗口标准宽高--
            this.ClientSize = new Size(860, 600);

            // 居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 无边框现代窗口
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用默认最大化最小化
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 设置背景为纯白
            this.BackColor = Color.White;

            // 保持较高级别的窗口显示层级
            this.TopMost = true;
        }

        /// <summary>
        /// 挂载并配置 WebView2
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 充满窗体
            _webView.Dock = DockStyle.Fill;
            _webView.DefaultBackgroundColor = Color.White;

            // 挂载至 Controls 集合
            this.Controls.Add(_webView);

            // 订阅 Load 事件进行异步初始化
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载时异步初始化 WebView2 并导航至 HTML 静态资源
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                if (this.IsDisposed || this.Disposing) return;

                // 设置本地 AppData 中的 WebView2 缓存生成路径 --硬编码: WebView2 数据目录--
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAddInDemo",
                    "WebView2Data"
                );

                Directory.CreateDirectory(userDataFolder);

                // 异步创建 WebView2 核心环境
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                if (this.IsDisposed || this.Disposing) return;

                // 初始化 CoreWebView2
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 注册通信消息监听
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 寻找前端静态资源路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string appDir = Tool.GetAppDirectory();

                // 候选路径搜索列表
                string[] candidatePaths = new string[]
                {
                    Path.Combine(baseDir, "Resources", "component_distribution.html"),
                    Path.Combine(baseDir, "..", "Resources", "component_distribution.html"),
                    Path.Combine(baseDir, "publish", "Resources", "component_distribution.html"),
                    Path.Combine(appDir, "Resources", "component_distribution.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "component_distribution.html")
                };

                string htmlPath = string.Empty;
                foreach (string candidate in candidatePaths)
                {
                    if (File.Exists(candidate))
                    {
                        htmlPath = candidate;
                        break;
                    }
                }

                // 导航至前端 Vue 3 页面
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    _webView.Source = new Uri(htmlPath);
                }
                else
                {
                    MessageBox.Show($"未找到元器件分布看板资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化看板 WebView2 发生异常: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogHelper.WriteLog($"[查看分布] ComponentDistributionForm 加载异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理来自前端 Vue 3 的 WebMessage 交互指令
        /// 遵循规则 3: 调用 ExcelServices 统一处理
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 解析前端传递的 JSON 字符串
                string messageJson = e.TryGetWebMessageAsString();
                if (string.IsNullOrWhiteSpace(messageJson)) return;

                using var doc = JsonDocument.Parse(messageJson);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;

                string action = actionProp.GetString() ?? string.Empty;

                switch (action)
                {
                    case "pageReady":
                        // 前端 Vue 3 挂载完毕
                        _isWebReady = true;
                        ExecuteScanAndPushData();
                        break;

                    case "getDistributionData":
                        // 前端主动请求刷新数据：若有行号，重新拉取最新 AB1 与数据行特征
                        if (_currentQuery != null && _currentQuery.SummaryRowIndex > 0)
                        {
                            // 规则 3: 调用 ExcelServices 统一提取最新查询参数 (含 AB1 源分类表)
                            var refreshedQuery = ExcelServices.ExtractComponentQueryFromSummarySheet(_currentQuery.SummaryRowIndex);
                            if (refreshedQuery != null)
                            {
                                // 更新当前入参缓存
                                _currentQuery = refreshedQuery;
                            }
                        }
                        // 执行最新定向扫描并推送到前端
                        ExecuteScanAndPushData();
                        break;

                    case "locateCell":
                        // 前端请求穿透定位目标单元格
                        string sheetName = root.TryGetProperty("sheetName", out var sProp) ? sProp.GetString() ?? "" : "";
                        int row = root.TryGetProperty("row", out var rProp) ? rProp.GetInt32() : 0;
                        int col = root.TryGetProperty("col", out var cProp) ? cProp.GetInt32() : 2;
                        int detRow = root.TryGetProperty("detRow", out var dProp) ? dProp.GetInt32() : 0;

                        if (!string.IsNullOrEmpty(sheetName) && row > 0)
                        {
                            // 规则 3: 调度公共服务层穿透定位，通过 ExcelAsyncUtil.QueueAsMacro 保证 Excel 主线程执行与聚光灯事件连通
                            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                            {
                                bool success = ExcelServices.LocateComponentInSheet(sheetName, row, col, detRow);
                                // 调度回 UI 线程回送定位结果消息
                                SafeInvoke(() =>
                                {
                                    PostMessageSafe("locateResult", new { success = success, sheetName = sheetName, row = row, detRow = detRow });
                                });
                            });
                        }
                        break;

                    case "updateSummaryFromDetail":
                        // 前端请求从分类明细表反向同步更新至元件汇总表 (方案 A 智能差量计算)
                        string upSheetName = root.TryGetProperty("sheetName", out var usProp) ? usProp.GetString() ?? "" : "";
                        int detailRowIndex = root.TryGetProperty("detailRow", out var drProp) ? drProp.GetInt32() : 0;
                        int sumRowIndex = root.TryGetProperty("summaryRow", out var srProp) ? srProp.GetInt32() : (_currentQuery != null ? _currentQuery.SummaryRowIndex : 0);

                        if (!string.IsNullOrEmpty(upSheetName) && detailRowIndex > 0 && sumRowIndex > 0)
                        {
                            // 规则 3: 窗口类不直接操作 Excel，通过 ExcelAsyncUtil.QueueAsMacro 调度公共服务执行
                            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                            {
                                // 执行差量更新方案 A 并更新到 C 列原型号规格
                                var updateResult = ExcelServices.UpdateSummaryComponentFromDetailRow(upSheetName, detailRowIndex, sumRowIndex);
                                // 调度回 UI 线程向前端 WebView2 推送执行结果
                                SafeInvoke(() =>
                                {
                                    PostMessageSafe("updateSummaryResult", updateResult);
                                });
                            });
                        }
                        break;

                    case "dragWindow":
                        // 标题栏拖拽窗口
                        SafeInvoke(() =>
                        {
                            ReleaseCapture();
                            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        });
                        break;

                    case "closeDialog":
                        // 关闭看板窗口
                        SafeInvoke(this.Close);
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[查看分布] WebMessage 处理异常: {ex.Message}");
            }
        }


        /// <summary>
        /// 执行后台全簿扫描并安全推送到前端 Vue 3
        /// </summary>
        private void ExecuteScanAndPushData()
        {
            Task.Run(() =>
            {
                try
                {
                    // 规则 3: 扫描逻辑全部在 ExcelServices 中执行
                    var data = ExcelServices.ScanComponentDistribution(_currentQuery);

                    // 调度回 UI 线程推送数据
                    SafeInvoke(() =>
                    {
                        PostMessageSafe("renderDistributionData", data);
                    });
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[查看分布] ExecuteScanAndPushData 异常: {ex.Message}");
                    SafeInvoke(() =>
                    {
                        PostMessageSafe("renderDistributionData", new ComponentDistributionResultDto
                        {
                            Success = false,
                            Message = $"查询分布失败: {ex.Message}"
                        });
                    });
                }
            });
        }

        /// <summary>
        /// 当看板已打开，用户在 Excel 中右键其他元件时无缝刷新看板
        /// </summary>
        public void UpdateQueryAndRefresh(ComponentDistributionQuery query)
        {
            _currentQuery = query;

            // 激活并前置窗口
            SafeInvoke(() =>
            {
                this.WindowState = FormWindowState.Normal;
                this.BringToFront();
                this.Activate();

                // 若前端已就绪，立即拉取最新数据
                if (_isWebReady)
                {
                    ExecuteScanAndPushData();
                }
            });
        }

        /// <summary>
        /// 线程安全向前端 WebView2 发送 JSON 消息
        /// </summary>
        private void PostMessageSafe(string actionName, object data)
        {
            if (this.IsDisposed || this.Disposing) return;
            if (_webView.CoreWebView2 == null) return;

            try
            {
                var payload = new
                {
                    action = actionName,
                    data = data
                };
                string json = JsonSerializer.Serialize(payload, JsonOptions);
                _webView.CoreWebView2.PostWebMessageAsString(json);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[查看分布] PostMessageSafe 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 跨线程调度 UI 动作安全助手
        /// </summary>
        private void SafeInvoke(Action action)
        {
            if (this.IsDisposed || this.Disposing) return;
            if (this.InvokeRequired)
            {
                this.BeginInvoke(action);
            }
            else
            {
                action();
            }
        }
    }
}
