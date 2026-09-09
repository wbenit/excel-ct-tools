using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的箱柜管理多功能宿主窗口
    /// 支持三种模式：批建箱柜 (batch)、编辑箱柜信息 (edit)、箱柜调序 (reorder)
    /// 遵循 STA 模态安全、无边框微秒级平滑位移拖拽与每 3 行 1 行中文注释规范
    /// </summary>
    public class CabinetManageForm : Form
    {
        // 声明 WebView2 浏览器核心控件
        private readonly WebView2 _webView;

        // 声明箱柜控制器实例
        private readonly CabinetController _controller;

        // 导入 Windows 原生 user32.dll 接口以支持无边框拖拽
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 消息通知接口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 窗口移动常量定义
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // JSON 序列化配置结构 (驼峰命名转换)
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        // 当前窗口的工作模式标识: batch / edit / reorder
        private string _mode;

        /// <summary>
        /// 构造函数：初始化窗体、指定工作模式并装载 WebView2
        /// </summary>
        /// <param name="mode">模式标识: batch / edit / reorder</param>
        public CabinetManageForm(string mode = "batch")
        {
            // 保存模式参数
            _mode = string.IsNullOrWhiteSpace(mode) ? "batch" : mode.Trim();

            // 实例化箱柜控制器
            _controller = new CabinetController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 配置窗体基础外观与尺寸
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 动态平滑切换工作台模式 (batch / edit / reorder)，支持单例复用
        /// </summary>
        /// <param name="mode">目标业务模式</param>
        public void SwitchMode(string mode)
        {
            // 更新业务模式
            _mode = string.IsNullOrWhiteSpace(mode) ? "batch" : mode.Trim();
            // 重新调整窗体尺寸与标题
            InitializeFormProperties();
            // 向前端广播最新初始化数据
            SendInitDataToFront();
        }

        /// <summary>
        /// 配置 Form 外观与根据业务模式自适应窗口尺寸
        /// </summary>
        private void InitializeFormProperties()
        {
            // 根据不同业务模式自适应设定窗口尺寸与标题
            if (_mode == "edit")
            {
                this.Text = "编辑箱柜信息";
                this.ClientSize = new Size(520, 480);
            }
            else if (_mode == "reorder")
            {
                this.Text = "箱柜调序";
                this.ClientSize = new Size(680, 560);
            }
            else
            {
                // 默认批建模式：提供充足视口容纳多列数据网格
                this.Text = "批建箱柜";
                this.ClientSize = new Size(880, 620);
            }

            // 屏幕居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 采用现代化无边框设计
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用最大化与最小化
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 任务栏显示
            this.ShowInTaskbar = false;

            // 启用双缓冲消除白屏闪烁
            this.DoubleBuffered = true;

            // 浅灰色边框底色
            this.BackColor = Color.FromArgb(240, 242, 245);
        }

        /// <summary>
        /// 初始化 WebView2 控件属性与挂载
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 控件停靠撑满全屏
            _webView.Dock = DockStyle.Fill;

            // 挂载至 Controls 集合
            this.Controls.Add(_webView);

            // 注册窗体 Load 事件，延迟初始化 Chromium 核心
            this.Load += async (s, e) =>
            {
                try
                {
                    // 确定独立的 WebView2 用户数据缓存目录
                    string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                    string userDataFolder = Path.Combine(appDataDir, "ExcelAddInDemo", "WebView2_CabinetManage");

                    // 异步创建核心运行环境
                    var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                    await _webView.EnsureCoreWebView2Async(env);

                    // 禁用右键默认上下文菜单
                    _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                    // 禁用快捷键缩放
                    _webView.CoreWebView2.Settings.IsZoomControlEnabled = false;

                    // 注册 WebMessage 跨进程消息通信总线
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                    // 加载前端 cabinet_manage.html 页面
                    LoadHtmlResource();
                }
                catch (Exception ex)
                {
                    // 记录初始化异常日志
                    LogHelper.WriteLog($"WebView2 初始化异常: {ex.Message}");
                    MessageBox.Show($"界面引擎加载失败: {ex.Message}", "系统错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
        }

        /// <summary>
        /// 检索并导航加载前端 cabinet_manage.html 静态资源
        /// </summary>
        private void LoadHtmlResource()
        {
            // 获取当前程序集基准目录
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            // 构造备选资源路径
            string[] searchPaths = new[]
            {
                Path.Combine(baseDir, "Resources", "cabinet_manage.html"),
                Path.Combine(baseDir, "..", "..", "Resources", "cabinet_manage.html"),
                Path.Combine(baseDir, "publish", "Resources", "cabinet_manage.html"),
                @"d:\code\excel-ct-tools\Resources\cabinet_manage.html" // 开发源码目录兜底 --硬编码--
            };

            string htmlPath = string.Empty;
            foreach (var path in searchPaths)
            {
                if (File.Exists(path))
                {
                    htmlPath = Path.GetFullPath(path);
                    break;
                }
            }

            if (!string.IsNullOrEmpty(htmlPath))
            {
                // 本地文件导航
                _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
            }
            else
            {
                // 提示未找到资源文件
                MessageBox.Show($"未找到箱柜管理界面资源文件 (cabinet_manage.html)，请确保文件已复制至输出目录。", "资源缺失", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// 跨进程 WebMessage 消息接收回调
        /// </summary>
        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 解析前端传递的 JSON 消息字符串
                string json = e.TryGetWebMessageAsString();
                if (string.IsNullOrWhiteSpace(json)) return;

                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                string action = root.GetProperty("action").GetString() ?? "";

                switch (action)
                {
                    // 窗口移动拖拽指令
                    case "dragWindow":
                        // 确保鼠标左键处于物理按下状态再释放捕获，防幽灵死锁
                        if ((MouseButtons & MouseButtons.Left) == MouseButtons.Left)
                        {
                            ReleaseCapture();
                            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        }
                        break;

                    // 关闭窗口指令
                    case "close":
                        this.DialogResult = DialogResult.Cancel;
                        this.Close();
                        break;

                    // 前端挂载就绪索取初始化数据
                    case "getInitData":
                        SendInitDataToFront();
                        break;

                    // 前端切换当前待编辑的箱柜指令
                    case "switchEditCabinet":
                        if (root.TryGetProperty("cabinetIndex", out var idxEl))
                        {
                            int targetK = idxEl.GetInt32();
                            var specificInfo = _controller.GetCabinetInfoByIndex(targetK);
                            PostMessageSafe(new
                            {
                                action = "editCabinetChanged",
                                editInfo = specificInfo
                            });
                        }
                        break;

                    // 提交批建箱柜指令
                    case "batchCreate":
                        if (root.TryGetProperty("data", out var dataEl))
                        {
                            var items = JsonSerializer.Deserialize<List<BatchCabinetItemDto>>(dataEl.GetRawText(), JsonOptions);
                            int createdCount = _controller.BatchCreateCabinets(items ?? new List<BatchCabinetItemDto>());
                            PostMessageSafe(new
                            {
                                action = "batchCreateResult",
                                success = createdCount > 0,
                                count = createdCount,
                                message = createdCount > 0 ? $"成功批量创建 {createdCount} 台箱柜！" : "未能成功创建箱柜，请检查清单柜号是否为空。"
                            });
                            if (createdCount > 0)
                            {
                                this.DialogResult = DialogResult.OK;
                                this.Close();
                            }
                        }
                        break;

                    // 提交编辑箱柜信息指令
                    case "saveCabinetInfo":
                        if (root.TryGetProperty("data", out var editEl))
                        {
                            var dto = JsonSerializer.Deserialize<CabinetEditDto>(editEl.GetRawText(), JsonOptions);
                            bool saveSuccess = _controller.SaveCabinetInfo(dto!);
                            PostMessageSafe(new
                            {
                                action = "saveCabinetInfoResult",
                                success = saveSuccess,
                                message = saveSuccess ? "箱柜信息已成功保存并同步！" : "保存失败，未找到匹配的箱柜定义名称。"
                            });
                            if (saveSuccess)
                            {
                                this.DialogResult = DialogResult.OK;
                                this.Close();
                            }
                        }
                        break;

                    // 提交箱柜调序指令
                    case "applyReorder":
                        if (root.TryGetProperty("newOrder", out var orderEl))
                        {
                            var orderList = JsonSerializer.Deserialize<List<int>>(orderEl.GetRawText(), JsonOptions);
                            bool reorderOk = _controller.ApplyCabinetReorder(orderList ?? new List<int>());
                            PostMessageSafe(new
                            {
                                action = "applyReorderResult",
                                success = reorderOk,
                                message = reorderOk ? "箱柜排序已成功应用并更新表格！" : "应用调序失败。"
                            });
                            if (reorderOk)
                            {
                                this.DialogResult = DialogResult.OK;
                                this.Close();
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"WebMessageReceived 消息处理异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 向前端安全回传初始化配置与当前箱柜数据包
        /// </summary>
        private void SendInitDataToFront()
        {
            try
            {
                // 获取当前箱柜编辑数据 (优先光标选中，回退首台箱柜)
                var editInfo = _controller.GetCabinetInfo();

                // 获取当前工作表的全部有效箱柜列表 (供调序与编辑下拉切换共同使用)
                var allCabinets = _controller.GetCabinetsForReorder();

                // 获取当前活动工作表名称
                string curCategory = "";
                try
                {
                    var ctx = Tool.GetActiveExcelContext();
                    if (ctx?.Sheet != null) curCategory = Convert.ToString(ctx.Sheet.Name) ?? "";
                }
                catch { }

                PostMessageSafe(new
                {
                    action = "initData",
                    mode = _mode,
                    currentCategory = curCategory,
                    editInfo = editInfo,
                    allCabinets = allCabinets,
                    reorderList = allCabinets
                });
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"SendInitDataToFront 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 窗体关闭时安全释放 WebView2 核心与解绑事件，杜绝进程僵死残留
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            try
            {
                if (_webView?.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
                }
                _webView?.Dispose();
            }
            catch { }
        }

        /// <summary>
        /// 线程安全的向 WebView2 前端派发 JSON 消息
        /// </summary>
        private void PostMessageSafe(object data)
        {
            if (this.IsDisposed || !this.IsHandleCreated || _webView?.CoreWebView2 == null) return;

            string json = JsonSerializer.Serialize(data, JsonOptions);
            this.SafeInvoke(() =>
            {
                try { _webView.CoreWebView2.PostWebMessageAsString(json); } catch { }
            });
        }

        /// <summary>
        /// 线程安全执行委托操作
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
    }
}
