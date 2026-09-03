using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的元器件数据管理操作悬浮窗
    /// 遵循规范：C# (Excel-DNA) + WebView2 + Vue 3 + Element Plus，主色调 #009688，非模态显示
    /// </summary>
    public class ComponentManageForm : Form
    {
        // WebView2 浏览器控件实例
        private readonly WebView2 _webView;

        // 业务控制器实例
        private readonly ComponentManageController _controller;

        // JSON 序列化通用选项
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        #region Windows API 窗口无边框拖拽支持

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int VK_LBUTTON = 0x01; // 物理鼠标左键虚拟键码
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        #endregion

        /// <summary>
        /// 构造函数：初始化无边框置顶窗口与 WebView2
        /// </summary>
        public ComponentManageForm()
        {
            _controller = new ComponentManageController();
            _webView = new WebView2();

            // 配置窗体几何与外观属性 (480x560 像素，精致悬浮面板)
            this.Text = "元器件数据管理";
            this.ClientSize = new Size(480, 560);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = true;
            this.TopMost = true; // 置顶悬浮，便于在 Excel 中划选单元格
            this.BackColor = Color.White;

            // 填充布局
            _webView.Dock = DockStyle.Fill;
            this.Controls.Add(_webView);

            // 注册生命周期事件
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 异步加载 WebView2 环境并导航加载 component_manage.html
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 获取本地数据缓存目录
                string userDataDir = Path.Combine(Tool.GetAppDataDirectory(), "WebView2_ComponentManage");
                var env = await CoreWebView2Environment.CreateAsync(null, userDataDir);
                await _webView.EnsureCoreWebView2Async(env);

                // 禁用默认右键菜单与底部状态栏
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = false;

                // 注册 Web 消息接收事件
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 检索 HTML 页面资源路径
                string appDir = Tool.GetAppDirectory();
                string htmlPath = Path.Combine(appDir, "Resources", "component_manage.html");
                if (!File.Exists(htmlPath))
                {
                    htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "component_manage.html");
                }

                if (File.Exists(htmlPath))
                {
                    _webView.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
                }
                else
                {
                    MessageBox.Show($"未找到元器件数据管理页面: {htmlPath}", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[ComponentManageForm] 加载 WebView2 异常: {ex.Message}");
                MessageBox.Show($"初始化元器件管理窗口失败: {ex.Message}", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 集中处理来自 Vue 3 前端界面的 postMessage 指令
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string rawJson = e.TryGetWebMessageAsString();
                if (string.IsNullOrWhiteSpace(rawJson)) return;

                using var doc = JsonDocument.Parse(rawJson);
                var root = doc.RootElement;
                if (!root.TryGetProperty("action", out var actionProp)) return;

                string action = actionProp.GetString() ?? string.Empty;

                switch (action)
                {
                    // 1. 前端就绪，初始化品牌列表并探测当前选区
                    case "ready":
                        string initDs = root.TryGetProperty("dataSource", out var initDsp) ? initDsp.GetString() ?? "cloud" : "cloud";
                        var brands = _controller.GetBrandStats(initDs);
                        var defaultNames = _controller.GetNamesByBrand(null, initDs);
                        var selection = _controller.DetectSelection();
                        PostMessageToWeb(new
                        {
                            action = "initData",
                            dataSource = initDs,
                            brands,
                            names = defaultNames,
                            selection
                        });
                        break;

                    // 1.1 切换数据源 (云端 vs 本地个人库)
                    case "changeDataSource":
                        string newDs = root.TryGetProperty("dataSource", out var ndsp) ? ndsp.GetString() ?? "cloud" : "cloud";
                        var dsBrands = _controller.GetBrandStats(newDs);
                        var dsNames = _controller.GetNamesByBrand(null, newDs);
                        PostMessageToWeb(new
                        {
                            action = "dataSourceChanged",
                            dataSource = newDs,
                            brands = dsBrands,
                            names = dsNames
                        });
                        break;

                    // 1.2 根据品牌获取对应的所有元器件名称
                    case "getNamesByBrand":
                        string queryBrand = root.TryGetProperty("brand", out var qbp) ? qbp.GetString() ?? "" : "";
                        string dsName = root.TryGetProperty("dataSource", out var dsn) ? dsn.GetString() ?? "cloud" : "cloud";
                        var brandNames = _controller.GetNamesByBrand(queryBrand, dsName);
                        PostMessageToWeb(new
                        {
                            action = "namesResult",
                            brand = queryBrand,
                            names = brandNames
                        });
                        break;

                    // 2. 探测当前 Excel 选区
                    case "detectSelection":
                        var selRes = _controller.DetectSelection();
                        PostMessageToWeb(new
                        {
                            action = "selectionResult",
                            selection = selRes
                        });
                        break;

                    // 3. 根据筛选条件拉取数据灌入 Excel 表格 (支持云端与个人库)
                    case "loadComponents":
                        string brand = root.TryGetProperty("brand", out var bp) ? bp.GetString() ?? "" : "";
                        string keyword = root.TryGetProperty("keyword", out var kp) ? kp.GetString() ?? "" : "";
                        string loadDs = root.TryGetProperty("dataSource", out var lds) ? lds.GetString() ?? "cloud" : "cloud";
                        var loadRes = _controller.LoadComponents(brand, keyword, loadDs);
                        PostMessageToWeb(new
                        {
                            action = "loadResult",
                            result = loadRes
                        });
                        break;

                    // 4. 选中行精准【更新】(支持云端与个人库)
                    case "updateSelected":
                        string updDs = root.TryGetProperty("dataSource", out var uds) ? uds.GetString() ?? "cloud" : "cloud";
                        var updateRes = _controller.UpdateSelected(updDs);
                        PostMessageToWeb(new
                        {
                            action = "actionResult",
                            op = "update",
                            result = updateRes
                        });
                        break;

                    // 5. 选中行精准【新增】(支持云端与个人库)
                    case "createSelected":
                        string crtDs = root.TryGetProperty("dataSource", out var cds) ? cds.GetString() ?? "cloud" : "cloud";
                        var createRes = _controller.CreateSelected(crtDs);
                        PostMessageToWeb(new
                        {
                            action = "actionResult",
                            op = "create",
                            result = createRes
                        });
                        break;

                    // 6. 选中行精准【删除】(支持云端与个人库)
                    case "deleteSelected":
                        string delDs = root.TryGetProperty("dataSource", out var dds) ? dds.GetString() ?? "cloud" : "cloud";
                        var deleteRes = _controller.DeleteSelected(delDs);
                        PostMessageToWeb(new
                        {
                            action = "actionResult",
                            op = "delete",
                            result = deleteRes
                        });
                        break;

                    // 7. 窗口平滑位移拖拽 (基于物理屏幕增量，纯非模态调度，彻底杜绝 Win32 模态循环死锁导致 Excel 崩溃)
                    case "moveWindow":
                        // 提取水平物理位移增量 deltaX
                        int deltaX = root.TryGetProperty("deltaX", out var dxp) ? dxp.GetInt32() : 0;
                        // 提取垂直物理位移增量 deltaY
                        int deltaY = root.TryGetProperty("deltaY", out var dyp) ? dyp.GetInt32() : 0;
                        // 仅当发生有效位移时调整窗口位置
                        if (deltaX != 0 || deltaY != 0)
                        {
                            SafeInvoke(() =>
                            {
                                // 直接更新窗体屏幕坐标，微秒级响应且绝不挂起 STA 消息泵
                                this.Location = new Point(this.Left + deltaX, this.Top + deltaY);
                            });
                        }
                        break;

                    // 7.1 旧版窗口拖拽指令兼容兜底 (带物理按键校验)
                    case "dragWindow":
                        SafeInvoke(() =>
                        {
                            // 物理校验：若用户在消息派发延迟期间已松开鼠标左键，直接丢弃，绝不触发系统模态拖拽死锁
                            if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) == 0)
                            {
                                return;
                            }

                            // 释放当前鼠标捕获
                            ReleaseCapture();
                            // 发送 WM_NCLBUTTONDOWN 触发原生移动循环
                            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        });
                        break;

                    // 8. 最小化窗口
                    case "minimizeWindow":
                        SafeInvoke(() => { this.WindowState = FormWindowState.Minimized; });
                        break;

                    // 9. 关闭窗口
                    case "closeWindow":
                        SafeInvoke(this.Close);
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[ComponentManageForm] 处理 Web 消息异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 向 Vue 3 前端安全推送 JSON 消息
        /// </summary>
        public void PostMessageToWeb(object data)
        {
            SafeInvoke(() =>
            {
                try
                {
                    if (_webView?.CoreWebView2 != null)
                    {
                        string json = JsonSerializer.Serialize(data, JsonOptions);
                        _webView.CoreWebView2.PostWebMessageAsString(json);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[ComponentManageForm] PostMessageToWeb 异常: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 线程安全调用辅助函数
        /// </summary>
        private void SafeInvoke(System.Action action)
        {
            try
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
            catch { }
        }
    }
}
