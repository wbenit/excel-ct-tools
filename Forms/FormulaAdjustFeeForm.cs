using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Controllers;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“公式法调费”无边框宿主窗口
    /// </summary>
    public class FormulaAdjustFeeForm : Form
    {
        // 声明 WebView2 浏览器主控件
        private readonly WebView2 _webView;

        // 声明公式法调费后端 WebAPI 控制器句柄
        private readonly FormulaAdjustFeeController _controller;

        // 导入 Windows 原生 user32.dll 接口用于拖拽窗口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 原生消息接口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 常量定义: 标题栏拖拽消息标识
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 通用 JSON 序列化配置结构
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化控制器与 WebView2 控件属性
        /// </summary>
        public FormulaAdjustFeeForm()
        {
            // 实例化 WebAPI 风格控制器
            _controller = new FormulaAdjustFeeController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 设置 Form 窗体尺寸与显示几何外观
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸 (880x640 像素)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题文本
            this.Text = "公式法调费";

            // 依据图一物理布局设定尺寸为 880x640 像素
            this.ClientSize = new Size(880, 640);

            // 设置屏幕中央弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 设为无边框样式样式
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用 WinForm 原生最大化
            this.MaximizeBox = false;

            // 启用最小化
            this.MinimizeBox = true;

            // 设置窗口背景填充色
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件并注册加载回调
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 控件满框充满 Form
            _webView.Dock = DockStyle.Fill;

            // 挂载至 Controls 控件集
            this.Controls.Add(_webView);

            // 绑定 Form Load 异步加载监听
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// Form 加载事件: 初始化 WebView2 环境并导航至 Resources/formula_adjust_fee.html
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 窗体已被销毁则退出
                if (this.IsDisposed || this.Disposing) return;

                // 设置本地 AppData 中的 WebView2 缓存生成路径
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAddInDemo",
                    "WebView2Data"
                );

                // 创建缓存文件夹
                Directory.CreateDirectory(userDataFolder);

                // 异步创建 WebView2 核心环境
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 判断窗体有效性
                if (this.IsDisposed || this.Disposing) return;

                // 初始化 CoreWebView2 核心对象
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 挂载前端交互消息回调
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 寻找目标 HTML 资源文件路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // 配置多级备选路径集 --避免调试与打包执行环境差异--
                string[] candidatePaths = new string[]
                {
                    Path.Combine(baseDir, "Resources", "formula_adjust_fee.html"),
                    Path.Combine(baseDir, "..", "Resources", "formula_adjust_fee.html"),
                    Path.Combine(baseDir, "publish", "Resources", "formula_adjust_fee.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "formula_adjust_fee.html")
                };

                // 保存匹配的目标文件路径
                string htmlPath = string.Empty;

                // 循环查找首个存在的目标文件
                foreach (string candidate in candidatePaths)
                {
                    if (File.Exists(candidate))
                    {
                        htmlPath = candidate;
                        break;
                    }
                }

                // 导航至目标 HTML 资源
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    _webView.Source = new Uri(htmlPath);
                }
                else
                {
                    MessageBox.Show($"未找到公式法调费界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化 WebView2 控件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 响应来自前端 Vue 3 发来的 JSON 交互消息
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 获取前端传递的消息字符串（优先读取 String，备选读取 Raw JSON）
                string messageJson = "";
                try
                {
                    // 尝试读取直接传入的字符串消息
                    messageJson = e.TryGetWebMessageAsString();
                }
                catch { }

                // 若尝试读取为空，则直接获取原生 WebMessageAsJson 内容
                if (string.IsNullOrEmpty(messageJson))
                {
                    messageJson = e.WebMessageAsJson;
                }

                // 若解析文本仍然为空则退出
                if (string.IsNullOrEmpty(messageJson)) return;

                // 解析 JSON 数据文档
                using var doc = JsonDocument.Parse(messageJson);

                // 获取 Root 根属性节点
                var root = doc.RootElement;

                // 获取 action 指令
                string action = root.TryGetProperty("action", out var actProp) ? actProp.GetString() ?? "" : "";

                // 根据指令分支处理
                switch (action)
                {
                    // 获取初始化公式数据集
                    case "getInitFormulaData":
                        // 调用后端 WebAPI 控制器获取所有公式组 (含各自独立的明细数据)
                        var groups = _controller.GetFormulaGroups();

                        // 找到默认选中的公式组，若无则取第一个
                        var defaultGroup = groups.FirstOrDefault(g => g.IsDefault) ?? groups.FirstOrDefault();
                        // 获取当前选中的公式明细
                        var details = defaultGroup != null && defaultGroup.Details != null && defaultGroup.Details.Count > 0
                            ? defaultGroup.Details
                            : _controller.GetFormulaDetails(defaultGroup?.Name ?? "多费用公式");

                        // 构造发送给 Vue 前端的数据包
                        var resData = new
                        {
                            action = "renderFormulaData",
                            groups = groups,
                            defaultGroupId = defaultGroup?.Id,
                            details = details
                        };

                        // 回发数据包给 Vue 前端 (跨线程安全)
                        PostWebMessageSafe(JsonSerializer.Serialize(resData, JsonOptions));
                        break;

                    // 保存公式组配置至 JSON 文件
                    case "saveFormulaData":
                        // 反序列化前端提交的最新公式组完整列表
                        if (root.TryGetProperty("groups", out var groupsProp))
                        {
                            // 解析公式组集合
                            var updatedGroups = JsonSerializer.Deserialize<List<FormulaGroupModel>>(groupsProp.GetRawText(), JsonOptions);
                            // 若数据非空则调用控制器持久化至本地 JSON 文件
                            if (updatedGroups != null && updatedGroups.Count > 0)
                            {
                                // 保存至磁盘
                                _controller.SaveAllFormulaGroups(updatedGroups);
                            }
                        }
                        break;

                    // 设为默认公式组并同步写回 CabinetTemplate.xlsx 模板文件
                    case "setDefaultFormulaGroup":
                        // 获取公式组名称
                        string defGroupName = root.TryGetProperty("groupName", out var dgnProp) ? dgnProp.GetString() ?? "" : "";
                        // 解析当前明细列表
                        List<FormulaItemModel>? defItems = null;
                        if (root.TryGetProperty("details", out var defDtProp))
                        {
                            defItems = JsonSerializer.Deserialize<List<FormulaItemModel>>(defDtProp.GetRawText(), JsonOptions);
                        }
                        // 委托公共服务层执行模板更新与汇总行对齐 (删除旧计费行并汇总行对齐写入模板新计费行)
                        ExcelServices.UpdateCabinetTemplateDefaultFee(defItems, defGroupName);
                        break;

                    // 执行应用公式调费
                    case "applyFormula":
                        // 获取调费作用域
                        string scope = root.TryGetProperty("targetScope", out var scProp) ? scProp.GetString() ?? "currentCabinet" : "currentCabinet";

                        // 获取所选公式组名称
                        string gName = root.TryGetProperty("groupName", out var gnProp) ? gnProp.GetString() ?? "多费用公式" : "多费用公式";

                        // 反序列化前端编辑传入的公式明细数据列表
                        List<FormulaItemModel>? items = null;
                        if (root.TryGetProperty("details", out var dtProp))
                        {
                            items = JsonSerializer.Deserialize<List<FormulaItemModel>>(dtProp.GetRawText(), JsonOptions);
                        }

                        // 跨线程安全委托给公共 Excel 服务层执行具体计算与写入
                        ExcelServices.ApplyFormulaAdjustFeeToExcel(scope, gName, items);

                        // 弹出操作完成友好提示 (UI 线程执行)
                        SafeInvoke(() =>
                        {
                            MessageBox.Show($"公式调费应用成功！范围: {scope}, 公式组: {gName}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        });
                        break;

                    // 最小化窗口
                    case "minimize":
                        SafeInvoke(() => this.WindowState = FormWindowState.Minimized);
                        break;

                    // 关闭窗口
                    case "close":
                        SafeInvoke(() => this.Close());
                        break;

                    // 响应窗口平滑位移拖拽 (基于非模态物理增量，彻底杜绝 Win32 模态循环死锁导致 Excel 崩溃)
                    case "moveWindow":
                        int deltaX = root.TryGetProperty("deltaX", out var dxProp) ? dxProp.GetInt32() : 0;
                        int deltaY = root.TryGetProperty("deltaY", out var dyProp) ? dyProp.GetInt32() : 0;
                        if (deltaX != 0 || deltaY != 0)
                        {
                            SafeInvoke(() =>
                            {
                                // 直接更新窗体屏幕坐标，微秒级响应且绝不挂起 STA 消息泵
                                this.Location = new Point(this.Left + deltaX, this.Top + deltaY);
                            });
                        }
                        break;

                    // 响应 HTML 顶栏拖拽指令 (旧版兼容兜底)
                    case "dragWindow":
                        SafeInvoke(() =>
                        {
                            // 释放当前鼠标捕获句柄
                            ReleaseCapture();

                            // 发送 WM_NCLBUTTONDOWN (0xA1) 消息触发原生无边框窗口拖拽
                            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                        });
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"处理公式法调费前端消息异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 跨线程安全向 WebView2 发送 JSON 消息
        /// </summary>
        private void PostWebMessageSafe(string json)
        {
            // 在 UI 线程中调度发送
            SafeInvoke(() =>
            {
                // 确保 WebView2 控件及其内核有效
                if (!this.IsDisposed && _webView?.CoreWebView2 != null)
                {
                    // 向前端页面发送 JSON 消息
                    _webView.CoreWebView2.PostWebMessageAsJson(json);
                }
            });
        }

        /// <summary>
        /// 安全跨线程调度 UI 动作，防止在句柄未创建或窗体已被释放时调用抛出异常
        /// </summary>
        private void SafeInvoke(Action action)
        {
            // 校验窗体句柄有效性与释放状态
            if (this.IsDisposed || !this.IsHandleCreated) return;

            // 根据是否跨线程选择 Invoke 或直接执行
            if (this.InvokeRequired)
            {
                // 跨线程安全调度
                this.Invoke(action);
            }
            else
            {
                // 主 UI 线程直接同步执行
                action();
            }
        }
    }
}
