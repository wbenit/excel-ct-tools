using System;
using System.IO;
using System.Drawing;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的聚光灯个性化外观设置宿主窗口
    /// 采用现代化无边框设计与绿蓝主调 (#009688)，支持非模态物理增量拖拽与实时联动预览
    /// </summary>
    public class SpotlightSettingForm : Form
    {
        // 声明 WebView2 浏览器容器控件
        private readonly WebView2 _webView;

        // 声明聚光灯业务数据控制器
        private readonly SpotlightController _controller;

        // 备份窗口打开时的原始配置快照，便于取消操作时无损恢复
        private SpotlightConfig _backupConfig;

        // 标记用户是否执行了明确的保存动作
        private bool _isSaved = false;

        // 通用 JSON 序列化规范：驼峰命名并忽略属性大小写
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            // 启用属性名驼峰命名转换
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            // 启用反序列化属性忽略大小写匹配
            PropertyNameCaseInsensitive = true,
            // 启用缩进美化输出
            WriteIndented = false
        };

        /// <summary>
        /// 构造函数：初始化窗体、控制器与 WebView2 控件属性
        /// </summary>
        public SpotlightSettingForm()
        {
            // 实例化聚光灯业务控制器
            _controller = new SpotlightController();

            // 备份当前全局运行配置副本以备取消还原
            _backupConfig = _controller.GetConfig().Clone();

            // 实例化内嵌 WebView2 浏览器控件
            _webView = new WebView2();

            // 初始化窗口基础几何与外观样式
            InitializeFormProperties();

            // 初始化 WebView2 控件布局并挂载窗体
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体几何尺寸、无边框风格与屏幕居中定位
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体显示标题
            this.Text = "聚光灯个性化设置";

            // 设置窗体显示宽度 520 与高度 480
            this.ClientSize = new Size(520, 480);

            // 设置窗体启动时在屏幕中央弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 采用现代化无边框风格，交由前端 Vue 精准绘制阴影与圆角
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用任务栏重复图标展示
            this.ShowInTaskbar = true;

            // 禁用最大化保持设计尺寸
            this.MaximizeBox = false;

            // 禁用最小化保持轻量体验
            this.MinimizeBox = false;

            // 设置窗体背景底色为白
            this.BackColor = Color.White;

            // 注册窗口关闭中事件，保证未保存时自动还原样式
            this.FormClosing += OnFormClosing;
        }

        /// <summary>
        /// 初始化 WebView2 控件布局并注册加载事件
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 设置 WebView2 铺满整个窗体工作区
            _webView.Dock = DockStyle.Fill;

            // 将 WebView2 添加至窗体控件集合
            this.Controls.Add(_webView);

            // 注册窗体加载事件以异步初始化环境
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载异步事件：初始化 CoreWebView2 运行环境并导航页面
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 若窗体已处于销毁状态则直接退出
                if (this.IsDisposed || this.Disposing) return;

                // 计算本地缓存数据存储目录
                string userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAddInDemo", "WebView2Data");

                // 确保缓存目录存在
                Directory.CreateDirectory(userDataFolder);

                // 创建独立的 CoreWebView2 环境对象
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 二次校验窗体生命周期有效性
                if (this.IsDisposed || this.Disposing) return;

                // 初始化 CoreWebView2 引擎内核
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 挂载前端交互消息监听事件
                if (_webView.CoreWebView2 != null)
                {
                    // 绑定 WebMessageReceived 消息处理委托
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 获取当前程序集 AppDomain 基准运行路径
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // 候选页面物理定位路径列表 (支持开发调试与打包发布环境)
                string[] candidatePaths = new string[]
                {
                    // 候选 1: 当前 AppDomain/Resources/spotlight_setting.html
                    Path.Combine(baseDir, "Resources", "spotlight_setting.html"),
                    // 候选 2: 当前工程目录/Resources/spotlight_setting.html
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "Resources", "spotlight_setting.html")
                };

                // 记录命中的有效 HTML 文件路径
                string htmlPath = string.Empty;

                // 遍历查找首个存在的 HTML 文件
                foreach (string candidate in candidatePaths)
                {
                    // 校验文件是否存在
                    if (File.Exists(candidate))
                    {
                        // 命中有效路径
                        htmlPath = candidate;
                        break;
                    }
                }

                // 若未在输出目录找到，尝试从工作区根目录兜底回退
                if (string.IsNullOrEmpty(htmlPath))
                {
                    // 尝试从工程物理路径获取
                    string fallbackPath = @"d:\code\excel-ct-tools\Resources\spotlight_setting.html";
                    // 校验兜底路径
                    if (File.Exists(fallbackPath)) htmlPath = fallbackPath;
                }

                // 若找到有效 HTML 页面资源
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    // 导航至本地 HTML 页面文件
                    _webView.Source = new Uri(htmlPath);
                }
                else
                {
                    // 弹出友好提示未找到前端页面
                    MessageBox.Show($"未找到聚光灯设置界面资源文件: {htmlPath}", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                // 捕获异常防止崩溃闪退
                MessageBox.Show($"初始化聚光灯设置 WebView2 失败: {ex.Message}", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 响应前端 Vue 3 发来的 JSON 交互消息
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 提取前端传入的消息文本 (双轨兼容：优先提取字符串，若为原生对象则读取 WebMessageAsJson)
                string messageJson = string.Empty;
                try { messageJson = e.TryGetWebMessageAsString(); } catch { }
                if (string.IsNullOrWhiteSpace(messageJson))
                {
                    try { messageJson = e.WebMessageAsJson; } catch { }
                }

                // 若消息依然为空则直接忽略
                if (string.IsNullOrWhiteSpace(messageJson)) return;

                // 解析 JSON 文档根节点
                using var doc = JsonDocument.Parse(messageJson);
                var root = doc.RootElement;

                // 读取动作 action 指令名称
                string action = root.TryGetProperty("action", out var actProp) ? actProp.GetString() ?? "" : "";

                // 响应窗口平滑物理增量位移指令 (彻底杜绝 Win32 模态挂起卡死)
                if (action == "moveWindow")
                {
                    // 读取水平偏移量
                    int deltaX = root.TryGetProperty("deltaX", out var dxProp) ? dxProp.GetInt32() : 0;
                    // 读取垂直偏移量
                    int deltaY = root.TryGetProperty("deltaY", out var dyProp) ? dyProp.GetInt32() : 0;
                    // 坐标非零时更新窗体位置
                    if (deltaX != 0 || deltaY != 0)
                    {
                        // 在 UI 线程调度平滑位移
                        SafeInvoke(() =>
                        {
                            // 物理增量坐标位移
                            this.Location = new Point(this.Left + deltaX, this.Top + deltaY);
                        });
                    }
                    return;
                }

                // 动作分支处理
                switch (action)
                {
                    // 页面初始化数据请求
                    case "getInitData":
                        // 获取当前全局生效配置
                        var currentCfg = _controller.GetConfig();
                        // 组装回发载荷 (同时附带 data 与 config 键以兼顾前端任意读取方式)
                        var initPayload = new
                        {
                            action = "initData",
                            data = currentCfg,
                            config = currentCfg
                        };
                        // 序列化为 JSON 字符串
                        string initJson = JsonSerializer.Serialize(initPayload, JsonOptions);
                        // 回发给前端 Vue 组件渲染
                        PostWebMessageAsStringSafe(initJson);
                        break;

                    // 实时联动预览指令 (用户调整滑块、选色、切模式时即时派发)
                    case "previewConfig":
                        // 兼容尝试提取 data 或 config 节点
                        JsonElement previewElem;
                        if (!root.TryGetProperty("data", out previewElem) && !root.TryGetProperty("config", out previewElem))
                        {
                            previewElem = root;
                        }
                        // 反序列化为配置模型
                        var previewCfg = JsonSerializer.Deserialize<SpotlightConfig>(previewElem.GetRawText(), JsonOptions);
                        // 若反序列化成功则调用控制器即时应用预览
                        if (previewCfg != null)
                        {
                            // 执行实时预览样式应用
                            _controller.ApplyPreview(previewCfg);
                        }
                        break;

                    // 保存配置指令
                    case "saveConfig":
                        // 兼容尝试提取 data 或 config 节点
                        JsonElement saveElem;
                        if (!root.TryGetProperty("data", out saveElem) && !root.TryGetProperty("config", out saveElem))
                        {
                            saveElem = root;
                        }
                        // 反序列化为配置模型
                        var saveCfg = JsonSerializer.Deserialize<SpotlightConfig>(saveElem.GetRawText(), JsonOptions);
                        if (saveCfg != null)
                        {
                            // 标记保存成功
                            _isSaved = true;
                            // 更新备份副本为当前已保存的最新对象
                            _backupConfig = saveCfg.Clone();
                            // 持久化保存并全局应用
                            _controller.SaveConfig(saveCfg);
                            // 回发保存成功消息
                            var saveResp = new { action = "saveSuccess", message = "保存成功" };
                            PostWebMessageAsStringSafe(JsonSerializer.Serialize(saveResp, JsonOptions));
                            // 保存成功后立即在 UI 线程安全关闭当前设置面板
                            SafeInvoke(this.Close);
                        }
                        break;

                    // 恢复出厂默认设置指令
                    case "resetDefault":
                        // 调用控制器恢复默认并写盘
                        var defCfg = _controller.ResetDefault();
                        // 更新备份副本为默认配置
                        _backupConfig = defCfg.Clone();
                        // 组装回发数据刷新前端界面 (同时附带 data 与 config 键)
                        var resetPayload = new
                        {
                            action = "onResetDefault",
                            data = defCfg,
                            config = defCfg
                        };
                        // 回传新数据
                        PostWebMessageAsStringSafe(JsonSerializer.Serialize(resetPayload, JsonOptions));
                        break;

                    // 关闭窗口指令 (多别名全面容错支持：close, closeWindow, cancel)
                    case "close":
                    case "closeWindow":
                    case "cancel":
                        // 安全关闭窗口 (未保存时将在 FormClosing 钩子中自愈还原)
                        SafeInvoke(this.Close);
                        break;
                }
            }
            catch (Exception ex)
            {
                // 异常记录与安全捕获
                var errPayload = new { action = "onError", message = ex.Message };
                PostWebMessageAsStringSafe(JsonSerializer.Serialize(errPayload, JsonOptions));
            }
        }

        /// <summary>
        /// 窗体即将关闭事件钩子：若用户未点击保存便关闭窗口，自动撤销临时预览并还原备份配置
        /// </summary>
        private void OnFormClosing(object? sender, FormClosingEventArgs e)
        {
            try
            {
                // 若用户未点击保存直接关闭
                if (!_isSaved && _backupConfig != null)
                {
                    // 恢复至打开设置面板前的备份快照
                    _controller.SaveConfig(_backupConfig);
                }
            }
            catch
            {
                // 静默容错保护
            }
        }

        /// <summary>
        /// 跨线程安全向 WebView2 前端发送 JSON 字符串消息
        /// </summary>
        private void PostWebMessageAsStringSafe(string text)
        {
            // 调度至 UI 消息循环
            SafeInvoke(() =>
            {
                // 确保窗体未销毁且 WebView2 内核就绪
                if (!this.IsDisposed && _webView?.CoreWebView2 != null)
                {
                    // 向前端发送文本
                    _webView.CoreWebView2.PostWebMessageAsString(text);
                }
            });
        }

        /// <summary>
        /// 安全跨线程调度辅助委托
        /// </summary>
        private void SafeInvoke(Action action)
        {
            // 校验窗体句柄有效性
            if (this.IsDisposed || !this.IsHandleCreated) return;
            // 判断是否需要 Invoke
            if (this.InvokeRequired)
            {
                // 异步调度
                this.BeginInvoke(action);
            }
            else
            {
                // 同步执行
                action();
            }
        }
    }
}
