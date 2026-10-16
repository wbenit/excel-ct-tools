using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“云物料”一体化工作台无边框弹窗宿主窗体
    /// 集成物料库、柜体计算器、母排计算器与载流量速查
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public class CloudMaterialHubForm : Form
    {
        // 声明 WebView2 浏览器核心控件
        private readonly WebView2 _webView;

        // 默认激活的选项卡名称
        private readonly string _initialTab;

        // Windows 原生 user32.dll 接口用于标题栏拖拽窗口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // SendMessage 原生系统消息分发接口
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 常量定义: 标题栏拖拽消息
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        // 通用 JSON 序列化配置结构
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        /// <summary>
        /// 构造函数: 初始化窗体属性与 WebView2 控件
        /// </summary>
        /// <param name="initialTab">默认激活选项卡 ("material" | "cabinet" | "copper" | "ampacity")</param>
        public CloudMaterialHubForm(string initialTab = "material")
        {
            // 保存初始选项卡
            _initialTab = initialTab;
            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 设置 Form 窗体尺寸与显示外观 (1180x760 像素，宽屏弹性展示)
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸 (宽 1180，高 760 像素，弹性自适应)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题文本
            this.Text = "云物料一体化工作台";

            // 依据多模块标签工作台布局设定尺寸为 1180x760 像素 --硬编码: 窗体初始尺寸--
            this.ClientSize = new Size(1180, 760);

            // 设置屏幕中央居中弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 设为无边框窗口样式以实现现代扁平UI
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
            // 控件满屏铺满 Form
            _webView.Dock = DockStyle.Fill;

            // 挂载至 Controls 控件集
            this.Controls.Add(_webView);

            // 绑定 Form Load 异步加载监听
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// Form 加载事件: 初始化 WebView2 环境并导航至 Resources/cloud_material_hub.html
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

                // 再次检查窗体有效性
                if (this.IsDisposed || this.Disposing) return;

                // 初始化 CoreWebView2 核心对象
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 挂载前端交互消息回调
                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                }

                // 寻找目标 HTML 资源文件路径 (多层级自适应检索防缺失)
                string htmlPath = FindHtmlResourcePath();

                // 导航至目标 HTML 资源
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    _webView.Source = new Uri(htmlPath);
                }
                else
                {
                    MessageBox.Show($"未找到云物料工作台界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                // 提示初始化失败
                MessageBox.Show($"初始化 WebView2 控件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 多候选自适应检索 cloud_material_hub.html 物理路径
        /// </summary>
        private string FindHtmlResourcePath()
        {
            string appDir = Tool.GetAppDirectory();
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            // 配置多级备选路径集 --避免调试与发布环境物理路径差异--
            string[] candidatePaths = new string[]
            {
                Path.Combine(appDir, "Resources", "cloud_material_hub.html"),
                Path.Combine(baseDir, "Resources", "cloud_material_hub.html"),
                Path.Combine(baseDir, "publish", "Resources", "cloud_material_hub.html"),
                Path.Combine(baseDir, "..", "Resources", "cloud_material_hub.html"),
                Path.Combine(Directory.GetCurrentDirectory(), "Resources", "cloud_material_hub.html")
            };

            foreach (string path in candidatePaths)
            {
                if (File.Exists(path))
                {
                    return path;
                }
            }

            return Path.Combine(baseDir, "Resources", "cloud_material_hub.html");
        }

        /// <summary>
        /// 响应来自前端 Vue 3 发来的 JSON 交互消息
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 获取消息文本
                string messageJson = "";
                try { messageJson = e.TryGetWebMessageAsString(); } catch { }
                if (string.IsNullOrEmpty(messageJson))
                {
                    try { messageJson = e.WebMessageAsJson; } catch { }
                }

                if (string.IsNullOrWhiteSpace(messageJson)) return;

                // 解析为 JsonDocument
                using var doc = JsonDocument.Parse(messageJson);
                var root = doc.RootElement;
                string action = root.TryGetProperty("action", out var actProp) ? actProp.GetString() ?? "" : "";

                // 根据指令分发处理
                switch (action)
                {
                    case "initHub":
                        // 前端加载完成后发送初始化参数
                        HandleInitHub();
                        break;

                    case "queryMaterials":
                        // 查询物料库 (支持四分类、型号、名称、品牌筛选)
                        HandleQueryMaterials(root);
                        break;

                    case "getAmpacityTable":
                        // 获取国标母排载流量速查表
                        HandleGetAmpacityTable(root);
                        break;

                    case "insertToExcel":
                        // 插入物料至当前光标所在行下方
                        HandleInsertToExcel(root);
                        break;

                    case "replaceInExcel":
                        // 替换 Excel 当前选中行物料
                        HandleReplaceInExcel(root);
                        break;

                    case "saveToDb":
                        // 将柜体或母排计算结果加入个人物料库
                        HandleSaveToDb(root);
                        break;

                    case "saveShellPriceRules":
                        // 持久化保存柜体板材材质与各厚度单价库配置
                        HandleSaveShellPriceRules(root);
                        break;

                    case "dragWindow":
                        // 标题栏拖拽移动
                        HandleDragWindow();
                        break;

                    case "minimizeWindow":
                        // 最小化窗口
                        this.WindowState = FormWindowState.Minimized;
                        break;

                    case "closeWindow":
                        // 关闭窗口
                        this.Close();
                        break;
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[CloudMaterialHubForm] 处理前端消息异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理初始化指令，通知前端激活选项卡并回传库内品牌、铜价参数及持久化的板材单价库
        /// </summary>
        private void HandleInitHub()
        {
            // 获取本地个人物料库的真实品牌统计列表
            var brandStats = PersonalComponentDbService.GetBrandStats();
            var brandList = new List<string>();
            foreach (var b in brandStats)
            {
                if (!string.IsNullOrWhiteSpace(b.Brand)) brandList.Add(b.Brand);
            }

            // 读取已持久化的计算定额与板材单价规则配置
            var quotationRules = ExcelServices.LoadQuotationRules();
            // 提取板材价格与厚度梯度配置
            var shellPriceConfig = quotationRules?.ShellPriceRules;

            var initData = new
            {
                action = "initHubResponse",
                initialTab = _initialTab,
                copperPrice = 111465.0, // --硬编码: 默认基准铜价--
                processingFee = 0.0,
                brands = brandList,
                shellPriceRules = shellPriceConfig
            };
            PostMessageToWeb(initData);
        }

        /// <summary>
        /// 保存前端修改的柜体板材材质与各板厚每平方单价库至持久化 JSON 配置文件
        /// </summary>
        /// <param name="root">前端发送的 JSON 数据根元素</param>
        private void HandleSaveShellPriceRules(JsonElement root)
        {
            try
            {
                // 读取当前全局报价与定额配置
                var rules = ExcelServices.LoadQuotationRules() ?? new QuotationRules();
                // 提取前端回传的板材单价数组
                if (root.TryGetProperty("materialPrices", out var pricesElem))
                {
                    // 反序列化为材质单价集合
                    var list = JsonSerializer.Deserialize<List<MaterialPriceItem>>(pricesElem.GetRawText(), JsonOptions);
                    if (list != null && list.Count > 0)
                    {
                        // 更新板材单价配置
                        rules.ShellPriceRules.MaterialPrices = list;
                        // 持久化保存到 quotation_rules.json
                        ExcelServices.SaveQuotationRules(rules);
                        // 向前端发送保存成功响应
                        PostMessageToWeb(new { action = "saveShellPriceRulesResponse", success = true, message = "板材单价库已成功持久化保存！" });
                        return;
                    }
                }
                // 未能获取有效数据时响应失败
                PostMessageToWeb(new { action = "saveShellPriceRulesResponse", success = false, message = "未解析到有效的板材单价数据" });
            }
            catch (Exception ex)
            {
                // 记录异常日志并通知前端
                LogHelper.WriteLog($"[CloudMaterialHubForm] 保存板材单价异常: {ex.Message}");
                PostMessageToWeb(new { action = "saveShellPriceRulesResponse", success = false, message = $"保存失败: {ex.Message}" });
            }
        }

        /// <summary>
        /// 处理物料库分页复合查询
        /// </summary>
        private void HandleQueryMaterials(JsonElement root)
        {
            var queryParams = new HubMaterialQueryParams
            {
                Category = root.TryGetProperty("category", out var c) ? c.GetString() ?? "全部" : "全部",
                ModelKeyword = root.TryGetProperty("modelKeyword", out var m) ? m.GetString() : null,
                NameKeyword = root.TryGetProperty("nameKeyword", out var n) ? n.GetString() : null,
                Brand = root.TryGetProperty("brand", out var b) ? b.GetString() : null,
                CertFilter = root.TryGetProperty("certFilter", out var cf) ? cf.GetString() ?? "all" : "all",
                PageIndex = root.TryGetProperty("pageIndex", out var pi) ? pi.GetInt32() : 1,
                PageSize = root.TryGetProperty("pageSize", out var ps) ? ps.GetInt32() : 50
            };

            // 查询本地物料库
            var result = PersonalComponentDbService.QueryHubMaterials(queryParams);

            // 回发前端渲染
            PostMessageToWeb(new
            {
                action = "queryMaterialsResponse",
                totalCount = result.TotalCount,
                items = result.Items
            });
        }

        /// <summary>
        /// 处理载流量速查表请求
        /// </summary>
        private void HandleGetAmpacityTable(JsonElement root)
        {
            decimal copperPrice = root.TryGetProperty("copperPrice", out var cp) ? cp.GetDecimal() : 111465m;
            bool isCopper = !root.TryGetProperty("materialType", out var mt) || mt.GetString() == "copper";

            // 获取国标载流量矩阵
            var list = PersonalComponentDbService.GetBusbarAmpacityTable(copperPrice, isCopper);

            PostMessageToWeb(new
            {
                action = "getAmpacityTableResponse",
                items = list
            });
        }

        /// <summary>
        /// 处理插入物料到 Excel
        /// </summary>
        private void HandleInsertToExcel(JsonElement root)
        {
            if (!root.TryGetProperty("item", out var itemElem)) return;

            var item = JsonSerializer.Deserialize<HubMaterialDto>(itemElem.GetRawText(), JsonOptions);
            if (item == null) return;

            // 在 Excel 主线程调度插入
            this.BeginInvoke(new Action(() =>
            {
                bool success = ExcelServices.InsertMaterialFromHub(item);
                PostMessageToWeb(new
                {
                    action = "insertToExcelResponse",
                    success = success,
                    message = success ? "物料已成功插入到 Excel！" : "插入失败，请确保位于分类表明细区"
                });
            }));
        }

        /// <summary>
        /// 处理替换 Excel 选中行物料
        /// </summary>
        private void HandleReplaceInExcel(JsonElement root)
        {
            if (!root.TryGetProperty("item", out var itemElem)) return;
            bool updateQty = root.TryGetProperty("updateQuantity", out var uq) && uq.GetBoolean();

            var item = JsonSerializer.Deserialize<HubMaterialDto>(itemElem.GetRawText(), JsonOptions);
            if (item == null) return;

            // 在 Excel 主线程调度替换
            this.BeginInvoke(new Action(() =>
            {
                bool success = ExcelServices.ReplaceMaterialInExcel(item, updateQty);
                PostMessageToWeb(new
                {
                    action = "replaceInExcelResponse",
                    success = success,
                    message = success ? "Excel 选中行物料已成功替换！" : "替换失败，请先选中有效单元格"
                });
            }));
        }

        /// <summary>
        /// 处理保存物料到本地物料库 (加入柜体库/母排库)
        /// </summary>
        private void HandleSaveToDb(JsonElement root)
        {
            if (!root.TryGetProperty("item", out var itemElem)) return;

            var item = JsonSerializer.Deserialize<HubMaterialDto>(itemElem.GetRawText(), JsonOptions);
            if (item == null) return;

            int savedId = PersonalComponentDbService.SaveHubMaterial(item);
            PostMessageToWeb(new
            {
                action = "saveToDbResponse",
                success = savedId > 0,
                id = savedId,
                message = savedId > 0 ? $"已成功加入本地物料库！(ID: {savedId})" : "保存至物料库失败"
            });
        }

        /// <summary>
        /// 响应无边框窗体拖拽移动
        /// </summary>
        private void HandleDragWindow()
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
        }

        /// <summary>
        /// 安全向前端发送 JSON 消息
        /// </summary>
        private void PostMessageToWeb(object data)
        {
            try
            {
                if (_webView.CoreWebView2 != null && !this.IsDisposed && !this.Disposing)
                {
                    string json = JsonSerializer.Serialize(data, JsonOptions);
                    _webView.CoreWebView2.PostWebMessageAsString(json);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CloudMaterialHubForm] PostMessageToWeb 异常: {ex.Message}");
            }
        }
    }
}
