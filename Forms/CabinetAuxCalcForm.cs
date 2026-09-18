using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAddInDemo.Forms
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“智能辅材与壳体计算”无边框置顶宿主窗口
    /// </summary>
    public class CabinetAuxCalcForm : Form
    {
        // WebView2 浏览器主控件
        private readonly WebView2 _webView;

        // 辅材壳体计算控制器
        private readonly CabinetAuxCalcController _controller;

        // 二次回路图纸对齐与绑定业务控制器
        private readonly SecondaryCircuitController _secController;

        // JSON 序列化选项
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化控制器与 WebView2 控件属性
        /// </summary>
        public CabinetAuxCalcForm()
        {
            // 实例化辅材壳体计算控制器
            _controller = new CabinetAuxCalcController();

            // 实例化二次回路图纸控制器
            _secController = new SecondaryCircuitController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 设置 Form 窗体尺寸与显示几何外观 (840x660 像素)
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体基本外观与尺寸 (840x660 像素，居中无边框)
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题
            this.Text = "智能辅材与壳体计算";

            // 依据方案 A 宽屏架构设定尺寸为 1200x800 像素
            // 确保二次回路绑定工作台具有超 1150px 的展开宽度与 700px+ 的矢量看图视口
            this.ClientSize = new Size(1200, 800);

            // 设置屏幕中央弹出
            this.StartPosition = FormStartPosition.CenterScreen;

            // 设为无边框样式
            this.FormBorderStyle = FormBorderStyle.None;

            // 禁用原生最大化
            this.MaximizeBox = false;

            // 禁用原生最小化
            this.MinimizeBox = false;

            // 设置窗口背景填充色
            this.BackColor = Color.White;
        }

        /// <summary>
        /// 初始化 WebView2 控件并注册加载回调
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 控件充满窗体
            _webView.Dock = DockStyle.Fill;

            // 挂载至 Controls 控件集
            this.Controls.Add(_webView);

            // 绑定 Form Load 异步加载监听
            this.Load += OnFormLoadAsync;
        }

        /// <summary>
        /// 窗体加载异步事件处理程序，初始化 WebView2 运行环境并导航页面
        /// </summary>
        private async void OnFormLoadAsync(object? sender, EventArgs e)
        {
            try
            {
                // 指定 WebView2 用户专属数据缓存目录，避免因程序目录无写权限抛出拒绝访问异常
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAddInDemo",
                    "WebView2Data"
                );

                // 创建缓存文件夹
                Directory.CreateDirectory(userDataFolder);

                // 异步创建 WebView2 核心环境并指定专属用户数据目录
                var webViewEnv = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // 判断窗体有效性
                if (this.IsDisposed || this.Disposing) return;

                // 初始化 CoreWebView2 核心对象
                await _webView.EnsureCoreWebView2Async(webViewEnv);

                // 禁用默认右键菜单
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;

                // 禁用 DevTools 生产环境安全策略 (可根据需要开启)
                _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                // 注册 Web 消息接收处理函数
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 检索 HTML 页面资源路径 (多重备选路径检测)
                string appDir = Tool.GetAppDirectory();
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string[] candidatePaths = new[]
                {
                    Path.Combine(appDir, "Resources", "cabinet_aux_calc.html"),
                    Path.Combine(baseDir, "Resources", "cabinet_aux_calc.html"),
                    Path.Combine(baseDir, "..", "Resources", "cabinet_aux_calc.html"),
                    Path.Combine(baseDir, "publish", "Resources", "cabinet_aux_calc.html"),
                    Path.Combine(Application.StartupPath, "Resources", "cabinet_aux_calc.html")
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

                // 导航至 HTML 页面 (使用 VirtualHostName 映射为安全域名，确保 WebWorker 与 WebAssembly 零跨域阻断，完美支持 DWG 矢量预览)
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    // 提取资源文件夹真实物理路径
                    string resDir = Path.GetDirectoryName(htmlPath)!;
                    // 将本地资源目录安全映射为虚拟域名 https://appassets.local
                    _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        "appassets.local",
                        resDir,
                        Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.Allow);

                    // 导航至虚拟主机安全页面
                    _webView.Source = new Uri("https://appassets.local/cabinet_aux_calc.html");
                }
                else
                {
                    MessageBox.Show($"未找到辅材壳体计算界面资源文件: {htmlPath}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化 WebView2 发生异常: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 安全跨线程调度 UI 动作
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

        /// <summary>
        /// 线程安全地向前端 WebView2 投递 JSON 报文
        /// </summary>
        /// <param name="json">待推送的 JSON 字符串</param>
        private void PostWebMessageSafe(string json)
        {
            // 调度到 UI 线程安全执行
            SafeInvoke(() =>
            {
                try
                {
                    // 校验 WebView2 核心环境有效性
                    if (_webView?.CoreWebView2 != null)
                    {
                        // 向前端发送字符串消息
                        _webView.CoreWebView2.PostWebMessageAsString(json);
                    }
                }
                catch (Exception ex)
                {
                    // 记录投递异常
                    System.Diagnostics.Debug.WriteLine($"[CabinetAuxCalcForm] PostWebMessageSafe 异常: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 处理来自前端 Vue 3 的 postMessage 请求
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string jsonString = string.Empty;
                try { jsonString = e.TryGetWebMessageAsString(); } catch { }
                if (string.IsNullOrEmpty(jsonString))
                {
                    try { jsonString = e.WebMessageAsJson; } catch { }
                }
                if (string.IsNullOrEmpty(jsonString)) return;

                using var doc = JsonDocument.Parse(jsonString);
                var root = doc.RootElement;

                if (!root.TryGetProperty("action", out var actionElement)) return;
                string action = actionElement.GetString() ?? string.Empty;

                // 响应窗口平滑位移拖拽 (基于非模态物理增量，彻底杜绝 Win32 模态循环死锁导致 Excel 崩溃)
                if (action == "moveWindow")
                {
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
                }
                // 响应最小化窗口
                else if (action == "minimize")
                {
                    // 调度最小化状态
                    SafeInvoke(() => this.WindowState = FormWindowState.Minimized);
                }
                // 响应关闭窗口 (使用 BeginInvoke 异步排队，彻底消除 WebMessageReceived 消息泵内部 Dispose 导致的死锁与超时)
                else if (action == "close")
                {
                    // 异步投递至 Windows 消息队列下一帧调度关闭，使当前的 IPC 消息调用安全返回退出
                    this.BeginInvoke(new Action(() => this.Close()));
                }
                // 响应获取初始上下文数据
                else if (action == "getContext")
                {
                    // 获取上下文数据实体
                    var context = _controller.GetInitialContext();
                    // 序列化回传报文
                    string resJson = JsonSerializer.Serialize(new
                    {
                        action = "initContext",
                        data = context
                    }, JsonOptions);
                    // 线程安全回送前端
                    PostWebMessageSafe(resJson);
                }
                // 响应切换工作表获取箱柜列表
                else if (action == "getCabinetsBySheet")
                {
                    // 提取目标工作表名称
                    string sheetName = root.TryGetProperty("sheetName", out var sn) ? sn.GetString() ?? string.Empty : string.Empty;
                    // 读取该工作表所有箱柜
                    var cabList = _controller.GetCabinetsBySheet(sheetName);
                    // 序列化箱柜列表数据
                    string resJson = JsonSerializer.Serialize(new
                    {
                        action = "cabinetsList",
                        data = cabList
                    }, JsonOptions);
                    // 线程安全回送前端
                    PostWebMessageSafe(resJson);
                }
                // 响应保存规则配置
                else if (action == "saveRules")
                {
                    // 提取规则配置对象
                    if (root.TryGetProperty("rules", out var rulesElem))
                    {
                        // 反序列化规则配置
                        var rules = JsonSerializer.Deserialize<QuotationRules>(rulesElem.GetRawText(), JsonOptions);
                        // 保存操作初始状态
                        bool success = false;
                        if (rules != null)
                        {
                            // 调度控制器保存规则
                            success = _controller.SaveRules(rules);
                        }
                        // 序列化保存结果报文
                        string resJson = JsonSerializer.Serialize(new
                        {
                            action = "saveRulesResult",
                            success = success
                        }, JsonOptions);
                        // 线程安全回送前端
                        PostWebMessageSafe(resJson);
                    }
                }
                // 响应单个箱柜推导分析 (采用 ExcelAsyncUtil.QueueAsMacro 异步宏调度，释放 WinForms UI 消息泵)
                else if (action == "analyzeCabinet")
                {
                    // 提取目标工作表名称
                    string sheetName = root.TryGetProperty("sheetName", out var sn) ? sn.GetString() ?? string.Empty : string.Empty;
                    // 提取目标箱柜 Det 定义名称
                    string detName = root.TryGetProperty("detName", out var dn) ? dn.GetString() ?? string.Empty : string.Empty;
                    // 规则配置可空对象
                    QuotationRules? rules = null;
                    // 反序列化前端传入的规则
                    if (root.TryGetProperty("rules", out var rulesElem))
                    {
                        // 解析为强类型规则对象
                        rules = JsonSerializer.Deserialize<QuotationRules>(rulesElem.GetRawText(), JsonOptions);
                    }
                    // 获取生效的规则实例
                    var effectiveRules = rules ?? new QuotationRules();

                    // 调度 Excel 异步宏队列执行，彻底杜绝在 UI 线程同步裸调 COM 造成界面鼠标拖拽与点击假死
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            // 调度业务控制器单柜推导分析
                            var result = _controller.AnalyzeSingleCabinet(sheetName, detName, effectiveRules);
                            // 序列化分析结果数据报文
                            string resJson = JsonSerializer.Serialize(new
                            {
                                action = "analyzeResult",
                                success = result != null,
                                data = result
                            }, JsonOptions);
                            // 线程安全回送前端 WebView2 渲染
                            SafeInvoke(() => PostWebMessageSafe(resJson));
                        }
                        catch (Exception exAnalyze)
                        {
                            // 记录分析异常日志
                            System.Diagnostics.Debug.WriteLine($"[CabinetAuxCalcForm] analyzeCabinet 异步异常: {exAnalyze.Message}");
                            // 失败安全反馈
                            SafeInvoke(() => PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "analyzeResult",
                                success = false,
                                data = (CabinetCalcResult?)null
                            }, JsonOptions)));
                        }
                    });
                }
                // 响应写入当前选中的箱柜
                else if (action == "writeCurrentCabinet")
                {
                    // 提取目标工作表名称
                    string sheetName = root.TryGetProperty("sheetName", out var sn) ? sn.GetString() ?? string.Empty : string.Empty;
                    // 提取目标箱柜 Det 定义名称
                    string detName = root.TryGetProperty("detName", out var dn) ? dn.GetString() ?? string.Empty : string.Empty;
                    QuotationRules? rules = null;
                    // 反序列化规则配置
                    if (root.TryGetProperty("rules", out var rulesElem))
                    {
                        rules = JsonSerializer.Deserialize<QuotationRules>(rulesElem.GetRawText(), JsonOptions);
                    }

                    // 调度单台箱柜写入业务
                    var res = _controller.WriteSingleCabinet(sheetName, detName, rules ?? new QuotationRules());
                    // 序列化回执报文
                    string resJson = JsonSerializer.Serialize(new
                    {
                        action = "writeCurrentCabinetResult",
                        success = res.Success,
                        cabinetName = res.CabinetName,
                        message = res.Message
                    }, JsonOptions);
                    // 线程安全回送前端
                    PostWebMessageSafe(resJson);
                }
                // 响应更新当前分类表所有箱柜
                else if (action == "updateCurrentCategory")
                {
                    // 提取目标分类工作表名称
                    string sheetName = root.TryGetProperty("sheetName", out var sn) ? sn.GetString() ?? string.Empty : string.Empty;
                    QuotationRules? rules = null;
                    // 反序列化规则配置
                    if (root.TryGetProperty("rules", out var rulesElem))
                    {
                        rules = JsonSerializer.Deserialize<QuotationRules>(rulesElem.GetRawText(), JsonOptions);
                    }

                    var effectiveRules = rules ?? new QuotationRules();

                    // 调度 Excel 异步宏队列执行，彻底解耦 Chromium IPC 与 Excel STA 线程
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            // 执行批量更新并实时回传阶段进度
                            var res = _controller.UpdateCurrentCategory(sheetName, effectiveRules, (percent, msgText) =>
                            {
                                // 线程安全向前端 WebView2 推送进度报文
                                SafeInvoke(() =>
                                {
                                    PostWebMessageSafe(JsonSerializer.Serialize(new
                                    {
                                        action = "updateProgress",
                                        percent = percent,
                                        message = msgText
                                    }, JsonOptions));
                                });
                            });

                            // 序列化最终回执报文
                            string resJson = JsonSerializer.Serialize(new
                            {
                                action = "updateCurrentCategoryResult",
                                success = res.Success,
                                count = res.UpdatedCabinets,
                                message = res.Message
                            }, JsonOptions);

                            // 线程安全回送前端
                            SafeInvoke(() => PostWebMessageSafe(resJson));
                        }
                        catch (Exception ex)
                        {
                            // 记录异常日志并回传失败提示
                            LogHelper.WriteLog($"[CabinetAuxCalcForm] updateCurrentCategory 异常: {ex.Message}");
                            SafeInvoke(() => PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "updateCurrentCategoryResult",
                                success = false,
                                count = 0,
                                message = $"更新失败: {ex.Message}"
                            }, JsonOptions)));
                        }
                    });
                }
                // 响应更新全工作簿所有分类表
                else if (action == "updateAllCategories")
                {
                    QuotationRules? rules = null;
                    // 反序列化规则配置
                    if (root.TryGetProperty("rules", out var rulesElem))
                    {
                        rules = JsonSerializer.Deserialize<QuotationRules>(rulesElem.GetRawText(), JsonOptions);
                    }

                    var effectiveRules = rules ?? new QuotationRules();

                    // 调度 Excel 异步宏队列执行，彻底解耦 Chromium IPC
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            // 执行全工作簿批量更新并实时回传多表进度
                            var res = _controller.UpdateAllCategories(effectiveRules, (percent, msgText) =>
                            {
                                // 线程安全向前端 WebView2 推送进度报文
                                SafeInvoke(() =>
                                {
                                    PostWebMessageSafe(JsonSerializer.Serialize(new
                                    {
                                        action = "updateProgress",
                                        percent = percent,
                                        message = msgText
                                    }, JsonOptions));
                                });
                            });

                            // 序列化最终回执报文
                            string resJson = JsonSerializer.Serialize(new
                            {
                                action = "updateAllCategoriesResult",
                                success = res.Success,
                                sheetCount = res.UpdatedSheets,
                                cabinetCount = res.UpdatedCabinets,
                                message = res.Message
                            }, JsonOptions);

                            // 线程安全回送前端
                            SafeInvoke(() => PostWebMessageSafe(resJson));
                        }
                        catch (Exception ex)
                        {
                            // 记录异常日志并回传失败提示
                            LogHelper.WriteLog($"[CabinetAuxCalcForm] updateAllCategories 异常: {ex.Message}");
                            SafeInvoke(() => PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "updateAllCategoriesResult",
                                success = false,
                                sheetCount = 0,
                                cabinetCount = 0,
                                message = $"全工作簿更新失败: {ex.Message}"
                            }, JsonOptions)));
                        }
                    });
                }
                // ==================== 二次回路图纸对齐与绑定工作台动作支持 ====================
                // 1. 获取已持久化的 DWG 目录与图纸文件
                else if (action == "getDwgDirs")
                {
                    // 获取二次图纸目录配置
                    var dwgData = _secController.GetDwgDirectoriesAndFiles();
                    // 异步推送至前端 WebView2
                    PostWebMessageSafe(JsonSerializer.Serialize(new
                    {
                        action = "dwgDirsLoaded",
                        data = dwgData
                    }, JsonOptions));
                }
                // 2. 调起文件夹选择器更换 DWG 图纸目录 (独立 STA 线程运行，避免阻塞主线程)
                else if (action == "selectDwgDir")
                {
                    // 提取目标目录类型 (回路图纸默认为 circuit)
                    string target = root.TryGetProperty("target", out var tgProp) ? (tgProp.GetString() ?? "circuit") : "circuit";
                    // 启动独立的 STA 线程
                    var dialogThread = new System.Threading.Thread(() =>
                    {
                        try
                        {
                            // 实例化本地文件夹选择器
                            using var fbd = new FolderBrowserDialog
                            {
                                Description = target == "layout" ? "请选择【二次排布图 DWG 文件】存放目录" : "请选择【同配置回路代号 DWG 原理图】存放目录",
                                ShowNewFolderButton = true
                            };

                            // 读取持久化配置的目录
                            var currentCfg = ConfigManager.Instance.Current?.SecondaryCircuit;
                            string? initialPath = target == "layout" ? currentCfg?.LayoutDwgDirectory : currentCfg?.CircuitDwgDirectory;
                            // 若存在初始路径则默认定位
                            if (!string.IsNullOrWhiteSpace(initialPath) && Directory.Exists(initialPath))
                            {
                                fbd.SelectedPath = initialPath;
                            }

                            // 弹出模态选择框
                            if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                            {
                                string chosenPath = fbd.SelectedPath;
                                // 线程安全切回主线程更新业务配置
                                SafeInvoke(() =>
                                {
                                    var setRes = _secController.SetDwgDirectory(target, chosenPath);
                                    PostWebMessageSafe(JsonSerializer.Serialize(new
                                    {
                                        action = "dwgDirSelected",
                                        result = setRes
                                    }, JsonOptions));
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            // 记录异常日志
                            LogHelper.WriteLog($"[CabinetAuxCalcForm] selectDwgDir 线程异常: {ex.Message}");
                        }
                    });
                    // 设置 STA 单元状态
                    dialogThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    dialogThread.IsBackground = true;
                    dialogThread.Start();
                }
                // 3. 扫描目录层级结构 (包含子文件夹与 DWG 图纸，异步后台线程池调度，杜绝阻塞 UI 线程)
                else if (action == "scanDirectoryHierarchy")
                {
                    // 提取待层级扫描路径
                    string? hierPath = root.TryGetProperty("dirPath", out var hpProp) ? hpProp.GetString() : null;
                    if (!string.IsNullOrWhiteSpace(hierPath))
                    {
                        // 调度后台线程池执行磁盘枚举，释放 WinForms 与 WebView2 主线程
                        System.Threading.Tasks.Task.Run(() =>
                        {
                            // 调度控制器轻量浅层扫描
                            var hierData = _secController.ScanDirectoryHierarchy(hierPath);
                            // 异步线程安全回发扫描结果数据
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "directoryHierarchyScanned",
                                data = hierData
                            }, JsonOptions));
                        });
                    }
                }
                // 4. 全局递归定位指定回路图号或文件名的具体物理路径及其所在父目录 (跨目录穿透反显图纸)
                else if (action == "locateAndHighlightDwg")
                {
                    // 提取检索图号与根目录
                    string? locCode = root.TryGetProperty("code", out var lcProp) ? lcProp.GetString() : null;
                    string? locRoot = root.TryGetProperty("rootDir", out var lrProp) ? lrProp.GetString() : null;
                    if (!string.IsNullOrWhiteSpace(locCode))
                    {
                        // 后台异步穿透嗅探定位
                        System.Threading.Tasks.Task.Run(() =>
                        {
                            var locResult = _secController.LocateDwgFile(locRoot, locCode);
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "dwgLocated",
                                targetCode = locCode,
                                result = locResult
                            }, JsonOptions));
                        });
                    }
                }
                // 5. 提取 DWG 原始二进制 Base64 数据供真实矢量视口 WebGL 渲染
                else if (action == "getDwgBinary")
                {
                    // 提取图纸文件路径
                    string? binPath = root.TryGetProperty("filePath", out var bpProp) ? bpProp.GetString() : null;
                    string? binPref = root.TryGetProperty("preferredType", out var bprefProp) ? bprefProp.GetString() : null;
                    if (!string.IsNullOrWhiteSpace(binPath))
                    {
                        // 读取二进制流并转为 Base64
                        var binData = _secController.GetDwgBinary(binPath, binPref);
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "dwgBinaryLoaded",
                            data = binData
                        }, JsonOptions));
                    }
                }
                // 6. 外部调起 AutoCAD 打开 DWG
                else if (action == "openInCad")
                {
                    // 提取目标文件路径
                    string? cadPath = root.TryGetProperty("filePath", out var cpProp) ? cpProp.GetString() : null;
                    string? cadPref = root.TryGetProperty("preferredType", out var cprefProp) ? cprefProp.GetString() : null;
                    if (!string.IsNullOrWhiteSpace(cadPath))
                    {
                        // 启动关联 CAD 软件
                        var cadRes = _secController.OpenInCad(cadPath, cadPref);
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "openInCadResult",
                            result = cadRes
                        }, JsonOptions));
                    }
                }
                // 7. 扫描当前活动 Excel 工作表中的全部二次元件组 (按型号去重输出)
                else if (action == "scanExcelComponentGroups")
                {
                    // 交由 Excel 原生主线程调度执行 COM 业务
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        // 扫描二次元件组清单 (去重聚合)
                        var groupRows = ExcelServices.ScanExcelComponentGroups();
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "excelComponentGroupsScanned",
                            groups = groupRows
                        }, JsonOptions));
                    });
                }
                // 8. 批量将二次元件组与回路图号绑定持久化写入 Excel 对应行的第 32 列
                else if (action == "saveExcelComponentGroupBindings")
                {
                    if (root.TryGetProperty("bindings", out var bProp))
                    {
                        // 反序列化待保存的绑定映射数组
                        var bindList = JsonSerializer.Deserialize<List<ExcelServices.ComponentGroupBindingSaveDto>>(bProp.GetRawText(), JsonOptions);
                        // 交由 Excel 纯净主线程调度批量回写
                        ExcelAsyncUtil.QueueAsMacro(() =>
                        {
                            var bindRes = ExcelServices.SaveExcelComponentGroupBindings(bindList);
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "excelComponentGroupBindingsSaved",
                                result = new
                                {
                                    success = bindRes.Success,
                                    count = bindRes.UpdatedCount,
                                    message = bindRes.Message
                                }
                            }, JsonOptions));
                        });
                    }
                }
                // 9. 获取全部二次方案列表 (用于视口联动呈现图2五大定额参数胶囊)
                else if (action == "getSchemes")
                {
                    // 获取二次方案全量列表
                    var schemes = _secController.GetSchemes();
                    PostWebMessageSafe(JsonSerializer.Serialize(new
                    {
                        action = "schemesLoaded",
                        data = schemes
                    }, JsonOptions));
                }
                // 10. 保存或更新二次回路方案与 BOM 至 personal_components.db
                else if (action == "saveSecondaryScheme")
                {
                    // 提取前端传入的方案实体载荷数据
                    string secSchemePayload = "{}";
                    if (root.TryGetProperty("data", out var rootPayload))
                    {
                        secSchemePayload = rootPayload.GetRawText();
                    }
                    // 调度二次回路控制器保存方案入库
                    var saveRes = _secController.SaveScheme(secSchemePayload);
                    // 异步推送保存结果至前端
                    PostWebMessageSafe(JsonSerializer.Serialize(new
                    {
                        action = "saveSecondarySchemeResult",
                        result = saveRes
                    }, JsonOptions));
                }
                // 11. 检索本地物料库元器件供编辑弹窗中的物料选择器使用
                else if (action == "searchMaterialComponents")
                {
                    // 提取物料检索关键字
                    string mKw = root.TryGetProperty("keyword", out var kwProp) ? (kwProp.GetString() ?? string.Empty) : string.Empty;
                    // 从本地 SQLite 数据库物料表中检索候选元件
                    var compList = Services.PersonalComponentDbService.SearchComponents(
                        searchKeyword: mKw,
                        name: null, current: null, pole: null, tripMode: null, brand: null, mustContainRules: null, maxResults: 60);
                    // 异步回发物料检索结果列表
                    PostWebMessageSafe(JsonSerializer.Serialize(new
                    {
                        action = "searchMaterialComponentsResult",
                        data = compList
                    }, JsonOptions));
                }
                // 12. 获取全部二次回路方案列表 (供“复制其他方案”选择使用)
                else if (action == "getSecondarySchemesForCopy")
                {
                    // 提取复制方案检索关键字
                    string copyKw = root.TryGetProperty("keyword", out var ckwProp) ? (ckwProp.GetString() ?? string.Empty) : string.Empty;
                    // 读取所有二次回路方案
                    var schemesForCopy = Services.PersonalComponentDbService.GetAllSecondarySchemes(copyKw);
                    // 异步推送方案选择列表
                    PostWebMessageSafe(JsonSerializer.Serialize(new
                    {
                        action = "getSecondarySchemesForCopyResult",
                        data = schemesForCopy
                    }, JsonOptions));
                }
                // 13. 保存板材材质价格列表与加工预留配置
                else if (action == "saveMaterialPrices")
                {
                    // 检查是否存在板材价格列表
                    if (root.TryGetProperty("prices", out var pricesProp))
                    {
                        // 反序列化板材材质价格集合
                        var pricesList = JsonSerializer.Deserialize<List<MaterialPriceItem>>(pricesProp.GetRawText(), JsonOptions);
                        CabinetAllowanceConfig? allowance = null;
                        // 检查是否存在加工预留放量配置
                        if (root.TryGetProperty("allowance", out var allowProp))
                        {
                            // 反序列化加工放量配置
                            allowance = JsonSerializer.Deserialize<CabinetAllowanceConfig>(allowProp.GetRawText(), JsonOptions);
                        }
                        // 调用服务持久化存储
                        ExcelServices.UpdateMaterialPrices(pricesList ?? new List<MaterialPriceItem>(), allowance);
                        // 异步推送保存成功反馈
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "saveMaterialPricesResult",
                            success = true,
                            message = "板材材质单价配置已成功保存！"
                        }, JsonOptions));
                    }
                }
                // 14. 将柜体钣金计算出的规格型号与单价回写至当前箱柜计费区
                else if (action == "applyShellToExcel")
                {
                    // 提取目标工作表名称
                    string targetSheet = root.TryGetProperty("sheetName", out var snProp) ? (snProp.GetString() ?? string.Empty) : string.Empty;
                    // 提取回写载荷实体
                    if (root.TryGetProperty("payload", out var plProp))
                    {
                        // 反序列化壳体回写载荷
                        var writePayload = JsonSerializer.Deserialize<CabinetShellWritePayload>(plProp.GetRawText(), JsonOptions);
                        // 通过宏队列 QueueAsMacro 调度纯净 Excel COM 线程执行
                        ExcelAsyncUtil.QueueAsMacro(() =>
                        {
                            try
                            {
                                // 获取 Excel Application
                                var app = ExcelDnaSafeAccessor.GetApplication();
                                var wb = app?.ActiveWorkbook;
                                // 定位目标分类工作表
                                var ws = wb?.Worksheets[targetSheet] as Microsoft.Office.Interop.Excel.Worksheet;
                                if (ws != null && writePayload != null)
                                {
                                    // 调度服务层执行计费区回写
                                    var writeRes = ExcelServices.WriteCalculatedShellToFeeArea(ws, writePayload);
                                    // 异步通知前端回写结果
                                    PostWebMessageSafe(JsonSerializer.Serialize(new
                                    {
                                        action = "applyShellToExcelResult",
                                        success = writeRes.Success,
                                        message = writeRes.Message
                                    }, JsonOptions));
                                }
                                else
                                {
                                    // 未找到目标工作表提示
                                    PostWebMessageSafe(JsonSerializer.Serialize(new
                                    {
                                        action = "applyShellToExcelResult",
                                        success = false,
                                        message = $"未找到工作表【{targetSheet}】"
                                    }, JsonOptions));
                                }
                            }
                            catch (Exception exWrite)
                            {
                                // 异常处理并返回提示
                                PostWebMessageSafe(JsonSerializer.Serialize(new
                                {
                                    action = "applyShellToExcelResult",
                                    success = false,
                                    message = $"回写执行异常: {exWrite.Message}"
                                }, JsonOptions));
                            }
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"处理前端WebMessage发生异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 窗体关闭中事件处理：仅解绑消息监听事件，杜绝死锁
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try
            {
                // 解绑 WebMessageReceived 事件防止悬空引用
                if (_webView?.CoreWebView2 != null)
                {
                    // 移除事件监听
                    _webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
                }
                // 注意：严禁在此处同步调用 _webView.Dispose()，避免底层 Chromium 句柄通信死锁挂起
            }
            catch (Exception ex)
            {
                // 记录关闭中异常
                System.Diagnostics.Debug.WriteLine($"[CabinetAuxCalcForm] OnFormClosing 异常: {ex.Message}");
            }
            // 调度基类关闭事件
            base.OnFormClosing(e);
        }

        /// <summary>
        /// 窗体已完全关闭事件处理：在窗体完全卸载退出后安全释放 WebView2 控件资源
        /// </summary>
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try
            {
                // 窗体已完全脱离 Windows 屏幕视口，安全释放 WebView2
                _webView?.Dispose();
            }
            catch (Exception ex)
            {
                // 记录释放异常信息
                System.Diagnostics.Debug.WriteLine($"[CabinetAuxCalcForm] OnFormClosed 释放异常: {ex.Message}");
            }
            // 调度基类关闭完成事件
            base.OnFormClosed(e);
        }
    }
}
