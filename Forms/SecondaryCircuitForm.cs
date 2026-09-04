using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using ExcelAddInDemo.Controllers;
using ExcelAddInDemo.Models;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelDna.Integration;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 基于 WebView2 + Vue 3 的“二次图回路方案与 BOM 管理中心”无边框宿主窗体
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，配置与硬编码显式标明
    /// </summary>
    public class SecondaryCircuitForm : Form
    {
        // 声明 WebView2 浏览器控件
        private readonly WebView2 _webView;

        // 声明后端控制器
        private readonly SecondaryCircuitController _controller;

        // 导入 Windows 原生 user32.dll 内存接口，用于支持无边框窗体鼠标按住拖拽
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        // 导入 SendMessage 原生函数发送系统标头拖拽消息
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        // 导入 GetAsyncKeyState 函数用于物理鼠标按键状态双重校验
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        // 常量定义: 标题栏拖拽消息标识与左键虚拟键码
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;
        private const int VK_LBUTTON = 0x01;

        // 通用 JSON 序列化配置结构 (驼峰命名与忽略大小写)
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 构造函数: 初始化控制器与窗体属性
        /// </summary>
        public SecondaryCircuitForm()
        {
            // 实例化二次方案管理控制器
            _controller = new SecondaryCircuitController();

            // 实例化 WebView2 控件
            _webView = new WebView2();

            // 配置窗体几何外观与样式
            InitializeFormProperties();

            // 配置并挂载 WebView2 控件
            InitializeWebViewControl();
        }

        /// <summary>
        /// 配置窗体几何尺寸、图标、无边框与居中属性
        /// </summary>
        private void InitializeFormProperties()
        {
            // 设置窗体标题
            this.Text = "二次图回路方案与 BOM 管理中心";
            // 获取用户当前主显示器的工作区有效几何尺寸
            var workingArea = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
            // 默认窗口宽度：自适应 1420 像素宽屏呈现，且不超过工作区 95%
            this.Width = Math.Max(1280, Math.Min(1460, (int)(workingArea.Width * 0.94))); // --硬编码-- 窗口默认宽度
            // 默认窗口高度：自适应 880 像素舒展呈现，且不超过工作区 92%
            this.Height = Math.Max(780, Math.Min(920, (int)(workingArea.Height * 0.90))); // --硬编码-- 窗口默认高度
            // 最小窗口尺寸门禁，保障图纸视口与表格完整展示
            this.MinimumSize = new Size(1180, 720);
            // 启动时在桌面居中显示
            this.StartPosition = FormStartPosition.CenterScreen;
            // 彻底去除系统原生边框以呈现纯净现代扁平效果
            this.FormBorderStyle = FormBorderStyle.None;
            // 设置底层背景色与浅灰色调对齐
            this.BackColor = Color.FromArgb(244, 247, 246);
            // 启用双缓冲减少界面闪烁
            this.DoubleBuffered = true;
        }

        /// <summary>
        /// 初始化 WebView2 控件布局并挂载至窗体
        /// </summary>
        private void InitializeWebViewControl()
        {
            // 设置 WebView2 控件充满整个窗体工作区
            _webView.Dock = DockStyle.Fill;
            // 将控件添加到窗体控件树中
            this.Controls.Add(_webView);

            // 异步初始化 WebView2 运行环境
            InitializeWebViewEnvironmentAsync();
        }

        /// <summary>
        /// 异步建立 WebView2 核心运行上下文
        /// </summary>
        private async void InitializeWebViewEnvironmentAsync()
        {
            try
            {
                // 获取插件专属 LocalAppData 临时数据保存目录
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string userDataFolder = Path.Combine(localAppData, "ExcelCTTools", "SecondaryCircuitWebView2"); // --硬编码-- 缓存路径

                // 异步创建核心运行环境实例
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                // 确保控件内部核心已成功初始化
                await _webView.EnsureCoreWebView2Async(env);

                // 禁用浏览器原生通用快捷键与右键菜单以保障原生应用体验
                _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
                _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                // 绑定前后端双向 WebMessage 接收事件
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // 寻找目标 HTML 资源文件路径 (多重备选路径检测)
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string[] candidatePaths = new string[]
                {
                    Path.Combine(baseDir, "Resources", "secondary_circuit_manage.html"),
                    Path.Combine(baseDir, "..", "Resources", "secondary_circuit_manage.html"),
                    Path.Combine(baseDir, "publish", "Resources", "secondary_circuit_manage.html"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Resources", "secondary_circuit_manage.html")
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

                // 导航至 HTML 页面 (使用 VirtualHostName 映射为安全域名，确保 WebWorker 与 WebAssembly 零跨域阻断)
                if (!string.IsNullOrEmpty(htmlPath) && File.Exists(htmlPath))
                {
                    // 提取资源根目录路径
                    string resDir = Path.GetDirectoryName(htmlPath)!;
                    // 将本地资源目录安全映射为 https://appassets.local
                    _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        "appassets.local",
                        resDir,
                        Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.Allow);

                    // 导航至虚拟主机安全页面
                    _webView.Source = new Uri("https://appassets.local/secondary_circuit_manage.html");
                }
                else
                {
                    MessageBox.Show($"未找到二次方案管理界面资源文件: secondary_circuit_manage.html", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                // 捕获环境初始化失败异常
                MessageBox.Show($"初始化 WebView2 运行环境失败: {ex.Message}", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 响应前端 Vue 3 发送的各类业务交互消息
        /// </summary>
        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 提取前端传递的原生 JSON 文本字符串
                string messageJson = e.TryGetWebMessageAsString();
                if (string.IsNullOrEmpty(messageJson)) return;

                // 解析 JSON 文档根节点
                using var doc = JsonDocument.Parse(messageJson);
                var root = doc.RootElement;

                // 提取动作标识符
                if (!root.TryGetProperty("action", out var actionProp)) return;
                string action = actionProp.GetString() ?? string.Empty;

                // 依据动作标识进行分发调度
                switch (action)
                {
                    // 1. 请求加载所有方案数据、组分类及可用物料品牌
                    case "loadSchemes":
                        string? kw = root.TryGetProperty("keyword", out var kwProp) ? kwProp.GetString() : null;
                        string? grp = root.TryGetProperty("groupName", out var grpProp) ? grpProp.GetString() : null;
                        // 查询方案列表
                        var schemes = _controller.GetSchemes(kw, grp);
                        // 查询二次排布图组
                        var groups = _controller.GetSecondaryGroups();
                        // 查询本地物料库中的全部品牌 (含二次元件)
                        var matBrands = _controller.GetMaterialBrands();
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "schemesLoaded",
                            schemes = schemes,
                            groups = groups,
                            materialBrands = matBrands
                        }, JsonOptions));
                        break;

                    // 1.1 单独获取本地物料库所有品牌
                    case "getMaterialBrands":
                        var allBrands = _controller.GetMaterialBrands();
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "materialBrandsLoaded",
                            materialBrands = allBrands
                        }, JsonOptions));
                        break;

                    // 2. 保存或更新单个方案
                    case "saveScheme":
                        if (root.TryGetProperty("scheme", out var sProp))
                        {
                            var saveRes = _controller.SaveScheme(sProp.GetRawText());
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "saveSchemeResult",
                                result = saveRes
                            }, JsonOptions));
                        }
                        break;

                    // 3. 删除方案
                    case "deleteScheme":
                        if (root.TryGetProperty("id", out var idProp) && idProp.TryGetInt32(out int delId))
                        {
                            var delRes = _controller.DeleteScheme(delId);
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "deleteSchemeResult",
                                result = delRes
                            }, JsonOptions));
                        }
                        break;

                    // 4. 从本地个人物料库检索元器件 (支持品牌与关键字组合，默认二次元件)
                    case "searchMaterials":
                        string? mKw = root.TryGetProperty("keyword", out var mKwProp) ? mKwProp.GetString() : null;
                        string? mBrand = root.TryGetProperty("brand", out var mBProp) ? mBProp.GetString() : null;
                        // 调度控制器执行多维检索
                        var materials = _controller.SearchMaterialComponents(mKw, mBrand);
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "materialsLoaded",
                            materials = materials
                        }, JsonOptions));
                        break;

                    // 5. 触发从当前活动 Excel 工作表批量扫描识别导入
                    case "importFromExcel":
                        var impRes = _controller.ImportFromActiveExcel();
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "importExcelResult",
                            result = impRes
                        }, JsonOptions));
                        break;

                    // 5.1 获取排布图与回路代号 DWG 本地图纸目录及文件列表
                    case "getDwgDirs":
                        var dwgData = _controller.GetDwgDirectoriesAndFiles();
                        PostWebMessageSafe(JsonSerializer.Serialize(new
                        {
                            action = "dwgDirsLoaded",
                            data = dwgData
                        }, JsonOptions));
                        break;

                    // 5.2 调起 WinForms 现代文件夹选择器设置图纸目录 (独立 STA 线程异步运行，彻底杜绝阻塞 WebView2 与 Excel 主线程)
                    case "selectDwgDir":
                        string target = root.TryGetProperty("target", out var tgProp) ? (tgProp.GetString() ?? "layout") : "layout";
                        // 启动独立的 STA 工作线程弹出文件夹对话框，确保绝对不阻塞主事件循环与 Chromium IPC
                        var dialogThread = new System.Threading.Thread(() =>
                        {
                            try
                            {
                                // 实例化文件夹选取器
                                using var fbd = new FolderBrowserDialog
                                {
                                    Description = target == "layout" ? "请选择【二次排布图 DWG 文件】存放目录" : "请选择【同配置回路代号 DWG 原理图】存放目录",
                                    ShowNewFolderButton = true
                                };

                                // 优先使用当前已配置的目录作为初始定位目录
                                var currentCfg = ConfigManager.Instance.Current?.SecondaryCircuit;
                                string? initialPath = target == "layout" ? currentCfg?.LayoutDwgDirectory : currentCfg?.CircuitDwgDirectory;
                                if (!string.IsNullOrWhiteSpace(initialPath) && Directory.Exists(initialPath))
                                {
                                    fbd.SelectedPath = initialPath;
                                }

                                // 模态弹窗在独立线程中独立运行，完全不阻塞主窗体与 WebView2
                                if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                                {
                                    string chosenPath = fbd.SelectedPath;
                                    // 线程安全切回主线程更新业务配置并推送给前端
                                    SafeInvoke(() =>
                                    {
                                        var setRes = _controller.SetDwgDirectory(target, chosenPath);
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
                                // 记录独立线程弹窗异常
                                LogHelper.WriteLog($"[SecondaryCircuitForm] selectDwgDir 线程异常: {ex.Message}");
                            }
                        });
                        // 显式指定 STA 单元模型以支持 Win32 Shell 接口
                        dialogThread.SetApartmentState(System.Threading.ApartmentState.STA);
                        dialogThread.IsBackground = true;
                        dialogThread.Start();
                        break;

                    // 5.3 重新扫描指定本地目录下的 DWG 文件
                    case "scanDwgDir":
                        string? scanPath = root.TryGetProperty("dirPath", out var dpProp) ? dpProp.GetString() : null;
                        if (!string.IsNullOrWhiteSpace(scanPath))
                        {
                            // 扫描该目录下 DWG 文件
                            var files = _controller.ScanDwgFiles(scanPath);
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "dwgFilesScanned",
                                dirPath = scanPath,
                                files = files
                            }, JsonOptions));
                        }
                        break;

                    // 5.3.1 扫描目录层级结构 (包含子文件夹与 DWG 图纸，支持双击下钻与面包屑)
                    case "scanDirectoryHierarchy":
                        string? hierPath = root.TryGetProperty("dirPath", out var hpProp) ? hpProp.GetString() : null;
                        if (!string.IsNullOrWhiteSpace(hierPath))
                        {
                            // 调度控制器执行综合层级扫描
                            var hierData = _controller.ScanDirectoryHierarchy(hierPath);
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "directoryHierarchyScanned",
                                data = hierData
                            }, JsonOptions));
                        }
                        break;

                    // 5.3.3 全局递归定位指定回路图号或文件名的具体物理路径及其所在父目录 (跨目录穿透反显图纸)
                    case "locateAndHighlightDwg":
                        string? locCode = root.TryGetProperty("code", out var lcProp) ? lcProp.GetString() : null;
                        string? locRoot = root.TryGetProperty("rootDir", out var lrProp) ? lrProp.GetString() : null;
                        if (!string.IsNullOrWhiteSpace(locCode))
                        {
                            // 异步在后台线程池执行全局穿透嗅探定位，确保 UI 极速响应
                            System.Threading.Tasks.Task.Run(() =>
                            {
                                var locResult = _controller.LocateDwgFile(locRoot, locCode);
                                PostWebMessageSafe(JsonSerializer.Serialize(new
                                {
                                    action = "dwgLocated",
                                    targetCode = locCode,
                                    result = locResult
                                }, JsonOptions));
                            });
                        }
                        break;

                    // 5.3.2 提取 DWG 原始二进制 Base64 数据供真实矢量视口 WebGL 渲染
                    case "getDwgBinary":
                        string? binPath = root.TryGetProperty("filePath", out var bpProp) ? bpProp.GetString() : null;
                        string? binPref = root.TryGetProperty("preferredType", out var bprefProp) ? bprefProp.GetString() : null;
                        if (!string.IsNullOrWhiteSpace(binPath))
                        {
                            // 调度控制器读取二进制流
                            var binData = _controller.GetDwgBinary(binPath, binPref);
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "dwgBinaryLoaded",
                                data = binData
                            }, JsonOptions));
                        }
                        break;

                    // 5.5 外部调起 AutoCAD 或系统关联程序打开 DWG
                    case "openInCad":
                        string? cadPath = root.TryGetProperty("filePath", out var cpProp) ? cpProp.GetString() : null;
                        string? cadPref = root.TryGetProperty("preferredType", out var cprefProp) ? cprefProp.GetString() : null;
                        if (!string.IsNullOrWhiteSpace(cadPath))
                        {
                            // 启动外部关联软件
                            var cadRes = _controller.OpenInCad(cadPath, cadPref);
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "openInCadResult",
                                result = cadRes
                            }, JsonOptions));
                        }
                        break;


                    // 5.7 扫描当前活动 Excel 工作表中的全部二次元件组 (按型号去重输出)
                    case "scanExcelComponentGroups":
                        // 使用 QueueAsMacro 交由 Excel 原生主线程执行，确保 COM 状态就绪
                        ExcelAsyncUtil.QueueAsMacro(() =>
                        {
                            // 调度 Excel 业务服务提取活动表格中的二次元件组清单 (去重聚合)
                            var groupRows = ExcelServices.ScanExcelComponentGroups();
                            // 序列化并通过 IPC 异步推送回前端 WebView2
                            PostWebMessageSafe(JsonSerializer.Serialize(new
                            {
                                action = "excelComponentGroupsScanned",
                                groups = groupRows
                            }, JsonOptions));
                        });
                        break;

                    // 5.8 批量将二次元件组与回路图号绑定持久化写入 Excel 对应行的第 32 列
                    case "saveExcelComponentGroupBindings":
                        if (root.TryGetProperty("bindings", out var bProp))
                        {
                            // 反序列化待保存的绑定映射数组
                            var bindList = JsonSerializer.Deserialize<List<ExcelServices.ComponentGroupBindingSaveDto>>(bProp.GetRawText(), JsonOptions);
                            // 关键保障：使用 QueueAsMacro 脱离 WebView2 回调上下文，交由 Excel 纯净主线程调度
                            ExcelAsyncUtil.QueueAsMacro(() =>
                            {
                                // 调度 Excel 服务批量回写第 32 列
                                var bindRes = ExcelServices.SaveExcelComponentGroupBindings(bindList);
                                // 将写入结果安全回传给前端
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
                        break;

                    // 6. 无边框窗口拖拽 (严格基于物理按键检测防幽灵死锁)
                    case "dragWindow":
                        SafeInvoke(() =>
                        {
                            // 物理校验当前左键是否仍处于按下状态
                            if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0)
                            {
                                // 释放鼠标捕获
                                ReleaseCapture();
                                // 发送非客户区左键按下消息以启动系统原生窗口拖拽
                                SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
                            }
                        });
                        break;

                    // 7. 最小化窗口
                    case "minimizeWindow":
                        // 安全调度最小化窗体状态
                        SafeInvoke(() => this.WindowState = FormWindowState.Minimized);
                        break;

                    // 8. 最大化或还原窗口
                    case "toggleMaximizeWindow":
                        // 切换最大化与普通尺寸
                        SafeInvoke(() =>
                        {
                            // 判断当前状态
                            this.WindowState = this.WindowState == FormWindowState.Maximized ? FormWindowState.Normal : FormWindowState.Maximized;
                        });
                        break;

                    // 9. 关闭当前窗口 (兼容 close 和 closeWindow 指令)
                    case "close":
                    case "closeWindow":
                        // 安全关闭当前窗体并释放资源
                        SafeInvoke(this.Close);
                        break;
                }
            }
            catch (Exception ex)
            {
                // 记录分发错误
                LogHelper.WriteLog($"[SecondaryForm] 处理前端消息异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 线程安全地向前端发送 WebMessage 消息
        /// </summary>
        private void PostWebMessageSafe(string json)
        {
            SafeInvoke(() =>
            {
                if (_webView != null && _webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.PostWebMessageAsString(json);
                }
            });
        }

        /// <summary>
        /// 遵循启发式经验 10 的 SafeInvoke 安全调用保护
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
