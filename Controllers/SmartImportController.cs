using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 智能导入箱柜 BOM 控制器
    /// 负责协调 WebView2 前端 IPC 交互、文件浏览、网格预览、双基线解析推导与最终导入
    /// </summary>
    public class SmartImportController
    {
        // 模板配置文件保存路径 (保存在当前运行目录或用户配置目录)
        private static readonly string TemplateConfigFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "CTtools", "smart_import_templates.json");

        // 统一 JSON 序列化选项
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        // 缓存当前已读取的外部工作表全量内存二维数组 (规则 7：避免重复打开外部工作簿产生红框弹窗)
        private object[,]? _cachedMatrix = null;
        // 缓存全量行数
        private int _cachedRows = 0;
        // 缓存全量列数
        private int _cachedCols = 0;
        // 缓存关联的文件绝对物理路径
        private string _cachedFilePath = string.Empty;
        // 缓存关联的工作表名称
        private string _cachedSheetName = string.Empty;
        // 同文件多 Sheet 预览结果缓存池 (切回看过的 Sheet 零 COM 开销瞬切)
        private readonly Dictionary<string, SmartImportPreviewResult> _sheetCache = new Dictionary<string, SmartImportPreviewResult>(StringComparer.OrdinalIgnoreCase);
        // 缓存最近一次全量推导解析出的箱柜列表，避免前后端之间来回传输数 MB 全量明细 JSON
        private List<ParsedCabinetModel>? _cachedCabinets = null;
        // 缓存最近一次全量解析的配置快照键
        private string _cachedParseKey = string.Empty;
        // 内存缓存访问锁，确保多线程异步读取安全
        private readonly object _cacheLock = new object();

        /// <summary>
        /// 异步 STA 线程中弹出文件选择对话框选择外部 Excel 标书
        /// </summary>
        /// <param name="callback">选择完成后的主线程安全回调</param>
        public void SelectExcelFileAsync(Action<string> callback)
        {
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    using var dialog = new System.Windows.Forms.OpenFileDialog
                    {
                        Title = "请选择外部 Excel 标书清单文件",
                        Filter = "Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*",
                        CheckFileExists = true,
                        Multiselect = false
                    };

                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.FileName))
                    {
                        callback?.Invoke(dialog.FileName);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog($"[SelectExcelFileAsync] 对话框异常: {ex.Message}");
                }
            });

            // 必须设定为单线程单元模式 (STA)
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }

        /// <summary>
        /// 读取外部 Excel 的工作表列表与前 40 行网格预览，并同步在内存建立整表只读矩阵缓存
        /// </summary>
        public SmartImportPreviewResult LoadPreviewGrid(string filePath, string sheetName = "")
        {
            lock (_cacheLock)
            {
                // 若切换了物理文件路径，清空旧文件的工作表缓存与解析缓存
                if (!string.Equals(filePath, _cachedFilePath, StringComparison.OrdinalIgnoreCase))
                {
                    _sheetCache.Clear();
                    _cachedCabinets = null;
                    _cachedParseKey = string.Empty;
                }
                // 若命中同文件下的同工作表缓存，直接秒级切换当前矩阵指针并返还预览，零 COM 开销
                else if (!string.IsNullOrWhiteSpace(sheetName) && _sheetCache.TryGetValue(sheetName, out var cachedPreview))
                {
                    _cachedMatrix = cachedPreview.FullMatrix;
                    _cachedRows = cachedPreview.FullRowCount;
                    _cachedCols = cachedPreview.FullColCount;
                    _cachedSheetName = cachedPreview.CurrentSheet;
                    _cachedCabinets = null;
                    _cachedParseKey = string.Empty;
                    return cachedPreview;
                }
            }

            // 调用静默引擎读取工作表与全量矩阵，默认列数放宽到 40 列以完整覆盖 AF 列 (第 32 列)
            var preview = SmartImportParserService.ReadPreviewGrid(filePath, sheetName, maxRows: 40, maxCols: 40);
            if (preview != null && preview.Success && preview.FullMatrix != null)
            {
                lock (_cacheLock)
                {
                    // 将整表二维矩阵安全缓存在控制器内存中
                    _cachedMatrix = preview.FullMatrix;
                    _cachedRows = preview.FullRowCount;
                    _cachedCols = preview.FullColCount;
                    _cachedFilePath = filePath;
                    _cachedSheetName = preview.CurrentSheet;
                    _cachedCabinets = null;
                    _cachedParseKey = string.Empty;
                    // 存入当前工作表缓存池
                    if (!string.IsNullOrWhiteSpace(preview.CurrentSheet))
                    {
                        _sheetCache[preview.CurrentSheet] = preview;
                    }
                }
            }
            return preview!;
        }

        /// <summary>
        /// 判断指定文件与工作表是否已经命中内存矩阵缓存 (命中即可在后台线程毫秒级纯内存解析，零 COM 调用)
        /// </summary>
        public bool HasCachedMatrix(string filePath, string sheetName)
        {
            lock (_cacheLock)
            {
                return _cachedMatrix != null &&
                    string.Equals(filePath, _cachedFilePath, StringComparison.OrdinalIgnoreCase) &&
                    (string.IsNullOrWhiteSpace(sheetName) || string.Equals(sheetName, _cachedSheetName, StringComparison.OrdinalIgnoreCase));
            }
        }

        /// <summary>
        /// 根据用户在样本柜上点选确立的规则，全量解析并返回精简预览卡片数据 (同时在后端缓存全量箱柜明细)
        /// </summary>
        public List<ParsedCabinetModel> ParsePreview(string filePath, string sheetName, SmartImportTemplateConfig config)
        {
            object[,]? matrix = null;
            int rows = 0;
            int cols = 0;
            string activeSheetName = sheetName;

            lock (_cacheLock)
            {
                // 若当前点选操作针对已缓存的文件与 Sheet，直接使用内存矩阵秒级推导，绝不重复触发 COM 打开文件
                if (_cachedMatrix != null &&
                    string.Equals(filePath, _cachedFilePath, StringComparison.OrdinalIgnoreCase) &&
                    (string.IsNullOrWhiteSpace(sheetName) || string.Equals(sheetName, _cachedSheetName, StringComparison.OrdinalIgnoreCase)))
                {
                    matrix = _cachedMatrix;
                    rows = _cachedRows;
                    cols = _cachedCols;
                    activeSheetName = _cachedSheetName;
                }
            }

            // 执行全量内存扫描 (未命中缓存时回退 COM 读取)
            List<ParsedCabinetModel> fullCabinets = (matrix != null)
                ? SmartImportParserService.ScanCabinetsFromMatrix(matrix, rows, cols, config)
                : SmartImportParserService.ParseAllCabinets(filePath, sheetName, config);

            // 构建配置指纹，将全量解析结果缓存于后端控制器，供点击导入时直接读取
            string parseKey = $"{filePath}|{activeSheetName}|{JsonSerializer.Serialize(config, JsonOptions)}";
            lock (_cacheLock)
            {
                _cachedCabinets = fullCabinets;
                _cachedParseKey = parseKey;
            }

            // 构造轻量级预览摘要集合返回给前端 (每台箱柜仅传前 3 项元件用于卡片展示，大幅减少 IPC 载荷与 Vue Proxy 开销)
            var summaries = new List<ParsedCabinetModel>(fullCabinets.Count);
            foreach (var cab in fullCabinets)
            {
                int totalCount = cab.Components?.Count ?? 0;
                // 仅截取前 3 项元器件用于前端卡片预览展示
                var previewComps = totalCount > 3 ? cab.Components!.GetRange(0, 3) : (cab.Components ?? new List<ParsedComponentModel>());
                summaries.Add(new ParsedCabinetModel
                {
                    SourceStartRow = cab.SourceStartRow,
                    SourceEndRow = cab.SourceEndRow,
                    CabinetName = cab.CabinetName,
                    // 透传箱柜名称、型号与备注供前端卡片预览
                    CabinetDisplayName = cab.CabinetDisplayName,
                    CabinetModel = cab.CabinetModel,
                    CabinetRemark = cab.CabinetRemark,
                    CabinetQuantity = cab.CabinetQuantity,
                    Components = previewComps,
                    TotalComponentCount = totalCount,
                    TotalMarkedPrice = cab.TotalMarkedPrice
                });
            }
            return summaries;
        }

        /// <summary>
        /// 执行一键创建新分类表并批量写入数据 (优先复用后端缓存的全量解析结果与内存矩阵，支持实时进度回调)
        /// </summary>
        public SmartImportExecuteResult ExecuteImport(SmartImportExecuteRequest request, Action<int, int, string>? progressCallback = null)
        {
            if (request != null && (request.Cabinets == null || request.Cabinets.Count == 0))
            {
                // 计算请求指纹比对后端缓存
                string reqKey = $"{request.FilePath}|{request.SheetName}|{JsonSerializer.Serialize(request.Config, JsonOptions)}";
                lock (_cacheLock)
                {
                    // 命中缓存直接复用完整箱柜明细
                    if (_cachedCabinets != null && _cachedCabinets.Count > 0 && string.Equals(reqKey, _cachedParseKey, StringComparison.Ordinal))
                    {
                        request.Cabinets = _cachedCabinets;
                    }
                    // 缓存配置有微小差异但矩阵仍在内存中，直接 1ms 纯内存重新推导，避免再次打开外部文件
                    else if (_cachedMatrix != null &&
                        string.Equals(request.FilePath, _cachedFilePath, StringComparison.OrdinalIgnoreCase) &&
                        (string.IsNullOrWhiteSpace(request.SheetName) || string.Equals(request.SheetName, _cachedSheetName, StringComparison.OrdinalIgnoreCase)))
                    {
                        request.Cabinets = SmartImportParserService.ScanCabinetsFromMatrix(_cachedMatrix, _cachedRows, _cachedCols, request.Config);
                    }
                }
            }

            return ExcelServices.ExecuteSmartImportToNewCategory(request!, progressCallback);
        }

        /// <summary>
        /// 保存用户自定义点选的导入规则模板
        /// </summary>
        public bool SaveTemplate(SmartImportTemplateConfig config)
        {
            if (config == null || string.IsNullOrWhiteSpace(config.TemplateName)) return false;

            try
            {
                var templates = GetTemplates();
                // 移除同名旧模板
                templates.RemoveAll(t => string.Equals(t.TemplateName, config.TemplateName, StringComparison.OrdinalIgnoreCase));
                templates.Add(config);

                string dir = Path.GetDirectoryName(TemplateConfigFile);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                string json = JsonSerializer.Serialize(templates, JsonOptions);
                File.WriteAllText(TemplateConfigFile, json);
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[SaveTemplate] 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 读取所有已保存的规则模板
        /// </summary>
        public List<SmartImportTemplateConfig> GetTemplates()
        {
            try
            {
                if (File.Exists(TemplateConfigFile))
                {
                    string json = File.ReadAllText(TemplateConfigFile);
                    var list = JsonSerializer.Deserialize<List<SmartImportTemplateConfig>>(json, JsonOptions);
                    if (list != null) return list;
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[GetTemplates] 异常: {ex.Message}");
            }
            return new List<SmartImportTemplateConfig>();
        }
    }
}
