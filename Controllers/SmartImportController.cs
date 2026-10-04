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
            // 调用静默引擎读取工作表与全量矩阵
            var preview = SmartImportParserService.ReadPreviewGrid(filePath, sheetName, maxRows: 40, maxCols: 15);
            if (preview != null && preview.Success && preview.FullMatrix != null)
            {
                // 将整表二维矩阵安全缓存在控制器内存中
                _cachedMatrix = preview.FullMatrix;
                _cachedRows = preview.FullRowCount;
                _cachedCols = preview.FullColCount;
                _cachedFilePath = filePath;
                _cachedSheetName = preview.CurrentSheet;
            }
            return preview;
        }

        /// <summary>
        /// 根据用户在样本柜上点选确立的规则，全量解析并返回结构化箱柜卡片数据 (优先内存矩阵零 I/O 极速推导)
        /// </summary>
        public List<ParsedCabinetModel> ParsePreview(string filePath, string sheetName, SmartImportTemplateConfig config)
        {
            // 若当前点选操作针对已缓存的文件与 Sheet，直接使用内存矩阵秒级推导，绝不重复触发 COM 打开文件
            if (_cachedMatrix != null &&
                string.Equals(filePath, _cachedFilePath, StringComparison.OrdinalIgnoreCase) &&
                (string.IsNullOrWhiteSpace(sheetName) || string.Equals(sheetName, _cachedSheetName, StringComparison.OrdinalIgnoreCase)))
            {
                // 直接内存扫描推导，耗时仅 1ms，零弹窗、零延迟
                return SmartImportParserService.ScanCabinetsFromMatrix(_cachedMatrix, _cachedRows, _cachedCols, config);
            }

            // 缓存未命中时执行全量读取与解析 (带全局静默防弹窗保护)
            return SmartImportParserService.ParseAllCabinets(filePath, sheetName, config);
        }

        /// <summary>
        /// 执行一键创建新分类表并批量写入数据
        /// </summary>
        public SmartImportExecuteResult ExecuteImport(SmartImportExecuteRequest request)
        {
            return ExcelServices.ExecuteSmartImportToNewCategory(request);
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
