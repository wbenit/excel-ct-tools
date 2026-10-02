using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 常规样式投标报表 WebView2 交互控制器
    /// </summary>
    public class TenderReportRegularController
    {
        // 缓存所属宿主窗体引用 (可空)
        private readonly Forms.TenderReportRegularForm? _form;

        // 全局 JSON 序列化选项：开启大小写忽略、驼峰命名对齐与原生中文防转义
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            // 允许反序列化时不区分属性名大小写
            PropertyNameCaseInsensitive = true,
            // 序列化时自动转换为前端标配的 camelCase 小写驼峰命名
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            // 保持中文字符原生输出，避免转义为 Unicode 编码
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="form">所属的 WinForms 宿主窗口</param>
        public TenderReportRegularController(Forms.TenderReportRegularForm? form = null)
        {
            _form = form;
        }

        /// <summary>
        /// 获取当前活动算价工作簿的工程信息与分类列表初始化数据
        /// </summary>
        /// <returns>JSON 字符串包含项目信息与分类列表</returns>
        public string GetInitialDataJson()
        {
            try
            {
                // 调用服务层一次性抓取初始数据
                var (projectInfo, categories) = ExcelServices.GetTenderReportInitialData();

                // 记录调试日志
                LogHelper.WriteLog($"[TenderReport] 获取初始数据成功: 项目=[{projectInfo.ProjectName}], 分类数={categories.Count}");

                // 获取持久化保存在 appsettings.json 中的报表高级偏好设置
                var currentSettings = ConfigManager.Instance.Current.TenderReport ?? new TenderReportSettings();

                // 包装为前端标准统一结构
                var response = new
                {
                    success = true,
                    data = new
                    {
                        projectInfo = projectInfo,
                        categories = categories,
                        settings = currentSettings
                    }
                };

                return JsonSerializer.Serialize(response, JsonOptions);
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[TenderReport] 读取初始数据异常: {ex.Message}");

                // 异常统一包装返回
                return JsonSerializer.Serialize(new
                {
                    success = false,
                    message = $"读取初始报表数据失败: {ex.Message}"
                }, JsonOptions);
            }
        }

        /// <summary>
        /// 执行常规样式投标报表导出生成
        /// </summary>
        /// <param name="configJson">前端提交的导出配置 JSON 字符串</param>
        /// <returns>JSON 格式的导出结果</returns>
        public string ExportReport(string configJson)
        {
            try
            {
                // 记录进入导出控制器的调试日志
                LogHelper.WriteLog($"[TenderReport] Controller 收到导出请求, 参数长度={configJson?.Length ?? 0}");

                // 检查入参字符串是否为空
                if (string.IsNullOrWhiteSpace(configJson))
                {
                    LogHelper.WriteLog("[TenderReport] 提交的导出参数为空，直接终止");
                    return JsonSerializer.Serialize(new TenderReportExportResult
                    {
                        Success = false,
                        Message = "提交的导出参数为空"
                    }, JsonOptions);
                }

                // 反序列化前端提交的配置模型
                var config = JsonSerializer.Deserialize<TenderReportExportConfig>(configJson, JsonOptions);
                if (config == null)
                {
                    LogHelper.WriteLog("[TenderReport] 导出参数反序列化失败，config 为 null");
                    return JsonSerializer.Serialize(new TenderReportExportResult
                    {
                        Success = false,
                        Message = "导出参数反序列化失败"
                    }, JsonOptions);
                }

                // 记录成功解析的工程名称与分类选择数量
                int selCatCount = config.SelectedCategories?.Count ?? 0;
                LogHelper.WriteLog($"[TenderReport] 导出参数解析成功: 工程=[{config.ProjectInfo?.ProjectName}], 选中分类数={selCatCount}");

                // 核心规则：自动将用户当前调整的报表偏好选项持久化写入本地 appsettings.json
                if (config.Settings != null)
                {
                    try
                    {
                        // 保存至单例配置管理器并写盘
                        ConfigManager.Instance.Current.TenderReport = config.Settings;
                        ConfigManager.Instance.SaveConfig(ConfigManager.Instance.Current);
                    }
                    catch (Exception exCfg)
                    {
                        // 配置保存失败不阻断导出主流程
                        LogHelper.WriteLog($"[TenderReport] 持久化报表偏好设置异常: {exCfg.Message}");
                    }
                }

                // 调用服务层生成引擎执行导出
                var result = ExcelServices.ExportTenderReportRegular(config);

                // 记录服务层导出结果指标
                LogHelper.WriteLog($"[TenderReport] 服务层导出完成: Success={result.Success}, Msg=[{result.Message}], File=[{result.OutputFilePath}]");

                return JsonSerializer.Serialize(result, JsonOptions);
            }
            catch (Exception ex)
            {
                // 记录控制器层未捕获的严重异常堆栈
                LogHelper.WriteLog($"[TenderReport] 导出控制器处理异常: {ex}");
                return JsonSerializer.Serialize(new TenderReportExportResult
                {
                    Success = false,
                    Message = $"导出处理异常: {ex.Message}"
                }, JsonOptions);
            }
        }

        /// <summary>
        /// 在 Windows 资源管理器中高亮定位指定已导出的文件
        /// </summary>
        /// <param name="filePath">导出的目标文件物理路径</param>
        public void OpenExportedFileLocation(string filePath)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                {
                    // 调用 explorer.exe /select 参数定位文件
                    Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[TenderReport] 打开文件位置异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 请求关闭宿主窗口
        /// </summary>
        public void CloseWindow()
        {
            try
            {
                _form?.SafeInvoke(() => _form.Close());
            }
            catch { }
        }
    }
}
