using System;
using System.Collections.Generic;
using System.Text.Json;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 云方案中心控制器，负责前端 WebView2 与 C# ExcelServices 之间的数据中转与分发
    /// </summary>
    public class CloudSolutionController
    {
        // 通用 JSON 序列化配置参数
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false
        };

        /// <summary>
        /// 分页多维检索方案列表
        /// </summary>
        public CloudSchemePageResult QuerySchemes(string queryJson)
        {
            try
            {
                // 反序列化前端传来的查询参数
                var query = JsonSerializer.Deserialize<CloudSchemeQueryDto>(queryJson, JsonOptions) 
                            ?? new CloudSchemeQueryDto();
                // 调度业务层服务执行多维检索
                return ExcelServices.QuerySchemes(query);
            }
            catch (Exception ex)
            {
                // 记录查询异常日志
                LogHelper.WriteLog($"[CloudSolutionController] QuerySchemes 异常: {ex.Message}");
                // 异常兜底返回空结果集
                return new CloudSchemePageResult();
            }
        }

        /// <summary>
        /// 获取单个方案详情 (含多图纸与 BOM 元器件)
        /// </summary>
        public CloudSchemeDetail? GetSchemeDetail(string id)
        {
            try
            {
                // 调度业务层根据 ID 查询完整详情
                return ExcelServices.GetSchemeDetailById(id);
            }
            catch (Exception ex)
            {
                // 记录详情获取异常
                LogHelper.WriteLog($"[CloudSolutionController] GetSchemeDetail 异常: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 切换方案的星标收藏状态
        /// </summary>
        public bool ToggleFavorite(string id)
        {
            try
            {
                // 调度业务层切换收藏
                return ExcelServices.ToggleSchemeFavorite(id);
            }
            catch (Exception ex)
            {
                // 记录收藏切换异常
                LogHelper.WriteLog($"[CloudSolutionController] ToggleFavorite 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 保存或另存为企业自建方案
        /// </summary>
        public bool SaveEnterpriseScheme(string schemeJson)
        {
            try
            {
                // 反序列化方案实体
                var scheme = JsonSerializer.Deserialize<CloudSchemeDetail>(schemeJson, JsonOptions);
                if (scheme == null) return false;
                // 调度业务层执行保存
                return ExcelServices.SaveEnterpriseScheme(scheme);
            }
            catch (Exception ex)
            {
                // 记录保存异常
                LogHelper.WriteLog($"[CloudSolutionController] SaveEnterpriseScheme 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 删除指定的企业自建方案
        /// </summary>
        public bool DeleteEnterpriseScheme(string id)
        {
            try
            {
                // 调度业务层执行物理删除
                return ExcelServices.DeleteEnterpriseScheme(id);
            }
            catch (Exception ex)
            {
                // 记录删除异常
                LogHelper.WriteLog($"[CloudSolutionController] DeleteEnterpriseScheme 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 将勾选的方案 BOM 插入当前活动 Excel 分类表
        /// </summary>
        public (bool Success, string Message) InsertSchemeToExcel(string insertDtoJson)
        {
            try
            {
                // 反序列化插入参数
                var dto = JsonSerializer.Deserialize<SchemeInsertToExcelDto>(insertDtoJson, JsonOptions);
                if (dto == null) return (false, "插入参数为空！");
                // 调度业务层执行 Excel 写入
                return ExcelServices.InsertSchemeBomToExcel(dto);
            }
            catch (Exception ex)
            {
                // 记录插入异常
                LogHelper.WriteLog($"[CloudSolutionController] InsertSchemeToExcel 异常: {ex.Message}");
                return (false, $"执行异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取当前配置的二次回路图纸根目录
        /// </summary>
        public string GetSecondaryCircuitDwgDir()
        {
            // 从配置管理器中读取二次图纸目录
            return ConfigManager.Instance.Current?.SecondaryCircuit?.CircuitDwgDirectory ?? string.Empty;
        }

        /// <summary>
        /// 更新并持久化保存二次回路图纸根目录
        /// </summary>
        public bool SetSecondaryCircuitDwgDir(string path)
        {
            try
            {
                // 使用 ConfigManager 专用的二次方案目录持久化更新接口
                ConfigManager.Instance.UpdateSecondaryDwgDirectories(null, path?.Trim() ?? string.Empty);
                return true;
            }
            catch (Exception ex)
            {
                // 记录保存异常日志
                LogHelper.WriteLog($"[CloudSolutionController] SetSecondaryCircuitDwgDir 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 扫描指定二次图纸根目录下的所有子文件夹及各自的 DWG 数量
        /// </summary>
        public List<SecondaryFolderItemDto> ScanSecondaryFolders(string? rootDir = null)
        {
            // 初始化返回列表
            var result = new List<SecondaryFolderItemDto>();
            // 若未传参则读取当前配置目录
            string targetDir = string.IsNullOrWhiteSpace(rootDir) ? GetSecondaryCircuitDwgDir() : rootDir.Trim();
            // 路径有效性与物理存在性判断
            if (string.IsNullOrWhiteSpace(targetDir) || !System.IO.Directory.Exists(targetDir))
            {
                return result;
            }

            try
            {
                // 枚举根目录下所有直接子目录
                var subDirs = System.IO.Directory.GetDirectories(targetDir);
                // 按目录名称字母自然拼音排序
                Array.Sort(subDirs, StringComparer.OrdinalIgnoreCase);

                foreach (var dir in subDirs)
                {
                    try
                    {
                        var di = new System.IO.DirectoryInfo(dir);
                        // 过滤隐藏目录
                        if ((di.Attributes & System.IO.FileAttributes.Hidden) != 0) continue;

                        // 快速统计该子文件夹下的 .dwg 文件数量
                        int dwgCount = System.IO.Directory.GetFiles(dir, "*.dwg", System.IO.SearchOption.TopDirectoryOnly).Length;

                        // 添加到返回项
                        result.Add(new SecondaryFolderItemDto
                        {
                            Name = di.Name,
                            FullPath = di.FullName,
                            DwgCount = dwgCount
                        });
                    }
                    catch (Exception exSub)
                    {
                        // 记录单个子目录读取异常
                        LogHelper.WriteLog($"[CloudSolutionController] 读取子文件夹 {dir} 异常: {exSub.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录枚举整体异常
                LogHelper.WriteLog($"[CloudSolutionController] ScanSecondaryFolders 异常: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 扫描指定子文件夹内的所有 DWG 文件并到 personal_components 数据库匹配参数组装卡片列表
        /// </summary>
        public List<SecondaryFolderDwgCardDto> GetFolderDwgCards(string folderPath, string folderName, string? keyword = null)
        {
            // 初始化卡片集合
            var cards = new List<SecondaryFolderDwgCardDto>();
            string cleanKw = keyword?.Trim() ?? string.Empty;

            try
            {
                // 1. 若明确指定了子文件夹路径且物理存在，扫描该指定子文件夹
                if (!string.IsNullOrWhiteSpace(folderPath) && System.IO.Directory.Exists(folderPath))
                {
                    // 扫描单目录并装配卡片
                    ScanAndAppendCards(folderPath, folderName, cleanKw, cards);
                }
                // 2. 若未指定子文件夹 (即前端处于"全部方案分类")，遍历根目录下所有子目录聚合展示
                else
                {
                    // 获取当前配置的二次方案根目录
                    string rootDir = GetSecondaryCircuitDwgDir();
                    if (!string.IsNullOrWhiteSpace(rootDir) && System.IO.Directory.Exists(rootDir))
                    {
                        // 获取所有直接子文件夹列表
                        var subDirs = System.IO.Directory.GetDirectories(rootDir);
                        Array.Sort(subDirs, StringComparer.OrdinalIgnoreCase);

                        foreach (var subDir in subDirs)
                        {
                            try
                            {
                                var di = new System.IO.DirectoryInfo(subDir);
                                // 过滤隐藏文件夹
                                if ((di.Attributes & System.IO.FileAttributes.Hidden) != 0) continue;
                                // 依次扫描每个子目录并追加至卡片列表
                                ScanAndAppendCards(subDir, di.Name, cleanKw, cards);
                            }
                            catch (Exception exSub)
                            {
                                // 记录单个子目录遍历异常
                                LogHelper.WriteLog($"[CloudSolutionController] 遍历子目录 {subDir} 异常: {exSub.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录组装卡片异常
                LogHelper.WriteLog($"[CloudSolutionController] GetFolderDwgCards 异常: {ex.Message}");
            }

            return cards;
        }

        /// <summary>
        /// 扫描单个目录下的所有 DWG 文件，并匹配数据库方案填充至卡片列表
        /// </summary>
        private void ScanAndAppendCards(string targetDir, string curFolderName, string cleanKw, List<SecondaryFolderDwgCardDto> cardList)
        {
            // 扫描当前目录下的所有 DWG 图纸文件
            var dwgFiles = Services.DwgPreviewService.ScanDwgFiles(targetDir);

            foreach (var file in dwgFiles)
            {
                // 去扩展名的图纸名称 (如 "WATSG")
                string dwgName = file.NameWithoutExt;

                // 到 personal_components.db 按照 applicable_codes 优先匹配对应方案
                var matchedScheme = Services.PersonalComponentDbService.FindSchemeByDwgName(dwgName);

                // 若传入了关键字，执行模糊过滤 (过滤文件名、分类名、品牌、描述、方案名)
                if (!string.IsNullOrEmpty(cleanKw))
                {
                    bool hitKw = dwgName.IndexOf(cleanKw, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 curFolderName.IndexOf(cleanKw, StringComparison.OrdinalIgnoreCase) >= 0;

                    if (matchedScheme != null)
                    {
                        hitKw = hitKw ||
                                (!string.IsNullOrEmpty(matchedScheme.Brand) && matchedScheme.Brand.IndexOf(cleanKw, StringComparison.OrdinalIgnoreCase) >= 0) ||
                                (!string.IsNullOrEmpty(matchedScheme.Description) && matchedScheme.Description.IndexOf(cleanKw, StringComparison.OrdinalIgnoreCase) >= 0) ||
                                (!string.IsNullOrEmpty(matchedScheme.SchemeName) && matchedScheme.SchemeName.IndexOf(cleanKw, StringComparison.OrdinalIgnoreCase) >= 0);
                    }

                    // 未命中检索关键字则跳过
                    if (!hitKw) continue;
                }

                // 提取 DWG 轻量高清缩略图 (三级降级梯队安全解析)
                var preview = Services.DwgPreviewService.GetDwgPreview(file.FullPath);

                // 组装前端卡片展示模型
                var card = new SecondaryFolderDwgCardDto
                {
                    FileName = file.FileName,
                    DwgName = dwgName,
                    FolderName = curFolderName,
                    FolderPath = targetDir,
                    FullPath = file.FullPath,
                    PreviewBase64 = preview?.Base64Image ?? string.Empty,
                    UpdateTime = file.LastModified
                };

                // 若在数据库中成功命中方案数据，填充参数
                if (matchedScheme != null)
                {
                    card.IsMatched = true;
                    card.SchemeId = matchedScheme.Id;
                    card.SchemeName = !string.IsNullOrWhiteSpace(matchedScheme.SchemeName) ? matchedScheme.SchemeName : dwgName;
                    card.Brand = matchedScheme.Brand ?? string.Empty;
                    card.Description = matchedScheme.Description ?? string.Empty;
                    card.CrossDoorCount = matchedScheme.CrossDoorCount;
                    card.HoleSpec = matchedScheme.HoleSpec ?? string.Empty;
                    card.LaborCost = matchedScheme.LaborCost;
                    card.MaterialCost = matchedScheme.TotalMaterialCost;
                    card.TotalCost = matchedScheme.TotalCost;
                    card.LayoutDwgName = matchedScheme.CadDrawingName ?? string.Empty;
                    card.SchemeData = matchedScheme;
                    if (!string.IsNullOrWhiteSpace(matchedScheme.UpdatedAt))
                    {
                        card.UpdateTime = matchedScheme.UpdatedAt;
                    }
                }
                else
                {
                    // 未在数据库中匹配到记录时，参数保留默认值，前端提供编辑入库入口
                    card.IsMatched = false;
                    card.SchemeId = 0;
                    card.SchemeName = dwgName;
                    card.Brand = string.Empty;
                    card.Description = string.Empty;
                    card.CrossDoorCount = 0.0;
                    card.HoleSpec = string.Empty;
                    card.LaborCost = 0.0;
                    card.MaterialCost = 0.0;
                    card.TotalCost = 0.0;
                    card.LayoutDwgName = string.Empty;
                    // 初始化一个默认实体供编辑使用
                    card.SchemeData = new SecondarySchemeEntity
                    {
                        SchemeName = dwgName,
                        GroupName = curFolderName,
                        ApplicableCodes = new List<string> { dwgName },
                        CadDrawingName = dwgName
                    };
                }

                // 加入卡片列表
                cardList.Add(card);
            }
        }

        /// <summary>
        /// 调用本地默认关联的 AutoCAD 打开该 DWG 文件
        /// </summary>
        public (bool Success, string Message) OpenDwgInCad(string fullPath)
        {
            // 调用底层服务执行 CAD 打开
            return Services.DwgPreviewService.OpenInDefaultCad(fullPath);
        }

        /// <summary>
        /// 保存或更新二次回路方案至 personal_components.db
        /// </summary>
        public (bool Success, int SchemeId, string Message) SaveSecondaryScheme(string schemeJson)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(schemeJson))
                {
                    return (false, 0, "方案数据为空！");
                }

                // 反序列化方案对象
                var scheme = JsonSerializer.Deserialize<SecondarySchemeEntity>(schemeJson, JsonOptions);
                if (scheme == null)
                {
                    return (false, 0, "方案数据反序列化失败！");
                }

                // 方案主名称校验
                if (string.IsNullOrWhiteSpace(scheme.SchemeName))
                {
                    return (false, 0, "方案名称不能为空！");
                }

                // 确保至少有一个回路代号
                if (scheme.ApplicableCodes == null || scheme.ApplicableCodes.Count == 0)
                {
                    scheme.ApplicableCodes = new List<string> { scheme.SchemeName.Trim() };
                }

                // 保存至本地 SQLite 数据库
                int id = Services.PersonalComponentDbService.SaveSecondaryScheme(scheme);
                if (id > 0)
                {
                    return (true, id, "方案保存成功！");
                }
                return (false, 0, "保存到 SQLite 数据库失败，请检查日志！");
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[CloudSolutionController] SaveSecondaryScheme 异常: {ex.Message}");
                return (false, 0, $"保存异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 检索本地物料库元器件供编辑弹窗中选型添加使用
        /// </summary>
        public List<ComponentApiDto> SearchMaterialComponents(string keyword)
        {
            try
            {
                // 检索本地物料库
                return Services.PersonalComponentDbService.SearchComponents(
                    searchKeyword: keyword,
                    name: null,
                    current: null,
                    pole: null,
                    tripMode: null,
                    brand: null,
                    mustContainRules: null,
                    maxResults: 60
                );
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CloudSolutionController] SearchMaterialComponents 异常: {ex.Message}");
                return new List<ComponentApiDto>();
            }
        }
    }
}
