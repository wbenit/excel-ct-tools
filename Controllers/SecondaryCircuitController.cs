using System;
using System.Collections.Generic;
using System.Text.Json;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 二次图控制回路方案与 BOM 管理控制器 (供 Vue 3 前端 WebView2 交互调用)
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，配置与硬编码显式标明
    /// </summary>
    public class SecondaryCircuitController
    {
        // 声明通用的 JSON 序列化规范
        private static readonly JsonSerializerOptions JsonOpt = new JsonSerializerOptions
        {
            // 驼峰命名转换
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            // 忽略大小写
            PropertyNameCaseInsensitive = true,
            // 缩进排版
            WriteIndented = true
        };

        /// <summary>
        /// 获取所有二次方案列表数据
        /// </summary>
        /// <param name="keyword">方案名、回路或 CAD 图名搜索关键字</param>
        /// <param name="groupName">二次组名称</param>
        /// <returns>二次方案实体列表</returns>
        public List<SecondarySchemeEntity> GetSchemes(string? keyword = null, string? groupName = null)
        {
            // 调用本地 SQLite 数据服务查询并动态补齐最新价格
            return PersonalComponentDbService.GetAllSecondarySchemes(keyword, groupName);
        }

        /// <summary>
        /// 根据 ID 获取方案详情
        /// </summary>
        /// <param name="id">方案主键 ID</param>
        /// <returns>方案详情实体</returns>
        public SecondarySchemeEntity? GetSchemeById(int id)
        {
            // 调用服务层获取单条方案
            return PersonalComponentDbService.GetSecondarySchemeById(id);
        }

        /// <summary>
        /// 保存或更新二次回路方案
        /// </summary>
        /// <param name="schemeJson">前端传递的方案 JSON 字符串</param>
        /// <returns>操作结果包含成功状态、生成 ID 与消息</returns>
        public object SaveScheme(string schemeJson)
        {
            try
            {
                // 空值安全校验
                if (string.IsNullOrWhiteSpace(schemeJson))
                {
                    return new { success = false, message = "提交的方案数据为空！" };
                }

                // 反序列化为方案实体对象
                var scheme = JsonSerializer.Deserialize<SecondarySchemeEntity>(schemeJson, JsonOpt);
                if (scheme == null)
                {
                    return new { success = false, message = "方案数据格式反序列化失败！" };
                }

                // 方案主名称必填校验
                if (string.IsNullOrWhiteSpace(scheme.SchemeName))
                {
                    return new { success = false, message = "二次方案主名称不能为空！" };
                }

                // 确保至少有一个适用回路代号 (若未填则默认使用方案名)
                if (scheme.ApplicableCodes == null || scheme.ApplicableCodes.Count == 0)
                {
                    scheme.ApplicableCodes = new List<string> { scheme.SchemeName.Trim() };
                }

                // 核心门禁：执行回路代号全局唯一性防重检测 (不允许跨方案绑定相同代号)
                var conflict = PersonalComponentDbService.CheckApplicableCodeConflict(scheme.ApplicableCodes, scheme.Id);
                // 若检测到与已有方案产生同名冲突
                if (conflict != null)
                {
                    // 记录拦截日志并返回明确的拦截提示信息
                    LogHelper.WriteLog($"[SecondaryController] SaveScheme 拦截: 代号【{conflict.Code}】与方案【{conflict.SchemeName}】冲突");
                    // 返回友好错误提示给前端
                    return new { success = false, message = $"回路代号【{conflict.Code}】已被方案【{conflict.SchemeName}】绑定使用，全局不允许重复！请在当前方案中移除该代号后重试。" };
                }

                // 调用 SQLite 服务层保存
                int id = PersonalComponentDbService.SaveSecondaryScheme(scheme);
                if (id > 0)
                {
                    return new { success = true, id = id, message = "二次回路方案保存成功！" };
                }
                else
                {
                    return new { success = false, message = "保存至 SQLite 数据库失败，请检查日志！" };
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志并返回错误
                LogHelper.WriteLog($"[SecondaryController] SaveScheme 异常: {ex.Message}");
                return new { success = false, message = $"保存异常: {ex.Message}" };
            }
        }

        /// <summary>
        /// 删除指定方案
        /// </summary>
        /// <param name="id">方案 ID</param>
        /// <returns>删除结果</returns>
        public object DeleteScheme(int id)
        {
            // 校验 ID
            if (id <= 0)
            {
                return new { success = false, message = "无效的方案 ID！" };
            }

            // 执行删除
            bool ok = PersonalComponentDbService.DeleteSecondaryScheme(id);
            // 返回结果
            return new { success = ok, message = ok ? "删除方案成功！" : "删除方案失败或记录不存在！" };
        }

        /// <summary>
        /// 获取所有已存在的二次组分类列表
        /// </summary>
        /// <returns>二次组分类名称集合</returns>
        public List<string> GetSecondaryGroups()
        {
            // 调用服务层获取去重组名
            return PersonalComponentDbService.GetSecondaryGroups();
        }

        /// <summary>
        /// 从本地个人物料库 components 表中快速检索元器件 (支持品牌与关键字组合多维查询)
        /// 遵循更正 2 原则：从本地物料表选择，保证二次元器件价格全局一致
        /// </summary>
        /// <param name="keyword">型号或名称关键字 (可选)</param>
        /// <param name="brand">指定品牌 (默认二次元件)</param>
        /// <returns>匹配到的物料列表 (最多返回前 60 条)</returns>
        public object SearchMaterialComponents(string? keyword = null, string? brand = null)
        {
            try
            {
                // 清洗关键字参数
                string? cleanKw = string.IsNullOrWhiteSpace(keyword) ? null : keyword.Trim();
                // 清洗品牌参数
                string? cleanBrand = string.IsNullOrWhiteSpace(brand) ? null : brand.Trim();

                // 调用本地物料库多维查询服务 (按品牌、型号或名称检索)
                var components = PersonalComponentDbService.SearchComponents(
                    searchKeyword: cleanKw,
                    name: null,
                    current: null,
                    pole: null,
                    tripMode: null,
                    brand: cleanBrand,
                    mustContainRules: null,
                    maxResults: 60);

                // 投影为轻量 DTO 返回给前端
                var result = new List<object>();
                foreach (var c in components)
                {
                    // 构造前端所需的物料属性结构
                    result.Add(new
                    {
                        id = c.Id,
                        name = c.Name,
                        model = c.Model,
                        brand = c.Brand,
                        price = c.Price,
                        remark = c.Remark
                    });
                }

                // 返回匹配集合
                return result;
            }
            catch (Exception ex)
            {
                // 记录查询异常
                LogHelper.WriteLog($"[SecondaryController] SearchMaterialComponents 异常: {ex.Message}");
                return new List<object>();
            }
        }

        /// <summary>
        /// 获取本地物料库中已有的所有元器件品牌列表 (默认包含"二次元件")
        /// </summary>
        /// <returns>去重排序后的品牌列表</returns>
        public List<string> GetMaterialBrands()
        {
            // 初始化默认品牌列表，优先加入二次元件
            var brandSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "二次元件" }; // --硬编码-- 默认二次元件品牌

            try
            {
                // 从数据库统计接口获取全部已入库品牌
                var stats = PersonalComponentDbService.GetBrandStats();
                foreach (var item in stats)
                {
                    // 过滤非空品牌名称
                    if (!string.IsNullOrWhiteSpace(item.Brand))
                    {
                        // 收入品牌集合
                        brandSet.Add(item.Brand.Trim());
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[SecondaryController] GetMaterialBrands 异常: {ex.Message}");
            }

            // 返回品牌清单
            return brandSet.ToList();
        }

        /// <summary>
        /// 触发从当前 Excel 活动工作表中批量扫描识别并导入二次图库
        /// </summary>
        /// <returns>执行结果统计报文</returns>
        public object ImportFromActiveExcel()
        {
            // 调用 ExcelServices 批量导入业务
            var result = ExcelServices.ImportSecondarySchemesFromActiveSheet();
            // 返回规范化 JSON 结构
            return new
            {
                success = result.Success,
                message = result.Message,
                schemeCount = result.SchemeCount,
                bomCount = result.BomCount
            };
        }

        /// <summary>
        /// 获取当前配置的排布图与回路代号 DWG 本地图纸目录以及对应的图纸文件列表
        /// </summary>
        public object GetDwgDirectoriesAndFiles()
        {
            // 从全局配置管理器读取二次图配置
            var secConfig = ConfigManager.Instance.Current?.SecondaryCircuit ?? new SecondaryCircuitSettings();
            string layoutDir = secConfig.LayoutDwgDirectory ?? string.Empty;
            string circuitDir = secConfig.CircuitDwgDirectory ?? string.Empty;

            // 分别扫描各自目录下的 DWG 文件
            var layoutFiles = !string.IsNullOrWhiteSpace(layoutDir) && System.IO.Directory.Exists(layoutDir)
                ? DwgPreviewService.ScanDwgFiles(layoutDir)
                : new List<DwgFileInfo>();

            var circuitFiles = !string.IsNullOrWhiteSpace(circuitDir) && System.IO.Directory.Exists(circuitDir)
                ? DwgPreviewService.ScanDwgFiles(circuitDir)
                : new List<DwgFileInfo>();

            // 返回综合数据模型
            return new
            {
                layoutDirectory = layoutDir,
                circuitDirectory = circuitDir,
                layoutFiles = layoutFiles,
                circuitFiles = circuitFiles
            };
        }

        /// <summary>
        /// 设置并持久化指定的 DWG 目录
        /// </summary>
        /// <param name="target">"layout" 表示排布图，"circuit" 表示回路代号</param>
        /// <param name="newDirectory">新目录物理路径</param>
        public object SetDwgDirectory(string target, string newDirectory)
        {
            if (string.Equals(target, "layout", StringComparison.OrdinalIgnoreCase))
            {
                // 更新二次排布图目录
                ConfigManager.Instance.UpdateSecondaryDwgDirectories(newDirectory, null);
            }
            else
            {
                // 更新回路代号图纸目录
                ConfigManager.Instance.UpdateSecondaryDwgDirectories(null, newDirectory);
            }

            // 扫描新目录下的文件并返回
            var files = DwgPreviewService.ScanDwgFiles(newDirectory);
            return new
            {
                success = true,
                target = target,
                directory = newDirectory,
                files = files
            };
        }

        /// <summary>
        /// 扫描指定本地目录下的 DWG 文件列表
        /// </summary>
        /// <param name="dirPath">目标物理路径</param>
        public List<DwgFileInfo> ScanDwgFiles(string dirPath)
        {
            // 直接委托服务层扫描
            return DwgPreviewService.ScanDwgFiles(dirPath);
        }

        /// <summary>
        /// 扫描指定目录下的子文件夹列表与 DWG 图纸文件，支持层级展开与溯源
        /// </summary>
        /// <param name="dirPath">目标物理路径</param>
        public DirectoryHierarchyResult ScanDirectoryHierarchy(string dirPath)
        {
            // 委托底层服务执行综合层级扫描
            return DwgPreviewService.ScanDirectoryHierarchy(dirPath);
        }

        /// <summary>
        /// 全局定位回路图纸所在的物理路径与父目录 (支持前端点击元件组跨文件夹反显)
        /// </summary>
        /// <param name="rootDir">搜索根目录 (若为空则自动取已配置的回路图纸目录)</param>
        /// <param name="codeOrName">图号或文件名</param>
        /// <returns>定位结果与文件实体</returns>
        public object LocateDwgFile(string? rootDir, string codeOrName)
        {
            // 若前端未传入根目录，自动使用已持久化保存的回路图纸根目录
            string searchRoot = rootDir ?? string.Empty;
            if (string.IsNullOrWhiteSpace(searchRoot))
            {
                searchRoot = ConfigManager.Instance.Current?.SecondaryCircuit?.CircuitDwgDirectory ?? string.Empty;
            }

            // 委托底层服务进行全局递归穿透定位
            var (success, parentDir, fileInfo, msg) = DwgPreviewService.LocateDwgInDirectory(searchRoot, codeOrName);
            return new
            {
                success = success,
                parentDir = parentDir,
                dwgItem = fileInfo,
                message = msg
            };
        }

        /// <summary>
        /// 获取 DWG 图纸的原始二进制 Base64 数据供前端真实矢量视口渲染
        /// </summary>
        /// <param name="filePath">物理路径或图纸名称</param>
        /// <param name="preferredType">优先查找类型 ("layout" 或 "circuit")</param>
        public object GetDwgBinary(string filePath, string? preferredType = null)
        {
            // 智能补全图纸真实物理绝对路径
            string targetPath = filePath;
            if (!System.IO.File.Exists(targetPath))
            {
                var secConfig = ConfigManager.Instance.Current?.SecondaryCircuit;
                string? layoutDir = secConfig?.LayoutDwgDirectory;
                string? circuitDir = secConfig?.CircuitDwgDirectory;
                string searchFileName = targetPath.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase)
                    ? targetPath
                    : targetPath + ".dwg";

                if (string.Equals(preferredType, "circuit", StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrWhiteSpace(circuitDir) && System.IO.File.Exists(System.IO.Path.Combine(circuitDir, searchFileName)))
                        targetPath = System.IO.Path.Combine(circuitDir, searchFileName);
                    else if (!string.IsNullOrWhiteSpace(layoutDir) && System.IO.File.Exists(System.IO.Path.Combine(layoutDir, searchFileName)))
                        targetPath = System.IO.Path.Combine(layoutDir, searchFileName);
                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(layoutDir) && System.IO.File.Exists(System.IO.Path.Combine(layoutDir, searchFileName)))
                        targetPath = System.IO.Path.Combine(layoutDir, searchFileName);
                    else if (!string.IsNullOrWhiteSpace(circuitDir) && System.IO.File.Exists(System.IO.Path.Combine(circuitDir, searchFileName)))
                        targetPath = System.IO.Path.Combine(circuitDir, searchFileName);
                }
            }

            // 读取底层二进制字节流并转为 Base64
            var (success, base64, err) = DwgPreviewService.ReadDwgBinaryBase64(targetPath);
            return new
            {
                success = success,
                fileName = Path.GetFileName(targetPath),
                fullPath = targetPath,
                base64Data = base64,
                errorMessage = err
            };
        }

        /// <summary>
        /// 获取单张 DWG 图纸的缩略图 Base64 预览及文件属性
        /// </summary>
        /// <param name="filePath">物理全路径，或相对于排布图/回路图目录的文件名</param>
        /// <param name="preferredType">可选优先查找类型: "layout" 或 "circuit"</param>
        public DwgPreviewResult GetDwgPreview(string filePath, string? preferredType = null)
        {
            string targetPath = filePath;
            // 若传入的仅为文件名或未包含盘符，尝试从配置的两个目录中补齐绝对路径
            if (!System.IO.File.Exists(targetPath))
            {
                var secConfig = ConfigManager.Instance.Current?.SecondaryCircuit;
                string? layoutDir = secConfig?.LayoutDwgDirectory;
                string? circuitDir = secConfig?.CircuitDwgDirectory;

                // 若缺少 .dwg 后缀则自动补全
                string searchFileName = targetPath.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase)
                    ? targetPath
                    : targetPath + ".dwg";

                // 按照指定偏好或顺序匹配存在性
                if (string.Equals(preferredType, "circuit", StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrWhiteSpace(circuitDir) && System.IO.File.Exists(System.IO.Path.Combine(circuitDir, searchFileName)))
                        targetPath = System.IO.Path.Combine(circuitDir, searchFileName);
                    else if (!string.IsNullOrWhiteSpace(layoutDir) && System.IO.File.Exists(System.IO.Path.Combine(layoutDir, searchFileName)))
                        targetPath = System.IO.Path.Combine(layoutDir, searchFileName);
                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(layoutDir) && System.IO.File.Exists(System.IO.Path.Combine(layoutDir, searchFileName)))
                        targetPath = System.IO.Path.Combine(layoutDir, searchFileName);
                    else if (!string.IsNullOrWhiteSpace(circuitDir) && System.IO.File.Exists(System.IO.Path.Combine(circuitDir, searchFileName)))
                        targetPath = System.IO.Path.Combine(circuitDir, searchFileName);
                }
            }

            // 调度服务层提取缩略图
            return DwgPreviewService.GetDwgPreview(targetPath);
        }

        /// <summary>
        /// 外部调起本地系统关联的 AutoCAD 打开该 DWG 文件
        /// </summary>
        public object OpenInCad(string filePath, string? preferredType = null)
        {
            var preview = GetDwgPreview(filePath, preferredType);
            var (success, msg) = DwgPreviewService.OpenInDefaultCad(preview.FullPath);
            return new { success = success, message = msg };
        }

        /// <summary>
        /// 在 Windows 资源管理器中定位并选中该 DWG 文件
        /// </summary>
        public object OpenInExplorer(string filePath, string? preferredType = null)
        {
            var preview = GetDwgPreview(filePath, preferredType);
            var (success, msg) = DwgPreviewService.OpenInExplorer(preview.FullPath);
            return new { success = success, message = msg };
        }
    }
}
