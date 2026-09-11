using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;
using ExcelAddInDemo.Forms;
using ExcelAddInDemo.Models;
using ExcelDna.Integration;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：云方案中心（行业方案、企业方案、BOM倍增插入与另存方案）
    /// </summary>
    public static partial class ExcelServices
    {
        // 存储本地预置与企业自建方案的 JSON 物理文件名称
        private const string CloudSchemesFileName = "cloud_schemes.json"; // --硬编码-- 方案持久化数据文件名

        // 方案默认存储子目录名称
        private const string DataDirectoryName = "data"; // --硬编码-- 数据持久化子目录

        // 维持云方案窗口的静态单例引用，杜绝重复多开
        private static CloudSolutionForm? _cloudSolutionForm = null;

        // 线程安全互斥锁
        private static readonly object _schemeLock = new object();

        // 内存中缓存的方案集合
        private static List<CloudSchemeDetail>? _cachedSchemes = null;

        // 通用 JSON 序列化配置
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 启动并弹出“云方案中心”现代化窗口 (非模态弹出，安全依附 Excel 主窗口)
        /// 彻底根除独立后台线程 COM 跨线程死锁与 Excel 硬禁用风险
        /// </summary>
        public static void ShowCloudSolutionDialog()
        {
            try
            {
                // 若窗体已打开且未释放，直接还原、置顶并聚焦
                if (_cloudSolutionForm != null && !_cloudSolutionForm.IsDisposed)
                {
                    // 若处于最小化状态则恢复正常尺寸
                    if (_cloudSolutionForm.WindowState == FormWindowState.Minimized)
                    {
                        _cloudSolutionForm.WindowState = FormWindowState.Normal;
                    }
                    // 推至最前台
                    _cloudSolutionForm.BringToFront();
                    // 激活窗口焦点
                    _cloudSolutionForm.Activate();
                    return;
                }

                // 实例化全新云方案窗口
                _cloudSolutionForm = new CloudSolutionForm();
                // 绑定窗口关闭事件，置空单例句柄
                _cloudSolutionForm.FormClosed += (s, e) => _cloudSolutionForm = null;

                // 安全获取 Excel 主窗口 HWND 句柄
                IntPtr excelHwnd = ExcelDnaSafeAccessor.GetWindowHandle();
                if (excelHwnd != IntPtr.Zero)
                {
                    // 作为 Excel 主窗口的 Owned 窗口非模态展示
                    _cloudSolutionForm.Show(new ExcelWin32Window(excelHwnd));
                }
                else
                {
                    // 独立非模态弹出
                    _cloudSolutionForm.Show();
                }
            }
            catch (Exception ex)
            {
                // 记录启动异常日志
                LogHelper.WriteLog($"[CloudSolution] 启动方案中心窗体异常: {ex.Message}");
                MessageBox.Show($"启动云方案中心失败: {ex.Message}", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 获取云方案持久化数据文件的绝对物理路径
        /// </summary>
        private static string GetCloudSchemesFilePath()
        {
            // 获取插件专属 data 目录物理路径 (优先检测 XLL 真实所在目录)
            string dataDir = Tool.GetAppDataDirectory();
            // 确保 data 目录存在
            if (!Directory.Exists(dataDir))
            {
                // 递归创建目录
                Directory.CreateDirectory(dataDir);
            }
            // 返回方案数据完整文件路径
            return Path.Combine(dataDir, CloudSchemesFileName);
        }

        /// <summary>
        /// 加载全量方案列表（若本地文件不存在则自动生成一套丰富高质量的官方预置方案）
        /// </summary>
        public static List<CloudSchemeDetail> GetOrInitAllSchemes()
        {
            lock (_schemeLock)
            {
                // 如果内存已有缓存则直接返回
                if (_cachedSchemes != null && _cachedSchemes.Count > 0)
                {
                    return _cachedSchemes;
                }

                // 获取数据文件路径
                string filePath = GetCloudSchemesFilePath();
                // 判断磁盘上是否存在方案数据文件
                if (File.Exists(filePath))
                {
                    try
                    {
                        // 读取文本内容
                        string json = File.ReadAllText(filePath);
                        // 反序列化为方案详情集合
                        var list = JsonSerializer.Deserialize<List<CloudSchemeDetail>>(json, JsonOptions);
                        if (list != null && list.Count > 0)
                        {
                            // 赋值缓存并返回
                            _cachedSchemes = list;
                            return _cachedSchemes;
                        }
                    }
                    catch (Exception ex)
                    {
                        // 记录反序列化失败日志并转为走预置兜底
                        LogHelper.WriteLog($"[CloudSolution] 读取方案文件异常: {ex.Message}，将初始化默认方案");
                    }
                }

                // 本地无数据或损坏，生成内置高质量预置方案
                _cachedSchemes = GenerateDefaultPresetSchemes();
                // 异步落盘保存
                SaveAllSchemesToDisk(_cachedSchemes);
                // 返回初始化方案列表
                return _cachedSchemes;
            }
        }

        /// <summary>
        /// 将内存方案集合序列化并写入磁盘持久化存储
        /// </summary>
        public static void SaveAllSchemesToDisk(List<CloudSchemeDetail> schemes)
        {
            lock (_schemeLock)
            {
                try
                {
                    // 获取目标路径
                    string filePath = GetCloudSchemesFilePath();
                    // 序列化为美化排版的 JSON 字符串
                    string json = JsonSerializer.Serialize(schemes, JsonOptions);
                    // 写入文本文件
                    File.WriteAllText(filePath, json);
                }
                catch (Exception ex)
                {
                    // 记录保存异常日志
                    LogHelper.WriteLog($"[CloudSolution] 方案持久化写入磁盘失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 分页多维检索方案列表（支持行业/企业、一次/二次/收藏过滤与热门排序）
        /// </summary>
        public static CloudSchemePageResult QuerySchemes(CloudSchemeQueryDto query)
        {
            // 加载全量方案
            var all = GetOrInitAllSchemes();
            // 开启 LINQ 管道过滤
            var q = all.AsEnumerable();

            // 1. 范围过滤 (行业方案 / 企业方案)
            if (!string.IsNullOrWhiteSpace(query.Scope))
            {
                // 严格匹配范围
                q = q.Where(s => string.Equals(s.Scope, query.Scope, StringComparison.OrdinalIgnoreCase));
            }

            // 2. 业务类别过滤 (一次方案 / 二次方案 / 我的收藏)
            if (string.Equals(query.SchemeType, CloudSchemeType.Favorite, StringComparison.OrdinalIgnoreCase))
            {
                // 仅筛选收藏标记为 true 的项
                q = q.Where(s => s.IsFavorite);
            }
            else if (!string.IsNullOrWhiteSpace(query.SchemeType))
            {
                // 匹配具体的一次或二次类型
                q = q.Where(s => string.Equals(s.SchemeType, query.SchemeType, StringComparison.OrdinalIgnoreCase));
            }

            // 3. 柜型过滤
            if (!string.IsNullOrWhiteSpace(query.CabinetModel) && query.CabinetModel != "全部") // --硬编码-- 全部不作过滤
            {
                // 包含柜型判定
                q = q.Where(s => s.CabinetModel.IndexOf(query.CabinetModel, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            // 4. 关键字检索 (方案名称、描述、作者或分类)
            if (!string.IsNullOrWhiteSpace(query.Keyword))
            {
                // 提取关键字去除首尾空格
                string kw = query.Keyword.Trim();
                // 满足任意维度模糊匹配
                q = q.Where(s => (s.SchemeName != null && s.SchemeName.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                              || (s.CabinetModel != null && s.CabinetModel.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                              || (s.CreatorName != null && s.CreatorName.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                              || (s.Description != null && s.Description.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0));
            }

            // 5. 排序规则
            if (string.Equals(query.SortBy, "latest", StringComparison.OrdinalIgnoreCase)) // --硬编码-- 最新更新排序
            {
                // 按更新时间降序
                q = q.OrderByDescending(s => s.UpdateTime);
            }
            else
            {
                // 默认热门排序：置顶在前，其次按浏览量+收藏量综合热度
                q = q.OrderByDescending(s => !string.IsNullOrEmpty(s.Tag))
                     .ThenByDescending(s => s.ViewCount + s.FavCount * 5);
            }

            // 获取过滤后的总记录数
            int total = q.Count();
            // 计算分页偏移量
            int pageIndex = Math.Max(1, query.PageIndex);
            // 确保每页条数合理
            int pageSize = query.PageSize > 0 ? query.PageSize : CloudSchemeDefaults.DefaultPageSize;

            // 提取分页子集
            var pagedItems = q.Skip((pageIndex - 1) * pageSize)
                              .Take(pageSize)
                              .Select(s => (CloudSchemeItem)s)
                              .ToList();

            // 构造分页结果对象并返回
            return new CloudSchemePageResult
            {
                Items = pagedItems,
                TotalCount = total,
                PageIndex = pageIndex,
                PageSize = pageSize
            };
        }

        /// <summary>
        /// 根据方案 ID 获取单个方案的完整详情（含多图纸与 BOM 明细）
        /// </summary>
        public static CloudSchemeDetail? GetSchemeDetailById(string id)
        {
            // 入参校验
            if (string.IsNullOrWhiteSpace(id)) return null;
            // 获取全量方案
            var all = GetOrInitAllSchemes();
            // 按 ID 精准查询
            var scheme = all.FirstOrDefault(s => s.Id == id);
            if (scheme != null)
            {
                // 增加浏览量并防并发
                scheme.ViewCount++;
                // 动态重算总计金额
                if (scheme.BomItems != null)
                {
                    scheme.TotalAmount = scheme.BomItems.Sum(b => (decimal)b.Quantity * b.QuotePrice);
                }
            }
            // 返回方案详情
            return scheme;
        }

        /// <summary>
        /// 切换方案的星标收藏状态
        /// </summary>
        public static bool ToggleSchemeFavorite(string id)
        {
            // 查询方案
            var scheme = GetSchemeDetailById(id);
            if (scheme == null) return false;

            // 状态反转
            scheme.IsFavorite = !scheme.IsFavorite;
            // 收藏数自增或自减
            scheme.FavCount = Math.Max(0, scheme.FavCount + (scheme.IsFavorite ? 1 : -1));
            // 异步写回磁盘
            SaveAllSchemesToDisk(GetOrInitAllSchemes());
            // 返回最新状态
            return scheme.IsFavorite;
        }

        /// <summary>
        /// 保存或更新企业方案
        /// </summary>
        public static bool SaveEnterpriseScheme(CloudSchemeDetail scheme)
        {
            // 基础非空校验
            if (scheme == null || string.IsNullOrWhiteSpace(scheme.SchemeName)) return false;

            lock (_schemeLock)
            {
                // 获取方案集合
                var all = GetOrInitAllSchemes();
                // 强制范围标记为企业方案
                scheme.Scope = CloudSchemeScope.Enterprise;
                // 更新修改时间
                scheme.UpdateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

                // 检查是否为已有方案更新
                int existIndex = all.FindIndex(s => s.Id == scheme.Id);
                if (existIndex >= 0)
                {
                    // 替换已有记录
                    all[existIndex] = scheme;
                }
                else
                {
                    // 若是新建确保有唯一 ID
                    if (string.IsNullOrWhiteSpace(scheme.Id))
                    {
                        scheme.Id = Guid.NewGuid().ToString("N");
                    }
                    // 插入到集合首位
                    all.Insert(0, scheme);
                }

                // 存盘持久化
                SaveAllSchemesToDisk(all);
                return true;
            }
        }

        /// <summary>
        /// 删除指定 ID 的企业自建方案
        /// </summary>
        public static bool DeleteEnterpriseScheme(string id)
        {
            // 入参校验
            if (string.IsNullOrWhiteSpace(id)) return false;

            lock (_schemeLock)
            {
                // 获取方案集合
                var all = GetOrInitAllSchemes();
                // 仅允许删除企业方案，杜绝删除系统行业方案
                int removed = all.RemoveAll(s => s.Id == id && s.Scope == CloudSchemeScope.Enterprise);
                if (removed > 0)
                {
                    // 存盘写回磁盘
                    SaveAllSchemesToDisk(all);
                    return true;
                }
                return false;
            }
        }

        /// <summary>
        /// 将用户在方案中心勾选确认的 BOM 清单一次性高效插入当前活动 Excel 分类表中
        /// 遵循规则 7：读写多个区域采用内存数组一次性写入；严格维护箱柜元器件区间
        /// </summary>
        public static (bool Success, string Message) InsertSchemeBomToExcel(SchemeInsertToExcelDto dto)
        {
            try
            {
                // 校验选中项
                if (dto == null || dto.SelectedBomItems == null || dto.SelectedBomItems.Count == 0)
                {
                    return (false, "待插入的 BOM 明细列表为空！");
                }

                // 获取 Excel 上下文
                var context = Tool.GetActiveExcelContext(null, null);
                if (context == null)
                {
                    return (false, "未能连接到活动 Excel 工作表，请确保 Excel 已打开且处于编辑状态！");
                }

                dynamic app = context.App;
                dynamic wb = context.Wb;
                dynamic ws = context.Sheet;

                // 检查是否要在当前选中箱柜中追加元器件
                if (dto.InsertMode == "currentCabinet")
                {
                    // 执行在当前选定箱柜元器件区插入逻辑
                    return AppendBomToCurrentCabinet(app, ws, dto);
                }
                else
                {
                    // 默认新建箱柜并将 BOM 写入新箱柜中
                    return CreateNewCabinetWithBom(app, ws, dto);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[CloudSolution] 方案写入 Excel 异常: {ex.Message}");
                return (false, $"写入 Excel 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 新建箱柜并将 BOM 矩阵一次性完整写入
        /// </summary>
        private static (bool Success, string Message) CreateNewCabinetWithBom(dynamic app, dynamic ws, SchemeInsertToExcelDto dto)
        {
            // 查询方案基础属性
            var scheme = GetSchemeDetailById(dto.SchemeId);
            string cabName = scheme?.SchemeName ?? "云方案箱柜";

            // 1. 调用已有的模板复制服务创建新箱柜块
            var cabInfo = CopyCabinetDetailFromTemplate(ws, 0, 0, cabName, app);
            if (cabInfo == null)
            {
                return (false, "创建新箱柜模板结构失败，请检查工作表格式！");
            }

            int detRow = cabInfo.DetailRow;
            int subsumRow = cabInfo.SubsumRow;
            int compStartRow = detRow + 2; // 规则 6: 元器件起始行为 Cab_Det + 2
            int compEndRow = subsumRow - 1; // 规则 6: 元器件终止行为 Cab_Subsum - 1
            int availableRows = compEndRow - compStartRow + 1;

            var items = dto.SelectedBomItems;
            int reqCount = items.Count;

            // 2. 检查空间，若元器件数量多于区域行数，先插入行 (严格遵循业务规则 6)
            if (reqCount > availableRows)
            {
                int rowsToInsert = reqCount - availableRows;
                // 在小计行上方插入空行
                dynamic insertRange = ws.Range[$"A{subsumRow}:A{subsumRow + rowsToInsert - 1}"];
                insertRange.EntireRow.Insert(-4121); // xlDown
                // 插入行后小计行下移
                subsumRow += rowsToInsert;
                compEndRow = subsumRow - 1;
            }

            // 3. 构建内存二维数组，准备一次性写入 (遵循规则 7)
            // 列定义对应常见分类表列: B(序号), C(名称), D(型号), E(品牌), F(单位), G(单价), H(数量)
            object[,] dataMatrix = new object[reqCount, 8];
            int loopMultiplier = dto.LoopMultiplier > 0 ? dto.LoopMultiplier : 1;

            for (int i = 0; i < reqCount; i++)
            {
                var item = items[i];
                // 计算乘算后的数量 (WL 勾选则回路数倍增)
                double finalQty = item.IsWlDoubled ? (item.Quantity * loopMultiplier) : item.Quantity;

                dataMatrix[i, 0] = i + 1;                  // B 列: 序号
                // C 列: 元件名称 (标准 Name 字段)
                dataMatrix[i, 1] = item.Name ?? "";
                // D 列: 规格型号 (标准 Model 字段)
                dataMatrix[i, 2] = item.Model ?? "";
                dataMatrix[i, 3] = item.Brand ?? "";         // E 列: 品牌
                dataMatrix[i, 4] = item.Unit ?? "台";        // F 列: 单位
                dataMatrix[i, 5] = (double)item.QuotePrice;  // G 列: 单价
                dataMatrix[i, 6] = finalQty;                 // H 列: 数量
                dataMatrix[i, 7] = (double)item.QuotePrice * finalQty; // I 列: 合价
            }

            // 4. 将构建完毕的二维数组一次性刷入工作表
            dynamic writeRange = ws.Range[$"B{compStartRow}:I{compStartRow + reqCount - 1}"];
            writeRange.Value2 = dataMatrix;

            return (true, $"已成功新建箱柜【{cabName}】并写入 {reqCount} 项元器件清单！");
        }

        /// <summary>
        /// 在当前光标所在的已有箱柜元器件区间末尾追加 BOM 项
        /// </summary>
        private static (bool Success, string Message) AppendBomToCurrentCabinet(dynamic app, dynamic ws, SchemeInsertToExcelDto dto)
        {
            // 扫描定位当前活动行所属箱柜
            dynamic activeCell = app.ActiveCell;
            int currentRow = activeCell?.Row ?? 0;
            if (currentRow <= 0) return (false, "未检测到有效活动单元格！");

            // 检索箱柜信息与小计行位置
            dynamic? activeWb = app.ActiveWorkbook;
            var cabinets = Tool.GetSheetValidCabinets(ws, activeWb);
            int targetK = 0;
            int detRow = 0;
            int subsumRow = 0;

            if (cabinets != null)
            {
                foreach (var cab in cabinets)
                {
                    int k = cab.Key;
                    var (_, cDet, cSub, cTol) = Tool.FindStandardCategoryRowIndexes((object)ws, k);
                    if (currentRow >= cDet && currentRow <= cTol)
                    {
                        targetK = k;
                        detRow = cDet;
                        subsumRow = cSub;
                        break;
                    }
                }
            }

            if (targetK == 0 || detRow == 0 || subsumRow == 0)
            {
                return (false, "当前光标未位于任何有效箱柜明细区域内，请先点击目标箱柜后再试！");
            }

            int compStartRow = detRow + 2;
            int compEndRow = subsumRow - 1;

            var items = dto.SelectedBomItems;
            int reqCount = items.Count;
            int loopMultiplier = dto.LoopMultiplier > 0 ? dto.LoopMultiplier : 1;

            // 在小计行上方直接插入 N 行空白行用于承载新物料 (遵循规则 6)
            dynamic insertRange = ws.Range[$"A{subsumRow}:A{subsumRow + reqCount - 1}"];
            insertRange.EntireRow.Insert(-4121); // xlDown

            // 构建二维数据矩阵
            object[,] dataMatrix = new object[reqCount, 8];
            for (int i = 0; i < reqCount; i++)
            {
                var item = items[i];
                double finalQty = item.IsWlDoubled ? (item.Quantity * loopMultiplier) : item.Quantity;

                dataMatrix[i, 0] = i + 1;
                // C 列: 元件名称 (标准 Name 字段)
                dataMatrix[i, 1] = item.Name ?? "";
                // D 列: 规格型号 (标准 Model 字段)
                dataMatrix[i, 2] = item.Model ?? "";
                dataMatrix[i, 3] = item.Brand ?? "";
                dataMatrix[i, 4] = item.Unit ?? "台";
                dataMatrix[i, 5] = (double)item.QuotePrice;
                dataMatrix[i, 6] = finalQty;
                dataMatrix[i, 7] = (double)item.QuotePrice * finalQty;
            }

            // 写入插入好的新区间中
            int writeStartRow = subsumRow; // 原小计行位置现已成为新插入空行的起始
            dynamic writeRange = ws.Range[$"B{writeStartRow}:I{writeStartRow + reqCount - 1}"];
            writeRange.Value2 = dataMatrix;

            // 规则 6: 新增元器件行后，全表定义名称与计费区域联动自适应刷新
            Tool.FixAndFillCabinetNamesForSheet(ws);

            return (true, $"已成功在当前箱柜中追加 {reqCount} 项元器件明细！");
        }

        /// <summary>
        /// 生成高质量官方预置方案母版（包含高压、低压、二次测控典型方案）
        /// </summary>
        private static List<CloudSchemeDetail> GenerateDefaultPresetSchemes()
        {
            var list = new List<CloudSchemeDetail>();

            // 1. KYN28-12 PT柜 (一次方案)
            list.Add(new CloudSchemeDetail
            {
                Id = "sch_ind_001",
                Scope = CloudSchemeScope.Industry,
                SchemeType = CloudSchemeType.Primary,
                SchemeName = "KYN28-12 PT柜(630A)",
                CabinetModel = "", // 对齐图2图3：箱柜型号留空
                Dimensions = "",   // 对齐图2图3：箱柜尺寸留空
                ThumbnailUrl = "https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=500&q=80",
                DrawingUrls = new List<string>
                {
                    "https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=1200&q=80",
                    "https://images.unsplash.com/photo-1581092335397-9583fe92d232?w=1200&q=80"
                },
                Description = "KYN28-12型户内金属铠装移开式开关设备，适用于3.6~12kV三相交流50Hz单母线及单母线分段系统，主要用于PT电压测量与母线避雷保护。",
                CategoryId = "cat_high_voltage",
                CategoryName = "高压开关柜方案",
                ViewCount = 12250,
                FavCount = 3680,
                IsFavorite = true,
                CreatorName = "电气天下", // 对齐图2图3修改人
                UpdateTime = "2017-12-08 09:44:28", // 对齐图2图3修改时间
                Tag = "顶",
                BomItems = new List<CloudSchemeBomItem>
                {
                    // 1. PT手车 (未选中，对齐图2)
                    new CloudSchemeBomItem { SortOrder = 1, Selected = false, Name = "PT手车", Model = "630A", Brand = "", Unit = "台", Quantity = 1.0, IsWlDoubled = true, CatalogPrice = 3500m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 3500m, TotalPrice = 3500m, Category = "高压手车", MaterialCode = "", Origin = "", Remark = "" },
                    // 2. 电压互感器 (未选中，对齐图2)
                    new CloudSchemeBomItem { SortOrder = 2, Selected = false, Name = "电压互感器", Model = "JDZX10-10 10/0.1kV", Brand = "大连第一互感器", Unit = "台", Quantity = 3.0, IsWlDoubled = true, CatalogPrice = 2200m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 2200m, TotalPrice = 6600m, Category = "互感器", MaterialCode = "", Origin = "", Remark = "" },
                    // 3. 高压熔断器 (选中，对齐图2)
                    new CloudSchemeBomItem { SortOrder = 3, Selected = true, Name = "高压熔断器", Model = "XRNP-10/0.5A", Brand = "上海一开", Unit = "台", Quantity = 3.0, IsWlDoubled = true, CatalogPrice = 44.94m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 44.94m, TotalPrice = 134.82m, Category = "熔断器", MaterialCode = "", Origin = "", Remark = "" },
                    // 4. 避雷器 (选中，对齐图2)
                    new CloudSchemeBomItem { SortOrder = 4, Selected = true, Name = "避雷器", Model = "HY5WZ2-17/45", Brand = "大连伏安", Unit = "台", Quantity = 3.0, IsWlDoubled = true, CatalogPrice = 360m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 360m, TotalPrice = 1080m, Category = "避雷器", MaterialCode = "", Origin = "", Remark = "" },
                    // 5. 带电显示器+传感器 (选中，对齐图2)
                    new CloudSchemeBomItem { SortOrder = 5, Selected = true, Name = "带电显示器+...", Model = "DXN11/Q1", Brand = "大连丰和", Unit = "套", Quantity = 1.0, IsWlDoubled = true, CatalogPrice = 258m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 258m, TotalPrice = 258m, Category = "显示装置", MaterialCode = "", Origin = "", Remark = "" },
                    // 6. 消谐器 (选中，对齐图2)
                    new CloudSchemeBomItem { SortOrder = 6, Selected = true, Name = "消谐器", Model = "HDXX-10", Brand = "安徽徽电科技", Unit = "只", Quantity = 1.0, IsWlDoubled = true, CatalogPrice = 1800m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 1800m, TotalPrice = 1800m, Category = "消谐装置", MaterialCode = "", Origin = "", Remark = "" },
                    // 7. 触头盒 (选中，对齐图2)
                    new CloudSchemeBomItem { SortOrder = 7, Selected = true, Name = "触头盒", Model = "630A", Brand = "", Unit = "只", Quantity = 3.0, IsWlDoubled = true, CatalogPrice = 120m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 120m, TotalPrice = 360m, Category = "绝缘件", MaterialCode = "", Origin = "", Remark = "" },
                    // 8. 穿墙套管 (选中，对齐图2)
                    new CloudSchemeBomItem { SortOrder = 8, Selected = true, Name = "穿墙套管", Model = "", Brand = "", Unit = "只", Quantity = 3.0, IsWlDoubled = true, CatalogPrice = 220m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 220m, TotalPrice = 660m, Category = "绝缘件", MaterialCode = "", Origin = "", Remark = "" },
                    // 9. 绝缘子 (选中，对齐图2)
                    new CloudSchemeBomItem { SortOrder = 9, Selected = true, Name = "绝缘子", Model = "", Brand = "", Unit = "只", Quantity = 3.0, IsWlDoubled = true, CatalogPrice = 60m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 60m, TotalPrice = 180m, Category = "绝缘件", MaterialCode = "", Origin = "", Remark = "" },
                    // 10. 不锈钢安装板 (选中，对齐图2)
                    new CloudSchemeBomItem { SortOrder = 10, Selected = true, Name = "不锈钢安装板", Model = "", Brand = "", Unit = "套", Quantity = 1.0, IsWlDoubled = true, CatalogPrice = 800m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 800m, TotalPrice = 800m, Category = "结构件", MaterialCode = "", Origin = "", Remark = "" },
                    // 11. 铜排 (选中，对齐图2)
                    new CloudSchemeBomItem { SortOrder = 11, Selected = true, Name = "铜排", Model = "TMY 3*40*6", Brand = "", Unit = "公斤", Quantity = 1.71, IsWlDoubled = true, CatalogPrice = 48m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 48m, TotalPrice = 82.08m, Category = "母排", MaterialCode = "", Origin = "", Remark = "" },
                    // 12. 微机保护测控装置 (使整单总计精准吻合图2图3中显示的 28222.42)
                    new CloudSchemeBomItem { SortOrder = 12, Selected = true, Name = "微机保护测控装置", Model = "WDZ-5200", Brand = "国电南自", Unit = "台", Quantity = 1.0, IsWlDoubled = false, CatalogPrice = 8500m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 8500m, TotalPrice = 8500m, Category = "微机保护", MaterialCode = "", Origin = "南京", Remark = "" },
                    // 13. 二次控制母线辅材 (使整单总计精准吻合 28222.42)
                    new CloudSchemeBomItem { SortOrder = 13, Selected = true, Name = "二次母线及辅材套件", Model = "KYN28-12配套", Brand = "标准配套", Unit = "套", Quantity = 1.0, IsWlDoubled = false, CatalogPrice = 4247.52m, Discount = 1.0m, QuoteDiscount = 1.0m, QuotePrice = 4247.52m, TotalPrice = 4247.52m, Category = "二次辅材", MaterialCode = "", Origin = "无锡", Remark = "" }
                }
            });

            // 2. MNS 低压进线柜 (一次方案)
            list.Add(new CloudSchemeDetail
            {
                Id = "sch_ind_002",
                Scope = CloudSchemeScope.Industry,
                SchemeType = CloudSchemeType.Primary,
                SchemeName = "MNS 低压主进线柜(2000A/3P)",
                CabinetModel = "MNS",
                Dimensions = "800*1000*2200",
                ThumbnailUrl = "https://images.unsplash.com/photo-1581092335397-9583fe92d232?w=500&q=80",
                DrawingUrls = new List<string>
                {
                    "https://images.unsplash.com/photo-1581092335397-9583fe92d232?w=1200&q=80"
                },
                Description = "低压抽出式开关柜进线单元，配置智能型框架万能断路器，支持双电源机械互锁与电气联锁保护。",
                CategoryId = "cat_low_voltage",
                CategoryName = "低压进线方案",
                ViewCount = 8940,
                FavCount = 2150,
                IsFavorite = false,
                CreatorName = "电气天下",
                UpdateTime = "2026-09-07 16:45",
                Tag = "热",
                BomItems = new List<CloudSchemeBomItem>
                {
                    new CloudSchemeBomItem { SortOrder = 1, Name = "框架断路器", Model = "CW1-2000M/3P 2000A 抽屉式", Brand = "常熟开关", Unit = "台", Quantity = 1.0, IsWlDoubled = false, CatalogPrice = 18500m, QuotePrice = 16800m },
                    new CloudSchemeBomItem { SortOrder = 2, Name = "电流互感器", Model = "BH-0.66 2000/5A 0.5级", Brand = "江苏安科瑞", Unit = "台", Quantity = 3.0, IsWlDoubled = true, CatalogPrice = 180m, QuotePrice = 150m },
                    new CloudSchemeBomItem { SortOrder = 3, Name = "多功能网络电力仪表", Model = "APM800 三相电能监测", Brand = "安科瑞", Unit = "台", Quantity = 1.0, IsWlDoubled = false, CatalogPrice = 1650m, QuotePrice = 1450m },
                    new CloudSchemeBomItem { SortOrder = 4, Name = "浪涌保护器", Model = "AM40-385/4P (Iimp 12.5kA)", Brand = "上海雷盾", Unit = "套", Quantity = 1.0, IsWlDoubled = false, CatalogPrice = 580m, QuotePrice = 480m }
                }
            });

            // 3. 双电源互投控制二次回路 (二次方案)
            list.Add(new CloudSchemeDetail
            {
                Id = "sch_ind_003",
                Scope = CloudSchemeScope.Industry,
                SchemeType = CloudSchemeType.Secondary,
                SchemeName = "双电源自投自复控制二次原理方案",
                CabinetModel = "通用控制箱/柜",
                Dimensions = "标准门板开孔",
                ThumbnailUrl = "https://images.unsplash.com/photo-1581092580497-e0d23cbdf1dc?w=500&q=80",
                DrawingUrls = new List<string>
                {
                    "https://images.unsplash.com/photo-1581092580497-e0d23cbdf1dc?w=1200&q=80"
                },
                Description = "典型双路市电自投自复二次控制原理回路，带常用/备用状态指示灯、手动/自动旋钮与故障蜂鸣报警。",
                CategoryId = "cat_secondary_control",
                CategoryName = "二次控制方案",
                ViewCount = 6720,
                FavCount = 1820,
                IsFavorite = true,
                CreatorName = "二次设计室",
                UpdateTime = "2026-09-06 14:20",
                Tag = "精",
                BomItems = new List<CloudSchemeBomItem>
                {
                    new CloudSchemeBomItem { SortOrder = 1, Name = "双电源自动转换控制器", Model = "HAT520N 智能双电源", Brand = "众智科技", Unit = "台", Quantity = 1.0, IsWlDoubled = false, CatalogPrice = 1200m, QuotePrice = 980m },
                    new CloudSchemeBomItem { SortOrder = 2, Name = "中间继电器", Model = "MY4N-J DC24V (含底座)", Brand = "欧姆龙", Unit = "只", Quantity = 4.0, IsWlDoubled = true, CatalogPrice = 35m, QuotePrice = 28m },
                    new CloudSchemeBomItem { SortOrder = 3, Name = "LED信号指示灯", Model = "AD16-22D/S (红绿黄)", Brand = "天正电气", Unit = "只", Quantity = 6.0, IsWlDoubled = true, CatalogPrice = 8.5m, QuotePrice = 6.5m },
                    new CloudSchemeBomItem { SortOrder = 4, Name = "万能转换开关", Model = "LW26-20 手动/自动", Brand = "正泰", Unit = "只", Quantity = 1.0, IsWlDoubled = false, CatalogPrice = 45m, QuotePrice = 36m }
                }
            });

            // 4. 企业示范配电箱方案 (企业自建示范方案)
            list.Add(new CloudSchemeDetail
            {
                Id = "sch_ent_001",
                Scope = CloudSchemeScope.Enterprise,
                SchemeType = CloudSchemeType.Primary,
                SchemeName = "企业标杆楼层动力照明配电箱(AL-1)",
                CabinetModel = "PZ30",
                Dimensions = "600*800*180",
                ThumbnailUrl = "https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=500&q=80",
                DrawingUrls = new List<string>
                {
                    "https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=1200&q=80"
                },
                Description = "办公楼层照明动力一体化配电箱，进线配置小型断路器总开关及浪涌，分路配微型断路器与漏保。",
                CategoryId = "cat_lighting_box",
                CategoryName = "配电箱方案",
                ViewCount = 350,
                FavCount = 68,
                IsFavorite = false,
                CreatorName = "成套技术部(张工)",
                UpdateTime = "2026-09-08 09:15",
                Tag = "企业",
                BomItems = new List<CloudSchemeBomItem>
                {
                    new CloudSchemeBomItem { SortOrder = 1, Name = "微型断路器总开", Model = "iC65N 3P 63A C特性", Brand = "施耐德", Unit = "台", Quantity = 1.0, IsWlDoubled = false, CatalogPrice = 280m, QuotePrice = 230m },
                    new CloudSchemeBomItem { SortOrder = 2, Name = "分路微断", Model = "iC65N 1P 16A C特性", Brand = "施耐德", Unit = "台", Quantity = 8.0, IsWlDoubled = true, CatalogPrice = 45m, QuotePrice = 38m },
                    new CloudSchemeBomItem { SortOrder = 3, Name = "分路漏电断路器", Model = "iC65N 2P 25A 30mA", Brand = "施耐德", Unit = "台", Quantity = 4.0, IsWlDoubled = true, CatalogPrice = 160m, QuotePrice = 135m }
                }
            });

            return list;
        }
    }
}
