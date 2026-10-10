using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// DrawCode 云端 WebAPI 官方通信服务客户端
    /// 统一管理用户登录、工作组选择、Windows DPAPI 凭证加密、会话持久化与租户切库
    /// </summary>
    public static class DrawCodeApiClient
    {
        // 全局单例 HttpClient 实例，复用底层 TCP 连接池
        private static readonly HttpClient _httpClient;

        // 本地用户会话文件持久化绝对路径 (%AppData%\CTtools\user_session.dat)
        private static readonly string SessionFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "CTtools",
            "user_session.dat" // 使用 .dat 后缀标明为 DPAPI 加密二进制数据
        );

        // 前端与后端统一绑定的 AES 密码对称加密密钥 (16 位 UTF8 密钥) --硬编码: 平台固定密钥--
        private const string AesKeyString = "2XNN4K8LC0ELVWN4";

        // JSON 序列化通用参数配置
        private static readonly JsonSerializerOptions JsonOpts = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 静态构造函数：初始化安全网络协议、连接池与默认 HTTP 请求头
        /// </summary>
        static DrawCodeApiClient()
        {
            try
            {
                // 启用 TLS 1.2 与 TLS 1.3 现代安全传输协议
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | (SecurityProtocolType)12288;
            }
            catch
            {
                try { ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12; } catch { }
            }

            // 初始化 HttpClient 处理器并开启响应流自动解压缩
            var handler = new HttpClientHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            // 创建长期驻留复用的 HttpClient 实例
            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(30)
            };

            // 预设标准请求头，模拟常规浏览器以防云端 WAF 网页防火墙拦截请求
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 ExcelAddInDemo/1.0");
            // 设置预期的媒体接收格式为 JSON 格式
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        /// <summary>
        /// 获取配置中 DrawCode 远程 WebAPI 基准访问地址（规范以 /api 结尾）
        /// </summary>
        public static string GetBaseUrl()
        {
            // 从全局配置管理器提取基准地址
            string url = ConfigManager.Instance.Current?.Api?.DrawCodeBaseUrl ?? string.Empty;
            // 若配置为空则回退到官方线上域名
            if (string.IsNullOrWhiteSpace(url))
            {
                url = "https://code.xingren.online/api"; // --硬编码: 官方线上默认域名--
            }
            url = url.Trim().TrimEnd('/');
            // 确保以 /api 结尾，兼容本地 5217 与线上域名
            if (!url.EndsWith("/api", StringComparison.OrdinalIgnoreCase))
            {
                url = $"{url}/api";
            }
            return url;
        }

        /// <summary>
        /// 获取动态配置的单次请求超时时限
        /// </summary>
        private static TimeSpan GetTimeout()
        {
            int sec = ConfigManager.Instance.Current?.Api?.TimeoutSeconds ?? 15;
            return TimeSpan.FromSeconds(sec > 0 ? sec : 15);
        }

        /// <summary>
        /// 检查响应头中是否存在 x-new-token 自动续期标记，若存在则无缝更新本地与内存凭证
        /// </summary>
        public static void CheckAndRefreshToken(HttpResponseMessage? response)
        {
            if (response == null) return;
            // 检测响应头中的新令牌标识
            if (response.Headers.TryGetValues("x-new-token", out var tokens))
            {
                string? newToken = tokens.FirstOrDefault();
                if (!string.IsNullOrWhiteSpace(newToken))
                {
                    // 更新全局静态内存令牌
                    ExcelServices.CurrentToken = newToken;
                    // 同步更新本地加密持久化凭证
                    var session = LoadSession();
                    if (session != null)
                    {
                        session.Token = newToken;
                        SaveSession(session);
                    }
                }
            }
        }

        /// <summary>
        /// 对明文密码执行 AES-128-ECB + PKCS7 对称加密（与 Web 前端加密算法 100% 对齐）
        /// </summary>
        /// <param name="plainPassword">用户输入的明文密码</param>
        /// <returns>Base64 编码后的密文字符串</returns>
        public static string EncryptPassword(string plainPassword)
        {
            if (string.IsNullOrEmpty(plainPassword)) return string.Empty;

            // 将 16 位秘钥转换为 UTF8 字节数组
            byte[] keyBytes = Encoding.UTF8.GetBytes(AesKeyString);
            // 将明文字符串转换为字节数组
            byte[] plainBytes = Encoding.UTF8.GetBytes(plainPassword);

            // 实例化系统标准 AES 加密算法引擎
            using var aes = Aes.Create();
            aes.Key = keyBytes;
            // 采用与前端 crypto-js 一致的 ECB 模式
            aes.Mode = CipherMode.ECB;
            // 采用 PKCS7 填充规范
            aes.Padding = PaddingMode.PKCS7;

            // 创建加密转换器
            using var encryptor = aes.CreateEncryptor();
            // 执行单次整块内存加密转换
            byte[] cipherBytes = encryptor.TransformFinalBlock(plainBytes, 0, plainBytes.Length);

            // 转换为 Base64 字符串返回
            return Convert.ToBase64String(cipherBytes);
        }

        /// <summary>
        /// 执行真实云端账号登录，成功后获取用户资料及所属的所有工作组列表供用户选择
        /// </summary>
        /// <param name="phone">用户手机号/账号</param>
        /// <param name="plainPassword">用户输入的明文密码</param>
        /// <returns>包含成功状态、用户基础信息及工作组列表的元组</returns>
        public static async Task<(bool Success, string Message, string Token, string UserId, string DisplayName, List<CompanyGroupDto> Groups)> LoginAsync(string phone, string plainPassword)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(phone) || string.IsNullOrWhiteSpace(plainPassword))
                {
                    return (false, "手机号与密码不能为空！", "", "", "", new List<CompanyGroupDto>());
                }

                // 1. 对明文密码执行标准 AES 加密
                string encryptedPwd = EncryptPassword(plainPassword);

                // 2. 拼装登录请求 URL
                string baseUrl = GetBaseUrl();
                string loginUrl = $"{baseUrl}/User/Login?phone={Uri.EscapeDataString(phone)}&password={Uri.EscapeDataString(encryptedPwd)}";

                using var cts = new CancellationTokenSource(GetTimeout());
                // 发起异步 HTTP GET 请求
                using var response = await _httpClient.GetAsync(loginUrl, cts.Token);
                string responseBody = await response.Content.ReadAsStringAsync();

                // 检查并刷新 Token
                CheckAndRefreshToken(response);

                if (!response.IsSuccessStatusCode)
                {
                    // 解析后端异常友好提示文本
                    string errorMsg = TryExtractErrorMessage(responseBody) ?? $"服务器响应异常 ({(int)response.StatusCode})";
                    return (false, errorMsg, "", "", "", new List<CompanyGroupDto>());
                }

                // 3. 解析登录响应数据包
                using var doc = JsonDocument.Parse(responseBody);
                var root = doc.RootElement;
                var dataNode = root;

                // 兼容外层带有 code / data 包装的数据结构
                if (root.TryGetProperty("data", out var wrappedData) && wrappedData.ValueKind == JsonValueKind.Object)
                {
                    dataNode = wrappedData;
                }

                // 提取 JWT Token 凭据
                string token = TryGetStringProperty(dataNode, "token", "Token");
                if (string.IsNullOrWhiteSpace(token))
                {
                    string err = TryExtractErrorMessage(responseBody) ?? "登录未能获取到有效的身份凭证 Token！";
                    return (false, err, "", "", "", new List<CompanyGroupDto>());
                }

                // 提取用户信息
                string userId = string.Empty;
                string displayName = phone;
                if (dataNode.TryGetProperty("userInfo", out var userNode) || dataNode.TryGetProperty("UserInfo", out userNode))
                {
                    userId = TryGetStringProperty(userNode, "userId", "UserId", "id", "Id");
                    displayName = TryGetStringProperty(userNode, "realName", "RealName", "name", "Name");
                    if (string.IsNullOrWhiteSpace(displayName)) displayName = phone;
                }

                // 4. 连带拉取当前用户加入的所有工作组列表
                var (groupSuccess, groups) = await GetUserGroupsAsync(token);

                return (true, "验证成功！", token, userId, displayName, groups);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[DrawCodeApiClient] LoginAsync 异常: {ex.Message}");
                return (false, $"网络通信异常: {ex.Message}", "", "", "", new List<CompanyGroupDto>());
            }
        }

        /// <summary>
        /// 携带 Token 凭证拉取当前用户加入的所有工作组列表 (对应后端 GET /api/Set/GetUserGroups)
        /// </summary>
        /// <param name="token">有效的 JWT 授权令牌</param>
        /// <returns>工作组列表集合</returns>
        public static async Task<(bool Success, List<CompanyGroupDto> Groups)> GetUserGroupsAsync(string token)
        {
            var resultList = new List<CompanyGroupDto>();
            try
            {
                if (string.IsNullOrWhiteSpace(token)) return (false, resultList);

                // 组装请求地址
                string baseUrl = GetBaseUrl();
                string groupsUrl = $"{baseUrl}/Set/GetUserGroups";

                // 创建包含 Authorization 头的独立请求包
                using var request = new HttpRequestMessage(HttpMethod.Get, groupsUrl);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                using var cts = new CancellationTokenSource(GetTimeout());
                // 执行请求
                using var response = await _httpClient.SendAsync(request, cts.Token);
                string responseBody = await response.Content.ReadAsStringAsync();

                // 检查并刷新 Token
                CheckAndRefreshToken(response);

                if (!response.IsSuccessStatusCode)
                {
                    return (false, resultList);
                }

                // 解析工作组 JSON
                using var doc = JsonDocument.Parse(responseBody);
                var root = doc.RootElement;
                var listNode = root;

                // 兼容外层 data 包装
                if (root.TryGetProperty("data", out var dataNode) && dataNode.ValueKind == JsonValueKind.Array)
                {
                    listNode = dataNode;
                }

                if (listNode.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in listNode.EnumerateArray())
                    {
                        var dto = JsonSerializer.Deserialize<CompanyGroupDto>(item.GetRawText(), JsonOpts);
                        if (dto != null && dto.Id > 0)
                        {
                            resultList.Add(dto);
                        }
                    }
                }

                return (true, resultList);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[DrawCodeApiClient] GetUserGroupsAsync 异常: {ex.Message}");
                return (false, resultList);
            }
        }

        /// <summary>
        /// 绑定用户选定的工作组，并更新全局状态及可选持久化
        /// </summary>
        public static void BindActiveSession(string token, string userId, string userName, string displayName, int groupId, string groupName, bool rememberMe)
        {
            // 更新 ExcelServices 静态运行内存
            ExcelServices.CurrentToken = token;
            ExcelServices.CurrentUserId = userId;
            ExcelServices.CurrentUserDisplayName = displayName;
            ExcelServices.CurrentGroupId = groupId;
            ExcelServices.CurrentGroupName = groupName;

            // 若勾选记住我，通过 Windows DPAPI 加密落盘
            if (rememberMe)
            {
                var session = new UserSessionConfig
                {
                    Token = token,
                    UserId = userId,
                    UserName = userName,
                    DisplayName = displayName,
                    GroupId = groupId,
                    GroupName = groupName,
                    RememberMe = true,
                    ExpireTime = DateTime.UtcNow.AddDays(15) // 会话有效期设为 15 天
                };
                SaveSession(session);
            }
            else
            {
                // 未勾选则清除以往持久化凭证
                ClearSession();
            }
        }

        /// <summary>
        /// 使用 Windows DPAPI (ProtectedData) 安全加密保存会话至本地磁盘
        /// </summary>
        public static void SaveSession(UserSessionConfig session)
        {
            try
            {
                if (session == null) return;
                string dir = Path.GetDirectoryName(SessionFilePath)!;
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                string json = JsonSerializer.Serialize(session, JsonOpts);
                byte[] plainBytes = Encoding.UTF8.GetBytes(json);
                // 使用 Windows DPAPI (CurrentUser 范围) 加密用户会话凭据，杜绝明文 Token 泄漏风险
                byte[] cipherBytes = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(SessionFilePath, cipherBytes);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[DrawCodeApiClient] SaveSession 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 从本地磁盘读取经 Windows DPAPI 加密的用户会话配置 (若过期或无效返回 null)
        /// </summary>
        public static UserSessionConfig? LoadSession()
        {
            try
            {
                if (!File.Exists(SessionFilePath)) return null;

                byte[] cipherBytes = File.ReadAllBytes(SessionFilePath);
                if (cipherBytes == null || cipherBytes.Length == 0) return null;

                string json;
                try
                {
                    // 优先使用 Windows DPAPI 解密
                    byte[] plainBytes = ProtectedData.Unprotect(cipherBytes, null, DataProtectionScope.CurrentUser);
                    json = Encoding.UTF8.GetString(plainBytes);
                }
                catch
                {
                    // 兼容旧格式未加密文本 (若存在)
                    json = Encoding.UTF8.GetString(cipherBytes);
                }

                if (string.IsNullOrWhiteSpace(json)) return null;

                var session = JsonSerializer.Deserialize<UserSessionConfig>(json, JsonOpts);
                if (session == null) return null;

                // 若未勾选记住我或已超过过期时间，则视为失效
                if (!session.RememberMe || (session.ExpireTime > DateTime.MinValue && session.ExpireTime < DateTime.UtcNow))
                {
                    return null;
                }

                return session;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[DrawCodeApiClient] LoadSession 异常: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 清理本地持久化的会话凭证 (用于注销或登出)
        /// </summary>
        public static void ClearSession()
        {
            try
            {
                if (File.Exists(SessionFilePath))
                {
                    File.Delete(SessionFilePath);
                }
            }
            catch { }
        }

        /// <summary>
        /// 查询当前工作组下的云端项目列表 (对应后端 POST /api/Project/GetProjectList?groupId={groupId})
        /// </summary>
        /// <param name="keyword">项目名称模糊搜索关键词</param>
        /// <param name="pageNum">当前页码 (1-based)</param>
        /// <param name="pageSize">每页记录条数</param>
        /// <returns>项目列表集合及状态描述</returns>
        public static async Task<(bool Success, string Message, List<ProjectSummaryDto> Projects)> GetProjectsAsync(string keyword = "", int pageNum = 1, int pageSize = 50)
        {
            var list = new List<ProjectSummaryDto>();
            try
            {
                // 校验身份凭据与工作组绑定状态
                string token = ExcelServices.CurrentToken;
                int groupId = ExcelServices.CurrentGroupId;
                if (string.IsNullOrWhiteSpace(token) || groupId <= 0)
                {
                    return (false, "请先登录并绑定工作组后再查询项目列表！", list);
                }

                // 拼装请求 URL（必须携带 groupId 用于多租户分库路由）
                string baseUrl = GetBaseUrl();
                string url = $"{baseUrl}/Project/GetProjectList?groupId={groupId}";

                // 构造分页与过滤查询参数
                var reqBody = new
                {
                    PageNum = pageNum,
                    PageSize = pageSize,
                    ProjectName = string.IsNullOrWhiteSpace(keyword) ? null : keyword.Trim()
                };

                // 序列化请求体
                string jsonBody = JsonSerializer.Serialize(reqBody, JsonOpts);
                using var request = new HttpRequestMessage(HttpMethod.Post, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

                using var cts = new CancellationTokenSource(GetTimeout());
                // 发起 HTTP 请求
                using var response = await _httpClient.SendAsync(request, cts.Token);
                string responseBody = await response.Content.ReadAsStringAsync();

                // 检查并刷新 Token
                CheckAndRefreshToken(response);

                if (!response.IsSuccessStatusCode)
                {
                    string err = TryExtractErrorMessage(responseBody) ?? $"获取项目失败 ({(int)response.StatusCode})";
                    return (false, err, list);
                }

                // 解析项目列表响应
                using var doc = JsonDocument.Parse(responseBody);
                var root = doc.RootElement;
                var arrayNode = root;

                // 💡 兼容外层带有 data 包装的数据结构，并同时兼容分页对象（data.items 或 data.list）
                if (root.TryGetProperty("data", out var dataNode))
                {
                    if (dataNode.ValueKind == JsonValueKind.Array)
                    {
                        arrayNode = dataNode;
                    }
                    else if (dataNode.ValueKind == JsonValueKind.Object)
                    {
                        // 💡 若返回的是包含 items 或 list 的标准分页对象，提取其子数组节点
                        if ((dataNode.TryGetProperty("items", out var itemsNode) || dataNode.TryGetProperty("list", out itemsNode) ||
                             dataNode.TryGetProperty("Items", out itemsNode) || dataNode.TryGetProperty("List", out itemsNode)) &&
                            itemsNode.ValueKind == JsonValueKind.Array)
                        {
                            arrayNode = itemsNode;
                        }
                    }
                }
                else if (root.ValueKind == JsonValueKind.Object)
                {
                    // 💡 兼容直接返回分页对象且无外层 data 包装的情形
                    if ((root.TryGetProperty("items", out var itemsNode) || root.TryGetProperty("list", out itemsNode) ||
                         root.TryGetProperty("Items", out itemsNode) || root.TryGetProperty("List", out itemsNode)) &&
                        itemsNode.ValueKind == JsonValueKind.Array)
                    {
                        arrayNode = itemsNode;
                    }
                }

                if (arrayNode.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in arrayNode.EnumerateArray())
                    {
                        try
                        {
                            // 优先通过标准反序列化解析项目概要
                            var proj = JsonSerializer.Deserialize<ProjectSummaryDto>(item.GetRawText(), JsonOpts);
                            if (proj != null && proj.Id > 0)
                            {
                                list.Add(proj);
                                continue;
                            }
                        }
                        catch { }

                        // 兜底方案：通过 JsonElement 逐字段弹性提取核心属性
                        try
                        {
                            int pId = 0;
                            if (item.TryGetProperty("id", out var idP) || item.TryGetProperty("Id", out idP))
                            {
                                if (idP.ValueKind == JsonValueKind.Number) pId = idP.GetInt32();
                                else if (idP.ValueKind == JsonValueKind.String && int.TryParse(idP.GetString(), out int parsedId)) pId = parsedId;
                            }
                            string pName = string.Empty;
                            if (item.TryGetProperty("projectName", out var nameP) || item.TryGetProperty("ProjectName", out nameP))
                            {
                                pName = nameP.GetString() ?? string.Empty;
                            }
                            string pType = string.Empty;
                            if (item.TryGetProperty("projectType", out var typeP) || item.TryGetProperty("ProjectType", out typeP))
                            {
                                pType = typeP.GetString() ?? string.Empty;
                            }

                            if (pId > 0)
                            {
                                list.Add(new ProjectSummaryDto
                                {
                                    Id = pId,
                                    ProjectName = pName,
                                    ProjectType = pType
                                });
                            }
                        }
                        catch { }
                    }
                }

                // 记录成功日志
                LogHelper.WriteLog($"[DrawCodeApiClient] GetProjectsAsync 成功获取 {list.Count} 条项目 (keyword: '{keyword}')");
                return (true, "获取成功", list);
            }
            catch (Exception ex)
            {
                // 记录详细异常日志
                LogHelper.WriteLog($"[DrawCodeApiClient] GetProjectsAsync 异常: {ex}");
                return (false, $"网络通信异常: {ex.Message}", list);
            }
        }

        /// <summary>
        /// 推送批量导入箱柜数据至云端系统图库 (对应后端 POST /api/Project/ImportExcelData?groupId={groupId})
        /// 支持用户确认后覆盖已有的相同 Order 序号箱柜
        /// </summary>
        /// <param name="importDto">包含箱柜列表、批次及是否覆盖标记的导入参数</param>
        /// <returns>导入成功结果统计</returns>
        public static async Task<CloudImportResultDto> ImportExcelDataAsync(CloudImportRequestDto importDto)
        {
            var result = new CloudImportResultDto();
            try
            {
                // 校验身份凭据与工作组绑定状态
                string token = ExcelServices.CurrentToken;
                int groupId = ExcelServices.CurrentGroupId;
                if (string.IsNullOrWhiteSpace(token) || groupId <= 0)
                {
                    result.Success = false;
                    result.Message = "未登录或未绑定工作组，无法推送箱柜！";
                    return result;
                }

                if (importDto == null || importDto.EBoxs == null || importDto.EBoxs.Count == 0)
                {
                    result.Success = false;
                    result.Message = "没有可推送的箱柜数据！";
                    return result;
                }

                // 拼装请求 URL（显式携带 groupId 供 CtmoGroupIdFilterAttribute 分库路由）
                string baseUrl = GetBaseUrl();
                string url = $"{baseUrl}/Project/ImportExcelData?groupId={groupId}";

                // 序列化请求体
                string jsonBody = JsonSerializer.Serialize(importDto, JsonOpts);
                using var request = new HttpRequestMessage(HttpMethod.Post, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(60)); // 导入数据耗时预留60秒
                // 发起 HTTP 请求
                using var response = await _httpClient.SendAsync(request, cts.Token);
                string responseBody = await response.Content.ReadAsStringAsync();

                // 检查并刷新 Token
                CheckAndRefreshToken(response);

                if (!response.IsSuccessStatusCode)
                {
                    string err = TryExtractErrorMessage(responseBody) ?? $"推送失败 ({(int)response.StatusCode})";
                    result.Success = false;
                    result.Message = err;
                    return result;
                }

                // 解析后端返回的 ExcelImportResultDto
                using var doc = JsonDocument.Parse(responseBody);
                var root = doc.RootElement;
                var dataNode = root;
                if (root.TryGetProperty("data", out var dNode) && dNode.ValueKind == JsonValueKind.Object)
                {
                    dataNode = dNode;
                }

                int eBoxCount = 0;
                if (dataNode.TryGetProperty("eBoxCount", out var bcProp) || dataNode.TryGetProperty("EBoxCount", out bcProp))
                {
                    eBoxCount = bcProp.GetInt32();
                }

                int eTwoCount = 0;
                if (dataNode.TryGetProperty("eTwoCount", out var tcProp) || dataNode.TryGetProperty("ETwoCount", out tcProp))
                {
                    eTwoCount = tcProp.GetInt32();
                }

                string msg = TryGetStringProperty(dataNode, "message", "Message");
                if (string.IsNullOrWhiteSpace(msg))
                {
                    msg = $"成功同步 {eBoxCount} 台箱柜至云端！";
                }

                result.Success = true;
                result.EBoxCount = eBoxCount > 0 ? eBoxCount : importDto.EBoxs.Count;
                result.ETwoCount = eTwoCount;
                result.Message = msg;
                return result;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[DrawCodeApiClient] ImportExcelDataAsync 异常: {ex.Message}");
                result.Success = false;
                result.Message = $"网络异常: {ex.Message}";
                return result;
            }
        }

        /// <summary>
        /// 实时从云端 WebAPI 拉取指定项目的系统图模板配置，并动态解析自定义字段标签名列表
        /// 保证非硬编码，每次依据云端项目配置实时拉取
        /// </summary>
        /// <param name="projectId">目标项目 ID</param>
        /// <param name="groupId">可选工作组 ID (若为 null 则自动使用当前工作组)</param>
        /// <returns>(是否成功, 提示信息, 自定义字段标签列表)</returns>
        public static async Task<(bool Success, string Message, List<string> Labels)> GetProjectCustomFieldLabelsAsync(int projectId, int? groupId = null)
        {
            var labels = new List<string>();
            try
            {
                // 1. 优先使用当前有效 Token，未登录则抛出提示
                string token = ExcelServices.CurrentToken;
                if (string.IsNullOrWhiteSpace(token))
                {
                    return (false, "用户未登录，无法拉取项目模板", labels);
                }

                // 2. 确定请求的目标工作组 ID
                int targetGroupId = groupId ?? ExcelServices.CurrentGroupId;

                // 3. 构造请求 URL: /Project/GetProjectTemplate?projectId={projectId}&groupId={targetGroupId}
                string baseUrl = GetBaseUrl();
                string url = $"{baseUrl}/Project/GetProjectTemplate?projectId={projectId}&groupId={targetGroupId}";

                // 4. 创建 HTTP GET 请求对象
                using var request = new HttpRequestMessage(HttpMethod.Get, url);
                // 附加 Bearer Token 认证头
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                // 附加工作组切库 Header
                request.Headers.Add("ctmo-group-id", targetGroupId.ToString());

                // 5. 设置 15 秒超时防护
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
                // 异步发送请求并获取响应
                using var response = await _httpClient.SendAsync(request, cts.Token);
                // 读取响应体内容文本
                string responseBody = await response.Content.ReadAsStringAsync();

                // 检查并自动刷新 Token
                CheckAndRefreshToken(response);

                // 若 HTTP 状态码非 200，记录日志并返回错误
                if (!response.IsSuccessStatusCode)
                {
                    string err = TryExtractErrorMessage(responseBody) ?? $"拉取模板失败 ({(int)response.StatusCode})";
                    return (false, err, labels);
                }

                if (string.IsNullOrWhiteSpace(responseBody))
                {
                    return (true, "项目未配置模板", labels);
                }

                // 6. 弹性解析 JSON：可能是 AjaxResult 包裹，也可能是直接的 JSON 串
                using var doc = JsonDocument.Parse(responseBody);
                var root = doc.RootElement;

                // 检查是否有 data 属性包裹
                string templateJsonStr = "";
                if (root.TryGetProperty("data", out var dataElem))
                {
                    // 若 data 为字符串形式的模板 JSON
                    if (dataElem.ValueKind == JsonValueKind.String)
                    {
                        templateJsonStr = dataElem.GetString() ?? "";
                    }
                    else if (dataElem.ValueKind == JsonValueKind.Object)
                    {
                        // 若 data 直接为对象形式
                        templateJsonStr = dataElem.GetRawText();
                    }
                }
                else
                {
                    // 若无 data 外壳，则 root 本身即为模板内容
                    templateJsonStr = responseBody;
                }

                if (string.IsNullOrWhiteSpace(templateJsonStr))
                {
                    return (true, "项目未配置自定义字段", labels);
                }

                // 7. 二次解析模板 JSON 提取 config.customContent
                using var tplDoc = JsonDocument.Parse(templateJsonStr);
                var tplRoot = tplDoc.RootElement;

                JsonElement customContentNode = default;
                bool foundCustomContent = false;

                // 方式 A: 尝试在 config.customContent 中查找
                if (tplRoot.TryGetProperty("config", out var configNode) && configNode.ValueKind == JsonValueKind.Object)
                {
                    if (configNode.TryGetProperty("customContent", out var ccNode) && ccNode.ValueKind == JsonValueKind.Array)
                    {
                        customContentNode = ccNode;
                        foundCustomContent = true;
                    }
                }

                // 方式 B: 若 config 下未找到，尝试直接在根节点查找 customContent
                if (!foundCustomContent && tplRoot.TryGetProperty("customContent", out var directCcNode) && directCcNode.ValueKind == JsonValueKind.Array)
                {
                    customContentNode = directCcNode;
                    foundCustomContent = true;
                }

                // 8. 遍历自定义字段数组提取 label
                if (foundCustomContent && customContentNode.ValueKind == JsonValueKind.Array)
                {
                    var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var item in customContentNode.EnumerateArray())
                    {
                        // 提取每个自定义字段的 label 属性 (如 "柜体尺寸", "防护等级")
                        if (item.TryGetProperty("label", out var lProp) && lProp.ValueKind == JsonValueKind.String)
                        {
                            string lVal = lProp.GetString()?.Trim() ?? "";
                            // 过滤空标签与重复字段
                            if (!string.IsNullOrEmpty(lVal) && set.Add(lVal))
                            {
                                labels.Add(lVal);
                            }
                        }
                    }
                }

                return (true, $"成功获取 {labels.Count} 个自定义字段", labels);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[DrawCodeApiClient] GetProjectCustomFieldLabelsAsync 异常: {ex.Message}");
                return (false, $"解析自定义字段异常: {ex.Message}", labels);
            }
        }

        /// <summary>
        /// 从云端获取指定项目的生产批次列表并计算最大批次
        /// 遵循规范：每 3 行至少 1 行中文注释
        /// </summary>
        /// <param name="projectId">云端项目 ID</param>
        /// <param name="groupId">可选工作组 ID (若为 null 则自动使用当前工作组)</param>
        /// <returns>当前最大生产批次（若尚无批次或拉取失败返回 0）</returns>
        public static async Task<int> GetProjectMaxProduceOrderAsync(int projectId, int? groupId = null)
        {
            try
            {
                if (projectId <= 0) return 0;
                string token = ExcelServices.CurrentToken;
                if (string.IsNullOrWhiteSpace(token)) return 0;

                int targetGroupId = groupId ?? ExcelServices.CurrentGroupId;
                string baseUrl = GetBaseUrl();
                // 构造请求 URL: /Project/GetProduceOrdersShow?projectId={projectId}&groupId={targetGroupId}
                string url = $"{baseUrl}/Project/GetProduceOrdersShow?projectId={projectId}&groupId={targetGroupId}";

                using var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                request.Headers.Add("ctmo-group-id", targetGroupId.ToString());

                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
                using var response = await _httpClient.SendAsync(request, cts.Token);
                string responseBody = await response.Content.ReadAsStringAsync();

                CheckAndRefreshToken(response);
                if (!response.IsSuccessStatusCode || string.IsNullOrWhiteSpace(responseBody)) return 0;

                using var doc = JsonDocument.Parse(responseBody);
                var root = doc.RootElement;
                var listNode = root;

                // 若被 data 属性包裹则下钻
                if (root.TryGetProperty("data", out var dNode) && dNode.ValueKind == JsonValueKind.Array)
                {
                    listNode = dNode;
                }

                int maxOrder = 0;
                if (listNode.ValueKind == JsonValueKind.Array)
                {
                    foreach (var elem in listNode.EnumerateArray())
                    {
                        // 兼容 produceOrder / ProduceOrder 字段名
                        if ((elem.TryGetProperty("produceOrder", out var poProp) || elem.TryGetProperty("ProduceOrder", out poProp))
                            && poProp.TryGetInt32(out int poVal))
                        {
                            if (poVal > maxOrder) maxOrder = poVal;
                        }
                    }
                }
                return maxOrder;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[DrawCodeApiClient] GetProjectMaxProduceOrderAsync 异常: {ex.Message}");
                return 0;
            }
        }

        #region 私有辅助工具方法

        // 尝试从多个备选属性名中提取非空字符串值
        private static string TryGetStringProperty(JsonElement element, params string[] propertyNames)
        {
            foreach (var name in propertyNames)
            {
                if (element.TryGetProperty(name, out var prop) && prop.ValueKind == JsonValueKind.String)
                {
                    string val = prop.GetString() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(val)) return val.Trim();
                }
            }
            return string.Empty;
        }

        // 尝试从后端返回的 JSON 串中提取错误描述文本
        private static string? TryExtractErrorMessage(string json)
        {
            try
            {
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                string msg = TryGetStringProperty(root, "message", "Message", "msg", "Msg", "error", "Error");
                return string.IsNullOrWhiteSpace(msg) ? null : msg;
            }
            catch
            {
                return null;
            }
        }

        #endregion
    }
}
