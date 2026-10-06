using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 登录请求数据传输对象
    /// </summary>
    public class LoginRequest
    {
        // 用户名或登录账号/手机号
        public string Username { get; set; } = string.Empty;

        // 用户登录密码字符串
        public string Password { get; set; } = string.Empty;

        // 是否记住登录状态标识
        public bool RememberMe { get; set; } = true;
    }

    /// <summary>
    /// 工作组选择请求数据传输对象
    /// </summary>
    public class SelectGroupRequest
    {
        // 身份认证授权令牌 Token 串
        public string Token { get; set; } = string.Empty;

        // 用户唯一标识 Id 编号
        public string UserId { get; set; } = string.Empty;

        // 用户登录账号名称
        public string UserName { get; set; } = string.Empty;

        // 显示给用户的真实姓名或昵称
        public string DisplayName { get; set; } = string.Empty;

        // 选定的工作组 Id
        public int GroupId { get; set; }

        // 选定的工作组名称
        public string GroupName { get; set; } = string.Empty;

        // 是否记住登录状态标识
        public bool RememberMe { get; set; } = true;
    }

    /// <summary>
    /// 登录响应结果统一数据包格式
    /// </summary>
    public class LoginResponse
    {
        // 操作是否成功标记
        public bool Success { get; set; }

        // 返回给前端的提示消息文本
        public string Message { get; set; } = string.Empty;

        // 身份认证授权令牌 Token 串
        public string Token { get; set; } = string.Empty;

        // 用户唯一标识
        public string UserId { get; set; } = string.Empty;

        // 当前登录用户的显示名称
        public string DisplayName { get; set; } = string.Empty;

        // 是否需要用户手动选择工作组 (当存在多个工作组时为 true)
        public bool NeedSelectGroup { get; set; }

        // 当前选定的工作组 ID
        public int SelectedGroupId { get; set; }

        // 当前选定的工作组名称
        public string SelectedGroupName { get; set; } = string.Empty;

        // 当前用户加入的所有工作组列表
        public List<CompanyGroupDto> Groups { get; set; } = new List<CompanyGroupDto>();
    }

    /// <summary>
    /// 用户身份认证及配置授权 WebAPI 控制器实现
    /// </summary>
    public class AuthController
    {
        // 缓存最近一次登录成功后的工作组列表，方便根据 Id 查找 Name
        private List<CompanyGroupDto> _cachedGroups = new List<CompanyGroupDto>();

        /// <summary>
        /// 执行真实云端 DrawCode 用户登录身份验证
        /// </summary>
        /// <param name="request">包含账号密码的登录参数数据结构</param>
        /// <returns>返回带有状态、Token及工作组列表的登录结果</returns>
        public async Task<LoginResponse> LoginAsync(LoginRequest request)
        {
            // 基础校验：判断输入的用户名和密码是否为空
            if (string.IsNullOrWhiteSpace(request.Username) || string.IsNullOrWhiteSpace(request.Password))
            {
                // 返回用户名或密码为空的失败响应数据结构
                return new LoginResponse
                {
                    Success = false,
                    Message = "手机号与密码不能为空！"
                };
            }

            // 调用 DrawCodeApiClient 发起真实登录认证
            var (success, msg, token, userId, displayName, groups) = await DrawCodeApiClient.LoginAsync(
                request.Username.Trim(),
                request.Password
            );

            // 登录失败时直接返回失败信息
            if (!success)
            {
                return new LoginResponse
                {
                    Success = false,
                    Message = msg
                };
            }

            // 缓存工作组列表
            _cachedGroups = groups ?? new List<CompanyGroupDto>();

            // 场景 1：用户无任何工作组
            if (_cachedGroups.Count == 0)
            {
                return new LoginResponse
                {
                    Success = false,
                    Message = "登录成功，但当前账号尚未加入任何工作组，请联系企业管理员分配！"
                };
            }

            // 场景 2：用户只有一个工作组，自动完成绑定
            if (_cachedGroups.Count == 1)
            {
                var onlyGroup = _cachedGroups[0];
                // 绑定会话并持久化
                DrawCodeApiClient.BindActiveSession(
                    token,
                    userId,
                    request.Username.Trim(),
                    displayName,
                    onlyGroup.Id,
                    onlyGroup.GroupName,
                    request.RememberMe
                );

                return new LoginResponse
                {
                    Success = true,
                    Message = $"登录成功！已自动接入工作组：{onlyGroup.GroupName}",
                    Token = token,
                    UserId = userId,
                    DisplayName = displayName,
                    NeedSelectGroup = false,
                    SelectedGroupId = onlyGroup.Id,
                    SelectedGroupName = onlyGroup.GroupName,
                    Groups = _cachedGroups
                };
            }

            // 场景 3：用户拥有多个工作组，让用户选择（遵循用户指令：让用户选择）
            return new LoginResponse
            {
                Success = true,
                Message = "身份验证成功，请选择要进入的工作组",
                Token = token,
                UserId = userId,
                DisplayName = displayName,
                NeedSelectGroup = true,
                Groups = _cachedGroups
            };
        }

        /// <summary>
        /// 用户在前端选中工作组后的确认绑定
        /// </summary>
        /// <param name="request">工作组选择参数</param>
        /// <returns>操作结果</returns>
        public Task<LoginResponse> SelectGroupAsync(SelectGroupRequest request)
        {
            // 校验选择的工作组 Id 有效性
            if (request.GroupId <= 0)
            {
                return Task.FromResult(new LoginResponse
                {
                    Success = false,
                    Message = "请选择有效的工作组！"
                });
            }

            // 提取工作组名称（优先从入参或缓存中查找）
            string groupName = request.GroupName;
            if (string.IsNullOrWhiteSpace(groupName))
            {
                var found = _cachedGroups.FirstOrDefault(g => g.Id == request.GroupId);
                groupName = found?.GroupName ?? $"工作组 #{request.GroupId}";
            }

            // 调用客户端绑定会话并根据 RememberMe 加密持久化落盘
            DrawCodeApiClient.BindActiveSession(
                request.Token,
                request.UserId,
                request.UserName,
                request.DisplayName,
                request.GroupId,
                groupName,
                request.RememberMe
            );

            // 返回绑定成功响应
            return Task.FromResult(new LoginResponse
            {
                Success = true,
                Message = $"已成功接入工作组：{groupName}",
                Token = request.Token,
                UserId = request.UserId,
                DisplayName = request.DisplayName,
                NeedSelectGroup = false,
                SelectedGroupId = request.GroupId,
                SelectedGroupName = groupName,
                Groups = _cachedGroups
            });
        }

        /// <summary>
        /// 获取当前已加载的会话状态信息
        /// </summary>
        public object GetCurrentSessionInfo()
        {
            return new
            {
                IsLogged = !string.IsNullOrWhiteSpace(ExcelServices.CurrentToken),
                UserId = ExcelServices.CurrentUserId,
                DisplayName = ExcelServices.CurrentUserDisplayName,
                GroupId = ExcelServices.CurrentGroupId,
                GroupName = ExcelServices.CurrentGroupName
            };
        }

        /// <summary>
        /// 执行注销登出
        /// </summary>
        public void Logout()
        {
            // 清理本地持久化
            DrawCodeApiClient.ClearSession();
            // 重置运行内存
            ExcelServices.CurrentToken = string.Empty;
            ExcelServices.CurrentUserId = string.Empty;
            ExcelServices.CurrentUserDisplayName = "未登录";
            ExcelServices.CurrentGroupId = 0;
            ExcelServices.CurrentGroupName = string.Empty;
        }
    }
}

