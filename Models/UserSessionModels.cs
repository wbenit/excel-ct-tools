using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Models
{
    /// <summary>
    /// 本地持久化用户会话凭据数据结构模型
    /// 保存在 %AppData%\CTtools\user_session.json 中以支持免密静默登录
    /// </summary>
    public class UserSessionConfig
    {
        // 身份认证授权令牌 Token 密钥串
        [JsonPropertyName("token")]
        public string Token { get; set; } = string.Empty;

        // 用户唯一标识 UserId
        [JsonPropertyName("userId")]
        public string UserId { get; set; } = string.Empty;

        // 用户登录账号/手机号
        [JsonPropertyName("userName")]
        public string UserName { get; set; } = string.Empty;

        // 用户对外展示真实姓名或昵称
        [JsonPropertyName("displayName")]
        public string DisplayName { get; set; } = string.Empty;

        // 当前用户绑定的云端工作组租户 ID (GroupId)
        [JsonPropertyName("groupId")]
        public int GroupId { get; set; }

        // 当前用户绑定的云端工作组租户名称
        [JsonPropertyName("groupName")]
        public string GroupName { get; set; } = string.Empty;

        // 是否勾选记住登录状态
        [JsonPropertyName("rememberMe")]
        public bool RememberMe { get; set; }

        // 会话过期绝对时间戳 (UTC)
        [JsonPropertyName("expireTime")]
        public DateTime ExpireTime { get; set; }
    }

    /// <summary>
    /// 云端返回的用户工作组简要信息传输对象 (对应后端 CompanyGroupDto)
    /// 用于多租户分库路由与切库识别
    /// </summary>
    public class CompanyGroupDto
    {
        // 租户物理分库与工作组核心唯一标识 ID (GroupId)
        [JsonPropertyName("id")]
        public int Id { get; set; }

        // 公司企业或工作组全称
        [JsonPropertyName("companyName")]
        public string CompanyName { get; set; } = string.Empty;

        // 统一对外暴露的工作组名称属性 (自动兼容 CompanyName 与 GroupName)
        [JsonPropertyName("groupName")]
        public string GroupName
        {
            get => !string.IsNullOrWhiteSpace(CompanyName) ? CompanyName : _groupName;
            set => _groupName = value ?? string.Empty;
        }
        private string _groupName = string.Empty;

        // 工作组企业 Logo 静态资源路径
        [JsonPropertyName("logoUrl")]
        public string LogoUrl { get; set; } = string.Empty;

        // 标记当前是否为激活状态的工作组
        [JsonPropertyName("isActive")]
        public bool IsActive { get; set; }

        // 当前用户在当前工作组的角色标识 ID
        [JsonPropertyName("roleId")]
        public int? RoleId { get; set; }

        // 当前用户在当前工作组的角色名称
        [JsonPropertyName("roleName")]
        public string RoleName { get; set; } = string.Empty;

        // 角色在工作组中的权限排名索引 (0 为最高管理员权限)
        [JsonPropertyName("roleIndex")]
        public int RoleIndex { get; set; } = -1;
    }
}
