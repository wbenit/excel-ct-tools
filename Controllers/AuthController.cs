using System;
using System.Threading.Tasks;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 登录请求数据传输对象
    /// </summary>
    public class LoginRequest
    {
        // 用户名或登录账号名称
        public string Username { get; set; } = string.Empty;

        // 用户登录密码字符串
        public string Password { get; set; } = string.Empty;

        // 是否记住登录状态标识
        public bool RememberMe { get; set; }
    }

    /// <summary>
    /// 用户信息模型对象
    /// </summary>
    public class UserDto
    {
        // 用户唯一标识 Id 编号
        public string UserId { get; set; } = string.Empty;

        // 显示给用户的真实姓名或昵称
        public string DisplayName { get; set; } = string.Empty;

        // 用户所属组织或部门名称
        public string Department { get; set; } = string.Empty;
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

        // 当前登录用户的详细扩展信息
        public UserDto? User { get; set; }
    }

    /// <summary>
    /// 用户身份认证及配置授权 WebAPI 控制器实现
    /// </summary>
    public class AuthController
    {
        /// <summary>
        /// 执行用户登录身份验证请求接口
        /// </summary>
        /// <param name="request">包含账号密码的登录参数数据结构</param>
        /// <returns>返回带有状态及 Token 的登录结果</returns>
        public async Task<LoginResponse> LoginAsync(LoginRequest request)
        {
            // 异步模拟网络请求延迟与后台计算
            await Task.Delay(300);

            // 基础校验：判断输入的用户名和密码是否为空
            if (string.IsNullOrWhiteSpace(request.Username) || string.IsNullOrWhiteSpace(request.Password))
            {
                // 返回用户名或密码为空的失败响应数据结构
                return new LoginResponse
                {
                    Success = false,
                    Message = "用户名与密码不能为空！"
                };
            }

            // 示例校验规则：模拟校验（实际应用中连接后端 .NET WebAPI 服务）
            if (request.Password.Length < 4)
            {
                // 返回密码长度不足的失败响应包
                return new LoginResponse
                {
                    Success = false,
                    Message = "密码长度至少需要 4 位字符！"
                };
            }

            // 构造生成随机或符合标准的 JWT Token 凭据
            string token = "Bearer_" + Guid.NewGuid().ToString("N");

            // 返回校验成功的授权与用户信息结果包
            return new LoginResponse
            {
                Success = true,
                Message = "身份认证成功，欢迎登录！",
                Token = token,
                User = new UserDto
                {
                    UserId = "U1001",
                    DisplayName = request.Username,
                    Department = "鑫壬成套技术部"
                }
            };
        }
    }
}
