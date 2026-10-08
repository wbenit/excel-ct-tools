using System;
using System.Text.Json;
using ExcelAddInDemo.Models;
using ExcelAddInDemo.Services;
using ExcelAddInDemo;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 总价一键调整控制器，负责处理来自前端 WebView2 的 IPC 请求与 JSON 数据序列化
    /// </summary>
    public class TotalPriceAdjustController
    {
        // 统一 JSON 序列化选项 (小驼峰命名，忽略大小写)
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        /// <summary>
        /// 获取总价一键调整初始化数据 (包含工程总价、成本、分类表列表及动态 Q 列标签)
        /// </summary>
        /// <returns>序列化后的 JSON 字符串</returns>
        public string GetInitData()
        {
            try
            {
                // 调用服务层扫描提取初始化数据
                var data = ExcelServices.GetTotalPriceAdjustInitData();
                // 封装响应对象
                var response = new
                {
                    action = "onInitData",
                    data = data
                };
                // 序列化为 JSON 字符串
                return JsonSerializer.Serialize(response, JsonOptions);
            }
            catch (Exception ex)
            {
                // 记录控制器异常
                LogHelper.WriteLog($"[TotalPriceAdjustController] GetInitData 异常: {ex.Message}");
                var errorResponse = new
                {
                    action = "onInitData",
                    data = new TotalPriceAdjustInitData
                    {
                        Success = false,
                        Message = $"获取初始化数据失败: {ex.Message}"
                    }
                };
                return JsonSerializer.Serialize(errorResponse, JsonOptions);
            }
        }

        /// <summary>
        /// 执行总价一键调整
        /// </summary>
        /// <param name="requestJson">前端提交的请求 JSON 报文</param>
        /// <returns>序列化后的执行结果 JSON 字符串</returns>
        public string ExecuteAdjust(string requestJson)
        {
            try
            {
                // 反序列化请求实体
                var request = JsonSerializer.Deserialize<TotalPriceAdjustRequest>(requestJson, JsonOptions);
                if (request == null)
                {
                    var invalidResponse = new
                    {
                        action = "onAdjustResult",
                        data = new TotalPriceAdjustResult
                        {
                            Success = false,
                            Message = "未能正确解析调价请求参数"
                        }
                    };
                    return JsonSerializer.Serialize(invalidResponse, JsonOptions);
                }

                // 调用服务层执行总价调整
                var result = ExcelServices.ExecuteTotalPriceAdjust(request);
                var response = new
                {
                    action = "onAdjustResult",
                    data = result
                };
                return JsonSerializer.Serialize(response, JsonOptions);
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[TotalPriceAdjustController] ExecuteAdjust 异常: {ex.Message}");
                var errorResponse = new
                {
                    action = "onAdjustResult",
                    data = new TotalPriceAdjustResult
                    {
                        Success = false,
                        Message = $"执行调价失败: {ex.Message}"
                    }
                };
                return JsonSerializer.Serialize(errorResponse, JsonOptions);
            }
        }
    }
}
