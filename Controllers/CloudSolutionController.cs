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
    }
}
