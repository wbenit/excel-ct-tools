using System;
using System.IO;
using System.Drawing;
using System.Collections.Generic;
using ExcelDna.Integration;
using ExcelAddInDemo.Controllers;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 业务服务分部类：企业设置数据与工作簿同步
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 将企业设置数据同步写回当前打开的活动工程工作簿（“项目信息”工作表）
        /// </summary>
        /// <param name="settings">企业设置数据模型</param>
        /// <returns>同步成功与否的结果元组 (成功标志, 详细反馈消息)</returns>
        public static (bool Success, string Message) SyncEnterpriseSettingsToActiveWorkbook(EnterpriseSettingsData settings)
        {
            // 校验入参数据对象有效性
            if (settings == null)
            {
                // 空数据拦截并返回提示
                return (false, "企业设置数据为空，无法同步。");
            }

            try
            {
                // 获取当前活动 Excel Application 实例
                dynamic app = ExcelDnaUtil.Application;
                // 提取当前打开的活动工作簿句柄
                dynamic activeWb = app?.ActiveWorkbook;
                if (activeWb == null)
                {
                    // 未打开任何工作簿时友好提示
                    return (false, "当前没有打开任何 Excel 工作簿。");
                }

                // 核心安全守门：校验当前工作簿是否为合法的成套工程工作簿 (必须包含【项目信息】主表)
                if (!Tool.IsProjectWorkbook(activeWb))
                {
                    // 非成套工程工作簿拦截并提示
                    return (false, "当前工作簿不是标准成套工程工作簿（未检测到【项目信息】主表）。");
                }

                // 获取【项目信息】工作表句柄 --硬编码: 项目信息工作表名--
                dynamic infoSheet = null;
                try { infoSheet = activeWb.Sheets["项目信息"]; } catch { }
                if (infoSheet == null)
                {
                    // 未找到对应工作表退出
                    return (false, "未找到【项目信息】工作表。");
                }

                // 1. 同步文本字段至【项目信息】各单元格 (规则 7 / 保证单点写入规范)
                // 顶栏单位名称 (Row 1) 与本企业信息单位名称 (Row 22)
                if (!string.IsNullOrWhiteSpace(settings.CompanyName))
                {
                    // 赋值 B1 单元格
                    infoSheet.Range["B1"].Value = settings.CompanyName;
                    // 赋值 B22 单元格
                    infoSheet.Range["B22"].Value = settings.CompanyName;
                }

                // 英文名称 (Cell B23) - 重点闭环：同步箭头所指单元格
                infoSheet.Range["B23"].Value = settings.EnglishName ?? string.Empty;

                // 本企业联系人 (Cell B24)
                infoSheet.Range["B24"].Value = settings.ContactPerson ?? string.Empty;

                // 本企业联系电话 (Cell B25)
                infoSheet.Range["B25"].Value = settings.ContactPhone ?? string.Empty;

                // 本企业销售地区 (Cell B26)
                infoSheet.Range["B26"].Value = settings.SalesRegion ?? string.Empty;

                // 报价人 (Cell B8)
                if (!string.IsNullOrWhiteSpace(settings.Quoter))
                {
                    // 同步报价人姓名
                    infoSheet.Range["B8"].Value = settings.Quoter;
                }

                // 2. 同步与替换【项目信息】主表表头 Logo 图片 - 重点闭环：同步箭头所指表头 Logo
                SyncLogoImageToSheet(infoSheet, settings.LogoBase64);

                // 3. 全量同步当前工作簿中所有【分类明细表】的表头 Logo 图片 (重点闭环：logo 同步到分类明细)
                int syncedCategoryCount = 0;
                try
                {
                    // 获取当前活动工作簿登记的所有分类明细工作表名称列表
                    var categoryNames = Tool.GetProjectCategorySheetNames(activeWb, forceRefresh: true);
                    var targetCatSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    if (categoryNames != null)
                    {
                        // 将登记的分类名称加入哈希集合
                        foreach (var name in categoryNames) targetCatSet.Add(name);
                    }

                    // 辅助清单非分类表排除集合
                    var helperSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    {
                        "采购清单", "领料清单", "人工清单", "成品交接单"
                    };

                    // 遍历工作簿中全部工作表，双轨守门确保无遗漏
                    foreach (dynamic ws in activeWb.Worksheets)
                    {
                        try
                        {
                            // 提取工作表名称
                            string wsName = Convert.ToString(ws.Name)?.Trim() ?? string.Empty;
                            if (string.IsNullOrWhiteSpace(wsName)) continue;

                            // 排除非分类表的系统保留工作表 (如封面、元件汇总表、屏柜汇总表等)
                            if (Tool.IsReservedOrReportSheet(wsName)) continue;
                            // 排除项目信息主表 (前面已单独同步)
                            if (string.Equals(wsName, "项目信息", StringComparison.OrdinalIgnoreCase)) continue;
                            // 排除辅助清单表
                            if (helperSheets.Contains(wsName)) continue;

                            // 判定是否为纳管的分类明细表 (在登记集合中或命中分类表结构判定)
                            bool isCategory = targetCatSet.Contains(wsName) || Tool.IsProjectCategorySheet(ws, activeWb);

                            // 若白名单未覆盖，进一步通过工作表左上角是否有图片/表头特征兜底探测
                            if (!isCategory)
                            {
                                try
                                {
                                    // 探测工作表图形
                                    foreach (dynamic s in ws.Shapes)
                                    {
                                        // 检查是否存在表头位置图形
                                        if (s.Top < 80 && s.Left < 220)
                                        {
                                            isCategory = true;
                                            break;
                                        }
                                    }
                                }
                                catch { }
                            }

                            // 若判定为分类明细表
                            if (isCategory)
                            {
                                // 执行分类表表头 Logo 替换写入
                                SyncLogoImageToSheet(ws, settings.LogoBase64);
                                // 累加同步成功计数
                                syncedCategoryCount++;
                            }
                        }
                        catch (Exception exCatItem)
                        {
                            // 记录单表同步异常日志
                            LogHelper.WriteLog($"[EnterpriseSettings] 同步分类工作表 [{ws?.Name}] Logo 异常: {exCatItem.Message}");
                        }
                    }
                }
                catch (Exception exCats)
                {
                    // 记录分类遍历异常日志
                    LogHelper.WriteLog($"[EnterpriseSettings] 扫描分类明细表同步 Logo 异常: {exCats.Message}");
                }

                // 组织友好的成功反馈信息
                string successMsg = syncedCategoryCount > 0 
                    ? $"已成功将企业设置与 Logo 同步至【项目信息】及 {syncedCategoryCount} 个分类明细表！" 
                    : "已成功将企业设置与 Logo 同步至当前打开的【项目信息】工作表！";

                // 同步完成返回成功结果
                return (true, successMsg);
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"[EnterpriseSettings] 同步至活动工作簿异常: {ex.Message}");
                // 容错返回错误描述
                return (false, $"同步过程中发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 同步更新或替换工作表左上角 Logo 图片
        /// </summary>
        /// <param name="sheet">目标工作表句柄</param>
        /// <param name="logoBase64">Logo 图片的 Base64 编码字符串</param>
        public static void SyncLogoImageToSheet(dynamic sheet, string logoBase64)
        {
            // 校验工作表有效性
            if (sheet == null) return;

            try
            {
                // 优先继承原旧 Logo 的物理坐标，实现 100% 原位无损替换
                float targetLeft = 4f;
                float targetTop = 4f;
                bool hasAnchorPos = false;

                // 先清理原有左上角 Logo 图片 (遍历查找落在 A1:C3 区域且 Top < 80, Left < 220 的 Shape)
                var shapesToDelete = new List<dynamic>();
                try
                {
                    // 遍历工作表所有图形对象
                    foreach (dynamic shape in sheet.Shapes)
                    {
                        try
                        {
                            // 检查位置在左上角表头 Logo 放置区域内 --硬编码: Logo 探测范围 Top < 80 且 Left < 220--
                            if (shape.Top < 80 && shape.Left < 220)
                            {
                                // 首次命中时抓取原图形实际坐标
                                if (!hasAnchorPos)
                                {
                                    // 继承原图水平坐标
                                    targetLeft = (float)shape.Left;
                                    // 继承原图垂直坐标
                                    targetTop = (float)shape.Top;
                                    hasAnchorPos = true;
                                }
                                // 收集命中的旧 Logo Shape
                                shapesToDelete.Add(shape);
                            }
                        }
                        catch { }
                    }

                    // 批量执行删除旧 Logo
                    foreach (var s in shapesToDelete)
                    {
                        // 删除旧图形
                        try { s.Delete(); } catch { }
                    }
                }
                catch (Exception exShape)
                {
                    // 记录清理旧 Logo 日志
                    LogHelper.WriteLog($"[EnterpriseSettings] 清理旧 Logo 图片异常: {exShape.Message}");
                }

                // 若未在表头探查到旧图形，以 A1 单元格左上角偏移 4pt 兜底
                if (!hasAnchorPos)
                {
                    try
                    {
                        // 提取 A1 单元格句柄
                        dynamic cellA1 = sheet.Range["A1"];
                        targetLeft = (float)cellA1.Left + 4f; // --硬编码: 兜底水平偏移 4pt--
                        targetTop = (float)cellA1.Top + 4f;   // --硬编码: 兜底垂直偏移 4pt--
                    }
                    catch { }
                }

                // 若 Base64 串为空，说明用户清空了 Logo，清理完后直接退出
                if (string.IsNullOrWhiteSpace(logoBase64))
                {
                    // 空 Logo 场景处理完成退出
                    return;
                }

                // 准备图片字节数组
                byte[] imageBytes;
                // 若传入的是现存物理文件路径则直接读取
                if (File.Exists(logoBase64))
                {
                    // 读取物理文件字节
                    imageBytes = File.ReadAllBytes(logoBase64);
                }
                else
                {
                    // 提取 Base64 纯数据部分 (兼容清洗 data:image/png;base64, 前缀)
                    string pureBase64 = logoBase64;
                    // 查找逗号分隔符索引
                    int commaIdx = pureBase64.IndexOf(',');
                    if (commaIdx >= 0)
                    {
                        // 截取纯数据内容
                        pureBase64 = pureBase64.Substring(commaIdx + 1);
                    }
                    // 将 Base64 字符串解码为图片字节数组
                    imageBytes = Convert.FromBase64String(pureBase64);
                }

                // 校验字节非空
                if (imageBytes == null || imageBytes.Length == 0) return;

                // 生成系统临时图片保存物理路径
                string tempFilePath = Path.Combine(Path.GetTempPath(), $"ct_logo_{Guid.NewGuid():N}.png");
                // 将图片字节写入本地临时文件
                File.WriteAllBytes(tempFilePath, imageBytes);

                try
                {
                    // 设置插入左边距与上边距
                    float left = targetLeft;
                    float top = targetTop;

                    // 读取图片真实尺寸并计算保真缩放比例
                    using (var memStream = new MemoryStream(imageBytes))
                    using (var gdiImg = System.Drawing.Image.FromStream(memStream))
                    {
                        // 设定 Logo 最大展示区域限制 (原模板标准尺寸为 105.75x39，设定上限 120x42 完美防溢出)
                        float maxW = 120f; // --硬编码: Logo 最大展示宽度 120pt--
                        float maxH = 42f;  // --硬编码: Logo 最大展示高度 42pt--

                        // 计算宽、高维度的缩放比例因子
                        float scaleW = maxW / gdiImg.Width;
                        float scaleH = maxH / gdiImg.Height;
                        // 取较小缩放比以保持等比完整展示
                        float scale = Math.Min(scaleW, scaleH);
                        // 若原图小于限定区域则不强制放大以防模糊失真
                        if (scale > 1.0f) scale = 1.0f;

                        // 计算最终几何宽高尺寸
                        float finalW = gdiImg.Width * scale;
                        float finalH = gdiImg.Height * scale;

                        // 调用 Excel COM Shapes.AddPicture 插入内嵌保存的图片 (LinkToFile: 0, SaveWithDocument: -1)
                        sheet.Shapes.AddPicture(tempFilePath, 0, -1, left, top, finalW, finalH);
                    }
                }
                finally
                {
                    // 安全删除临时生成的本地图片文件
                    try { if (File.Exists(tempFilePath)) File.Delete(tempFilePath); } catch { }
                }
            }
            catch (Exception ex)
            {
                // 记录插入 Logo 异常
                LogHelper.WriteLog($"[EnterpriseSettings] 插入/同步 Logo 图片异常: {ex.Message}");
            }
        }
    }
}
