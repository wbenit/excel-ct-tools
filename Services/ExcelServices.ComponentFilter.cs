using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 公共服务分部类: 元器件多选行当前列筛选与同箱柜共存标记服务
    /// </summary>
    public static partial class ExcelServices
    {
        // 单元格原始颜色快照数据模型，用于在取消筛选时 100% 还原用户自定义标记色
        private class CellColorSnapshot
        {
            // 单元格所在物理行号
            public int Row { get; set; }
            // 单元格所在物理列号
            public int Col { get; set; }
            // 原先是否为无填充色 (xlNone: -4142)
            public bool HasNoColor { get; set; }
            // 原先的具体 OLE 颜色数值
            public object? OriginalColor { get; set; }
        }

        // 缓存工作表在筛选前的单元格原始背景色快照字典 (Key: 工作表名称，确保跨表独立)
        private static readonly Dictionary<string, List<CellColorSnapshot>> _filterOriginalColorSnapshots 
            = new Dictionary<string, List<CellColorSnapshot>>(StringComparer.OrdinalIgnoreCase);

        // 相邻箱柜区分：白底 (第一台/奇数台)
        // --硬编码: 第一台/奇数台白底色--
        private static readonly Color AlternateWhiteColor = Color.White;

        // 相邻箱柜区分：青底 (第二台/偶数台)，采用柔和淡青绿 Hex #E0F2F1，对齐 #009688 主题
        // --硬编码: 第二台/偶数台淡青底色 RGB--
        private static readonly Color AlternateCyanColor = Color.FromArgb(224, 242, 241);

        // Excel 常数: 无填充色 ColorIndex (xlNone)
        // --硬编码: Excel xlNone 常数--
        private const int XlNoneColorIndex = -4142;

        // Excel 常数: 手动计算模式 (xlCalculationManual)
        // --硬编码: Excel 手动计算常数--
        private const int XlCalcManual = -4135;

        // Excel 常数: 自动计算模式 (xlCalculationAutomatic)
        // --硬编码: Excel 自动计算常数--
        private const int XlCalcAutomatic = -4105;

        /// <summary>
        /// 根据用户当前选区（单选或多选行）在当前激活列的内容，执行多选联合筛选并进行同箱柜共存标记
        /// </summary>
        public static void FilterComponentsBySelection()
        {
            try
            {
                // 获取 Excel Application 动态实例
                dynamic? app = ExcelDnaUtil.Application;
                // 校验 Excel 全局实例有效性
                if (app == null) return;

                // 获取当前工作簿活动选区 Range
                dynamic? selection = app.Selection;
                // 获取当前活动单元格 ActiveCell
                dynamic? activeCell = app.ActiveCell;
                // 获取当前活动工作表 ActiveSheet
                dynamic? activeSheet = app.ActiveSheet;

                // 若选区或工作表为空，直接退出
                if (selection == null || activeSheet == null)
                {
                    // 弹出友好提示告知用户未选定单元格
                    MessageBox.Show("未检测到有效单元格选区，请先选中需要筛选的单元格或行！", "筛选提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    // 中断退出
                    return;
                }

                // 确定目标物理列号（以当前活动单元格所在列为准，支持用户跨列拖选时聚焦当前列）
                int targetCol = (activeCell != null) ? (int)activeCell.Column : 2;

                // 准备去重保存用户选中的筛选关键字列表
                List<string> filterKeywords = ExtractFilterKeywordsFromSelection(selection, targetCol, activeCell);

                // 若未能提取到任何有效文本，友好提示并返回
                if (filterKeywords.Count == 0)
                {
                    // 弹出友好提示告知用户所选单元格为空
                    MessageBox.Show("所选单元格在当前列的内容为空，请选择包含有效元器件名称或型号的单元格后重试！", "筛选提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    // 中断退出
                    return;
                }

                // 规则 8 架构安全防护：在操作前自愈与校验当前工作表的定义名称拓扑结构
                Tool.FixAndFillCabinetNamesForSheet(activeSheet);

                // 获取所属工作簿对象
                dynamic? activeWb = null;
                // 向上追溯所属工作簿句柄
                try { activeWb = activeSheet.Parent; } catch { }

                // 获取当前工作表中所有合规箱柜锚点 (按物理行号升序排列)
                var validCabinets = Tool.GetSheetValidCabinets((object)activeSheet, (object?)activeWb);

                // 若当前工作表不存在箱柜定义名称（可能为元件汇总表或普通数据表），走通用平铺行筛选降级方案
                if (validCabinets == null || validCabinets.Count == 0)
                {
                    // 执行普通扁平数据表的当前列多选筛选降级逻辑
                    FilterFlatSheetByKeywords(app, activeSheet, targetCol, filterKeywords);
                    // 结束并返回
                    return;
                }

                // 执行分类明细表专属的箱柜区间多选筛选与同柜标记逻辑
                FilterCategorySheetByCabinets(app, activeSheet, validCabinets, targetCol, filterKeywords);
            }
            catch (Exception ex)
            {
                // 记录筛选处理异常日志
                LogHelper.WriteLog($"执行多选行当前列筛选异常: {ex.Message}\r\n{ex.StackTrace}");
                // 弹出异常提示信息
                MessageBox.Show($"筛选执行失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 从用户选区中提取目标列的非空文本作为筛选关键词集合（去重保序）
        /// </summary>
        private static List<string> ExtractFilterKeywordsFromSelection(dynamic selection, int targetCol, dynamic? activeCell)
        {
            // 用于去重判定的哈希集合 (忽略大小写)
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            // 保存有序的关键词结果列表
            List<string> result = new List<string>();

            try
            {
                // 获取选区中的子区域 Areas
                dynamic areas = selection.Areas;
                // 获取子区域数量
                int areaCount = areas.Count;

                // 遍历每一个选区 Area
                for (int a = 1; a <= areaCount; a++)
                {
                    // 获取当前子区域 Range
                    dynamic area = areas[a];
                    // 获取子区域行数与列数
                    int rowCount = area.Rows.Count;
                    // 获取子区域起始物理列与结束物理列
                    int startCol = area.Column;
                    // 计算结束列号
                    int endCol = startCol + area.Columns.Count - 1;

                    // 若该子区域未覆盖目标列 targetCol，则跳过
                    if (targetCol < startCol || targetCol > endCol)
                    {
                        // 不在目标列范围内跳过
                        continue;
                    }

                    // 计算目标列在当前 Area 内部的相对列索引 (1-based)
                    int relCol = targetCol - startCol + 1;

                    // 若行数较多采用数组批量读入内存 (规则 7)
                    if (rowCount > 1)
                    {
                        // 提取目标列在该 Area 中的一维/二维值数组
                        dynamic colRange = area.Columns[relCol];
                        // 一次性读入内存
                        object[,] valMatrix = colRange.Value2 as object[,];

                        // 遍历二维数组行
                        if (valMatrix != null)
                        {
                            // 遍历提取文本
                            for (int r = 1; r <= rowCount; r++)
                            {
                                // 转换为纯文本并去除首尾空白
                                string text = Convert.ToString(valMatrix[r, 1])?.Trim() ?? "";
                                // 若非空且未记录过，加入列表
                                if (!string.IsNullOrWhiteSpace(text) && seen.Add(text))
                                {
                                    // 添加到关键词列表
                                    result.Add(text);
                                }
                            }
                        }
                    }
                    else
                    {
                        // 单行单单元格读取
                        dynamic cell = area.Cells[1, relCol];
                        // 转换为纯文本并去除首尾空白
                        string text = Convert.ToString(cell.Value2)?.Trim() ?? "";
                        // 若非空且未记录过，加入列表
                        if (!string.IsNullOrWhiteSpace(text) && seen.Add(text))
                        {
                            // 添加到关键词列表
                            result.Add(text);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录选区提取异常
                LogHelper.WriteLog($"提取选区关键词异常: {ex.Message}");
            }

            // 若选区遍历后仍为空，尝试读取当前活动单元格 ActiveCell 的值作为兜底
            if (result.Count == 0 && activeCell != null)
            {
                // 读取活动单元格文本
                string fallbackText = Convert.ToString(activeCell.Value2)?.Trim() ?? "";
                // 若非空则加入结果
                if (!string.IsNullOrWhiteSpace(fallbackText))
                {
                    // 加入兜底关键词
                    result.Add(fallbackText);
                }
            }

            // 返回提取出的有效关键词列表
            return result;
        }

        /// <summary>
        /// 分类明细表专属逻辑：遍历箱柜明细块执行多选筛选，识别同柜共存并打上淡青绿高亮标记
        /// </summary>
        private static void FilterCategorySheetByCabinets(
            dynamic app,
            dynamic activeSheet,
            List<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets,
            int targetCol,
            List<string> filterKeywords)
        {
            // 汇总全表所有命中的元器件物理行号
            HashSet<int> allHitRows = new HashSet<int>();
            // 汇总命中了【同箱柜共存】（同时集齐全部筛选关键词）的元器件物理行号
            HashSet<int> sameCabinetMatchedRows = new HashSet<int>();
            // 记录同时集齐全部筛选词的箱柜数量
            int sameCabinetCount = 0;

            // 按照从上到下的顺序收集包含命中行的箱柜数据模型列表: (箱柜键, 箱柜锚点, 命中物理行列表, 是否集齐全部已选元件)
            var hitCabinetsOrdered = new List<(int CabIdx, Models.CabinetAnchorModel Anchor, List<int> HitRows, bool IsFullMatch)>();

            // 遍历工作表中每一个合法箱柜
            foreach (var kv in validCabinets)
            {
                // 提取箱柜锚点数据模型
                var anchor = kv.Value;
                // 校验 Det 与 Subsum 锚点 Range 是否有效
                if (anchor?.Det == null || anchor?.Subsum == null) continue;

                // 提取箱柜标题信息行物理行号
                int detRow = anchor.Det.Row;
                // 提取底部明细小计行物理行号
                int subsumRow = anchor.Subsum.Row;

                // 根据规则 6：Cab_Det.row + 2 为元器件起始行
                int compStartRow = detRow + 2;
                // 根据规则 6：Cab_Subsum.row - 1 为元器件终止行
                int compEndRow = subsumRow - 1;

                // 若该箱柜元器件区间无有效行，跳过当前箱柜
                if (compEndRow < compStartRow) continue;

                // 计算该箱柜元器件总行数
                int compRowCount = compEndRow - compStartRow + 1;

                // 规则 7：采用二维数组一次性将该箱柜目标列的所有单元格数据读取到内存中
                dynamic colRange = activeSheet.Range[activeSheet.Cells[compStartRow, targetCol], activeSheet.Cells[compEndRow, targetCol]];
                // 转换为 object[,] 二维矩阵
                object[,] valMatrix = colRange.Value2 as object[,];
                // 校验数组有效性
                if (valMatrix == null) continue;

                // 记录当前箱柜命中的关键词集合 (去重统计命中种类数)
                HashSet<string> hitKwsInCabinet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                // 记录当前箱柜中所有命中的元器件物理行号
                List<int> hitRowsInCabinet = new List<int>();

                // 遍历该箱柜每一行元器件
                for (int r = 1; r <= compRowCount; r++)
                {
                    // 获取当前行单元格纯文本
                    string cellValue = Convert.ToString(valMatrix[r, 1])?.Trim() ?? "";
                    // 若内容为空则跳过空行
                    if (string.IsNullOrWhiteSpace(cellValue)) continue;

                    // 计算在 Excel 工作表中的真实绝对物理行号
                    int physicalRow = compStartRow + r - 1;

                    // 检查是否命中了用户选中的任一关键词 (使用 Contains 包含匹配，支持型号或名称模糊识别)
                    foreach (string kw in filterKeywords)
                    {
                        // 进行忽略大小写的子串包含匹配
                        if (cellValue.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // 记录命中的关键词类别
                            hitKwsInCabinet.Add(kw);
                            // 记录命中的行号
                            hitRowsInCabinet.Add(physicalRow);
                            // 命中后跳出关键字循环，避免单行重复记录
                            break;
                        }
                    }
                }

                // 若当前箱柜存在命中行
                if (hitRowsInCabinet.Count > 0)
                {
                    // 将当前箱柜命中行并入全表总命中集合
                    foreach (int r in hitRowsInCabinet) allHitRows.Add(r);

                    // 核心业务判定：如果用户选择了 2 个或以上关键词，且当前箱柜【同时集齐了所有筛选关键词】
                    bool isFullSameCabinetMatch = (filterKeywords.Count >= 2 && hitKwsInCabinet.Count >= filterKeywords.Count);

                    // 若判定为同箱柜全部共存匹配
                    if (isFullSameCabinetMatch)
                    {
                        // 累加同柜箱柜计数
                        sameCabinetCount++;
                        // 将该箱柜所有命中行加入同柜高亮行集合
                        foreach (int r in hitRowsInCabinet) sameCabinetMatchedRows.Add(r);
                    }

                    // 登记该命中箱柜在连续有序列表中 (用于相邻箱柜斑马纹交替区分)
                    hitCabinetsOrdered.Add((kv.Key, anchor, hitRowsInCabinet, isFullSameCabinetMatch));
                }
            }

            // 若全表扫描完毕未找到任何匹配行，弹出友好提示并返回
            if (allHitRows.Count == 0)
            {
                // 拼接所选关键词字符串
                string kwStr = string.Join("、", filterKeywords);
                // 弹出提示告知用户未找到匹配项
                MessageBox.Show($"未在当前分类表中检索到包含 [{kwStr}] 的元器件行！", "筛选结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // 中断退出
                return;
            }

            // 计算整表需要扫描管理的最大物理行号 (取 UsedRange 与各箱柜 Tolsum 的最大值)
            int maxRow = 100;
            try
            {
                // 获取工作表已用区域 UsedRange
                dynamic used = activeSheet.UsedRange;
                // 计算 UsedRange 底部行号
                maxRow = Math.Max(maxRow, (int)used.Row + (int)used.Rows.Count);
            }
            catch { }

            // 遍历箱柜提取最大的 Tolsum 行号
            foreach (var kv in validCabinets)
            {
                // 若 Tolsum 有效，更新最大行号
                if (kv.Value?.Tolsum != null)
                {
                    // 取当前最大值与 Tolsum 行号的最大值
                    maxRow = Math.Max(maxRow, (int)kv.Value.Tolsum.Row + 5);
                }
            }

            // 记录当前的屏幕更新与事件状态，并在进入大批量操作前挂起以保证极速流畅
            bool prevUpdating = true;
            // 记录先前的自动重算模式
            int prevCalculation = XlCalcAutomatic;
            try
            {
                // 保存先前屏幕更新状态
                prevUpdating = app.ScreenUpdating;
                // 冻结屏幕重绘
                app.ScreenUpdating = false;
                // 冻结 COM 事件触发
                app.EnableEvents = false;
                // 读取原计算模式
                prevCalculation = (int)app.Calculation;
                // 切换为手动重算模式，消除万次公式级联重算风暴
                app.Calculation = (dynamic)XlCalcManual;
            }
            catch { }

            try
            {
                // 步骤 1：若本工作表先前已有快照未还原（如多次连续筛选），先还原旧快照，杜绝用户原色丢失
                RestoreColorSnapshotsForSheet(activeSheet);

                // 步骤 2：对本次命中的所有元器件行 A~M 列执行原始背景色快照备份（100% 保护用户自定义标记色）
                BackupColorSnapshotsForRows(activeSheet, allHitRows);

                // 步骤 3：执行“全部隐藏，只留下筛选出的行”
                // 将 1 到 maxRow 中所有不属于 allHitRows 的行划分为离散区间，批量设置 Hidden = true
                ApplyBatchRowVisibility(activeSheet, 1, maxRow, allHitRows);

                // 步骤 4：相邻箱柜斑马纹交替上色（第一台白底，第二台青底，第三台白底...）
                int whiteOle = ColorTranslator.ToOle(AlternateWhiteColor);
                // 获取淡青底 OLE 数值
                int cyanOle = ColorTranslator.ToOle(AlternateCyanColor);

                // 遍历命中箱柜序列
                for (int cabOrder = 0; cabOrder < hitCabinetsOrdered.Count; cabOrder++)
                {
                    // 提取当前命中箱柜数据
                    var cabItem = hitCabinetsOrdered[cabOrder];
                    // 偶数序号 (0, 2, 4...) 赋予白底，奇数序号 (1, 3, 5...) 赋予青底，相邻箱柜边界极其鲜明
                    int targetOleColor = (cabOrder % 2 == 0) ? whiteOle : cyanOle;

                    // 遍历该箱柜内的所有命中行
                    foreach (int row in cabItem.HitRows)
                    {
                        try
                        {
                            // 选取该行 A 列至 M 列核心区域
                            dynamic rowRange = activeSheet.Range[$"A{row}:M{row}"];
                            // 赋予对应箱柜的交替背景底色
                            rowRange.Interior.Color = targetOleColor;
                        }
                        catch { }
                    }
                }

                // 步骤 5：在 Excel 底部状态栏展示醒目的筛选汇总统计
                string statusMsg = (sameCabinetCount > 0)
                    ? $"[筛选完成] 已筛选出 {allHitRows.Count} 行元器件 (共 {hitCabinetsOrdered.Count} 台箱柜，已按白/青底交替区分，其中 {sameCabinetCount} 台同时包含已选元器件)"
                    : $"[筛选完成] 已筛选出 {allHitRows.Count} 行元器件 (共 {hitCabinetsOrdered.Count} 台箱柜，已按白/青底交替区分)";
                // 写入 Excel 状态栏
                try { app.StatusBar = statusMsg; } catch { }
            }
            finally
            {
                // 无论是否发生异常，100% 恢复 Excel 原始渲染、计算与事件环境
                try
                {
                    // 恢复公式计算模式
                    app.Calculation = (dynamic)prevCalculation;
                    // 恢复事件侦听
                    app.EnableEvents = true;
                    // 恢复屏幕刷新
                    app.ScreenUpdating = prevUpdating;
                }
                catch { }
            }
        }

        /// <summary>
        /// 针对普通平铺工作表（无箱柜拓扑）的当前列多选筛选降级逻辑
        /// </summary>
        private static void FilterFlatSheetByKeywords(dynamic app, dynamic activeSheet, int targetCol, List<string> filterKeywords)
        {
            // 获取 UsedRange 范围
            dynamic used = activeSheet.UsedRange;
            // 获取起始行与总行数
            int startRow = used.Row;
            // 计算总行数
            int rowCount = used.Rows.Count;
            // 计算结束行
            int endRow = startRow + rowCount - 1;

            // 若行数极少则不处理
            if (rowCount <= 1) return;

            // 规则 7：一次性读入目标列数据矩阵
            dynamic colRange = activeSheet.Range[activeSheet.Cells[startRow, targetCol], activeSheet.Cells[endRow, targetCol]];
            // 转换为二维矩阵
            object[,] matrix = colRange.Value2 as object[,];
            // 校验矩阵
            if (matrix == null) return;

            // 收集命中的物理行号
            HashSet<int> hitRows = new HashSet<int>();

            // 遍历所有数据行
            for (int r = 1; r <= rowCount; r++)
            {
                // 获取当前行文本
                string val = Convert.ToString(matrix[r, 1])?.Trim() ?? "";
                // 若空则跳过
                if (string.IsNullOrWhiteSpace(val)) continue;

                // 绝对物理行号
                int physicalRow = startRow + r - 1;

                // 检查是否包含任一关键字
                foreach (string kw in filterKeywords)
                {
                    // 子串包含匹配
                    if (val.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // 记录命中行
                        hitRows.Add(physicalRow);
                        // 结束当前行检查
                        break;
                    }
                }
            }

            // 若无命中行
            if (hitRows.Count == 0)
            {
                // 提示未找到
                MessageBox.Show("未在当前工作表中检索到符合条件的行！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // 退出
                return;
            }

            // 挂起刷新执行批量行显隐
            bool prevUpdating = true;
            try
            {
                // 记录并挂起刷新
                prevUpdating = app.ScreenUpdating;
                // 关闭重绘
                app.ScreenUpdating = false;
                // 关闭事件
                app.EnableEvents = false;

                // 将 1 到 endRow 批量隐藏除 hitRows 以外的所有行
                ApplyBatchRowVisibility(activeSheet, 1, endRow, hitRows);

                // 状态栏提示
                try { app.StatusBar = $"[筛选完成] 已筛选出 {hitRows.Count} 行数据"; } catch { }
            }
            finally
            {
                // 恢复环境
                try
                {
                    // 开启事件
                    app.EnableEvents = true;
                    // 开启屏幕更新
                    app.ScreenUpdating = prevUpdating;
                }
                catch { }
            }
        }

        /// <summary>
        /// 批量应用行可见性：将非保留行聚合为连续物理区间批量执行 EntireRow.Hidden = true，杜绝逐行 COM 卡顿
        /// </summary>
        private static void ApplyBatchRowVisibility(dynamic sheet, int minRow, int maxRow, HashSet<int> keepVisibleRows)
        {
            try
            {
                // 记录当前连续待隐藏区间的起始行号 (-1 表示当前不在隐藏区间内)
                int hideStart = -1;

                // 从 minRow 线性扫描至 maxRow
                for (int r = minRow; r <= maxRow; r++)
                {
                    // 判断当前行是否需要隐藏 (即不属于 keepVisibleRows)
                    bool shouldHide = !keepVisibleRows.Contains(r);

                    // 若当前行需要隐藏
                    if (shouldHide)
                    {
                        // 若尚未开启连续区间，以此行为起点
                        if (hideStart == -1) hideStart = r;
                    }
                    else
                    {
                        // 当前行需要保持显示：若先前存在未闭合的隐藏区间，在此闭合执行
                        if (hideStart != -1)
                        {
                            // 批量设置该连续区间的整行隐藏
                            sheet.Range[$"A{hideStart}:A{r - 1}"].EntireRow.Hidden = true;
                            // 重置区间标记
                            hideStart = -1;
                        }
                        // 确保该保留行处于显示状态
                        sheet.Rows[r].Hidden = false;
                    }
                }

                // 循环结束后，若尾部仍有未闭合的隐藏区间，执行最终批量隐藏
                if (hideStart != -1)
                {
                    // 批量隐藏尾部连续区间
                    sheet.Range[$"A{hideStart}:A{maxRow}"].EntireRow.Hidden = true;
                }
            }
            catch (Exception ex)
            {
                // 记录批量显隐异常
                LogHelper.WriteLog($"批量设置行显隐异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 清除当前工作表的元器件筛选状态：一键恢复所有行显示并精准无损还原用户原本的标记底色
        /// </summary>
        public static void ClearComponentFilter(object? sheet = null)
        {
            try
            {
                // 获取 Excel Application 动态实例
                dynamic? app = ExcelDnaUtil.Application;
                // 校验有效性
                if (app == null) return;

                // 若未显式传入工作表，使用当前活动工作表
                dynamic? targetSheet = sheet ?? app.ActiveSheet;
                // 校验工作表有效性
                if (targetSheet == null) return;

                // 挂起屏幕刷新以保证瞬发恢复
                bool prevUpdating = true;
                try
                {
                    // 保存刷新状态
                    prevUpdating = app.ScreenUpdating;
                    // 暂停屏幕更新
                    app.ScreenUpdating = false;
                    // 暂停事件处理
                    app.EnableEvents = false;

                    // 1. 一行代码极速解除整表所有行的隐藏状态，100% 恢复全貌
                    targetSheet.Rows.EntireRow.Hidden = false;

                    // 2. 核心技术点：无损还原单元格原始背景色（完美保留用户原本所有的自定义标记色）
                    bool restoredBySnapshot = RestoreColorSnapshotsForSheet(targetSheet);

                    // 3. 容错兜底：若无快照存在（例如首次打开直接点清除），才执行常规清除
                    if (!restoredBySnapshot)
                    {
                        dynamic? targetWb = null;
                        // 追溯所属工作簿
                        try { targetWb = targetSheet.Parent; } catch { }
                        // 获取有效箱柜列表
                        var validCabinets = Tool.GetSheetValidCabinets((object)targetSheet, (object?)targetWb);

                        // 若存在箱柜，清除箱柜元器件区间背景色
                        if (validCabinets != null && validCabinets.Count > 0)
                        {
                            // 清除箱柜高亮底色
                            ClearCabinetHighlightColors(targetSheet, validCabinets);
                        }
                        else
                        {
                            // 若为普通表，清除已用区域的背景色
                            try { targetSheet.UsedRange.Interior.ColorIndex = XlNoneColorIndex; } catch { }
                        }
                    }

                    // 4. 在状态栏给出恢复提示
                    try { app.StatusBar = "已清除筛选，已恢复显示全部行并还原原有标记底色。"; } catch { }
                }
                finally
                {
                    // 恢复事件与屏幕更新
                    try
                    {
                        // 开启事件
                        app.EnableEvents = true;
                        // 开启屏幕更新
                        app.ScreenUpdating = prevUpdating;
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                // 记录清除筛选异常日志
                LogHelper.WriteLog($"清除筛选异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 对指定的所有行 A~M 列执行原始背景色快照备份（记录 ColorIndex 与 Color，保证无损还原）
        /// </summary>
        private static void BackupColorSnapshotsForRows(dynamic sheet, IEnumerable<int> rows)
        {
            try
            {
                // 获取工作表名称作为快照隔离标识
                string sheetKey = Convert.ToString(sheet.Name) ?? "DefaultSheet";
                // 准备单元格快照列表
                var list = new List<CellColorSnapshot>();

                // 遍历所有待上色的物理行号
                foreach (int row in rows)
                {
                    // 遍历 A 列至 M 列 (1 到 13 列) --硬编码: 13 列覆盖成套报价核心列--
                    for (int col = 1; col <= 13; col++)
                    {
                        try
                        {
                            // 获取单元格对象
                            dynamic cell = sheet.Cells[row, col];
                            // 获取背景色索引
                            int cIdx = (int)cell.Interior.ColorIndex;

                            // 若原先为无填充色
                            if (cIdx == XlNoneColorIndex)
                            {
                                // 登记无填充色记录
                                list.Add(new CellColorSnapshot { Row = row, Col = col, HasNoColor = true });
                            }
                            else
                            {
                                // 登记具体 OLE 颜色数值
                                list.Add(new CellColorSnapshot { Row = row, Col = col, HasNoColor = false, OriginalColor = cell.Interior.Color });
                            }
                        }
                        catch { }
                    }
                }

                // 保存至静态快照缓存
                _filterOriginalColorSnapshots[sheetKey] = list;
            }
            catch (Exception ex)
            {
                // 记录备份异常
                LogHelper.WriteLog($"备份原始底色快照异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 从快照中无损精准还原工作表单元格原始背景色（完美保留用户原本的所有自定义标记色）
        /// </summary>
        private static bool RestoreColorSnapshotsForSheet(dynamic sheet)
        {
            try
            {
                // 获取工作表名称
                string sheetKey = Convert.ToString(sheet.Name) ?? "DefaultSheet";

                // 检查是否存在当前工作表的快照
                if (_filterOriginalColorSnapshots.TryGetValue(sheetKey, out var list) && list != null && list.Count > 0)
                {
                    // 遍历快照记录
                    foreach (var snap in list)
                    {
                        try
                        {
                            // 获取单元格
                            dynamic cell = sheet.Cells[snap.Row, snap.Col];
                            // 若原本为无填充色
                            if (snap.HasNoColor)
                            {
                                // 还原为无填充色
                                cell.Interior.ColorIndex = XlNoneColorIndex;
                            }
                            else if (snap.OriginalColor != null)
                            {
                                // 100% 原样精准还原用户原本自定义的标记色
                                cell.Interior.Color = snap.OriginalColor;
                            }
                        }
                        catch { }
                    }

                    // 清除已还原的快照缓存
                    _filterOriginalColorSnapshots.Remove(sheetKey);
                    // 标记成功还原
                    return true;
                }
            }
            catch (Exception ex)
            {
                // 记录还原异常
                LogHelper.WriteLog($"还原原始底色快照异常: {ex.Message}");
            }

            // 未命中快照
            return false;
        }

        /// <summary>
        /// 清除所有箱柜元器件区间的临时高亮背景色 (安全还原为 xlNone 无填充色)
        /// </summary>
        private static void ClearCabinetHighlightColors(dynamic sheet, List<KeyValuePair<int, Models.CabinetAnchorModel>> validCabinets)
        {
            // 遍历每一个箱柜
            foreach (var kv in validCabinets)
            {
                try
                {
                    // 提取锚点
                    var anchor = kv.Value;
                    // 校验锚点 Range
                    if (anchor?.Det == null || anchor?.Subsum == null) continue;

                    // 提取起始行与结束行
                    int compStartRow = anchor.Det.Row + 2;
                    // 计算结束行
                    int compEndRow = anchor.Subsum.Row - 1;

                    // 若行数合法
                    if (compEndRow >= compStartRow)
                    {
                        // 选取该箱柜元器件区间的 A 列至 M 列
                        dynamic compRange = sheet.Range[$"A{compStartRow}:M{compEndRow}"];
                        // 设置背景色为无填充色 (xlNone: -4142)
                        compRange.Interior.ColorIndex = XlNoneColorIndex;
                    }
                }
                catch { }
            }
        }
    }
}
