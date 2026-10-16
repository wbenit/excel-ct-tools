using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo
{
    /// <summary>
    /// ExcelServices 业务服务层扩展分部类：云物料一体化工作台、物料插入与替换回写
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释
    /// </summary>
    public static partial class ExcelServices
    {
        // 缓存当前已打开的云物料工作台单例窗体实例
        private static Forms.CloudMaterialHubForm? _cloudMaterialHubForm;

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“云物料”一体化工作台
        /// </summary>
        /// <param name="initialTab">默认激活的选项卡 ("material" | "cabinet" | "copper" | "ampacity")</param>
        public static void ShowCloudMaterialHubDialog(string initialTab = "material")
        {
            try
            {
                // 获取当前活动工作表
                dynamic app = ExcelDnaUtil.Application;
                dynamic activeSheet = app.ActiveSheet;
                // 校验工作表有效性 (规则 9: 不符合要求直接阻拦)
                if (activeSheet == null)
                {
                    MessageBox.Show("未检测到活动的 Excel 工作表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 操作前先行执行规则 6 自愈检查，保障定义名称结构纯净 (规则 8)
                Tool.FixAndFillCabinetNamesForSheet(activeSheet);

                // 以非模态方式展示云物料工作台窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _cloudMaterialHubForm, () => new Forms.CloudMaterialHubForm(initialTab));
            }
            catch (Exception ex)
            {
                // 捕获异常防止崩溃
                MessageBox.Show($"弹出云物料工作台失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// <summary>
        /// 将单条物料插入到当前 Excel 分类表的元器件区域 (插入到当前光标所在行的下方)
        /// 严格遵循规则 6、7、8、9
        /// </summary>
        /// <param name="item">待插入的物料条目</param>
        /// <returns>操作是否成功</returns>
        public static bool InsertMaterialFromHub(HubMaterialDto item)
        {
            // 包装为单条列表调用批量插入方法
            if (item == null) return false;
            // 转发调用批量物料插入方法
            return InsertMaterialsFromHub(new List<HubMaterialDto> { item });
        }

        /// <summary>
        /// 批量将多条物料明细插入到当前 Excel 分类表的元器件区域 (支持多行明细连续插入)
        /// 严格遵循规则 6、7、8、9
        /// </summary>
        /// <param name="items">待插入的物料明细列表</param>
        /// <returns>操作是否成功</returns>
        public static bool InsertMaterialsFromHub(List<HubMaterialDto> items)
        {
            try
            {
                // 空校验判断
                if (items == null || items.Count == 0) return false;
                // 获取 Excel 宿主对象
                dynamic app = ExcelDnaUtil.Application;
                // 获取当前活动工作表
                dynamic ws = app.ActiveSheet;
                // 规则 9: 检查工作表有效性，无效直接拦截报错
                if (ws == null)
                {
                    // 弹出警示消息
                    MessageBox.Show("当前没有激活的 Excel 工作表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // 规则 8: 校验并纠正规则 6 定义名称架构
                Tool.FixAndFillCabinetNamesForSheet(ws);

                // 获取当前活动单元格
                dynamic activeCell = app.ActiveCell;
                // 提取光标所在行号
                int cursorRow = activeCell != null ? (int)activeCell.Row : 0;

                // 扫描当前工作表中的所有箱柜定义名称 Cab_Det_k 和 Cab_Subsum_k
                int targetCabinetK = -1;
                // 目标箱柜明细行号
                int targetDetRow = -1;
                // 目标箱柜小计行号
                int targetSubsumRow = -1;
                // 首个有效明细行备选
                int firstValidDetRow = -1;
                // 首个有效小计行备选
                int firstValidSubsumRow = -1;

                // 循环遍历工作表定义名称 (上限 500)
                for (int k = 1; k <= 500; k++)
                {
                    // 尝试提取 Cab_Det_k
                    dynamic detNameObj = null;
                    // 从工作表作用域捕获
                    try { detNameObj = ws.Names.Item($"Cab_Det_{k}"); } catch { }
                    // 若为空则从工作簿作用域捕获
                    if (detNameObj == null)
                    {
                        try { detNameObj = ws.Parent.Names.Item($"Cab_Det_{k}"); } catch { }
                    }
                    // 若依然为空则跳过当前编号
                    if (detNameObj == null) continue;

                    // 尝试提取 Cab_Subsum_k
                    dynamic subsumNameObj = null;
                    // 从工作表级别尝试提取
                    try { subsumNameObj = ws.Names.Item($"Cab_Subsum_{k}"); } catch { }
                    // 工作表未命中尝试从工作簿提取
                    if (subsumNameObj == null)
                    {
                        try { subsumNameObj = ws.Parent.Names.Item($"Cab_Subsum_{k}"); } catch { }
                    }
                    // 未命中则跳过
                    if (subsumNameObj == null) continue;

                    // 获取当前箱柜的明细行与小计行物理行号
                    int detR = (int)detNameObj.RefersToRange.Row;
                    // 小计行行号
                    int subR = (int)subsumNameObj.RefersToRange.Row;

                    // 记录首个有效箱柜作为兜底目标
                    if (firstValidDetRow <= 0)
                    {
                        // 记录首个明细行
                        firstValidDetRow = detR;
                        // 记录首个小计行
                        firstValidSubsumRow = subR;
                        // 记录编号
                        targetCabinetK = k;
                    }

                    // 判定光标是否落入此箱柜范围 (包括从 Det 行至 Subsum 行)
                    if (cursorRow >= detR && cursorRow <= subR)
                    {
                        // 命中目标箱柜
                        targetCabinetK = k;
                        // 锁定明细行
                        targetDetRow = detR;
                        // 锁定小计行
                        targetSubsumRow = subR;
                        break;
                    }
                }

                // 若未命中当前光标所属箱柜，回退至首个有效箱柜
                if (targetDetRow <= 0)
                {
                    // 回退明细行
                    targetDetRow = firstValidDetRow;
                    // 回退小计行
                    targetSubsumRow = firstValidSubsumRow;
                }

                // 规则 9: 检查是否具有有效明细区，否则直接阻拦报错
                if (targetDetRow <= 0 || targetSubsumRow <= 0)
                {
                    // 提示用户无有效箱柜区域
                    MessageBox.Show("未在当前工作表中找到有效的箱柜明细区域（Cab_Det / Cab_Subsum）！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // 元器件有效区域：从 Cab_Det.Row + 2 至 Cab_Subsum.Row - 1 (规则 6)
                int compStartRow = targetDetRow + 2;
                // 元器件终止行
                int compEndRow = targetSubsumRow - 1;

                // 统计待插入的物料明细行总数
                int count = items.Count;

                // 1. 扫描当前箱柜元器件区域，统计空白行总数 (规则 7: 二维数组批量读入内存)
                int totalCompRows = Math.Max(0, compEndRow - compStartRow + 1);
                // 记录元器件区域内部总空白行数
                int emptyRowCount = 0;
                // 标记每一行是否为空行
                bool[] isRowEmpty = new bool[totalCompRows];

                if (totalCompRows > 0)
                {
                    // 一次性读取元器件区域 B列(名称) 与 C列(型号)
                    dynamic checkRange = ws.Range[$"B{compStartRow}:C{compEndRow}"];
                    // 转换为二维数组对象
                    object[,] checkVals = (object[,])checkRange.Value2;
                    // 循环遍历元器件区域各行
                    for (int r = 1; r <= totalCompRows; r++)
                    {
                        // 提取 B 列名称文本
                        string bText = checkVals[r, 1]?.ToString()?.Trim() ?? string.Empty;
                        // 提取 C 列型号文本
                        string cText = checkVals[r, 2]?.ToString()?.Trim() ?? string.Empty;
                        // 判断本行是否属于空白行
                        bool empty = string.IsNullOrEmpty(bText) && string.IsNullOrEmpty(cText);
                        // 记录到数组
                        isRowEmpty[r - 1] = empty;
                        // 若为空白行则累计计数
                        if (empty) emptyRowCount++;
                    }
                }

                // 2. 决定物理插入的目标起始行号 (用户指令: 应该在光标行插入)
                int insertRow = cursorRow;
                // 如果光标在元器件区之外，则智能对齐到元器件有效区域
                if (insertRow < compStartRow || insertRow > compEndRow)
                {
                    // 寻找元器件区域最后一个有内容的行
                    int lastUsedRow = compStartRow - 1;
                    // 反向扫描已使用行
                    for (int r = totalCompRows - 1; r >= 0; r--)
                    {
                        // 命中最后非空行
                        if (!isRowEmpty[r])
                        {
                            lastUsedRow = compStartRow + r;
                            break;
                        }
                    }
                    // 定位在最后已用行的下一行
                    insertRow = Math.Min(compEndRow, lastUsedRow + 1);
                    // 边界保护
                    if (insertRow < compStartRow) insertRow = compStartRow;
                }

                // 3. 计算光标行起连续可用空白行数
                int cursorRowIdx = insertRow - compStartRow;
                // 统计从光标行开始的连续空行
                int consecutiveEmptyFromCursor = 0;
                if (cursorRowIdx >= 0 && cursorRowIdx < totalCompRows)
                {
                    // 向后遍历空行
                    for (int r = cursorRowIdx; r < totalCompRows; r++)
                    {
                        // 若连续为空则自增
                        if (isRowEmpty[r]) consecutiveEmptyFromCursor++;
                        // 遇到非空行即终止
                        else break;
                    }
                }

                // 4. 规则 6: 检查待插入行数是否超过元器件区域的空白行数 (计算相差行 diffRows)
                int diffRows = Math.Max(0, count - emptyRowCount);

                // 判定：若光标行本身是空行且后续连续空行足够容纳 count 行，直接复用写入
                if (cursorRowIdx >= 0 && isRowEmpty[cursorRowIdx] && consecutiveEmptyFromCursor >= count)
                {
                    // 空间充裕，无需整行平移或插行
                }
                // 否则（光标行已有物料需下移，或区域空行不足）：执行平移插行并精准平衡相差行
                else
                {
                    // 在光标行处插入 count 行以容纳新数据并使原有内容顺延下移
                    for (int idx = 0; idx < count; idx++)
                    {
                        // 获取光标所在行对象
                        dynamic newRow = ws.Rows[insertRow];
                        // 向下平移插入整行
                        newRow.Insert(1); // 1 = xlShiftDown
                        // 复制相邻行单元格格式
                        dynamic refFormatRow = ws.Rows[insertRow > compStartRow ? (insertRow - 1) : (insertRow + count)];
                        // 执行格式复制
                        refFormatRow.Copy();
                        // 获取新插入行
                        dynamic insertedRow = ws.Rows[insertRow];
                        // 仅粘贴格式保持样式一致
                        insertedRow.PasteSpecial(-4122); // -4122 = xlPasteFormats
                        // 清除剪贴板状态
                        app.CutCopyMode = false;
                    }

                    // 规则 6: 保持净增加行数精确等于相差行 diffRows (避免无节制扩展表格)
                    int deleteEmptyRows = Math.Min(count, emptyRowCount);
                    // 若原本存在可用空行，则从小计行上方扣减对应的多余空行
                    if (deleteEmptyRows > 0)
                    {
                        // 计算插行后当前小计行的实际物理行号
                        int currentSubsumRow = targetSubsumRow + count;
                        // 计算待删除空行的起始行
                        int delStartRow = currentSubsumRow - deleteEmptyRows;
                        // 圈定待删除区域
                        dynamic delRange = ws.Range[$"A{delStartRow}:A{currentSubsumRow - 1}"];
                        // 执行整行向上删除
                        delRange.EntireRow.Delete(-4121); // -4121 = xlShiftUp
                    }
                }

                // 5. 规则 7: 采用数组一次性将物料数据写入内存并回写单元格
                object[,] writeBuffer = new object[count, 13]; // 覆盖 A 到 M 列
                // 遍历物料明细填充二维数组
                for (int i = 0; i < count; i++)
                {
                    // 获取当前物料 DTO 对象
                    var curItem = items[i];
                    // A 列序号公式指向明细表头行 (targetDetRow + 1)
                    int compHeaderRow = targetDetRow + 1;
                    writeBuffer[i, 0] = $"=ROW()-ROW(A${compHeaderRow})";
                    // B 列物料名称
                    writeBuffer[i, 1] = curItem.Name;
                    // C 列型号规格
                    writeBuffer[i, 2] = curItem.Model;
                    // D 列品牌生产厂家
                    writeBuffer[i, 3] = string.IsNullOrWhiteSpace(curItem.Brand) ? "国优" : curItem.Brand;
                    // E 列计量单位
                    writeBuffer[i, 4] = string.IsNullOrWhiteSpace(curItem.Unit) ? "米" : curItem.Unit;
                    // F 列数量 (长度或重量)
                    writeBuffer[i, 5] = curItem.Quantity > 0 ? curItem.Quantity : 1.0;
                    // G 列单价 (空或与 M 列联动)
                    writeBuffer[i, 6] = DBNull.Value;
                    // H 列总价
                    writeBuffer[i, 7] = DBNull.Value;
                    // I 列备注
                    writeBuffer[i, 8] = curItem.Remark ?? string.Empty;
                    // J 列成本单价
                    writeBuffer[i, 9] = DBNull.Value;
                    // K 列成本总价
                    writeBuffer[i, 10] = DBNull.Value;
                    // L 列毛利
                    writeBuffer[i, 11] = DBNull.Value;
                    // M 列单价数值 (纯数值或由公式联动)
                    writeBuffer[i, 12] = (double)curItem.Price;
                }

                // 一次性批量回写到目标单元格行区域 (覆盖 insertRow 到 insertRow + count - 1)
                dynamic writeRange = ws.Range[$"A{insertRow}:M{insertRow + count - 1}"];
                // 批量内存赋值
                writeRange.Value2 = writeBuffer;

                // 6. 规则 6 & 8: 刷新当前箱柜元器件 A 列序号自适应公式与计费小计公式，根治序号断裂与 #REF!
                try
                {
                    // 重新探测当前箱柜最新的物理行号拓扑 (兼容插删行后的物理位移)
                    var curIndexes = Tool.FindStandardCategoryRowIndexes(ws, targetCabinetK);
                    // 提取当前明细行号
                    int curDetRow = curIndexes.cabDetRow > 0 ? curIndexes.cabDetRow : targetDetRow;
                    // 计算净增行数
                    int netDiff = (cursorRowIdx >= 0 && isRowEmpty[cursorRowIdx] && consecutiveEmptyFromCursor >= count) ? 0 : (count - Math.Min(count, emptyRowCount));
                    // 提取当前小计行号
                    int curSubsumRow = curIndexes.cabSubsumRow > 0 ? curIndexes.cabSubsumRow : (targetSubsumRow + netDiff);
                    // 提取当前总计行号
                    int curTolsumRow = curIndexes.cabTolsumRow > 0 ? curIndexes.cabTolsumRow : (curSubsumRow + 5);
                    // 计算元器件起始行
                    int curCompStartRow = curDetRow + 2;

                    // 刷新整柜元器件 A 列序号(=ROW()-ROW(A${headerRow}))、单价总价联动与小计自愈公式
                    RefreshCabinetFeeAreaFormulas(ws, curDetRow, curCompStartRow, curSubsumRow, curTolsumRow);

                    // 若实际产生了物理插行删行，自愈校准全表定义名称拓扑
                    if (netDiff != 0)
                    {
                        Tool.FixAndFillCabinetNamesForSheet(ws, forceRebuild: false);
                    }
                }
                catch (Exception exRefresh)
                {
                    // 记录自愈异常容错日志
                    LogHelper.WriteLog($"[CloudMaterialHub] 插入物料后自愈刷新异常: {exRefresh.Message}");
                }

                // 选中新插入的首行提示用户
                dynamic focusCell = ws.Range[$"B{insertRow}"];
                // 获得焦点高亮
                focusCell.Select();

                return true;
            }
            catch (Exception ex)
            {
                // 记录异常并提示
                MessageBox.Show($"插入物料到 Excel 异常: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// 替换 Excel 当前光标选中行物料数据 (用户决策 A2: 直接覆盖该行的名称、型号、单价等字段)
        /// </summary>
        /// <param name="item">新物料条目</param>
        /// <param name="updateQuantity">是否同步覆盖 F 列数量</param>
        /// <returns>操作是否成功</returns>
        public static bool ReplaceMaterialInExcel(HubMaterialDto item, bool updateQuantity)
        {
            try
            {
                // 获取活动工作表与单元格
                dynamic app = ExcelDnaUtil.Application;
                dynamic ws = app.ActiveSheet;
                dynamic activeCell = app.ActiveCell;
                if (ws == null || activeCell == null)
                {
                    MessageBox.Show("当前没有激活的单元格！请先选中需要替换的元器件行。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                int targetRow = (int)activeCell.Row;
                // 规则 8: 校验并纠正规则 6
                Tool.FixAndFillCabinetNamesForSheet(ws);

                // 读取当前行的数量与原有内容 (规则 7: 数组读取)
                dynamic readRange = ws.Range[$"A{targetRow}:M{targetRow}"];
                object[,] existingVals = (object[,])readRange.Value2;

                // 准备写回数组
                object[,] writeBuffer = new object[1, 13];
                for (int c = 0; c < 13; c++)
                {
                    writeBuffer[0, c] = existingVals[1, c + 1];
                }

                // 覆盖 B 列名称
                writeBuffer[0, 1] = item.Name;
                // 覆盖 C 列型号
                writeBuffer[0, 2] = item.Model;
                // 覆盖 D 列品牌
                if (!string.IsNullOrWhiteSpace(item.Brand))
                {
                    writeBuffer[0, 3] = item.Brand;
                }
                // 覆盖 E 列单位
                if (!string.IsNullOrWhiteSpace(item.Unit))
                {
                    writeBuffer[0, 4] = item.Unit;
                }
                // 若指定更新数量，覆盖 F 列数量
                if (updateQuantity && item.Quantity > 0)
                {
                    writeBuffer[0, 5] = item.Quantity;
                }
                // 覆盖 I 列备注
                if (!string.IsNullOrWhiteSpace(item.Remark))
                {
                    writeBuffer[0, 8] = item.Remark;
                }
                // 覆盖 M 列单价
                writeBuffer[0, 12] = (double)item.Price;

                // 规则 7: 数组一次性写回
                readRange.Value2 = writeBuffer;

                return true;
            }
            catch (Exception ex)
            {
                // 捕获异常提示
                MessageBox.Show($"替换 Excel 物料失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
    }
}
