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

                // 5. 规则 7: 纯数据矩阵与公式列解耦安全写入，杜绝 put_Formula 对含特殊字符型号解析崩溃
                // 覆盖 B 列到 N 列 (共 13 列纯数据字段)
                object[,] dataMatrix = new object[count, 13];
                // 遍历物料列表填充纯数据矩阵 (含名称、型号、厂家、单位、数量、备注、系数、单价、折扣)
                for (int i = 0; i < count; i++)
                {
                    // 获取当前物料 DTO 对象
                    var curItem = items[i];
                    // 索引 0 对应 B 列: 物料名称
                    dataMatrix[i, 0] = curItem.Name;
                    // 索引 1 对应 C 列: 型号规格 (纯文本写入，杜绝 COM 解析 -/+ 异常)
                    dataMatrix[i, 1] = curItem.Model;
                    // 索引 2 对应 D 列: 品牌生产厂家 (若为空默认国优) --硬编码: 默认厂家--
                    dataMatrix[i, 2] = string.IsNullOrWhiteSpace(curItem.Brand) ? "国优" : curItem.Brand;
                    // 索引 3 对应 E 列: 计量单位 (若为空默认米) --硬编码: 默认计量单位--
                    dataMatrix[i, 3] = string.IsNullOrWhiteSpace(curItem.Unit) ? "米" : curItem.Unit;
                    // 索引 4 对应 F 列: 数量 (长度或重量)
                    dataMatrix[i, 4] = curItem.Quantity > 0 ? curItem.Quantity : 1.0;
                    // 索引 5 对应 G 列: 预留销售单价列 (通过独立 Formula 单独写入)
                    dataMatrix[i, 5] = DBNull.Value;
                    // 索引 6 对应 H 列: 预留销售总价列 (通过独立 Formula 单独写入)
                    dataMatrix[i, 6] = DBNull.Value;
                    // 索引 7 对应 I 列: 备注说明
                    dataMatrix[i, 7] = curItem.Remark ?? string.Empty;
                    // 索引 8 对应 J 列: 预留成本单价列 (通过独立 Formula 单独写入)
                    dataMatrix[i, 8] = DBNull.Value;
                    // 索引 9 对应 K 列: 预留成本总价列 (通过独立 Formula 单独写入)
                    dataMatrix[i, 9] = DBNull.Value;
                    // 索引 10 对应 L 列: 报出/加价系数 (默认 1.0) --硬编码: 默认加价系数--
                    dataMatrix[i, 10] = 1.0;
                    // 索引 11 对应 M 列: 面价/基准单价数值
                    dataMatrix[i, 11] = (double)curItem.Price;
                    // 索引 12 对应 N 列: 采购折扣系数 (默认 1.0) --硬编码: 默认折扣系数--
                    dataMatrix[i, 12] = 1.0;
                }

                // 规则 7: 一次性将 B 到 N 列纯数据批量写入工作表
                dynamic dataRange = ws.Range[$"B{insertRow}:N{insertRow + count - 1}"];
                // 批量内存赋值
                dataRange.Value2 = dataMatrix;

                // 准备 A、G、H、J、K 五列的独立标准公式矩阵 (规则 7 单列整块写入)
                object[,] formulasA = new object[count, 1];
                // 销售单价公式矩阵
                object[,] formulasG = new object[count, 1];
                // 销售总价公式矩阵
                object[,] formulasH = new object[count, 1];
                // 成本单价公式矩阵
                object[,] formulasJ = new object[count, 1];
                // 成本总价公式矩阵
                object[,] formulasK = new object[count, 1];

                // 目标箱柜明细表头行号 (targetDetRow + 1)
                int compHeaderRow = targetDetRow + 1;
                // 循环组装各行标准成套计算公式
                for (int i = 0; i < count; i++)
                {
                    // 当前实际物理行号
                    int curR = insertRow + i;
                    // A 列: 自适应序号公式 --硬编码: 动态序号公式--
                    formulasA[i, 0] = $"=ROW()-ROW(A${compHeaderRow})";
                    // G 列: 销售单价标准联动公式 =ROUND(M*L*N, 2) --硬编码: 销售单价公式--
                    formulasG[i, 0] = $"=IF(AND(B{curR}=\"\",C{curR}=\"\"),\"\",ROUND(M{curR}*L{curR}*N{curR},2))";
                    // H 列: 销售总价标准联动公式 =ROUND(F*G, 2) --硬编码: 销售总价公式--
                    formulasH[i, 0] = $"=IF(AND(B{curR}=\"\",C{curR}=\"\"),\"\",ROUND(F{curR}*G{curR},2))";
                    // J 列: 成本单价标准联动公式 =ROUND(M*N, 2) --硬编码: 成本单价公式--
                    formulasJ[i, 0] = $"=IF(AND(B{curR}=\"\",C{curR}=\"\"),\"\",ROUND(M{curR}*N{curR},2))";
                    // K 列: 成本总价标准联动公式 =ROUND(J*F, 2) --硬编码: 成本总价公式--
                    formulasK[i, 0] = $"=IF(AND(B{curR}=\"\",C{curR}=\"\"),\"\",ROUND(J{curR}*F{curR},2))";
                }

                // 规则 7: 分别按单列 Range 单次 COM 批量写入公式，0 型号词法冲突
                ws.Range[$"A{insertRow}:A{insertRow + count - 1}"].Formula = formulasA;
                // 写入 G 列销售单价公式
                ws.Range[$"G{insertRow}:G{insertRow + count - 1}"].Formula = formulasG;
                // 写入 H 列销售总价公式
                ws.Range[$"H{insertRow}:H{insertRow + count - 1}"].Formula = formulasH;
                // 写入 J 列成本单价公式
                ws.Range[$"J{insertRow}:J{insertRow + count - 1}"].Formula = formulasJ;
                // 写入 K 列成本总价公式
                ws.Range[$"K{insertRow}:K{insertRow + count - 1}"].Formula = formulasK;

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
        /// 替换 Excel 当前光标选中行物料数据 (用户决策 A2: 直接覆盖该行的名称、型号、单价等字段，并确保公式完备)
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

                // 读取当前行的数量与原有内容 (规则 7: 覆盖 A 到 N 列，共 14 列)
                dynamic readRange = ws.Range[$"A{targetRow}:N{targetRow}"];
                object[,] existingVals = (object[,])readRange.Value2;

                // 准备写回数组 (B 到 N 列共 13 列纯数据)
                object[,] writeBuffer = new object[1, 13];
                for (int c = 0; c < 13; c++)
                {
                    // 复制现有单元格数值
                    writeBuffer[0, c] = existingVals[1, c + 2];
                }

                // 覆盖 B 列名称
                writeBuffer[0, 0] = item.Name;
                // 覆盖 C 列型号 (纯文本赋值，防 COM 解析崩溃)
                writeBuffer[0, 1] = item.Model;
                // 覆盖 D 列品牌 (若非空则更新)
                if (!string.IsNullOrWhiteSpace(item.Brand))
                {
                    writeBuffer[0, 2] = item.Brand;
                }
                // 覆盖 E 列单位 (若非空则更新)
                if (!string.IsNullOrWhiteSpace(item.Unit))
                {
                    writeBuffer[0, 3] = item.Unit;
                }
                // 若指定更新数量，覆盖 F 列数量
                if (updateQuantity && item.Quantity > 0)
                {
                    writeBuffer[0, 4] = item.Quantity;
                }
                // 覆盖 I 列备注
                if (!string.IsNullOrWhiteSpace(item.Remark))
                {
                    writeBuffer[0, 7] = item.Remark;
                }
                // 检查补齐 L 列报出系数 (若为空或非有效数值则初始化为 1.0)
                object oldL = writeBuffer[0, 10];
                if (oldL == null || !double.TryParse(oldL.ToString(), out double lNum) || lNum <= 0)
                {
                    // 赋默认报出系数 --硬编码: 默认系数--
                    writeBuffer[0, 10] = 1.0;
                }
                // 覆盖 M 列单价数值
                writeBuffer[0, 11] = (double)item.Price;
                // 检查补齐 N 列采购折扣 (若为空或非有效数值则初始化为 1.0)
                object oldN = writeBuffer[0, 12];
                if (oldN == null || !double.TryParse(oldN.ToString(), out double nNum) || nNum <= 0)
                {
                    // 赋默认采购折扣 --硬编码: 默认折扣--
                    writeBuffer[0, 12] = 1.0;
                }

                // 规则 7: 纯数据矩阵一次性写回 B 到 N 列
                dynamic writeRange = ws.Range[$"B{targetRow}:N{targetRow}"];
                writeRange.Value2 = writeBuffer;

                // 补齐 G 列销售单价标准联动公式 --硬编码: 销售单价公式--
                ws.Cells[targetRow, 7].Formula = $"=IF(AND(B{targetRow}=\"\",C{targetRow}=\"\"),\"\",ROUND(M{targetRow}*L{targetRow}*N{targetRow},2))";
                // 补齐 H 列销售总价标准联动公式 --硬编码: 销售总价公式--
                ws.Cells[targetRow, 8].Formula = $"=IF(AND(B{targetRow}=\"\",C{targetRow}=\"\"),\"\",ROUND(F{targetRow}*G{targetRow},2))";
                // 补齐 J 列成本单价标准联动公式 --硬编码: 成本单价公式--
                ws.Cells[targetRow, 10].Formula = $"=IF(AND(B{targetRow}=\"\",C{targetRow}=\"\"),\"\",ROUND(M{targetRow}*N{targetRow},2))";
                // 补齐 K 列成本总价标准联动公式 --硬编码: 成本总价公式--
                ws.Cells[targetRow, 11].Formula = $"=IF(AND(B{targetRow}=\"\",C{targetRow}=\"\"),\"\",ROUND(J{targetRow}*F{targetRow},2))";

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
