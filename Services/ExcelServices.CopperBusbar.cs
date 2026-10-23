using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace ExcelAddInDemo
{
    /// <summary>
    /// 母排铜排计算与物料回填相关业务服务（写入 ExcelServices 公共服务层）
    /// </summary>
    public static partial class ExcelServices
    {
        /// <summary>
        /// 唤起智能母排/铜排计算器弹窗工作台
        /// </summary>
        public static void OpenCopperBusbarCalcDialog()
        {
            try
            {
                // 获取当前活动工作簿句柄
                dynamic app = ExcelDnaUtil.Application;
                // 获取活动工作表对象
                dynamic activeSheet = app.ActiveSheet;
                // 非工作表则直接阻拦报错（规则 9）
                if (activeSheet == null)
                {
                    MessageBox.Show("未检测到活动的 Excel 工作表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 操作前先行执行规则 6 自愈检查，保障定义名称结构纯净（规则 8）
                Tool.FixAndFillCabinetNamesForSheet(activeSheet);

                // 获取 Excel 主窗口的 Win32 HWND 句柄，避免多线程模态死锁与卡顿
                IntPtr excelHwnd = ExcelDnaUtil.WindowHandle;
                // 实例化基于 WebView2 的母排计算器窗体
                using var form = new CopperBusbarCalcForm();
                // 在 Excel 主线程中以模态对话框显示
                form.ShowDialog(new ExcelWin32Window(excelHwnd));
            }
            catch (Exception ex)
            {
                // 捕获异常并弹出错误提示
                MessageBox.Show($"打开母排/铜排计算器失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 从当前活动工作表提取所有箱柜及其 W 列电流、X 列极数数据
        /// </summary>
        public static List<CopperCabinetDto> GetCabinetsForBusbarCalc()
        {
            // 初始化返回的箱柜列表
            var resultList = new List<CopperCabinetDto>();
            // 获取 Excel 宿主对象
            dynamic app = ExcelDnaUtil.Application;
            // 获取活动工作表
            dynamic ws = app.ActiveSheet;
            // 校验活动工作表有效性
            if (ws == null) return resultList;

            // 规则 8: 在读取前确保当前表格的规则 6 定义名称完整
            Tool.FixAndFillCabinetNamesForSheet(ws);

            // 提取当前表中最大使用行数
            int maxRow = 1;
            try { maxRow = ws.UsedRange.Rows.Count + ws.UsedRange.Row; } catch { maxRow = 1000; }
            if (maxRow < 20) maxRow = 500;

            // 规则 7: 一次性将 A 到 X 列数据读入内存二维数组 (X 列为第 24 列)
            dynamic fullRange = ws.Range[$"A1:X{maxRow}"];
            object[,] fullMatrix = ConvertTo2DArray(fullRange.Value2, maxRow, 24);

            // 扫描当前工作表中的所有箱柜定义名称 Cab_Det_k
            int k = 1;
            while (true)
            {
                // 尝试获取 Cab_Det_k 定义名称
                dynamic detNameObj = null;
                try { detNameObj = ws.Names.Item($"Cab_Det_{k}"); } catch { }
                if (detNameObj == null)
                {
                    // 若工作表级没有，尝试从工作簿级提取
                    try { detNameObj = ws.Parent.Names.Item($"Cab_Det_{k}"); } catch { }
                }

                // 若连续未找到则说明箱柜扫描完毕
                if (detNameObj == null) break;

                // 提取 Cab_Subsum_k 定义名称（小计行）
                dynamic subsumNameObj = null;
                try { subsumNameObj = ws.Names.Item($"Cab_Subsum_{k}"); } catch { }
                if (subsumNameObj == null)
                {
                    try { subsumNameObj = ws.Parent.Names.Item($"Cab_Subsum_{k}"); } catch { }
                }

                // 提取箱柜信息物理行号
                int detRow = 0;
                int subsumRow = 0;
                try { detRow = detNameObj.RefersToRange.Row; } catch { break; }
                try { subsumRow = subsumNameObj != null ? subsumNameObj.RefersToRange.Row : (detRow + 15); } catch { subsumRow = detRow + 15; }

                // 提取箱柜名称与柜型（位于 detRow 的 B 列）
                string cabName = fullMatrix[detRow, 2]?.ToString()?.Trim() ?? $"箱柜 {k}";
                // 提取柜型标识（若包含 GGD/MNS/GCS/GCK 则提取，否则默认 GGD）
                string cabType = "GGD"; // --硬编码: 默认柜型--
                if (cabName.IndexOf("MNS", StringComparison.OrdinalIgnoreCase) >= 0) cabType = "MNS";
                else if (cabName.IndexOf("GCS", StringComparison.OrdinalIgnoreCase) >= 0) cabType = "GCS";
                else if (cabName.IndexOf("GCK", StringComparison.OrdinalIgnoreCase) >= 0) cabType = "GCK";
                else if (cabName.IndexOf("KYN", StringComparison.OrdinalIgnoreCase) >= 0) cabType = "KYN28";

                // 规则 6: 元器件起始行为 detRow + 2，终止行为 subsumRow - 1
                int compStartRow = detRow + 2;
                int compEndRow = subsumRow - 1;

                // 主进线电流初始化
                int mainCurrent = 0;
                // 分支回路列表
                var branchList = new List<CopperBranchDto>();

                // 遍历元器件区间扫描 W 列（第 23 列）与 X 列（第 24 列）
                for (int r = compStartRow; r <= compEndRow && r <= maxRow; r++)
                {
                    // 提取 B 列元器件名称与 C 列规格型号
                    string compName = fullMatrix[r, 2]?.ToString()?.Trim() ?? "";
                    string compModel = fullMatrix[r, 3]?.ToString()?.Trim() ?? "";

                    // 提取 W 列（电流）文本与数值
                    object wVal = fullMatrix[r, 23];
                    string wText = wVal?.ToString()?.Trim() ?? "";
                    // 提取 X 列（极数）文本
                    object xVal = fullMatrix[r, 24];
                    string xText = xVal?.ToString()?.Trim() ?? "3P"; // --硬编码: 默认3P极数--

                    // 正则提取 W 列中的纯数字作为额定电流
                    int amp = ParseAmpere(wText);

                    // 若识别到有效电流值
                    if (amp > 0)
                    {
                        // 取当前箱柜所有电流中的最大值作为主进线开关电流
                        if (amp > mainCurrent) mainCurrent = amp;

                        // 将该回路收录为分支引下排
                        branchList.Add(new CopperBranchDto
                        {
                            CircuitName = !string.IsNullOrEmpty(compName) ? $"{compName} ({compModel})" : $"回路 {branchList.Count + 1}",
                            Current = amp,
                            Poles = xText,
                            RouteType = (cabType == "MNS" || cabType == "GCS") ? "drawer" : "top_down",
                            BendType = "L",
                            Spec = RecommendBranchSpec(amp),
                            Pieces = 1,
                            UnfoldedMm = 800, // --硬编码: 默认分支展开长--
                            Qty = ParsePolesQty(xText)
                        });
                    }
                }

                // 若未能从 W 列识别出电流，给予常规标准进线电流兜底
                if (mainCurrent <= 0) mainCurrent = 630; // --硬编码: 兜底主开关电流--

                // 推荐主母排规格与拼数
                var (recSpec, recPieces) = RecommendMainSpecAndPieces(mainCurrent);

                // 创建该箱柜传输模型
                var cabDto = new CopperCabinetDto
                {
                    Id = $"cab_{k}",
                    No = k.ToString(),
                    Name = cabName,
                    Type = cabType,
                    Width = (mainCurrent >= 2000) ? 1000 : 800, // --硬编码: 大电流柜宽1000小电流800--
                    MainCurrent = mainCurrent,
                    MainBusbars = new List<CopperMainBusbarDto>
                    {
                        new CopperMainBusbarDto { Phase = "A", Current = mainCurrent, RecommendedSpec = recSpec, Spec = recSpec, Pieces = recPieces, LengthMm = (mainCurrent >= 2000 ? 1000 : 800), LapMm = 150 },
                        new CopperMainBusbarDto { Phase = "B", Current = mainCurrent, RecommendedSpec = recSpec, Spec = recSpec, Pieces = recPieces, LengthMm = (mainCurrent >= 2000 ? 1000 : 800), LapMm = 150 },
                        new CopperMainBusbarDto { Phase = "C", Current = mainCurrent, RecommendedSpec = recSpec, Spec = recSpec, Pieces = recPieces, LengthMm = (mainCurrent >= 2000 ? 1000 : 800), LapMm = 150 },
                        new CopperMainBusbarDto { Phase = "N", Current = mainCurrent / 2, RecommendedSpec = RecommendHalfSpec(recSpec), Spec = RecommendHalfSpec(recSpec), Pieces = 1, LengthMm = (mainCurrent >= 2000 ? 1000 : 800), LapMm = 150 },
                        new CopperMainBusbarDto { Phase = "PE", Current = mainCurrent / 2, RecommendedSpec = "TMY-60x6", Spec = "TMY-60x6", Pieces = 1, LengthMm = (mainCurrent >= 2000 ? 1000 : 800), LapMm = 100 }
                    },
                    BranchBusbars = branchList
                };

                // 加入结果集合
                resultList.Add(cabDto);
                // 递增索引继续扫描下一个箱柜
                k++;
            }

            // 返回提取出的箱柜列表
            return resultList;
        }

        /// <summary>
        /// 将计算出的铜排物料明细行按规则 6 安全插入到 Excel 各箱柜元器件区中
        /// </summary>
        public static bool InsertCopperBusbarBomToExcel(CopperWriteBackPayload payload)
        {
            // 校验入参有效性
            if (payload == null || payload.Cabinets == null || payload.Cabinets.Count == 0) return false;

            // 获取 Excel 宿主句柄
            dynamic app = ExcelDnaUtil.Application;
            // 获取活动工作表
            dynamic ws = app.ActiveSheet;
            // 校验工作表有效性
            if (ws == null) return false;

            // 规则 8: 操作前执行规则 6 架构自愈
            Tool.FixAndFillCabinetNamesForSheet(ws);

            // 判定当前是按重量计价模式还是按长度计价模式
            bool isWeightMode = string.Equals(payload.UnitMode, "weight", StringComparison.OrdinalIgnoreCase);
            // 提取综合单价 (铜价 + 加工费)
            double copperPricePerKg = payload.CopperUnitPrice > 0 ? payload.CopperUnitPrice : 82.50; // --硬编码: 默认单价--

            // 暂停屏幕刷新加速写入 (操作完成后恢复)
            app.ScreenUpdating = false;
            try
            {
                // 倒序处理各箱柜，避免上方插行导致下方箱柜物理行号漂移
                for (int k = payload.Cabinets.Count; k >= 1; k--)
                {
                    // 获取当前箱柜回填数据
                    var cab = payload.Cabinets[k - 1];

                    // 获取当前箱柜的 Cab_Det 与 Cab_Subsum 定义名称
                    dynamic detNameObj = null;
                    dynamic subsumNameObj = null;
                    try { detNameObj = ws.Names.Item($"Cab_Det_{k}"); } catch { }
                    if (detNameObj == null) { try { detNameObj = ws.Parent.Names.Item($"Cab_Det_{k}"); } catch { } }
                    try { subsumNameObj = ws.Names.Item($"Cab_Subsum_{k}"); } catch { }
                    if (subsumNameObj == null) { try { subsumNameObj = ws.Parent.Names.Item($"Cab_Subsum_{k}"); } catch { } }

                    // 若未找到箱柜标签则跳过该柜
                    if (detNameObj == null || subsumNameObj == null) continue;

                    // 获取真实的明细行号与小计行号
                    int detRow = detNameObj.RefersToRange.Row;
                    int subsumRow = subsumNameObj.RefersToRange.Row;

                    // 规则 6: 元器件起始行为 detRow + 2，终止行为 subsumRow - 1
                    int compStartRow = detRow + 2;
                    int compEndRow = subsumRow - 1;

                    // 扫描当前元器件区间寻找最后一个非空行 (规则 7 内存扫描 B 列与 C 列)
                    int totalCompRows = Math.Max(0, compEndRow - compStartRow + 1);
                    int lastUsedIndex = 0;
                    if (totalCompRows > 0)
                    {
                        dynamic checkRange = ws.Range[$"B{compStartRow}:C{compEndRow}"];
                        object[,] checkMatrix = ConvertTo2DArray(checkRange.Value2, totalCompRows, 2);
                        for (int r = 1; r <= totalCompRows; r++)
                        {
                            string bText = checkMatrix[r, 1]?.ToString()?.Trim() ?? "";
                            string cText = checkMatrix[r, 2]?.ToString()?.Trim() ?? "";
                            if (!string.IsNullOrEmpty(bText) || !string.IsNullOrEmpty(cText))
                            {
                                lastUsedIndex = r;
                            }
                        }
                    }

                    // 计算最后一个非空元器件行的物理行号
                    int lastUsedRow = (lastUsedIndex > 0) ? (compStartRow + lastUsedIndex - 1) : (compStartRow - 1);
                    // 计算当前剩余可用空行数
                    int availableEmptyRows = Math.Max(0, compEndRow - lastUsedRow);

                    // 构建待写入当前箱柜的铜排物料条目列表 (水平主排、垂直分支排、地排)
                    var insertItems = BuildCopperBomItemsForCabinet(cab, isWeightMode, copperPricePerKg);
                    int reqCount = insertItems.Count;
                    if (reqCount == 0) continue;

                    // 规则 6: 检查空行是否充足，若不足则在小计行上方执行安全插行
                    if (reqCount > availableEmptyRows)
                    {
                        int rowsToInsert = reqCount - availableEmptyRows;
                        // 在 Cab_Subsum 行上方执行向下插入 (xlShiftDown = -4121)
                        dynamic insertRange = ws.Range[$"A{subsumRow}:A{subsumRow + rowsToInsert - 1}"];
                        insertRange.EntireRow.Insert(-4121);

                        // 小计行下移
                        subsumRow += rowsToInsert;
                        compEndRow = subsumRow - 1;

                        // 维护真正小计行的求和公式（保持自适应）
                        ws.Cells[subsumRow, 8].Formula = $"=ROUND(SUM(H{compStartRow}:INDEX(H:H,ROW()-1)),2)";
                    }

                    // 计算写入起始物理行号
                    int writeStartRow = lastUsedRow + 1;

                    // 构建 B 到 N 列纯数据矩阵 (共 13 列，规则 7 内存批量构建，防型号 COM 崩溃)
                    object[,] dataMatrix = new object[reqCount, 13];
                    // 遍历插入物料构建纯数据与常数字段
                    for (int i = 0; i < reqCount; i++)
                    {
                        var item = insertItems[i];
                        // 索引 0 对应 B 列: 名称 (矩形铜母线)
                        dataMatrix[i, 0] = item.Name;
                        // 索引 1 对应 C 列: 型号 (如 TMY-2*(120x10)，纯文本写入)
                        dataMatrix[i, 1] = item.Model;
                        // 索引 2 对应 D 列: 生产厂家 (国优) --硬编码: 生产厂家--
                        dataMatrix[i, 2] = "国优";
                        // 索引 3 对应 E 列: 计量单位 (kg 或 米)
                        dataMatrix[i, 3] = item.Unit;
                        // 索引 4 对应 F 列: 数量 (重量 或 长度)
                        dataMatrix[i, 4] = item.Qty;
                        // 索引 5 对应 G 列: 预留销售单价列 (通过独立 Formula 单独写入)
                        dataMatrix[i, 5] = DBNull.Value;
                        // 索引 6 对应 H 列: 预留销售总价列 (通过独立 Formula 单独写入)
                        dataMatrix[i, 6] = DBNull.Value;
                        // 索引 7 对应 I 列: 备注 (水平主排/分支引下/接地PE)
                        dataMatrix[i, 7] = item.Remark;
                        // 索引 8 对应 J 列: 预留成本单价列 (通过独立 Formula 单独写入)
                        dataMatrix[i, 8] = DBNull.Value;
                        // 索引 9 对应 K 列: 预留成本总价列 (通过独立 Formula 单独写入)
                        dataMatrix[i, 9] = DBNull.Value;
                        // 索引 10 对应 L 列: 报出系数 (默认 1.0) --硬编码: 默认加价系数--
                        dataMatrix[i, 10] = 1.0;
                        // 索引 11 对应 M 列: 表价/单价数值 (按重量为单价，按长度为单价数值)
                        dataMatrix[i, 11] = item.PriceFormulaOrVal;
                        // 索引 12 对应 N 列: 采购折扣 (默认 1.0) --硬编码: 默认折扣系数--
                        dataMatrix[i, 12] = 1.0;
                    }

                    // 规则 7: 纯数据矩阵一次性写回 B 到 N 列
                    dynamic destDataRange = ws.Range[$"B{writeStartRow}:N{writeStartRow + reqCount - 1}"];
                    destDataRange.Value2 = dataMatrix;

                    // 准备 A、G、H、J、K 五列独立标准联动公式矩阵
                    object[,] formulasA = new object[reqCount, 1];
                    // 销售单价公式矩阵
                    object[,] formulasG = new object[reqCount, 1];
                    // 销售总价公式矩阵
                    object[,] formulasH = new object[reqCount, 1];
                    // 成本单价公式矩阵
                    object[,] formulasJ = new object[reqCount, 1];
                    // 成本总价公式矩阵
                    object[,] formulasK = new object[reqCount, 1];

                    // 目标箱柜明细表头行号 (detRow + 1)
                    int compHeaderRow = detRow + 1;
                    for (int i = 0; i < reqCount; i++)
                    {
                        int curRow = writeStartRow + i;
                        // A 列自适应序号公式绑定当前箱柜明细表头行 --硬编码: 序号公式模板--
                        formulasA[i, 0] = $"=ROW()-ROW(A${compHeaderRow})";
                        // G 列销售单价标准联动公式 =ROUND(M*L*N, 2) --硬编码: 销售单价公式--
                        formulasG[i, 0] = $"=IF(AND(B{curRow}=\"\",C{curRow}=\"\"),\"\",ROUND(M{curRow}*L{curRow}*N{curRow},2))";
                        // H 列销售总价标准联动公式 =ROUND(F*G, 2) --硬编码: 销售总价公式--
                        formulasH[i, 0] = $"=IF(AND(B{curRow}=\"\",C{curRow}=\"\"),\"\",ROUND(F{curRow}*G{curRow},2))";
                        // J 列成本单价标准联动公式 =ROUND(M*N, 2) --硬编码: 成本单价公式--
                        formulasJ[i, 0] = $"=IF(AND(B{curRow}=\"\",C{curRow}=\"\"),\"\",ROUND(M{curRow}*N{curRow},2))";
                        // K 列成本总价标准联动公式 =ROUND(J*F, 2) --硬编码: 成本总价公式--
                        formulasK[i, 0] = $"=IF(AND(B{curRow}=\"\",C{curRow}=\"\"),\"\",ROUND(J{curRow}*F{curRow},2))";
                    }

                    // 规则 7: 单列单次 COM 批量写入公式，0 型号词法冲突
                    ws.Range[$"A{writeStartRow}:A{writeStartRow + reqCount - 1}"].Formula = formulasA;
                    // G 列单价公式写入
                    ws.Range[$"G{writeStartRow}:G{writeStartRow + reqCount - 1}"].Formula = formulasG;
                    // H 列总价公式写入
                    ws.Range[$"H{writeStartRow}:H{writeStartRow + reqCount - 1}"].Formula = formulasH;
                    // J 列成本单价公式写入
                    ws.Range[$"J{writeStartRow}:J{writeStartRow + reqCount - 1}"].Formula = formulasJ;
                    // K 列成本总价公式写入
                    ws.Range[$"K{writeStartRow}:K{writeStartRow + reqCount - 1}"].Formula = formulasK;
                }

                // 写入全部完毕后，执行全局规则 6 自愈刷新，校准定义名称拓扑与超链接
                Tool.FixAndFillCabinetNamesForSheet(ws);
                return true;
            }
            finally
            {
                // 恢复屏幕刷新
                app.ScreenUpdating = true;
            }
        }

        /// <summary>
        /// 构建待插入单个箱柜的铜排物料条目列表
        /// </summary>
        private static List<CopperInsertItem> BuildCopperBomItemsForCabinet(CopperCabinetDto cab, bool isWeightMode, double copperPrice)
        {
            // 初始化插入明细列表
            var list = new List<CopperInsertItem>();
            // 空对象安全防护
            if (cab == null) return list;

            // 1. 水平主母排 (汇总 A/B/C/N 相，独立于 PE 排)
            if (cab.MainBusbars != null && cab.MainBusbars.Count > 0)
            {
                // 提取 A 相作为代表主排规格
                var aRow = cab.MainBusbars[0];
                // 累计水平主排总重量与总长度
                double mainTotalWeight = 0;
                double mainTotalLength = 0;
                // 遍历母排列表汇总
                foreach (var mb in cab.MainBusbars)
                {
                    // PE 排留作第 3 项接地铜母线单独处理
                    if (string.Equals(mb.Phase, "PE", StringComparison.OrdinalIgnoreCase)) continue;
                    // 计算单项总长度 (含搭接裕量)
                    double l = (mb.LengthMm + mb.LapMm) / 1000.0 * mb.Pieces;
                    // 累加长度
                    mainTotalLength += l;
                    // 累加重量
                    mainTotalWeight += l * mb.UnitWeight;
                }

                // 计算水平主母排综合单米理论重量
                double mainUnitWeight = mainTotalLength > 0 ? Math.Round(mainTotalWeight / mainTotalLength, 3) : aRow.UnitWeight;
                // 格式化型号：多拼则输出如 TMY-2*(120x10)，单拼输出如 TMY-120x10
                string mainModel = FormatBusbarModel(aRow.Spec, aRow.Pieces);

                // 添加水平主母排物料行
                list.Add(new CopperInsertItem
                {
                    Name = "矩形铜母线", // B 列名称
                    Model = mainModel, // C 列型号: 如 TMY-2*(120x10)
                    Unit = isWeightMode ? "kg" : "米", // E 列单位
                    Qty = isWeightMode ? Math.Round(mainTotalWeight, 2) : Math.Round(mainTotalLength, 2), // F 列数量: 重量或长度
                    Remark = "矩形铜母线 (水平主排)", // I 列备注
                    // M 列单价: 按重量为纯单价数值，按长度为公式 =每米重量*铜价
                    PriceFormulaOrVal = isWeightMode ? (object)copperPrice : $"={mainUnitWeight}*{copperPrice}"
                });
            }

            // 2. 垂直分支引下母排 (按回路汇总)
            if (cab.BranchBusbars != null && cab.BranchBusbars.Count > 0)
            {
                // 分支累计重量与长度初始化
                double branchTotalWeight = 0;
                double branchTotalLength = 0;
                // 提取首个回路规格与拼数作为代表
                string branchSpec = cab.BranchBusbars[0].Spec;
                int branchPieces = cab.BranchBusbars[0].Pieces;

                // 循环统计各回路引下线用量
                foreach (var br in cab.BranchBusbars)
                {
                    // 计算展开总米数
                    double l = (br.UnfoldedMm / 1000.0) * br.Pieces * br.Qty;
                    // 累加长度
                    branchTotalLength += l;
                    // 累加重量 (按米重 1.42 兜底或标准米重计算)
                    branchTotalWeight += l * 1.42; // --硬编码: 默认分支米重--
                }

                // 若存在分支排用量
                if (branchTotalWeight > 0 || branchTotalLength > 0)
                {
                    // 计算分支排综合单米理论重量
                    double branchUnitWeight = branchTotalLength > 0 ? Math.Round(branchTotalWeight / branchTotalLength, 3) : 1.42;
                    // 格式化型号
                    string branchModel = FormatBusbarModel(branchSpec, branchPieces);

                    // 添加分支引下母排物料行
                    list.Add(new CopperInsertItem
                    {
                        Name = "矩形铜母线", // B 列名称
                        Model = branchModel, // C 列型号
                        Unit = isWeightMode ? "kg" : "米", // E 列单位
                        Qty = isWeightMode ? Math.Round(branchTotalWeight, 2) : Math.Round(branchTotalLength, 2), // F 列数量: 重量或长度
                        Remark = "矩形铜母线 (分支引下)", // I 列备注
                        // M 列单价: 按重量为纯单价，按长度为公式 =每米重量*铜价
                        PriceFormulaOrVal = isWeightMode ? (object)copperPrice : $"={branchUnitWeight}*{copperPrice}"
                    });
                }
            }

            // 3. 接地铜母线 PE
            // 尝试从箱柜母排中获取 PE 相规格与长度
            var peRow = cab.MainBusbars?.Find(b => string.Equals(b.Phase, "PE", StringComparison.OrdinalIgnoreCase));
            // 提取 PE 规格 (默认 TMY-40x4)
            string peSpec = peRow?.Spec ?? "TMY-40x4"; // --硬编码: 默认接地排型号--
            // 提取 PE 长度 (默认柜宽 1.0米)
            double peLength = peRow != null ? (peRow.LengthMm + peRow.LapMm) / 1000.0 : 1.0;
            // 提取 PE 米重 (TMY-40x4 约为 1.42 kg/m)
            double peUnitWeight = peRow != null ? peRow.UnitWeight : 1.42;
            // 计算 PE 总重量
            double peTotalWeight = peLength * peUnitWeight;

            // 添加接地母排物料行
            list.Add(new CopperInsertItem
            {
                Name = "接地铜母线", // B 列名称
                Model = peSpec, // C 列型号: 如 TMY-40x4
                Unit = isWeightMode ? "kg" : "米", // E 列单位
                Qty = isWeightMode ? Math.Round(peTotalWeight, 2) : Math.Round(peLength, 2), // F 列数量: 重量或长度
                Remark = "接地铜母线 PE", // I 列备注
                // M 列单价: 按重量为纯单价，按长度为公式 =每米重量*铜价
                PriceFormulaOrVal = isWeightMode ? (object)copperPrice : $"={peUnitWeight}*{copperPrice}"
            });

            // 返回构建好的物料清单
            return list;
        }

        /// <summary>
        /// 格式化母排型号字符串 (如 TMY-120x10, 2拼 -> TMY-2*(120x10))
        /// </summary>
        private static string FormatBusbarModel(string spec, int pieces)
        {
            // 空值保护兜底
            if (string.IsNullOrWhiteSpace(spec)) return "TMY-40x4";
            // 单拼保持原样
            if (pieces <= 1) return spec;
            // 剥离 TMY- 前缀提取纯截面规格
            string coreDim = spec.Replace("TMY-", "").Trim();
            // 组装符合国标图纸标准的拼数表达式，如 TMY-2*(120x10)
            return $"TMY-{pieces}*({coreDim})";
        }

        /// <summary>
        /// 从字符串解析安培数 (如 2000A 解析为 2000)
        /// </summary>
        private static int ParseAmpere(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            var match = Regex.Match(text, @"\d+");
            if (match.Success && int.TryParse(match.Value, out int val))
            {
                return val;
            }
            return 0;
        }

        /// <summary>
        /// 解析极数对应的根数
        /// </summary>
        private static int ParsePolesQty(string poles)
        {
            if (string.IsNullOrWhiteSpace(poles)) return 3;
            if (poles.IndexOf("4", StringComparison.OrdinalIgnoreCase) >= 0) return 4;
            if (poles.IndexOf("1", StringComparison.OrdinalIgnoreCase) >= 0) return 1;
            if (poles.IndexOf("2", StringComparison.OrdinalIgnoreCase) >= 0) return 2;
            return 3;
        }

        /// <summary>
        /// 根据主进线电流推荐主母排规格与拼数
        /// </summary>
        private static (string Spec, int Pieces) RecommendMainSpecAndPieces(int current)
        {
            if (current >= 3200) return ("TMY-120x10", 3);
            if (current >= 2000) return ("TMY-120x10", 2);
            if (current >= 1600) return ("TMY-100x10", 2);
            if (current >= 1250) return ("TMY-80x8", 2);
            if (current >= 800) return ("TMY-80x8", 1);
            if (current >= 630) return ("TMY-60x6", 1);
            return ("TMY-50x5", 1);
        }

        /// <summary>
        /// 推荐中性排半规格
        /// </summary>
        private static string RecommendHalfSpec(string mainSpec)
        {
            if (mainSpec.Contains("120x10")) return "TMY-100x10";
            if (mainSpec.Contains("100x10")) return "TMY-80x8";
            if (mainSpec.Contains("80x8")) return "TMY-60x6";
            return "TMY-40x4";
        }

        /// <summary>
        /// 根据分支电流推荐分支铜排规格
        /// </summary>
        private static string RecommendBranchSpec(int current)
        {
            if (current >= 630) return "TMY-50x5";
            if (current >= 400) return "TMY-40x4";
            if (current >= 250) return "TMY-30x4";
            return "TMY-30x4";
        }
    }

    #region 数据传输 DTO 模型定义

    /// <summary>
    /// 箱柜传输模型
    /// </summary>
    public class CopperCabinetDto
    {
        public string Id { get; set; } = "";
        public string No { get; set; } = "";
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Width { get; set; } = 800;
        public int MainCurrent { get; set; } = 630;
        public double TotalWeight { get; set; } = 0;
        public List<CopperMainBusbarDto> MainBusbars { get; set; } = new List<CopperMainBusbarDto>();
        public List<CopperBranchDto> BranchBusbars { get; set; } = new List<CopperBranchDto>();
    }

    /// <summary>
    /// 水平主母排 DTO
    /// </summary>
    public class CopperMainBusbarDto
    {
        public string Phase { get; set; } = "A";
        public int Current { get; set; } = 630;
        public string RecommendedSpec { get; set; } = "TMY-60x6";
        public string Spec { get; set; } = "TMY-60x6";
        public int Pieces { get; set; } = 1;
        public int LengthMm { get; set; } = 800;
        public int LapMm { get; set; } = 150;
        public double UnitWeight { get; set; } = 3.20;
        public double TotalWeight { get; set; } = 0;
    }

    /// <summary>
    /// 分支引下排 DTO
    /// </summary>
    public class CopperBranchDto
    {
        public string CircuitName { get; set; } = "";
        public int Current { get; set; } = 250;
        public string Poles { get; set; } = "3P";
        public string RouteType { get; set; } = "drawer";
        public string BendType { get; set; } = "L";
        public string Spec { get; set; } = "TMY-30x4";
        public int Pieces { get; set; } = 1;
        public int UnfoldedMm { get; set; } = 800;
        public int Qty { get; set; } = 3;
        public double TotalWeight { get; set; } = 0;
    }

    /// <summary>
    /// 回填载荷 DTO
    /// </summary>
    public class CopperWriteBackPayload
    {
        public string Mode { get; set; } = "detailBOM";
        public string UnitMode { get; set; } = "weight"; // "weight" 或 "length"
        public double TotalWeight { get; set; } = 0;
        public double TotalLength { get; set; } = 0;
        public double TotalCost { get; set; } = 0;
        public double CopperUnitPrice { get; set; } = 82.50;
        public List<CopperCabinetDto> Cabinets { get; set; } = new List<CopperCabinetDto>();
    }

    /// <summary>
    /// 插入物料项结构
    /// </summary>
    public class CopperInsertItem
    {
        public string Name { get; set; } = "矩形铜母线";
        public string Model { get; set; } = "";
        public string Unit { get; set; } = "kg";
        public double Qty { get; set; } = 0;
        public string Remark { get; set; } = "";
        public object PriceFormulaOrVal { get; set; } = 0;
    }

    #endregion
}
