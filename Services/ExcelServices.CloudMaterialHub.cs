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
        /// 将物料插入到当前 Excel 分类表的元器件区域 (插入到当前光标所在行的下方)
        /// 严格遵循规则 6、7、8、9
        /// </summary>
        /// <param name="item">待插入的物料条目</param>
        /// <returns>操作是否成功</returns>
        public static bool InsertMaterialFromHub(HubMaterialDto item)
        {
            try
            {
                // 获取 Excel 宿主对象
                dynamic app = ExcelDnaUtil.Application;
                dynamic ws = app.ActiveSheet;
                // 规则 9: 检查工作表有效性
                if (ws == null)
                {
                    MessageBox.Show("当前没有激活的 Excel 工作表！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // 规则 8: 校验并纠正规则 6 定义名称架构
                Tool.FixAndFillCabinetNamesForSheet(ws);

                // 获取当前活动单元格
                dynamic activeCell = app.ActiveCell;
                int cursorRow = activeCell != null ? (int)activeCell.Row : 0;

                // 扫描当前工作表中的所有箱柜定义名称 Cab_Det_k 和 Cab_Subsum_k
                int targetCabinetK = -1;
                int targetDetRow = -1;
                int targetSubsumRow = -1;
                int firstValidDetRow = -1;
                int firstValidSubsumRow = -1;

                // 循环遍历工作表定义名称
                for (int k = 1; k <= 500; k++)
                {
                    // 尝试提取 Cab_Det_k
                    dynamic detNameObj = null;
                    try { detNameObj = ws.Names.Item($"Cab_Det_{k}"); } catch { }
                    if (detNameObj == null)
                    {
                        try { detNameObj = ws.Parent.Names.Item($"Cab_Det_{k}"); } catch { }
                    }
                    if (detNameObj == null) continue;

                    // 尝试提取 Cab_Subsum_k
                    dynamic subsumNameObj = null;
                    try { subsumNameObj = ws.Names.Item($"Cab_Subsum_{k}"); } catch { }
                    if (subsumNameObj == null)
                    {
                        try { subsumNameObj = ws.Parent.Names.Item($"Cab_Subsum_{k}"); } catch { }
                    }
                    if (subsumNameObj == null) continue;

                    // 获取当前箱柜的明细行与小计行
                    int detR = (int)detNameObj.RefersToRange.Row;
                    int subR = (int)subsumNameObj.RefersToRange.Row;

                    // 记录首个有效箱柜作为兜底
                    if (firstValidDetRow <= 0)
                    {
                        firstValidDetRow = detR;
                        firstValidSubsumRow = subR;
                        targetCabinetK = k;
                    }

                    // 判定光标是否落入此箱柜范围 (包括从 Det 行至 Subsum 行)
                    if (cursorRow >= detR && cursorRow <= subR)
                    {
                        targetCabinetK = k;
                        targetDetRow = detR;
                        targetSubsumRow = subR;
                        break;
                    }
                }

                // 若未命中当前光标所属箱柜，回退至首个有效箱柜
                if (targetDetRow <= 0)
                {
                    targetDetRow = firstValidDetRow;
                    targetSubsumRow = firstValidSubsumRow;
                }

                // 规则 9: 检查是否具有有效明细区，否则直接阻拦报错
                if (targetDetRow <= 0 || targetSubsumRow <= 0)
                {
                    MessageBox.Show("未在当前工作表中找到有效的箱柜明细区域（Cab_Det / Cab_Subsum）！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // 元器件有效区域：从 Cab_Det.Row + 2 至 Cab_Subsum.Row - 1 (规则 6)
                int compStartRow = targetDetRow + 2;
                int compEndRow = targetSubsumRow - 1;

                // 决定物理插入的目标行号 (用户指令: 插入到当前光标所在行的下方)
                int insertRow = cursorRow + 1;
                // 如果光标在元器件区之外，则默认插入在元器件区首行
                if (insertRow <= compStartRow || insertRow > targetSubsumRow)
                {
                    insertRow = compStartRow;
                }

                // 检查是否需要在 insertRow 处插入新行
                // 规则 6: 如果元器件数量多于区域行数，先要插入行。
                // 检查光标所在行下方是否已有内容，有内容则插入新行以防覆盖
                dynamic targetRowRange = ws.Range[$"B{insertRow}:C{insertRow}"];
                object[,] cellVals = (object[,])targetRowRange.Value2;
                bool isTargetRowOccupied = cellVals[1, 1] != null && !string.IsNullOrWhiteSpace(cellVals[1, 1].ToString())
                                        || cellVals[1, 2] != null && !string.IsNullOrWhiteSpace(cellVals[1, 2].ToString());

                // 如果目标行已被占用，执行整行插入并复制上一行格式
                if (isTargetRowOccupied)
                {
                    dynamic newRow = ws.Rows[insertRow];
                    newRow.Insert(1); // 1 = xlShiftDown
                    // 复制格式
                    dynamic prevRow = ws.Rows[insertRow - 1];
                    prevRow.Copy();
                    dynamic insertedRow = ws.Rows[insertRow];
                    insertedRow.PasteSpecial(-4122); // -4122 = xlPasteFormats
                    app.CutCopyMode = false;
                }

                // 规则 7: 采用数组一次性将物料数据写入内存并回写单元格
                object[,] writeBuffer = new object[1, 13]; // 覆盖 A 到 M 列
                // A 列序号公式
                writeBuffer[0, 0] = $"=ROW()-ROW(A$6)";
                // B 列名称
                writeBuffer[0, 1] = item.Name;
                // C 列型号规格
                writeBuffer[0, 2] = item.Model;
                // D 列品牌生产厂家
                writeBuffer[0, 3] = string.IsNullOrWhiteSpace(item.Brand) ? "国优" : item.Brand;
                // E 列计量单位
                writeBuffer[0, 4] = string.IsNullOrWhiteSpace(item.Unit) ? "个" : item.Unit;
                // F 列数量
                writeBuffer[0, 5] = item.Quantity > 0 ? item.Quantity : 1.0;
                // G 列单价 (空或与 M 列联动)
                writeBuffer[0, 6] = DBNull.Value;
                // H 列总价
                writeBuffer[0, 7] = DBNull.Value;
                // I 列备注
                writeBuffer[0, 8] = item.Remark ?? string.Empty;
                // J 列成本单价
                writeBuffer[0, 9] = DBNull.Value;
                // K 列成本总价
                writeBuffer[0, 10] = DBNull.Value;
                // L 列毛利
                writeBuffer[0, 11] = DBNull.Value;
                // M 列单价数值 (纯数值或由公式联动)
                writeBuffer[0, 12] = (double)item.Price;

                // 一次性回写到目标单元格行
                dynamic writeRange = ws.Range[$"A{insertRow}:M{insertRow}"];
                writeRange.Value2 = writeBuffer;

                // 选中新插入的行提示用户
                dynamic focusCell = ws.Range[$"B{insertRow}"];
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
