using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelAddInDemo.Forms;
using ExcelAddInDemo.Models;

namespace ExcelAddInDemo.Controllers
{
    /// <summary>
    /// 箱柜管理控制器：提供批建箱柜、编辑箱柜信息、箱柜调序的 WebAPI 风格接口及弹窗唤起
    /// 遵循 WebView2 解耦与每 3 行至少 1 行中文注释规范
    /// </summary>
    public class CabinetController
    {
        // 缓存单例非模态窗体引用，杜绝窗口多开与资源冲突
        private static CabinetManageForm? _cabinetManageFormInstance;

        /// <summary>
        /// 唤起基于 WebView2 + Vue 3 的箱柜管理工作台窗口 (非模态，安全不阻塞 Excel 消息循环)
        /// </summary>
        /// <param name="mode">操作模式: batch (批建) / edit (编辑) / reorder (调序)</param>
        public static void ShowCabinetDialog(string mode = "batch")
        {
            try
            {
                // 若窗体实例已存在且未被销毁，直接切换工作台模式并前置获得焦点
                if (_cabinetManageFormInstance != null && !_cabinetManageFormInstance.IsDisposed)
                {
                    _cabinetManageFormInstance.SwitchMode(mode);
                    _cabinetManageFormInstance.BringToFront();
                    _cabinetManageFormInstance.Activate();
                    return;
                }

                // 以统一非模态安全方式展示箱柜管理窗口，防止 Chromium 消息循环与 Excel COM 产生死锁
                ExcelServices.ShowModelessForm(ref _cabinetManageFormInstance, () => new CabinetManageForm(mode));
            }
            catch (Exception ex)
            {
                // 记录异常日志
                LogHelper.WriteLog($"唤起箱柜管理窗口异常: {ex.Message}");
            }
        }

        /// <summary>
        /// WebAPI 接口: 获取当前光标命中的箱柜信息 (用于“编辑箱柜信息”窗体初始化)
        /// </summary>
        public CabinetEditDto? GetCabinetInfo()
        {
            // 调用 Excel 业务服务层提取当前选中箱柜信息
            return ExcelServices.GetActiveCabinetEditInfo();
        }

        /// <summary>
        /// WebAPI 接口: 按序号获取指定箱柜信息 (用于“编辑箱柜信息”下拉切换)
        /// </summary>
        public CabinetEditDto? GetCabinetInfoByIndex(int cabinetK)
        {
            // 调用业务服务层按序号获取指定箱柜数据
            return ExcelServices.GetCabinetEditInfoByIndex(cabinetK);
        }

        /// <summary>
        /// WebAPI 接口: 保存用户在前端界面修改的箱柜数据
        /// </summary>
        /// <param name="dto">待保存的箱柜数据传输对象</param>
        public bool SaveCabinetInfo(CabinetEditDto dto)
        {
            // 校验入参有效性
            if (dto == null) return false;
            // 调度业务服务层回写顶部汇总行与底部明细表头
            return ExcelServices.SaveCabinetEditInfo(dto);
        }

        /// <summary>
        /// WebAPI 接口: 执行批量创建箱柜
        /// </summary>
        /// <param name="items">批建明细列表</param>
        public int BatchCreateCabinets(List<BatchCabinetItemDto> items)
        {
            // 校验列表有效性
            if (items == null || items.Count == 0) return 0;
            // 调度业务层关闭渲染高速批量插行建柜
            return ExcelServices.BatchCreateCabinets(items);
        }

        /// <summary>
        /// WebAPI 接口: 获取当前工作表中全部有效箱柜列表以供调序展示
        /// </summary>
        public List<CabinetOrderItemDto> GetCabinetsForReorder()
        {
            // 扫描提取当前分类表所有箱柜
            return ExcelServices.GetCategoryCabinetsForReorder();
        }

        /// <summary>
        /// WebAPI 接口: 应用用户排列好的箱柜新顺序
        /// </summary>
        /// <param name="newOrderKList">新排序的箱柜序号列表</param>
        public bool ApplyCabinetReorder(List<int> newOrderKList)
        {
            // 校验序号集合有效性
            if (newOrderKList == null || newOrderKList.Count <= 1) return true;
            // 调度内存整块重排算法写回表格
            return ExcelServices.ApplyCabinetReorder(newOrderKList);
        }
    }
}
