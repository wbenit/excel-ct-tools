using System;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Ribbon GUI 界面控制器，负责自定义选项卡与 Excel 菜单回调交互
    /// </summary>
    [ComVisible(true)]
    // 特别标注加壳特性：排除类与回调方法的名称重命名混淆，确保商业加壳后 Ribbon XML 反射调用不受影响
    [Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class RibbonController : ExcelRibbon
    {
        // 缓存 Ribbon UI 控制接口句柄
        private IRibbonUI? _ribbon;

        /// <summary>
        /// 重写 GetCustomUI 方法，返回 Excel 选项卡 UI 结构的 XML 源码
        /// </summary>
        public override string GetCustomUI(string ribbonId)
        {
            // 定义包含选项卡 (Tab)、分组 (Group)、按钮 (Button)、复选框 (CheckBox) 及编辑框 (EditBox) 的 XML
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='tabDemo' label='鑫壬成套服务'>
        <group id='grpData' label='数据写入操作'>
          <button id='btnInsert' label='插入当前时间' imageMso='DateInsert' size='large' onAction='OnInsertDataClicked' />
          <button id='btnClear' label='清空选中区域' imageMso='EditClear' size='large' onAction='OnClearRangeClicked' />
        </group>
        <group id='grpConfig' label='交互与格式'>
          <checkBox id='chkHighlight' label='开启自动高亮' getPressed='GetAutoHighlightPressed' onAction='OnAutoHighlightToggled' />
          <editBox id='txtMessage' label='自定义内容:' getText='GetCustomMessageText' onChange='OnCustomMessageChanged' />
          <button id='btnApply' label='批量填充文本' imageMso='FontColorPicker' size='large' onAction='OnApplyCustomTextClicked' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        /// <summary>
        /// Ribbon 加载完成后的回调函数
        /// </summary>
        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            // 保存 Ribbon UI 对象引用以便需要时刷新 UI
            _ribbon = ribbon;
        }

        /// <summary>
        /// 按钮回调：插入当前时间戳
        /// </summary>
        public void OnInsertDataClicked(IRibbonControl control)
        {
            // 执行业务层逻辑，在活动单元格填入时间数据
            ExcelServices.InsertTimestampAndData();
        }

        /// <summary>
        /// 按钮回调：清空当前选中区域
        /// </summary>
        public void OnClearRangeClicked(IRibbonControl control)
        {
            // 执行业务层逻辑，清除选中单元格的格式与内容
            ExcelServices.ClearActiveRange();
        }

        /// <summary>
        /// 复选框状态获取回调：加载复选框是否选中的初始化状态
        /// </summary>
        public bool GetAutoHighlightPressed(IRibbonControl control)
        {
            // 读取业务服务类中高亮标识状态
            return ExcelServices.IsAutoHighlightEnabled;
        }

        /// <summary>
        /// 复选框状态改变回调：响应用户勾选/取消勾选操作
        /// </summary>
        public void OnAutoHighlightToggled(IRibbonControl control, bool pressed)
        {
            // 将最新的勾选状态同步给业务服务类
            ExcelServices.IsAutoHighlightEnabled = pressed;
        }

        /// <summary>
        /// 编辑框默认文本获取回调：加载文本框初始显示内容
        /// </summary>
        public string GetCustomMessageText(IRibbonControl control)
        {
            // 读取业务服务类中保存的文本框值
            return ExcelServices.CustomMessageText;
        }

        /// <summary>
        /// 编辑框内容修改回调：响应用户在界面输入新文本
        /// </summary>
        public void OnCustomMessageChanged(IRibbonControl control, string text)
        {
            // 更新保存自定义消息文本
            ExcelServices.CustomMessageText = text;
        }

        /// <summary>
        /// 按钮回调：批量将编辑框文本应用填入选中单元格区域
        /// </summary>
        public void OnApplyCustomTextClicked(IRibbonControl control)
        {
            // 调用业务逻辑批量写入选中的单元格
            ExcelServices.ApplyCustomTextToSelection();
        }
    }
}
