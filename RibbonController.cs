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
        // 静态缓存 Ribbon UI 控制接口，支持跨类触发 Ribbon 状态重绘
        private static IRibbonUI? _ribbonInstance;

        /// <summary>
        /// 全局静态触发 Ribbon 界面状态刷新
        /// </summary>
        public static void InvalidateRibbon()
        {
            try
            {
                // 调用原生 Invalidate 刷新选项卡各控件状态
                _ribbonInstance?.Invalidate();
            }
            catch { }
        }

        /// <summary>
        /// 重写 GetCustomUI 方法，返回 Excel 选项卡 UI 结构的 XML 源码
        /// </summary>
        public override string GetCustomUI(string ribbonId)
        {
            // 返回 100% 符合 Office CustomUI 规范的标准 Ribbon UI XML 定义
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='tabDemo' label='鑫壬'>
        <!-- 账户控件分组 -->
        <group id='grpAccount' label='账户'>
          <!-- “我的” 下拉菜单控件 -->
          <menu id='menuUser' label='我的' imageMso='UserKey' size='large'>
            <!-- 企业设置按钮 -->
            <button id='btnEnterprise' label='企业设置' imageMso='Properties' onAction='OnMenuAction' />
            <!-- 上传头像按钮 -->
            <button id='btnUploadAvatar' label='上传头像' imageMso='ContactCard' onAction='OnMenuAction' />
            <!-- 我的资料按钮 -->
            <button id='btnProfile' label='我的资料' imageMso='ContactCard' onAction='OnMenuAction' />
            <!-- 一周小结按钮 -->
            <button id='btnWeekly' label='一周小结' imageMso='TableProperties' onAction='OnMenuAction' />
            <!-- 查看排行榜按钮 -->
            <button id='btnRanking' label='查看排行榜' imageMso='Rating' onAction='OnMenuAction' />
            <!-- 会员中心(订单发票) 按钮 -->
            <button id='btnVip' label='会员中心(订单发票)' imageMso='Currency' onAction='OnMenuAction' />
            <!-- ExWinner官网 按钮 -->
            <button id='btnOfficialSite' label='ExWinner官网' imageMso='WebPagePreview' onAction='OnMenuAction' />
            <!-- 分享 菜单项 -->
            <menu id='menuShare' label='分享' imageMso='Share'>
              <!-- 分享链接按钮 -->
              <button id='btnShareLink' label='分享链接' onAction='OnMenuAction' />
            </menu>
            <!-- 退出登录按钮 -->
            <button id='btnLogout' label='退出' imageMso='CloseWindow' onAction='OnMenuAction' />
          </menu>
        </group>
        <!-- “我的项目” 功能分组 -->
        <group id='grpProjects' label='我的项目'>
          <!-- 本机项目下拉菜单 -->
          <menu id='menuLocalProject' label='本机项目' imageMso='TableProperties' size='large'>
            <!-- 本机项目列表项 -->
            <button id='btnLocalProj1' label='默认本机项目' onAction='OnMenuAction' />
          </menu>
          <!-- 云项目下拉菜单 -->
          <menu id='menuCloudProject' label='云项目' imageMso='ServerProperties' size='large'>
            <!-- 云项目列表项 -->
            <button id='btnCloudProj1' label='默认云项目' onAction='OnMenuAction' />
          </menu>
        </group>
        <!-- ①建项目→ 功能分组 -->
        <group id='grpBuildProject' label='①建项目→'>
          <!-- 新建项目按钮 -->
          <button id='btnNewProject' label='新建项目' imageMso='FileNew' size='large' onAction='OnMenuAction' />
          <!-- 自动组价下拉菜单 -->
          <menu id='menuAutoPrice' label='自动组价' imageMso='TableStyles' size='large'>
            <!-- 自动组价项 -->
            <button id='btnAutoPriceSub' label='自动组价' onAction='OnMenuAction' />
          </menu>
          <!-- 国网报价下拉菜单 -->
          <menu id='menuStateGridQuote' label='国网报价' imageMso='WebPagePreview' size='large'>
            <!-- 国网报价项 -->
            <button id='btnStateGridQuoteSub' label='国网报价' onAction='OnMenuAction' />
          </menu>
          <!-- 分类功能按钮 (SplitButton 支持大图标一键直达与下拉菜单) -->
          <splitButton id='splitCategory' size='large'>
            <!-- 顶部大图标一键直接触发新建分类 -->
            <button id='btnCategorySub' label='新建分类' imageMso='GroupOutline' onAction='OnMenuAction' />
            <!-- 下拉菜单列表 -->
            <menu id='menuCategory' label='分类'>
              <button id='btnCategorySubMenu' label='新建分类' onAction='OnMenuAction' />
            </menu>
          </splitButton>
          <!-- 箱柜下拉菜单 -->
          <menu id='menuCabinet' label='箱柜' imageMso='CreateForm' size='large'>
            <!-- 1. 新建箱柜按钮 -->
            <button id='btnNewCabinet' label='新建箱柜' imageMso='CreateForm' onAction='OnMenuAction' />
            <!-- 2. 新建无明细箱柜按钮 -->
            <button id='btnNewCabinetNoDetail' label='新建无明细箱柜' onAction='OnMenuAction' />
            <!-- 3. 批建箱柜按钮 -->
            <button id='btnBatchNewCabinet' label='批建箱柜' onAction='OnMenuAction' />
            <!-- 4. 编辑箱柜信息按钮 -->
            <button id='btnEditCabinet' label='编辑箱柜信息' imageMso='EditPage' onAction='OnMenuAction' />
            <!-- 5. 剪切箱柜按钮 -->
            <button id='btnCutCabinet' label='剪切箱柜' imageMso='Cut' onAction='OnMenuAction' />
            <!-- 6. 复制箱柜按钮 -->
            <button id='btnCopyCabinet' label='复制箱柜' imageMso='Copy' onAction='OnMenuAction' />
            <!-- 7. 插入复制的箱柜按钮 -->
            <button id='btnInsertCopiedCabinet' label='插入复制的箱柜' imageMso='Paste' onAction='OnMenuAction' />
            <!-- 8. 删除箱柜按钮 -->
            <button id='btnDeleteCabinet' label='删除箱柜' imageMso='Delete' onAction='OnMenuAction' />
            <!-- 9. 箱柜调序按钮 -->
            <button id='btnReorderCabinet' label='箱柜调序' imageMso='SortAscending' onAction='OnMenuAction' />
            <!-- 10. 导入箱柜BOM按钮 -->
            <button id='btnImportCabinetBOM' label='导入箱柜BOM' imageMso='ImportTextFile' onAction='OnMenuAction' />
            <!-- 11. 智能导入箱柜BOM按钮 -->
            <button id='btnSmartImportCabinetBOM' label='智能导入箱柜BOM' imageMso='ImportXml' onAction='OnMenuAction' />
          </menu>
        </group>
        <!-- ②录元件 功能分组 -->
        <group id='grpInputComponents' label='②录元件'>
          <!-- 智能识图下拉菜单 -->
          <menu id='menuSmartOCR' label='智能识图' imageMso='FindDialog' size='large'>
            <!-- 智能识图项 -->
            <button id='btnSmartOCRSub' label='智能识图' onAction='OnMenuAction' />
          </menu>
          <!-- 云方案按钮 -->
          <button id='btnCloudSolution' label='云方案' imageMso='ServerProperties' size='large' onAction='OnMenuAction' />
          <!-- 云物料下拉菜单 -->
          <menu id='menuCloudMaterial' label='云物料' imageMso='TableProperties' size='large'>
            <!-- 云物料库项 -->
            <button id='btnCloudMaterialSub' label='云物料库' onAction='OnMenuAction' />
          </menu>
          <!-- 型号识别(提取极数与电流) 按钮 -->
          <button id='btnModelParamParser' label='识别极数电流' imageMso='AutoFilter' size='large' screentip='型号识别极数电流' supertip='自动从型号中识别并提取电流和极数，支持双通道顺位流水线与白名单过滤' onAction='OnMenuAction' />
          <!-- 元器件数据管理按钮 (支持在 Excel 中直接查看、批量筛选、选中行更新/新增/删除) -->
          <button id='btnComponentManage' label='元器件管理' imageMso='TableInsertRowsAbove' size='large' screentip='元器件数据管理' supertip='在 Excel 中按品牌和名称筛选元器件数据，支持对选中行进行精准更新、新增和删除' onAction='OnMenuAction' />
          <!-- 二次元件组规则管道(生成二次) 按钮 -->
          <button id='btnComponentGroupRule' label='生成二次元件' imageMso='TableFormulaDialog' size='large' screentip='二次元件组规则管道' supertip='基于可视化动态规则管道自动识别箱柜元件特征，生成二次元件组并自动写入套数' onAction='OnMenuAction' />
          <!-- 二次图方案与 BOM 库按钮 -->
          <button id='btnSecondaryCircuitManage' label='二次方案库' imageMso='QueryShowTable' size='large' screentip='二次图方案与BOM管理' supertip='管理二次原理图控制回路方案、同配置多回路映射、门板开孔、人工工费及BOM物料定额' onAction='OnMenuAction' />
        </group>
        <!-- ③调价格→ 功能分组 -->
        <group id='grpAdjustPrice' label='③调价格→'>
          <!-- 元件批调下拉菜单 -->
          <menu id='menuBatchAdjust' label='元件批调' imageMso='AutoSum' size='large'>
            <!-- 1. 汇总调价按钮 -->
            <button id='btnSummaryAdjustPrice' label='汇总调价' imageMso='AutoSum' screentip='汇总调价' supertip='汇总项目元件，快速调改名称、型号、...、价格...' onAction='OnMenuAction' />
            <!-- 2. 分布调价按钮 -->
            <button id='btnDistributedAdjustPrice' label='分布调价' imageMso='PivotTableInsert' onAction='OnMenuAction' />
            <!-- 3. 筛选调价按钮 -->
            <button id='btnFilterAdjustPrice' label='筛选调价' imageMso='Filter' onAction='OnMenuAction' />
            <!-- 4. 一键匹配价格按钮 -->
            <button id='btnAutoMatchPrice' label='一键匹配价格' imageMso='Pointer' onAction='OnMenuAction' />
            <!-- 5. 母排用量一键预估按钮 -->
            <button id='btnEstimateBusbarUsage' label='母排用量一键预估' imageMso='ChartInsert' onAction='OnMenuAction' />
            <!-- 6. 箱体尺寸一键预估级联菜单 -->
            <menu id='menuEstimateCabinetSize' label='箱体尺寸一键预估' imageMso='FlashFill'>
              <!-- 箱体尺寸一键预估子项 -->
              <button id='btnEstimateCabinetSizeSub' label='箱体尺寸一键预估' onAction='OnMenuAction' />
            </menu>
            <!-- 7. 母排一键改价按钮 -->
            <button id='btnBusbarBatchPrice' label='母排一键改价' imageMso='TableInsertRowsAbove' onAction='OnMenuAction' />
            <!-- 8. 箱体一键改价按钮 -->
            <button id='btnCabinetBatchPrice' label='箱体一键改价' imageMso='ShapeCube' onAction='OnMenuAction' />
            <!-- 9. 多方案报价级联菜单 -->
            <menu id='menuMultiPlanQuote' label='多方案报价' imageMso='FileNew'>
              <!-- 多方案报价子项 -->
              <button id='btnMultiPlanQuoteSub' label='多方案报价' onAction='OnMenuAction' />
            </menu>
          </menu>
          <!-- 智能算料/辅材壳体计算 按钮 (大图标直达) -->
          <button id='btnCabinetAuxCalc' label='辅材壳体计算' imageMso='CalculateNow' size='large' screentip='辅材壳体与配电智能计算' supertip='智能推导匹配壳体尺寸、计算铜排母线用量、一次及二次接线辅材与装配人工费，支持全参数动态配置' onAction='OnMenuAction' />
          <!-- 费用设定下拉菜单 -->
          <menu id='menuFeeSetting' label='费用设定' imageMso='Currency' size='large'>
            <!-- 1. 公式法调费按钮 (对应图二样式与提示) -->
            <button id='btnFormulaAdjustFee' label='公式法调费' imageMso='Calculator' screentip='公式法调费' supertip='按明细总价，套用公式算成套费(管理费、利润等)。' onAction='OnMenuAction' />
            <!-- 2. 系数法调费按钮 -->
            <button id='btnCoefficientAdjustFee' label='系数法调费' imageMso='PercentSymbol' onAction='OnMenuAction' />
            <!-- 3. 智能算料按钮 -->
            <button id='btnSmartMaterial' label='智能算料' imageMso='AutoSum' onAction='OnMenuAction' />
            <!-- 4. 费用公式转值按钮 -->
            <button id='btnFormulaToValue' label='费用公式转值' imageMso='PasteValues' onAction='OnMenuAction' />
            <!-- 5. 总价一键调整按钮 -->
            <button id='btnOneKeyAdjustTotal' label='总价一键调整' imageMso='Gauge' onAction='OnMenuAction' />
            <!-- 6. 总价高级调整按钮 -->
            <button id='btnAdvancedAdjustTotal' label='总价高级调整' imageMso='Diamond' onAction='OnMenuAction' />
            <!-- 默认费用设定子菜单选项 -->
            <button id='btnFeeSettingSub' label='费用设定' onAction='OnMenuAction' />
          </menu>
          <!-- 撤销/还原下拉菜单 -->
          <menu id='menuUndoRedo' label='撤销/还原' imageMso='Undo' size='large'>
            <!-- 撤销/还原项 -->
            <button id='btnUndoRedoSub' label='撤销/还原' onAction='OnMenuAction' />
          </menu>
        </group>
        <!-- ④出报表 功能分组 -->
        <group id='grpExportReports' label='④出报表'>
          <!-- 标书报表下拉菜单 -->
          <menu id='menuTenderReport' label='标书报表' imageMso='PrintPreviewAndPrint' size='large'>
            <!-- 标书报表项 -->
            <button id='btnTenderReportSub' label='标书报表' onAction='OnMenuAction' />
          </menu>
          <!-- 材料统计下拉菜单 -->
          <menu id='menuMaterialStat' label='材料统计' imageMso='ChartInsert' size='large'>
            <!-- 材料统计项 -->
            <button id='btnMaterialStatSub' label='材料统计' onAction='OnMenuAction' />
          </menu>
        </group>
        <!-- 辅助项 功能分组 -->
        <group id='grpAuxiliary' label='辅助项'>
          <!-- 聚光灯行列高亮切换按钮 -->
          <toggleButton id='btnToggleSpotlight' label='聚光灯' imageMso='PivotTableVisualFilter' size='large' getPressed='GetSpotlightPressed' onAction='OnSpotlightAction' screentip='行列聚光灯 (Ctrl+Alt+L)' supertip='以十字半透明柔和色彩高亮选中单元格所在行与列，无损Excel撤销重做(Ctrl+Z)且零文件修改' />
          <!-- 联动CAD夹点显示切换按钮 -->
          <toggleButton id='btnToggleCadSync' label='联动CAD' imageMso='SelectionPane' size='large' getPressed='GetCadSyncPressed' onAction='OnCadSyncAction' screentip='联动AutoCAD夹点' supertip='选中行时自动读取AA列句柄，在AutoCAD中即时高亮并激活夹点显示' />
          <!-- 右键菜单模式切换按钮 -->
          <button id='btnToggleContextMenuMode' label='右键菜单模式' imageMso='ControlsGallery' size='large' screentip='切换右键菜单模式' supertip='在【业务专属菜单】与【Excel 原生右键菜单】之间彻底二选一切换' onAction='OnMenuAction' />
          <!-- 项目工具下拉菜单 -->
          <menu id='menuProjectTools' label='项目工具' imageMso='Tools' size='large'>
            <!-- 智能输入按钮 -->
            <button id='btnSmartInput' label='智能输入' imageMso='SmartArtInsert' screentip='智能输入' supertip='配置元器件去重词库与C列输入智能联动选项' onAction='OnMenuAction' />
            <!-- 项目工具项 -->
            <button id='btnProjectToolsSub' label='项目工具' onAction='OnMenuAction' />
          </menu>
          <!-- 企业DHub按钮 -->
          <button id='btnEnterpriseDHub' label='企业DHub' imageMso='ServerProperties' size='large' onAction='OnMenuAction' />
          <!-- 设置按钮 -->
          <button id='btnSettings' label='设置' imageMso='OptionButton' size='large' onAction='OnMenuAction' />
        </group>
        <!-- 服务 功能分组 -->
        <group id='grpService' label='服务'>
          <!-- 服务下拉菜单 -->
          <menu id='menuService' label='服务中心' imageMso='Help' size='large'>
            <!-- 在线客服按钮 -->
            <button id='btnServiceSub' label='在线客服' onAction='OnMenuAction' />
          </menu>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        /// <summary>
        /// 聚光灯切换按钮状态获取回调：读取当前聚光灯激活状态
        /// </summary>
        public bool GetSpotlightPressed(IRibbonControl control)
        {
            // 读取 ExcelServices 中的聚光灯开关状态
            return ExcelServices.IsSpotlightEnabled;
        }

        /// <summary>
        /// 聚光灯切换按钮点击回调：开启或关闭聚光灯功能
        /// </summary>
        public void OnSpotlightAction(IRibbonControl control, bool isPressed)
        {
            // 同步切换聚光灯状态
            if (isPressed)
            {
                // 开启聚光灯
                ExcelServices.EnableSpotlight();
            }
            else
            {
                // 关闭聚光灯
                ExcelServices.DisableSpotlight();
            }
        }

        /// <summary>
        /// 切换按钮状态获取回调：读取是否开启与 AutoCAD 夹点联动
        /// </summary>
        public bool GetCadSyncPressed(IRibbonControl control)
        {
            // 读取 CadSyncClient 中的联动开关状态
            return Services.CadSyncClient.SyncToCadEnabled;
        }

        /// <summary>
        /// 切换按钮点击回调：切换与 AutoCAD 夹点联动开关
        /// </summary>
        public void OnCadSyncAction(IRibbonControl control, bool isPressed)
        {
            // 将用户最新切换的状态同步至 CadSyncClient
            Services.CadSyncClient.SyncToCadEnabled = isPressed;
        }

        /// <summary>
        /// 按钮回调：弹出基于 WebView2 + Vue 3 的用户登录与配置窗口
        /// </summary>
        public void OnShowLoginClicked(IRibbonControl control)
        {
            // 调用业务层服务在独立 STA 线程中弹出登录窗体
            ExcelServices.ShowLoginDialog();
        }

        /// <summary>
        /// Ribbon 加载完成后的回调函数
        /// </summary>
        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            // 保存 Ribbon UI 对象引用以便需要时刷新 UI
            _ribbon = ribbon;
            // 同步更新静态实例引用以支持外部类调用 InvalidateRibbon
            _ribbonInstance = ribbon;
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
        /// Ribbon 下拉菜单及菜单按钮通用响应回调
        /// </summary>
        public void OnMenuAction(IRibbonControl control)
        {
            // 读取触发响应的控件唯一标识 ID
            string controlId = control.Id;

            // 响应“新建箱柜”按钮指令
            if (controlId == "btnNewCabinet")
            {
                // 调用业务层服务：直接在顶部“成套产品报价清单”插入行并复制模板明细
                ExcelServices.CreateNewCabinetFromSelection();
            }
            // 响应“删除箱柜”按钮指令
            else if (controlId == "btnDeleteCabinet")
            {
                // 调用业务层服务：删除当前选中的箱柜及其明细数据
                ExcelServices.DeleteCabinetFromSelection();
            }
            // 响应“企业设置”按钮指令
            else if (controlId == "btnEnterprise")
            {
                // 弹出基于 WebView2 + Vue 3 的“企业设置”窗口
                ExcelServices.ShowEnterpriseSettingsDialog();
            }
            // 响应“新建项目”按钮指令
            else if (controlId == "btnNewProject")
            {
                // 弹出基于 WebView2 + Vue 3 的“新建项目”窗口
                ExcelServices.ShowCreateProjectDialog();
            }
            // 响应“我的资料”按钮指令
            else if (controlId == "btnProfile")
            {
                // 弹出用户登录与配置窗体
                ExcelServices.ShowLoginDialog();
            }
            // 响应“公式法调费”及“费用设定”按钮指令
            else if (controlId == "btnFormulaAdjustFee" || controlId == "btnFeeSettingSub")
            {
                // 弹出基于 WebView2 + Vue 3 的“公式法调费”窗口
                ExcelServices.ShowFormulaAdjustFeeDialog();
            }
            // 响应“新建分类”按钮指令 (支持 splitButton 顶部直达与下拉菜单项)
            else if (controlId == "btnCategorySub" || controlId == "btnCategorySubMenu")
            {
                // 弹出基于 WebView2 + Vue 3 的“新建分类”窗口
                ExcelServices.ShowCategoryDialog();
            }
            // 响应“汇总调价”按钮指令
            else if (controlId == "btnSummaryAdjustPrice")
            {
                // 弹出基于 WebView2 + Vue 3 的“汇总调价”窗口
                ExcelServices.ShowSummaryAdjustPriceDialog();
            }
            // 响应“智能输入”按钮指令
            else if (controlId == "btnSmartInput")
            {
                // 弹出基于 WebView2 + Vue 3 的“智能输入配置”窗口
                ExcelServices.ShowSmartInputDialog();
            }
            // 响应“识别极数电流”按钮指令
            else if (controlId == "btnModelParamParser")
            {
                // 弹出基于 WebView2 + Vue 3 的“型号参数识别设置”窗口
                ExcelServices.ShowModelParamParserDialog();
            }
            // 响应“生成二次元件 (规则管道)”按钮指令
            else if (controlId == "btnComponentGroupRule")
            {
                // 弹出基于 WebView2 + Vue 3 的“二次元件组规则管道构建器”窗口
                ExcelServices.ShowComponentGroupBuilderDialog();
            }
            // 响应“二次方案库”按钮指令
            else if (controlId == "btnSecondaryCircuitManage")
            {
                // 弹出基于 WebView2 + Vue 3 的“二次图回路方案与 BOM 管理中心”窗口
                ExcelServices.ShowSecondaryCircuitManageDialog();
            }
            // 响应“元器件管理”按钮指令
            else if (controlId == "btnComponentManage")
            {
                // 弹出基于 WebView2 + Vue 3 的“元器件数据管理”悬浮窗口
                ExcelServices.ShowComponentManageDialog();
            }
            // 响应“智能辅材壳体计算”/“智能算料”/“箱体尺寸预估”按钮指令
            else if (controlId == "btnCabinetAuxCalc" || controlId == "btnSmartMaterial" || controlId == "btnEstimateCabinetSizeSub")
            {
                // 弹出基于 WebView2 + Vue 3 的“智能辅材与壳体计算”工作台
                ExcelServices.ShowCabinetAuxCalcDialog();
            }
            // 响应“切换右键菜单模式”按钮指令
            else if (controlId == "btnToggleContextMenuMode")
            {
                // 调度宏一键切换右键菜单模式并提示
                ExcelEventManager.MacroToggleContextMenuMode();
            }
        }
    }
}
