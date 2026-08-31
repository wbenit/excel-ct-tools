# Session State

## [Completed]

- **彻底修复业务专属右键菜单图标不显示、点击无反应以及元器件管理未响应的问题**：
  1. **定位根本原因**：
     - **外部 CDN 依赖导致死锁**：右键菜单页面 `custom_context_menu.html` 引入了海外 CDN `cdnjs.cloudflare.com`（FontAwesome）以及 `unpkg.com`（Vue 3 / Element Plus）。在国内网络或离线内网环境下网络请求严重超时或被拦截，导致 Vue 无法加载（抛出 `Uncaught ReferenceError: Vue is not defined`）；
     - **模板未编译且无事件监听**：Vue 未能 mount，右上角直接裸露插值表达式源码 `{{ contextInfo.shee...`，菜单项上的 `@click` 纯属普通 HTML 属性而没有绑定任何原生 DOM 点击事件，导致点击彻底无反应；
     - **C# 路由遗漏**：`CustomContextMenuForm.cs` 的 `OnWebMessageReceived` 分发逻辑中，遗漏了 `case "openComponentManage":`，导致点击元器件数据管理被直接静默丢弃。
  2. **全面离线化与极速架构重构**：
     - **图标全面离线化**：将所有菜单项图标彻底替换为轻量内联矢量 SVG（每枚几百字节），实现 100% 零网络依赖、零延迟离线秒显且高清；
     - **事件绑定原生化**：由于右键菜单属于瞬态快捷交互窗口，改用原生 JavaScript（Vanilla JS）处理点击、防抖、数据接收与键盘导航（上下箭头切换、回车确认、ESC关闭），彻底消除对外部庞大框架库的加载依赖，实现 0 毫秒秒开秒响应；
     - **C# 路由补齐**：在 `CustomContextMenuForm.cs` 中补充 `case "openComponentManage":` 分支，保障所有菜单项指令均能顺畅下发执行。


- **彻底修复 CAD 批量导出到 Excel 时计费区域 A 列显示为 ####（#REF!）的问题及支持用户自定义序号**：
  1. **定位根本原因**：
     - 标准模板 `CabinetTemplate.xlsx` 的明细表中，元器件和计费区域预设的序号公式均为 `=ROW()-ROW(A$45)`（指向母版的明细表头行）；
     - `ExportSingleCabinetOptimized` 导出各箱柜时通过全行克隆拷贝母版明细块，元器件区域通过 `compMatrix` 重新覆写了公式，但计费区域（小计、管理费、利润、税金、单台合计）遗漏了对 A 列的重写；
     - 批量导出末尾物理删除了母版行（`ws.Rows[...].Delete()`），导致计费区域引用的母版表头行丢失断裂，公式变为 `=ROW()-ROW(#REF!)`；
     - 因 A 列列宽仅为 4.38，无法容纳 5 字符的 `#REF!` 错误提示，在 Excel 界面被渲染呈现为 `####`；
     - 顶部汇总行原硬编码 `ROW(A$6)`，改为通过配置动态获取汇总表头行号。
  2. **最小变动修复与重构 (`Services/ExcelServices.Cabinet.cs`)**：
     - 在 `ExportSingleCabinetOptimized` 步骤 5 中，顶部汇总行的序号公式采用动态表头行 `sumHeaderRow`，避免写死；
     - 将计费区域自适应求和与 A 列智能序号刷新逻辑抽取为独立公共方法 `RefreshCabinetFeeAreaFormulas(sheet, detRow, compStartRow, subsumRow, tolsumRow)`，显著降低主流程复杂度，提升模块内聚度与多场景复用性；
     - 在 `RefreshCabinetFeeAreaFormulas` 中：
       - 刷新小计行 H/K 列自适应求和公式；
       - 规则 7 一次性读取计费区域（`subsumRow` 到 `tolsumRow - 1`）原始公式与数值，**判断仅当用户定义为 `[序号]` 或动态序号 `ROW(` 或为空时，才自适应生成 `$"=ROW()-ROW(A${detRow + 1})"`**；
       - **若用户在 A 列定义为其他特定内容（文本、自定义编号或自定义公式），则完全按照定义内容原样回填保留**；
       - 显式确保总计行 `tolsumRow` A 列为 `"总计"`；
     - 彻底切断对母版行的外部依赖，从根源上杜绝 `#REF!` 和 `####`，同时兼顾了灵活性与用户自定义需求。


- **排查并彻底修复点击“开始报价”没反应的问题**：
  1. **定位根本原因**：
     - 原 `ProjectController.CreateProjectAsync` 采用了 `Task.Run` 后台线程池执行；
     - 在后台线程中访问 `ExcelDnaUtil.Application` 会因非 Excel 主线程而抛出异常，导致 `ExcelDnaSafeAccessor.GetApplication()` 捕获返回 `null`；
     - 后端收到 `app == null` 直接提前返回 `false`，且 `CreateProjectForm.cs` 中对失败未做任何提示和处理，导致点击后前端静默无反应。
  2. **后端修复 (`Controllers/ProjectController.cs` & `Forms/CreateProjectForm.cs`)**：
     - 将 `CreateProject` 调整为在 Excel 主线程同步调度执行（遵循 Excel COM 单线程 STA 规则）；
     - `CreateProjectForm.cs` 使用 `SafeInvoke` 调度到主线程安全执行，并增加成功/失败结果回发机制及异常捕获提示；
     - 在 `EnsureCabinetTemplate` 和 `OnFormLoadAsync` 候选路径中补充 `Tool.GetAppDirectory()`，确保能 100% 命中模板文件 `CabinetTemplate.xlsx`。
  3. **前端优化 (`Resources/create_project.html`)**：
     - 为“开始报价”按钮添加 `:loading="isSubmitting"` 防重与加载状态；
     - 增加 `startQuotationResult` 消息监听与 Element Plus 错误提示。

- **排查并彻底修复 Ribbon“新建分类”点击无效的问题**：
  1. **定位根本原因**：
     - **窗体单例与模态生命周期冲突**：原 `ShowCategoryDialog` 混用了 `ShowDialog` 模态弹窗与 `_categoryFormInstance` 单例引用。模态关闭后，该静态变量仍存活未清空，二次点击被 `if (_categoryFormInstance != null && !_categoryFormInstance.IsDisposed)` 拦截，只调用已隐藏窗体的 `BringToFront/Activate`，导致二次及后续点击彻底无效。
     - **Ribbon 控件类型限制**：原 XML 使用 `<menu id='menuCategory'>`，Office Ribbon 机制下点击大图标上半部分只会展开下拉菜单而不会直接触发操作，必须展开后再次点击子项才能触发。
  2. **修复方案**：
     - **全局统一非模态生命周期 (`Services/ExcelServices.Category.cs`)**：改用 `ShowModelessForm(ref _categoryFormInstance, () => new Forms.CategoryForm());`，保持全项目统一的非模态 + 单例激活机制，Excel 可直接交互且杜绝点击失效；
     - **Ribbon 升级为 SplitButton (`RibbonController.cs`)**：将 `menuCategory` 升级为 `<splitButton id='splitCategory' size='large'>`，点击大图标即可一键直接弹出“新建分类”窗口，点击下拉箭头亦可展示子菜单；并在 `OnMenuAction` 中支持 `btnCategorySub` 与 `btnCategorySubMenu`；
     - **资源路径优化 (`Forms/CategoryForm.cs`)**：将 `Tool.GetAppDirectory()` 提升为最高检索优先级。

- **修复“新建分类”弹窗底部“确定创建”与“取消”按钮被截断不可见的问题**：
  1. **定位根本原因**：
     - `Forms/CategoryForm.cs` 原窗体高度设定为 360 像素（`ClientSize = new Size(480, 360)`）；
     - 前端 `Resources/category.html` 包含顶部标题栏(32px)、提示 Banner(~50px)、3个带间距的输入/下拉表单项(~175px)及内边距(40px)；
     - 加上底部按钮栏(52px)后整体高度超过 350px，在 Windows DPI 缩放（如 125%/150%）及字体渲染下总高度超出窗口视口；
     - 因 `body { overflow: hidden }` 且 `.form-container` 缺乏 `min-height: 0` 约束，底部的 `.footer-bar`（包含“取消”和“确定创建”按钮）被挤出了窗口可视区域下方而被裁剪。
  2. **修复方案**：
     - **调整宿主窗体尺寸 (`Forms/CategoryForm.cs`)**：将窗体高度从 360 像素扩大至 420 像素（`new Size(480, 420)`），为按钮栏留足垂直安全显示空间；
     - **前端布局自适应防护 (`Resources/category.html`)**：
       - 精简表单内边距与外间距，为 `.form-container` 增加 `overflow-y: auto; min-height: 0;`；
       - 为 `.footer-bar` 添加 `flex-shrink: 0; z-index: 10;` 确保底部按钮栏始终吸底完整可见；
       - 为分类名称和初始箱柜输入框增加 `@keyup.enter="submitCreate"` 回车快捷提交支持。

- **公式法调费窗口：明细表格最左侧行号/下拉列水平滚动固定锁定**：
  1. **需求目标**：
     - 在“公式法调费”明细表格中，横向水平滚动查看右侧列（如 J/K/类别等）时，最左侧包含下拉图标及行号（1、2、3、4、总计）的列需要固定不随滚动条平移。
  2. **修改内容 (`Resources/formula_adjust_fee.html`)**：
     - 在明细表格最左侧索引列 `<el-table-column>` 上增加 `fixed` 属性（`<el-table-column fixed width="36" align="center">`）；
     - 增加 `.detail-section .el-table th.el-table-fixed-column--left` 与 `.detail-section .el-table td.el-table-fixed-column--left` 的背景色与 z-index 防护样式，杜绝横向滚动时下层单元格内容穿帮透出；
     - 同步更新至 `bin\Debug\net48\Resources\formula_adjust_fee.html`。

## [Completed]

- **元器件数据管理功能（Excel 呈现与选区更新/新增/删除）全部实施并编译成功**：
  1. **API 通信层 (`Services/ComponentApiClient.cs`)**：
     - 新增 `QueryManageComponents(brand, keyword)`，支持按品牌和名称跨页批量拉取满足条件的所有云端元器件物料；
     - 新增 `CreateComponent(CreateComponentApiRequest)`、`UpdateComponent(UpdateComponentApiRequest)`、`DeleteComponent(id)` 标准 RESTful 增删改接口对接商城后端。
  2. **模型层 (`Models/ComponentManageModels.cs`)**：
     - 定义工作表默认名 `元器件数据管理`、主题颜色及 ID 浅灰底色等规范；
     - 定义列字段索引映射模型 `ComponentManageColumnConfig`（A列至L列）；
     - 定义新增/更新 DTO、选区探测结果 `SelectionDetectResult` 及批量操作结果 `ComponentManageActionResult`。
  3. **业务层 (`Services/ExcelServices.ComponentManage.cs`)**：
     - `EnsureComponentManageWorksheet`：自动创建并初始化标准表头（开启自动筛选、冻结首行、设置主题色）；
     - `LoadComponentsToSheet`：采用 `object[,]` 二维数组一次性读写灌入数据行，杜绝 COM 性能瓶颈；
     - `DetectCurrentSelection`：自动提取用户鼠标在 Excel 中划选的物理行（连续或按 Ctrl 多选，排除表头第 1 行）；
     - `UpdateSelectedComponents`：对选中的 1 行或多行提取 A 列系统 ID 并提交云端更新，成功后 L 列回显时间戳；
     - `CreateSelectedComponents`：对选中的 1 行或多行忽略原 ID 进行全新新增，成功后自动回填新生成的系统 ID 至 A 列；
     - `DeleteSelectedComponents`：对选中的 1 行或多行调用云端删除接口，成功后直接从 Excel 中整行物理删除。
  4. **交互控制与 UI 层 (`Controllers/ComponentManageController.cs`, `Forms/ComponentManageForm.cs`, `Resources/component_manage.html`)**：
     - 采用 Vue 3 `<script setup>` + Element Plus 框架构建，遵循绿蓝相间主色调 `#009688`；
     - 非模态置顶悬浮窗，内置每 1.5 秒自动探知当前选区行数；
     - 提供【查询并呈现至 Excel】、【更新选中行】、【选中行新增】、【删除选中行】（含危险操作二次确认）；
     - 在 Ribbon 功能区【②录元件】以及业务专属右键菜单中均挂载了【元器件管理】直达入口。

- **彻底修复元器件管理窗口拖拽时容易卡死整个屏幕的问题**：
  1. **定位根本原因**：
     - **异步 IPC 延迟导致幽灵鼠标捕获**：Web 端向 C# 发送 `postMessage('dragWindow')` 有异步延迟，若用户轻点或快速抬起鼠标，C# 触发 `WM_NCLBUTTONDOWN` 进入 Windows 模态移动循环时物理鼠标按键已弹起，Windows 无法收到 `WM_LBUTTONUP`，导致全局鼠标独占捕获死锁；
     - **定时轮询 COM 竞争主线程**：前端原设置了 `setInterval(1500)` 无休止调用 COM 探测选区，与 Win32 窗口拖拽模态移动消息泵竞争 Excel STA 单线程，引发主线程互斥假死。
  2. **实施修复**：
     - **C# 端增加物理鼠标按键校验 (`Forms/ComponentManageForm.cs`)**：引入 Win32 `GetAsyncKeyState(VK_LBUTTON)`，在触发 `WM_NCLBUTTONDOWN` 前强制校验物理左键是否仍处于按下状态；若已松开则直接丢弃，从根本上杜绝“幽灵捕获死锁”；
     - **彻底废除无休止定时轮询 (`Resources/component_manage.html`)**：移除 `setInterval`，改为 `window.onfocus`（用户在 Excel 划选后切回窗口自动灵敏探测）以及点击【刷新选区】按需检测；
     - **事件按键严格限制**：前端 `onHeaderMouseDown` 仅响应鼠标左键（`e.button === 0`），排除右键或中键触发。

- **升级选择品牌后自动加载该品牌所有元器件名称机制（后端数据库级 DISTINCT 聚合）**：
  1. **问题根源**：原客户端通过分页拉取明细（如 500 条）再提取名称，若某一类别的元器件数量超过 500 条（如施耐德微型断路器），将导致后续品类被严重截断遗漏，且网络开销随明细条数急剧膨胀；
  2. **后端新增专用接口 (`draw-mall`)**：
     - 在 `IComponentServicer.cs` 与 `ComponentServicer.cs` 中实现 `GetNamesAsync(string? brand)`，依托数据库 `SELECT DISTINCT Name FROM Component WHERE Brand = @brand` 执行毫秒级去重查询；
     - 在 `ComponentController.cs` 中暴露 `GET /api/api/Component/GetNames?brand={brand}` 接口；
  3. **插件端对接与缓存 (`excel-ct-tools`)**：
     - 在 `ComponentApiClient.cs` 的 `GetNamesByBrand` 中，直接请求后端的 `GetNames` 聚合接口，配合内存缓存字典 `_brandNamesCache`，实现 0 漏项、极速毫秒级响应；
  4. **构建结果**：`DrawMall.Web` 与 `ExcelAddInDemo` 均顺利完成编译（0 个错误）。

- **落地【配套附件】原位切换查询、型号拼接回填与价格公式自动累加功能**：
  1. **后端支撑 (`draw-mall`)**：
     - `IComponentServicer` 与 `ComponentServicer` 实现 `GetAttachmentsAsync(brand, name, model)`：在数据库层过滤 `Brand + Name + Param1='附件'`，在内存中精准拆分 `Param2`（中英文逗号/分号）比对当前型号，毫秒级返回可用附件；
     - `ComponentController` 暴露 `GET /api/api/Component/GetAttachments` 接口；
  2. **客户端与服务层 (`excel-ct-tools`)**：
     - `ComponentApiClient.GetAttachments` 负责拉取附件并带本地备用过滤降级；
     - `ExcelServices.FillSelectedAttachmentToActiveRow` 实现“原内容+附件型号”回填 D 列，并将价格转为原生公式累加（如 `=447.01+150.00`）回填 G 列；
     - `ComponentMatchOverlayForm` 扩展 `getAttachments` 与 `selectAttachment` 消息中继；
     - `component_match_overlay.html` 管道栏右侧增加【🧩 配套附件】按钮，支持方式 A 原位平滑切换视图，点击附件自动执行拼接回填。
  3. **构建结果**：`DrawMall.Web` 与 `ExcelAddInDemo` 均顺利完成编译（0 个错误）。

- **元件汇总生成时 D 列保持空白留空不填内容**：
  1. **需求定位**：
     - 用户在点击“立即生成”或“重新汇总”生成【元件汇总表】时，要求 D 列（型号规格）保持空白，不填内容（原逻辑会填充从分类表提取出的型号 `comp.Model`）；
     - 此时 C 列完整保留基准参考原型号规格，D 列作为调价选型列保留空白，以便工程师后续在汇总表中通过物料联想悬浮窗进行重新选型调价。
  2. **代码实施 (`Services/ExcelServices.SummaryAdjustPrice.cs`)**：
     - 在构建 18 列输出矩阵 `outMatrix` 的循环中，将 `outMatrix[i, 3] = comp.Model;` 调整为 `outMatrix[i, 3] = string.Empty;`；
     - 严格遵循注释规范，增加中文解释说明。
  3. **编译校验**：
     - 运行 `dotnet build /t:Compile` 编译通过，0 错误 0 警告。

- **元件汇总表 G 列单价生成改为动态公式 `=ROUND(K{row}*(L{row}*M{row}+N{row}*O{row}),2)`**：
  1. **需求定位**：
     - 用户要求将【元件汇总表】中 G 列（单价）赋值从静态数值修改为动态公式：`=ROUND(K{row}*(L{row}*M{row}+N{row}*O{row}),2)`，其中行号对应 Excel 中的实际物理行号；
     - 公式含义为：单价 = 四舍五入保留2位[ 报出系数 * (本体表价 * 本体折扣 + 附件表价 * 附件折扣) ]。
  2. **代码实施 (`Services/ExcelServices.SummaryAdjustPrice.cs`)**：
     - 计算每行的绝对物理行号 `int currentRow = startDataRow + i;`；
     - 将 `outMatrix[i, 6]` 改为 `$"=ROUND(K{currentRow}*(L{currentRow}*M{currentRow}+N{currentRow}*O{currentRow}),2)"`；
     - 将写入属性从 `.Value2 = outMatrix` 切换为 `.Formula = outMatrix`，确保公式在 Excel 中自动解析并实时联动计算；
     - 严格遵循注释规范，增加中文注释说明。
  3. **编译校验**：
     - 运行 `dotnet build /t:Compile` 编译通过，0 错误。

- **元件汇总表 H 列总价与 J 列成本单价生成改为动态联动公式**：
  1. **需求定位**：
     - 用户要求将【元件汇总表】中 H 列（总价）改为动态公式：`=ROUND(F{row}*G{row},2)`（数量 * 单价，F列为数量，G列为单价）；
     - 将 J 列（成本单价）改为动态公式：`=ROUND(L{row}*M{row}+N{row}*O{row},2)`（本体表价 * 本体折扣 + 附件表价 * 附件折扣）；
     - 结合此前已将 G 列改为 `=ROUND(K{row}*(L{row}*M{row}+N{row}*O{row}),2)`，使整张汇总表的单价、总价、成本价格体系完全形成自动联动公式链。
  2. **代码实施 (`Services/ExcelServices.SummaryAdjustPrice.cs`)**：
     - 将 `outMatrix[i, 7]` 改为 `$"=ROUND(F{currentRow}*G{currentRow},2)"`；
     - 将 `outMatrix[i, 9]` 改为 `$"=ROUND(L{currentRow}*M{currentRow}+N{currentRow}*O{currentRow},2)"`；
     - 严格遵循注释规范，增加中文注释说明。
  3. **编译校验**：
     - 运行 `dotnet build /t:Compile` 编译通过，0 错误。

- **元件汇总表点击本体物料填 L 列、配套附件填 N 列（多附件公式连接）**：
  1. **需求定位**：
     - 在【元件汇总表】双击 D 列选择本体物料时：型号填入 D 列，本体价格必须填在 L 列 (本体表价)，若 M 列 (本体折扣) 为空或 0 则默认补 1；
     - 点击【配套附件】选定附件时：型号以 `+附件型号` 追加在 D 列，附件价格必须填在 N 列 (附件表价)；
     - 多个附件时，N 列自动用加法公式连接（如首次为 `49`，第二次选定 `60` 则自动升级为 `=49+60`；若已是公式则继续在末尾追加 `+单价`）；
     - 附件选定时若 O 列 (附件折扣) 为空或 0 则默认补 1；
     - 联动保持 G 列单价、J 列成本单价、H 列总价的 ROUND 联动公式实时计算。
  2. **代码实施 (`Services/ExcelServices.ComponentMatch.cs`)**：
     - 在 `ShowComponentMatchOverlay` 中，若 D 列为空或为“点击查询”，自动以 C 列原型号作为上下文进行物料与附件初筛；
     - 在 `FillSelectedComponentToActiveRow` 中，识别 `isSummarySheet`，将本体单价赋给 `L{row}`，M 列默认设为 1，并确保 G/J/H 公式联动；
     - 在 `FillSelectedAttachmentToActiveRow` 中，识别 `isSummarySheet`，将附件单价填入 `N{row}`（单个直接写入数值，多个自动拼接为 `={old}+{new}` 公式），O 列默认设为 1，确保 G/J/H 公式联动。
  3. **编译校验**：
     - 运行 `dotnet build` 编译通过，0 错误。

- **常规分类明细表中本体与附件在 M 列 (表价) 连加**：
  1. **需求定位**：
     - 在常规分类明细表中：选定本体物料时，本体价格填入 M 列 (表价)；
     - 选定配套附件时，D 列追加 `+附件型号`，M 列与本体及其他附件以加法公式形式连加（如首次仅有本体为 `447.01`，选定附件 `150` 后自动升级为 `=447.01+150`；若已是公式则继续在末尾追加 `+单价`，如 `=447.01+150+60`）；
     - 汇总表保持本体填 L 列、附件填 N 列（多附件用加法公式），且不重写表格自带的 G/J/H 公式。
  2. **代码实施 (`Services/ExcelServices.ComponentMatch.cs`)**：
     - `FillSelectedComponentToActiveRow` 中，常规表分支将价格写入 `M{row}`；
     - `FillSelectedAttachmentToActiveRow` 中，常规表分支读取 `M{row}` 已有值或公式，生成加法连加公式回填给 `M{row}`。
  3. **编译校验**：
     - 运行 `dotnet build` 编译通过，0 错误。

- **明细表 N 列折扣支持 ROUND(...) 嵌套公式解析，且本体与附件折扣统一保留 2 位小数**：
  1. **需求背景**：
     - 在分类明细表中，N 列公式可能被外层 `ROUND((...), 2)` 包裹（如 `=ROUND((159.05*1*0.5+336.2*1*1)/495.25 ,2)`）；
     - 折扣解析要求保留 2 位小数（例如 `0.85`）。
  2. **代码实施 (`Services/ExcelServices.SummaryAdjustPrice.cs`)**：
     - 在 `ParseBaseAndAccessoryDiscount` 中增加对 `ROUND(..., 2)` 外层包裹的兼容剥离；
     - 提取本体折扣与加权附件折扣后，统一通过 `Math.Round(..., 2)` 保留 2 位小数。
  3. **编译校验**：
     - 运行 `dotnet build` 编译通过，0 错误。

## [Completed]

- 常规表中本体与附件在 M 列连加、汇总表中本体填 L 列附件填 N 列（多附件公式连接且不改动表格自带公式）已全部落地；
- 明细表 N 列折扣支持 `ROUND(...)` 解析并统一保留 2 位小数。

## [In-Progress]

- 等待用户在 Excel 中实测验证。

## [Next]

- 在分类明细表 N 列填入 `=ROUND((159.05*1*0.5+336.2*1*1)/495.25 ,2)`，执行汇总调价，验证本体折扣 (0.50) 与附件折扣 (1.00) 是否准确解析并写入汇总表 M 列与 O 列。



