# Session State

## [Completed]

- **彻底修复辅材壳体计算中心 (cabinet_aux_calc.html) 界面横向溢出与组件移出屏幕故障**：
  1. **锁定容器横向滚动**：为 `.main-body` 与各卡片容器增加 `overflow-x: hidden !important; min-width: 0; width: 100%; box-sizing: border-box;`，从源头杜绝页面产生水平滚动条，彻底切断 Chromium/WebView2 焦点自动平移（`scrollIntoView`）导致左侧组件被卷出视口的路径；
  2. **网格与卡片响应式重构**：
     - 将 `.param-grid.three-col` 改造为 `repeat(auto-fit, minmax(220px, 1fr))`，在 125%/150% 等高 DPI 缩放或小屏下自适应折行，Label 宽度由 125px 微调至 105px 紧凑对齐；
     - 将 `.results-grid` 4 列卡片改为 `repeat(auto-fit, minmax(160px, 1fr))`，空间不足时平滑转为 2x2 网格；
     - 算式明细字符串增加 `word-break: break-all; white-space: normal;`，表格容器设置 `min-width: 0;` 防止撑宽父级；
  3. **编译与热同步**：
     - 代码已热同步至 `bin\Debug\net48\Resources\cabinet_aux_calc.html`；
     - `dotnet build /t:Compile` 编译通过：0 错误。

- **彻底排除主元器件（第1个元件）参与导线与分支排计算的约束修复 (`Services/ExcelServices.CabinetAuxCalc.cs`)**：
  1. **问题排查定位**：
     - 在基础特征扫描循环中，原代码对 `i = 0` 的首行主元器件也统计了极数与数量，将其极线头塞入 `currentWireMap` 中，导致主进线总开关被错误折算了一次出线导线；
     - 分支排部分原先遍历为 `i = 1` 开始，确认主元器件未参与分支排；
  2. **精准排除与单点修复**：
     - 在统计 `currentWireMap` 时增加 `if (i > 0 && !comp.IsFireTransformer && !comp.IsCurrentTransformer)` 门控约束；
     - 确保首行主进线开关（`i == 0`）与互感器穿心件绝不计入一次导线线头数，主进线开关不计算分支排也不计算导线，仅出线分路元件（`i > 0`）计算出线导线与分支排；
  3. **编译构建验证**：
     - `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误 0 警告。


- **落实“无水平排则出线分支不做排只能做线”的工程约束规则 (`Services/ExcelServices.CabinetAuxCalc.cs`, `Models/CabinetAuxCalcModels.cs`, `Resources/cabinet_aux_calc.html`)**：
  1. **出线分支排门控约束 (`Services/ExcelServices.CabinetAuxCalc.cs`)**：
     - 在分支四中引入 `hasHorizontalBus` 判定：仅当满足水平排条件（`hasHorizontalBus == true`）且存在大电流出线断路器时，才计算出线分支铜排；
     - 若未满足水平排（`hasHorizontalBus == false`），出线分支坚决不做分支排，不累加铜排重量也不生成分支排算式；
  2. **一次配线全量承接与规格扩展 (`Services/ExcelServices.CabinetAuxCalc.cs`, `Models/CabinetAuxCalcModels.cs`)**：
     - 在一次导线计算遍历 `currentWireMap` 时，将门控放行逻辑调整为 `!hasHorizontalBus || cur < rules.CopperRules.BranchMinCurrent`；
     - 当无水平排时，所有出线回路（包括大电流回路）自动全部转为一次导线配线（“只能做线”）；
     - 在 `PrimaryWireSpecTable` 中扩充 160A（BV-50）、250A（BV-70）、9999A（BV-95）大截面导线规格条目，保障大电流走线选型精准；
  3. **前端界面提示更新与热同步 (`Resources/cabinet_aux_calc.html`)**：
     - 更新铜排定额面板关于出线分支排的业务提示：“仅在满足水平排且出线额定电流 > XXA 时触发；若未满足水平排，则出线分支不做排，全部按一次导线计算”；
     - 前端模板已热同步至 `bin\Debug\net48\Resources\cabinet_aux_calc.html`；
  4. **编译构建验证**：
     - `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误 0 警告。


- **彻底修复出线分支排铜排型号直接借用水平主排单重的问题，重构为基于各出线回路额定电流独立选型与透明推导 (`Models/CabinetAuxCalcModels.cs`, `Services/ExcelServices.CabinetAuxCalc.cs`, `Resources/cabinet_aux_calc.html`)**：
  1. **问题根因定位**：
     - 原代码在扫描大电流塑壳断路器时仅统计台数，计算理论重量时直接乘以水平主排每米单重 `mainBusWeightPerMeter`；
     - 水平主排是按照主进线开关（如 400A）选型为 `TMY-40*4`（1.424 kg/m），而出线回路断路器可能是 160A、250A 等，导致出线分支排重量被成倍虚高，且算式中未能展现真实分支规格；
  2. **核心业务与计算引擎重构 (`Services/ExcelServices.CabinetAuxCalc.cs`)**：
     - 解除与水平主排的强耦合，建立按规格分组统计字典 `branchBusGroupMap`；
     - 循环出线回路时，满足大电流门限的断路器根据自身额定电流调用 `GetBusbarSpecItem` 精准匹配对应 TMY 规格与每米理论单重；
     - 依据规格分组聚合台数并汇总涉及的电流档位，按单台基准长独立计算各组理论重量并累加至 `copperWeight`；
     - 算式明细输出全面透明化，如 `出线分支排 (出线160A共3台 | TMY-25*3 | 0.668kg/m): 3台 × 1.0m × 0.668kg/m = 2.00 KG`，若有多种规格则逐行清晰分列；
  3. **配置模型扩展与界面联动 (`Models/CabinetAuxCalcModels.cs`, `Resources/cabinet_aux_calc.html`)**：
     - `CopperConfig` 扩充 `BranchBusUnitLength`（出线分支铜排单台基准展开长，单位：米，默认 1.0），标注 `--硬编码--`；
     - 铜排母线定额面板新增“分支排单台长”输入项，更新底部业务提示文字，并在 Vue `initContext` 中补充响应式默认值安全兜底；
     - 热同步最新前端模板至 `bin\Debug\net48\Resources\cabinet_aux_calc.html`；
  4. **编译构建与代码规范验证**：
     - 执行 `dotnet build /t:Compile` 编译通过：0 错误 0 警告；
     - 严格遵循新增代码每 3 行包含至少 1 行中文注释。


- **物料智能匹配悬浮窗 (component_match_overlay.html / ComponentMatchOverlayForm.cs) 必含标签快捷删除、内联编辑与模式 B 双态分流**：
  1. **痛点消除**：解决因预设型号必含规则过严导致 0 条匹配时，用户在悬浮窗内无法原地调整条件的阻碍；
  2. **交互升级**：
     - 必含标签右侧集成 `✕` 快捷删除按钮，Hover 变浅红放大，点击即时移除约束并触发重搜；
     - 双击标签文字无缝切换为 Mini Input（`inline-rule-input`），支持 Enter 确认保存、Esc 取消；
     - 管道末尾新增 `+ 必含` 按钮，点击弹出输入框支持动态追加新必含条件；
     - 键盘防冲突保护：内联编辑拦截 Enter / Esc 冒泡，防止误触发全局回填或关闭窗口；
  3. **模式 B 双态分流（临时试探 vs 持久化保存/重置）**：
     - 默认增删改仅在当前窗口运行时过滤管道生效，即时异步重搜；
     - 检测到改动时界面自动露出 `[⚡ 临时]` 标识、`[💾 保存]` 按钮（调用 C# 写入磁盘 JSON 配置文件）与 `[↺ 重置]` 按钮（还原默认规则）；
  4. **全流程验证**：
     - `dotnet build /t:Compile` 编译通过：0 错误；
     - HTML 模板已热同步至 `bin\Debug\net48\Resources\`；
     - 通过浏览器自动化子代理完成新增、双击编辑、删除、保存与重置全套交互真机驱动测试并截图/录像存档。


- **公式法调费窗口 (formula_adjust_fee.html) A 列(序号)与 B 列(元件名称)横向滚动固定锁定**：
  1. **固定列扩展**：
     - 在明细表格 `el-table` 中，为 A 列（序号）和 B 列（元件名称）添加 `fixed` 属性，B 列明确固定宽度 `width="100"`；
     - 与最左侧的行标记索引列（fixed 36px）联动，在横向拖动滚动条时，序号列与元件名称列紧随其后牢牢固定在左侧，不随 C、D、E、F 等数据列滚动；
  2. **样式与视觉层级防护**：
     - 在 CSS 中增强固定列在 `tr:hover` 悬停时的背景底色覆盖（`#f8fafc`），杜绝滚动时单元格穿帮并保持一致的高亮悬停交互；
  3. **实时热同步与浏览器自动化真机测试**：
     - 同步更新源码及 `bin\Debug\net48\` 下的 HTML 文件；
     - 通过浏览器端自动化驱动测试并截图，验证向右拖动滚动条后，行序号、A 列、B 列完全固定在左侧，C/D/E/F 列顺畅滚动隐藏在固定列后方，H/J/K 等列正常露出的预期效果。

- **彻底重构全项目所有 10 个 WebView2 窗口的拖拽架构，彻底消除导致 Excel 卡死与崩溃重启隐患 (全量 C# 窗体与前端 HTML 模板)**：
  1. **故障根因彻底明确**：
     - 原无边框窗口通过前端向 C# 发送 `dragWindow`，C# 调用 `SendMessage(WM_NCLBUTTONDOWN, HTCAPTION, 0)` 同步阻塞 API；
     - 该调用在 Windows 层面启动非客户区模态移动循环（Modal Move Loop），在松开鼠标前霸占主线程；
     - 窗口均在 Excel STA 主线程上创建，导致 Excel 整个主事件循环与 OLE 管道在拖拽期间完全停摆；
     - WebView2 的渲染子进程因 IPC 等待宿主应答超时引发挂死检测，叠加 Excel COM 保护机制（Access Violation 0xC0000005）最终使 Excel 崩溃并触发 Office 自动恢复重启。
  2. **全面落地的统一新标准架构 (PointerEvents + rAF 节流 + 非模态坐标更新)**：
     - **前端规范**：全部替换为基于 Pointer Events 的 `onHeaderPointerDown`（兼容触控与鼠标）；调用 `setPointerCapture` 保证光标不丢；引入 `window.devicePixelRatio` 精准适配高分屏 1:1 跟踪；通过 `requestAnimationFrame` 合并高频移动，向 C# 派发带物理增量 `{ action: 'moveWindow', deltaX, deltaY }`；CSS 标题栏样式补充 `touch-action: none` 阻断系统手势干扰；
     - **后端规范**：各窗体统一拦截 `case "moveWindow":`，基于 `SafeInvoke` 直接执行非阻塞的 `this.Location = new Point(this.Left + deltaX, this.Top + deltaY)`，耗时在微秒级，完全不阻塞 STA 线程，不进入任何 Win32 模态循环；保留原有的 `dragWindow` 作为向后兼容兜底；
  3. **全部 10 个窗口改造清单 (100% 达成)**：
     - ① `Forms/ComponentManageForm.cs` & `Resources/component_manage.html` (元器件数据管理)
     - ② `Forms/CategoryForm.cs` & `Resources/category.html` (新建分类)
     - ③ `Forms/CabinetAuxCalcForm.cs` & `Resources/cabinet_aux_calc.html` (辅材壳体计算)
     - ④ `Forms/ComponentGroupBuilderForm.cs` & `Resources/component_group_builder.html` (二次元件组构建)
     - ⑤ `Forms/ComponentMatchForm.cs` & `Resources/component_match_dialog.html` (物料匹配与规则)
     - ⑥ `Forms/CreateProjectForm.cs` & `Resources/create_project.html` (开始报价/新建项目)
     - ⑦ `Forms/FormulaAdjustFeeForm.cs` & `Resources/formula_adjust_fee.html` (公式法调费)
     - ⑧ `Forms/ModelParamParserForm.cs` & `Resources/model_param_parser.html` (规格型号参数提取)
     - ⑨ `Forms/SmartInputForm.cs` & `Resources/smart_input.html` (智能输入设置，新增 `SafeInvoke` 封装)
     - ⑩ `Forms/SummaryAdjustPriceForm.cs` & `Resources/summary_adjust_price.html` (汇总调价，支持配置视图与紧凑编辑条双视图平滑拖拽)
  4. **编译与同步交付验证**：
     - 全量项目编译通过：`0 个错误`；
     - 所有前端 HTML 模板已全部热同步至 `bin\Debug\net48\Resources\`。

- **彻底修复“右键切换到旧快捷键（原生菜单）时 Excel 直接卡死且进程无法结束”故障 (`Forms/CustomContextMenuForm.cs`, `ExcelEventManager.cs`)**：
  1. **故障根因诊断**：
     - **根因 A (WebView2 与 Windows 模态弹窗 IPC 死锁)**：点击“切换为 Excel 原生右键菜单”时，在 WebView2 的 `WebMessageReceived` 同步回调链条中直接调用了 `MessageBox.Show(...)` 模态弹窗。WebView2 底层 IPC 消息抽泵被 Windows 限制性模态对话框循环阻断，导致 `msedgewebview2.exe` 等待宿主确认与 Excel STA 主线程等待弹窗关闭形成跨进程死锁；
     - **根因 B (右键事件流中执行全局 CommandBars 遍历与删除)**：在 `OnSheetBeforeRightClick` 中，原生模式与自定义模式下每次右键均调用 `RemoveContextMenuControls()`。在 Excel 触发右键事件（`WM_CONTEXTMENU`）内部组装菜单的瞬间，通过 COM 循环遍历所有 200+ 个 CommandBars 并执行控件查找/删除，直接引发 Excel C++ 菜单管道重入死锁；
     - **根因 C (进程结束不了的原因)**：Excel 处于 COM RPC/ALPC 同步等待内核态时，常规任务管理器“结束任务”（先发 `WM_CLOSE`）无法唤醒挂起的线程，导致 Excel 沦为占用句柄与 XLL 文件锁的僵尸进程 (Zombie Process)；
  2. **系统级重构与彻底修复**：
     - **脱离 WebMessage 回调链**：在 `CustomContextMenuForm.cs` 中，菜单动作点击后立即隐藏/关闭浮窗，通过 `ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro` 将动作派发到 Excel 纯净宏队列中异步执行，彻底解耦 WebView2 IPC 通信与后续业务/弹窗；
     - **原生右键 0 干扰、0 延迟放行**：在 `ExcelEventManager.cs` 的 `OnSheetBeforeRightClick` 中，当处于原生右键模式时，执行 `cancel = false; return;` 直接放行，彻底移除事件中的 `RemoveContextMenuControls()` 遍历；
     - **精准清理 Cell 菜单**：将 `RemoveContextMenuControls` 改造为直接按键名安全读取 `commandBars["Cell"]` 进行单点清理，彻底废止 200+ CommandBars 的全局危险循环；
  3. **编译与打包验证**：
     - `dotnet build -c Release /p:ExcelDnaPack=true` 成功编译与打包，0 错误 0 警告，成功生成便携版加壳插件。

- **修复个人物料库在搜索下拉悬浮窗中查询不到的问题 (`Forms/ComponentMatchOverlayForm.cs`, `Services/PersonalComponentDbService.cs`, `Resources/component_match_overlay.html`, `Services/ComponentApiClient.cs`)**：
  1. **数据源路由补全 (`ComponentMatchOverlayForm.cs`)**：
     - 在候选初始加载 (`PushInitialCandidates`)、即时模糊搜索 (`searchKeyword`) 与配套附件查询 (`getAttachments`) 中增加针对 `_filterConfig.DataSource == "personal"` 的分支路由，彻底解决之前悬浮框始终请求云端 WebAPI 导致本地库无响应的问题；
     - 在 `initCandidates` 消息中透传 `dataSource` 状态至前端；
  2. **智能降级与宽容匹配检索 (`PersonalComponentDbService.cs`)**：
     - 解除对 `ComponentApiClient` 的外部依赖，内置纯静态无副作用的 `ExtractIntegerCurrent` 与 `NormalizePolesParam` 方法；
     - 额定电流与极数支持空值宽容匹配（当字段为空时自动在型号字段中匹配 `16A`、`3P` 等规格）；
     - 增加智能降级机制：当用户在搜索框中主动输入关键字时，若因 Excel 单元格中带入的名称等其他约束导致 0 命中，系统自动以“当前品牌 + 用户搜索关键字”执行宽容检索，确保用户输入如 `16`、`AC30` 时能立即展现匹配的物料列表；
     - 增加对“全部品牌 / 全部 / All”等统称标签的过滤排除，防止生成错误的 `brand = '全部品牌'` 约束；
  3. **前端视觉呈现优化 (`component_match_overlay.html`)**：
     - 在过滤管道标签栏增加数据源徽章（`[💻 个人库]` vs `[🌐 云端库]`），让用户清晰直观获知当前检索管道归属；
  4. **全流程验证**：
     - 编译 `dotnet build` 0 错误 0 警告；
     - 通过反射与多维参数实测验证：输入 `16`、品牌 `国优`、即使行内带入无关名称 `微型断路器`，均能 100% 成功命中并返回本地库中全部 6 条符合物料。

- **全新个人物料库 (SQLite 本地免安装) 与全链路双向联动升级 (`Services/PersonalComponentDbService.cs`, `ExcelAddInDemo.csproj`, `Controllers/ComponentMatchController.cs`, `Controllers/ComponentManageController.cs`, `Forms/ComponentMatchForm.cs`, `Forms/ComponentManageForm.cs`, `Services/ExcelServices.ComponentMatch.cs`, `Services/ExcelServices.ComponentManage.cs`, `Resources/component_match_dialog.html`, `Resources/component_manage.html`)**：
  1. **免安装绿色便携 SQLite 架构设计与底层实现**：
     - 引入 `Stub.System.Data.SQLite.Core.NetFramework 1.0.119.0`，在 `.csproj` 中建立 `CopyDependenciesToPublish` 目标，自动发布 `x86/x64` 原生 `SQLite.Interop.dll`，100% 绿色便携免安装；
     - 建立 `PersonalComponentDbService.cs`，存储定位 `%LocalAppData%\ExcelCTTools\data\personal_components.db`，表结构与云端 MySQL `components` 1:1 镜像对齐，支持自愈建表与复合索引自动创建；
     - 实现完整的数据层：品牌聚合统计、根据品牌取名称、模糊与必含搜索、事务批量新增、批量更新与删除。
  2. **图 1【元器件物料匹配与品牌规则设置】升级 (`component_match_dialog.html`, `ComponentMatchForm.cs`, `ComponentMatchController.cs`, `ExcelServices.ComponentMatch.cs`)**：
     - 在规则设置顶部增加【0. 物料数据源设置 (Data Source)】单选胶囊（`云端公共库` vs `本地个人库`）；
     - 切换数据源时自动动态刷新品牌聚合统计网格与对应品牌标签数量；
     - 模拟测试与选区匹配时携带当前选中的数据源，`ExecuteBatchMatchWithDb` 自动按需从本地 SQLite 或云端反查并回填；
     - 底部操作栏新增 `[📂 管理物料]` 按钮，支持一键调出元器件数据管理窗口。
  3. **图 2【元器件数据管理】升级 (`component_manage.html`, `ComponentManageForm.cs`, `ComponentManageController.cs`, `ExcelServices.ComponentManage.cs`)**：
     - 顶部增加【管理物料库源】单选切换（`云端公共库` vs `本地个人库`）；
     - 切换到本地个人库时，品牌与名称下拉列表自动绑定本地 SQLite 统计数据；
     - 选区精准更新、新增、删除操作根据所选数据源分流处理，支持在本地个人库中执行增删改并在 Excel 中回填状态与 ID。
  4. **构建与运行验证**：
     - `dotnet build` 编译 0 错误；
     - 实测验证 SQLite 数据库自动建库建表（`components` 表与 3 组复合索引）正常就绪，全链路无缝闭环。
  1. **规则4与工程算法闭环**：
     - 实现了多分类长词优先降序匹配（Maximal Match），彻底杜绝短代号吞噬长型号；
     - 实现了断路器遇漏电自动升格机制（微型断路器+漏电 ➔ 微型漏电，塑壳断路器+漏电 ➔ 塑壳漏电）；
     - 实现了短字符/单双字母安全边界保护（如阻止 16A 的 A 误判为接触器、阻止 400V 的 V 误判为浪涌，针对施耐德 Acti9 A9 系列设立专用保护）；
     - 实现了工业命名 KB0 / KBO 兼容归一化与中文品名强直通机制。
  2. **从 Excel 规则选区一键同步特征库**：
     - 公共服务层 `ImportCategoryDictFromExcelSelection` 支持纯内存二维数组直读用户在 Excel 中框选的特征表（第一行是类别名称，下方是代号）；
     - 控制器与窗体安全交互，前端一键同步并即时生效，极大简化字典维护。
  3. **Excel 批量识别与二维数组极速回填**：
     - 扩展 `ExecuteBatchModelParse` 支持名称输出列（B列）、最小/最大电流列、极数列、脱扣列的同时解析；
     - 支持“仅填空白单元格 (OnlyEmpty)”与“强制覆盖 (OverwriteAll)”策略；
     - 严格遵循规范第 7 条，纯内存二维数组批量读入写回，无 COM 卡顿。
  4. **Vue 3 + Element Plus 前端界面重构**：
     - 增加“2. 元器件名称/类别识别通道”配置卡片，提供类别胶囊切换面板与代号 Tag 池；
     - 升级沙盒实时测试预览，支持透明展示类别决策轨迹。
  5. **严格代码规范与测试验证**：
     - 新增代码每 3 行包含至少 1 行中文注释，配置硬编码标注 `--硬编码--`；
     - 编写反射实测脚本验证微断、微漏升格、塑壳漏电、接触器、施耐德 A9 微断、KB0 等所有用例全部 100% 通过。

- **严格对齐 `tmy.DrawIO` 最新流程图：未满足水平排（塑壳数不足或电流和不足）均流转至垂直母排判定 (`Services/ExcelServices.CabinetAuxCalc.cs` & `Resources/cabinet_aux_calc.html`)**：
  1. **解析最新流程图连线**：
     - 节点 6（塑壳电流和比较）的 `false` 分支连线指向节点 16（判定垂直母排）；
     - 节点 5（塑壳数量比较）的 `false` 分支连线同样指向节点 16；
     - 明确未做水平排时，只要分路总电流和 $\ge$ 门限且主开电流 $>$ 门限，即可触发垂直母排。
  2. **核心业务判定代码重构 (`Services/ExcelServices.CabinetAuxCalc.cs`)**：
     - 条件一：出线塑壳台数 $\ge$ 门限 且 塑壳电流和 $\ge$ 门限 $\implies$ 采用水平排（按 4 极塑壳数定 3/4 根）；
     - 条件二（`else if`）：未满足水平排时，若分路总电流和 $\ge$ 门限 且 主开电流 $>$ 门限 $\implies$ 采用垂直母排（按主开极数定 3/4 根）；
     - 条件三（`else`）：均不满足时主母排置空（不设水平排也不设垂直母排）；
     - 保持零地排（标配 1 根）、垂直 N 排与大电流分支排的独立核算；
     - 严格遵循每 3 行包含至少 1 行中文注释。
  3. **前端界面提示同步 (`Resources/cabinet_aux_calc.html`)**：
     - 更新铜排定额设置页底部的业务公式提示，并同步发布至 `bin\Debug\net48`；
  4. **编译验证**：
     - 执行 `dotnet build /t:Compile` 编译通过，0 错误。

- **优化铜排配置栅格排版与响应式深度合并，彻底解决输入框文字遮挡与数字不显示问题 (`Resources/cabinet_aux_calc.html`)**：
  1. **排版挤压消除**：将每个 `.param-row-item` 后面的 `.param-unit` 精简为纯单位（`台`、`A`、`mm`、`米`），把长篇业务说明剥离并统一收纳至底部的公式与业务提示微卡片中，彻底根除双列网格互相重叠覆盖遮挡输入框的问题；
  2. **响应式深度合并与默认值兜底**：在 `initContext` 中采用细粒度属性赋值保持 Vue 3 `reactive` 对象的响应式追踪，并对历史旧版本 JSON 缺少的新字段（`mccbCountThreshold: 2`, `mccbCurrentSumThreshold: 250`, `branchTotalCurrentThreshold: 300`, `mainSwitchCurrentThreshold: 250`, `fourPoleMccbThreshold: 1` 等）全面自动赋予规范默认值；
  3. **清理重复 DOM**：删除铜排定额页中残留的重复特殊元件与母排表格标签，已重新同步至 `bin\Debug\net48\Resources\cabinet_aux_calc.html`。

- **落地基于 `tmy.DrawIO` 的全新铜排制作规则引擎与配置项扩充 (`Models/CabinetAuxCalcModels.cs` & `Services/ExcelServices.CabinetAuxCalc.cs` & `Resources/cabinet_aux_calc.html`)**：
  1. **配置模型层升级 (`Models/CabinetAuxCalcModels.cs`)**：
     - 在 `CopperConfig` 中扩充全新规则门限：`MccbCountThreshold` (出线塑壳数量门限，默认 2 台)、`MccbCurrentSumThreshold` (塑壳电流和门限，默认 250A)、`BranchTotalCurrentThreshold` (分路总电流和门限，默认 300A)、`MainSwitchCurrentThreshold` (主进线开关电流门限，默认 250A)；
     - 🌟 **4极水平排判定门限**：新增 `FourPoleMccbThreshold`（默认 1 台，分路中 4 极塑壳 $\ge$ 该值采用 4 根水平排，否则 3 根）；
     - 🌟 **特殊元器件关键字列表**：新增 `SpecialComponentKeywords`（默认 `["双电源", "ATS", "火灾探测器", "火灾互感器", "电气火灾"]`，支持用户动态增删）；
     - **工艺系数与边距**：补充 `WidthDeduction` (120mm)、`HeightDeduction` (300mm)、`VerticalBaseLength` (1.2m)、`LoadExtensionRatio` (0.1m)、`LoadExtensionStepCurrent` (150A) 与 `BranchMinCurrent` (100A)。
  2. **业务计算引擎重构 (`Services/ExcelServices.CabinetAuxCalc.cs`)**：
     - **主开关与分路解耦**：首行自动识别为主进线开关提取 $I_{main}$ 与极数 $P_{main}$，后续行为分路元件；
     - **分支一（水平排 vs 垂直母排）**：
       - 若分路塑壳数 $\ge$ 门限且塑壳电流和 $\ge$ 门限 $\to$ 判定为水平排，若 4 极塑壳数 $\ge$ 门限为 4 根 $(W-\Delta W)\times 4$，否则为 3 根 $(W-\Delta W)\times 3$；
       - 若未做水平排，且分路总电流和 $\ge$ 门限且主开关电流 $>$ 门限 $\to$ 判定为垂直母排，按 $[1.2 + 0.1 \times (\sum I / 150)] \times 极数$ 计算；
     - **分支二（垂直 N 排）**：命中特殊元器件或主开关为 4 极时，自动生成 1 根 $(H - \Delta H)$ 垂直 N 排；
     - **分支三（零地排）**：标配 1 根 $(W - \Delta W)$ 零地排；
     - **分支四（出线分支排）**：出线电流 $>$ 门限的塑壳数量 $\times 1.0\text{m}$。
  3. **前端配置界面与交互完善 (`Resources/cabinet_aux_calc.html`)**：
     - 铜排定额页重构为：主母排决策门限（含 4 极塑壳门限）、尺寸扣除量与垂直排工艺系数、特殊元器件关键字标签库（支持添加/删除 Tag）、母排规格表；
     - 绿色/蓝绿色主题（`#009688`）无边框设计，实时联动推导算式展示。
  4. **代码规范**：
     - 全程遵循每 3 行包含至少 1 行中文注释，配置硬编码均规范标注 `--硬编码--`。

- **落地铜排计算公式可视化、设置即时联动与附件（双电源/互感器等）动态规则尺寸联动功能 (`Models/CabinetAuxCalcModels.cs` & `Services/ExcelServices.CabinetAuxCalc.cs` & `Resources/cabinet_aux_calc.html`)**：
  1. **模型层与规则可扩展性升级 (`Models/CabinetAuxCalcModels.cs`)**：
     - 新增 `AttachmentBusbarRule` 动态规则模型（包含关键字、适用结构、横向排数、纵向排数、固定补偿、规格选用与启用开关）；
     - `CopperConfig` 增加 `AttachmentRules` 规则库（预设倒T/I型双电源、ATS、火灾互感器等标准规则，消除静态硬编码）；
     - `CabinetCalcResult` 扩展 `CopperFormulaDetails` 明细算式列表。
  2. **业务与动态几何尺寸联动引擎 (`Services/ExcelServices.CabinetAuxCalc.cs`)**：
     - 重构铜排计算算法：主母排根据倒T型（四极补偿）或I型（三极补偿）结合有效柜宽计算；
     - 动态遍历 `AttachmentRules`：实现附件排长 $N_W \times (W - \Delta W) + N_H \times (H - \Delta H) + L_{\text{固定补偿}}$ 与箱柜宽高及排数的完全参数化联动；
     - 细化分支铜排与过渡搭接排算式，生成结构化且完全透明的推导过程字符串。
  3. **交互控制与 UI 呈现升级 (`Resources/cabinet_aux_calc.html`)**：
     - **【⚡ 铜排母线定额】配置页**：补齐柜宽/柜高扣除量，并新增【附件与特殊元器件铜排动态影响规则库】Element-Plus 动态数据表格，支持新增/删除/修改规则及一键恢复默认；
     - **【📊 智能推导与回写】视图**：在铜排结果下方直观展示【⚡ 铜排母线推导明细与尺寸联动算式】，全流程透明展开主排、动态附件排与分支排的具体算式；
     - 调参即时联动，全表推导即时重算。
  4. **编译构建**：
     - 执行 `dotnet build /t:Compile` 编译通过，0 错误，严格遵循每 3 行包含一行中文注释规范。


- **物料智能联想下拉悬浮窗全链路异步非阻塞性能优化 (`Forms/ComponentMatchOverlayForm.cs` & `Services/ComponentApiClient.cs` & `Resources/component_match_overlay.html`)**：
  1. **主线程解耦**：
     - `ComponentApiClient` 实现真正的 `SearchComponentsAsync`、`QueryComponentsAsync`、`GetAttachmentsAsync` 异步 HTTP 请求；
     - `ComponentMatchOverlayForm` 的 `searchKeyword`、`getAttachments` 及初始数据加载全部迁移到后台 `Task.Run` 工作线程，**彻底解除 UI/STA 主线程网络等待阻塞**；
  2. **请求版本与防乱序机制**：
     - 增加自增 `_searchReqCounter` 与 `_latestSearchReqId`，快速打字时自动丢弃旧请求返回，消除数据闪烁与乱序覆盖；
  3. **前端防抖与交互提速**：
     - 前端防抖调整为 120ms，界面 Loading 动画在非阻塞主线程中 60fps 平滑旋转，输入打字 0 卡顿、0 丢字；
  4. **Excel 选区联动提速**：
     - `ShowComponentMatchOverlay` 弹出悬浮窗时不再在主线程等待网络响应，瞬间弹窗并异步填充候选数据，Excel 光标移动毫无顿挫感；
  5. **编译校验**：
     - `dotnet build /t:Compile` 编译通过，0 错误。
  1. **正向生成汇总表提取校准 (`GenerateComponentSummarySheet`)**：
     - 从分类明细表提取元器件时，严格按照：W(电流)、X(极数)、Y(脱扣)、Z(附件)、AA(BlockName)、AB(BlockCategory) 列提取并写入汇总表的 T、U、V、W、X、Y 列；
  2. **反向一键更新回写校准 (`UpdateFromComponentSummarySheet`)**：
     - 从元件汇总表写回分类明细表时，修正原先写入旧列的问题，严格回写至：
       - **W 列 (索引 23)**: `Current` (额定电流)
       - **X 列 (索引 24)**: `Poles` (极数)
       - **Y 列 (索引 25)**: `trip` (脱扣方式)
       - **Z 列 (索引 26)**: `Accessory` (配套附件)
       - **AA 列 (索引 27)**: `BlockName` (图块名称 / 扩展参数1)
       - **AB 列 (索引 28)**: `BlockCategory` (图块类别 / 扩展参数2)
  3. **编译校验**：
     - 执行 `dotnet build /t:Compile` 编译通过，0 错误。
  1. **模型列映射规范重构 (`Models/CabinetAuxCalcModels.cs`)**：
     - **W 列 (第 23 列)**: 额定电流 (`Current`)
     - **X 列 (第 24 列)**: 极数 (`Poles`)
     - **Y 列 (第 25 列)**: 脱扣类型/脱扣方式 (`Trip`)（新增属性）
     - **Z 列 (第 26 列)**: 附件描述 (`Accessory`)
     - **AA 列 (第 27 列)**: 图块名称 (`BlockName`)
     - **AB 列 (第 28 列)**: 图块类别 (`BlockCategory`)
  2. **元器件矩阵批量读取引擎适配 (`Services/ExcelServices.CabinetAuxCalc.cs`)**：
     - 依据规则 7，将元器件批量读取范围从 `A:Z` 扩展至 `A:AB`（共 28 列，`ws.Range[$"A{compStartRow}:AB{compEndRow}"]`）；
     - 逐行提取时，严格按照 W(电流)、X(极数)、Y(脱扣)、Z(附件)、AA(图块名)、AB(图块类别) 列索引进行读取与赋值；
  3. **编译校验**：
     - 执行 `dotnet build /t:Compile` 编译通过，0 错误。


- **直接匹配数据库脱扣数据列 (`DrawMall.Ability` & `excel-ct-tools`)**：
  1. **服务端接口与 DTO 增强 (`DrawMall.Ability.Docking/Dto/ComponentDtos.cs` & `DrawMall.Ability/ComponentServicer.cs`)**：
     - `ComponentQueryDto` 新增 `Tripping` 脱扣方式查询属性；
     - `ComponentServicer.GetPagedListAsync` 直接在数据库层面按 `x.Tripping == query.Tripping` 进行精确匹配过滤；
  2. **客户端接口请求联动 (`Services/ComponentApiClient.cs`)**：
     - 在 `queryParams` 中添加 `Tripping={cleanTrip}`，直接向 WebAPI 传递脱扣入参，由数据库返回匹配 `tripping` 数据列的结果，彻底移除客户端本地猜测代码。

- **元器件明细行与汇总表 CAD 句柄导出/读取列由 AA 列调整为 AD 列 (第 30 列)**：
  1. **箱柜批量导出引擎 (`Services/ExcelServices.Cabinet.cs`)**：
     - 将元器件明细行的 CAD 句柄写入列调整为 `AD` 列（`sheet.Range[$"AD{compStartRow}:AD{compEndRow}"]`）；
     - 将箱柜信息行图元坐标范围写入列调整为 `AD` 列（第 30 列）；
  2. **汇总调价引擎 (`Services/ExcelServices.SummaryAdjustPrice.cs`)**：
     - 将明细表批量读取范围从 `A:AA` 扩展至 `A:AD`（30 列）；
     - 直接从 `AD` 列（索引 30）提取 CAD 句柄，不保留对旧版 `AA` 列的向下兼容；
     - 汇总表生成时将 CAD 句柄批量写入 `AD` 列；
     - 汇总调价一键更新反向同步时批量读取范围同样同步为 `A:AD`（30 列）；
  3. **选区联动与二次元件组 (`ExcelEventManager.cs` & `Services/ExcelServices.ComponentGroup.cs`)**：
     - 在 `OnSheetSelectionChange` 中直接读取 `AD` 列；
     - 在二次元件组规则插入时，直接提取与写入 `AD/AE` 列（第 30/31 列）；
  4. **编译校验**：
     - 执行 `dotnet build /t:Compile` 编译通过，0 错误。

- **重构物料智能匹配全参数回填与汇总调价双向同步引擎**：
  1. **物料匹配回填引擎 (`Services/ExcelServices.ComponentMatch.cs`)**：
     - **常规分类明细表分支**：
       - B 列 (col 2): 名称 (`item.Name`)
       - C 列 (col 3): 规格型号 (`item.Model`)
       - D 列 (col 4): 生产厂家/品牌 (`item.Brand`)
       - M 列 (col 13): 表价 (`item.Price`)
       - V 列 (col 22): 扩展参数1 (`item.Param1`)
       - W 列 (col 23): 扩展参数2 (`item.Param2`)
       - X 列 (col 24): 额定电流 (`item.Current`)
       - Y 列 (col 25): 极数 (`item.Poles`)
       - Z 列 (col 26): 附件/备注 (`item.Remark`)
       - 回填完成后自动清除 C/D 列背景底色。
     - **元件汇总表分支**：
       - B 列 (col 2): 名称 (`item.Name`)
       - D 列 (col 4): 型号 (`item.Model`)
       - I 列 (col 9): 品牌/生产厂家 (`item.Brand`)
       - L 列 (col 12): 本体表价 (`item.Price`)
       - M 列 (col 13): 本体折扣 (若空/0则自动补 1 保障公式联动)
       - P 列 (col 16): 备注 (`item.Remark`)
       - T 列 (col 20): 额定电流 (`item.Current`)
       - U 列 (col 21): 极数 (`item.Poles`)
       - V 列 (col 22): 脱扣方式 (`item.Tripping`)
       - W 列 (col 23): 附件列 (初始置空，选附件时自动回填附件脱扣方式 `attachment.Tripping`)
       - X 列 (col 24): 参数1 (`item.Param1`)
       - Y 列 (col 25): 参数2 (`item.Param2`)
       - 回填完成后自动清除 D 列背景底色。
     - **附件追加回填 (`FillSelectedAttachmentToActiveRow`)**：
       - 常规分类表中自适应更新 C 列（型号+附件型号\*数量）与 M 列表价连加公式；
       - 元件汇总表中更新 D 列、N 列附件表价加法公式，并**将附件脱扣方式回填/累加至 W 列**。
  2. **汇总调价双向同步引擎 (`Services/ExcelServices.SummaryAdjustPrice.cs`)**：
     - **生成汇总表**：
       - 读取明细表 A~AA 27 列范围，提取扩展参数 V(参数1)、W(参数2)、X(额定电流)、Y(极数)、AA(CAD句柄)；
       - 在汇总表渲染 T4:Y5 扩展参数表头（T: 额定电流, U: 极数, V: 脱扣方式, W: 附件, X: 参数1, Y: 参数2）；
       - 矩阵一次性批量写入 T~Y 列参数数据及 AA 列 CAD 句柄；
       - 精确设置 T~Y 列宽与 A5:Y5 自动筛选。
     - **一键更新反向同步**：
       - 读取汇总表 A~Y 25 列数据矩阵（包含 T: 电流, U: 极数, V: 脱扣, W: 附件, X: 参数1, Y: 参数2）；
       - 将修改后的参数反写回各分类明细表对应行的 V 列(参数1)、W 列(参数2)、X 列(额定电流)、Y 列(极数)。
  3. **代码规范与验证**：
     - 严格遵循每 3 行包含一行中文注释规范；
     - 遵循最小修改原则，读写大区域遵循规则 7 数组批量进出内存。

- **落地【智能辅材与壳体计算】及动态定额规则配置中心功能**：
  1. **模型与配置持久层 (`Models/CabinetAuxCalcModels.cs`)**：
     - 定义 `QuotationRules` 聚合模型（`GeneralConfig`, `ShellConfig`, `CopperConfig`, `AuxConfig`, `LaborConfig`）；
     - 实现规则持久化至 `quotation_rules.json`（支持铜价、加点系数、综合税率、壳体匹配名称、标准尺寸库、接线空间阶梯、二次元件定额表等自由配置）；
     - 根据实际表格结构对齐元器件列字段映射：B 列(名称)、C 列(型号)、F 列(数量)、V 列(图块类别)、W 列(图块名称)、X 列(额定电流)、Y 列(极数)、Z 列(附件描述)；
     - **新增结构化一次导线长度计算配置模型 (`PrimaryWireLengthConfig`)**：包含基础垂直预留高度(130mm)、火灾互感器增量(100mm)、普通互感器增量(130mm)、落地柜门限(1600mm)、柜宽系数(0.7)、裕量放大系数(1.1)与配电箱宽系数(0.6)。
  2. **核心业务与计算引擎 (`Services/ExcelServices.CabinetAuxCalc.cs`)**：
     - 遵循规则 7，采用 2D 数组一次性批量读入箱柜元器件有效数据（读取范围扩展至 `A{compStartRow}:Z{compEndRow}` 覆盖 26 列）；
     - **全面接入 `WireLengthConfig` 动态推导一次线长度与用量**：自动根据箱高判断落地柜或配电箱，结合互感器加成折算各规格总米数并差异化计价；
     - 完整复原原 VBA 算法：CAD/本地特征库双轨降级、容积率安全系数选型、倒T/I型主母排+ATS双电源+通长N排+分支铜排计算、一次导线+二次元件辅材定额、平铺面积+二次装配人工费；
     - **严格落地用户指定的壳体回写规则**：优先扫描计费区（`Cab_Subsum` 到 `Cab_Tolsum-1`）B 列匹配 `ShellMatchName`（默认“箱体”），命中则写入 C 列；未命中则写入 `Cab_Det` 的 C 列并将 B 列设为匹配名称；
     - 铜排若 $>0$ 自动插入/更新 TMY 数量公式与总价，辅材与人工费动态公式写入，确保全表联动重算。
  3. **交互控制与 UI 视觉深度重构 (`Forms/CabinetAuxCalcForm.cs`, `Resources/cabinet_aux_calc.html`)**：
     - 将窗体尺寸调整为 `960x720` 像素，保证视觉空间宽裕舒展；
     - **全面去除所有按钮、卡片、窗口阴影 (`box-shadow: none !important`)**，呈现纯净现代工程扁平风；
     - **定额规则二级 Tabs 彻底解耦独立**：
       - **Tab 1: 🔌 一次配线定额**（一次导线长度折算参数 + 图 3 一次导线规格选型与单价表）；
       - **Tab 2: 🧩 二次元件定额**（图 4 二次元件接线与工价定额表单独作为独立 Tab 页，全宽展开并支持添加）；
       - **Tab 3: 📦 壳体选型规则**（壳体匹配名称、安全系数、标准尺寸库）；
       - **Tab 4: ⚡ 铜排母线定额**（母排结构门限、补偿参数）；
       - **Tab 5: 🛠️ 费率与结构补贴**（基础辅材与补贴、平铺制作工价、税率乘数）；
     - 使用严谨的双列 Grid 网格布局重构参数表单，固定 Label 最小宽度（125px），彻底解决文字与输入框拥挤重叠问题。
  4. **交互控制与 UI 呈现 (`Controllers/CabinetAuxCalcController.cs`, `Forms/CabinetAuxCalcForm.cs`, `Resources/cabinet_aux_calc.html`)**：
     - 基于 WebView2 + Vue 3 `<script setup>` + Element Plus 框架构建，主色调 `#009688`；
     - 提供【📊 智能推导与回写】与【⚙️ 规则与定额配置中心】双视图；
     - 在 Ribbon 功能区【③调价格→】增加【辅材壳体计算】大按钮，并在右键菜单中挂载直达入口。
  5. **编译构建**：
     - 执行 `dotnet build` 编译通过，0 错误，严格遵循每 3 行包含一行中文注释规范。

- **修复元件明细表 M 列（表价/面价）拆分本体与附件价格时附件乘积公式无法解析的问题**：
  1. **定位根本原因**：
     - 原 `ParseBaseAndAccessoryPrice` 方法在按 `+` 拆分出各个加项后，对第二项及后续附件项直接使用 `decimal.TryParse` 解析；
     - 当公式形如 `=159.05+336.2*2` 时，附件项为 `"336.2*2"`（包含乘号），`decimal.TryParse` 失败返回 `false`，导致附件价格解析结果为 `0`；
     - 另外原逻辑未对外层 `ROUND(...)` 和圆括号进行脱壳处理，当存在 `ROUND` 包裹时会导致所有项均解析失败。
  2. **修复方案 (`Services/ExcelServices.SummaryAdjustPrice.cs`)**：
     - 增加 `EvaluateTermValue` 辅助方法，支持对单项表达式进行乘法运算（如 `"336.2*2"` 计算为 `672.4`，`"159.05*1"` 计算为 `159.05`）；
     - 在 `ParseBaseAndAccessoryPrice` 中增加对 `ROUND(...)` 与外层圆括号的脱壳处理；
     - 第一项（本体项）与后续项（附件项）均通过 `EvaluateTermValue` 精确计算并累加，确保本体表价与附件总表价均能准确解析。

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
     - 用户要求将【元件汇总表】中 H 列（总价）改为动态公式：`=ROUND(F{row}*G{row},2)`（数量 \* 单价，F列为数量，G列为单价）；
     - 将 J 列（成本单价）改为动态公式：`=ROUND(L{row}*M{row}+N{row}*O{row},2)`（本体表价 _ 本体折扣 + 附件表价 _ 附件折扣）；
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

- **落地配套附件数量入口与型号(+型号*数量)/价格(+单价*数量)公式联动回填**：
  1. **需求背景与目标**：
     - 在物料智能联想下拉窗的【配套附件】模式中，支持用户为断路器等主体添加多个同类附件（如 2 个辅助触头 `OF*2`、2 个分励脱扣器 `MX*2` 等）；
     - 要求提供清晰直观的数量入口，选定附件后实现型号（`+附件型号*数量`）与价格公式（`+单价*数量`）的精准联动回填。
  2. **前端界面与交互 (`Resources/component_match_overlay.html`)**：
     - 在配套附件模式管道栏右侧增加精致的步进数量控制器 `数量: [ - ] [ 1 ] [ + ]`（支持输入与点击步进，范围 1~99，默认 1）；
     - 数量 > 1 时，候选卡片型号右侧呈现 `×数量` 提示徽标，价格实时计算并展示乘积小计（如 `(¥336.20 × 2) ¥672.40`）；
     - 底部指引更新为 `[点击/回车] 自动拼接 +附件*数量 并累加单价公式`；
     - 选定附件派发 `selectAttachment` 时携带 `quantity` 数量参数。
  3. **C# 控制与业务层 (`Forms/ComponentMatchOverlayForm.cs`, `Services/ExcelServices.ComponentMatch.cs`)**：
     - `ComponentMatchOverlayForm` 解析提取 Web 消息中的 `quantity` 参数并传给服务层；
     - `ExcelServices.FillSelectedAttachmentToActiveRow` 扩展 `int quantity = 1` 支持：
       - D 列型号拼接：`quantity > 1` 时格式化为 `+附件型号*数量`（如 `+MX*2`），`quantity == 1` 时为 `+附件型号`；
       - 价格公式：`quantity > 1` 时生成 `+单价*数量`（如 `+336.2*2`），常规表写入 M 列（如 `=159.05+336.2*2`），汇总表写入 N 列（如 `=336.2*2` 或 `=150+80*2`）；
       - 与 `SummaryAdjustPrice.cs` 中的 `EvaluateTermValue` 与 `ParseBaseAndAccessoryPrice` 完美闭环兼容。
  4. **编译校验**：
     - 运行 `dotnet build` 编译通过，0 错误。

- **落地 Excel 选中行联动 AutoCAD 夹点显示（单选/多选行支持 + 自动缩放视野 + 50ms 防抖 + Ribbon 切换开关）**：
  1. **需求背景与目标**：
     - 用户在 Excel 明细表或分类表中选中单行或多选连续/跨行时，自动提取所有覆盖行的 AA 列（第 27 列）CAD 文字句柄，通过轻量级命名管道向 AutoCAD 发送即时夹点高亮与自动对焦缩放通知；
  2. **客户端与防抖控制 (`Services/CadSyncClient.cs`)**：
     - 提供全局 `SyncToCadEnabled` 联动开关与 `AutoZoomEnabled` 自动视角缩放开关；
     - 内置 50ms `System.Threading.Timer` 防抖调度器，在快速按键切换行时避免管道拥堵；
     - 异步非阻塞发送至 `CadExcelHandleSyncPipe` 管道，超时 50ms 即焚，CAD 未开启时静默忽略不卡顿 Excel。
  3. **选区事件捕获 (`ExcelEventManager.cs`)**：
     - 在 `OnSheetSelectionChange` 中检测选区是否覆盖 C 列（或整行多选）：
     - 批量循环遍历所选所有行，提取 AA 列全部句柄去重合并，并携带 `autoZoom=true` 调用 `CadSyncClient.SendHandlesDebounced` 发送。
  4. **Ribbon 功能区切换按钮 (`RibbonController.cs`)**：
     - 在辅助项分组中增加【⚡ 联动CAD】（`btnToggleCadSync`）大图标切换按钮，支持用户随时一键启停联动。

- **汇总调价时提取首个元器件句柄至“元件汇总表” AA 列，并支持点击 C 列在 AutoCAD 中即时夹点高亮显示**：
  1. **需求背景与目标**：
     - 在执行“汇总调价”并生成“元件汇总表”时，打通汇总调价行与 AutoCAD 原图元之间的关联映射；
     - 提取每个分组内首个有效元器件的 CAD 句柄并保存至“元件汇总表”的 AA 列（第 27 列）；
     - 当工程师在“元件汇总表”中点击 C 列（原型号规格列）时，AutoCAD 自动平移缩放至对应图元并高亮显示原生蓝色夹点。
  2. **代码实施 (`Services/ExcelServices.SummaryAdjustPrice.cs` & `ExcelEventManager.cs`)**：
     - **聚合模型扩展**：在 `AggregatedComponent` 类中新增 `Handle` 属性；
     - **读取范围扩展至 AA 列**：在 `GenerateComponentSummarySheet` 中将二维数组批量读取范围扩展至 `A{compStartRow}:AA{compEndRow}`（覆盖 27 列），并在逐行解析时提取 `valMatrix[r, 27]` 存入 `AggregatedComponent.Handle`；
     - **提取首个有效句柄**：在 `GroupBy` 分组聚合返回实体时，通过 `g.Select(x => x.Handle).FirstOrDefault(h => !string.IsNullOrWhiteSpace(h))` 提取组内首个有效非空句柄；
     - **批量写入汇总表 AA 列**：在写入 18 列数据矩阵后，构造 `handleMatrix` 并通过 `Range["AA...:AA..."].Value2 = handleMatrix` 一次性批量写入汇总表 AA 列；
     - **C 列点击体验优化**：在 `ExcelEventManager.OnSheetSelectionChange` 中，当在“元件汇总表”中选中 C 列时，保持向 CAD 命名管道推送 AA 列句柄以激活夹点高亮与自动缩放，同时隐藏输入覆盖框，避免遮挡和影响用户查看图元。

- **落地汇总调价【一键更新】反向同步更新至各分类表箱柜明细功能**：
  1. **需求背景与目标**：
     - 用户在【元件汇总表】中对各物料进行重新选型、添加配套附件、调整厂家品牌、报出系数、本体与附件价格及折扣后，点击【一键更新】按钮，将调价结果全自动精准反向同步更新至各个分类工作表中对应的箱柜明细区域中，并联动刷新整张报价表的公式计算体系。
  2. **代码实施与架构设计 (`Services/ExcelServices.SummaryAdjustPrice.cs` & `Controllers/SummaryAdjustPriceController.cs` & `Resources/summary_adjust_price.html`)**：
     - **数据提取与规则模型 (`SummaryAdjustItem`)**：从“元件汇总表”第 6 行起批量读入（A6:R{maxRow} 及 AA 列），提取元件名称、原型号规格、修改后的型号规格、厂家、报出系数、本体/附件表价与折扣、备注及 U 列原始型号；
     - **智能公式构造与加权折扣计算**：
       - 若包含附件表价（N 列 > 0 或为公式）：M 列自动组装为本体与附件连加公式（如 `=159.05+336.2*2`），N 列按本体与附件金额加权计算综合折扣并保留 2 位小数；
       - 若无附件：M 列写入本体表价（数值或公式），N 列写入本体折扣；
     - **全工作簿箱柜明细精准匹配与批量回写 (规则 6 & 规则 7)**：
       - 遍历所有有效分类工作表，通过 `Tool.GetSheetValidCabinets` 定位箱柜元器件有效区域（`Cab_Det + 2` 至 `Cab_Subsum - 1`）；
       - 采用 2D 数组一次性批量读入内存（覆盖 A~AA 共 27 列）；
       - 依据 `名称 + 原型号`（或 U 列原始型号）精准匹配对应行，批量更新 C 列（型号）、D 列（厂家）、I 列（备注）、L 列（系数）、M 列（表价公式/数值）、N 列（折扣）；
       - 若箱柜有更新，通过 2D 数组 `.Formula = matrix` 一次性批量写回工作表；
     - **联动机制保障**：分类表明细行的销售单价(G)、销售总价(H)、成本单价(J)、成本总价(K)均由原自适应公式自动重算，箱柜小计(`Cab_Subsum`)、计费区域与顶部箱柜汇总行(`Cab_Sum`)全链路自动联动刷新；
     - **前端交互与统计反馈**：更新完成后向前端回发包含更新分类表数、箱柜数及明细项数的统计报文，并通过 Element Plus 弹窗友好提示成功信息。
  3. **编译校验**：
     - 运行 `dotnet build /t:Compile` 编译通过，0 错误，严格遵循每 3 行包含一行中文注释规范。

- **升级汇总调价【一键更新】中 N 列折扣公式格式为 `=ROUND((本体*折扣+附件*折扣)/总表价,2)` 样式**：
  1. **需求定位**：
     - 用户要求在一键更新回写分类明细表时，N 列（折扣）更新为形如 `=ROUND((159.05*0.5+736.2*1)/831.45,2)` 标准动态公式样式；
  2. **代码实施 (`Services/ExcelServices.SummaryAdjustPrice.cs`)**：
     - 在 `UpdateFromComponentSummarySheet` 中，当存在配套附件时（N 列 > 0 或为公式）：
       - 提取本体项 `baseTerm = $"{baseExpr}*{baseDiscStr}"`（如 `159.05*0.5`）；
       - 提取附件项 `accTerm = $"{accExpr}*{accDiscStr}"`（支持多附件拆分加项，如 `736.2*1` 或 `336.2*2*1`）；
       - 计算总表价分母 `totalListPrice = basePrice + accPrice`（如 `831.45`）；
       - 组合生成标准动态公式：`nColContent = $"=ROUND(({baseTerm}+{accTerm})/{totalListStr},2)"`；
     - 与既有 `ParseBaseAndAccessoryDiscount` 解析逻辑完全对称兼容闭环。
  3. **编译校验**：
     - 运行 `dotnet build /t:Compile` 编译通过，0 警告，0 错误。

## [Completed]

- 修复汇总调价【一键更新】中因公式字符串插值缺失 `$` 前缀导致 Excel COM 抛出“内存不足”的底层异常；
- 汇总调价【一键更新】未勾选名称时覆盖 B 列名称，且无论何种条件均更新 E 列单位已全部落地；
- 汇总调价【一键更新】基于汇总前配置（工作表范围过滤 + 多维合并条件精准对齐 + AB1 单元格配置持久化）已全部落地；
- 汇总调价“合并条件”中“名称”复选框已改为可自由勾选与取消，后端 GroupBy 动态响应；
- 汇总调价【一键更新】中 N 列折扣生成 `=ROUND((本体*折扣+附件*折扣)/总表价,2)` 样式动态公式已全部落地；
- 汇总调价【一键更新】反向精准同步回写至各分类表箱柜明细并联动全表公式重算已全部落地；
- **二次图回路方案与 BOM 数据管理系统落地实施 (全面闭环)**：
  1. **实体建模与持久化升级**：
     - 在 `Models/SecondaryCircuitModels.cs` 中实现 `SecondarySchemeEntity` 与 `SecondaryBomItem`，包含同配置回路代号列表（`ApplicableCodes`）、`CadDrawingName`（CAD 图名）、跨门数、开孔、人工费及动态驱动计算属性（材料费、综合总费用）；
     - 在 `PersonalComponentDbService.cs` 中自愈增加 `secondary_circuit_schemes` 关系表与多维复合索引；并在 `PersonalComponentDbService.SecondaryCircuit.cs` 中实现 CRUD、按代号/CAD图名检索、事务批量导入及物料库最新单价动态对齐机制；
  2. **Excel 批量智能解析导入**：
     - 在 `ExcelServices.SecondaryCircuit.cs` 中实现 `ImportSecondarySchemesFromActiveSheet`，严格遵循规则 3 与规则 7，通过 2D 内存数组一次性扫描当前工作表，智能识别主方案行（拆分同配置逗号代号）与子 BOM 明细行；
  3. **前端控制与 Vue 3 工作台**：
     - 创建 `Controllers/SecondaryCircuitController.cs` 与 `Forms/SecondaryCircuitForm.cs`，落实 Win32 物理按键检测防幽灵鼠标死锁；
     - 创建 `Resources/secondary_circuit_manage.html`，遵循项目 `#009688` 扁平设计、`<script setup>` 结构，具备子 BOM 展开行、回路代号 Tag 动态维护、以及从本地个人物料库极速选型的滑出抽屉；
  4. **菜单集成与构建验证**：
     - 在 Ribbon XML 中挂载【二次方案库】大图标按钮并绑定打开回调；
     - 运行 `dotnet build` 编译打包通过，0 警告，0 错误，严格遵循每 3 行至少 1 行中文注释规范。

## [Completed]

- 修复汇总调价【一键更新】中因公式字符串插值缺失 `$` 前缀导致 Excel COM 抛出“内存不足”的底层异常；
- 汇总调价【一键更新】未勾选名称时覆盖 B 列名称，且无论何种条件均更新 E 列单位已全部落地；
- 汇总调价【一键更新】基于汇总前配置（工作表范围过滤 + 多维合并条件精准对齐 + AB1 单元格配置持久化）已全部落地；
- 汇总调价“合并条件”中“名称”复选框已改为可自由勾选与取消，后端 GroupBy 动态响应；
- 汇总调价【一键更新】中 N 列折扣生成 `=ROUND((本体*折扣+附件*折扣)/总表价,2)` 样式动态公式已全部落地；
- 汇总调价【一键更新】反向精准同步回写至各分类表箱柜明细并联动全表公式重算已全部落地；
- 修复 UpdateFromComponentSummarySheet 中 buildConfig 与 activeWb 变量未定义的编译错误；
- **二次图回路方案与 BOM 数据管理系统（SQLite持久化 + 本地物料库强一致绑定 + CAD图名关联 + 同配置多回路映射 + Excel批量导入 + Vue3管理工作台）全链路落地闭环**；
- **二次方案管理工作台 5 项细节调优全面落地**：
  1. 彻底修复主界面右上角关闭按钮失效问题（切断拖拽事件冒泡防 Win32 消息循环吞噬，双向支持 `close` 和 `closeWindow`）；
  2. “所属二次组”全面更名为“二次排布图”（覆盖主表列头、工具栏下拉筛选、弹窗表单）；
  3. 彻底去除图二内容（从编辑弹窗和主表格中彻底移除 CAD 图名/图号维护项，紧凑排布）；
  4. “二次线跨门”全面修改为“二次线跨门根数”（步进器强制整数，表格及导入均四舍五入为整数根数）；
  5. 二次 BOM 子物料清单受控高度自适应滚动修复（启用 Element Plus 原生 `max-height="220"` 属性，超出行数时自动出现垂直滚动条且表头固定吸顶，弹窗设置 `max-height: 72vh; overflow-y: auto;` 杜绝遮挡底部按钮）；
  6. 物料库检索抽屉支持品牌搜索，默认指定“二次元件”（抽屉顶部集成品牌筛选/输入下拉框，默认二次元件，支持自动带出全部物料库已有品牌，且支持首屏与空关键字直接按品牌拉取）。

## [Completed]

- 修复汇总调价【一键更新】中因公式字符串插值缺失 `$` 前缀导致 Excel COM 抛出“内存不足”的底层异常；
- 汇总调价【一键更新】未勾选名称时覆盖 B 列名称，且无论何种条件均更新 E 列单位已全部落地；
- 汇总调价【一键更新】基于汇总前配置（工作表范围过滤 + 多维合并条件精准对齐 + AB1 单元格配置持久化）已全部落地；
- 汇总调价“合并条件”中“名称”复选框已改为可自由勾选与取消，后端 GroupBy 动态响应；
- 汇总调价【一键更新】中 N 列折扣生成 `=ROUND((本体*折扣+附件*折扣)/总表价,2)` 样式动态公式已全部落地；
- 汇总调价【一键更新】反向精准同步回写至各分类表箱柜明细并联动全表公式重算已全部落地；
- 修复 UpdateFromComponentSummarySheet 中 buildConfig 与 activeWb 变量未定义的编译错误；
- 二次图回路方案与 BOM 数据管理系统全链路落地闭环；
- 主界面关闭按钮修复、二次排布图更名、移除CAD图名项、跨门根数整数化、BOM垂直滚动条自适应等 5 项反馈全面解决；
- 本地个人物料库检索抽屉支持品牌检索，默认选中“二次元件”，支持全部品牌切换与型号/名称组合搜索。

## [In-Progress]

- 监听用户在实际 Excel 环境与图纸交互中的体验反馈与调优需求。

## [Next]

- 协助用户进行实际 Excel 二次图表格的导入验证，或根据图纸识别联动需求进一步打通与《CabinetAuxCalc》辅材中心的跨门线及开孔工费一键应用。

