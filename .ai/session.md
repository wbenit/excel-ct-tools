# 一、 商业加壳与代码安全注意事项 🔒

## Ribbon 回调必须保留名称（Exclude Obfuscation）

- **规则**：Excel 菜单通过 XML 字符串（如 onAction="MyMethod"）利用反射机制调用 C# 方法。
- **注意**：在 RibbonController.cs 中新增的 Ribbon 回调函数，必须确保该类或方法拥有 `[Obfuscation(Exclude = true)]` 标记，避免加壳重命名导致找不到方法。
- **最佳实践**：将界面回调集中放在 RibbonController.cs，而将业务逻辑解耦在单独类中。

---

# 二、 启发式经验与踩坑结晶 💡

1. **Excel Ribbon CustomUI 严苛 Schema 防隐藏**：
   - 使用不存在的 `imageMso` 或空的 `group label=''` 会导致 Excel 静默抛弃整个自定义 Ribbon Tab。需使用全版本通用的标准 imageMso 名称。

2. **Excel-DNA 安全弹窗与防闪退机制**：
   - 禁止使用独立线程 `Application.Run()` 弹出对话框，子线程未捕获异常会导致 Excel 进程闪退并触发 COM 硬禁用报错（`HRESULT: 0x800A03EC`）。
   - 正确方式：获取 `ExcelDnaUtil.WindowHandle` 模态附着主窗口弹出 `form.ShowDialog(new ExcelWin32Window(excelHwnd))`。

3. **Excel-DNA Modeless 非模态窗体与线程安全机制**：
   - 保持弹窗时 Excel 可自由编辑：使用 `ShowModelessForm`，调用 `form.Show(new ExcelWin32Window(excelHwnd))` 挂载 Owner 为 Excel 主句柄，既避免窗口下沉，又保证 Excel 保持非阻塞可编辑状态。
   - 单例生命周期复用：维护窗体引用，已打开时调用 `BringToFront()` / `Activate()` 避免重复多实例。
   - **CoreWebView2 必须在 UI 线程访问**：异步方法（如 `await Task.Run(...)`）在非模态环境下恢复时可能处于 Worker 线程，所有 `PostWebMessageAsJson` 与 WinForms 控件交互必须通过 `SafeInvoke`（`this.Invoke` / `this.BeginInvoke`）切换至 UI 线程，防止 `CoreWebView2 can only be accessed from the UI thread` 异常。

4. **Excel-DNA 单文件打包 (-packed.xll) 依赖包含**：
   - NuGet 依赖（WebView2, System.Text.Json 等）须在 `.dna` 文件中声明 `<Reference Path="...dll" Pack="true" />` 方可打入 packed.xll 中。
   - 前端资源需在 csproj 配置 `CopyDependenciesToPublish` Target 目标发布，C# 层配合多级备用路径检索。

5. **原生 HTML 内嵌 Vue 模板中禁用 `/>` 自闭合语法（防按钮 DOM 被吞）**：
   - 在 HTML 文件挂载 Vue 3 时，浏览器 DOM 解析器会将 `<el-input ... />` 中的 `/>` 忽略，把紧随其后的同级 `<el-button>` 误解析为 `<el-input>` 的子元素，导致 Element Plus 渲染时直接抛弃该按钮节点。
   - **规则**：所有 Vue 自定义组件在 HTML 页面中必须使用显式闭合标签（如 `<el-input></el-input>`）。

6. **Excel COM 定义名称（Defined Names）与模板句柄生命周期**：
   - 从模板工作簿复制插入内容后需**立即关闭 `templateWb` 句柄**，防止名称误挂载在只读模板上被静默销毁。
   - `Names.Add` 的 `RefersTo` 参数必须传入以 `=` 开头的标准公式字符串（如 `$"='{sheetName}'!$A${row}"`）。

---

# 三、 最新进度与状态记录 🚀

- **已完成架构与功能**：
  1. **Ribbon 工具栏**：实现了全套“我的项目”、“①建项目→”、“②录元件”、“③调价格→”、“④出报表”等分组与按钮菜单。
  2. **全局 UI 主题规范**：根据规则要求统一升级所有 Web 窗口为**绿蓝相间主题，主色调为 `#009688`**。
  3. **“我的企业设置”窗口**：基于 C# + WebView2 + Vue 3 (Element Plus, `<script setup>`)，反显解析正常。
  4. **“新建项目”窗口与控制器**：支持项目创建、 SaveAs 模板清洗、 Win32 强力置顶与视口最大化。
  5. **“新建箱柜”功能与定义名称/超链接联动**。
  6. **分类工作表标准行号智能动态探测（彻底杜绝硬编码行号 +1 偏移与 #REF! 异常）**。
  7. **“智能输入模式”与选择表联动（元器件去重词库与属性联动回填）**。
  8. **元器件行自适应空行判断公式矩阵升级（覆盖 F/G/H/J/K/L/N/Q 列）**。
  9. **新建箱柜行数与元器件行数完全基于 CabDetRowIndex 与 CabTolsumRowIndex 动态推导**。
  10. **实现 Ribbon 工具栏【删除箱柜】（btnDeleteCabinet）轻量精准删除功能**。
  11. **抽取 Excel 活动环境与有效箱柜扫描公共方法至 Tool.cs（消除多处重复冗余代码）**。
  12. **抽取光标智能匹配所属箱柜算法至 Tool.cs（FindCabinetByRow 与 GetActiveCabinet）**。
  13. **增强 CollectAllDefinedNames 与 GetSheetValidCabinets 空值智能补齐重建机制**。
  14. **彻底杜绝 DLR 动态解析（Dynamic Call）传播导致可空结构体（Nullable<KeyValuePair>）操作符异常**。
  15. **顶部汇总与选区多选箱柜智能识别及自底向上批量删除机制（已落地）**：
      - **业务与架构**：
        1. 在 [Tool.cs](file:///e:/Ace/excel-ct-tools/Tool.cs) 中新增 `GetSelectedCabinets` 公共方法，遍历 `Selection.Areas` 提取所有命中的行号并通过 `FindCabinetByRow` 自动去重与升序排列返回；
        2. 重构 `Tool.GetActiveCabinet` 统一复用 `GetSelectedCabinets`；
        3. 在 [Services/ExcelServices.Cabinet.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.Cabinet.cs) 中升级 `DeleteCabinetFromSelection` 交互层，支持单箱柜与多选箱柜的批量确认提示；
        4. 在 [Services/ExcelServices.Cabinet.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.Cabinet.cs) 中实现 `DeleteCabinets` 核心方法，采用严格的自底向上（从大到小）物理删除算法（先删明细块、再删汇总行），彻底杜绝删除行导致上方行号偏移的问题，并提供全表箱柜全选删除时的初始空白骨架重置保护与 4 个定义名称安全清理。
  16. **每个箱柜明细块自带 3 行报价人信息体系重构（已落地）**：
      - **业务与架构**：
        1. 确立每个箱柜明细块从大标题（`Cab_Det - 3`）至总计行下方 3 行报价人信息（`Cab_Tolsum + 3`）的完整物理闭环；
        2. 在 [Services/ExcelServices.Cabinet.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Services/ExcelServices.Cabinet.cs) 的 `CreateNewCabinet` 中将 `lastDetBlockEnd` 精准计算为 `maxTolsumRow + 3`，模板明细块复制源完整覆盖至 `templateTolsumRow + 3`，新总计行精准推导为 `newDetailStartRow + (templateTolsumRow - templateStartRow)`，使每个新创建的箱柜均完整自带专属 3 行报价人信息；
        3. 强化顶部汇总表插行判定（结合 A、B 列），并在插行触发时对下方明细行进行 `+1` 偏移补偿；
        4. 同步升级 [Services/ExcelServices.Cabinet.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Services/ExcelServices.Cabinet.cs) 的 `DeleteCabinets`（删除区间为 `detRow - 3` 至 `tolsumRow + 3`）及 [Tool.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Tool.cs) 的 `FindCabinetByRow`（命中区间扩展至 `tolR + 3`）；
        5. 修正汇总表 G 列单价公式为标准指向总计行的 `$"=H{newTolsumRow}"`。
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。
  17. **新建箱柜动态计费项自适应与多余元器件行智能删除清洗机制（已落地）**：
      - **业务与架构**：
        1. 在 [Services/ExcelServices.Cabinet.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Services/ExcelServices.Cabinet.cs) 的 `CreateNewCabinet` 中，动态从被复制的源箱柜中提取计费区域跨度 `feeSpan = srcTolsumRow - srcSubsumRow + 1`；
        2. 基于 [appsettings.json](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/appsettings.json) 的 `CabDetRowIndex` 与 `CabTolsumRowIndex` 计算标准总跨度（默认 28 行），动态推导标准元器件行数 `standardCompRows = standardTotalSpan - 2 - feeSpan`；
        3. 复制源箱柜完整明细块后，计算超出标准行数的多余行 `extraRows = srcCompRows - standardCompRows`，在元器件区域末尾执行物理向上删除（`Delete xlShiftUp`），下方小计、总计及报价人行号同步向上偏移 `extraRows`；
        4. 对修剪至标准规格的元器件区域调用 `Tool.BuildComponentRowsMatrix` 一次性批量写回纯净的自适应空行公式矩阵（A~Q 列），彻底杜绝复制产生的臃肿空白行并完成图号/备注等表头旧数据清洗。
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。
  18. **FindStandardCategoryRowIndexes 极简智能探测升级（支持任意箱柜序号 K 与未覆盖回退末位箱柜）**：
      - **业务与架构**：
        1. 扩展 [Tool.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Tool.cs) 中的 `FindStandardCategoryRowIndexes(dynamic sheet, int cabinetK = 1)` 方法；
        2. 彻底复用底层成熟的 `GetSheetValidCabinets` 强类型集合（内置双作用域扫描与智能重建），优先匹配目标 `cabinetK`；
        3. 若目标 `cabinetK` 未完整覆盖（如新建箱柜时该序号尚未生成），直接回退取表中最后一个有效箱柜（`validCabinets.Last()`）的物理行号分布；
        4. 彻底移除冗余的 80+ 行 `UsedRange` 二维文本扫描块，在未识别到有效箱柜时直接返回配置基准默认值；
        5. 保持 `cabinetK = 1` 默认参数，无缝向下兼容原有单参调用。
  19. **CreateNewCabinet 复用 Tool.GetActiveCabinet 智能获取当前选中箱柜**：
      - **业务与架构**：
        1. 在 [Services/ExcelServices.Cabinet.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Services/ExcelServices.Cabinet.cs) 的 `CreateNewCabinet` 中，调用 `Tool.GetActiveCabinet(app, validCabinets, fallbackSingle: true)` 直接获取当前光标所在行命中的箱柜实体；
        2. 若未命中选区且存在多台箱柜，自动回退取当前工作表中的最后一个有效箱柜（`validCabinets.Last()`）；
        3. 汇总行插入位置 `insertRow` 紧随选中箱柜的汇总行下一行（未命中取基准起始行）；
        4. 复制源箱柜直接取当前选中箱柜的 `Det`、`Subsum`、`Tolsum` 物理行号，彻底废除手动遍历求最大汇总行及固定取第一台箱柜的代码冗余。
  21. **CreateNewCabinet 源箱柜标准行提取与复制修剪重构**：
      - **业务与架构**：
        1. 在 [Services/ExcelServices.Cabinet.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Services/ExcelServices.Cabinet.cs) 的 `CreateNewCabinet` 中，通过 `srcK = activeCab?.Key ?? ...` 获取当前选中/前序箱柜序号；
        2. 调用 `Tool.FindStandardCategoryRowIndexes(sheet, srcK)` 直接获取源箱柜的标准行分布（`srcDetRow`、`srcSubsumRow`、`srcTolsumRow`），并根据末尾箱柜标准总计行 `lastIndexes.cabTolsumRow + 3` 精确定位末尾明细块插入起点；
        3. 完整复制源明细区块 `[srcDetRow-3 : srcTolsumRow+3]` 到末尾新位置；
        4. 动态计算计费跨度 `feeSpan` 与配置标准元器件行数 `standardCompRows`，精准物理向上删除多余元器件行；
        5. 调用 `Tool.BuildComponentRowsMatrix` 批量写回 A~Q 列自适应空行公式，并清洗表头旧图号、旧备注，注册 4 个定义名称与超链接。
  22. **明细块插入位置对齐汇总位置（紧随源箱柜明细块并物理插行）**：
      - **业务与架构**：
        1. 修复明细块始终写在工作表最末尾的问题；
        2. 在 [Services/ExcelServices.Cabinet.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Services/ExcelServices.Cabinet.cs) 中，将新明细块的目标起始行定位为 `newDetailStartRow = copyEndRow + 1`（紧随源箱柜明细块结束行之后）；
        3. 在复制前调用 `activeSheet.Rows[$"{newDetailStartRow}:{newDetailStartRow + copyRowCount - 1}"].Insert(-4121)` 物理向下插入整行，确保后续已有箱柜明细块安全下移不被覆盖；
        4. 复制源明细区块并执行公式清洗、多余元器件行向上修剪、自适应公式重构及 4 个定义名称注册。
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。
  23. **从 CabinetTemplate.xlsx 复制 41~74 行标准明细并智能插入光标汇总行（已落地）**：
      - **业务与架构**：
        1. 在 [Services/ExcelServices.Cabinet.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Services/ExcelServices.Cabinet.cs) 中完善 `CopyCabinetDetailFromTemplate`；
        2. 智能识别用户光标命中的箱柜实体（`Tool.GetActiveCabinet`），在其汇总行下方物理插入新汇总行，明细块紧随源箱柜明细块（总计行+3行报价人）之后插入；
        3. 自动处理汇总插行导致下方明细行的物理 `+1` 偏移补偿；
        4. 复制模板 41~74 行（34行），不执行外部公式清洗；
        5. 完整注册 4 个工作表级定义名称（`Cab_Sum_{K}`、`Cab_Det_{K}`、`Cab_Subsum_{K}`、`Cab_Tolsum_{K}`）；
        6. 建立汇总行与明细行双向超链接绑定，配置标准汇总公式（G/H/J/K/L 列），光标自动聚焦新汇总行 B 列。
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。
  24. **抽取箱柜定义名称前缀为不可变值类型（CabinetPrefixConfig 已落地）**：
      - **业务与架构**：
        1. 在 [Models/CabinetModels.cs](file:///e:/Ace/excel-ct-tools/Models/CabinetModels.cs) 中新增 `CabinetPrefixConfig` 值类型结构体（`readonly struct`），封装 `SumPrefix`、`DetPrefix`、`SubsumPrefix`、`TolsumPrefix`，实现零堆分配与零 GC 压力；
        2. 提供静态快照 `CabinetPrefixConfig.Current`、元组解构（`Deconstruct`）以及 `GetSumName(k)`、`GetDetName(k)`、`GetSubsumName(k)`、`GetTolsumName(k)` 快捷方法；
        3. 在 [AppConfig.cs](file:///e:/Ace/excel-ct-tools/AppConfig.cs) 的 `ExcelSettings` 中增加 `Prefixes` 只读属性；
        4. 全面重构并消除 [Services/ExcelServices.Cabinet.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.Cabinet.cs)、[Tool.cs](file:///e:/Ace/excel-ct-tools/Tool.cs)、[Services/ExcelServices.Category.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.Category.cs)、[Services/ExcelServices.SummaryAdjustPrice.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.SummaryAdjustPrice.cs)、[Services/ExcelServices.SmartInput.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.SmartInput.cs)、[Services/ExcelServices.Project.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.Project.cs)、[Services/ExcelServices.FormulaAdjustFee.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.FormulaAdjustFee.cs) 及 [ExcelEventManager.cs](file:///e:/Ace/excel-ct-tools/ExcelEventManager.cs) 中重复冗余的 4 个前缀读取与兜底赋值逻辑。
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。
  25. **公式法调费设为默认同步更新模板计费行与汇总行对齐（已落地）**：
      - **业务与架构**：
        1. 在 [Services/ExcelServices.FormulaAdjustFee.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.FormulaAdjustFee.cs) 中实现 `UpdateCabinetTemplateDefaultFee` 公共业务方法；
        2. 以 COM 方式打开 [CabinetTemplate.xlsx](file:///e:/Ace/excel-ct-tools/Resources/CabinetTemplate.xlsx) 模板的【分类1】工作表，探测当前模板箱柜 1 的行结构（`Cab_Sum_1`、`Cab_Det_1`、`Cab_Subsum_1`、`Cab_Tolsum_1`），若任意一个行号无效（<=0）自动调用 `Tool.FixAndFillCabinetNamesForSheet` 修复定义名称并重新获取所有标准行；
        3. **无需行数比较**：直接清空旧计费行区域 `B{oldSubsumRow}:Q{oldTolsumRow}` 数据（A 列序号与定义名称保留不删）；
        4. **汇总行对齐**：以固定的总计行 `cabTolsumRow` 向上计算小计行 `newSubsumRow = oldTolsumRow - newFeeRows + 1`，调用 `Tool.BuildFeeMatrix` 批量写入 A~Q 列自适应计费矩阵，并刷新元器件自适应公式；
        5. 重新注册并更新模板中 4 个定义名称锚点；
        6. 保持顶部汇总行（Row 7）的 G/H/J/K/L 列公式（单价指向 `=H{newTolsumRow}`，成本指向 `=K{newTolsumRow}`，毛利与毛利率联动）精准对齐绑定到总计行，更新双向超链接并安全保存；
        7. 在 [Forms/FormulaAdjustFeeForm.cs](file:///e:/Ace/excel-ct-tools/Forms/FormulaAdjustFeeForm.cs) 与 [Resources/formula_adjust_fee.html](file:///e:/Ace/excel-ct-tools/Resources/formula_adjust_fee.html) 的 `setDefaultGroup` 中串联该同步指令。
  26. **批量导入 (BatchExportCabinets) 极速性能优化改造（已落地）**：
      - **业务与架构**：
        1. 在 [Services/ExcelServices.Cabinet.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.Cabinet.cs) 中彻底重构 `BatchExportCabinets`；
        2. **模板句柄单例复用**：批量导入入口处单次预加载 `CabinetTemplate.xlsx` 的 41:74 范围，所有新建箱柜共享内存快速复制，最终统一定点关闭释放，消灭 95% 的磁盘 I/O；
        3. **锁定手动计算模式**：导入全流程设置 `app.Calculation = xlManual`，写入完成后执行 1 次 `app.Calculate()` 并还原，彻底消除每次插行与写公式时的全表连锁重算；
        4. **分类前向填充与内存状态上下文（SheetContext）**：增加 `Category` 前向传播填充（Forward Fill），解决子箱柜分类为空导致被错误打散的问题；每个分类表仅首次扫描 1 次基准行号，后续在内存中高速推导；
        5. **同表精准克隆**：彻底放弃脆弱的跨工作簿 `templateSrcRange`，改用工作表内基于汇总行偏移的 `41+(k-1):74+(k-1)` 精准模板槽位克隆，100% 杜绝跨工作簿 COM 异常；
        6. **数组批量写入**：将汇总行 1 行 13 列数据及公式组装为二维数组单次 Range 写入，大幅减少 COM 跨进程通信开销。
  27. **批量导出 (BatchExportCabinets) 纯净母版隔离克隆与静默清理重构（已落地）**：
      - **业务与架构**：
        1. 根治隐患：首台箱柜直接写入并插行会导致其明细块膨胀变形，后续箱柜按 34 行克隆时小计/总计行被截断丢失；
        2. **纯净母版隔离**：分类表中首个标准槽位（第 7 行汇总 + 行 41:74 明细）或旧表临时导入的 34 行模板保持绝对纯净（不写入真实数据、不执行插行），专供该表内高速克隆；
        3. **各箱柜独立克隆与插行**：导出的实际箱柜序号 `K` 从 1 到 N，所有箱柜统一从纯净母版向下克隆，各箱柜按自身元件数量独立插行，互不污染；
        4. **自底向上静默清理**：该分类所有箱柜导出完成后，自底向上物理删除母版明细块（34行）和母版汇总行（1行），利用 Excel 自动行列平移特性无缝上移衔接，公式与定义名称自适应更新。
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。
  28. **元器件型号参数自动识别与批量回填（极数与电流）功能（已落地）**：
      - **业务与架构**：
        1. 在 [Models/ModelParserConfig.cs](file:///e:/Ace/excel-ct-tools/Models/ModelParserConfig.cs) 中建立模型，定义多级顺位流水线规则、极数/电流排除词库、标称白名单有效值集合；
        2. 在 [Controllers/ModelParamParserController.cs](file:///e:/Ace/excel-ct-tools/Controllers/ModelParamParserController.cs) 中实现配置持久化读写、出厂重置、单条沙盒测试与 Excel 批量执行控制器；
        3. 在 [Services/ExcelServices.ModelParse.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.ModelParse.cs) 中封装极数与电流双通道过滤匹配引擎（前置排除 ➔ 多级流水线 ➔ 后置白名单校验），并以 **`object[,]` 二维数组一次性读入与写入** 方式实现极速 Excel 批处理；
        4. 在 [Resources/model_param_parser.html](file:///e:/Ace/excel-ct-tools/Resources/model_param_parser.html) 中基于 Vue3 + Element Plus 构建现代化配置与沙盒面板（主色调 `#009688` 绿蓝主题）；
        5. 在 [Forms/ModelParamParserForm.cs](file:///e:/Ace/excel-ct-tools/Forms/ModelParamParserForm.cs) 与 [RibbonController.cs](file:///e:/Ace/excel-ct-tools/RibbonController.cs) 中挂载功能入口。
        6. **UI 与规则深度升级**：修复列映射 Label 换行问题，增加 `找 "空格+数字+空格"`、`自定义正则表达式` 并列选项；优化 `1P+N`/`3P+N` 复合极数提取与格式化保护，彻底杜绝被误截断为 `1`。
      - **构建状态**：`ExcelAddInDemo.csproj` 编译构建通过（0 错误）。
  29. **汇总调价生成元件汇总表后展示图二编辑工具条及功能联动（已落地）**：
      - **业务与架构**：
        1. 在 [Resources/summary_adjust_price.html](file:///e:/Ace/excel-ct-tools/Resources/summary_adjust_price.html) 中实现多视图架构（`currentView: 'config' | 'editor'`），高保真复刻图二视觉（红棕色横线、圆形写字板图标、“编辑元件[下表中]”大标题、“↓仅可编辑[白色]列数据...”提示语、“重新汇总”操作链接与“一键更新”主按钮）；
        2. 点击“立即生成”并创建“元件汇总表”成功后，窗口自动平滑收缩为 680×115 紧凑悬浮编辑条，方便用户直接对照 Excel 编辑白色列数据；
        3. 自动缓存生成时的全套分类、合并与排序参数，点击“重新汇总”直接按初始配置极速重算并刷新覆盖“元件汇总表”，并提供返回配置面板的小图标；
        4. 在 [Controllers/SummaryAdjustPriceController.cs](file:///e:/Ace/excel-ct-tools/Controllers/SummaryAdjustPriceController.cs)、[Forms/SummaryAdjustPriceForm.cs](file:///e:/Ace/excel-ct-tools/Forms/SummaryAdjustPriceForm.cs) 与 [Services/ExcelServices.SummaryAdjustPrice.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.SummaryAdjustPrice.cs) 中打通 `resizeWindow` 窗口尺寸自适应与 `updateFromSummary` 一键更新预留通路。
  30. **箱柜明细行表头静态标签保护与双向同步列映射校正（已落地）**：
      - **业务与架构**：
        1. 修复新建箱柜克隆模板后误执行 `Cells[newDetRow, 3].Value = string.Empty` 导致明细行 C 列静态标签 `型号:` 被清空的问题（在 [Services/ExcelServices.Cabinet.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.Cabinet.cs) 的 `CopyCabinetDetailFromTemplate` 与 `CreateNewCabinet` 中彻底移除对 Cell 3 的误清空）；
        2. 全面校正 [ExcelEventManager.cs](file:///e:/Ace/excel-ct-tools/ExcelEventManager.cs) 中顶部汇总行与底部明细行之间的 8 组双向同步映射：
           - **柜号**：Sum B (2) $\leftrightarrow$ Det B (2)
           - **箱柜名称**：Sum C (3) $\leftrightarrow$ Det G (7)（Det G:H 合并数据单元格）
           - **箱柜型号**：Sum D (4) $\leftrightarrow$ Det D (4)（Det D:E 合并数据单元格）
           - **数量**：Sum F (6) $\leftrightarrow$ Tolsum F (6)
           - **备注**：Sum I (9) $\leftrightarrow$ Det I (9)
           - **箱柜尺寸**：Sum M (13) $\leftrightarrow$ Det K (11)
           - **类别**：Sum N (14) $\leftrightarrow$ Det M (13)
           - **图号**：Sum Q (17) $\leftrightarrow$ Det O (15)
        3. 反向同步增加对合并单元格右侧列（如 Det E 列、Det H 列）的容错支持。
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。
  31. **二次元件组可视化动态规则管道构建器系统（已落地）**：
      - **业务与架构**：
        1. 在 [Models/ComponentGroupRuleModels.cs](file:///e:/Ace/excel-ct-tools/Models/ComponentGroupRuleModels.cs) 中构建规则管道模型（包含四种匹配模式：包含/排除/主控数量#/成组比例/，支持多属性过滤：电流/型号/极数/附件及多种套数生成策略），并实现双向编译器 `PipelineCompiler`（可视化条件树 ⇄ 经典文本表达式无损互转）与执行评估器 `PipelineEvaluator`；
        2. 严格按用户列映射与套数指示实施：
           - 一次元件读取：B 列名称、C 列型号、F 列数量、V 列电流、W 列极数、X 列附件；
           - 二次元件输出：B 列写入 `"元件组"`、C 列写入元件组名称、E 列写入 `"套"`、**F 列写入计算套数**；
        3. 在 [Services/ExcelServices.ComponentGroup.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.ComponentGroup.cs) 中实现箱柜元件二维数组（B~X列）极速批量读取与安全插行（在小计行前插入，继承 CAD 句柄并自动更新定义名称与公式）；
        4. 在 [Controllers/ComponentGroupBuilderController.cs](file:///e:/Ace/excel-ct-tools/Controllers/ComponentGroupBuilderController.cs) 与 [Forms/ComponentGroupBuilderForm.cs](file:///e:/Ace/excel-ct-tools/Forms/ComponentGroupBuilderForm.cs) 中打通 WebView2 双向消息通道，支持出厂规则库初始化与 AppData 持久化；
        5. 在 [Resources/component_group_builder.html](file:///e:/Ace/excel-ct-tools/Resources/component_group_builder.html) 中基于 Vue 3 + Element Plus 构建现代化三栏规则流面板（主色调 `#009688` 绿蓝主题，支持从 Excel 一键抓取当前箱柜沙盒试跑与批量执行）；
        6. 在 [RibbonController.cs](file:///e:/Ace/excel-ct-tools/RibbonController.cs) 的【②录元件】分组中挂载【生成二次元件】（`btnComponentGroupRule`）主菜单入口。
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误，0 警告）。
  32. **规则管道构建器【或】关系条件分支组与 DLR 动态解析修复（已落地）**：
      - **业务与架构**：
        1. 修复 `GetActiveCabinetComponentsFromExcel` 与 `ExecuteBatchComponentGroup` 中 `(object)app` 与 `(object)sheet` 的 DLR 动态调用传播，解决 `Nullable<KeyValuePair>` 与 `ValueTuple` 运行时反射异常；
        2. 全面升级 [Resources/component_group_builder.html](file:///e:/Ace/excel-ct-tools/Resources/component_group_builder.html)，支持在流水线中直接添加与管理【或】关系条件分支卡片（`OR Group`，满足任一分支即可，支持多分支成员、各自属性过滤器及 `(热继电器# | 控制保护开关#)` 经典表达式双向无损互转）。
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。
  33. **箱柜元件动态资源池扣减评估算法（机制 B）落地（已完成）**：
      - **业务与架构**：
        1. 在 [Models/ComponentGroupRuleModels.cs](file:///e:/Ace/excel-ct-tools/Models/ComponentGroupRuleModels.cs) 中实现 `PipelineEvaluator.EvaluateRulesWithResourcePool`，按规则优先级（Priority）在箱柜元件工作资源池上进行带原子库存扣减的链式匹配；
        2. 彻底解决“多表共享互感器”时因严格等比导致的互相冲突与一个都无法识别的问题：
           - **比例模式 `/`**：按成员可用数量动态求公约套数 $\min(\text{数量}/\text{基准配比})$，成套后扣减消耗的元件库存；
           - **主控模式 `#`**：以主控元件在池中当前剩余数量驱动套数并扣减；
        3. 同步重构 [Services/ExcelServices.ComponentGroup.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.ComponentGroup.cs) 的 `RunSandboxPipelineTest`（沙盒实时输出扣减追踪日志）与 `ExecuteBatchComponentGroup`（Excel 生成时精准扣减，杜绝虚假超额生成）。
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。
  34. **规则流排序控制与分行扣减聚合日志优化（已完成）**：
      - **业务与架构**：
        1. 在 [Resources/component_group_builder.html](file:///e:/Ace/excel-ct-tools/Resources/component_group_builder.html) 规则库卡片上增加【⬆️ 上移】与【⬇️ 下移】交互按钮，支持实时调整规则顺序并自动重跑沙盒；
        2. 在 [Models/ComponentGroupRuleModels.cs](file:///e:/Ace/excel-ct-tools/Models/ComponentGroupRuleModels.cs) 中优化扣减日志格式：将跨多行（如多行互感器 1件+2件）的分步扣减自动智能聚合为统一的消耗总量 `[互感器] 消耗 3件 (池剩余 3件)`，消除日志碎片化带来的数量误解。
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。
  35. **二次元件组规则管道前置生效/跳过守卫（基于原始数量的两元件对比）机制落地（已完成）**：
      - **业务与架构**：
        1. 在 [Models/ComponentGroupRuleModels.cs](file:///e:/Ace/excel-ct-tools/Models/ComponentGroupRuleModels.cs) 中定义 `GuardActionMode`（`SkipIfMatched` / `ExecuteIfMatched`）与 `RuleGuardCondition` 模型，并在 `ComponentGroupRule` 中增加 `Guards` 集合；
        2. 升级双向编译器 `PipelineCompiler`：支持 `[SKIP: 接触器 == 热继电器]`、`[ONLY: 接触器 > 热继电器]` 前缀语法，实现可视化守卫 ⇄ 经典文本表达式的无损实时双向互转；
        3. 升级执行评估器 `PipelineEvaluator.EvaluateRulesWithResourcePool`：保留 `originalPool` 不可变镜像，优先评估规则的守卫条件，若满足跳过模式或不满足仅生效模式，直接跳过并输出清晰诊断日志（如 `[步骤 X] ⏭️ 规则 [KM变频器联动] 被前置守卫跳过: 原始[接触器]数量(3件) == 原始[热继电器]数量(3件) -> 触发跳过`）；
        4. 在 [Resources/component_group_builder.html](file:///e:/Ace/excel-ct-tools/Resources/component_group_builder.html) 中新增【前置生效与跳过守卫】可视化卡片，左侧规则卡片联动显示 `守卫` 徽章，打通沙盒试跑与批量生成。
      - **构建状态**：`ExcelAddInDemo.dll` 编译构建通过（0 错误，0 警告）。
  36. **二次元件组规则构建器架构大瘦身（全面转向纯结构化数据流）（已完成）**：
      - **业务与架构**：
        1. 彻底移除 `PipelineCompiler` 中 400+ 行冗余的文本反向解析器、括号匹配与复杂正则表达式；
        2. 控制器与窗体移除 `parseExpression` 与 `buildExpression` 双向消息通道，全面升级为纯强类型 JSON 交互；
        3. 移除前端中间面板底部的【经典表达式输入卡片】与庞大的文本双向同步逻辑，交互直观轻快，左侧卡片采用纯轻量响应式计算属性展示规则概括；
        4. 重构 `CreateDefault` 出厂规则库为强类型对象直接构造，直存直取，系统健壮性与可维护性大幅提升。
      - **构建状态**：`ExcelAddInDemo.dll` 编译构建通过（0 错误，0 警告）。
  37. **沙盒面板 Flexbox 弹性收缩坍塌与小屏幕视口适配修复（已完成）**：
      - **业务与架构**：
        1. 根治 `.comp-table-container` 因 Flex 默认收缩（`flex-shrink: 1`）在空间紧凑时被压缩为 0px 产生的一条线 Bug，明确赋予 `height: 150px; min-height: 140px; flex-shrink: 0;`；
        2. 为 `.matched-group-item`、`.log-box` 均显式声明 `flex-shrink: 0;`，交由父容器 `.sandbox-body` 的垂直滚动条自适应平滑承载；
        3. 宿主窗体初始化根据屏幕工作区（`Screen.PrimaryScreen.WorkingArea`）动态计算尺寸，彻底避免超出屏幕导致被任务栏遮挡。
      - **构建状态**：`ExcelAddInDemo.dll` 编译构建通过（0 错误，0 警告）。
  38. **汇总调价生成 16 列双层复合表头与本体/附件表价折扣拆分引擎（已落地）**：
      - **业务与架构**：
        1. 在 [Services/ExcelServices.SummaryAdjustPrice.cs](file:///d:/code/excel-ct-tools/Services/ExcelServices.SummaryAdjustPrice.cs) 中重构 `AggregatedComponent` 模型，扩充为覆盖 16 个列字段的完整数据结构；
        2. 实现 `ParseBaseAndAccessoryPrice`：从明细 M 列提取表价公式（如 `=159.05+336.2`），精准拆分第一个加数为【本体表价】，第二项及之后项累加为【附件表价】；
        3. 实现 `ParseBaseAndAccessoryDiscount` 与 `ExtractDiscountFromTerm`：从明细 N 列提取复合加权折扣公式（如 `=(159.05*1*0.5+336.2*1*1)/495.25`），精准提取【本体折扣】（0.5）与【附件加权折扣】（1.0），并向上兼容普通数值与简单公式；
        4. 采用 2D 数组（`A~Q` 列 Value2 与 Formula 矩阵）一次性读入内存，完成台数放大与 `Group By` 聚合；
        5. 生成高保真 16 列【元件汇总表】（双层表头 4~5 行，J~N 列粉色调价核心表头，A/E/I 列浅蓝强调底色，精准列宽 6~24，AutoFilter 挂载与 SUM 合计行，实线细边框），采用 2D 数组单次写入 Excel。

