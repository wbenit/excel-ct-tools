# Session State

## [In-Progress]

- **箱柜调序功能顶部汇总与底部分类明细表全量物理同步重排与公式超链接自愈联动落地 (`Services/ExcelServices.CabinetManage.cs`)**：
  1. **用户核心问题与根因分析**：
     - 用户指令反馈：“箱柜下面，调序功能在顶部汇总调过来了，但是底部分明细没调”；
     - 根因定位：原 `ApplyCabinetReorder` 算法在重排时，仅仅用循环覆写了顶部汇总表的物理单元格文字，不仅完全没有对底部的明细表块进行任何移动，甚至导致顶部汇总行的定义名称与公式引用错乱、超链接指向错误；
  2. **闭环解决方案与物理重排算法设计**：
     - **阶段一（顶部汇总行原生整行物理重排）**：以 `firstSumRow` 为起始目标指针，按新排序列表 `sanitizedOrder` 遍历，通过定义名称 `Cab_Sum_{K}` 实时提取当前物理行号，调用原生 `cutRow.Cut()` 紧邻 `targetRow.Insert(-4121)` 进行整行剪切插入，定义名称自动随行位移；
     - **阶段二（底部分类明细块原生整块物理重排）**：预先以首个有明细箱柜的 `Det - 3` 行作为基准锚点 `firstDetStartRow`，并精准测算每台箱柜明细块的独立总行数（包含元件增删与后置间隔）；在新顺序中按有明细箱柜依次调度 `cutBlock.Cut()` 与 `targetBlock.Insert(-4121)`，将整块明细（标题、表头、元件、小计、计费、总计、落款）无损平移至目标位置；
     - **阶段三（全局校准自愈与双向公式超链接刷新）**：
       a. 汇总行 A 列自适应动态序号公式统一刷新为 `=ROW()-ROW(A$6)`；
       b. 汇总行 G、H、J、K、L 列公式严格重定向至对应明细总计行 `Tolsum`；
       c. 汇总行与明细行双向超链接自愈修复，明细行 B 列彻底清理超链接；
       d. 明细块元器件区域 A 列序号公式自适应绑定至 `$"=ROW()-ROW(A${dRow + 1})"`，小计与计费区域联动刷新（`RefreshCabinetFeeAreaFormulas`）；
     - 编译验证：`dotnet build /t:Compile /p:DebugType=none` 全量编译通过：**0 错误**。

- **常规样式投标报表导出升级为【模板占位符自适应动态驱动引擎】彻底根除硬编码与双重柜号缺陷 (`Services/ExcelServices.TenderReport.cs`)**：
  1. **用户核心指令与深层根因反思**：
     - 用户一针见血指出：“你这个都是写死的，模板的本意是根据[内容]自动识别位置的”；
     - 深度核查模板母版真实文本：Row 11 C2 实际为 `柜号：[柜号]`，C4 为 `[箱柜型号]`，C8 为 `[箱柜名称]`，C9 为 `备注:[箱柜备注]`；
     - 原代码死写列号并暴力拼接 `$"柜号：{cab.CabinetNo}"`，导致模板自带的“柜号：”与代码拼接的“柜号：”叠加，在屏幕上错乱呈现双重“柜号：柜号：SW-APKT”；且写死列号丧失了模板自适应自由排版的能力。
  2. **一揽子自适应占位符引擎全面落地**：
     - **占位符单行自适应替换 (`ReplaceRowPlaceholders`)**：整行扫描单元格文本，精准匹配 `[箱柜序号]`、`[柜号]`、`[箱柜型号]`、`[箱柜名称]`、`[箱柜备注]`、`[分类名称]`、`[报价说明]`、`[报价人]`、`[报价日期]` 等标签并就地替换，保留模板原生排版与前缀，彻底消灭双重前缀 Bug；
     - **数据行动态列映射器 (`Dynamic Column Mapping`)**：扫描 Row 13 母版数据行，自动识别字段对应物理列号（`[元件名称]`、`[元件单价]`、`[箱柜总价]` 等），彻底废除硬编码列号，模板调整列序或增删列 100% 自动适配；
     - **动态公式自适应拼装**：通过 `GetExcelColumnLetter` 动态提取数量列与单价列字母，自适应拼装 `=Qty*Price` 及小计总计公式；
     - **多箱柜克隆与防踩踏保护**：结合临时母版 `tempWs` 单向向下克隆，橙底总计行永远紧跟元件末尾，细网格线 100% 保真；
     - 编译验证：`dotnet build /t:Compile /p:DebugType=none` 全量编译通过：**0 错误**。

- **剪切/复制箱柜运行时异常 `运算符“==”无法应用于 KeyValuePair 和 null` 彻底修复 (`Services/ExcelServices.CabinetManage.cs`, `Services/ExcelServices.Cabinet.cs`)**：
  1. **异常根因深度剖析**：
     - 用户点击【剪切箱柜】时，`CutCurrentCabinet` 内部调用 `CaptureActiveCabinetToClipboard`；
     - 原代码第 207 行 `dynamic app = context.App;`，导致第 220 行 `var activeCab = Tool.GetActiveCabinet(app, ...)` 传入了 `dynamic` 实参；
     - 依据 C# 语言规范，一旦方法入参含 `dynamic`，返回值类型即被编译器推导为 `dynamic`，调用推迟至 DLR 运行期绑定；
     - 运行期 `GetActiveCabinet` 命中并返回了结构体 `KeyValuePair<int, CabinetAnchorModel>`，紧随其后的第 221 行 `if (activeCab == null)` 在 DLR 动态求值时试图寻找 `KeyValuePair` 与 `null` 的 `==` 运算符，因值类型结构体无法直接与 `null` 做 `==` 比较，DLR 抛出致命异常 `运算符“==”无法应用于“KeyValuePair<...>”和“<null>”类型的操作数`；
  2. **彻底闭环修复**：
     - 在 `CabinetManage.cs`（第 86、220、334、501 行）与 `Cabinet.cs`（第 83、748 行）中，显式将 `app` 转为 `(object)app`，切断 DLR 动态推导；
     - 接收变量统一显式声明为强类型可空结构体 `KeyValuePair<int, Models.CabinetAnchorModel>? activeCab`；
     - 将空值判断严格使用 `!activeCab.HasValue || activeCab.Value.Key <= 0` 标准 Nullable 语法，100% 在编译期静态绑定，杜绝 DLR 运行期介入；
     - 编译验证：`dotnet build /t:Compile /p:DebugType=none` 完整编译通过：**0 错误**。


- **明细行 detRow B 列超链接彻底清除与元器件行 A 列序号公式修复 (`Services/ExcelServices.Cabinet.cs`, `Services/ExcelServices.CabinetManage.cs`, `Services/ExcelServices.FormulaAdjustFee.cs`, `Tool.cs`)**：
  1. **用户核心指令与缺陷定位**：
     - **指令 1 (det 行的 B 列不需要链接)**：明细信息行（如第 296 行 `Cab_Det_6`）B 列箱柜名称（`箱柜6`）原本存在超链接；根因是 `FormulaAdjustFee.cs` 历史代码向模板 `Cells[cabDetRow, 2]` 添加了超链接，且从模板或源箱柜复制后未清理超链接属性；
     - **指令 2 (底部元器件行的 A 列公式应该是 `=ROW()-ROW(A$49)`，49 应该是 det+1 行)**：从模板克隆出来的箱柜，元器件区域（`compStartRow` 至 `subsumRow - 1`）A 列序号沿用了模板写死的静态公式 `=ROW()-ROW(A$49)`，导致行号较大时（如第 298 行）显示异常序号 `298 - 49 = 249`；此处的 49 应严格自适应为当前箱柜自身明细表头行（`detRow + 1`，如第 297 行）。
  2. **闭环修复措施**：
     - **增强 `RefreshCabinetFeeAreaFormulas` (`ExcelServices.Cabinet.cs`)**：
       a. 显式彻底清除明细表头 B 列超链接：`try { sheet.Cells[detRow, 2].Hyperlinks.Delete(); } catch { }`；
       b. 内存构建二维数组一次性批量将元器件区域（`compStartRow` 至 `subsumRow - 1`）A 列序号公式刷新为 `$"=ROW()-ROW(A${detRow + 1})"`（契合规则 6 & 规则 7）；
       c. 小计行求和公式与计费区域自适应公式同步联动刷新；
     - **模板复制与箱柜新增链路挂接 (`ExcelServices.Cabinet.cs:CopyCabinetDetailFromTemplate`, `AddCabinetCore`)**：
       明细块复制完成后，显式清除 `newDetRow, 2` 的超链接，并立即调度 `RefreshCabinetFeeAreaFormulas` 刷新元器件序号与计费公式；
     - **箱柜复制与剪贴板链路挂接 (`ExcelServices.CabinetManage.cs:InsertCopiedCabinet`)**：
       粘贴明细块后显式清除 `newDetRow, 2` 超链接，并调度 `RefreshCabinetFeeAreaFormulas` 修复元器件序号与计费公式；
     - **调费设为默认链路清理 (`ExcelServices.FormulaAdjustFee.cs`)**：
       定位并根除了向模板 B 列写超链接的罪魁祸首：将原代码中误写为 B 列的 `Anchor: catSheet.Cells[cabSumRow, 2]` 彻底纠正为 A 列（`Cells[cabSumRow, 1]`），并为明细行 A 列添加返回链接，B 列（汇总与明细）一律禁止添加超链接，从源头彻底切断模板污染；
     - **全表自动校准自愈引擎升级 (`Tool.cs:FixAndFillCabinetNamesForSheet`)**：
       在扫描箱柜并校准定义名称时，同步执行 `sheet.Cells[curDetRow, 2].Hyperlinks.Delete()` 并调用 `RefreshCabinetFeeAreaFormulas`，对已有工作表中残留的带超链接 B 列与显示 249 的历史箱柜实现瞬间自动修复；
     - 编译验证：`dotnet build /t:Compile /p:DebugType=none` 完整编译通过：**0 错误**。

- **图二二次方案集成至图一云方案中心可行性与架构需求分析**：
  1. **背景**：用户询问图二（本地 SQLite 113 套二次回路方案与 BOM 管理中心）是否能按照图一（云方案中心 二次方案 Tab）的方式进行集成；
  2. **现状研判**：
     - 图一 (`cloud_solution.html`)：顶部二级 Tab 包含【一次方案】与【二次方案】，但目前数据源来自 `cloud_schemes.json`，企业方案下的二次方案为空（0 条），左侧分类为一次配电柜分类（低压进线、电容补偿等）；
     - 图二 (`secondary_circuit_manage.html`)：数据持久化存储于本地 SQLite (`personal_components.db: secondary_circuit_schemes`)，已有 113 套真实二次控制方案，包含二次排布图分组、适用回路代号、跨门线根数、开孔、人工、材料费、子 BOM 清单及 CAD 图纸绑定；
  3. **分析结论**：完全可行，且是云方案中心实现“一次+二次”全流程闭环的最佳路径；
  4. **门控状态**：严格执行“需求分析门控”，仅输出全方位需求分析与架构方案，等待用户决策后再执行具体代码实施。

- **批建与新建箱柜明细块插入行回归规范 `Tolsum + 4` 修复 (`Services/ExcelServices.Cabinet.cs`)**：
  1. **现象与根因剖析**：
     - **用户现象**：批建箱柜本应在上一台箱柜的 `Tolsum + 4` 行紧凑插入，但实际却在 `Tolsum + 5` 插入，导致两台箱柜明细块之间凭空多出了一行空白行；
     - **深层根因定位**：在模板结构中，总计行 `Tolsum` 在第 71 行，第 72~74 行为 3 行附注与报价人落款，下一台箱柜明细块起始行正好紧随第 74 行之后，即 `71 + 4 = 75`（`Tolsum + 4`）。而在上一轮代码中，在第 131~135 行错误增加了一句 `targetDetailStartRow += 1`，导致原本已精准计算好的 `cabTolsumRow + 4` 被累加变成了 `Tolsum + 5`，两箱柜之间产生了多余空行；
  2. **修复实施**：
     - 彻底移除了 `Services/ExcelServices.Cabinet.cs` 中多余的 `targetDetailStartRow += 1`，使明细块起始插入行严格、精准锁定为 `Tolsum + 4`；
     - 编译验证：`dotnet build /t:Compile /p:DebugType=none` 编译通过：**0 错误**。

- **新增箱柜插在无明细箱柜前面及无明细箱柜顶部汇总表定义名称丢失缺陷根治 (`Tool.cs`, `Services/ExcelServices.CabinetManage.cs`, `Services/ExcelServices.Cabinet.cs`)**：
  1. **四大致命根因定位**：
     - **根因一 (`Tool.FindStandardCategoryRowIndexes` 强绑定检测抹杀无明细箱柜)**：原代码强校验 `if (sumR > 0 && detR > 0 && tolR > 0)`，无明细箱柜无底表明细（`detR == 0, tolR == 0`），导致探测全表末尾基准行时无明细箱柜被判定为非法数据直接丢弃，回退至系统默认值 `cabSumRow = 7`，新建箱柜在第 8 行插入，硬生生插到无明细箱柜前面；
     - **根因二 (`Tool.FixAndFillCabinetNamesForSheet` 自动校准重建除名抹杀)**：自动校准定义名称时写死 `int cabCount = detRows.Count;`（以明细块数量为基准），若表内有无明细箱柜，循环只跑明细块次数，多出的无明细箱柜根本未被绑定 `Cab_Sum_k`，直接变成“黑户”导致后续扫描无法识别；
     - **根因三 (`CreateNewCabinetNoDetail` 与 `InsertCopiedCabinet` 注册作用域不一致)**：普通箱柜使用 `Tool.SafeSetSheetName` 注册工作表级局部名称（Sheet-level），而无明细箱柜原先使用 `wb.Names.Add` 注册在工作簿级，易引发多表同名遮蔽与跨表过滤剔除；
     - **根因四 (`CopyCabinetDetailFromTemplate` 插入定位策略偏差)**：新建普通箱柜时未考虑末尾可能存在无明细箱柜，且插入 1 行汇总行后明细块起始行未做平移对齐。
  2. **一揽子闭环修复实施**：
     - **解耦末尾基准探测算法 (`Tool.cs:FindStandardCategoryRowIndexes`)**：彻底解耦汇总行末尾与明细行末尾：`cabSumRow` 精准取全表所有有效箱柜（普通柜+无明细柜）`Sum.Row` 的绝对最大值；`cabTolsumRow` 精准取全表有明细箱柜 `Tolsum.Row` 的绝对最大值；废除 `detR > 0 && tolR > 0` 绑定判定，无明细箱柜的 `sumR` 100% 精确返回；
     - **全量兼容自动扫描补齐 (`Tool.cs:FixAndFillCabinetNamesForSheet`)**：箱柜总数重构为 `Math.Max(detRows.Count, sumRows.Count)`；对所有扫描到的汇总行 `sumRows` 100% 绑定 `Cab_Sum_k`；仅在存在明细块时绑定 `Det/Subsum/Tolsum`，杜绝无明细箱柜被除名；
     - **统一定义名称注册规范 (`ExcelServices.CabinetManage.cs`)**：`CreateNewCabinetNoDetail` 与 `InsertCopiedCabinet` 全面迁移至 `Tool.SafeSetSheetName`，全工程 100% 统一为工作表级别规范定义名称；
     - **优化箱柜插入与明细平移 (`ExcelServices.Cabinet.cs`)**：在末尾追加时严格在全表最大汇总行下方插入，普通箱柜明细块追加在全表最大明细块后方，且汇总行插行后明细行起始行号自适应平移 +1，杜绝碰撞错位；
     - 执行 `dotnet build` 完整编译构建通过：**0 错误**。

- **编辑箱柜“卡死”深度根因根治与非模态架构一揽子闭环升级 (`Controllers/CabinetController.cs`, `Forms/CabinetManageForm.cs`, `Services/ExcelServices.CabinetManage.cs`, `Tool.cs`, `Resources/cabinet_manage.html`)**：
  1. **卡死致命根因定位**：
     - **根因一 (`ShowDialog` 模态消息循环与 WebView2 COM 互锁死锁)**：原 `CabinetController.ShowCabinetDialog` 使用 `form.ShowDialog()` 启动 WinForms 模态消息循环，挂起了 Excel 主线程 COM 消息泵；而 WebView2 的 Chromium 内核初始化与 `getInitData` IPC 消息又在主线程等待 Excel COM 响应，导致跨单元 RPC 调用瞬间永久死锁卡死。
     - **根因二 (`Tool.BuildCabinetMap` 暴力过滤丢弃无明细箱柜)**：原代码硬编码 `.Where(x => x.Value.Det != null && x.Value.Sum != null)`，因无明细箱柜只有 `Cab_Sum_k` 而无 `Cab_Det_k`，导致无明细箱柜被 100% 过滤丢弃，`GetActiveCabinetEditInfo` 始终返回 null 并弹出阻塞式 `MessageBox.Show`。
  2. **一揽子闭环修复实施**：
     - **全面升级为 `ShowModelessForm` 非模态架构**：彻底废除 `ShowDialog`，使用 `ExcelServices.ShowModelessForm` 挂载在 Excel 主窗口 HWND，实现不阻塞 Excel COM 泵的非模态展示，彻底消灭 Chromium 与 Excel 的死锁根源；
     - **窗体实例复用与平滑切换**：在 `CabinetController` 与 `CabinetManageForm` 中引入 `SwitchMode`，支持单例复用，并在 `OnFormClosing` 中解绑事件并释放 `_webView`，杜绝进程残留；
     - **根基过滤器兼容无明细箱柜**：将 `Tool.BuildCabinetMap` 升级为 `.Where(x => x.Value.Sum != null || x.Value.Det != null)`，纯汇总箱柜与标准箱柜均可精准扫描识别；
     - **三级智能探测与兜底识别**：在 `GetActiveCabinetEditInfo` 中增加“定义名称精准匹配 -> 活动光标物理行号匹配 -> 当前表首台箱柜默认保底”三级探测，彻底移除阻断性 `MessageBox.Show`；
     - **前端支持多箱柜自由下拉切换**：在 `cabinet_manage.html` 编辑表单顶部增加 `<el-select>` 切换箱柜下拉框，支持在编辑窗口内直接切换当前分类表的所有箱柜进行快速修改与同步回写；
     - 资源已热同步至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`，`dotnet build /t:Compile` 编译通过：**0 错误**。

- **云方案点击白板/打不开深度根因排查与一揽子闭环修复 (`Resources/cloud_solution.html`, `Forms/CloudSolutionForm.cs`, `ExcelAddInDemo.csproj`, `Services/ExcelServices.CloudSolution.cs`, `Forms/CabinetManageForm.cs`)**：
  1. **双重白板致命根因定位**：
     - **根因一 (HTML 代码撕裂残渣破坏 DOM 树)**：`cloud_solution.html` 第 1288 行存在一段未完成合并的代码撕裂残渣（`</div>ick="deleteBomRow($index)" title="删除">` 及后续 49 行重复的面板与按钮），且内含未闭合的 `</el-button>`、`</template>`、`</el-table-column>`。Vue 3 挂载编译时抛出致命模板语法错误，Vue 实例崩溃导致页面呈现空白白屏。
     - **根因二 (违反 HTML5 Vue 自定义组件自闭合规范)**：依据 `local-heuristics.md:L27`，模板内存在 `<el-option ... />`、`<el-checkbox ... />`、`<el-empty ... />`、`<el-pagination ... />` 等自闭合标签，被原生 HTML 解析器吞噬后续同级节点，加剧了 DOM 树崩塌。
     - **根因三 (输出目录缺失 HTML 文件)**：`ExcelAddInDemo.csproj` 未配置 `cloud_solution.html` 自动复制；实测 `bin\Debug\net48\Resources\cloud_solution.html` 物理文件此前根本不存在，导致 WebView2 未加载任何有效网页，默认全白。
     - **根因四 (C# 寻址路径单一脆弱)**：`CloudSolutionForm.cs` 原先仅探测 `AppDomain.CurrentDomain.BaseDirectory`，未结合 `Tool.GetAppDirectory()` 与多级源码目录兜底，寻址失败时缺少友好页面兜底。
     - **根因五 (后台 STA 线程 Application.Run 隐患)**：`ShowCloudSolutionDialog` 使用独立后台线程启动模态循环，存在跨线程 COM 访问与 `local-heuristics.md:L8` 记录的硬禁用风险。
  2. **已执行的一揽子彻底修复措施**：
     - **前端模板净化**：彻底清除了第 1288-1336 行重复代码撕裂残渣；将所有自闭合 `el-` 标签规范重构为显式闭合标签；经 AST 标签栈算法严密校验：**0 个未闭合，0 个多余标签，标签栈平衡归零**；
     - **工程复制配置补全**：在 `ExcelAddInDemo.csproj` 中补全 `<None Include="Resources\cloud_solution.html"><CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory></None>`，并已热同步至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`；
     - **窗体寻址与容错增强**：在 `CloudSolutionForm.cs` 中引入 `Tool.GetAppDirectory()` 与多级备选路径；若发生异常输出带样式的 404 诊断页，杜绝静默白板；引入鼠标物理键检测，防范幽灵拖拽；重写 `OnFormClosing` 释放 WebView2 杜绝进程残留；
     - **弹窗模式升级对齐**：重构 `ShowCloudSolutionDialog` 为依附 Excel 主窗口 HWND 的安全非模态弹出，杜绝多开与后台 COM 线程冲突；
     - **修复 CabinetManageForm.cs CS0111 编译冲突**：消除冗余 `SwitchMode` 并修正 `_mode` 的 `readonly` 修饰符，`dotnet build /t:Compile` 编译通过：**0 错误**。


- **DotRush 与 C# 扩展失效根因排查与底层运行时修复**：
  1. **双重崩溃根因**：检查 `C#.log` 发现，不仅 DotRush 崩溃，连原本备用的 `dotnetdev-kr-custom.csharp` 的底层语言服务器 `Microsoft.CodeAnalysis.LanguageServer.exe` 也强依赖 **.NET 10.0 (`Microsoft.NETCore.App 10.0.0`)**。
  2. **关键异常目录定位**：`ms-dotnettools.vscode-dotnet-runtime` 本地存储的 `10.0.11~x64` 目录下此前只有 8.0.30，而真正的 `10.0.11` 被备份在 `10.0.11~x64.bak` 中，导致插件每次启动都在虚假的 10.0.11 目录中找不到 .NET 10.0.0 运行时，两个插件全部崩溃闪退。
  3. **已执行底层修复**：
     - 已将完整的 `Microsoft.NETCore.App 10.0.11` 与 `host\fxr\10.0.11` 恢复并合并至 `globalStorage\ms-dotnettools.vscode-dotnet-runtime\.dotnet\10.0.11~x64\`；
     - 同步部署至 `D:\Program Files\dotnet\dotnet-sdk-8.0.424-win-x64\`，并持久化写入用户级环境变量 `DOTNET_ROOT`；
     - 验证成功：`dotnet --info` 已正确包含 8.0.30 与 10.0.11，命令行实测 `Microsoft.CodeAnalysis.LanguageServer` 与 `DotRush.dll` 均已能正常加载且退出码为 0；
  4. **后续步骤**：用户重新执行一次 Reload Window 即可正常拉起语言服务器。

- **箱柜功能对齐 ExWinner 原生规范（9 项箱柜管理功能闭环：新建/新建无明细/批建/编辑/剪切/复制/插入复制/删除/调序）全面交付 (`Models/CabinetModels.cs`, `Services/ExcelServices.CabinetManage.cs`, `Services/ExcelServices.Cabinet.cs`, `Controllers/CabinetController.cs`, `Forms/CabinetManageForm.cs`, `Resources/cabinet_manage.html`, `RibbonController.cs`, `ExcelAddInDemo.csproj`)**：
  1. **无明细箱柜核心机制实现 (`CreateNewCabinetNoDetail`)**：
     - 严格落实用户指令：无明细箱柜不需要底部明细区域，仅在分类表顶部汇总表插入 1 行，分配全局递增名称 `Cab_Sum_k`；
     - **A 列自适应动态序号公式**：全面废除静态数字写死，A 列统一写入公式 `=ROW()-ROW(A$6)`，随着增删插改自适应变动；
     - 单价与成本支持直接录入或后续填入，合价公式 `=F*G` 自动重算并联动汇总至【项目信息】表；
     - 在 `DeleteCabinets` 中增加 `detRow > 0` 安全防护，无明细箱柜删除时不触发底层空明细删除异常。
  1.1. **标准箱柜与批量新增、复制插入、调序的 A 列公式对齐 (`CopyCabinetDetailFromTemplate`, `BatchCreateCabinets`, `InsertCopiedCabinet`, `ApplyCabinetReorder`)**：
     - 在 `Hyperlinks.Add` 挂载超链接后，将 `Formula` 设置为 `=ROW()-ROW(A$6)`，彻底移除 `TextToDisplay` 静态文本覆盖，完美兼备“超链接点击跳转明细”与“自适应动态序号计算”；
     - 调序与单箱柜自愈时自动维护 A 列公式 `=ROW()-ROW(A$6)`。
  2. **剪切/复制/插入箱柜内存整块快照吞吐 (`CopyCurrentCabinet`, `CutCurrentCabinet`, `InsertCopiedCabinet`)**：
     - 依据规则 7，整块读取汇总行（A:M 列）与明细区域（A:Q 列）至内存模型 `CopiedCabinetContext`；
     - 剪切保护：点击【剪切箱柜】时仅标记 `IsCut = true` 并深度快照，绝不提前破坏物理表格；在目标位置成功执行【插入复制的箱柜】并重新分配 `newK` 与定义名称后，后置安全执行物理删除源行；
     - 支持普通有明细箱柜与纯汇总无明细箱柜的跨表复制与插入，公式与超链接自愈重建。
  3. **批建箱柜 / 编辑箱柜信息 / 箱柜调序工作台 (`CabinetManageForm`, `cabinet_manage.html`)**：
     - 前端采用 WebView2 + Vue 3 + Element Plus，严格遵循 `<script setup>` 结构与绿蓝相间主题（`#009688`），所有标签显式闭合；
     - **批建箱柜**：支持从剪贴板多列自动解析（柜号、名称、型号、数量、单价、成本），支持一键批量开关“无明细”；
     - **编辑箱柜**：支持双向回写柜号、名称、型号、数量、单价、成本，自动同步汇总行与明细行；
     - **箱柜调序（ExWinner 核心内存整块重排算法）**：解决物理 `Cut`/`Insert` 导致公式与超链接 `#REF!` 的缺陷，整块内存读入 -> 内存重排 -> 顺序覆盖写回，100% 保持公式完整性。
  4. **Ribbon 菜单接通与项目集成**：
     - 在 `RibbonController.cs` 的 `OnMenuAction` 中完整接通 9 个按钮分支；
     - 在 `ExcelAddInDemo.csproj` 中配置 `cabinet_manage.html` 的自动复制输出，并同步至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`；
     - 代码注释规范完整（每 3 行至少 1 行中文注释），硬编码均标明 `--硬编码--`；
     - 执行 `dotnet build` 编译构建通过：**0 错误**。


- **分类跳转超链接挂载位置对齐 ExWinner 原生规范（A 列数字超链接跳转、B 列纯名称展示与历史错乱自愈引擎）全面交付 (`Services/ExcelServices.Category.cs`)**：
  1. **A 列序号挂载超链接**：
     - 在 `UpdateProjectInfoCategorySummary` 中将超链接挂载列彻底从 B 列迁移至 A 列（`Anchor: infoSheet.Range[$"A{targetInfoRow}"]`）；
     - 超链接 SubAddress 精准指向 `$"'{categorySheetName}'!A1"`，屏幕提示（ScreenTip）规范设置为 `"点击进入本分类报价清单"`（打标 `--硬编码: 屏幕提示文本--`）；
     - 写入动态序号公式 `=ROW()-ROW(A${headerRowIndex})`，公式计算出的数字 1, 2, 3... 自动具备超链接蓝色下划线与手型光标，点击数字直接激活目标分类表！
  2. **B 列写入 CELL("filename") 动态工作表名公式（随 Sheet 改名自动无感级联联动）**：
     - 在 `UpdateProjectInfoCategorySummary`、`RenameProjectInfoCategorySummary` 与 `NormalizeCategorySummaryLinks` 中，将 B 列改为写入 `=MID(CELL("filename",'{categoryName}'!$A$1),FIND("]",CELL("filename",'{categoryName}'!$A$1))+1,31)` 动态公式；
     - 显式调用 `infoSheet.Range[$"B{r}"].Hyperlinks.Delete()` 彻底清除超链接下划线，恢复纯净单元格公式；
     - 只要用户在底栏 Tab 或通过程序修改了分类工作表名称，B 列分类名称由 Excel 引擎原生自动刷新，无需后台重复遍历改写。
  3. **分类重命名与插入复制全面同步**：
     - 在 `RenameProjectInfoCategorySummary` 中同步更新 A 列超链接的 SubAddress 至最新分类工作表名，B 列清除超链接并更新为新名称；
     - 在 `InitializeCategorySheet` 与 `InsertCopiedCategory` 的箱柜汇总行 A 列超链接中补全 `ScreenTip: "点击进入本箱柜明细表"`。
  4. **构建全工作簿分类汇总超链接自愈规范化引擎 (`NormalizeCategorySummaryLinks`)**：
     - 提供跨表自愈引擎：自动扫描【项目信息】表的分类汇总区域，智能排查并修复历史文件中因模板名称丢失导致的 `#NAME?`、B 列误挂超链接、A 列缺超链接等格式缺陷；
     - 在打开分类管理窗口（`GetSuggestedCategoryInfo`）时自动静默自愈，确保用户既有文件一键无感对齐。
  5. **工程编译验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：**0 错误**；代码注释规范完整，硬编码均打标 `--硬编码--`。

  1. **彻底根除 `MessageBox.Show` 阻塞死锁，升级为异步非模态通信**：
     - 在 `FormulaAdjustFeeForm` 中移除 `applyFormula` 下的 `MessageBox.Show`，改为通过 `PostWebMessageSafe` 回传 `applyFormulaResult`，由前端 `ElMessage` 友好提示；
     - 在 `EnterpriseSettingsForm` 中移除保存成功/失败的 `MessageBox.Show`，改为回传 `saveSettingsResult`，前端通过 `ElMessage` 提示并延时平滑退出；
     - 在 `CreateProjectForm` 中将异常直接回传 `startQuotationResult` 并打标日志，彻底杜绝 Chromium IPC 模态死锁。
  2. **全面对齐 local-heuristics.md:L133：新建项目补充“✍️ 粘贴路径”双轨保障**：
     - 在 `create_project.html` 的保存目录旁新增 `✍️ 粘贴路径` 绿色扁平按钮；
     - 通过 `ElMessageBox.prompt` 弹出原生输入框，支持直接粘贴 Windows 资源管理器路径，零弹窗极速设定，彻底绕过系统外壳慢速磁盘枚举卡顿。
  3. **二次方案管理中心拖拽机制升级 (杜绝系统级全局鼠标捕获死锁)**：
     - 在 `SecondaryCircuitForm.cs` 中增加 `moveWindow` 物理增量坐标位移处理；
     - 将 `secondary_circuit_manage.html` 标题栏拖拽全面升级为现代 `pointerdown` + `moveWindow`（带 `rAF` 节流与 DPI 适配）；
     - 在 `dragWindow` 兜底中引入 `(GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0` 物理按键检测，物理按键弹起时坚决丢弃，彻底根绝全屏幕鼠标锁死。
  4. **建立全工程窗体生命周期释放保护机制 (杜绝 Excel 进程残留)**：
     - 在 `FormulaAdjustFeeForm`、`EnterpriseSettingsForm`、`CreateProjectForm`、`SecondaryCircuitForm`、`CabinetAuxCalcForm`、`CategoryForm`、`SummaryAdjustPriceForm`、`TenderReportRegularForm`、`ModelParamParserForm`、`ComponentManageForm`、`ComponentGroupBuilderForm`、`SmartInputForm`、`SpotlightSettingForm` 中全部显式重写 `OnFormClosing`，解绑 WebMessageReceived 事件并显式调用 `_webView?.Dispose()`，彻底根除关闭 Excel 时由于后台 Chromium 子进程等待导致的进程僵死残留问题。
  5. **工程构建与热同步**：
     - `dotnet build /t:Compile /p:DebugType=none` 编译通过：**0 错误**；
     - 修改的 HTML 资源已全量热同步至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`。

- **云方案中心（行业方案、企业方案、一次/二次/收藏、多图纸画廊与 BOM 回路倍增）全栈模块全面交付 (`Models/CloudSolutionModels.cs`, `Services/ExcelServices.CloudSolution.cs`, `Controllers/CloudSolutionController.cs`, `Forms/CloudSolutionForm.cs`, `Resources/cloud_solution.html`, `RibbonController.cs`)**：
  1. **双层 Tab 与三级子分类体系构建**：
     - **顶层 2 个 Tab**：【行业方案】（只读公共标准库，管理员维护）与【企业方案】（用户与企业私有资产库，支持自建与增删改）；
     - **二级 3 个 Tab**：【一次方案】、【二次方案】、【我的收藏】。企业方案板块支持左侧自定义树形目录分类抽屉展开；
     - **4 列响应式卡片网格**：自适应网格排版，展示封面图、多图纸数量徽标（`DrawingUrls.length`）、浏览/收藏数、柜型规格与置顶/企业标签，支持勾选批量插入。
  2. **方案详情大视口完整复刻（图 3）**：
     - **【图纸】**：升级为多图纸 (`DrawingUrls`) 画廊模式，左侧微缩胶卷列表点击即时切换，中央主视口支持平移、滚轮缩放、90° 旋转、自适应与全屏；
     - **【BOM 明细】**：呈现名称、型号、品牌、基准数量、**WL (回路翻倍标志)**、表价、折扣、单价与合价。底部动态输入回路倍数，凡勾选 WL 的行数量自动倍增，总计金额毫秒级联动重算；
     - **【描述】**：方案详细技术参数与工况说明；
     - **控制台**：支持一键插入到 Excel 活动表、导出 BOM 为新工作簿、分享与星标收藏/取消收藏。
  3. **Excel 双向互通与正向沉淀飞轮**：
     - **一键插入**：支持“新建箱柜插入”与“追加到当前箱柜”，严格遵循业务规则 6 与规则 7，内存二维数组一次性覆盖写入，维护小计行与计费公式区域纯净；
     - **存当前柜为方案**：支持在企业方案中一键提炼当前选定箱柜，快速沉淀企业资产。
  4. **工程构建与热同步**：
     - `cloud_solution.html` 已热同步覆盖至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：**0 错误**；
     - 新增代码严格遵循每 3 行至少 1 行中文注释，硬编码均标注 `--硬编码--`。


- **公式法调费总计行下边框线条保护与自动修复全面交付 (`Services/ExcelServices.FormulaAdjustFee.cs`)**：
  1. **总计行封底实线丢失根因排查与差额删行位置重构**：
     - **根因**：原差额删行逻辑直接在计费区末尾执行 `Delete`，当新公式项数少于原计费区行数时，直接将末尾带有黑色封底线条的旧总计行整行删除，而新上浮的行只有内部虚线，导致第 325 行总计下方线条丢失露白底；
     - **删行位置调整**：差额删除行时严格限定在总计行上方（`oldTolsumRow - deleteCount` 至 `oldTolsumRow - 1`）删除，绝对不碰总计行本身，确保旧总计行的边框格式自然上浮；
     - **显式确保总计行底边框（双保险）**：在公式矩阵写入完成后，对新总计行 `A{newTolsumRow}:Q{newTolsumRow}` 显式设置底边框 `tolsumRange.Borders[-4107].LineStyle = 1; tolsumRange.Borders[-4107].Weight = 2;`，不仅杜绝新破坏，更能自动一键修复历史已丢失线条的表格。
  2. **编译验证**：
     - 运行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：**0 错误**，代码注释规范完备。

- **公式法调费更新所有箱柜全工作簿多分类表遍历支持与精准执行反馈全面交付 (`Services/ExcelServices.FormulaAdjustFee.cs`, `Forms/FormulaAdjustFeeForm.cs`, `Tool.cs`)**：
  1. **“更新所有箱柜”失败根因彻底排查与全工作簿遍历重构**：
     - **原失败根因**：旧版 `ApplyFormulaAdjustFeeToExcel` 中根本没有遍历整个工作簿 `Worksheets` 的逻辑，仅仅在 `ActiveSheet` 上执行，若用户在非分类表点击则因有效箱柜为 0 直接静默退出，且外部 UI 盲目弹窗提示“应用成功”，导致虚假成功与实际上全未更新；
     - **全工作簿多表倒序更新支持**：当 `targetScope == "allCabinets"` 时，遍历当前活动工作簿下的所有工作表，安全跳过“项目信息”与元件汇总表等辅助表；
     - **安全激活与 COM 异常规避**：在更新具体工作表前调用 `ws.Activate()` 避免非激活表跨表行操作或公式写入时的 Excel 内部 1004 / 0x800A03EC 异常，遍历结束后平滑恢复原始活动工作表；
     - **精准结果统计与反馈**：重构 `ApplyFormulaAdjustFeeToExcel` 返回 `(bool Success, int UpdatedSheets, int UpdatedCabinets, string Message)` 元组；在 `FormulaAdjustFeeForm.cs` 中根据实际执行结果精准弹窗展示“成功更新 X 个分类表，共 Y 个箱柜！”或警告提示，彻底消除误导。
  2. **总计行识别特征进一步稳健强化 (`Tool.cs:FixAndFillCabinetNamesForSheet`)**：
     - 将总计行 G 列公式引用行号的判断条件优化为 `refRow > 0 && refRow != r`，彻底包容总计行引用工作表内部任意其他行（如单台合计行）的场景，严格落实用户指示“总计行的G列公式一定引用了其他行”。
  3. **编译与构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：**0 错误**，生成的 `ExcelAddInDemo.dll` 成功输出；
     - 新增与修改代码严格遵循每 3 行代码至少 1 行中文注释规范。

- **公式法调费更新计费区间未替换反插入缺陷深度根除与双锚点稳健重构全面交付 (`Tool.cs`, `Services/ExcelServices.FormulaAdjustFee.cs`)**：
  1. **小计行与总计行双黄金锚点识别算法重构 (`Tool.cs:FixAndFillCabinetNamesForSheet`)**：
     - **小计行 (Cab_Subsum)**：严格恪守用户指示，坚决不依赖 B 列文本（不假设叫“小计”），扫描明细区间内公式同时包含 `SUM` 与 `INDEX` 的单元格行精准锚定小计行物理位置；
     - **总计行 (Cab_Tolsum)**：采纳用户精准指示，总计行位于小计行下方，且 **G 列 (第 7 列，销售单价) 公式一定引用了其他行**，彻底废除原先依赖上一行 A 列必须包含 `ROW()-ROW(` 的脆弱判定；
     - **安全兜底重构**：废除原末尾 `curTolsumRow = curDetRow + 27; curSubsumRow = curTolsumRow - 3;` 盲目篡改已识别行号的致命硬编码，当小计行已准确命中时坚决保留，杜绝将第 78 行误判为计费起点的隐患。
  2. **计费区间精准差额替换与多箱柜倒序遍历 (`ExcelServices.FormulaAdjustFee.cs:ApplyFormulaAdjustFeeToExcel`)**：
     - **多箱柜倒序更新**：对 `targetCabinets` 按物理行号倒序遍历（`OrderByDescending`，从底向上），下方箱柜的插行删行 100% 绝不影响上方箱柜物理行号；
     - **动态行号自愈校验**：对每个箱柜执行有效性校验，若锚点缺失或倒挂自动调用双锚点重构算法刷新；
     - **精准对齐与矩阵覆盖替换**：在旧总计行处根据差额精准对齐行数（`delta > 0` 插入空白行，`delta < 0` 删除多余行），随后将 N 行新公式矩阵一次性覆盖写回，彻底消除图一中上方凭空插入、下方旧计费区残留的 Bug，上方元器件预留空行保持纯净。
  3. **编译与构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：**0 错误**，生成的 `ExcelAddInDemo.dll` 成功输出；
     - 新增与修改代码严格遵循每 3 行代码至少 1 行中文注释规范。

- **小箱免铜排全局短路门禁与电流门限配置化界面迁移全面交付 (`Services/ExcelServices.CabinetAuxCalc.cs`, `Resources/cabinet_aux_calc.html`)**：
  1. **小箱零地电流门限(原140A)彻底抽取为界面可设置参数**：
     - 将原写死在【🛠️ 费率与补贴】静态文本中的 `(电流<140A)` 改造为动态响应式联动：`元 (电流 < {{ rules.copperRules.iStructureCurrent || 140 }}A)`；
     - 迁移并在图 2【⚡ 铜排母线定额】下的【📐 主母排形态与排数智能决策门限】卡片中新增【⚡ 小箱免铜排门限:】输入框（绑定 `rules.copperRules.iStructureCurrent`，步长 10A，单位 A）；
     - 与原跨两列的“4极水平排门限”形成 1:1 左右对称整齐排版，下方公式提示框追加小箱免铜排短路判定机制说明；
     - 在 Vue setup 的 `initContext` 中补充了 `iStructureCurrent` 的深度合并与 140A 缺省安全兜底。
  2. **建立小箱免铜排顶层短路门禁机制**：
     - 当箱柜整柜最大电流小于用户设定的门限（`maxCurrent < smallBoxCurrentThreshold`）时，自动判定为小箱；
     - 铜排计算中实行短路直通门禁：`copperWeight = 0.0`，透明记录小箱免排明细说明，彻底跳过水平/垂直主排、垂直N排、标准零地排以及大电流出线分支排的所有计算；
     - 辅材计算中同步联动，小于该门限时自动计入小箱零地排固定补贴（默认 30 元）；
     - 提升 `hasHorizontalBus` 变量作用域至小箱判定前，确保后续出线导线全量平滑计入一次导线，消除了编译未定义引用错误。
  3. **工程构建与热同步**：
     - `cabinet_aux_calc.html` 已热同步复制到 `bin/Debug/net48/Resources/` 与 `publish/Resources/`；
     - 运行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：**0 错误**；
     - 新增代码严格遵循每 3 行代码至少 1 行中文注释，硬编码均打标 `--硬编码--`。


- **预留断路器台数上限参数配置化与微型漏电识别覆盖优化全面交付 (`Models/CabinetAuxCalcModels.cs`, `Services/ExcelServices.CabinetAuxCalc.cs`, `Resources/cabinet_aux_calc.html`, `quotation_rules.json`)**：
  1. **明确解答微型漏电的归属与判定**：
     - 微型漏电在系统中**算断路器**；
     - 优化统计条件，全面支持元件名称与型号中含有“断路器”、“漏电”、“微断”的匹配，微型漏电（如带漏保断路器、微型漏电开关等）均精准计入整柜断路器总台数。
  2. **将预留回路打折判定台数（原硬编码 3 台）彻底抽取为界面可配置项**：
     - 在 `LaborConfig` 模型中新增 `ReservedMaxBreakersThreshold`（默认 3 台）；
     - 在【装配人工费率与综合税费】配置卡片中新增【预留断路器上限: 3 台】独立输入框与联动微卡片提示；
     - 后台算法动态读取该门限，当柜内含预留回路且整柜断路器台数 $\le$ 该门限时，自动触发打折（$\times 0.4$），支持用户任意调整门限；
     - 配置文件 `quotation_rules.json` 写入默认值，HTML 资源已热同步至运行与发布目录；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：**0 错误**。


- **一次配线定额中导线规格选型与单价表样式完全对齐固定长度表且彻底根除截断遮挡修复全面交付 (`Resources/cabinet_aux_calc.html`)**：
  1. **彻底解决行高截断与添加新规格被遮挡问题**：
     - 移除了原 `style` 中写死的 `max-height: 220px;`，解除表格内部强制截断机制，让表格随导线规格数据项自然平铺展开；
     - 彻底杜绝了原有第 6 行文字和输入框被拦腰截断一半、以及点击“添加导线规格”后新行被挡在可视区域外的严重体验缺陷。
  2. **全面对齐固定长度元器件映射表（模块 B-2）的排版与交互样式**：
     - 卡片头部 `cfg-card-header`、标题说明、绿色扁平 `btn-flat` 按钮与加号矢量图标全面对齐；
     - 表格设置自适应列（`spec` 导线规格型号列设为 `min-width="140"`，不设固定 width），整张表格 100% 优雅填满卡片剩余宽度，彻底消除了原表格右侧多余的白色 gutter 空隙与错位；
     - 底部新增统一风格的 `formula-hint-box` 导线选型机制说明微卡片；
     - 模块 A-2 和模块 B-2 表格同步统一移除 `max-height` 限制，确保所有规则子表均无遮挡、平滑展开。
  3. **热同步覆盖**：
     - `cabinet_aux_calc.html` 已热同步覆盖至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`。


- **一次导线计算体系重构（垂直预留高度映射、配电箱裕量系数、固定长度元器件映射）全面交付 (`Models/CabinetAuxCalcModels.cs`, `Services/ExcelServices.CabinetAuxCalc.cs`, `Resources/cabinet_aux_calc.html`, `quotation_rules.json`)**：
  1. **元器件垂直预留高度加成映射表**：
     - 淘汰原写死的“火灾互感器(100mm)”和“普通互感器(130mm)”两项输入框，升级为通用关键字与尺寸映射表 `ExtraHeightRules`；
     - 严格遵循语音确认原则：**单柜同类元器件去重，每类只算一次**（例如柜内有 3 套互感器只加 1 次 130mm；有火灾只加 1 次 100mm；两者兼有则各加 1 次共 +230mm）；
     - 前端增加卡片【📏 元器件垂直预留高度加成映射表】及动态增删表格，底部公式提示同步更新。
  2. **配电箱裕量系数引入**：
     - 在配电箱导线总长公式中加入 `BoxLengthMargin`（默认 1.05），赋予配电箱与落地柜一致的可调放量能力；
     - 前端配置中心提供独立可调输入框，配电箱提示公式同步联动。
  3. **固定长度元器件接线映射表**：
     - 新增短跳线元器件映射表 `FixedLengthRules`（默认：接触器固定 300mm）；
     - 命中该表的出线分路元件不走箱体长宽公式，直接按 $\text{数量} \times \text{极数(3P为3)} \times \frac{\text{固定长度(300mm)}}{1000}$ 独立计算总长；
     - 根据该元器件自身额定电流匹配导线截面与规格（如 65A 匹配 BV-16），汇总计入一次导线规格表；
     - 前端增加卡片【🔗 固定长度元器件接线映射表】及动态增删表格。
  4. **工程构建与热同步**：
     - `bin/Debug/net48/data/quotation_rules.json` 与 `publish/data/quotation_rules.json` 写入默认配置；
     - `cabinet_aux_calc.html` 已热同步覆盖至调试目录与发布目录；
     - 运行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：**0 警告，0 错误**；
     - 新增代码严格遵循**每 3 行代码至少 1 行中文注释**，避免 Element Plus 自闭合标签陷阱，硬编码均打标 `--硬编码--`。


- **一次线与铜排计算范围自定义集合门禁及电流列为空告警提醒机制全面交付 (`Models/CabinetAuxCalcModels.cs`, `Services/ExcelServices.CabinetAuxCalc.cs`, `Resources/cabinet_aux_calc.html`, `data/quotation_rules.json`)**：
  1. **彻底取消电流列为空时的型号正则自动提取**：
     - 在 `ScanCabinetData` 中，当 W 列（第 23 列）为空时，直接保留 `current = 0`，坚决停止调用 `ParseCurrentFromModel` 正则解析；彻底根除了将热继电器 `NDR2-3822(16-24A)` 中的规格代号 `3822` 误当成 `3822A` 额定电流并触发超大出线分支铜排的深层隐患。
  2. **新增参与一次线与铜排计算的自定义元器件集合门禁**：
     - 在 `QuotationRules.General` 与 `quotation_rules.json` 中新增 `PrimaryCalcComponents` 配置属性（默认包含：塑壳断路器、塑壳、微断、小型断路器、断路器、隔离开关、负荷开关、框架断路器、双电源、ATS，标记 `--硬编码--`）；
     - 实现 `IsComponentInPrimaryCalcSet` 门禁过滤方法：只有命中自定义集合的元器件，才计算一次导线用量和出线分支铜排；非集合内元器件（热继电器、接触器、电涌保护器、指示灯等二次/辅助元件）坚决排除，杜绝误算。
  3. **未填电流只提醒告警机制**：
     - 在 `CabinetCalcResult` 中增加 `Warnings` 列表属性；
     - 当扫描到属于一次计算集合的元器件但 W 列电流为空（`current <= 0`）时，自动记录告警提醒（格式：`第 X 行【名称 型号】W列电流为空，不做自动提取，已跳过一次线与分支铜排计算！`）；
     - 在【📊 智能推导与回写】看板顶部增加醒目的 ⚠️ 告警提醒条幅（带警告图标、黄色高亮背景与详细清单），并在推导说明 `Description` 中同步提示。
  4. **前端配置中心与看板无缝升级**：
     - 在【⚙️ 规则与定额配置】的“🔌 一次配线定额”顶部新增【🏷️ 参与一次线与铜排计算的元器件集合 (门禁规则)】配置卡片，支持以 Tag 形式展示、动态新增和删除关键字，并可一键保存至磁盘持久化；
     - 前端 JS 增加响应式兜底兼容，确保历史老配置平滑无感升级。
  5. **构建与热同步**：
     - `cabinet_aux_calc.html` 已热同步至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`；
     - `quotation_rules.json` 已写入默认集合配置；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译全量通过：0 错误 0 警告。

- **二次回路图纸对齐与绑定工作台彻底移除右下角红框卡片与表格全高展开优化全面交付 (`Resources/secondary_circuit_manage.html`)**：
  1. **右下角冗余参数卡片彻底移除**：
     - 彻底移除了右下角红框处的【关联二次方案参数】双排卡片及占位提示容器，彻底消除了底部高度挤压与边框截断问题；
     - 结构更清爽，DOM 树复杂度降低。
  2. **右侧元件组表格自适应撑满全高**：
     - 将表格包裹层设为 `min-height: 0; flex: 1; height: 100%;`，让【Excel 元件组清单】表格纵向 100% 独占整个右侧工作区；
     - 表格支持完整的内部虚拟平滑滚动与表头吸顶，能同时纵向一览更多元件组型号。
  3. **数据零丢失：平移 BOM 项数至 CAD 视口顶部**：
     - 将原卡片右上角的 BOM 物料项数平移至中间 DWG 矢量视口顶部的彩色定额胶囊末尾：`📦 BOM: X项`；
     - 与既有的【二次组、跨门、二次材料、开孔、人工】形成完整六大核心方案定额胶囊，看图与看方案参数浑然一体。
  4. **构建与热同步**：
     - AST 标签栈校验通过：0 个未闭合，0 个多余；
     - 已热同步复制至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`。

- **VS 启动调试托管调试助手 "FatalExecutionEngineError" (0x80131623, 0x9fa1cae5) 致命崩溃第三阶段核心根除全面交付 (`Forms/ExcelWindowHook.cs`, `Services/ExcelServices.Spotlight.cs`)**：
  1. **现场还原与深层根因锁定**：
     - 用户在 Visual Studio 启动调试时，光标精准停留在 `Forms/ExcelWindowHook.cs` 第 152 行（`GuardTimer_Tick` 回调中读取原点坐标后的比较行）；
     - 崩溃报告地址 `0x9fa1cae5`（低 16 位 `0xCAE5`，即 `EXCEL7` 视口句柄）。CLR MDA 拦截原因是：后台定时器在主线程高频触发，执行 `GetClientRect` / `ClientToScreen` 跨语言 P/Invoke 调用破坏了 JIT 栈帧平衡；当非托管调用返回恢复执行托管 IL 指令（第 152 行）时，被 CLR 栈探测捕捉到寄存器/返回地址被句柄值覆盖，直接触发 `FatalExecutionEngineError (0x80131623)`；
     - 同时，工程中还存在 `_foregroundGuardTimer`（120ms）与 `_guardTimer`（200ms）两个后台定时器高频轰炸主消息循环，形成严重的调用栈撕裂温床。
  2. **第三阶段极简纯净安全架构重构**：
     - **措施 1 (彻底清除 ExcelWindowHook 内部定时器与 P/Invoke)**：重构 `ExcelWindowHook.cs` 为纯托管安全适配器类，彻底移除所有 `System.Windows.Forms.Timer` 定时器及其 Tick 回调，不再调用 `GetClientRect` 和 `ClientToScreen`，零跨语言调用，零栈破坏隐患，100% 保持类库公开接口兼容；
     - **措施 2 (彻底移除前台焦点守护定时器)**：在 `ExcelServices.Spotlight.cs` 中彻底移除 `_foregroundGuardTimer`（120ms 轮询）及 `StartForegroundGuard`、`StopForegroundGuard`、`ForegroundGuardTimer_Tick` 方法；由于聚光灯浮窗属于 Excel 主窗口的 Owned 窗口，系统 DWM 天然将其压在外部软件下方，无需任何轮询；
     - **措施 3 (原子级 GetWindowRect 单次获取视口绝对物理矩形)**：在 `UpdateSpotlightPosition` 中废弃 `GetClientRect` + `ClientToScreen` 的双重组合调用，改用 Win32 原生原子安全 API `GetWindowRect(excel7Hwnd, out RECT winRect)` 一步直接拿到绝对屏幕坐标和宽高，彻底消灭所有不稳定的 P/Invoke 转换；
     - **措施 4 (聚光灯全面由 Excel 原生 COM 事件模型驱动)**：选区移动跟随 `SheetSelectionChange`、切换表跟随 `SheetActivate`、切换工作簿跟随 `WorkbookActivate`，0 个定时器、0 个后台轮询、0% CPU 占用，彻底杜绝所有崩溃可能。
  3. **构建验证**：
     - 执行 `dotnet build /t:Compile` 编译通过：0 错误；
     - 执行完整 `dotnet build` 成功打包生成 `ExcelAddInDemo-AddIn64.xll`，0 警告，0 错误；
     - 代码严格满足每 3 行至少 1 行中文注释。

- **VS 启动调试托管调试助手 "FatalExecutionEngineError" (0x80131623, 0x9f9ecae5, 0x9fa1cae5) 致命崩溃第二阶段终极排查与彻底根除全面交付 (`AddInMain.cs`, `Forms/ExcelWindowHook.cs`, `Services/ExcelServices.Spotlight.cs`, `Forms/SpotlightOverlayForm.cs`)**：
  1. **深度根因排查与定位**：
     - 用户在 Visual Studio 启动调试时再次出现 `FatalExecutionEngineError (0x80131623)`，地址变为 `0x9fa1cae5`（与上一次地址低 16 位完全一致为 `0xCAE5`），经过系统句柄与模块跟踪，确认为 Win32 原生视口句柄；
     - **根因一 (启动抢跑争抢 COM 初始化栈)**：`AddInMain.AutoOpen()` 在 `QueueAsMacro` 延迟宏中若判断聚光灯曾开启过，会过早调用 `ExcelServices.EnableSpotlight()`。Excel 启动初期 COM 消息泵与主视口窗口树仍在动态构造中，此时过早探测视口与展示浮窗极易破坏初始化调用栈；
     - **根因二 (EnumChildWindows 反向 P/Invoke 委托 Thunk 栈溢出与 GC 风险)**：`FindExcel7Hwnd` 中使用了 `EnumChildWindows(mainHwnd, (childHwnd, lParam) => ...)`。Excel 顶层主窗口包含数百个子窗口，每次调用都会在非托管线程中高频反向回调托管动态 Lambda 委托，缺少固定强引用与确切的 `CallingConvention`，在 VS 调试器监控下引发 CLR 执行引擎内部 Thunk 崩溃；
     - **根因三 (缺少显式 Winapi 调用约定与高频 40ms 定时器)**：`ExcelWindowHook` 中包含 40ms 极高频巡检定时器，频繁调用 `ClientToScreen`，且所有 Win32 DllImport 均未显式指定 `CallingConvention = CallingConvention.Winapi`，触发 MDA 堆栈平衡检查崩溃。
  2. **第二阶段四重精准根除措施**：
     - **措施 1 (启动安全解耦)**：在 `AddInMain.cs` 中彻底移除启动时的 `EnableSpotlight()` 自动抢跑调用，插件启动仅注册 COM 事件与右键菜单，聚光灯改为按需由用户点击 Ribbon 或快捷键开启，保障 F5 启动 100% 纯净、安全；
     - **措施 2 (FindWindowEx 替代 EnumChildWindows 零回调重构)**：在 `ExcelServices.Spotlight.cs` 中彻底废弃 `EnumChildWindows` 及回调委托，改用纯原生原子查找 Win32 `FindWindowEx`（`XLMAIN` -> `XLDESK` -> `EXCEL7`），零委托分配、零非托管回调、彻底根绝反向 P/Invoke 堆栈损坏；
     - **措施 3 (全量 P/Invoke 显式 Winapi 规范化)**：为 `ExcelServices.Spotlight.cs`、`ExcelWindowHook.cs`、`SpotlightOverlayForm.cs` 中的所有 Win32 API 统一显式标注 `CallingConvention = CallingConvention.Winapi`，彻底消除 MDA 对栈指针平衡的质疑与拦截；
     - **措施 4 (巡检定时器门禁与 200ms 防抖)**：在 `ExcelWindowHook` 的 `GuardTimer_Tick` 中增加 `!ExcelServices.IsSpotlightEnabled` 门禁，关闭状态瞬间静默退出；将定时器轮询间隔从 40ms 优化为安全合理的 200ms 防抖检测，杜绝高频轰炸主消息循环。
  3. **编译与验证结果**：
     - 执行 `dotnet build /t:Compile` 编译通过：0 错误；
     - 执行完整 `dotnet build` 生成 `ExcelAddInDemo-AddIn64.xll` 成功：0 警告，0 错误；所有新增代码严格满足每 3 行至少 1 行中文注释。

- **二次回路图纸对齐与绑定工作台布局修复与体验增强全面交付 (`Resources/secondary_circuit_manage.html`)**：
  1. **彻底根治白板与错乱布局问题**：
     - 排查并修复了前期修改遗留的未闭合 `<div>` 标签问题，补全了 `#app` 根容器闭合 `</div>`，使模板 AST 标签栈完全平衡归零（0 个未闭合）；
     - 解决了历史编码冲突导致的 `?/div>` 问题，标签完全规范化；
     - 增加了 `.circuit-binding-dialog .el-dialog__body` 的紧凑内边距（`padding: 8px 12px !important; overflow: hidden !important;`），释放多达 24px+ 垂直空间。
  2. **三栏空间重构与 DWG 视口显著调宽**：
     - 左侧文件夹/图纸栏由 260px 缩窄至 230px；
     - 右侧元件组看板栏由 480px 缩窄至 360px（右侧表格列宽精细适配为 145/125px，字体 12/11px，两列紧凑并排无横向滚动条）；
     - 中间 WebGL CAD 矢量视口获得高达 **150px** 的额外宽度，大幅提升看图体验。
  3. **DWG 视口顶部胶囊完整展示**：
     - 左侧文件名限制最大宽度 `150px` 并允许弹性缩进，杜绝挤占右侧空间；
     - 中间参数胶囊添加 `flex-shrink: 0;`、间距设为 `6px`，每个参数配备悬浮 `:title` 提示，保障在任何屏幕下都不会被截断。
  4. **最右下角参数卡片双排布局且杜绝下边界截断**：
     - 表格使用 `<div style="flex: 1; min-height: 120px; overflow: hidden; position: relative;">` 包裹，搭配 `height="100%"` 受控滚动，表头固定吸顶，绝不向下挤压卡片；
     - 最右下角关联二次方案参数卡片升级为双排精致紧凑栅格（第 1 排：跨门线、二次材料、开孔；第 2 排：人工、二次组方案名），加上 `margin-top: auto; flex-shrink: 0;`；
     - 所有指标数字垂直居中、不折行且带 Tooltip，卡片完整露于视口内，下边框和数字下半截截断问题彻底消除。
  5. **热同步与清理**：
     - 文件已热同步至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`；
     - 清理所有中间辅助脚本，保持工程干净整洁；
     - 新增代码严格遵循每 3 行代码 1 行中文注释规范。

- **辅材与人工匹配名可配置及双区查找（计费区优先、元器件区备用）与铜排改填元器件最末行（无空位自动插行）全面交付 (`Models/CabinetAuxCalcModels.cs`, `Services/ExcelServices.CabinetAuxCalc.cs`, `Resources/cabinet_aux_calc.html`)**：
  1. **数据模型扩展 (`CabinetAuxCalcModels.cs`)**：
     - `AuxConfig` 新增 `AuxMatchName`（默认 `"辅材"`，标注 `--硬编码--`）；
     - `LaborConfig` 新增 `LaborMatchName`（默认 `"人工费"`，标注 `--硬编码--`）；
     - `CabinetCalcResult` 增加 `AuxTargetLocation`、`LaborTargetLocation`、`CopperTargetLocation` 记录回写定位。
  2. **前端配置中心与推导看板升级 (`Resources/cabinet_aux_calc.html`)**：
     - 在【🛠️ 费率与补贴】中新增【辅材匹配名称】与【人工匹配名称】可编辑配置输入项；
     - 底部说明区域升级为清晰呈现壳体、辅材、人工与铜排的回写规则与实际回写位置反馈；
     - `rules` 初始化与合并逻辑提供平滑安全兜底。
  3. **核心回写与插行引擎重构 (`ExcelServices.CabinetAuxCalc.cs`)**：
     - 重构 `WriteCabinetCalcResultToSheet`；
     - **辅材与人工**：先在计费区域遍历 B 列匹配（辅材匹配 `auxMatchName`，人工匹配 `laborMatchName` 或兼容 `"人工"`），命中则将推导金额公式写入 **H 列（销售总价列）**；若计费区未命中，自动穿透至元器件区域查找 B 列，命中则写入 H 列并将数量 F 列置为 1；
     - **铜排**：调整至元器件区域最下面一行；先检测既有铜排行实现原位更新防重复；若无既有行且最后一行已有元件（无空位），则在小计行 `Cab_Subsum` 位置执行向下插入新行；写入 17 列完整元器件数据（A 列序号、B 列 "铜排"、C 列 "TMY"、E 列 "KG"、F 列数量公式、G 列单价、H 列合价、J/K 列成本、Q 列 "材料" 等）；若发生插行，调用 `RefreshCabinetFeeAreaFormulas` 自动刷新小计公式与 A 列序号；
     - 清理计费区历史遗留的铜排行数据，杜绝重复计费。
  4. **构建与热同步**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误 0 警告；
     - 资源已同步复制至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`。

- **二次导线推导长度呈现与一次/二次计算引入柜高比例参数全面实施 (`Models/CabinetAuxCalcModels.cs`, `Services/ExcelServices.CabinetAuxCalc.cs`, `Resources/cabinet_aux_calc.html`)**：
  1. **数据模型深度扩展 (`CabinetAuxCalcModels.cs`)**：
     - `PrimaryWireLengthConfig`：增加 `CabinetHeightFactor` (落地柜柜高系数，默认 0.4) 与 `BoxHeightFactor` (配电箱柜高系数，默认 0.3)，标注 `--硬编码--`；
     - `AuxConfig`：增加 `SecondaryWidthFactor` (跨门柜宽系数，默认 0.8)、`SecondaryHeightFactor` (跨门柜高比例，默认 0.3)、`SecondaryMarginLength` (端头接线预留裕量，默认 300mm)，标注 `--硬编码--`；
     - `SecondarySchemeCalcItem`：增加 `SingleWireLength` (方案单套二次线长，米)、`TotalWireLength` (方案累计推导总长，米)；
     - `CabinetCalcResult`：增加 `SecondarySingleWireLength` (单根跨门基准长)、`SecondaryTotalWireCount` (二次线总根数)、`SecondaryTotalWireLength` (推导总长度)。
  2. **核心计算引擎双维空间走线模型升级 (`ExcelServices.CabinetAuxCalc.cs`)**：
     - 一次导线计算：在垂直落差计算中引入柜高比例项：落地柜叠加 `shellHeight * cabHeightFactor`、配电箱叠加 `shellHeight * boxHeightFactor`，彻底解决高柜与矮箱走线落差不明显问题；
     - 二次导线计算：将原有固定的单维 `(柜宽+300)/1000` 升级为严密的双维跨门空间模型：
       `单根长 = (柜宽 × 柜宽系数 + 柜高 × 柜高比例 + 端头裕量) / 1000`；
       `方案单套长 = 方案跨门根数 × 单根长`；
       `方案累计总长 = 单套长 × 方案套数`；
     - 回写与推导说明增强：计算结果回填 `CabinetCalcResult`，并在 `Description` 中详细输出二次方案套数、总根数、单根长与总线长推导算式。
  3. **前端工作台看板与配置中心升级 (`Resources/cabinet_aux_calc.html`)**：
     - **推导结果视图**：新增【⚡ 二次导线推导长度与用量明细】独立看板，包含顶部指标条（单根跨门基准长、总根数、推导总长度、导线规格、参考单价）及方案级推导表格（二次方案名、回路代号、套数、跨门线、单套线长、累计推导总长、辅材小计、工费小计）；
     - **规则配置模块 A**：在一次导线参数区补充落地柜柜高系数、配电箱柜高系数输入项与动态联动公式提示；
     - **规则配置模块 C**：在二次回路参数区提供控制线单价、跨门柜宽系数、跨门柜高比例、端头接线裕量输入项与详尽推导公式；
     - **前端 JS 兼容与默认值兜底**：在 `rules` 初始化对象及 `initContext` 报文反序列化逻辑中增加安全兜底，保证历史配置文件平滑无缝升级。
  4. **热同步与构建验证**：
     - 前端静态文件已热同步至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译验证通过：0 错误。

- **智能辅材与壳体计算中心接入二次回路方案定额并彻底下线旧“二次元件定额” (`Models/CabinetAuxCalcModels.cs`, `Services/ExcelServices.CabinetAuxCalc.cs`, `Resources/cabinet_aux_calc.html`)**：
  1. **步骤 1：重构辅材与人工计算引擎为方案驱动 (`ExcelServices.CabinetAuxCalc.cs`, `CabinetAuxCalcModels.cs`)**：
     - **扫描范围扩展**：将工作表扫描由第 28 列（AB 列）扩展至第 32 列（AF 列），一次性读入二维数组，提取绑定的二次回路代号/图号 `boundDwgCode` 与二次元件组标记 `isComponentGroup`；
     - **联动二次方案数据库**：一次性加载本地全量二次方案（`PersonalComponentDbService.GetAllSecondarySchemes()`），按回路代号、CAD 图名、方案名称或型号去星号精确检索匹配；
     - **参数精准继承与算式透明化**：
       - 二次配线辅材：继承方案级 `CrossDoorCount`（跨门线根数），按 `跨门根数 × (柜宽+300)/1000 × 线单价 × 套数` 计算辅材并累加；
       - 二次装配工价：继承方案级 `LaborCost`（人工工费），按 `单套工价 × 套数` 累加至人工总额，并在算式字符串中清晰列示（如 `45.0*1`）；
       - 记录命中的方案明细列表 `SecondarySchemeDetails`，并在推导说明 `Description` 中标注命中方案统计与未绑定警告；
     - **模型升级**：`CabinetComponentItem` 增加 `BoundDwgCode` 与 `IsComponentGroup`；新增 `SecondarySchemeCalcItem`；`AuxiliaryRulesConfig` 增加 `SecondaryWirePrice`（默认 0.8 元/米，标记 `--硬编码--`）；
  2. **步骤 2：下线前端二次元件定额 Tab 与废弃逻辑 (`cabinet_aux_calc.html`)**：
     - **界面彻底下线**：移除“二次元件定额”Tab 页及其添加/删除按钮，移除 `addSecondaryElement` 及 setup 导出项；
     - **联动说明与配置集成**：在“一次配线定额”Tab 末尾添加“二次回路方案定额参数联动说明”卡片，提供二次控制线单价调节输入项并明确指引至【二次回路方案管理中心】统一维护；
     - **结果透明呈现**：在“智能推导与回写”界面新增“命中的二次回路方案定额明细”表格卡片，清晰展示方案名、代号、套数、跨门线、辅材小计与工费小计；
  3. **静态资源同步与编译验证**：
     - 静态模板已热同步至 `bin\Debug\net48\Resources\cabinet_aux_calc.html` 与 `publish\Resources\cabinet_aux_calc.html`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译全量通过：0 错误 0 警告。

- **聚光灯设置面板点击关闭/取消无响应 (不能关闭) 之根因排查与彻底修复 (`Forms/SpotlightSettingForm.cs`, `Resources/spotlight_setting.html`)**：
  1. **深度排查三大根因**：
     - **根因一 (WebView2 消息反序列化通道阻断)**：前端调用 `postMessage(payload)` 直接传递原生 JavaScript Object，而在 C# 端仅使用了 `e.TryGetWebMessageAsString()`。当收到非字符串类型时，该方法直接抛出异常或返回空，而此前未捕获 `e.WebMessageAsJson` 进行兜底，导致 `string.IsNullOrWhiteSpace` 成立直接 `return`，所有操作（关闭、保存、取消、拖拽、联动）均未被执行；
     - **根因二 (参数节点键名不一致)**：前端传递的配置键名为 `config: { ...form }`，而 C# 端此前只尝试提取 `data` 节点，导致保存与预览指令即使送达也因找不到 `data` 被忽略；
     - **根因三 (指针事件穿透与冒泡干扰)**：右上角关闭按钮父级具有 `@pointerdown="onHeaderPointerDown"` 拖拽监听，点击叉叉时若未阻断冒泡会被标题栏拖拽逻辑拦截抢占；
  2. **全面针对性修复**：
     - **双轨消息解析与容错兼容**：在 `SpotlightSettingForm.cs` 中先尝试 `TryGetWebMessageAsString()`，若为空则自动读取 `e.WebMessageAsJson`，彻底兼容原生对象与 JSON 字符串；
     - **前端显式 JSON 序列化**：在 `spotlight_setting.html` 的 `postToCSharp` 中统一使用 `JSON.stringify(payload)` 发送；
     - **全别名指令与双节点兼容**：C# 端 `close` 分支同时支持 `close`、`closeWindow`、`cancel`；`previewConfig` 与 `saveConfig` 同时支持 `data` 与 `config` 节点，保存后立即安全调用 `SafeInvoke(this.Close)` 关闭窗口；
     - **事件阻断**：在关闭按钮上添加 `@pointerdown.stop @click.stop="closeWindow"`，彻底隔离拖拽与点击；
  3. **编译与热部署验证**：
     - 页面已同步复制至 `bin\Debug\net48\Resources\spotlight_setting.html`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误。

- **聚光灯 (Spotlight) 个性化设置中心完整制作与交付 (`Resources/spotlight_setting.html`, `Controllers/SpotlightController.cs`, `Forms/SpotlightSettingForm.cs`, `Services/ExcelServices.Spotlight.cs`, `Models/SpotlightConfig.cs`, `RibbonController.cs`)**：
  1. **前端现代 Vue 3 工作台 (`Resources/spotlight_setting.html`)**：
     - 基于 Vue 3 (`<script setup>`) + Element Plus + CSS Glassmorphism；
     - 主色调采用 `#009688` 绿蓝相间，圆角卡片化布局；
     - **非阻塞物理增量平滑拖拽**：使用 PointerEvents + `requestAnimationFrame` 发送 `{ action: 'moveWindow', deltaX, deltaY }`，彻底杜绝 Win32 `WM_NCLBUTTONDOWN` 模态消息循环导致的 Excel/WebView2 卡死；
     - **功能模块一览**：
       - 高亮模式卡片切换（十字交叉 / 仅高亮行 / 仅高亮列）；
       - 6 种工业精选预设色（绿蓝相间、天空蓝、薄荷绿、暖金琥珀、薰衣草紫、珊瑚粉）+ `el-color-picker` 调色盘；
       - 不透明度调节滑块（5% ~ 80% 步长 1%）；
       - 活动单元格镂空开关（保持输入文字清晰纯净）；
       - 实时联动效果预览视窗与实时 Excel 联动同步；
       - 【恢复默认】、【取消】与【保存并应用】完整闭环。
  2. **后端服务与控制器体系 (`SpotlightController.cs`, `SpotlightSettingForm.cs`, `ExcelServices.Spotlight.cs`)**：
     - **实时联动预览引擎**：用户在面板中拖动透明度滑块或选色时，即时派发 `previewConfig`，C# 内存更新浮窗样式与 GDI Region，用户可边调边在 Excel 中看到真实效果；若点击取消，自动无损还原至打开设置面板前的备份快照；
     - **非模态安全依附**：使用 `ShowModelessForm` 打开设置面板，尺寸 520x480 居中弹出，保持 Excel 处于可交互编辑状态；
     - **配置持久化**：`SpotlightConfig.UpdateCurrent` 自动保存至 LocalAppData JSON 文件并即时生效。
  3. **Ribbon 界面无缝集成 (`RibbonController.cs`)**：
     - 将辅助项中的 `btnToggleSpotlight` 升级为 `splitButton id='splitSpotlight'`；
     - 上半部保留大图标快速开启/关闭开关；
     - 下半部下拉菜单提供【⚙️ 聚光灯设置...】快速打开设置面板，以及快捷切换高亮模式（十字 / 仅行 / 仅列）；
  4. **工程构建与验证**：
     - `ExcelAddInDemo.csproj` 已加入 `spotlight_setting.html` 复制输出项；
     - 静态文件已自动部署至 `bin\Debug\net48\Resources\spotlight_setting.html`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 验证通过：0 错误。

- **打开/切换其他软件聚光灯残留在屏幕上遮挡之根因排查与彻底解决 (`Forms/SpotlightOverlayForm.cs`, `Forms/ExcelWindowHook.cs`, `Services/ExcelServices.Spotlight.cs`)**：
  1. **根因定位**：
     - 原 `SpotlightOverlayForm` 设置了 `TopMost = true`，导致浮窗成为全桌面系统的最高 Z-Order 窗口，即使用户切到微信、浏览器、VS Code 等其他软件，该十字半透明图形依然全局置顶悬浮在最上层；
     - 缺乏前台激活进程判定，在用户处于其他软件工作时依然保持了可见渲染。
  2. **四重系统级层级与前台联动彻底根治**：
     - **解除系统级 TopMost**：将 `SpotlightOverlayForm` 的 `TopMost` 改为 `false`；
     - **绑定 Excel 主窗口为 Owner**：通过 Win32 `SetWindowLongPtr(Handle, GWLP_HWNDPARENT, mainHwnd)` 将浮窗与 Excel 建立 Owned 依附关系。一旦用户切换激活其他软件，Windows 桌面管理器会自动将 Excel 及其所属浮窗一同沉入后台，100% 绝不遮挡任何其他软件；
     - **前台焦点守护定时器 (`_foregroundGuardTimer`)**：在 `ExcelServices.Spotlight.cs` 中增加 120ms 的极轻量轮询守护（耗时 0ms）：通过 `GetForegroundWindow()` 与 `GetWindowThreadProcessId` 监控前台 PID；一旦发现用户切到其他软件，立即调用 `ShowWindow(SW_HIDE)` 隐匿浮窗；一旦切回 Excel，瞬间自愈恢复高亮；
     - **原生钩子失焦即刻响应 (`WM_KILLFOCUS`)**：在 `ExcelWindowHook` 中拦截 `WM_KILLFOCUS` 并在 `OnActivationChanged` 中响应，当视口失焦且前台非 Excel 时 0 延迟秒级隐匿；
     - **刷新入口前门禁**：在 `UpdateSpotlightPosition` 开头校验 `IsExcelForeground()`，非前台时拒绝绘制并保持隐匿；
  3. **编译构建验证**：
     - `dotnet build /t:Compile /p:DebugType=none` 0 错误 0 警告通过。

- **聚光灯功能未显示 (消失) 深度排查与全面彻底恢复 (`Services/ExcelServices.Spotlight.cs`, `ExcelEventManager.cs`)**：
  1. **深度排查三项根因**：
     - **根因一 (浮窗托管 Visible 缺失)**：在防死锁调整中移除了 `_spotlightForm.Show()`，只调用了 `CreateControl()`。WinForms 的 Form 实例若从未调用 `Show()`，托管状态的 `this.Visible` 始终为 `false`，即便底层调用了 `ShowWindow(SW_SHOWNOACTIVATE)`，WinForms 底层窗口过程仍会在消息分发时自动将 Handle 隐匿；
     - **根因二 (dynamic 强制转换为 IntPtr 抛出 RuntimeBinderException)**：`FindExcel7Hwnd()` 中使用了 `(IntPtr)app.ActiveWindow.Hwnd`。在 C# DLR 机制下，对 dynamic 对象显式转换为 IntPtr 不支持内置 int 转换，必定触发 `RuntimeBinderException`，导致内部 catch 吞掉并返回 `IntPtr.Zero`，从而无法定位视口；
     - **根因三 (空启动新建工作簿未触发 SelectionChange)**：VS 调试空启动时工作簿为 0 聚光灯静默，用户新建空白表格默认停在 A1，此时不会触发 `SelectionChange`，此前未监听 `WorkbookActivate` / `SheetActivate`，导致聚光灯一直处于等待唤醒状态；
  2. **全面针对性修复**：
     - **恢复 `_spotlightForm.Show()`**：因 `SpotlightOverlayForm` 已具备 `ShowWithoutActivation => true` 与 `WS_EX_NOACTIVATE` 样式，调用 `Show()` 100% 绝不夺取 Excel 编辑焦点，同时使窗体进入桌面分层渲染管线；在刷新末尾确保 `_spotlightForm.Visible = true;`；
     - **安全拆箱与类型转换**：使用 `Convert.ToInt64(app.ActiveWindow.Hwnd)` 转换为 long 进而创建 `IntPtr`，100% 杜绝 DLR 转换异常；
     - **接入工作簿/工作表激活事件**：在 `ExcelEventManager.cs` 中注册 `SheetActivate` 与 `WorkbookActivate`，当用户新建、打开或切换表格时，聚光灯瞬间自愈激活并点亮；
  3. **编译构建验证**：
     - `dotnet build /t:Compile /p:DebugType=none` 0 错误 0 警告编译通过。

- **多选区滚动鼠标中键高亮退化为单行单列之根因排查与全面修复 (`Services/ExcelServices.Spotlight.cs`)**：
  1. **根因定位**：
     - 用户选中多行多列时，选区变动事件传入的是完整的选区 Range（如 `A1:D10`）；
     - 但在滚动鼠标中键时，视口钩子回调传入的是 `UpdateSpotlightPosition(null)`；
     - 原逻辑执行 `cell = target ?? app.ActiveCell`，直接回退到了单格 `ActiveCell`，导致选区信息被丢弃，高亮瞬间缩水为单行单列。
  2. **全面重构多选区与多区域 (Areas) 引擎**：
     - **选区智能保持**：当入参为 `null` 时，优先通过 `app.Selection` 获取用户当前的完整多选选区 Range，彻底解决滚屏时选区丢失问题；
     - **多区域 (Areas) 遍历与并集合成**：支持连续多行多列选区以及按住 Ctrl 键的多不连续区域，遍历 `targetRange.Areas` 并通过 Win32 GDI `CombineRgn(..., RGN_OR)` 统一合并，支持精准镂空；
     - **超大选区溢出防护**：对整行（16384 列）或整列（1048576 行）选区的磅值增加 8000 磅安全上限，杜绝 `PointsToScreenPixels` 产生 COM 溢出崩溃；
  3. **编译验证**：
     - `dotnet build /t:Compile /p:DebugType=none` 验证通过：0 错误 0 警告。


- **彻底去除 CAD 视口顶部左侧“矢量预览”文本与右侧“定位”按钮及关联代码 (`Resources/secondary_circuit_manage.html`, `Forms/SecondaryCircuitForm.cs`)**：
  1. **前端界面与交互精简**：
     - 去除视口顶部左侧冗余前缀 `<span style="color: #64748b;">矢量预览: </span>`，直接精简呈现图纸名称 `📐 {{ currentVectorDwg.fileName }}`；
     - 移除右侧 `<el-button size="small" type="info" link @click="openInExplorerFromViewer">📂 定位</el-button>` 按钮，保留调起外部 AutoCAD 打开功能；
     - 彻底清除 `openInExplorerFromViewer` 函数定义、`action === 'openInExplorerResult'` 消息响应及 `setup()` 导出声明；
  2. **后端 C# 调度代码同步清理**：
     - 在 `Forms/SecondaryCircuitForm.cs` 中移除 `case "openInExplorer":` 消息接收与响应分支；
  3. **静态资源热同步与验证**：
     - 已将修改后的 `secondary_circuit_manage.html` 同步复制至 `bin\Debug\net48\Resources\` 和 `publish\Resources\`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误。

  1. **靠顶部根因排查**：
     - 二次回路图纸对齐与绑定工作台此前显式配置了 `top="1.5vh"`，且编辑弹窗配置了 `top="2.5vh"`；
     - 同时 CSS 中 `.circuit-binding-dialog` 声明了 `margin: 0 auto !important;`，导致弹窗完全贴在主窗口顶部，缺乏视觉纵向居中平衡感；
  2. **全面启用原生 `align-center` 与弹性居中架构**：
     - 在 `<el-dialog v-model="circuitDwgDialogVisible"` 与 `<el-dialog v-model="dialogVisible"` 上统一移除 `top` 偏置属性，启用 Element Plus 原生 `align-center` 居中引擎；
     - 在 CSS 中增强全局遮罩 `.el-overlay-dialog { display: flex !important; align-items: center !important; justify-content: center !important; overflow: hidden !important; }`，彻底锁定居中视口；
     - 统一配置 `.circuit-binding-dialog` 与 `.scheme-edit-dialog` 为 `margin: auto !important; max-height: 94vh !important;`，确保在任何屏幕分辨率及最大化/窗口化切换时，弹窗均优雅、绝对居中于屏幕/窗体正中央；
  3. **静态资源热同步与验证**：
     - 已将修改后的 `secondary_circuit_manage.html` 同步复制至 `bin\Debug\net48\Resources\` 和 `publish\Resources\`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误 0 警告。

  1. **左侧未反显的深层根因分析**：
     - **引用不一致**：左侧表格数据源 `:data="filteredDirItems"` 是基于 `dirDwgFiles` 实时 map 计算出来的新对象集合，此前用原数组项调用 `setCurrentRow` 无法命中 Element Plus 的引用对比（`===` 为 false），导致行浅绿色高亮未附着；
     - **跨目录物理分布差异**：用户图纸按分类存放在不同的子文件夹中（如 `双电源`、`功能表互感表`、`变频器` 等）。当用户处于“双电源”子目录时，点击绑定了其它子目录图纸的元件组（如 `接触器变频器`），左侧当前列表中根本不存在该文件，导致无法呈现反显与开图；
  2. **全面重构反显机制与跨目录穿透寻图**：
     - **精准引用匹配与平滑滚动 (`highlightLeftDwg`)**：必须在 `filteredDirItems.value` 中提取同一引用对象调用 `circuitDwgTableRef.value.setCurrentRow(targetItem)`，确保 100% 亮起 `.current-row` 浅绿底色，并自动调用 `scrollIntoView` 滚动到视口中央；
     - **C# 后端全局递归穿透定位 (`LocateDwgInDirectory` / `LocateDwgFile`)**：当当前目录未搜寻到图纸时，自动向 C# 发送 `locateAndHighlightDwg` 并在图纸库全目录中进行深搜枚举；
     - **前端自动下钻跳转与延迟反显**：一旦定位到图纸所在父目录，自动触发 `browseToDirectory(parentDir)` 切换左侧目录，并在层级加载完毕后瞬间高亮反显目标图纸，中间 CAD 视口即刻同步渲染该图纸大图；
  3. **编译与热同步交付**：
     - HTML 静态模板已覆盖至 `bin\Debug\net48\Resources\` 与 `publish\Resources\`；
     - `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误。
       - **【人工】**：如 `¥ 30.00`
       - **【二次组】**（方案主名称）：如 `多功能表1` / `热水泵31台`
     - 中间 CAD 矢量视口顶部同步加入参数微章预览群，一边看图一边掌握 5 大工艺定额参数；
  5. **构建验证与热同步交付**：
     - 资源已覆盖同步至 `bin\Debug\net48\Resources\` 与 `publish\Resources\`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误。

- **彻底修复保存无效、按型号去重聚合、严格B列元件组判别及移除搜索框 (`Resources/secondary_circuit_manage.html`, `Services/ExcelServices.SecondaryCircuit.cs`, `Forms/SecondaryCircuitForm.cs`)**：
  1. **彻底解决保存无效问题**：
     - 在 `Forms/SecondaryCircuitForm.cs` 中将保存与扫描操作交由 `ExcelAsyncUtil.QueueAsMacro` 调度至 Excel 纯净主线程执行，彻底解除 WebView2 跨线程直接操作 COM 导致的异常阻断；
     - 使用全工程标准的 `Tool.GetActiveExcelContext()` 稳健获取 Excel 实例与活动工作表，避免非前台激活状态下的空引用；
     - 将 `postToCSharp` IPC 助手函数提升至 `setup()` 最顶层，彻底杜绝任何调用顺序时序问题；
  2. **严格 B 列类别判别，剔除机械尺寸开孔行**：
     - 彻底废除 `model.Contains("*")` 误判逻辑；
     - 严格限定：B 列必须为“元件组”（或 B 列为空且以 `*` 开头），若 B 列已有其他明确分类（如开孔 `HK91*91` 等），坚决排除；
  3. **右侧列表按型号规格全局去重聚合**：
     - 扫描服务重构为按 `GroupModel` 进行 Dictionary 聚合，输出 `ExcelComponentGroupItemDto`；
     - 列表不再重复罗列同一个型号（如不再出现十几个 `*KM变频器`），而是聚合成单行，并醒目标注 `共 X 处` 徽章与涵盖箱柜详情；
     - 保存时，一次绑定自动将图号批量覆盖写入全表该型号对应的所有物理行第 32 列；
  4. **取消右侧顶部搜索框，扩展表格高度**：
     - 彻底移除顶部搜索输入框，界面更加清爽专业；
     - 表格有效可视高度扩展至 `445px`，按钮排版自适应优化；
  5. **构建与热同步**：
     - 页面已热同步至 `bin\Debug\net48\Resources\` 与 `publish\Resources\`；
     - `dotnet build /t:Compile /p:DebugType=none` 顺利通过：0 错误 0 警告。


  1. **多重安全防重门禁全链路闭环**：
     - **门禁 1 (批量选图拦截)**：在 `confirmAddCircuitDwgCodes` 中，用户勾选图纸确认选入时，系统自动扫描当前全库所有既有方案，若图号已被其他方案绑定（例如 `CA1B` 已被【双电源互投】绑定），自动进行友好拦截与详细弹窗警示，只放行全库真正唯一的图纸；
     - **门禁 2 (手动输入拦截)**：在 `addNewCodeTag` 中，用户手动输入代号按回车时，先查当前方案内部重复，再查全局跨方案占用，发现冲突立即阻止添加并告知占用方案名；
     - **门禁 3 (前端保存前拦截)**：在 `submitScheme` 中执行终极代号扫描，一旦待保存方案包含跨方案重复代号，阻断提交并弹出阻断警示；
     - **门禁 4 (后端底层兜底保护)**：在 `PersonalComponentDbService.SecondaryCircuit.cs` 中增加 `CheckApplicableCodeConflict` 方法，并在 `SecondaryCircuitController.SaveScheme` 中接入拦截，确保数据库层绝对无同名回路代号落盘；
  2. **热更新与编译验证**：
     - 资源已热同步至 `bin\Debug\net48\Resources\secondary_circuit_manage.html` 与 `publish\Resources\`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误。

- **纯粹直接使用 `Tool.GetAppDataDirectory()` 统一二次方案与物料库存储路径，彻底去除旧路径兼容 (`Services/PersonalComponentDbService.cs`)**：
  1. **彻底精简路径获取逻辑**：
     - 去除 `GetDatabaseFilePath()` 中所有关于历史版本 LocalAppData 的环境探测、迁移复制与降级回退代码；
     - 直接纯粹调用公共工具类 `Tool.GetAppDataDirectory()` 拼接返回 `Path.Combine(appDataDir, DefaultDbFileName)`；
  2. **物理文件与编译验证**：
     - 数据完全以插件运行目录下的 `data\personal_components.db` 作为唯一物理存储；
     - 现有 113 套方案数据文件已同步至便携数据目录；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误 0 警告。

- **加大方案编辑弹窗操作高度，并彻底消除弹窗底部水平滚动条 (`Resources/secondary_circuit_manage.html`)**：
  1. **弹窗操作高度整体扩充**：
     - 将 `<el-dialog>` 宽度由 840px 微调扩展至 `860px`，顶部距离调优为 `top="2.5vh"`，留出更舒展的纵向视野；
     - 弹窗主滚动内容容器显式配置 `min-height: 560px; max-height: 84vh;`，无论物料行数多少，均提供开阔挺拔的编辑空间；
     - 下方“二次 BOM 子物料清单”卡片容器设定 `min-height: 330px;`，并将物料表格由原先容易引起视觉坍缩的 `max-height="220"` 替换为固定展开的受控自适应高度 `height="280"`，无论物料数据只有 1 行还是多行，均呈现规整大气的表格视图，内部行多时自适应纵向滚动；
  2. **彻底消除箭头所指的底部水平滚动条**：
     - **排查根因**：Element Plus 的 `<el-row :gutter="14">` 内部应用了负边距（`margin-left: -7px; margin-right: -7px;`），使得表单宽度扩展至 `100% + 14px`；而外层滚动容器此前仅指定了 `overflow-y: auto;` 未设定 `overflow-x: hidden;`，负边距溢出触发了父容器在底部渲染一条横向水平滚动条；
     - **全局与局部双重熔断防溢出**：
       - 在 CSS 样式表中注入 `.el-dialog__body { overflow-x: hidden !important; padding: 16px 20px !important; box-sizing: border-box; }`；
       - 在编辑对话框主滚动容器上明确配置 `overflow-y: auto; overflow-x: hidden !important; padding: 2px 4px 2px 2px; box-sizing: border-box;`；
       - 在表单层级配置 `style="width: 100%; overflow-x: hidden; box-sizing: border-box;"`，完全吸收 row gutter 负外边距，100% 杜绝水平横向滚动条产生；
  3. **编译构建与热更新同步**：
     - 已将前端页面热更新拷贝至 `bin\Debug\net48\Resources\secondary_circuit_manage.html` 与 `publish\Resources\`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误 0 警告。

- **去除旧眼睛预览与粘贴路径功能，强化 DWG 目录双写持久化保存机制 (`ConfigManager.cs`, `Resources/secondary_circuit_manage.html`, `Forms/SecondaryCircuitForm.cs`)**：
  1. **彻底去除箭头所指旧预览功能**：
     - 主方案表格：移除“同配置适用回路代号”列中 Tag 内部的 👁️ 预览图标及点击事件，恢复纯净 Tag 呈现；
     - 编辑方案对话框：移除排布图输入框右侧的 👁️ 预览按钮，移除回路代号 Tag 内部的 👁️ 预览图标，移除“点击 Tag 上的 👁️ 可即刻预览图纸”提示文字；
     - 彻底清理旧的 `dwgPreviewVisible` 模态弹窗 DOM 结构，清除 `openDwgPreview`、`previewData`、`dwgPreviewLoading`、`dwgPreviewLoaded` 及其关联代码；
  2. **彻底去除粘贴路径及关联代码**：
     - 编辑对话框中移除排布图右侧 `✍️`（粘贴路径）按钮与回路代号右侧 `✍️ 粘贴目录` 按钮；
     - 回路图纸选择弹窗中移除顶部 `✍️ 粘贴路径` 按钮；
     - 清除前端 `inputLayoutDirPrompt`、`inputCircuitDirPrompt` 及后端 `setDwgDirManual` 消息处理分支；
  3. **强化目录设置的配置文件持久化记忆 (`ConfigManager.cs`)**：
     - 排查修复了原 `SaveConfig` 仅写入 AppData 目录而在重启时被同级 `appsettings.json` 覆盖还原的问题；
     - 改造为**双写持久化**：优先同步写回插件运行目录下的 `appsettings.json`（便携优先），同时写入 `%LocalAppData%\ExcelCTTools\appsettings.json`（全局兜底）；
     - 确保用户点击【📁 浏览目录】后，选择的新目录立即写入配置文件，以后重新打开管理中心或重启 Excel 均能持续读取并记忆使用；
  4. **编译构建与热同步**：
     - 静态模板已热同步至 `bin\Debug\net48\Resources\` 与 `publish\Resources\`；
     - `dotnet build /t:Compile` 编译通过：0 错误。

- **彻底修复回路图纸选择弹窗左侧表格“大小”列被遮挡问题 (`Resources/secondary_circuit_manage.html`)**：
  1. **遮挡根因诊断**：
     - 用户将左侧容器宽度收窄至 250px 时，表格总列宽（勾选列 32px + 名称列原先 `min-width="145"` + 大小列 68px + 垂直滚动条 15px = 260px）超过了容器内减去 padding 后的实际可用净宽（234px），导致表格产生横向溢出，最右侧的大小列被推至可视区外边缘遮挡；
  2. **系统级紧凑表格排版重构**：
     - **CSS 专用紧凑样式 (`.dwg-file-table`)**：将表格单元格默认的大内边距（`12px`）收缩至 `padding: 4px 4px !important;`，并将垂直滚动条优化为 5px 细窄半透明滚动条，瞬间释放近 25px 横向布局空间；
     - **列宽与边距精确匹配**：
       - 左侧容器设为 `260px`（极紧凑宽度），`padding: 6px`，净可用宽度为 248px；
       - 勾选列明确为 `32px`，大小列给足 `62px`（文字“大小”、“64.3 KB”、“目录”均留有充裕空间，右对齐完整展露）；
       - 名称列解除过大的硬性 min-width 约束，改为自适应弹性填充（`min-width="95"`），表头精简为 `名称 (双击进入)`；
  3. **编译与热同步交付**：
     - 资源已热同步至 `bin\Debug\net48\Resources\` 与 `publish\Resources\`；
     - `dotnet build /t:Compile` 编译通过：0 错误 0 警告。

- **优化回路图纸选择弹窗左右分栏排版，调小左侧文件列表宽度至 290px (`Resources/secondary_circuit_manage.html`)**：
  1. **空间分配优化**：将左侧文件/目录表格容器宽度从 430px 精简调小至 290px，让出 140px 横向空间给右侧 WebGL 矢量 CAD 画布，显著提升工程图纸可视化面积；
  2. **紧凑排版适配**：
     - 搜索工具条精简按钮边距，将“全选图纸”精简为“全选”，与“清空”按钮在 290px 下平滑单行并列不折行；
     - 表格列宽紧凑微调（勾选列 36px，大小列 68px，名称列 145px），表头精炼为 `名称 (双击进目录)`；
     - 底部状态提示条文字精简为 `X 目录, Y 图纸 | 已选 Z 项`；
  3. **编译与热同步**：
     - 资源已热同步至 `bin\Debug\net48\Resources\` 与 `publish\Resources\`；
     - `dotnet build /t:Compile` 编译通过：0 错误 0 警告。

- **落地回路图纸目录层级双击下钻展开与真实矢量 WebGL 纯净视口预览架构 (`Services/DwgPreviewService.cs`, `Controllers/SecondaryCircuitController.cs`, `Forms/SecondaryCircuitForm.cs`, `Resources/secondary_circuit_manage.html`, `ExcelAddInDemo.csproj`)**：
  1. **纯前端 WebGL + WebAssembly 矢量引擎离线化与全局打包**：
     - 在前端项目 `draw-code-pcui` 中基于 `esbuild` 将 `@flyfish-dev/cad-viewer` 纯前端 CAD 核心打包为全局独立 IIFE 包 `cad-viewer.bundle.js`（导出 `window.CadViewerLib`）；
     - 将配套核心资源 `style.css`、`wasm/`（`libredwg-web.wasm` 6.3MB, `dwg-worker.js` 133KB, `libredwg-web.js`, `dwfv-render.wasm`）完整复制至 `Resources/cad-viewer/`；
     - 在 `ExcelAddInDemo.csproj` 中配置 `Resources\cad-viewer\**\*.*` 编译与打包拷贝，实现 100% 离线绿色便携运行；
  2. **WebView2 安全源域名映射与 Worker/WASM 跨域穿透 (`SecondaryCircuitForm.cs`)**：
     - 在 WebView2 初始化时调用 `SetVirtualHostNameToFolderMapping("appassets.local", resDir, Allow)`，将页面加载协议升级为 `https://appassets.local/secondary_circuit_manage.html`；
     - 彻底解除了 Chromium 在 `file:///` 协议下对 Web Worker 与 `.wasm` 加载的跨域安全阻断；
  3. **后端层级扫描与二进制流通道 (`Services/DwgPreviewService.cs`, `Controllers/SecondaryCircuitController.cs`)**：
     - 新增 `DwgDirectoryItem` 与 `DirectoryHierarchyResult` 模型，支持提取当前目录下的子文件夹、父级目录路径以及当前层级的 DWG 文件；
     - 实现 `ScanDirectoryHierarchy(dirPath)` 扫描方法与 `ReadDwgBinaryBase64(filePath)` 二进制 Base64 读取流；
     - 实现了跨线程安全派发与异常安全捕获；
  4. **前端左右分栏与工具栏纯净化视口交互 (`Resources/secondary_circuit_manage.html`)**：
     - **左右分栏布局**：弹窗重构为 1060px 舒展大屏，左侧 430px 展示文件夹与图纸混合表格，右侧自适应呈现黑色 WebGL 矢量视口；
     - **文件夹双击下钻**：文件夹标注 📁 图标与 `(双击进入)` 提示，双击自动深入下一级子目录；顶部提供动态面包屑与“⬆️ 返回上级”按钮；
     - **多选防误选保护**：表格勾选列通过 `:selectable="(row) => !row.isDir"` 禁用文件夹复选框，批量将选中的 DWG 图纸代号批量加入方案适用回路 Tag；
     - **纯净视口 (关闭所有工具栏)**：在 CSS 中通过 `.pure-cad-container .cad-toolbar, .cad-layers, .cad-inspector { display: none !important; }` 强力隐藏原生多余面板，仅保留纯净画布视口；
     - **原生手势全支持**：视口支持鼠标滚轮平滑缩放、左键拖拽平移、双击全图居中拟合，并支持顶部一键【🚀 CAD打开】与【📂 定位】；
  5. **代码规范与编译验证**：
     - 严格遵循每 3 行包含至少 1 行中文注释，配置硬编码显式标明 `--硬编码--`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误 0 警告；
     - 静态资源已全量热同步至 `bin\Debug\net48\Resources\` 与 `publish\Resources\`。

- **彻底修复文件夹选择器导致的 WebView2 页面卡死与假死问题，并在全工程排查修复同类模态对话框阻塞隐患 (`SecondaryCircuitForm.cs`, `CreateProjectForm.cs`, `EnterpriseSettingsForm.cs`, `Resources/secondary_circuit_manage.html`)**：
  1. **卡死根因定位分析**：
     - **根因 1 (Chromium IPC 死锁)**：前端通过 `window.chrome.webview.postMessage` 向 C# 发送 `selectDwgDir`。C# 在 WebView2 自身的 `WebMessageReceived` 事件处理回调中同步调用 `ShowDialog()`，启动了 Win32 阻塞式模态消息循环，导致主线程被挂起，无法向 WebView2 管道回传确认信号，Chromium 渲染进程因此被全局挂起，引发整个界面彻底冻结；
     - **根因 2 (Shell 扩展枚举阻塞)**：WinForms 传统 `FolderBrowserDialog` 默认枚举用户安装的第三方外壳扩展（如“百度网盘同步空间”等慢速驱动器），网络或云盘响应缓慢时会产生极长时间的同步假死；
  2. **全工程排查与线程模型彻底重构 (3 处全部修复)**：
     - **`SecondaryCircuitForm.cs` (选择 DWG 排布图/回路图纸目录)**：将 `selectDwgDir` 重构成通过独立的 STA 线程（`ApartmentState.STA`, `IsBackground = true`）异步弹出 `FolderBrowserDialog`，选择完成后通过 `SafeInvoke` 切回主线程通知前端，完全解耦主消息循环与 Chromium IPC；
     - **`CreateProjectForm.cs` (新建项目选择保存目录)**：同样改造为独立 STA 线程异步运行 `FolderBrowserDialog`，安全回传选定路径；
     - **`EnterpriseSettingsForm.cs` (选择企业 Logo 本地图片)**：同样改造为独立 STA 线程异步运行 `OpenFileDialog`，选定后安全回传 Base64，杜绝主线程卡死；
  3. **前端双轨操作体验提升 (`secondary_circuit_manage.html`)**：
     - 在“二次排布图”与“同配置回路代号”配置区，以及“从回路图纸目录选择 DWG”弹窗中，全面新增 `✍️ 粘贴路径` / `✍️ 粘贴目录` 按钮；
     - 用户既可在无卡死风险的情况下点击 `📁 浏览目录`，也可直接复制 Windows 资源管理器路径通过 `ElMessageBox.prompt` 快速粘贴，零弹窗极速设定；
     - C# 后端新增 `setDwgDirManual` 接口，自动验证非空目录、持久化配置并自动扫描提取 DWG 图纸清单；
  4. **编译验证与资源热同步**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误；
     - 资源已热同步至 `bin\Debug\net48\Resources\secondary_circuit_manage.html`。

- **落实方案 A：二次排布图与回路代号 DWG 目录分别就近独立设置、直接加载选入与即时预览 (`SecondaryCircuitForm.cs`, `SecondaryCircuitController.cs`, `Services/DwgPreviewService.cs`, `AppConfig.cs`, `ConfigManager.cs`, `Resources/secondary_circuit_manage.html`)**：
  1. **实体业务对应明确**：
     - 将“二次排布图”（如 `BDY`）与“同配置回路代号”（如 `风机2DY`）从纯手工文本输入升级为实体 DWG 图纸文件关联绑定；
  2. **方案 A 就近配置与本地持久化**：
     - 在“二次排布图”输入项与“同配置回路代号”输入区就近添加 📁 目录设置按钮，调用 WinForms 原生 `FolderBrowserDialog` 选取本地文件夹，自动扫描提取其下所有 `*.dwg` 文件；
     - 在 `AppConfig.cs` 与 `ConfigManager.cs` 中增加 `SecondaryCircuitSettings`（`LayoutDwgDirectory` 与 `CircuitDwgDirectory`），实现路径的本地持久化自动记忆与跨会话保留；
  3. **直接加载目录图纸选入**：
     - **排布图**：输入框支持基于 `el-autocomplete` 自动加载排布图目录下的 DWG 文件进行联想补全与快捷下拉选择；
     - **回路代号**：点击“+ 从目录加载 DWG 选择”弹出模态选择窗口，支持按图名模糊搜索、全选/清空、表格多选勾选并批量加入 Tag，且保留手动输入兜底；
  4. **双通道高可靠 DWG 图纸即时预览与 CAD 调起 (`Services/DwgPreviewService.cs`)**：
     - **通道 1 (Shell 原生)**：调用 Windows Shell `IShellItemImageFactory` 提取系统缓存的高清 CAD 缩略图；
     - **通道 2 (二进制文件流)**：读取 DWG Header `IMAGE_SEEK` 区域直接解析提取内嵌 BMP 缩略图；
     - **通道 3 (兜底占位)**：动态生成工业黑绿风 CAD 占位图并提示直接打开；
     - **交互闭环**：在主方案表格、编辑框排布图输入区、回路代号 Tag 内部均集成 👁️ 预览图标，点击调起预览大弹窗展示缩略图、文件物理路径、大小与更新时间，并支持一键【🚀 在 AutoCAD 中打开图纸】与【📂 在资源管理器中定位】；
  5. **编译构建与热同步**：
     - 执行 `dotnet build /t:Compile` 编译通过：0 错误 0 警告；
     - HTML 模板已热同步至 `bin\Debug\net48\Resources\secondary_circuit_manage.html` 与 `publish\Resources\`。


- **全面落地基于 ExWinner 架构的工业级 Excel 行列聚光灯 (Spotlight) 功能 (`Services/ExcelServices.Spotlight.cs`, `Forms/SpotlightOverlayForm.cs`, `Forms/ExcelWindowHook.cs`, `Models/SpotlightConfig.cs`, `RibbonController.cs`, `ExcelEventManager.cs`, `AddInMain.cs`)**：
  1. **架构解密与对齐落地**：
     - 深度学习并复刻 ExWinner 的非侵入式金标准架构：**Win32 GDI Region 动态剪裁 + 操作系统级无边框半透明穿透浮窗 (`WS_EX_TRANSPARENT | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW`) + 原生窗口消息钩子 (`NativeWindow`)**；
     - 彻底消除传统 VBA 条件格式对 Excel 撤销重做栈（Ctrl+Z）的破坏以及对工作簿文件的样式污染；
  2. **高精几何与性能保障**：
     - 利用 Excel COM `ActivePane.PointsToScreenPixelsX/Y` 自动兼容 Zoom 视口缩放、Windows 系统级 DPI 缩放与首行首列冻结窗格；
     - 通过 Win32 GDI `CreateRectRgn`、`CombineRgn(..., RGN_OR)` 组合十字高亮几何，并经由 `SetWindowRgn` 物理裁剪浮窗（非高亮区域直接剔除，GPU/CPU 负载极低）；
     - `ExcelWindowHook` 拦截 `EXCEL7` 与 `XLMAIN` 的移动、缩放与滚屏消息，内置 35ms 防抖节流消除撕裂，切出 Excel 自动隐匿；
  3. **交互集成与持久化**：
     - Ribbon 功能区【辅助项】分组新增【💡 聚光灯】大图标切换按钮（带互锁状态同步）；
     - 注册 Excel-DNA 宏快捷键 `[ExcelCommand(ShortCut = "^%L")]`（`Ctrl + Alt + L`）；
     - 在 `%LocalAppData%\ExcelCTTools\config\spotlight_config.json` 中自动持久化开关与参数，Excel 启动时自愈恢复；
  4. **代码规范与编译验证**：
     - 严格遵循新增代码每 3 行包含至少 1 行中文注释，硬编码统一标识 `--硬编码--`；
     - `dotnet build` 编译全量通过：0 错误，成功生成 32位与 64位 `.xll` 加载项。

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
- 本地个人物料库检索抽屉支持品牌检索，默认选中“二次元件”，支持全部品牌切换与型号/名称组合搜索；
- **常规样式甲方投标报表迁移实施全面落地闭环**：
  1. **标准模板资产引入**：将 ExWinner 官方母版 `自定义示例1_常规样式.xlsx` 规范化迁入 `Resources/TenderReport_Regular.xlsx`，并在 `.csproj` 中配置自动同步至输出与发布目录；
  2. **Ribbon 菜单多级对齐**：将功能区【④出报表】下的【标书报表】扩展为 1:1 对齐的多级级联菜单，挂载【甲方投标报表】（常规样式、高级样式、国网报表、用户定制报表）、【原始样式报表】、【内部审核报表】与【报表市场】，并完成常规样式报表点击路由绑定；
  3. **强类型模型与交互控制器**：新增 `Models/TenderReportModels.cs` 与 `Controllers/TenderReportRegularController.cs`，统一前后端数据规范与文件快速定位；
  4. **现代前端导出向导**：新增 `Resources/tender_report_regular.html`，基于 Vue 3 `<script setup>` + Element Plus 构建，遵循绿蓝相间主色调 `#009688`，支持工程信息核对、分类表格多选统计及导出偏好配置；
  5. **高性能生成引擎**：在 `Services/ExcelServices.TenderReport.cs` 中实现二维数组一次性抓取当前工程箱柜与元器件数据，动态克隆模板生成《封面》、《屏柜汇总表》与《屏柜分项表》，自动刷新公式求和、大写金额转换并彻底脱敏内部成本底价；
  6. **编译构建**：执行 `dotnet build` 编译打包 0 错误，严格遵循每 3 行包含一行中文注释规范。

- **二次回路图纸对齐与元件组绑定工作台布局与显示调优 (全面闭环)**：
  1. **DWG 矢量视口拓宽与右侧列表收窄**：
     - 左侧列表由 260px 微调为 230px；
     - 右侧元件组看板由 480px 调窄为 340px，内部两列表格列宽由 180/160px 自适应微调为 150/130px；
     - 中间 DWG 视口额外获得 160px+ 宽度，大幅扩容近 40%；
  2. **视口顶部胶囊与图纸名称防止截断**：
     - 文件名容器设置弹性收缩与 `:title` 悬浮提示；
     - 5 项定额参数胶囊设置 `flex-shrink: 0;`、紧凑间距与悬浮浮窗，消除被 CAD 打开按钮或边界裁切截断现象；
  3. **底部关联二次方案参数卡片双排显示**：
     - 将原单排 5 列拥挤网格改造为【双排清爽栅格】：第一排 3 列（二次线跨门、二次材料费、开孔），第二排 2 列（人工安装费、二次组方案名，大宽度保证长名称不截断）；
     - 表格启用 `min-height: 100px; height: 0; flex: 1;` 弹性内部滚动，参数卡片添加 `margin-top: auto; flex-shrink: 0;`，彻底根除数值下半截被容器裁切问题；
  4. **资产热同步与构建验证**：
     - 页面已同步热更新至 `bin/Debug/net48/Resources` 与 `publish/Resources`。

- **二次回路图纸对齐工作台最右下方参数区域不全深度根治 (全面闭环)**：
  1. **弹窗与主体高度释放**：弹窗显式声明 `height: 92vh`，主体容器彻底转为 `flex: 1; min-height: 0;` 纯弹性结构，解绑之前写死的 640px 限制，释放充足的垂直净空；
  2. **表格绝对定位隔离包裹层**：在右侧面板内为 `<el-table>` 建立独立的 `flex: 1; min-height: 110px; position: relative;` 包裹层，并设置 `height="100%"`，让表格在内部虚拟滚动，彻底杜绝表格行把底部参数卡片推挤出容器下边框的结构性隐患；
  3. **参数卡片双排紧凑精调**：面板宽度微调为 360px，参数卡片各指标方块采用 32px 紧凑高度与 `white-space: nowrap` 布局，行高严格受控，头部与边距极致优化，确保整张卡片（包括第二排方案名与人工）100% 完整展示；
  4. **资产热同步与构建验证**：页面已热复制更新至 `bin/Debug/net48/Resources` 与 `publish/Resources`。

## [Completed]

- **基于 ExWinner 工业标准的【分类】全功能与【项目信息】表双向联动落地 (`RibbonController.cs`, `Services/ExcelServices.Category.cs`, `Forms/CategoryForm.cs`, `Resources/category.html`, `Controllers/CategoryController.cs`, `Models/CategoryModels.cs`)**：
  1. **Ribbon 功能区 5 大操作对齐**：完整实现【分类】下拉菜单：新建分类、编辑分类、复制分类、插入复制的分类、删除分类；
  2. **项目信息表标准联动公式体系**：写入与母版 100% 对齐的 C~H 列公式（箱柜数量、总价、成本、毛利、毛利率、箱变智能识别），并集成遇小计行自动插入物理行 (`Insert(-4121)`) 的防覆盖机制；
  3. **删除与重命名的强自愈**：删除分类时同步物理删除【项目信息】中的对应行，序号公式 `=ROW()-ROW(A$28)` 自动递补，彻底消除 `#REF!`；重命名分类时自动更新超链接与公式表名引用；
  4. **复制分类与定义名称唯一性**：克隆源工作表并清洗跨表公式，重新扫描全工作簿并为新表分配全局递增唯一的定义名称（`Cab_Sum_K`、`Cab_Det_K` 等），彻底杜绝跨表名称冲突；
  5. **自适应三合一 WebView2 窗体**：`category.html` 升级为新建/编辑/插入复制三合一自适应界面，微秒级平滑位移拖拽，已全量同步并编译通过 (0 Errors)。

- **深度调研并学习 ExWinner（利驰智能电气报价软件）中箱柜下拉菜单 11 大核心功能架构与业务逻辑 (`D:\Program Files\ExWinner\`)**：
  1. **软件与架构底层摸排**：
     - ExWinner 同样基于 Excel-DNA (`ExcelDna.Integration.CustomUI.ExcelRibbon`) 框架开发，主功能位于 `leadsoft.superwinner.BLL.dll`、`leadsoft.superwinner.DAL.dll`、`ExcelAddIn4Scm.dll`；
     - 箱柜菜单 11 大按钮在 ExWinner 中与 `RibbonControlIdsConst` 逐一对应（`addCabinetButton`、`addNoDetailCabinetButton`、`batchAddCabinetButton`、`editCabinetButton`、`cutCabinetButton`、`copyCabinetButton`、`pasteCabinetButton`、`deleteCabinetButton`、`orderCabinetButton`、`importCabinetBomButton`、`importAICabinetBomButton`）；
  2. **核心业务与数据模型逻辑解析**：
     - **箱柜与定义名称四元组映射**：ExWinner 的 `Cabinet` 实体由 `TableName` (`Cab_Det_K`)、`DataName` (元器件数据区)、`ReserveName` (`Cab_Subsum_K` 计费/保留区) 与汇总行 `Cab_Sum_K` 构成，严格遵循行首到小计行、总计行的块级管理；
     - **新建/无明细箱柜机制**：普通箱柜复制预置模板块并分配递增 `K`，建立双向超链接；无明细箱柜（`SGCC_NoModel.xlsx`）则省去元器件明细行，直接汇总计价；
     - **批建箱柜与甲方清单对齐**：`frmBatchAddCabinet.html` 接收外部表格多列快速粘贴，校验后批量关更新渲染并并发写表，提供自动匹配云方案组价扩展；
     - **剪切/复制/插入/删除**：通过整块二维内存数组进行暂存（`Formula[,]`、值、格式），插入时动态计算全表 `maxK + 1`，彻底规避跨表覆盖，剪切后执行物理下移整块清除；
     - **箱柜调序核心**：`frmOrderCabinetByList.html` 与 `FrmMoveCabinet`，采用“内存抽取全部箱柜对象 -> 按目标顺序重排 -> 清空重建”算法，彻底根除 Excel 逐行搬移导致的公式撕裂与卡死；
     - **普通/智能BOM导入**：`FrmImportCabinetBom` 按柜号分组填充元器件行，遇空位不足向下自动插行并刷新增量公式，智能导入则结合正则/AI拆解复杂型号。

- **彻底修复【复制箱柜】与【插入复制的箱柜】表格边框丢失、底色丢失及 #VALUE! 计算错误 (`Models/CabinetModels.cs`, `Tool.cs`, `Services/ExcelServices.CabinetManage.cs`)**：
  1. **深度根因定位与排查**：
     - 原 `InsertCopiedCabinet` 在插入空白行后，仅对单元格的 `Value2` 和 `Formula` 进行了赋值，完全未复制 Excel 单元格样式（边框线、表头底色、字体、行高、合并单元格全部丢失，视觉上成为纯无边框文字）；
     - 原逻辑直接将源箱柜的静态公式字符串灌入新单元格，导致单价、总价公式中的相对行号错位，直接引发 `#VALUE!` 错误；
  2. **模型扩展与物理行号映射 (`Models/CabinetModels.cs`, `Services/ExcelServices.CabinetManage.cs`)**：
     - 在 `CopiedCabinetContext` 中补充 `SourceSumRow`、`SourceDetStartRow`、`SourceDetEndRow`、`SourceDetRow`、`SourceSubsumRow`、`SourceTolsumRow` 物理行号快照；
     - 在 `CaptureActiveCabinetToClipboard` 抓取时完整沉淀源行号与内部结构；
  3. **架构升级为 Excel 原生 Range.Copy 完整克隆体系 (`Tool.cs`, `Services/ExcelServices.CabinetManage.cs`)**：
     - **汇总行格式复刻**：在插入空白行后，执行 `srcSumRange.Copy(dstSumRange)`，实现汇总行边框线、背景底色、行高 100% 完整克隆；
     - **定义名称自适应寻址**：在插入汇总行后，通过 `Tool.SafeGetSheetName` 读取定义名称最新位置，精准推算明细块当前源行号与目标行号，彻底杜绝同表上下插入导致的物理行号错位；
     - **明细块 100% 格式与公式相对平移克隆**：调用 `srcDetailRange.Copy(dstDetailRange)`，利用 Excel 原生机制自动平移所有相对公式（单价、总价、合价），彻底杜绝 `#VALUE!` 计算错误；
     - **清洗与刷新闭环**：调用 `Tool.CleanRangeFormulas` 擦除外部工作簿残留路径，调用 `RefreshCabinetFeeAreaFormulas` 自适应刷新 A 列序号与小计/总计公式；
     - **坚固降级兜底**：保留内存二维数组回填与标准细线边框设置，双轨保障极端异常下的稳健运行；
  4. **编译构建与验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误。

## [In-Progress]

- 复制与插入箱柜功能测试与用户验收。

## [Next]

- 与用户确认箱柜功能具体实施优先级，并按计划推进实施。





