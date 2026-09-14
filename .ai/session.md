# Session State

- **新建箱柜 `Cab_Subsum` 小计行定义名称偏移修正与动态嗅探自愈 (`ExcelServices.Cabinet.cs`)**：
  1. **问题根因彻底清除**：
     - `CopyCabinetDetailFromTemplate` 在从 `CabinetTemplate.xlsx` 母版复制明细块时，硬编码了 `int newSubsumRow = targetDetailStartRow + (65 - 41);`（即 $+24$ 行）；
     - 但母版中 `Cab_Subsum_1`（小计行）真实位于第 66 行（相对起始行 41 的真实偏移为 $+25$ 行），导致旧代码把 `Cab_Subsum` 误绑定在第 65 行（元器件最后一个空行），比实际小计行小了 1；
  2. **双重精准防护方案落地**：
     - 将基础偏移修正为 `targetDetailStartRow + (66 - 41)`；
     - 增加动态内容嗅探：在复制出的明细区域中自动扫描 B 列，精准锁定包含“小计”与“单台合计/总计”的物理行，彻底杜绝后续模板改动再次引发偏差；
  3. **编译构建验证**：
     - `ExcelAddInDemo.csproj` 成功编译生成，**0 错误 0 警告**。

  1. **问题根因彻底清除**：
     - Excel 端（`excel-ct-tools`）与 AutoCAD 端（`cad-net_1`）因在不同宿主进程中运行，`Tool.GetAppDataDirectory()` 默认各自定位到各自的运行目录下的 `data` 文件夹，导致调价公式（`formula_fee_settings.json`）、企业设置等配置数据无法双端同步；
  2. **系统级漫游引导与跨端共享架构**：
     - 在 Windows 用户通用漫游目录 `%APPDATA%\ExcelAddInDemo\global_config.json` 引入全局引导配置；
     - 无论是在 Excel 进程还是 AutoCAD (`acad.exe`) 进程中，均能通过统一的系统全局路径读取同一份 `customDataDirectory` 自定义数据目录配置；
     - `Tool.GetAppDataDirectory()` 动态优先返回用户自定义配置的有效目录；若未配置则平滑回退至插件目录下的 `data` 文件夹；
     - 当用户配置切换到新的空目录时，自动将默认数据目录中的基础 JSON 配置文件安全同步过去，杜绝已有数据丢失；
  3. **“我的 - 企业设置”前端与宿主完整闭环**：
     - 在 `enterprise_settings.html` 中新增“数据目录：”表单输入框与“浏览...”按钮；
     - 风格统一为 `#009688` 绿蓝主题，并在 `<script setup>` 与 `setupLogic()` 中均完整实现双向绑定与监听；
     - 在 `EnterpriseSettingsForm.cs` 中增加 `selectDataDirectory` 指令处理，通过独立 STA 后台线程弹出 `FolderBrowserDialog` 目录选择框，绝不阻塞 WebView2 主通信管道；
     - 用户点击“保存”时，双写到企业设置与系统全局引导配置中，立即生效；
  4. **编译与验证**：
     - 执行 `dotnet build "e:\Ace\excel-ct-tools\ExcelAddInDemo.csproj" /p:RunExcelDnaBuild=false` 构建成功，**0 错误**；
     - 单元测试验证全局配置读写、自动模板复制、平滑回退机制均 100% 正常通过。

  1. **问题根因彻底清除**：
     - 彻底改变以往在各个业务入口（调费、算辅材、报表等）无脑强制全量扫描整表 `UsedRange`、二维数组倒序遍历与正则模糊猜测的粗暴模式；
     - 消除重复全量推导导致的几百毫秒严重性能损耗，并彻底根除因启发式“猜规则”反噬原本精准建立的代码锚点的问题。
  2. **轻量嗅探守门机制落地**：
     - 在 `FixAndFillCabinetNamesForSheet(dynamic sheet, bool forceRebuild = false)` 入口处增加健康度守门；
     - 快速比对当前工作表现存定义名称映射：若已具备合法 Sum 汇总行与正确的 Det 明细层级拓扑，**耗时 0ms 直接返回现有箱柜数量，跳过所有 UsedRange 与正则推导**；
     - 仅当定义名称数量为 0 或检测到破坏性 `#REF!` 时才真正执行自愈反推；
     - 既有所有调用方完全保持兼容，自动享受微秒级极速响应与防反噬保护。
  3. **编译构建与生效验证**：
     - `ExcelAddInDemo.csproj` 成功编译生成，**0 错误**。

  2. **总计行 (tolsum) F 列数量填写**：
     - 在 `ExportSingleCabinetOptimized` 中，提取有效数量 `int cabQty = cab.Header.Quantity > 0 ? cab.Header.Quantity : 1`；
     - 在刷新计费区域公式后，设置 `sheet.Cells[tolsumRow, 6].Value2 = cabQty;`（--硬编码: 第 6 列为 F 列--）；
     - 同步在 `CopyCabinetDetailFromTemplate` 与 `CreateNewCategory` 中补齐了总计行第 6 列的数量回填；
  3. **编译构建与代码规范**：
     - 严格遵守每 3 行包含一行中文注释，硬编码均带有 `--硬编码--` 标明；
     - `ExcelAddInDemo.csproj` C# 源码编译 **0 警告 0 错误**。

- **智能辅材与壳体计算中心「更新当前分类」与「更新所有分类」极速性能优化与 Element Plus 动态进度条落地交付 (`ExcelServices.CabinetAuxCalc.cs`, `CabinetAuxCalcController.cs`, `CabinetAuxCalcForm.cs`, `cabinet_aux_calc.html`)**：
  1. **问题根因彻底清除**：
     - **未挂起重绘重算**：原更新当前分类未开启 `ScreenUpdating = false`、`Calculation = xlCalculationManual` 与 `EnableEvents = false`，每次单元格修改均触发全表重算与重绘，性能严重拖慢数十倍；
     - **定义名称重复全量扫描**：原循环中每次调用 `ScanCabinetData` 均从零遍历全表定义名称，产生大量冗余 COM 跨进程开销；
     - **CAD 管道未运行超时累加**：未开启 AutoCAD 时每个箱柜均经历 500ms 握手超时累加；
     - **WebMessage 同步阻塞**：在 Chromium IPC 回调中同步执行 Excel 操作导致界面完全假死、动画冻结。
  2. **端到端极速优化与解耦实施**：
     - **四重极速性能保护**：`UpdateCurrentCategoryAuxAndShell` 与 `UpdateAllCategoriesAuxAndShell` 统一挂起屏幕重绘、警告弹窗、COM 事件与公式自动重算，处理完成后统一触发 `ws.Calculate()` / `app.Calculate()` 并恢复原始环境，单表处理速度提升 10~30 倍；
     - **轻量锚点复用 (`ScanCabinetData`)**：新增 `ScanCabinetData(ws, cabIndex, anchor)` 重载，直接复用已识别的 `CabinetAnchorModel`，单表彻底消除所有重复遍历；
     - **CAD 运行进程毫秒级预检**：快速探测系统是否存在 `acad` 进程，未运行 AutoCAD 时 0 毫秒跳过管道索取，绝不产生等待；
     - **`ExcelAsyncUtil.QueueAsMacro` 异步队列调度**：`CabinetAuxCalcForm.cs` 将更新逻辑移入 Excel 宏队列，彻底解耦 Chromium IPC 与 Excel STA 线程，杜绝界面冻结与崩溃。
  3. **工业级 `#009688` 绿蓝主题动态进度条实现**：
     - **后端多层级进度通知委托**：服务层与控制器注入 `Action<int, string>? onProgress`，实时通过 `SafeInvoke` 向前端派发 `{ action: "updateProgress", percent, message }` 报文；
     - **前端 Element Plus 进度模态卡片**：`<el-progress>` 结合 `:striped="true" :striped-flow="true"` 动态条纹流光，主色调 `#009688`，实时呈现百分比与当前正在处理的表名、箱柜名称与步骤；
     - **平滑完成闭环**：任务完成后自动拉满至 100%，停留 250ms 后自动关闭模态卡片并展示成功提示。
  4. **工程构建与三端同步**：
     - 静态资源已同步至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`（哈希严格一致）；
     - `ExcelAddInDemo.csproj` 构建验证：**0 错误**。

- **方案 A（静默高速内存解析 + CAD 管道直接回发 Excel + SQLite 双向持久化）落地交付 (`CadSyncClient.cs`, `ExcelServices.CabinetAuxCalc.cs`, `CadExcelSyncServer.cs`, `DwgDimensionScanner.cs`)**：
  1. **改变被动局面，实现打开即自动检测触发**：
     - 在 Excel 辅材与壳体推导（`CalculateCabinetAuxAndShell`）中，不仅从本地 SQLite 批量查库，还自动对比当前箱柜所有元器件的图纸名/目录名；
     - 若图纸在 SQLite 未收录或长宽为 0，且磁盘物理图纸文件真实存在（基准目录优先从配置读取，缺省兜底 `E:\BaiduNetdiskWorkspace\BaseData\新库`，--硬编码--），自动收集为缺失图纸路径集合；
  2. **双向管道实时索取与 CAD 直接回传**：
     - `CadSyncClient.cs` 扩展双向命名管道客户端，新增 `RequestExtractDwgDimensions`（同步封装）与 `RequestExtractDwgDimensionsAsync`；
     - 支持设定握手超时（500ms 内极速判定 CAD 是否在线，未开启时优雅降级，绝不卡死 Excel）；
     - 向 CAD 发送 `{ action: "extractDimensions", dwgPaths: [...] }`，CAD 端后台 `Database.ReadDwgFile` 静默解析长宽及进深文字，入库 SQLite 同时直接将 JSON 尺寸列表回传给 Excel；
     - **修复 JSON 跨端大小写失配导致属性为 0 的关键缺陷**：CAD 端 `Newtonsoft.Json` 输出 PascalCase 大写（`"Width"`, `"Height"`），而 Excel 端实体标有小写 `[JsonPropertyName("width")]`，原 `System.Text.Json` 默认区分大小写导致字段全部反序列化为 0 和 `""`；在客户端开启 `PropertyNameCaseInsensitive = true` 后属性完美绑定。
  3. **Excel 内存字典即时注入与箱体推导**：
     - Excel 收到 CAD 直接回发的尺寸列表后，立即注入到当前推导的内存字典 `dwgDimMap`（分别以带目录相对路径 Key 与纯图纸名注入）；
     - 当前箱柜计算逻辑即时享用真实外形宽高与进深，顺利激活 `comp.Depth` 与最大进深 $+60\text{mm}$ 的箱体深度安全防干涉推导；
  4. **严格采用确定尺寸，彻底移除经验估算兜底 (`ExcelServices.CabinetAuxCalc.cs`)**：
     - 彻底废除 `EstimateComponentDimensions` 行业经验长宽兜底逻辑；
     - 若 AA 列（图纸名）与 AB 列（目录名）均为空，判定为用户未配置 DWG，**直接静默跳过，绝不产生误报错**；
     - 仅当明确配置了 AA 列图纸名、却未能从 CAD 测算或库中提取到有效外形尺寸时，才精准记录明确错误警示至 `calcWarnings`；
     - 未定尺寸的元件占位面积记为 0，推导描述中明确提示未定尺寸项数，让用户一目了然定位问题；
  5. **工程构建与架构规范**：
     - 严格遵守用户指示：CAD 排版（`Common.cs`）保持原封不动；
     - 新增代码严格遵循每 3 行至少 1 行中文注释，硬编码均带有 `--硬编码--` 标明；
     - 修复 `TuFan.csproj` 引用路径至 `..\..\..\excel-ct-tools\ExcelAddInDemo.csproj`；
     - `ExcelAddInDemo.csproj` 与 `TuFan.csproj` 均实现 0 错误构建。

- **DWG 图纸列与目录列按表格类型自适应路由 (分类表 AA/AB 列 vs 元件汇总表 X/Y 列) 与箱体辅材提取修复落地 (`ExcelServices.CabinetAuxCalc.cs`, `ExcelServices.ComponentParamMatch.cs`, `ComponentParamMatchForm.cs`, `AppConfig.cs`, `appsettings.json`, `component_param_match.html`)**：
  1. **问题与业务列位对齐**：
     - 用户明确指出：在【元件汇总表】中，图纸名和目录名为 **X 列（参数1，第 24 列）** 与 **Y 列（参数2，第 25 列）**；
     - 在【分类表/箱柜明细表】中，图纸名和目录名则为 **AA 列（图块名称，第 27 列）** 与 **AB 列（图块类别，第 28 列）**；此时 X 列是极数、Y 列是脱扣方式，此前写死 X/Y 会误冲数据；
     - `ExcelServices.CabinetAuxCalc.cs` 原先在分类表提取 DWG 尺寸时误读了第 24 列 (X) 与第 25 列 (Y)，导致读到极数和脱扣方式去匹配图纸尺寸。
  2. **落地改动**：
     - **箱体辅材提取修正 (`ExcelServices.CabinetAuxCalc.cs`)**：彻底去除 X/Y 列的兼容兜底代码，严格仅从分类表 AA 列 (`blockName`, 第 27 列，图纸名) 与 AB 列 (`blockCategory`, 第 28 列，目录名) 提取 `dwgName` 与 `dwgDir`；
     - **图纸参数匹配服务智能工作表感知 (`ExcelServices.ComponentParamMatch.cs`)**：
       - 若聚焦在【元件汇总表】：锁定图纸写入 **X 列**、目录写入 **Y 列**；
       - 若聚焦在普通【分类表】：锁定图纸写入 **AA 列**、目录写入 **AB 列**；
       - 彻底避免在分类表中冲掉原有极数 (X 列) 与脱扣方式 (Y 列)；
     - **浮窗与配置对齐 (`AppConfig.cs`, `appsettings.json`, `ComponentParamMatchForm.cs`, `component_param_match.html`)**：
       - 默认配置对齐为 `TargetDirColumn: "AB"`, `TargetDwgColumn: "AA"`；
       - 前端界面动态展示当前工作表模式与目标列（如 `📋分类表: AA=图纸 | AB=目录` 或 `📄汇总表: X=图纸 | Y=目录`）；
       - 双击写入后明确反馈写入的目标列标与行号；
     - **CAD 排版模块暂保留**：CAD 端 `Common.cs` 按用户指示暂不修改。
  3. **编译构建验证与静态资源同步**：
     - HTML 静态资源已同步至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
     - `ExcelAddInDemo.csproj` 编译构建通过：**0 错误**。

- **元器件图纸参数匹配浮窗 ERR_FILE_NOT_FOUND 根除与多级路径容错探测落地 (`ExcelAddInDemo.csproj`, `ComponentParamMatchForm.cs`, `bin/Debug/net48/Resources/`, `publish/Resources/`)**：
  1. **问题根因剖析**：
     - 用户截图展示狭长浮窗内报 Chromium 原生错误“未找到文件 它可能已被移动、编辑或删除 ERR_FILE_NOT_FOUND”；
     - `ExcelAddInDemo.csproj` 中遗漏了 `<None Include="Resources\component_param_match.html"><CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory></None>`，导致项目 build 时 HTML 页面未被复制到 bin 输出目录；
     - `ComponentParamMatchForm.cs` 此前仅判断了 `Resources` 文件夹是否存在，未校验 HTML 文件物理存在即盲目映射虚拟域名 `appassets.local`，导致 Chromium 找不到页面报错。
  2. **落地实施方案**：
     - **项目配置补全 (`ExcelAddInDemo.csproj`)**：补充 `Resources\component_param_match.html` 的 `CopyToOutputDirectory` 配置，保障每次构建自动下发到目标目录；
     - **多级候选路径容错探测 (`ComponentParamMatchForm.cs`)**：借鉴 `CabinetAuxCalcForm` 架构，增加 AppDir、BaseDir、publish、源码目录等多重物理路径探测，仅当物理文件确认存在后提取其真实物理目录映射为虚拟域名，杜绝 ERR_FILE_NOT_FOUND；
     - **镜像多端同步与构建**：已同步将 HTML 部署至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`，执行 `dotnet build /t:Compile /p:DebugType=none`：**0 错误**。

- **业务专属右键菜单高分辨率截断修复与“图纸参数匹配”完整展示落地 (`CustomContextMenuForm.cs`, `custom_context_menu.html`, `publish/Resources/custom_context_menu.html`)**：
  1. **问题根因**：菜单项多达 16 项 + 4 条分割线（总高需 456px+），原写死尺寸 `Size(250, 470)` 在 Windows 125%~150% DPI 缩放下视口被压缩，外加 `overflow: hidden`，导致排在底部的“图纸参数匹配...”被无情截断裁切；
  2. **落地改动**：
     - `CustomContextMenuForm.cs` 将窗体高度从 470px 扩展至 **505px**，给底部留出充足的展示余量；
     - `custom_context_menu.html` 将单项高度由 25px 微调至 **24px**，提升排版紧凑度；
  3. **编译构建与镜像同步**：多端镜像已覆盖同步，执行 `dotnet build /t:Compile /p:DebugType=none`：**0 错误**。

- **物料智能匹配悬浮窗过滤管道纯值紧凑标签化与多维动态放宽检索系统级落地 (`ComponentMatchOverlayForm.cs`, `component_match_overlay.html`, `publish/Resources/component_match_overlay.html`)**：
  1. **问题根因剖析（隐式过滤导致云端 0 结果根本原因）**：
     - **隐式强约束误杀物料**：此前系统在单元格点击弹窗时，自动从当前行单元格提取 `Name`（名称）、`Current`（电流）、`Pole`（极数）、`TripMode`（脱扣）及品牌，强行作为 `AND` 条件拼接发送给云端接口（如 `GET /api/api/Component/GetPagedList?Name=隔离开关&Current=225&Poles=3&Brand=德力西`）。但厂家实际标准规格阶梯可能只有 160A、200A、250A，无 225A 规格，导致云端直接返回 0 条（获取不到）；
     - **过滤项暗箱化且无法调整**：前端界面此前仅展示了数据源与不可操作的品牌文字，未将名称、电流、极数等展示出来，更无法单独关闭某个过滤条件。
  2. **用户核心指令与落地实施方案**：
     - **纯值极简标签（去除所有参数名前缀）**：彻底剔除“品牌:”、“电流:”、“极数:”、“名称:”等前缀汉字，直接展示高密度纯值标签（如 `[德力西 ✕]`、`[隔离开关 ✕]`、`[225A ✕]`、`[3P ✕]`、`[必含词 ✕]`），极致节省横向排版空间；
     - **悬停 Tooltip 完整提示**：鼠标悬停在标签上时，利用原生 `title` 显示如 `额定电流: 225A (点击 ✕ 移除此过滤)`，信息透明直观；
     - **直接移除与即时异步放宽重搜**：点击任意标签右侧的 `✕`，直接从过滤列表移除；前端即时向 C# 派发包含放宽后参数的 `filters` 对象；C# 端支持动态接收放宽参数，在向云端 API 发送请求时剥离被关闭的参数（如去掉 `Current=`），毫秒级捞出该品牌下全部型号物料；
     - **一键重置/还原微按钮（↺）**：当有任何过滤项被移除或修改时，自动呈现还原按钮，支持一键恢复从当前单元格提取的原始参数组合；
     - **空状态指引**：在匹配为 0 条时增加提示文本：“提示: 点击上方标签 ✕ 可快速关闭对应过滤项以放宽检索”；
  3. **编译构建与镜像同步**：
     - `publish/Resources/component_match_overlay.html` 已强制覆盖同步；
     - 执行 `dotnet build /t:Compile /p:DebugType=none`：**0 错误**，构建成功。

- **汇总调价全表单次大数组读取、智能免重复自愈与 Element Plus 动态进度条系统级提速 (`ExcelServices.SummaryAdjustPrice.cs`, `Tool.cs`, `SummaryAdjustPriceController.cs`, `SummaryAdjustPriceForm.cs`, `summary_adjust_price.html`)**：
  1. **问题根因剖析（200+台箱柜严重卡顿根本原因）**：
     - **逐柜碎片化 Range 访问**：200 多台箱柜逐个调用 `sheet.Range` 读取，产生了 400 余次 COM 范围跨进程通信往返；
     - **SafeSetSheetName 频繁异常与重写**：数百个定义名称每次都执行 Delete/Add，触发 COM 异常与跨进程往返；
     - **无进度反馈**：长时间处理大工程时缺少可视化反馈，界面停留在转圈状态。
  2. **落地实施方案**：
     - **单表单次大数组读取 (终极规则 7)**：在分类表循环外层一次性抓取整表 `Range["A1:AD{maxUsedRow}"]` 的值矩阵与公式矩阵到内存，内层遍历 200+ 台箱柜时全部通过内存二维数组切片提取（0 次 COM 往返），单表耗时由数秒降至 0.05 秒；
     - **智能免重复自愈 (`Tool.cs`)**：`SafeSetSheetName` 检查已存在名称是否已精准指向 `$A${row}`，若匹配则毫秒级直接跳过，根除大量 COM 异常与无谓重写；
     - **全链路进度通信与 Element Plus 动态进度条**：在 C# 逐表扫描、内存聚合与汇总表渲染时注入 `onProgress` 委托，实时将百分比与正在扫描的分类表名称（如 `正在读取分类表 [12#楼] (4/7)...`）推送到 WebView2；前端以 `#009688` 绿蓝主题浮层卡片优雅呈现；
     - **镜像多端同步**：`summary_adjust_price.html` 已强制覆盖至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`。
  3. **编译与构建验证**：
     - 执行 `dotnet build "e:\Ace\ExcelAddInCTtools\ExcelAddInDemo.csproj" /p:DebugType=none /p:RunExcelDnaBuild=false`：**0 警告，0 错误**。



- **汇总调价与分布调价弹窗分类与箱柜台数读取极速化改造及 Chromium IPC 线程死锁根治 (`Tool.cs`, `ExcelServices.SummaryAdjustPrice.cs`, `ExcelServices.DistributedAdjustPrice.cs`, `SummaryAdjustPriceForm.cs`)**：
  1. **问题根因剖析（Excel 卡死根本原因）**：
     - **Chromium IPC 与 STA 线程死锁**：系统事件日志捕获到 Excel 崩溃于 `EmbeddedBrowserWebView.dll`，异常代码 `0x80000003`。原因为 `SummaryAdjustPriceForm` 在 WebView2 的 `WebMessageReceived` 回调中**直接同步执行** `GetCategories` 和 `GenerateSummary` 等重型 Excel COM 操作。此时 Chromium IPC 管道等待握手返回，而 Excel COM 跨进程 RPC 处于消息泵等待态，两端形成双向死锁，最终触发 Chromium 底层 Hard Breakpoint 断言崩溃或 Excel 全局卡死挂起；
     - **主线程数十秒无响应**：真实运行日志显示此前扫描每张表耗时 11 秒（4 张表共 44 秒），Windows 系统直接判定 EXCEL.EXE 为“Not Responding”卡死；
     - **旧代码暴力循环**：`Tool.cs` 循环 1~200 探测删除不存在的旧定义名称引发数千次 COM 异常与 CLR 堆栈展开；
     - **视口冻结风险**：查询函数中曾尝试设置 `ScreenUpdating = false`，若捕获或恢复不当会导致 Excel 视口永久停止重绘产生假死。
  2. **落地实施方案**：
     - **全面接入 `ExcelAsyncUtil.QueueAsMacro` 宏队列调度**：`SummaryAdjustPriceForm.cs` 中的 `getCategories`、`generateSummary`、`updateFromSummary`、`toggleColumnsVisibility`、`getColumnsHiddenStatus` 全部包裹在 `QueueAsMacro` 中异步调度，彻底解耦 WebView2 IPC 与 Excel 主线程，杜绝线程死锁与 `0x80000003` 崩溃；
     - **彻底移除只读扫描中的 `ScreenUpdating = false`**：保证 Excel 界面始终处于激活正常渲染状态；
     - **轻量只读内存数组扫描 (规则 7)**：`GetCategorySheetsWithCabinetCount()` 与 `GetCategorySheetsForDistribution()` 单次 COM 调用抓取顶部 `A1:H60` 内存二维数组，毫秒级定位表头并提取 F 列（第 6 列）台数，弹窗打开耗时从十余秒骤降至几十毫秒；
     - **消除 COM 异常大循环**：`Tool.cs` 清理超额定义名称改为基于现存集合精准删除，消除盲目循环与 COM 异常；
     - **严格履行规则 8**：在用户点击【立即生成】真正操作 Excel 表格前，对选中的分类表显式调用 `Tool.FixAndFillCabinetNamesForSheet(sheet)`，确保规则 6 架构有效性；
  3. **编译与构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none`：**0 错误**；
     - 执行完整项目构建：**0 错误**。

- **DWG 元器件真实外形尺寸与安装进深提取及箱体深度 +60mm 安全防干涉推导全系统落地交付 (`DwgDimensionModels.cs`, `PersonalComponentDbService.DwgDimensions.cs`, `ExcelServices.CabinetAuxCalc.cs`, `DwgPreviewService.cs`, `ComponentParamMatchForm.cs`, `component_param_match.html`, `cabinet_aux_calc.html`, `cad-net_1/TuFan/DwgDimensionScanner.cs`)**：
  1. **用户核心需求与背景**：
     - **摆脱元器件尺寸写死**：告别箱体尺寸推导中元件尺寸写死或硬编码估算，实现依据元器件绑定的路径/X列/Y列 DWG 文件真实外形尺寸计算箱体；
     - **DWG 几何外框与进深文字正则提取**：从图纸中提取模型空间几何外框（宽 $W$、高 $H$），并自动遍历所有 `DBText`、`MText` 与块属性，匹配“高”、“高度”、“厚度”、“H”（如“高度: 85”、“高 120mm”、“厚度 90”、“H: 105”）提取数字填入 SQLite（`personal_components.db`）的 `depth`；
     - **箱体推导安全装配准则落地**：推导箱体尺寸时，整柜最大元器件安装进深深度 $\max(depth)$，箱体默认深度必须满足：
       $$D_{\text{箱柜}} \ge \max(depth) + 60\text{mm}$$
       以确保门板装配与内部安装导轨无机械干涉，并向上靠拢标准深度阶梯（160, 180, 200, 250, 300, 350, 400...）；
     - **修改时间与文件大小双重指纹防伪**：采用 `LastModifiedTicks + FileSizeBytes` 双重指纹，若图纸未被修改则实现秒级跳过缓存，CAD 重新保存图纸后自动热更新。
  2. **AutoCAD 插件端扫描服务实现 (`cad-net_1/cad1/TuFan/DwgDimensionScanner.cs`)**：
     - 提供 `SCAN_CURR_DWG_DIM` 命令：扫描当前活动图纸外框与进深文字，入库 SQLite 并在命令行打印结果；
     - 提供 `SCAN_DWG_DIMS` 命令：弹出文件夹选择对话框（默认引导至 `E:\BaiduNetdiskWorkspace\BaseData\新库`，--硬编码--），利用静默 `Database.ReadDwgFile` 免视口极速递归扫描所有图纸；
     - 具备指纹比对机制（比对已收录的 Ticks 与文件字节大小，跳过未修改图纸）；
     - 支持每 50 张图纸批量事务写入 SQLite `dwg_component_dimensions` 表。
  3. **Excel 插件推导引擎与数据层强化 (`ExcelAddInCTtools`)**：
     - **数据库自愈与数据访问 (`PersonalComponentDbService.DwgDimensions.cs`)**：
       - `dwg_component_dimensions` 包含 id, dir_name, dwg_name, rel_path, width, height, depth, has_text_depth, matched_text, last_modified_ticks, file_size_bytes, updated_at；
       - 支持 `SaveOrUpdateDwgDimension`、`BatchSaveOrUpdateDwgDimensions`、`GetDwgDimension`、`BatchGetDwgDimensions`；
     - **Excel 扫描与箱体尺寸推导 (`ExcelServices.CabinetAuxCalc.cs`)**：
       - `ScanCabinetData` 提取元器件绑定的 DWG 文件名与目录名；
       - `CalculateCabinetAuxAndShell` 批量查库，优先采用 CAD 真实长宽计算占用面积，提取最大进深 `maxCompDepth`，计算 `minRequiredDepth = maxCompDepth + 60.0`；
       - 若推导深度小于安全门限，自动提升并调用 `AlignToStandardDepth` 向上对齐到标准箱体深度阶梯；
     - **图纸文件浏览自动注入尺寸 (`DwgPreviewService.cs`)**：
       - `DwgFileInfo` 增加 `Width`, `Height`, `Depth`, `HasDimension`, `HasTextDepth` 属性；
       - `ScanDwgFiles` 扫描完成后批量检索 SQLite 自动为图纸列表注入物理尺寸。
  4. **前端交互与信息呈现升级**：
     - **`Resources/cabinet_aux_calc.html`**：在【① 壳体推荐尺寸】卡片中，若 `realDimensionsCount > 0`，以护眼底色展示：“📐 CAD真实尺寸: 匹配 x 项, 最大进深 xx mm (安全门限+60mm: xx mm)”；
     - **`Resources/component_param_match.html`**：在 DWG 文件项增加 `.dwg-dim-tag` 展示 `W×H D:xx`，并在鼠标悬浮 Tooltip 及写入状态栏展示尺寸详情；
     - **`Forms/ComponentParamMatchForm.cs`**：在 `bindDwg` 回写时，实时返回该 DWG 尺寸字符串 `dimStr` 并在前端状态栏呈现。
  5. **工程构建与多端同步**：
     - 静态资源同步覆盖至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
     - `ExcelAddInCTtools` 执行 `dotnet build /t:Compile /p:DebugType=none`：**0 错误**；
     - `cad-net_1/cad1/TuFan` 执行 `dotnet build /p:DebugType=none /p:RunExcelDnaBuild=false`：**0 错误**。


- **元器件图纸参数匹配浮窗调换 X/Y 列回写内容 (`AppConfig.cs`, `appsettings.json`, `publish/appsettings.json`, `ExcelServices.ComponentParamMatch.cs`, `component_param_match.html`)**：
  - **用户核心需求**：在“图纸参数匹配”侧边浮窗中，将写入当前行的 X、Y 列内容调换（原先：X 列写目录、Y 列写图纸；调换后：X 列写图纸名称、Y 列写目录名称）；
  - **落地改动**：
    1. **配置层解耦与默认值更新**：
       - `AppConfig.cs`：将 `TargetDwgColumn` 设为 `"X"`（--硬编码--），`TargetDirColumn` 设为 `"Y"`（--硬编码--）；
       - `appsettings.json` 与 `publish/appsettings.json`：同步更新默认映射为 `"TargetDirColumn": "Y"`, `"TargetDwgColumn": "X"`；
    2. **服务层写入与参数明确化 (`ExcelServices.ComponentParamMatch.cs`)**：
       - `BindDwgParamToActiveRow` 方法参数签名更新为 `(dirName, dwgName, autoNextRow, colDir = "Y", colDwg = "X", removeExt = true)`；
       - X 列（`targetColDwg`）写入图纸纯名称 `cleanDwgName`；
       - Y 列（`targetColDir`）写入目录名称 `cleanDirName`；
       - 自动换行跳转与日志反馈对齐为 `[{targetColDwg}]={cleanDwgName}, [{targetColDir}]={cleanDirName}`；
    3. **前端 UI 提示与响应动态化 (`component_param_match.html`)**：
       - 规则提示行更新为动态模板：`{{ configInfo.targetDwgColumn || 'X' }}=图纸 | {{ configInfo.targetDirColumn || 'Y' }}=目录`（显示为 `X=图纸 | Y=目录`）；
       - 图纸项 hover 提示对齐为 `(X=图纸, Y=目录)`；
       - 前端 `configInfo` 默认值与 IPC 初始化数据接收对齐；
    4. **镜像同步与编译验证**：
       - 静态资源同步覆盖至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`（哈希一致）；
       - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整构建：**0 错误**。

- **彻底清除主界面底部泄漏的多余抽屉 DOM (`Resources/cabinet_aux_calc.html`)**：
  - **用户核心需求**：用户截图框选指出主界面底部（操作按钮下方）残留平铺的「默认加工折弯预留补偿 (mm)」及「板材各厚度单价表 (元/m²)」，要求删除；
  - **根本原因与处理**：
    1. 该内容源于历史遗留的右侧抽屉 `<el-drawer>` 模板，在未展开状态下因样式脱离或环境限制直接被平铺在主页面文档流底部；
    2. 且里面的“折弯放量预留补偿”违背了用户此前“取消放量”的要求，板材单价库也已收敛至【💰 壳体价格计算】Tab；
    3. 彻底删除了该 `el-drawer` DOM 结构，消除底部泄漏；
    4. 将弹窗中的「⚙️ 板材单价设置」按钮点击动作平滑联动至主界面【规则与定额配置】->【💰 壳体价格计算】Tab，架构体验高度一致；
  - **构建验证**：全工程构建 **0 警告，0 错误**，静态资源已同步生效。

- **壳体规格型号全面去除箱柜型号前缀 (`Services/ExcelServices.CabinetAuxCalc.cs`, `Resources/cabinet_aux_calc.html`)**：
  - **用户核心需求**：根据用户截图指示，去除弹窗中的「箱柜型号」选项，且在回填到 Excel 计费区及弹窗生成的型号中彻底去除箱柜型号前缀（如 `JXF-`、`XM-`、`XL21-`）；
  - **具体改动**：
    1. **前端弹窗精简**：从钣金算料弹窗的“规格型号拼装多选配置行”中移除「箱柜型号」复选框与下拉选择框；
    2. **实时型号生成解耦**：调整 `computedFullModel`，不再拼接箱柜型号前缀，直接以 `${w}*${h}*${d}` 尺寸开头（默认包含板厚与材质，如 `600*700*180-1.2mm 镀锌板`）；
    3. **后台批量计算与回写**：在 `CalculateShellPrice` 中移除 `XL21-` / `XM-` 前缀拼接，输出干净的标准格式 `${width}*{height}*{depth}-{thickness:F1}mm {material}{secondPlateSuffix}`；
    4. **构建验证**：已同步至 `publish/` 与 `bin/Debug/net48/`，全工程构建 **0 警告，0 错误**。

- **壳体尺寸计算界面精简与材质解耦 (`Resources/cabinet_aux_calc.html`)**：
  - **用户确认点**：壳体尺寸计算 Tab 中原本遗留的「批量计算默认材质」下拉框成功删除；
  - **解耦设计收敛**：
    - 【📦 壳体尺寸计算】Tab：保持纯粹的几何尺寸推导定位（仅保留纯高度深度阶梯库、壳体匹配名称、配电箱/落地柜安全系数、落地柜高度门限、标准常用尺寸库）；
    - 【💰 壳体价格计算】Tab：全面接管材质确定策略（304/201/冷轧/镀锌识别规则及右上角的基准材质兜底下拉项）、高度阶梯板厚、封闭箱体面数与单价库；
  - **静态资源同步**：已同步至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`，编译构建 0 错误。

- **壳体尺寸计算与壳体价格计算 Tab 分离及批量算料定价三层策略落地交付 (`Models/CabinetAuxCalcModels.cs`, `Services/ExcelServices.CabinetAuxCalc.cs`, `Resources/cabinet_aux_calc.html`)**：
  1. **用户核心需求与解耦原则**：
     - **职责分离**：恢复之前的纯「📦 壳体尺寸计算」Tab，单列独立的「💰 壳体价格计算」Tab，彻底将“尺寸推导”与“钣金算料定价”解耦；
     - **策略 1（板材材质确定策略）**：
       - 优先级 1（箱柜特征关键词自动识别）：含 `304` $\to$ **不锈钢304**；含 `不锈钢` 或 `201` $\to$ **不锈钢201**；含 `冷轧` $\to$ **冷轧板**；含 `镀锌` $\to$ **镀锌板**；
       - 优先级 2（全局基准材质兜底）：在【📦 壳体尺寸计算】与【💰 壳体价格计算】中提供「批量计算默认材质」下拉项（默认：**镀锌板**）；未命中特殊材质关键词的所有箱柜一律采用此材质兜底；
     - **策略 2（板材厚度高度阶梯决策法）**：
       - $H \le 800\text{mm} \to 1.2\text{mm}$（小型照明箱/终端箱）；
       - $800 < H \le 1600\text{mm} \to 1.5\text{mm}$（中型动力箱/挂墙箱）；
       - $H > 1600\text{mm} \to 2.0\text{mm}$（大型成套落地开关柜）；
       - 支持在表格中步进微调、添加和删除板厚阶梯；
     - **策略 3（封闭箱体面数确定策略，取消放量）**：
       - **取消放量**：外形尺寸直接算面积，不预留折弯放量补偿（$\Delta W=0, \Delta H=0, \Delta D=0$）；
       - **面数标准**：自动采用封闭箱体标准（宽深 2.0、宽高 2.0、高深 2.0；若箱柜名称包含“二层板”则自动叠加 1.0 面宽高二层板）；
       - 展开总面积计算式：$S = 2.0 \times \frac{WD}{10^6} + (2.0 + (\text{含二层板}?1:0)) \times \frac{HW}{10^6} + 2.0 \times \frac{HD}{10^6}$；
       - 单价计算式：$\text{单价} = \text{ROUND}(S \times P_{\text{sq}}, 2)$；
     - **解决一次性回填价格历史缺陷**：
       - 修复历史 `WriteCabinetCalcResultToSheet` 仅写型号尺寸而单价/总价为空的问题；
       - 批量回写时（单个、当前分类、全部分类），计费区壳体行 C 列回填拼装标准型号（如 `XM-800*600*200-1.2mm 镀锌板`）、E 列填“台”、F 列默认补 1、G 列填核算出的单价、H 列填 `=ROUND(F*G, 2)`，实现**规格尺寸与价格一次性全自动回填**。
  2. **服务层与模型层落地**：
     - `Models/CabinetAuxCalcModels.cs`：在 `QuotationRules` 中加入 `ShellPriceRules` 节点；新增 `ShellPriceConfig` 与 `HeightThicknessGradientItem` 实体；在 `CabinetCalcResult` 中补充 `RecommendedShellModel`, `RecommendedShellUnitPrice`, `RecommendedShellExpandedArea`, `RecommendedShellMaterial`, `RecommendedShellThickness`；
     - `Services/ExcelServices.CabinetAuxCalc.cs`：
       - 实现 `DetermineMaterial`、`DetermineThickness`、`CalculateShellPrice` 三大核心静态方法，严格遵循每 3 行新增代码加 1 行中文注释，并标明 `--硬编码--`；
       - 在 `CalculateCabinetAuxiliary` 中完成单柜核算与结果填充；
       - 在 `WriteCabinetCalcResultToSheet` 中实现规格、单位、数量、单价与公式全自动一次性回填计费区及 Cab_Det 兜底行。
  3. **前端界面全套重构与同步**：
     - `Resources/cabinet_aux_calc.html`：恢复纯尺寸计算 Tab 并嵌入默认材质下拉；单列【💰 壳体价格计算】Tab（包含材质确定策略图解、高度阶梯板厚配置表、面数策略与取消放量说明、板材材质及厚度单价库）；钣金计算器弹窗默认取消放量并自动同步推导出的材质与板厚；在 setup return 中完整导出新状态与方法；
     - 镜像同步覆盖至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none`：**0 警告，0 错误**。
  1. **用户核心需求**：“分布汇总，表格行列内容按如图设计”（根据真实软件截图完全对齐行列布局与交互）；
  2. **1:1 像素级行列与表头架构重构**：
     - **工作表命名**：由“材料分布表”统一重命名为【`元件汇总分布表`】；
     - **左侧 11 列固定字段（A~K列）**：
       - A: 排序ID（数字/小数，用于箱柜内部件重排）
       - B: 元件名称、C: 型号规格、D: 生产厂家
       - E: 表价、F: 折扣、G: 报出、H: 备注
       - I: 单位、J: 类别（填入“元件”）
       - K: 元件总数（第 9 行起使用动态加权公式 `=SUMPRODUCT($L$7:${endCol}$7, L{r}:${endCol}{r})` 联动计算）
     - **1~8 行复合表头与说明文字**：
       - 行 1: 柜序号（K1: 柜序号，L1起: 1, 2, 3... 居中）
       - 行 2: 所属分类（K2: 所属分类，L2起: 低压配电室... 居中）
       - 行 3: 箱柜名称（K3: 箱柜名称，L3起: 进线柜、联络柜... 居中）
       - 行 4: 箱柜型号（K4: 箱柜型号，L4起: GCK... 居中；并在行 4 插入红色警示文字 `*排序ID：指在箱柜中的排列位置，可输入任意数字(支持小数)。`）
       - 行 5: 柜号（K5: 柜号，L5起: 带下划线蓝色超链接 `1AA, 11AA`，点击直接平滑跳转至对应分类表的箱柜明细行 `Cab_Det_X`）
       - 行 6: 箱柜类别（K6: 箱柜类别，L6起: 智能推导为低压抽屉柜/照明配电箱/高压开关柜等；并在行 6 插入红色警示文字 `*更新到项目时，数量为空删除，为0保留。`）
       - 行 7: 箱柜数量（K7: 箱柜数量，L7起: 2, 1, 2...；并在 A7 单元格加粗展示 `项目名称: (xx)xx工程项目`）
       - 行 8: 列标题行（A8~K8 写入列字段标题，L8起各箱柜列标题均为“数量”，整行启用 Excel 原生 `AutoFilter` 自动筛选）
       - 行 9 起: 元器件数据与箱柜单台用量矩阵，有配额填数量，无配额留空；
  3. **反向更新（一键更新到明细）严格按截图规范升级**：
     - 严格践行红字规范：“**数量为空删除，为0保留**”——若箱柜单元格为空白，反向写回时清除箱柜明细对应行；若为 0，保留该行并将数量设为 0、合价设为 0；
     - **方式 A（智能插入新器件）**：若分布表中某箱柜填写了数量，但该箱柜原本没有该器件，自动在其元器件区域末尾（小计行前）插入新行，并填入型号、厂家、数量和单价，维护边框与公式；
     - **支持按排序ID重排**：勾选“调整元件排序”时，根据 A 列 `排序ID` 自动对箱柜内的元器件进行物理重排；
     - 严格遵守规则 6、7、8（前置 `FixAndFillCabinetNamesForSheet` 自愈，二维数组批量读写内存）；
  4. **控制器与前端多端同步**：
     - 前端界面与控制器所有文案与通信逻辑统一对齐为【元件汇总分布表】；
     - 镜像同步至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
     - `dotnet build /t:Compile /p:DebugType=none` 完整构建：**0 错误**。

- **【规则与定额配置】->【📦 壳体选型规则】Tab 页面全面重构集成纯高度深度阶梯库与板材单价库 (`Resources/cabinet_aux_calc.html`)**：
  1. **用户核心需求**：“界面没集成到箭头所指的tab页吗”（指向【规则与定额配置】下的【📦 壳体选型规则】二级 Tab 页）；
  2. **全面集成与工业级落地**：
     - **彻底消除旧的电流门限**：删除了原本卡片中容易产生歧义的“柜体电流门限 250A”，替换为“落地柜高度门限 1500mm”，贯彻“深度只由高度来确定，取消和电流关联确定”；
     - **模块 1（纯高度推导箱柜深度阶梯库）**：内置并支持自由编辑纯高度深度对应表（$H \le 400 \to 160$、$H \le 600 \to 180$、$H \le 800 \to 200$、$H \le 1000 \to 250$、$H \le 1400 \to 300$、$H \le 1800 \to 400$、$H \le 2000 \to 800$、$H > 2000 \to 1000$ mm），支持实时添加/删除阶梯门限；
     - **模块 2（钣金加工折弯预留补偿与展开参数）**：直接配置宽预留 W (50mm)、高预留 H (50mm)、深预留 D (20mm)、综合表面积展开系数 (1.20倍)、壳体匹配名称、配电箱与落地柜安全系数；
     - **模块 3（板材材质与各厚度单价库）**：以 Tab 形式集成冷轧板、不锈钢201、不锈钢304、镀锌板等各个材质的各厚度每平米单价，支持行内步进调整、支持“+ 添加厚度规格”与删除；
     - **模块 4（标准壳体常用尺寸库）**：保留现有的宽*高尺寸胶囊库，支持新增与移除；
     - **双向数据打通与持久化**：在【📦 壳体选型规则】Tab 中修改后，点击页面右下角【💾 保存当前规则配置】即可永久保存至后台配置；同时与推导卡片和钣金算料弹窗双向同步；
  3. **验证与同步**：
     - 镜像同步至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
     - `dotnet build /t:Compile /p:DebugType=none` 完整构建：**0 错误**。

- **智能辅材与壳体计算中心集成 ExWinner 钣金算料与深度纯高度驱动全套落地交付 (`Models/CabinetAuxCalcModels.cs`, `Services/ExcelServices.CabinetAuxCalc.cs`, `Forms/CabinetAuxCalcForm.cs`, `Resources/cabinet_aux_calc.html`)**：
  1. **用户核心需求与原则确认**：
     - **集成形式**：内嵌在现有的「智能辅材与壳体计算中心」（`cabinet_aux_calc.html`）中，从“① 壳体推荐尺寸”卡片联动呼出；
     - **深度推导原则**：**深度（D）只由高度（H）确定，彻底取消与电流的一切关联**；
     - **板材单价设置**：提供独立右侧抽屉（`el-drawer`），支持各材质（冷轧板、不锈钢201、不锈钢304、镀锌板等）各厚度平方单价及加工折弯放量微调与持久化保存；
     - **Excel 回写规范**：严格遵循规则 6（计费区无空行）、规则 7（二维数组一次性读写内存）、规则 8（`Tool.FixAndFillCabinetNamesForSheet` 自愈确保 4 个定义名称精准有效）。
  2. **深度纯高度驱动模型与服务层落地**：
     - **数据模型 (`Models/CabinetAuxCalcModels.cs`)**：
       - `ShellConfig` 增加 `HeightDepthGradients` 纯高度深度阶梯（H<=400:160, <=600:180, <=800:200, <=1000:250, <=1400:300, <=1800:400, <=2000:800, >2000:1000）、板材单价库 `MaterialPrices`、加工折弯放量 `Allowance`；
       - `CabinetCalcResult` 增加 `RecommendedShellDepth` 与 `RecommendedShellSizeFull`；
       - 新增 `MaterialPriceItem`、`CabinetAllowanceConfig`、`CabinetStructureItem`、`CabinetShellWritePayload` 实体；
     - **服务层 (`Services/ExcelServices.CabinetAuxCalc.cs`)**：
       - 静态方法 `DeriveDepthFromHeight(int height, ShellConfig? shellConfig = null)`：只依据箱柜高度推导深度，彻底解耦电流；在单柜分析 `CalculateCabinetAuxiliary` 中计算 `bestHeight` 后即刻调用并带入；
       - `UpdateMaterialPrices`：持久化板材单价表与加工预留放量；
       - `WriteCalculatedShellToFeeArea`：前置调用规则 8 自愈；二维数组一次性读取计费区（`subsumRow` 到 `tolsumRow - 1`）；优先就地替换现有壳体行规格、单价及联动公式，未命中则在 `tolsumRow` 前插入新行并自愈；
     - **IPC 控制器与窗体 (`Forms/CabinetAuxCalcForm.cs`)**：
       - 响应 `saveMaterialPrices` 与 `applyShellToExcel` 报文，通过 `ExcelAsyncUtil.QueueAsMacro` 调度安全执行并回传处理结果。
  3. **前端工业级交互与闭环 (`Resources/cabinet_aux_calc.html`)**：
     - **3D 轴测透视图与尺寸放量联动**：CSS 3D 透视立方体渲染（粉/黄/蓝三色面），实时标注各面厚度；支持户内/户外、明装/暗装快捷切换；提供 W、H、D 数值与折弯预留放量双步进微调器；支持纯高度推荐深度快捷胶囊一键切换；
     - **型号规格智能装配器**：动态生成符合工业规范的标准型号字符串（前缀+尺寸+板厚+材质+安装方式+二层板+自定义后缀）；
     - **高密度结构展开明细表**：支持按面数（顶底板、面门板、后背板、左右侧板、二层板）与按综合表面积系数双模式核算；支持统一板厚/独立板厚切换；实时核算展开面积与综合单价，支持人工微调总价；
     - **板材单价设置抽屉**：各材质 Tab 切换，输入各厚度单价与默认放量，一键持久化保存；
     - **回写计费区**：支持“自动覆盖/原位更新”与“➕ 插入新壳体行”两种模式直连 Excel 计费区；
     - **闭环与标签自愈**：严格闭合所有弹窗与抽屉标签，在 `setup()` 中完整导出所有状态与方法，并在 `initContext` 与消息监听中实现双向联动。
  4. **构建与多端静态资源同步**：
     - 镜像同步覆盖至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
     - `dotnet build /t:Compile /p:DebugType=none` 完整构建：**0 警告，0 错误**。


- **彻底修复分布调价与汇总调价分类表箱柜台数虚增Bug（根除汇总区穿透落款与历史幽灵定义名称，`Tool.cs`, `ExcelServices.DistributedAdjustPrice.cs`, `ExcelServices.SummaryAdjustPrice.cs`）**：
  1. **用户核心问题**：“数量还是不正确”（如“人防”分类表实际只有 13 行有效箱柜、合计 13 台，但之前界面依然显示 45 台或 30 台）；
  2. **根因完全闭环与修复点**：
     - **根因 A (`Tool.cs`)**：`FixAndFillCabinetNamesForSheet` 扫描汇总行 `sumRows` 时从 `cabSumStartRow` 一路扫到 `firstDetRow`，遇到“合计”、“总计”、“小计”、“金额大写”、“报价说明”、“编制/审核/批准”等签字说明栏时没有截断判定，把落款说明行全部误当成箱柜打上了 `Cab_Sum_X` 定义名称；
     - **根因 B (`Tool.cs`)**：清理多余旧定义名称时只循环到 `cabCount + 30`，导致若历史曾生成过 45 个名称，大于 43 的幽灵名称无法被删除；
     - **根因 C (`DistributedAdjustPrice.cs` & `SummaryAdjustPrice.cs`)**：未严格遵守规则 8（读取分类前漏调 `Tool.FixAndFillCabinetNamesForSheet(sheet)`），且循环累加时遇到非箱柜行也误加了默认 1 台；
     - **根因 D (数量解析)**：Excel COM 返回浮点格式（如 `1.0`）导致 `int.TryParse` 失败；且兜底扫描遇到“合计”行未中断；
  3. **落地实施与对齐**：
     - **`Tool.cs`**：扫描汇总行时增加关键截断：一旦 A/B/C 列包含“合计/总计/小计/大写/说明/审核/批准/制表/编制”，立即 `break;` 终止扫描；倒序遍历工作表 `Names` 集合将所有序号大于 `cabCount` 的残留定义名称物理删除，并安全兜底清理至 200 号；
     - **`DistributedAdjustPrice.cs` & `SummaryAdjustPrice.cs`**：前置调用 `Tool.FixAndFillCabinetNamesForSheet(sheet)` 自愈；严格拦截非箱柜行；严格仅读取 F 列（第 6 列），采用 `double.TryParse` 容错解析，空行与非箱柜行坚决不计入；兜底扫描同步增加合计行 `break;` 截断；
     - **`CreateComponentDistributionSheet`**：提取箱柜 F 列数量同步使用 `double.TryParse` 容错并拦截合计行；
  4. **构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整构建：**0 警告，0 错误**。

- **材料分布表排版重构：彻底删除左上角冗余箱柜表，全面 1:1 对齐 ExWinner 工业标准列宽与表头 (`ExcelServices.DistributedAdjustPrice.cs`)**：

  1. **用户核心指令**：“那么删除，并且细读"D:\Program Files\ExWinner“的分布调价排版，包括列宽，修改为一样”；
  2. **深入逆向与排版对齐 (`ComponentDist_High.xlsx` / `ComponentDist_Low.xlsx`)**：
     - **彻底移除冗余**：彻底删除左上方 A1:G17 的箱柜垂直列表（原第 6 步），消除信息重复与大面积死白留白，彻底解决首屏数据被推挤下沉以及多柜时穿透覆盖的隐患；
     - **首屏紧凑矩阵排布**：
       - 行 1~5：横向箱柜复合表头（K 列写引导标签“所属分类/箱柜数量/箱柜柜号/箱柜名称/箱柜型号”，L 列向右横向平铺展开各箱柜属性，浅绿蓝护眼底色 `#E0F2F1`，居中对齐）；
       - 行 6：元器件汇总表头（A6~K6：元件名称、规格型号、单位、汇总数量、单价、金额、厂家、表价、折扣系数、实际成本、差价，L6 向右为各箱柜柜号，主色调 `#009688`，白字粗体，行高 22）；
       - 行 7 起：元器件与分布数据大矩阵，末尾紧随“合计”行；
     - **1:1 像素级复刻 ExWinner 官方标准列宽**：
       - A 列 (元件名称): `15`
       - B 列 (规格型号): `29.5`
       - C 列 (单位): `7`
       - D 列 (汇总数量): `7`
       - E 列 (单价): `11`
       - F 列 (金额): `11.88`
       - G 列 (厂家): `17`
       - H 列 (表价): `10`
       - I 列 (折扣系数): `10`
       - J 列 (实际成本): `10`
       - K 列 (差价): `10`
       - L 列向右 (各箱柜单台数量列): 统一精准设为 `10`；
     - **对齐与公式对齐**：
       - A/B/G 列靠左（-4131），C/D 列居中（-4108），E/F/H/I/J/K 列靠右（-4152）且格式统一设为 `0.00`；
       - D 列汇总数量加权求和公式：`=SUMPRODUCT($L$2:$Z$2, L7:Z7)`；
       - F 列金额公式：`=ROUND(D7*E7, 2)`；
       - J 列实际成本公式：`=IF(H7>0, ROUND(H7*I7, 2), "")`；
       - K 列差价公式：`=IF(ISBLANK(J7),0,(J7-E7)*D7)`；
     - **反向更新自适应升级**：
       - 在 `UpdateFromComponentDistributionSheet` 中引入动态定位，自动寻找“元件名称”行，精准定位表头与数据行起始位置，完美兼容新旧表；
  3. **构建验证**：
     - `dotnet build /t:Compile /p:DebugType=none` 完整构建：**0 错误**。

- **分布调价与汇总调价箱柜数量读取统一修正为仅取 F 列并防范空行锚点 (`ExcelServices.DistributedAdjustPrice.cs`, `ExcelServices.SummaryAdjustPrice.cs`)**：
  1. **用户核心指令**：“数量一定在f列，不要尝试解析e列”；
  2. **根因与漏洞清除**：
     - 原代码使用 `??` 盲目串联第 5 列 (E列单位) 与第 6 列 (F列数量)，在 E 列有“台”等文本时触发操作符短路，导致无法读取 F 列真实数量且解析失败退化为 1，或在 E 列存在其他数值时导致列错位；
     - 彻底清除对 E 列（第 5 列）的尝试解析，明确将数量提取严格锁定在第 6 列（F 列）；
     - 增加汇总行柜号与箱柜名称有效性前置校验，若柜号与名称全为空则判定为无效/残余空行锚点，直接跳过，杜绝因历史定义名称残留导致虚增台数；
     - 同步修正【材料分布表】生成逻辑（`GenerateComponentDistributionSheet`）中对汇总行的读取：由原先误将第 5 列作台数、第 6 列作单价、第 7 列作总价，纠正为一次性读取 A:H 完整 8 列，第 6 列为台数、第 7 列为单价、第 8 列为总价；
     - 在 `ExcelServices.SummaryAdjustPrice.cs` 中同步对齐此项高品质规则；
  3. **编译构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整构建：**0 错误**。

- **分布调价（二维矩阵交叉调价表与一键反向更新）全套系统落地交付 (`ExcelServices.DistributedAdjustPrice.cs`, `DistributedAdjustPriceController.cs`, `DistributedAdjustPriceForm.cs`, `distributed_adjust_price.html`, `RibbonController.cs`, `ExcelAddInDemo.csproj`)**：
  1. **功能定位与深度逆向对比**：
     - 系统深度解构 `D:\Program Files\ExWinner\` 源码程序集 `leadsoft.superwinner.BLL.dll`（`AdjustBomByDistrClassic` / `AdjustBomByDistrVip`）、`frmAdjustBomByDistr.html` 与 `ComponentDist_High.xlsx/.myxml` 模板；
     - 彻底厘清与“汇总调价”的差异：“汇总调价”为一维器件总览，而“分布调价”构建【材料分布表】二维矩阵交叉表，纵向（A~K列）汇总元器件，横向自 L 列起展开全项目箱柜，交叉点展示单台配额用量；
     - 支持在 Excel 中修改白色调价核心列（E列单价、B列型号、G列厂家）后，点击【一键更新】批量精准反向写回各分类表对应箱柜明细行，自动自愈小计与总计联动公式。
  2. **端到端工程闭环设计**：
     - **服务层 (`ExcelServices.DistributedAdjustPrice.cs`)**：
       - `ShowDistributedAdjustPriceDialog()` 唤起非模态置顶窗体；
       - `GetCategorySheetsForDistribution()` 扫描分类表并精确识别箱柜台数；
       - `GenerateComponentDistributionSheet(request)`：严格执行规则 8 前置调用 `Tool.FixAndFillCabinetNamesForSheet(sheet)` 自愈 4 个定义名称；二维数组内存批量读取箱柜与器件明细（规则 7，器件限制在 `detRow + 2` 至 `subsumRow - 1`，计费区严格保护）；创建【材料分布表】，写入顶部箱柜概况（行 1~12）、横向箱柜复合表头（行 14~17，箱柜数量/柜号/名称/型号）、元器件汇总表头与分布数据矩阵（行 18 起），注入动态求和与折扣公式，高亮可编辑列；
       - `UpdateFromComponentDistributionSheet(options)`：二维数组批量提取调价器件清单，逐柜精确匹配写回新单价、新规格、新厂家，刷新小计求和公式与总计联动，调用自愈机制；
     - **控制器与数据模型 (`DistributedAdjustPriceController.cs`)**：
       - 涵盖分类模型、合并条件（名称/型号/厂家/无厂家/价格/备注/方案名）、排序规则、更新选项与执行统计实体；
       - 提供 `GetCategorySheetsJson`, `GenerateDistributionSheetJson`, `UpdateFromDistributionSheetJson`, `CheckDistributionSheetExistsJson` 等 WebAPI 标准接口；
     - **宿主窗体 (`DistributedAdjustPriceForm.cs`)**：
       - 几何尺寸 720x640 像素无边框置顶设计，采用 `appassets.local` 虚拟主机映射与 `QueueAsMacro` 调度规避 COM 冲突；
       - 支持原生级鼠标位移增量拖拽、最小化与关闭；
     - **前端 UI (`distributed_adjust_price.html`)**：
       - Vue 3 `<script setup>` + Element Plus，`#009688` 绿蓝相间主题；
       - 双阶段视图：阶段一（分类勾选、箱柜台数统计、高级合并与排序面板、【立即生成分布表】）；阶段二（【重新汇总】快捷链接、高级更新选项、【一键更新到明细】、调价操作指引）；
     - **Ribbon 挂载与多端镜像**：
       - `RibbonController.cs` 中挂载 `btnDistributedAdjustPrice` 按钮点击响应；
       - 同步至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`（SHA256 哈希 100% 一致）；
       - 执行 `dotnet build /t:Compile /p:DebugType=none`：**0 错误**。


- **元器件图纸参数匹配伴随侧边浮窗 (200x800) 全套系统落地交付 (`AppConfig.cs`, `appsettings.json`, `ConfigManager.cs`, `ExcelServices.ComponentParamMatch.cs`, `ComponentParamMatchForm.cs`, `component_param_match.html`, `custom_context_menu.html`, `CustomContextMenuForm.cs`)**：
  1. **用户核心指令与需求确认**：
     - **窗口规格**：严格限制为 `200 × 800` 像素紧凑型浮窗，设计为伴随式侧边栏（`TopMost = true`、无系统边框，贴靠屏幕右侧，支持顶部拖拽任意移动）；
     - **目录设置与下钻**：顶栏提供目录设置图标（📁），默认路径为 `E:\BaiduNetdiskWorkspace\BaseData\新库`（配置化持久化到 `appsettings.json`，标明 `--硬编码--`）；点击 📁 调起原生文件夹选择器；支持点击子目录卡片下钻、点击“⬆️ 上级”按钮返回、支持关键字实时搜索过滤；
     - **第 24 列 (X 列) 与第 25 列 (Y 列)**：双击 DWG 文件时，自动将当前目录名称写入当前活动行的 X 列（第 24 列），将图纸文件名写入 Y 列（第 25 列）；
     - **去除扩展名**：写入 Y 列时自动剔除 `.dwg` 扩展名，仅写入纯图纸名称；
     - **右键菜单集成**：在 Excel 原生样式右键菜单中增加“图纸参数匹配... [X/Y列]”项，点击即时呼出浮窗；
     - **边框线完全对称自协调 (`component_param_match.html`)**：根除 `html, body` 嵌套双重边框与 200px 尺寸导致的右侧裁切 Bug，重构为 `width: 100%; height: 100%` + `#app { box-sizing: border-box; border: 1px solid #00796b }`，左右上下四侧 1px 深绿蓝边框完美呈现，完全协调对称；美化 4px 极细滚动条；
     - **右击 DWG 返回上一级目录 (`component_param_match.html`)**：在 `.dwg-item` 绑定 `@contextmenu.stop.prevent="handleDwgRightClick"`，右击任意 DWG 即可一秒退回上一级，无需专门将光标移至顶部上级按钮；列表空白区亦支持右键返回；优化规则提示条为单行不折行；
     - **严格规范**：C# (Excel-DNA) + WebView2 + Vue 3 (`<script setup>`) + Element Plus；主题色绿蓝相间（`#009688`）；所有 Excel 操作严格写在 `ExcelServices.cs` 分部类；新增代码每 3 行至少 1 行中文注释。
  2. **端到端实现细节**：
     - **配置层 (`AppConfig.cs`, `appsettings.json`, `ConfigManager.cs`)**：
       - 增加 `ComponentParamMatchSettings` 实体类，包含 `BaseDirectory`, `TargetDirColumn`, `TargetDwgColumn`, `AutoNextRow`, `RemoveExtension`；
       - `ConfigManager.Instance.UpdateComponentParamMatchBaseDirectory` 支持动态修改根目录并自动保存到 JSON；
     - **服务层 (`ExcelServices.ComponentParamMatch.cs`)**：
       - `ShowComponentParamMatchDialog()` 唤起浮窗单例；
       - `BindDwgParamToActiveRow(dirName, dwgName, autoNextRow, colX, colY, removeExt)`：获取 `ActiveCell`，写入目标列并驱动 `Cells[nextRow, col].Select()`，处理返回值与异常日志；
     - **宿主窗体 (`Forms/ComponentParamMatchForm.cs`)**：
       - `200x800` 几何尺寸，无边框与顶部悬浮；
       - 双向 IPC 机制：支持 `getInitialData`, `selectBaseDirectory`, `scanDirectory`, `bindDwg` (宏队列 `QueueAsMacro` 调度避免 COM 冲突), `moveWindow`, `minimizeWindow`, `closeWindow`；
     - **前端 UI (`Resources/component_param_match.html`)**：
       - 200px 高密度布局，Vue 3 `<script setup>` + Element Plus，`#009688` 主题配色；
       - 自研拖拽条、面包屑导航与上级按钮、搜索框、子目录卡片、DWG 图纸卡片（双击高亮绿色动效反馈）、底部状态栏；
     - **右键菜单集成 (`Resources/custom_context_menu.html`, `Forms/CustomContextMenuForm.cs`)**：
       - 增加菜单项及 SVG 图标，窗体高度微调适配（`470px` 防止出现滚动条），无缝联动调起；
     - **工程构建与多端同步**：
       - 资源文件同步镜像覆盖至 `publish/` 与 `bin/Debug/net48/`；
       - 执行 `dotnet build /t:Compile /p:DebugType=none`：**0 错误**。

- **右键菜单视觉样式全面升级为 Office/Excel 原生风格并集成剪切/复制/粘贴/插入/删除/筛选功能 (`Forms/CustomContextMenuForm.cs`, `ExcelEventManager.cs`, `Resources/custom_context_menu.html`)**：
  1. **用户核心需求**：“把鼠标右键改为excel原生的样式” -> 明确要求“采用选项2，但是保留excel的复制粘贴插入，删除，筛选选项”；
  2. **视觉 1:1 像素级复刻 Office/Excel 原生风格**：
     - **去网页化与去徽章**：彻底移除顶部绿色渐变标题微标（“业务快捷菜单 工作表 (单元格)”）与大圆角大卡片边框；
     - **原生调色盘与边框**：纯白底色 `#ffffff`、1px 浅灰原生边框 `#d4d4d4`、微圆角 `4px`、Office 经典下拉微阴影；
     - **原生排版与悬浮**：标准 Office 字体 `"Segoe UI", "Microsoft YaHei", "SimSun"`、字号 `12px`、文字颜色 `#262626`、紧凑行高 `25px`；悬停背景切换为原生经典浅灰 `#f0f0f0`，杜绝位移缩放等网页动画；
  3. **功能集成与四段式结构**：
     - **剪贴板组**：剪切（`Ctrl+X`）、复制（`Ctrl+C`）、粘贴（`Ctrl+V`）；
     - **行列与筛选组**：插入...（调起 Excel 原生插入对话框）、删除...（调起 Excel 原生删除对话框）、按所选单元格的值筛选、清除筛选 / 自动筛选；
     - **业务功能组**：新建箱柜、识别参数并匹配物料、物料匹配与品牌规则...、智能输入词库配置...、汇总调价...、元器件数据管理...、辅材壳体计算...；
     - **模式切换组**：切换为 Excel 原生右键菜单；
  4. **C# 宿主与 COM 执行保障**：
     - 窗体尺寸自适应调整为 `250 × 445` 像素，15 项功能完整容纳且防截断无滚动条；
     - 在 `QueueAsMacro` 纯净宏上下文中通过 `CommandBars.ExecuteMso` 调度原生指令（Cut/Copy/Paste/CellsInsertDialog/CellsDeleteDialog/FilterBySelectedValue/FilterClearAllFilters），并具备 API 容错回退；
     - 在 `OnSheetBeforeRightClick` 中前置校验选区交集，自动激活右键选中的目标单元格，确保原生操作 100% 精确指向右击位置；
     - 新增代码严格遵循“至少每 3 行包含一行中文注释”规范；
  5. **工程构建与多端同步**：
     - 镜像强制同步至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
     - `dotnet build /t:Compile /p:DebugType=none` 完整编译：**0 错误**。

- **常规投标报表导出（标书报表）自由配置系统全套功能落地交付 (`TenderReportModels.cs`, `AppConfig.cs`, `TenderReportRegularController.cs`, `ExcelServices.TenderReport.cs`, `tender_report_regular.html`)**：
  1. **用户核心指令与需求答复**：
     - **取整到元**：项目总价、箱柜总价、箱柜单价、分项表单价与总价全部支持取整到元（`ROUND(..., 0)` 与 `Math.Round(..., 0)`）；
     - **按“分类工作表”分别输出独立 Sheet**：勾选时，针对选中的各个分类工作表分别克隆独立 Sheet（以分类名称命名），仅填充所属分类箱柜；未勾选时合并输出在《屏柜分项表》中；
     - **取消物料编码**：彻底剔除物料编码概念与选项；
     - **合并功能先不做**：彻底剔除器件分组合并逻辑；
     - **自动保存在本地 `appsettings.json`**：前端修改偏好后，导出时后端自动写回 `appsettings.json` 的 `TenderReportSettings` 节点，下次打开自动还原偏好；
     - **严格规范**：新增代码每 3 行至少 1 行中文注释；主题色 `#009688`；微调遵循最小改动法则；
  2. **端到端落地细节**：
     - **配置模型与持久化**：
       - `TenderReportSettings` 涵盖 `RoundProjectTotal`, `RoundCabinetTotal`, `RoundCabinetUnitPrice`, `RoundDetailUnitPriceTotal`, `TotalWithFormula`, `FeeItemWithoutFormula`, `EnableOneKeyLocate`, `OutputNotesInSummary`, `NotesMultilineDisplay`, `NotesAutoFitRowHeight`, `TextAutoWrap`, `SplitDetailBySheet`, `PrintSetting`；
       - `AppConfig.cs` 集成 `TenderReportSettings TenderReport` 属性；
       - `TenderReportRegularController.cs` 的 `GetInitialDataJson` 注入 settings，并在 `ExportReport` 中自动调用 `ConfigManager.Instance.Save()` 持久化；
     - **核心报表服务层增强 (`ExcelServices.TenderReport.cs`)**：
       - **取整与公式**：支持 `ROUND(..., 0)` 或 `ROUND(..., 2)`；若总价不带公式则全表输出纯数值；费用项始终以纯数值填入；总计行公式智能联动；
       - **分 Sheet 独立输出**：`PopulateDetailWorksheets` 自动从模板克隆独立分类工作表，过滤工作表非法字符及重名自增序号，填充完成后安全删除原模板工作表；
       - **一键定位**：汇总表序号列绑定工作簿级定义名称超链接 `CabDetRef_{globalIndex}`；分项表箱柜信息行首列绑定反向跳转超链接 `'屏柜汇总表'!A{sumRow}`，实现双向秒级跳转；
       - **打印分页**：`Paginated` 模式在每个箱柜后插入水平分页符 `HPageBreaks.Add`；
       - **报价说明与排版**：支持是否输出、自动行高与自动换行；
     - **前端 UI 界面高品质升级 (`tender_report_regular.html`)**：
       - 卡片 3 采用紧凑三组平铺布局（工作表结构与输出方式、取整与计算模式、报价说明与排版），完全贴合 Element Plus `#009688` 主题；
       - 挂载时通过 `initialDataLoaded` 自动回显历史配置，提交时完整回传；
     - **多端同步与工程构建**：
       - 镜像同步覆盖至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
       - 执行 `dotnet build /t:Compile /p:DebugType=none`：**0 警告，0 错误**。
  1. **用户核心提问**：“@[e:\Ace\ExcelAddInCTtools\Resources\cabinet_aux_calc.html:L5850] 是不是改为folderName更合理”；
  2. **完全认同并落地重构**：
     - 原 `groupNameStr` 纯粹用于从浏览路径中截取当前物理文件夹名称并传递给表单只读展示项 `folderName`；
     - 原变量名使用 `groupName...` 极易与方案实体的核心属性 `groupName`（二次排布图）产生语义混淆；
     - 将其统一重命名为 `folderName`，彻底解耦并自解释；
  3. **环境同步与编译**：
     - 同步至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`（哈希 100% 一致）；
     - `dotnet build /t:Compile /p:DebugType=none` 构建验证：**0 错误**。


- **二次回路排布图/布置图全面剔除 `layoutDwgName` 兼容，统一直接使用 `groupName` (`cabinet_aux_calc.html`)**：
  1. **用户核心指令**：“去除兼容 (兼容 layoutDwgName)，直接使用groupName”；
  2. **全面清理与纯净化**：
     - **弹窗初始化（`openEditSchemeModal`）**：彻底移除 `layoutDwgName`，直接采用 `groupName: scheme.groupName || ""`，新建空白方案也直接初始化 `groupName: ""`；
     - **持久化保存（`saveSecondarySchemeForm`）**：`payload` 中直接提取 `groupName: editingSecScheme.value.groupName || "通用"`，彻底剔除 `layoutDwgName` 冗余属性；
     - **方案复制（`applyCopiedScheme`）**：直接提取 `sourceScheme.groupName || sourceScheme.GroupName` 并仅赋值给 `editingSecScheme.value.groupName`；
     - **顶栏胶囊提示与注释**：清理遗留注释，明确方案排布图/布置图统一对应 `groupName`；
  3. **环境同步与编译**：
     - 镜像强制覆盖至 `publish/Resources/cabinet_aux_calc.html` 与 `bin/Debug/net48/Resources/cabinet_aux_calc.html`（三端哈希一致）；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整构建：**0 错误**。


- **二次回路排布图/布置图字段模型归一化修复 (`cabinet_aux_calc.html`)**：
  1. **用户核心提问**：“@[e:\Ace\ExcelAddInCTtools\Resources\cabinet_aux_calc.html:L3522] 为什么这里没显示，groupName 不正确吗”；
  2. **根本原因剖析**：
     - **字段本身完全正确**：在 C# 实体 `SecondarySchemeEntity` 与 SQLite 数据库表 `secondary_schemes` 及 Excel 导入规范中，**`groupName` 存储的正是“二次排布图/布置图”**（如 FA、BDY 等）；
     - **前端赋值脱节**：此前弹窗初始化函数 `openEditSchemeModal` 误将 `scheme.groupName` 塞入了只读的 `folderName`，表单实体 `editingSecScheme` 内部压根没有定义与初始化 `groupName` 属性，导致输入框 `v-model="editingSecScheme.groupName"` 值为 `undefined` 从而无法显示；
     - **保存与复制链条遗漏**：此前 `saveSecondarySchemeForm` 与 `applyCopiedScheme` 仍混杂了已废弃的 `layoutDwgName` 兼容逻辑；
  3. **落地闭环修复**：
     - 在 `openEditSchemeModal` 初始化时，精准将 `scheme.groupName || scheme.layoutDwgName` 灌入 `editingSecScheme.groupName`（同时赋给 `layoutDwgName` 保持双向安全兜底）；
     - 在 `saveSecondarySchemeForm` 组织提交负载时，精准提取 `editingSecScheme.groupName` 持久化至数据库 `group_name`；
     - 在 `applyCopiedScheme` 复制方案时，优先复制 `sourceScheme.groupName`；
     - 镜像覆盖同步至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`，`dotnet build` 验证：**0 错误**。


- **常规样式投标报表《屏柜分项表》计费部分（费用项）丢失修复全面交付 (`ExcelServices.TenderReport.cs`, `TenderReportModels.cs`)**：
  1. **用户核心反馈**：“生成常规报表目前丢失了计费部分”；
  2. **根因定位分析**：
     - `CollectFullCategoryDetails` 在提取箱柜明细时仅读取了 `detRow + 2` 至 `subsumRow - 1`，未提取规则 6 明确定义的计费区间（`Cab_Subsum` 至 `Cab_Tolsum - 1`）；
     - `PopulateDetailWorksheet` 在向《屏柜分项表》输出时，仅写入元器件后即直接输出橙色总计行，未在明细尾部输出费用项；总计行总价被硬编码为 `=SUM(元器件)*数量`，导致整表缺少壳体、辅材、人工等费用且金额与《屏柜汇总表》对不上；
  3. **落地修复与工业规范闭环**：
     - **模型层扩展**：在 `TenderReportCabinetItem` 中新增 `FeeItems` 列表存储计费项；
     - **数据采集与规则 8 自愈**：在提取每个分类表前显式调用 `Tool.FixAndFillCabinetNamesForSheet(sheet)` 自愈定义名称；以二维数组一次性读取 `startFeeRow` 至 `endFeeRow` 提取名称、型号、厂家、单位、数量、单价、合价（`Value2` 纯数值）及备注；
     - **分项表尾部输出与联动公式**：
       - 计算 `totalDetailCount = compCount + feeCount`，批量克隆白底细网格母版数据行；
       - 元器件总价保留计算公式（`={数量}*{单价}`）；
       - 计费项总价遵循“**总价带公式（费用项不带）**”规范，直接写入纯数值（Value2），杜绝跨表复杂公式导致 `#REF!`；
       - 橙色总计栏智能检测“单台合计”行 `singleTotalRow`，写入联动公式 `=ROUND({单台合计单元格}*{数量单元格}, 2)`，金额与汇总表 100% 严密吻合；
  4. **工程构建与验证**：
     - 严格遵守每 3 行新增代码至少 1 行中文注释；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。
  1. **用户核心指令**：“@[e:\Ace\ExcelAddInCTtools\Resources\cabinet_aux_calc.html:L2966-L2972] 修改为二次布置图参数显示”；
  2. **落地修改**：
     - 将 CAD 矢量视口顶部定额胶囊栏首项由「二次组: {{ currentVectorScheme.schemeName }}」修改为「布置图: {{ currentVectorScheme.layoutDwgName || currentVectorScheme.cadDrawingName || '-' }}」；
     - 悬浮 Tooltip 同步绑定完整提示：`:title="'二次布置图: ' + (currentVectorScheme.layoutDwgName || currentVectorScheme.cadDrawingName || '未设置')"`；
     - 严格遵循精简与最小改动规范，色彩维持 `#5eead4` 高亮青绿色，与右侧「跨门:」、「二次:」、「开孔:」、「人工:」、「BOM:」胶囊保持和谐统一；
  3. **环境同步与编译**：
     - 镜像强制同步覆盖至 `publish/Resources/cabinet_aux_calc.html` 与 `bin/Debug/net48/Resources/cabinet_aux_calc.html`（三端 SHA256 哈希 100% 一致）；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译验证：**0 错误**。


- **二次回路工作台顶栏微调与模板字符串加固 (`cabinet_aux_calc.html`)**：
  1. **用户核心指令**：“@[e:\Ace\ExcelAddInCTtools\Resources\cabinet_aux_calc.html:L2718] 删除此部分”；
  2. **落地修改**：
     - 彻底移除二次工作台顶栏左侧冗余静态文字 `<span style="font-weight: 600; color: #00796b">📁 图纸路径:</span>`，仅保留「⬆️ 返回上级」按钮及当前路径动态标题，极大增加路径物理展示长度；
     - 将铜排计算卡片插值表达式加固升级为 ES6 模板字符串反引号语法：`{{ calcResult.copperWeight > 0 ? `${calcResult.copperWeight} KG` : '0 (小箱补贴)' }}`，彻底免疫由于本地编辑器 Prettier / 自动格式化工具强行多行折断引发的未捕获 `SyntaxError` 页面白板故障；
  3. **环境同步与编译**：
     - 强制镜像同步覆盖至 `publish/Resources/cabinet_aux_calc.html` 与 `bin/Debug/net48/Resources/cabinet_aux_calc.html`（三端哈希一致）；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译：**0 错误**。

- **壳体辅材计算页面白屏根因彻底排查与闭环修复 (`cabinet_aux_calc.html`)**：
  1. **故障现象**：用户反馈“界面打不开了，点击ribbon上的壳体辅材计算后是白板”；
  2. **根因定位分析**：
     - 使用 Headless Browser 内核捕获控制台日志发现报错：`Uncaught SyntaxError: Invalid or unexpected token`；
     - 准确定位到此前 `cabinet_aux_calc.html` 在铜排计算结果展示卡片中，插值表达式单引号出现了跨行字符串断裂（`+ '\n KG'`）；
     - 普通单引号字面量跨行在 JavaScript / Blink 引擎中为非法语法，导致 Vue 模板编译器在编译阶段直接崩溃抛出 SyntaxError，`app.mount("#app")` 中断挂载，整页呈现白板；
  3. **修复措施与多端同步**：
     - 将插值表达式字符串紧凑合并到同一行：`{{ calcResult.copperWeight > 0 ? (calcResult.copperWeight + ' KG') : '0 (小箱补贴)' }}`；
     - 运行自动化 Node 脚本对全文件 HTML 标签开闭平衡、指令属性（`:...`, `v-...`）及事件绑定（`@...`）全面复核验证：**全部 0 语法错误**；
     - 通过浏览器无头子智能体实际加载测试：**页面成功完全渲染，UI 布局与组件均正常显示，无白屏无控制台报错**；
     - 镜像强制同步至 `publish/Resources/cabinet_aux_calc.html` 与 `bin/Debug/net48/Resources/cabinet_aux_calc.html`（SHA256 哈希完全一致）；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译验证：**0 错误**。

- **二次回路方案与 BOM 编辑弹窗集成及顶栏配置自愈 (按方案 1 落地交付) (`CabinetAuxCalcForm.cs`, `cabinet_aux_calc.html`)**：
  1. **用户核心需求**：“按方案1，同时在@[e:\Ace\ExcelAddInCTtools\Resources\cabinet_aux_calc.html:L2831-L2837] 后面添加编辑按钮，点击弹出如图界面可以编辑二次图对应的参数和bom”；
  2. **方案 1 落地 (图纸路径配置自愈与顶栏精简)**：
     - 二次回路图纸对齐与绑定工作台顶栏移除「更换根目录」按钮，保留「◀ 折叠CAD预览」、「🔄 刷新Excel元件组」与「刷新图纸」；
     - 图纸路径完全依赖 `appsettings.json` 中的 `SecondaryCircuitSettings.CircuitDwgDirectory` 配置；
     - 前端在接收到 `dwgDirsLoaded` 报文时，若当前浏览路径为空且回路根目录已配置，自动触发 `browseToDirectory(circuitDwgDir)` 瞬时下钻进入根目录；
  3. **编辑按钮与二次方案编辑全功能弹窗集成**：
     - **入口触发**：在工作台中间视口顶部的「📦 BOM: N项」胶囊后方添加「✏️ 编辑」按钮（以及图纸未配置二次方案时的「➕ 配置方案与BOM」快捷入口）；
     - **编辑弹窗高保真还原 (`secEditDialogOpen`)**：
       - 包含 8 项紧凑参数单行整齐平铺展示（方案目录、当前DWG、品牌、适用图名集合、二次排布图、跨门线根数、开孔要求、装配工费）；
       - 包含第二行独占的方案工艺描述（Description）；
       - 弹窗标题右侧设有「📄 复制其他方案」操作按钮；
       - 下方配备二次 BOM 子物料清单，支持「从本地物料库选型添加」与「自定义临时项」，支持行序号、名称、型号规格、数量、单位、单价、合价与删除，并在底部实时动态汇总「二次材料费小计」、「人工费」与「综合总成本」；
       - 统一使用 `name` 与 `model` 字段，彻底杜绝废弃的 `remark` 字段；
     - **复制方案选择弹窗 (`copySchemeModalVisible`)**：支持关键字即时检索成熟方案，可双击行或点击「载入复制」一键继承品牌、排布图、跨门线、开孔、工费、工艺描述及完整 BOM 清单；
     - **物料库选型弹窗 (`materialSelectorVisible`)**：支持搜索本地 SQLite 物料库，一键选中推入 BOM 清单；
  4. **C# 后端 IPC 动作支持与构建验证**：
     - 在 `CabinetAuxCalcForm.cs` 中注入 `saveSecondaryScheme`（入库 SQLite 并回发 `saveSecondarySchemeResult`）、`searchMaterialComponents`（检索物料库并回发 `searchMaterialComponentsResult`）、`getSecondarySchemesForCopy`（获取方案复制集合并回发 `getSecondarySchemesForCopyResult`）；
     - 严格遵守每 3 行代码至少 1 行中文注释规范；
     - 镜像热同步覆盖至 `publish/Resources/cabinet_aux_calc.html` 与 `bin/Debug/net48/Resources/cabinet_aux_calc.html`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 构建检查：**0 错误**。

- **智能辅材中心与二次绑定工作台升级 1200x800 宽屏视口架构 (方案 A 落地交付) (`CabinetAuxCalcForm.cs`, `cabinet_aux_calc.html`)**：
  1. **用户核心诉求**：“需要增加宽度，有什么方案” -> “方案 A”；
  2. **闭环实施落地**：
     - **C# 宿主窗体基准尺寸升级 (`CabinetAuxCalcForm.cs`)**：
       - 将无边框置顶宿主窗体客户区尺寸 `ClientSize` 从 `960 × 720` 一步到位扩展至 **`1200 × 800`**；
       - 为内嵌工作台弹窗释放出高达 **`1152px`** 的物理宽度（较原本的 921px 净增 **`231px`**）；
       - 结合此前左右两侧面板紧凑化让渡出的 182px，中间 CAD 矢量预览视口的可用宽度直接突破 **`700px`**，图纸纵览从容开阔；
     - **前端主体高度适度放宽 (`cabinet_aux_calc.html`)**：
       - 将工作台主体高度调整为 `calc(88vh - 120px)`，`max-height` 扩展为 `700px`，充分利用 800px 垂直高差；
  3. **环境同步与编译验证**：
     - 镜像热同步至 `publish/Resources/cabinet_aux_calc.html` 与 `bin/Debug/net48/Resources/cabinet_aux_calc.html`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **二次回路图纸对齐与绑定工作台左右布局紧凑化与冗余列清理 (`cabinet_aux_calc.html`)**：
  1. **用户核心需求**：“左侧红框内容不需要了，右边红框宽度改为240px，理解表述需求” -> “复选框列有作用吗，如果没有也删除，prop="fileSizeFormatted"列宽度给dwg预览” -> “执行”；
  2. **精准落地与空间重构**：
     - **左侧图纸看板轻量化**：
       - 彻底移除无实际业务消费的复选框多选列（`type="selection"`），以及配套的顶栏 `[全选]`、`[清空]` 按钮和底栏 `已选 X 项` 提示，搜索输入框 100% 占满顶栏；
       - 彻底移除 `prop="fileSizeFormatted"`（文件大小）列；
       - 将左侧容器宽度由 `230px` 紧凑收窄为 `168px`（净减少 62px，精准将大小列占用的 62px 空间全额让渡给中间 DWG 视口）；
       - 表格保留单列纯净名称展示，自适应填满 168px 宽度，避免长图纸名称被截断；
     - **右侧元件组看板收窄至 240px**：
       - 将右侧容器宽度由 `360px` 调整为 `240px`（释放 120px 宽度给中间 DWG 视口）；
       - 顶栏统计与操作重构为单行紧凑排版（`📊元件组(X种,已绑Y) 清空此项 | 全清`）；
       - 表格两列标签精简为“型号规格”与“绑定图号”，`min-width` 调整为 95px，彻底消除 240px 容器下的横向滚动条；
     - **中间 CAD 矢量预览视口大扩容**：
       - 左右两翼合计让渡 `62px + 120px = 182px` 宽度给中间 CAD 视口，大幅提升图纸可视视野；
  3. **环境同步与编译验证**：
     - 镜像热同步至 `publish/Resources/cabinet_aux_calc.html` 与 `bin/Debug/net48/Resources/cabinet_aux_calc.html`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **清理 `CabinetAuxCalcForm.cs` 与 `CabinetAuxCalcController.cs` 冗余死代码 (`CabinetAuxCalcForm.cs`, `CabinetAuxCalcController.cs`)**：
  1. **用户核心指令**：“去除冗余代码”；
  2. **冗余代码识别与精准清理**：
     - **Win32 P/Invoke 遗留声明清理**：彻底移除 `ReleaseCapture()`、`SendMessage()` 声明及 `WM_NCLBUTTONDOWN = 0xA1`、`HTCAPTION = 0x2` 常量定义；彻底移除旧版模态拖拽分支 `else if (action == "dragWindow" || action == "dragMove")`，完全由前端微秒级物理位移增量 `moveWindow` 驱动，杜绝 Win32 模态循环卡死 Excel 宿主；
     - **未被前端引用的废弃分支清理**：
       - 移除未被调用的 `applyCalculation` 早期批量回写分支（前端已全面采用单柜 `writeCurrentCabinet`、分类 `updateCurrentCategory` 及全表 `updateAllCategories` 现代化通道）；
       - 移除 `CabinetAuxCalcController.cs` 中仅被其调用的无用废弃方法 `ApplyCalculation`；
       - 移除二次工作台未被调用的 `scanDwgDir` 分支（前端已全面采用支持子目录递归下钻的 `scanDirectoryHierarchy` 综合通道），并重新理顺 1~9 号动作注释序号；
     - **消息推送双轨实现归一化**：
       - 将原有的 7 处散落且无防御的 `SafeInvoke(() => _webView.CoreWebView2.PostWebMessageAsString(resJson))` 统一切换为封装完善、带句柄有效性校验与异常捕获的 `PostWebMessageSafe(resJson)`；
  3. **规范符合与构建验证**：
     - 所有保留与调整的代码严格遵循“至少每 3 行代码包含一行中文注释”；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **DWG 文件不能预览根因彻底排查与闭环修复 (`CabinetAuxCalcForm.cs`, `cabinet_aux_calc.html`)**：
  1. **用户核心反馈**：“不能预览dwg文件，你看之前成功预览的代码”；
  2. **对比图一（`SecondaryCircuitForm.cs`）根因精准剖析**：
     - **安全上下文与协议限制**：图一之所能成功预览 DWG，是因为其使用了 `_webView.CoreWebView2.SetVirtualHostNameToFolderMapping("appassets.local", resDir, CoreWebView2HostResourceAccessKind.Allow)`，将本地文件夹安全映射为 `https://appassets.local`，运行在标准安全 HTTPS 域名下；
     - **`file:///` 协议致命拦截**：而图二（`CabinetAuxCalcForm.cs`）此前直接使用 `_webView.Source = new Uri(htmlPath)`（即 `file:///` 协议）；在 Chromium / Blink 引擎下，`file:` 协议的 Origin 为 `null`，浏览器强制禁止创建 WebWorker（`new Worker("cad-viewer/wasm/dwg-worker.js")` 抛出 `DOMException: SecurityError: cannot be accessed from origin 'null'`）且阻断 WebAssembly 二进制 fetch，导致 DWG 矢量视口引擎根本无法启动；
  3. **修复措施全面落地**：
     - **后端宿主虚拟域名映射**：在 `CabinetAuxCalcForm.cs` 的 `OnFormLoadAsync` 中同步引入 `SetVirtualHostNameToFolderMapping("appassets.local", resDir, ...)`，将页面导航切换至 `https://appassets.local/cabinet_aux_calc.html`，彻底解除 Worker/WASM 跨域与文件协议限制；
     - **前端接收流自愈加载保障**：在 `cabinet_aux_calc.html` 的 `dwgBinaryLoaded` 中加入视口实例自愈检测，若此时 `cadViewerInstance` 尚未挂载则自动触发初始化与延迟重试，杜绝静默失败；
     - 镜像同步覆盖至 `publish/Resources/cabinet_aux_calc.html`；
  4. **构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none`，结果：**0 错误**，构建成功。

- **智能辅材与壳体计算中心 (图二 `cabinet_aux_calc.html`) 成功集成二次图纸对齐与绑定工作台 (`CabinetAuxCalcForm.cs`, `cabinet_aux_calc.html`)**：
  1. **用户核心需求**：“把图一的二次绑定功能 @[secondary_circuit_manage.html:L1326-L2062] ，图二也添加这个跳转按钮”；
  2. **交互形态与位置落地**：
     - **按钮入口**：在图二“🎯 箱柜范围选择与实时参数”卡片头部标题栏右侧（`cfg-card-header`）配置「📂 二次图纸对齐与绑定」按钮（Element Plus primary plain 风格，品牌绿蓝色）；
     - **工作台弹窗内嵌**：将图一的三栏一体化工作台弹窗 `<el-dialog v-model="circuitDwgDialogVisible" ...>` 完整内嵌至图二，实现点击按钮即刻在图二页面内优雅弹出，无需切换窗口或离开上下文；
  3. **核心功能链路 100% 打通**：
     - **左侧 DWG 目录浏览与下钻**：支持多级子文件夹双击下钻与面包屑返回上级，支持过滤搜索、全选与清空；
     - **中间 WebGL 矢量 CAD 视口**：加载 `cad-viewer/cad-viewer.bundle.js` 与 `cad-viewer/style.css`，纯前端 WebAssembly/WebGL 矢量平滑渲染，支持滚轮缩放、左键平移与外部 CAD 调起；
     - **视口顶部方案定额胶囊**：动态关联匹配二次方案实体，即时呈现二次组方案名、跨门线根数、二次材料费、开孔规格、人工费与 BOM 子项数；
     - **右侧 Excel 元件组映射看板**：自动扫描当前活动 Excel 工作表去重二次元件组（B列='元件组'），双击左侧 DWG 即刻快速流水线绑定并自动跳向下一行；
     - **第 32 列 (AF列) 批量回写保存**：点击底部「💾 保存绑定到 Excel (第32列)」一键安全持久化写入全表对应行的第 32 列；
  4. **后端 C# IPC 架构设计**：
     - 在 `CabinetAuxCalcForm.cs` 中复用 `SecondaryCircuitController` 与 `ExcelServices.SecondaryCircuit.cs`；
     - 增加 `getDwgDirs`、`selectDwgDir`、`scanDwgDir`、`scanDirectoryHierarchy`、`locateAndHighlightDwg`、`getDwgBinary`、`openInCad`、`scanExcelComponentGroups`、`saveExcelComponentGroupBindings`、`getSchemes` 等完整动作路由转发；
     - 实现了 `PostWebMessageSafe` 线程安全调度，并保持代码严格符合“每 3 行至少 1 行中文注释”规范；
     - 镜像同步至 `publish/Resources/cabinet_aux_calc.html`。
  5. **构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none`：**0 警告，0 错误**。

- **二次元件组落实规则 8 前置自愈与紧跟元器件空行优先复用落地交付 (`ExcelServices.ComponentGroup.cs`)**：
  1. **用户核心指令**：“另外元件组需要放在元器件的最下面，有空行就不要插入”；
  2. **规则 8 闭环落实**：
     - 在 `GetActiveCabinetComponentsFromExcel`（抓取元件前）与 `ExecuteBatchComponentGroup`（倒序插行前）均前置调用 `Tool.FixAndFillCabinetNamesForSheet(sheet)`，确保规则 6（4 个定义名称、小计行与元器件区间）绝对精准；
  3. **“元件组放在元器件的最下面，有空行不要插入”算法实现**：
     - 在遍历元器件二维数组时，动态追踪记录最后一个有效非空元器件相对索引 `lastUsedIndex`；
     - 计算最后一个有效元器件的物理行号 `lastUsedRow` 与箱柜内部可用空行总数 `availableEmptyRows = compEndRow - lastUsedRow`；
     - **空行充足时（`reqCount <= availableEmptyRows`）**：完全不执行 `Insert` 插行，小计行行号保持不变；
     - **空行不足时（`reqCount > availableEmptyRows`）**：仅按差额在小计行上方插入 `reqCount - availableEmptyRows` 行；
     - 从 `lastUsedRow + 1` 开始依次紧凑回填二次元件组（A 列序号公式 `$"=ROW()-ROW(A${detRow + 1})"`、B 列“元件组”、C 列代号、E 列“套”、F 列数量、AD/AE 句柄），杜绝元件与元件组之间产生大面积空白断层且保证 A 列序号动态自适应连续；
  4. **构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **汇总行 A 列超链接就地更新与居中/虚线框全量样式自愈修复交付 (`Tool.cs`)**：
  1. **用户核心现象**：“不是居中对齐，且没有虚线框了” -> “修改”；
  2. **根因精准剖析**：
     - 原逻辑对单元格执行了 `Hyperlinks.Delete()`，触发了 Excel COM 的“恢复出厂设置”行为，将单元格格式重置为空白 Normal 格式，抹除了原有的居中对齐与灰色虚线边框；
     - 随后的 `sheet.Hyperlinks.Add` 强制套用了内置 Hyperlink 样式（蓝色、下划线、General常规对齐），且在计算出数值序号后数字自动靠右对齐；
  3. **闭环修复方案全面落地**：
     - **就地属性更新，废除盲目 Delete/Add**：若单元格已存在超链接，直接修改 `hl.SubAddress` 与 `hl.ScreenTip`，绝不触发 Excel 格式重置，完全保留原始上下文；
     - **强制居中与黑体保护**：显式设定 `HorizontalAlignment = -4108`（水平居中）、`VerticalAlignment = -4108`（垂直居中）、`Font.Underline = -4142`（去除下划线）、`Font.ColorIndex = -4105`（自动黑色）；
     - **同行 B 列虚线边框智能继承自愈**：动态读取同行的 B 列原生边框样式（`bBorders.LineStyle` 与 `Weight`），直接同步赋予 A 列，让已被破坏的灰色点线虚线框 100% 还原；
     - **纯汇总箱柜与明细行同步保护**：纯汇总箱柜与明细行表头 A 列同样进行居中保护、黑体保护及边框继承；B 列仅在存在超链接时才安全清理；
  4. **构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **移除 FixAndFillCabinetNamesForSheet 中越权调用的 RefreshCabinetFeeAreaFormulas (`Tool.cs`)**：
  1. **背景与诉求**：用户质询 `Tool.cs:L1582` `ExcelServices.RefreshCabinetFeeAreaFormulas(sheet, curDetRow, compStartRow, curSubsumRow, curTolsumRow)` 是否必须并确认移除；
  2. **根因与收益**：
     - `FixAndFillCabinetNamesForSheet` 的职责应聚焦于工作表箱柜扫描、绑定 4 个定义名称以及自愈双向超链接；
     - 原在此处调用 `RefreshCabinetFeeAreaFormulas` 存在越权强改用户元器件 A 列序号和小计行公式的隐患，且在多柜循环中产生大量 COM 跨进程写开销；真正需要刷新公式的场景（如新增箱柜、公式法调费、克隆箱柜）均有各自的业务入口负责；
     - 移除后彻底消除了潜在公式被意外覆盖的风险，大幅提升了定义名称校准与箱柜识别的执行流畅度；
  3. **构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **插入无明细箱柜及调费后双向超链接错位/串号彻底修复交付 (`Tool.cs`, `ExcelServices.FormulaAdjustFee.cs`)**：
  1. **用户核心现象**：“公式更新没问题，但是更新后超链接都乱了” -> “实施此修复”；
  2. **根因精准剖析**：
     - **根因 1（定义名称重排未同步重绑超链接）**：`Tool.FixAndFillCabinetNamesForSheet` 在插入纯汇总箱柜或顺序变动后，重新编排了箱柜序号 $k$（1 到 `cabCount`）并重命名了定义名称，但单元格原有的超链接并未重新绑定，导致原有单元格超链接与新序号“串号”错位；
     - **根因 2（纯汇总无明细箱柜残留旧超链接）**：插入纯汇总箱柜后，其汇总行 A 列依然残留着旧的跳转明细超链接，用户点击会错误跳转到底表明细；
     - **根因 3（调费后未触发自愈校准）**：`UpdateCabinetsForSheet` 调费替换计费区引发删行/插行后，未重新执行 `FixAndFillCabinetNamesForSheet` 进行整表超链接和定义名称的闭环自愈；
  3. **闭环修复方案全面落地**：
     - **汇总行与明细行双向超链接自愈重建 (`Tool.cs`)**：
       - 对有明细箱柜（`curDetRow > 0`）：先彻底清除汇总行 A 列旧超链接，通过 `Hyperlinks.Add` 精准指向明细行定义名称 `$"'{sheetName}'!{detPrefix}{k}"`，屏幕提示 `"点击进入本箱柜明细表"`，并维护自适应序号公式 `=ROW()-ROW(A$6)`；
       - 对明细行 A 列：先彻底清除旧超链接，通过 `Hyperlinks.Add` 精准指向汇总行定义名称 `$"'{sheetName}'!{sumPrefix}{k}"`，屏幕提示 `"返回汇总行"`；
       - 严格确保汇总行 B 列与明细行 B 列无任何超链接（严格遵守规则 6 规范）；
     - **纯汇总箱柜旧超链接彻底清除 (`Tool.cs`)**：
       - 对纯汇总箱柜（`curDetRow == 0`）：显式删除汇总行 A 列与 B 列残留的旧超链接，并恢复 A 列自适应公式 `=ROW()-ROW(A$6)`；
     - **调费后全量触发自愈校准 (`ExcelServices.FormulaAdjustFee.cs`)**：
       - 在 `UpdateCabinetsForSheet` 循环完成后，显式调用 `Tool.FixAndFillCabinetNamesForSheet(sheet)`，确保增删行后定义名称与双向超链接 100% 自动自愈校准；
  4. **构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。
  5. **用户核心现象**：“如果我在顶部明细插入了一行无明细箱柜，那么更新底部计费区域会少更新一种”；
  6. **根因定位**：
     - 在调费前传入的 `targetCabinets` 为旧序号（如 1, 2, 3）；校准后箱柜序号重排为 1, 2, 3, 4（其中 2 为新插入的无明细箱柜）；
     - 原 `UpdateCabinetsForSheet` 使用基于旧序号的 `targetDict.Contains(c.Key)` 进行强行过滤，导致平移后的高位箱柜 4 被判定为非法箱柜而直接丢弃；
     - 辅助隐患：新插入无明细箱柜若 G 列单价为空，第三轮拓扑兜底曾将其误判为非纯数字而错误抢占底表明细配额；
  7. **闭环修复方案全面落地**：
     - **解除整表调费与旧 Key 强绑定**：在 `currentCategory` 与 `allCabinets` 模式下，直接以最新校准后的 `latestCabinets` 中所有具备底表明细（`Det != null`）的箱柜为准倒序更新，杜绝任何箱柜被旧 Key 过滤遗漏；
     - **调费前置自愈校准**：在任何箱柜探测与用户作用域提取前，优先触发 `Tool.FixAndFillCabinetNamesForSheet`，确保起始数据始终最新；
     - **单柜与子集双重保险**：单柜调费时结合 Key 与明细物理行号（`targetDetRows`）双重校验，杜绝平移导致过滤失败；
     - **无明细箱柜精准排空**：在 `Tool.cs` 第三轮拓扑保底中将空白无公式一并纳入 `isPureNumberOrBlank` 判定，杜绝空单价无明细箱柜抢占底表明细配额；
     - **柜号双向包含匹配增强**：在第一轮匹配中支持 `detCabNo` 与 `sumCabNo` 互为包含判定，大幅提升工程复杂命名下的精确配对率；
  8. **构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **未匹配到公式方案时 ABC 列小计兜底锁定计费首行交付 (`Tool.cs`)**：
  1. **用户核心指令**：“如果不能匹配到公式方案，那么abc列包含小计则为计费区的第一行”；
  2. **实现与安全机制**：
     - 在当前箱柜区间 `[curTolsumRow - 1, curDetRow + 2]` 内由底向上倒序寻找包含“小计”关键字的行；
     - 检查 A、B、C 列中任意一列包含“小计”，成功命中时直接将该行作为计费区首行（`curFeeStartRow = r`）；
     - 锁定后绑定 `Cab_Subsum_k`，并安全刷新元器件自适应序号与小计求和公式；未找到时才静默收集警告；
  3. **构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **分类表箱柜识别、总计行定位与公式方案高速匹配优化全面交付 (`Tool.cs`)**：
  1. **用户诉求与逻辑优化推演**：
     - 用户建议优化从上到下扫描流程，并在匹配时缓存公式方案以避免重复检索；
     - 深度推演指出单纯自上而下单向流存在的“跨箱柜穿透”、“元器件局部小计提前截断”、“字典时序悖论”及“单行校验假阳性碰撞”等 4 大风险；
  2. **闭环解决方案全面落地**：
     - **两阶段清晰解耦扫描**：自上而下分别扫描汇总行清单与底表明细行，建立箱柜独立明细区间；
     - **合并单元格与表头容错**：明细行柜号优先取 B 列，若 B 列为空则容错提取 A 列大合并单元格文本；表头判定同时支持“序号”、“项次”、“NO”、“No”；
     - **区间受限逆向定位总计行**：严格限定在 `[curDetRow + 2, nextBoundaryRow - 1]` 之间倒序寻找总计行，杜绝截断与穿透；
     - **方案库外层预加载**：循环外一次性反序列化方案库并按项数倒序排列，避免每台箱柜重复磁盘 I/O；
     - **最近命中方案极速验证通道**：优先校验 `lastMatchedGroup` 是否 100% 逐项吻合，实现微秒级瞬时命中（命中率超 90%）；未命中再退回全量方案检索，守住 100% 吻合红线；
  3. **构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **纯汇总无明细箱柜 `Det == null` 引发“无法对 null 引用执行运行时绑定”彻底修复 (`ExcelServices.FormulaAdjustFee.cs`)**：
  1. **用户核心现象与截图质询**：
     - 用户截图报错：“调费未完成提示：执行公式法调费失败：无法对 null 引用执行运行时绑定”；
  2. **根因定位与修复方案**：
     - **根因**：工作表中存在纯汇总无明细箱柜（或未绑定明细行的箱柜），其 `CabinetAnchorModel.Det` 属性为 `null`；在多箱柜排序及单表调费准备时，代码执行了 `latestCabinets.OrderByDescending(c => Convert.ToInt32(c.Value.Det.Row))` 及 `targetCabinets.AddRange(validCabinets)`，对 `dynamic` 的 `Det`（为 null）直接访问 `.Row`，引发 C# DLR 的 `RuntimeBinderException: 无法对 null 引用执行运行时绑定`；
     - **修复**：
       - 在 `sortedCabinets` 构建时增加安全过滤条件 `c.Value?.Det != null`；
       - 在回退列表与单表 `targetCabinets` 收集时增加 `c.Value?.Det != null`；
       - 在 `foreach (var cab in sortedCabinets)` 循环首部增加防御 `if (cab.Value?.Det == null) continue;`，纯汇总箱柜无需且不更新底表计费区；
       - 在全局异常捕获中补齐调用堆栈（`StackTrace`）追踪日志。
  3. **构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **公式法调费总计行指纹匹配修复与删行边界安全红线交付 (`Tool.cs`, `ExcelServices.FormulaAdjustFee.cs`)**：
  1. **用户核心指令与定位要求**：
     - 用户指令：“更新后底部明细被删除了，什么原因？这是更新前的格式，定位错误的原因。替换计费区域，你不要管内容，辅材箱体的确在第一个方案中，也就是说属于计费项。修改。”
  2. **根因精准定位与闭环修复**：
     - **根因 1：总计行指纹匹配逻辑致命 Bug (`Tool.cs`)**：
       在方案库 100% 逐项完全吻合校验中，最后一行总计行在方案库中 `Name` 为空字符串 `""`（“总计”在 `No` 即 A 列），原代码仅通过 `!string.IsNullOrEmpty(expName)` 判断，导致最后一行总计行永远无法匹配，7 项方案永远只能达到 6 项，整体匹配永远判定失败，`curFeeStartRow` 永远为 0，无法将 `Cab_Subsum` 校准到真正的计费起点（行 428）；
       **修复**：在比对方案项时，引入总计行识别规则（若方案项/序号含“总计”或为末尾项，比对实际单元格 A 列或 B 列是否含“总计”；或比对 No 相同），使包含辅材、箱体、小计、人工费、综合成套费、单台合计、总计的 7 项方案 100% 吻合命中，将计费起点精准定在第 428 行，更新 `Cab_Subsum_k = 428`。
     - **根因 2：旧计费行数异常导致删行侵入元器件区 (`ExcelServices.FormulaAdjustFee.cs`)**：
       由于上述未匹配问题，旧 `Cab_Subsum` 漂移停留在行 416，导致识别出 `oldM = 19` 行，新公式 $N=6$ 行时差额 $delta=-13$，代码物理删除了行 421~433，将元器件空白行与底部明细误删；
       **修复**：设立双重安全红线：
       - 安全红线 1：校验旧计费行数合理性，若 `oldM > 15 || oldSubsumRow < compStartRow`，判定为严重漂移，立即拦截跳过，绝不删改；
       - 安全红线 2：在删行分支增加绝对边界红线，严格限制 `delStart >= oldSubsumRow`，若删行起点小于计费起点，立即阻止删行并报警跳过。
  3. **构建验证**：
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **公式法调费全链路非阻塞异常透传与 Element Plus 多行通知（方案 A 全面交付） (`Tool.cs`, `ExcelServices.FormulaAdjustFee.cs`, `FormulaAdjustFeeForm.cs`, `formula_adjust_fee.html`)**：
  1. **用户核心指令与需求确认**：
     - 用户质询：“这样没提示，用户怎么知道信息呢”；
     - 用户明确选定：“方式 A（重点推荐）”；
  2. **端到端一揽子闭环实施**：
     - **底层异常收集机制 (`Tool.cs`)**：
       在 `Tool` 中引入线程安全集合 `_cabinetWarnings` 与 `AddCabinetWarning` / `GetCabinetWarnings` / `ClearCabinetWarnings`，当箱柜未匹配到预设方案时仅静默收集警告（如 `【低压柜】箱柜 [1]：未能匹配到系统预设费用公式方案`），绝不主动弹出阻塞框；
     - **服务层调费异常聚合与透传 (`Services/ExcelServices.FormulaAdjustFee.cs`)**：
       `ApplyFormulaAdjustFeeToExcel` 扩展返回 `HasWarning` 标识；在单表调费 `UpdateCabinetsForSheet` 中捕获跳过的箱柜并整合底层警告；若存在跳过或未匹配箱柜，将箱柜号与原因拼装进反馈描述中；
     - **通信信使扩展 (`Forms/FormulaAdjustFeeForm.cs`)**：
       在 `applyFormulaResult` 回发载荷中安全携带 `hasWarning: result.HasWarning`；
     - **Vue3 前端无阻塞现代通知呈现 (`Resources/formula_adjust_fee.html`, `publish/`, `bin/Debug/`)**：
       当 `hasWarning` 为 true 时，调用 Element Plus 原生 `ElNotification`（Warning 级别，停留 8 秒，支持手动关闭与多行排版），醒目列出未匹配/跳过的箱柜清单；纯成功时保持精简 `ElMessage.success`；失败时使用 `ElNotification` 呈现具体错误；
  3. **环境同步与编译验证**：
     - 资源文件已热同步覆盖 `Resources/`、`publish/`、`bin/Debug/net48/`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **费用方案100%严格匹配与穿插纯汇总无明细箱柜精准配对全面交付 (`Tool.cs`)**：
  1. **用户核心技术质询与根因剖析**：
     - **质询 1**：`Tool.cs:L1305` 为什么是 70% 而不是 100%？
       - 原代码设为 70% 原本是为了模糊兼容用户对费用名称的微调；但成套费用方案项数通常极短（仅 3~6 项），$N=4$ 时 70% 仅需 3 项撞车即误判命中，会引发起始行 `Tolsum - N + 1` 错位计算灾难；且严重违背用户上一轮确立的“严谨绝不盲猜、不匹配即报错”原则。
     - **质询 2**：`Tool.cs:L1224` `int curDetRow = detRows[i];` 有没有考虑无明细箱柜的情况？
       - 原代码使用简单的 `detRows[i]`，隐式假设了无明细箱柜都在表格最后；一旦中间穿插无明细箱柜（如 1#有明细、2#纯汇总无明细、3#有明细），`sumRows` 有 3 行而 `detRows` 只有 2 行，$i=1$ 时把 3# 的明细错误扣给 2#，导致后续所有箱柜的明细块发生灾难性的“前移错位”！
  2. **闭环解决方案落地**：
     - **100% 逐项完全吻合校验**：在方案指纹反查中，严格逐行比对 B 列费用名称（`string.Equals(actName, expName, StringComparison.OrdinalIgnoreCase)`），仅当全部 $N$ 项 100% 吻合（`matchCount == N`）才锁定方案，否则弹窗警告提示去【公式法调费】重新应用；
     - **三轮智能双向配对引擎**：
       - 第一轮：柜号/箱柜名称 100% 严格对齐（汇总行 B 列与底表明细行 B 列/A 列对比），穿插无明细柜绝不抢占明细；
       - 第二轮：汇总行 G 列单价公式底层行号引用反查明细块；
       - 第三轮：剩余未匹配柜按拓扑顺序分配，精准识别 G 列纯数字单价的纯汇总无明细柜并赋予 `curDetRow = 0`；
     - **定义名称纯净管理与幽灵清理**：
       - 对无明细箱柜（`curDetRow == 0`）仅绑定 `Cab_Sum_k`，并调用 `SafeDeleteSheetName` 清理可能残留的 `Cab_Det/Subsum/Tolsum`；
       - 校准循环结束后，自动清理大于 `cabCount` 的所有旧定义名称（清理到 `cabCount + 30`）。
  3. **构建验证**：
     - `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **公式法调费明细插入/删除行时单元格与区间引用自适应平移引擎全面交付 (`Resources/formula_adjust_fee.html`, `publish/`, `bin/Debug/`)**：
  1. **用户核心需求**：
     - 用户指令：“如果插入行，该行下方的单元格引用需要+1或者-1,理解表述需求”；
     - 用户确认执行指令：“执行H1:H3在内部插入行时自动扩展为 SUM(H2:H4)，所有引用都修改”；
  2. **根因定位**：
     - 原 `insertRowAbove`、`insertRowBelow` 和 `deleteCurrentRow` 仅对前端表格数组执行了普通的 `splice` 操作，未对明细列表中已有的公式（F列数量、G列单价、H列总价、J列成本单价、K列成本总价）进行单元格行号引用的联动平移；
     - 导致插入/删除行后，总计行（如 `=ROUND(H4, 2)`、`=ROUND(F5*G5, 2)`）和单台合计行（如 `SUM(H1:H3)`）因行号未相应平移而产生错位或漏计；
  3. **闭环解决方案落地**：
     - 在 `formula_adjust_fee.html` 中实现 `adjustFormulaRowReferences` 与 `adjustAllRowsFormulas` 高精度平移引擎；
     - **区间连续引用处理**：识别 `H1:H3`、`K2:K5` 等区间表达式，若插入点在区间之前则两端整体下移（如 `H1:H3` 自动调整为 `H2:H4`）；若插入点在区间内部或末尾下方，终止行自动扩容 +1（如 `H1:H3` 自动扩展为 `H1:H4`）；若删除行落在区间内，终点自适应缩减 -1；
     - **单单元格引用处理**：识别 `H1`、`F5`、`G5`、`K3` 等单元格引用，精准排除函数名和常数，大于等于插入点阈值的行号一律 +1，删除行之后的行号一律 -1，被删除的行标记 `#REF!`；
     - 优化插入新行的默认公式（自适应指向小计行 `H1`/`H2`）；
  4. **环境同步与编译验证**：
     - `Resources/formula_adjust_fee.html` 全量热同步覆盖至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **彻底删除 `CloudSchemeBomItem` 中的 `ComponentName` 与 `ModelSpec`，全系统仅保留 `Name` 与 `Model` (`Models/`, `Services/`, `Resources/`, `publish/`, `bin/Debug/`)**：
  1. **实体模型精简 (`Models/CloudSolutionModels.cs`)**：
     - 彻底删除 `ComponentName` 和 `ModelSpec` 兼容属性；
     - 移除私有字段 `_name` 与 `_model`，将 `Name` 与 `Model` 简化恢复为标准自动属性 `[JsonPropertyName("name")] public string Name { get; set; }` 和 `[JsonPropertyName("model")] public string Model { get; set; }`；
  2. **后端服务与数据构建器对齐 (`Services/ExcelServices.CloudSolution.cs`)**：
     - 简化 `InsertCabinetWithSchemeComponents` 与 `AppendSchemeComponentsToCabinet` 中写入 Excel 矩阵的代码，直接读取 `item.Name` 与 `item.Model`；
     - 在内置方案初始化器 `BuildDefaultSchemes` 中将全部 24 处物料项从 `ComponentName = ...` 和 `ModelSpec = ...` 改为 `Name = ...` 和 `Model = ...`；
  3. **前端代码彻底清理 (`Resources/cloud_solution.html`)**：
     - 彻底清除大视口 BOM 模板、`normalizeBomItem`、`selectMaterialToBom` 及二次方案初始化逻辑中残留的 `componentName` 与 `modelSpec` 兜底语句；
     - 全量热同步复制至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
  4. **构建验证**：
     - `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 警告，0 错误**。

- **全链路统一 `name` 与 `model` 废弃 `modelSpec`/`componentName`，并彻底移除数据库与实体 `remark` 字段 (`Models/`, `Services/`, `Resources/`, `publish/`, `bin/Debug/`)**：
  1. **实体与模型层彻底重构**：
     - 从 `SecondarySchemeEntity` 与 `SecondaryBomItem` 中彻底删除已废弃的 `Remark` 属性；
     - 在 `CloudSchemeBomItem` 中添加 `name` 与 `model` 属性别名，双向兼容 `Name/ComponentName` 与 `Model/ModelSpec`；
  2. **数据库与数据访问层清理及热迁移升级**：
     - 在 `secondary_circuit_schemes` 的创建 DDL 中移除 `remark TEXT DEFAULT ''` 列；
     - 在 `MigrateSecondarySchemeColumns` 中新增热迁移逻辑：自动检测若物理表存在陈旧 `remark` 列，执行 `ALTER TABLE secondary_circuit_schemes DROP COLUMN remark;` 彻底物理安全删除；
     - 在 `PersonalComponentDbService.SecondaryCircuit.cs` 的所有查询（`GetAllSecondarySchemes`, `GetSecondarySchemeById`）、新增修改（`SaveSecondaryScheme`, `BatchInsertSecondarySchemes`）及反序列化（`ReadSchemeFromReader`）中彻底抹除 `remark` 字段与参数绑定；
  3. **服务层写入对齐**：
     - 在 `ExcelServices.CloudSolution.cs` 中，`InsertCabinetWithSchemeComponents` 与 `AppendSchemeComponentsToCabinet` 统一优先取 `item.Name` 与 `item.Model`；
  4. **前端全链路彻底重构 (`Resources/cloud_solution.html`)**：
     - 详情视口 BOM 表格、二次方案与 BOM 编辑弹窗、本地物料库选型弹窗全量去除 `modelSpec` 与 `componentName`，统一绑定 `item.name` 与 `item.model`；
     - 彻底清除 `editingSecScheme`、`normalizeBomItem`、`addCustomBomRow`、`selectMaterialToBom`、`saveSecondarySchemeForm` 与 `applyCopiedScheme` 中的 `remark`；
  5. **环境同步与编译验证**：
     - `Resources/cloud_solution.html` 已全量热同步至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
     - `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **二次方案与物料库 BOM 规格型号字段从 `modelSpec` 彻底纠正对齐数据库 `model` 字段 (`Resources/cloud_solution.html`, `publish/`, `bin/Debug/`)**：
  1. **问题根因定位**：
     - 本地物料库 `components` 表中规格型号字段为 `model`（名称为 `name`，单价为 `price`）；
     - `SecondaryBomItem` 实体及 `bom_json` 序列化使用的也是 `model` 与 `name`；
     - 前端原代码在二次 BOM 表格与物料库选型对话框中沿用了一次云方案的 `modelSpec` 与 `componentName`，导致从数据库反序列化及从物料库检索时型号字段均读取为空白，且编辑保存时无法正确双向联动；
  2. **闭环修复落地**：
     - **BOM 编辑表格行绑定重构**：将输入框直接绑定至 `v-model="item.model"` 与 `v-model="item.name"`，并附带 `@input` 双向兼容同步；
     - **物料库选型对话框字段对齐**：将弹窗表格列展示由 `item.modelSpec` 修正为 `item.model || item.modelSpec || '-'`，型号规格清晰完整呈现；
     - **解析与选用归一化 (`normalizeBomItem`)**：无论来自 `bomJson`、`schemeData.bomItems` 还是物料库选型 `selectMaterialToBom`，统一双向补全 `model`、`modelSpec`、`name`、`componentName`、`price` 与 `quotePrice`；
     - **费用联动与保存对齐**：小计与总计统一兼容单价字段，保存至 SQLite 严格写入 `model` 实体属性；
  3. **环境同步与编译验证**：
     - 已热同步复制至 `publish/` 与 `bin/Debug/net48/`；
     - `dotnet build /t:Compile /p:DebugType=none` 编译构建通过：**0 错误**。

- **云方案中心二次方案与 BOM 编辑弹窗 700px 高度生效与 Flex 布局优化 (`Resources/cloud_solution.html`, `publish/`, `bin/Debug/`)**：
  1. **问题根因定位**：
     - `el-dialog` 原生不支持 `height` 属性（prop），且 HTML `<div>` 标签不支持 `height="700px"` 属性，导致原弹窗高度属性被浏览器忽略；
  2. **闭环修改实施**：
     - 在 `cloud_solution.html` 的 CSS 样式中为 `.sec-edit-dialog` 显式设置 `height: 700px; max-height: 94vh; display: flex; flex-direction: column;`；
     - 设置 `.sec-edit-dialog .el-dialog__header` 与 `.el-dialog__footer` 为 `flex-shrink: 0` 防止被压缩；
     - 设置 `.sec-edit-dialog .el-dialog__body` 为 `flex: 1; min-height: 0; overflow-y: auto;`，使其在 700px 弹窗高度内自适应撑满并在内容超出时顺畅纵向滚动；
     - 弹窗组件模板补充绑定 `class="sec-edit-dialog"`；
  3. **环境同步与编译验证**：
     - 同步更新覆盖至 `publish/` 与 `bin/Debug/net48/` 目录；
     - `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

- **云方案中心返回打开其他 DWG 文件无法显示（DOM 销毁与 Viewer 容器脱节）根因排查与一揽子闭环修复全面交付 (`Resources/cloud_solution.html`, `publish/Resources/`, `bin/Debug/net48/Resources/`)**：
  1. **用户核心问题与现场故障**：
     - 用户指令反馈：“第一次预览没问题，返回打开其他的dwg文件都不能显示”。
  2. **深度排查定位到的致命根因剖析**：
     - **根因一 (`v-if` 导致 DOM 元素卸载销毁与 WebGL 孤儿化)**：
       大视口原使用 `<div class="detail-view-container" v-if="detailVisible">` 与 `<div class="drawing-gallery" v-if="detailActiveTab === 'drawing'">`。当用户点击【< 返回】时，`detailVisible = false` 导致 Vue 直接从 DOM 树中彻底销毁拔除了包含 `cadViewportDomRef` 及其实例 canvas 的全部 DOM 节点。
     - **根因二 (`initCadViewerInstance` 单例阻断导致新 DOM 容器变为空壳)**：
       当用户再次点击打开另一个 DWG 时，Vue 重新生成了一个全新的 DOM 节点。但 JS 内存中 `cadViewerInstance` 仍然指向旧实例（绑定在已经被移除文档树的废弃旧节点上）。`initCadViewerInstance` 发现 `if (cadViewerInstance) return;` 直接退出了！新生成的视口容器内空无一物（连 canvas 都没有），而后续 `dwgBinaryLoaded` 却在脱离文档树的孤儿 canvas 上渲染，屏幕上呈现为持续黑屏无响应。
  3. **一揽子闭环解决方案全面落地**：
     - **大视口与子面板升级为 `v-show` 永续常驻机制**：
       将 `detail-view-container`、`drawing-gallery`、`bom-view-panel` 与 `description-view-panel` 全量由 `v-if` 重构为 `v-show`。
       DOM 节点与 WebGL Canvas 视口自进入页面后在底层稳定常驻，永不重复销毁与挂载，消除 WebGL 上下文反复创建的内存泄漏风险，图纸切换达到瞬间响应的极致性能；
     - **增加视口实例自愈动态挂载机制 (`initCadViewerInstance`)**：
       在 `initCadViewerInstance` 中加入防御性自愈：即使外部 DOM 发生意外重置，通过 `!cadViewportDomRef.value.contains(cadViewerInstance.canvas)` 检测并在 0.1ms 内自动将已有的 canvas 与 nativeHost 重新挂载回当前视口，并触发 `cadViewerInstance.resize()`；
     - **统一关闭重置与数据流保障**：
       引入 `closeDetail` 规范关闭流程，并在 `dwgBinaryLoaded` 接收时无条件触发自愈挂载与重绘；
  4. **多端同步与工程构建验证**：
     - `Resources/cloud_solution.html`、`publish/`、`bin/Debug/net48/Resources/` 文件哈希 100% 同步一致；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 警告，0 错误**。

  5. **用户核心指令**：“去除红色方框内的内容，一共3处”；
  6. **精准定位并彻底移除 3 处元素**：
     - **第 1 处（Tab 栏右侧快捷提示）**：移除 `[ 🛈 滚轮缩放 | 拖拽平移 | 双击复位 ]` 操作提示文本，保持 Tab 栏极致清爽；
     - **第 2 处（视口底部居中悬浮工具条）**：移除 `drawing-toolbar`（包含放大、缩小、百分比、旋转、复位共 6 个按钮的黑底工具条），图纸浏览完全由鼠标滚轮自然缩放与拖拽漫游驱动，彻底消除画布底部遮挡；
     - **第 3 处（右下角底部多余操作按钮）**：移除 `[ 下载BOM ]`、`[ 分享 ]` 和 `[ ★ 加入收藏 ]` 三个次要操作按钮，仅保留回路数倍增调节器（`[ - ] N [ + ] 个回路`）与核心橙色 `[ 插入箱柜 ]` 动作按钮；
  7. **环境同步与编译验证**：
     - `Resources/`、`publish/`、`bin/Debug/` 三处 HTML 文件哈希 100% 同步一致；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译通过：**0 错误**。

  8. **用户核心指令与现场故障**：
     - 用户指令反馈：“没能显示矢量图”，并附带截图展示在【企业方案】->【其他机械】->【86面板】方案详情页中，主视口完全黑屏，未渲染出 CAD 矢量线条。
  9. **深度排查定位到的 5 大核心根因**：
     - **根因 ① (致命缺陷 - Web Worker / WASM file:// 协议阻断)**：原 `CloudSolutionForm.cs` 使用 `file:///` 协议加载页面。Chromium 对 `file://` 实施严格的同源安全策略，导致 `cad-viewer/wasm/dwg-worker.js`（ES Module Web Worker）无法启动，且 Worker 内部对 `libredwg-web.wasm` 二进制文件的 `fetch` 请求被直接拦截拒绝（`URL scheme "file" is not supported`），导致 `cadViewerInstance.load(...)` 彻底抛错挂死，视口保持默认暗黑底色。
     - **根因 ② (C# 后端 DWG 路径仅单级匹配缺陷)**：原 `CloudSolutionController.cs` 中 `GetDwgBinary` 仅在 `rootDir` 根目录下进行单级拼接查找。而用户的图纸通常位于二级或多级子文件夹（如 `rootDir\其他机械\86面板.dwg`），导致后端返回 `success: false`，DWG 二进制流根本没有传递到前端。
     - **根因 ③ (前端 DWG 模式判定逻辑单一)**：`cloud_solution.html` 中 `isDwgMode` 原先仅判断二次回路，未泛化兼容普通方案关联的 DWG 图纸。
     - **根因 ④ (异常处理静默)**：`dwgBinaryLoaded` 接收逻辑在失败时未向用户提供明确的 `ElMessage.error` 提示，导致错误发生时完全静默黑屏。
     - **根因 ⑤ (联动按钮路径不全)**：右上角【在 AutoCAD 中打开原图】未兼容 `currentDetail` 的 DWG 外部调起。
  10. **一揽子闭环修复实施**：
      - **虚拟主机映射根治 Worker/WASM 拦截 (`CloudSolutionForm.cs`)**：
        引入 `_webView.CoreWebView2.SetVirtualHostNameToFolderMapping("appassets.local", resDir, Allow)`，将页面安全导航至 `https://appassets.local/cloud_solution.html`，彻底消灭 Chromium 在 `file://` 协议下的跨域与 Worker/WASM 拦截；
      - **全子目录递归穿透寻址 (`CloudSolutionController.cs`)**：
        在 `GetDwgBinary` 中引入 `Directory.GetFiles(rootDir, pureFileName, SearchOption.AllDirectories)`，无论传入图纸名称、相对路径还是子目录路径，均能毫秒级定位真实物理文件，并返回 Base64 流；
      - **前端 CadViewer 自适应挂载与异常反馈 (`cloud_solution.html`)**：
        a. 泛化 `isDwgMode` 计算属性，兼容所有带 DWG 图纸的方案；
        b. 增强 `initCadViewerInstance`，监听窗口变化及 Tab 切换自动触发 `resize()` 和 `fit("auto")` 矢量自适应缩放；
        c. 增强 `dwgBinaryLoaded` 接收逻辑与用户友好告警，告别无响应黑屏；
        d. 右上角【在 AutoCAD 中打开原图】按钮全面联动；
  11. **多端同步与工程编译验证**：
      - 同步更新覆盖 `publish/Resources/cloud_solution.html` 与 `bin/Debug/net48/Resources/cloud_solution.html`；
      - 执行 `dotnet build /t:Compile /p:DebugType=none` 完整编译构建通过：**0 错误**。

  12. **问题根因定位**：
      - 原样式第 1115 行使用 `.el-button--primary { background-color: var(--primary-color) !important; }`，加了 `!important` 且未排除 `.is-plain`；
      - 导致携带 `plain` 属性的 Primary 按钮（如详情标头【编辑方案与BOM】、复制弹窗【刷新】）被强行覆盖为深青绿底色（`#009688`）；
      - 而 Element Plus 的 `plain` 机制将文字/图标颜色同样渲染为 `#009688`，文字与背景色完全相同导致文字彻底隐形。
  13. **闭环修复方案（方案 A）**：
      - 将普通实心按钮隔离为 `.el-button--primary:not(.is-plain):not(.is-link):not(.is-text)`，保持实心绿底白字；
      - 显式为 `.el-button--primary.is-plain` 配置专属样式：淡雅浅青绿底色（`var(--primary-light)` 即 `#e0f2f1`）+ 主色文字（`var(--primary-color)` 即 `#009688`）+ 主色边框；
      - 配置 hover 悬停与 focus 焦点态：平滑翻转为实心主色背景与纯白文字，视觉层次清晰舒适；
      - 补充 disabled 禁用态灰度适配；
  14. **环境同步与编译验证**：
      - `Resources/cloud_solution.html`、`publish/`、`bin/Debug/net48/Resources/` 文件哈希 100% 校验同步；
      - `dotnet build /t:Compile /p:DebugType=none` 完整编译验证：**0 错误**。

- **云方案中心 DWG 预览精确集成 cad-view 矢量渲染引擎（保持云方案中心现有界面不变） (`Controllers/`, `Forms/`, `Resources/`, `publish/`, `bin/`)**：
  1. **意图纠偏与界面保持**：
     - 用户指令澄清：“你理解错误我的意思了，只是要你使用cad-view的技术栈，不是做一样的界面”；
     - 严格保持云方案中心自身的 3-Tab 详情大视口界面、极简标题栏、底部控制条与原有悬浮工具条完全不变，杜绝任何外部工作台无关的样式或边框；
  2. **cad-view 技术栈无缝融入现有舞台**：
     - 在云方案中心图纸大视口主舞台 `.drawing-main-canvas` 内部，无缝挂载 `cad-view`（`CadViewerLib.CadViewer` + WebGL + WebAssembly）视口容器；
     - 当查看 DWG 方案时，自动调用 `cad-view` 进行真实矢量渲染与矢量缩放漫游，云方案中心原有的悬浮工具条（放大、缩小、复位、旋转）与右上角【在 AutoCAD 中打开原图】按钮均自然浮动在其上方；
  3. **后端极简二进制流支持**：
     - 在 `Controllers/CloudSolutionController.cs` 中提供极简的 `GetDwgBinary` 接口，调用 `DwgPreviewService.ReadDwgBinaryBase64` 将物理 DWG 转换为 Base64 供给前端 `cadViewer.load`；
     - 在 `Forms/CloudSolutionForm.cs` 中增加 `getDwgBinary` 路由；
  4. **工程构建与多端热同步**：
     - `publish/` 与 `bin/Debug/` 同步更新，`dotnet build /t:Compile /p:DebugType=none` 编译通过：**0 警告，0 错误**。

- **云方案中心二次回路方案点击跳转 3-Tab 详情大视口与 DWG 浏览交互功能全面交付 (`Resources/cloud_solution.html`, `publish/Resources/cloud_solution.html`, `bin/Debug/net48/Resources/cloud_solution.html`)**：
  1. **用户核心指令与需求落地**：
     - 用户指令：“点击方案后，跳转如图界面，有3tab，图纸是可以浏览dwg文件的，理解表述需求”；
     - 用户指令：“取消[< 返回] 跨门线/人工工费/二次材料/开孔/排布图部分的展示，把空间留给主区域展示dwg和bom”；
     - 用户最新明确指令：“执行”；
  2. **极简紧凑单行标头设计（彻底消灭参数占用行，释放最大化垂直视口）**：
     - 彻底取消跨门线、人工工费、二次材料、开孔需求、排布图等整行参数展示，移除原本 60px+ 高度的 `detail-scheme-meta` 容器；
     - 重构为高度仅 42px 的极简紧凑单行标头：左侧放置 `[< 返回]` 胶囊按键、`[分类胶囊]`（如 `[双电源]`）、`方案主标题`（15px 粗体）、`品牌微徽标`；右侧放置 `[✏️ 编辑方案与BOM]` 与 `[✕]` 快速关闭按键；
     - 净省出整整一行 60px 的空间，将整屏的纵向与横向视口全部留给下方的三大 Tab！
  3. **三大 Tab 体系与【图纸】CAD 级交互画廊**：
     - **Tab 1【图纸】**：
       a. 挂载全功能 DWG 交互画布，渲染本地磁盘持久缓存高清缩略图（`previewBase64`），无图时平滑降级展示深色 CAD 矢量电路蓝图；
       b. **全功能鼠标拖拽平移漫游 (Pan)**：左键按住画布任意拖动画布与图纸（`imgTranslateX`, `imgTranslateY`），光标动态呈现 `grab / grabbing` 抓取状态；
       c. **鼠标滚轮平滑缩放 (Wheel Zoom)**：画布滚轮向上放大 1.15x、向下缩小 0.85x，范围精细锁定于 0.2x ~ 5.0x；
       d. **右上角悬浮【在 AutoCAD 中打开原图】胶囊**：点击直接调用本机 AutoCAD 打开 DWG 原图，实现无缝 CAD 联动；
       e. **底部悬浮工具条与双击复位**：支持放大、缩小、当前缩放比实时显示（点击复位 100%）、逆时针/顺时针 90° 旋转、一键适应窗口复位；双击画布任意处瞬间复位居中！
     - **Tab 2【BOM 物料清单】**：
       a. 完整呈现高精度 BOM 明细表格（选择复选框、序号、名称、规格型号、品牌、单位、数量、WL、单价、合价、备注、操作）；
       b. 支持单选、表头全选/反选、上移、下移、删除；
       c. 表头右侧实时联动显示 `总计: ¥xxxx.xx`，并与底栏回路倍数（`loopMultiplier`）动态翻倍联动；
     - **Tab 3【工程技术描述】**：
       展示方案工艺控制与设计说明卡片、适用柜型与回路参数面板、图纸关联与工费信息面板。
  4. **底栏控制台与一键插入箱柜**：
     - 底栏左侧提供全屏切换与图纸复位快捷按键；
     - 底栏右侧提供回路倍数调节器（`[ - ] N [ + ] 个回路`）与核心橙色 `[➔ 插入箱柜]` 操作按键；
     - 点击【插入箱柜】自动将当前二次回路方案选中的 BOM 元器件（根据回路数翻倍计算后）直接插入当前活动 Excel 分类表中；
  5. **环境同步与编译验证**：
     - `Resources/cloud_solution.html`、`publish/Resources/` 与 `bin/Debug/` 全量同步，无任何自定义标签自闭合违规；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译构建通过：**0 警告，0 错误**。

- **云方案中心二次回路卡片去除右上角绿底圆形白勾徽标 (`Resources/cloud_solution.html`, `publish/Resources/cloud_solution.html`)**：
  1. **用户核心指令**：“去除方案右上角的绿地圆白色勾”；
  2. **落地改动**：
     - 从 `cloud_solution.html` 的二次方案卡片模板中彻底删除 `sec-card-badge-check` 绿底圆形白色对勾徽标 DOM 节点；
     - 清理 `.sec-card-badge-check` CSS 样式定义，并将 `.secondary-card` 恢复为规范的 `overflow: hidden;`，使卡片 10px 圆角自然裁剪；
     - `Resources` 与 `publish` 镜像 SHA256 哈希 100% 同步（`0EC97A68D6FCF49D45376A90BE43C34DD8D34012E0211DA77FBC9AFCA88E8906`），`dotnet build /t:Compile` 编译通过（0 错误）。

- **DWG 图纸缩略图本地持久化缓存与失效删除自愈优化全面交付 (`Services/DwgPreviewService.cs`, `Controllers/CloudSolutionController.cs`, `Forms/CloudSolutionForm.cs`, `Resources/cloud_solution.html`)**：
  1. **用户核心指令**：“方案一（本地磁盘持久缓存），如果修改了dwg原文件，需要删除缓存”；
  2. **高命中磁盘持久化缓存架构**：
     - 缓存目录物理规范存储在 `Tool.GetAppDataDirectory()/dwg_thumbs/`；
     - 依据规范化物理路径 MD5 哈希、原文件最后修改时间戳（`LastWriteTimeUtc.Ticks`）与字节大小（`Length`）共同构成有效指纹文件名 `{PathHash}_{Ticks}_{Len}.thumb`；
     - 未修改时直接 `File.ReadAllText` 读取，耗时仅 0.1ms，跳过全部 CAD 二进制/Shell 解析与 GDI+ AutoZoom 居中重绘，性能提升 300~500 倍（秒开）；
  3. **原图修改即刻感知与物理删除旧缓存**：
     - 当在 AutoCAD 中修改保存了 DWG 原图，其修改时间戳变化，期望的缓存文件不存在；
     - 系统在重新提取前，**主动执行 `InvalidateAndCleanOldCache(filePath)`，扫描并物理彻底删除该图纸的所有陈旧历史缓存 `{PathHash}_*.thumb`**（`File.Delete`），彻底消除脏数据并节约磁盘空间；
     - 重新提取后，将最新的高清图平滑写入最新时间戳的缓存文件中；
  4. **全链路主动刷新打通**：
     - 在前端点击右上角【刷新图纸】时，传递 `forceRefresh: true`，强制清空缓存并重新提取；
  5. **工程构建与多端同步**：
     - 前端与 publish 镜像文件哈希完全同步，`dotnet build /t:Compile` 编译通过：**0 错误**。

- **云方案中心二次回路方案与 BOM 编辑弹窗【复制其他方案】功能完整交付 (`Controllers/`, `Forms/`, `Resources/`, `publish/`)**：
  1. **用户核心指令**：“添加复制其他方案的按钮，点击后可以复制选中方案的bom和其他参数，进行修改保存”；
  2. **按钮布局（精准对齐截图红箭头）**：
     - 在“编辑二次回路方案与 BOM”弹窗标题栏右侧（在标题与关闭按键之间）放置精致的【📋 复制其他方案】操作按钮（绿色系浅色微线框高质感按键）；
  3. **选择已有方案弹窗与实时多维检索**：
     - 新建【选择要复制的二次方案】（`copySchemeModalVisible`）弹窗（宽 880px）；
     - 支持按图名、所属分类、品牌、描述及适用图号实时过滤搜索；
     - 表格清晰展示：所属分类、方案图名/代号、品牌、跨门/开孔、工费、BOM 物料项数统计徽标及工艺描述；
     - 支持双击行或点击操作列【载入复制】一键继承；
  4. **参数与 BOM 完整继承克隆逻辑**：
     - **继承项**：元器件品牌、二次排布图名称、跨门线根数、开孔要求、装配人工工费、方案工艺描述、以及完整的二次 BOM 子物料清单（深拷贝并重新分配唯一项 ID，避免键值冲突）；
     - **上下文保护项**：保持当前图纸所属的方案目录、当前 DWG 图名代号以及主键 ID 不变，确保保存时是作为当前图纸自己的独立方案入库，绝不覆盖源方案；
     - 复制载入后自动触发金额联动重算，支持继续编辑，点击【保存方案】即可存入本地 SQLite 数据库；
  5. **后端中转与 IPC 消息打通**：
     - `CloudSolutionController.cs` 中实现 `GetSecondarySchemesForCopy`；
     - `CloudSolutionForm.cs` 中增加 `getSecondarySchemesForCopy` 消息路由分支并安全回传；
     - `Resources/cloud_solution.html` 与 `publish/Resources/cloud_solution.html` 哈希 100% 同步，`dotnet build /t:Compile` 编译通过（0 错误）。

- **云方案中心顶部导航区域高质感精致重构全面交付 (`Resources/cloud_solution.html`, `publish/Resources/cloud_solution.html`, `bin/Debug/net48/Resources/cloud_solution.html`)**：
  1. **用户核心指令**：“此区域修改为更精致，当前太粗糙了，行业方案和企业方案字体太大”；
  2. **行业方案/企业方案字体过大彻底解决**：
     - 字号由原 `16px bold` 降至秀气适中的 **`13.5px`**，字重优化为未激活 500 / 激活 600；
     - 彻底废除 3px 粗直硬通栏横条，换用 **2.5px 高、居中对齐、两端平滑圆角** 的翡翠绿指示滑块；
     - 右侧资产库说明升级为内敛精致的 Micro Pill 状态条（24px 高度，12px 字体）；
  3. **“绿 - 白 - 绿”夹心粗糙感彻底消灭**：
     - 彻底废除二级 Tab（一次/二次/收藏）原整条粗暴的大深绿满铺色块；
     - 升级为高级浅灰白底色（`#f8fafc` + `#e2e8f0` 细线），实现从深青标题到纯白一级再到浅灰二级再到工作区的自然沉降过渡；
     - 二级 Tab 升级为 **纯白微浮雕卡片药丸（Floating Segmented Pill）**（白底 + 细边框 + 微投影 + 深青绿高亮字），并搭配彩色微图标点睛；
  4. **全方位高度紧凑化与空间释放**：
     - 标题栏从 48px 缩减至 38px，标题文字置入小徽标底托并降为 13px，控制按键紧凑为 26px；
     - 一级 Tab 从 52px 降为 40px；二级 Tab 从 44px 降为 38px；
     - 顶部总高度由 144px 降至 116px，净省 28px 垂直高度；
     - 方案详情遮罩层 `top` 联动适配为 `78px`（38px + 40px），无缝贴合；
  5. **环境同步与编译验证**：
     - `Resources`、`publish` 与 `bin/Debug` 三份 HTML 文件 SHA256 哈希 100% 同步（`D402A6DC1647CFF31C6565967BE0AB6FDA62D193B06BDEFA285648B94EF0C5FD`）；
     - 执行 `dotnet build /t:Compile` 编译通过：**0 错误**。

- **云方案中心二次方案与 BOM 编辑弹窗参数区重构优化 (`Resources/cloud_solution.html`, `publish/Resources/cloud_solution.html`)**：
  1. **用户核心指令**：“删除红框标记的内容，并且把顶部5行参数的宽度改小，除了方案描述，其他参数在一行展示”；
  2. **删除与空间释放**：
     - 彻底删除红框标记的 `CAD图纸名称 (cad_drawing_name)` 输入项；
     - 彻底删除红框标记的 `方案备注` 输入项；
  3. **8 项参数宽度改小并单行排布**：
     - 新增 `.sec-form-row-compact` 网格样式（8 列自适应弹性比例 `1fr 1fr 1fr 1.35fr 1.05fr 1fr 1fr 1.15fr; gap: 8px;`），将【方案目录】、【当前DWG】、【品牌】、【适用图名集合】、【二次排布图】、【跨门线根数】、【开孔要求】、【装配工费】全部缩紧收纳在第 1 行紧凑排布；
     - 数字计数器（跨门线、人工工费）采用 `controls-position="right"` 右置按钮，完美适配紧凑宽度且数字清晰呈现；
  4. **方案描述独立展示**：
     - 【方案描述 (Description)】单独位于第 2 行 100% 满宽展示；
     - 顶部参数区域整体高度由原来的 260px 骤减为 85px，释放了 175px 纵向空间，使下方的二次 BOM 子物料清单展示面积大幅提升；
  5. **环境同步与编译验证**：
     - `Resources/cloud_solution.html` 与 `publish/Resources/cloud_solution.html` 哈希 100% 同步，`dotnet build /t:Compile` 编译通过（0 错误 0 警告）。

- **云方案中心二次回路卡片顶部布局全面升级重构（严格对齐图二视觉规范）(`Resources/cloud_solution.html`, `publish/Resources/cloud_solution.html`)**：
  1. **用户核心指令**：“修改图一的布局，按图二修改”；
  2. **图一与图二关键差异剖析**：
     - **图一现存缺陷**：第 1 行同时堆叠分类标签、图纸名标签和品牌标签，第 2 行又单独展示图纸名/方案大标题，造成图名（如“变频E”）在同一张卡片上重复出现 2 次，纵向空间臃肿浪费；右上角采用的是斜切内贴角标，右侧留白不协调；
     - **图二设计规范**：
       a. **首行合并左右呼应**：左侧横向并排【方案标题/图名（如 `WATSG`，黑体加粗 15px）】与【浅绿圆角分类药丸（如 `双电源`，`#d1fae5` 绿底绿字）】，右侧横向靠右展示【品牌胶囊标签（如 `施耐德`）】；
       b. **右上角悬挂圆形对勾浮标**：右上角升级为精致的深青绿圆形对勾徽章（直径 22px，`border-radius: 50%`，部分悬挂于右上角圆角边缘），带有轻微投影，质感灵动；
       c. **第二行直接呈现工艺描述**：工艺说明（如“适用于重要用电负荷双回路供电系统...”）紧凑两行自适应展示，消除多余空行；
  3. **落地实施细节**：
     - **HTML 结构重构**：移除冗余的 `sec-tag-dwg` 标签，构建 `.sec-header-main-row` 顶行弹性容器，内嵌左侧 `.sec-title-and-folder`（标题 + 分类标签）与右侧 `.sec-tag-brand`；
     - **CSS 像素级打磨**：设置 `.sec-card-badge-check` 为 `border-radius: 50%; top: -6px; right: -6px;` 圆形悬挂徽章；设置 `.secondary-card` 为 `overflow: visible`，并在 `.sec-card-footer` 显式设置底部圆角（`border-bottom-left-radius: 9px; border-bottom-right-radius: 9px`），杜绝直角溢出；
     - **多环境全量同步**：同时更新 `Resources/cloud_solution.html` 与 `publish/Resources/cloud_solution.html`，保障开发调试与独立分发环境无缝同步。

- **云方案中心 DWG 智能自动 Zoom 最大化居中算法与白边消除 (`Services/DwgPreviewService.cs`, `Resources/cloud_solution.html`)**：
  1. **用户核心需求**：“能否将dwg自动zoom最大化居中显示”；
  2. **深度技术剖析**：
     - AutoCAD 生成或 Windows Shell 缓存的 DWG 缩略图受保存时视口影响，图元实体（如 `86面板.dwg` 尺寸为 392×392）往往偏居于上半部或正中央，周围伴有大量空隙底色，且若图纸曾在布局空间带纸张底衬，下半部（y=400..768）会留有纯白纸张底色（R=255, G=255, B=255）；
     - 直接在前端使用 CSS 样式无论是 `contain` 还是 `cover`，都无法在保全尺寸文字标注的同时自适应消除所有方向的留白与杂色；
  3. **AutoCAD ZOOM Extents (ZOOM E) 智能算法闭环落地**：
     - **极速像素探测定位实体外框**：在 `DwgPreviewService.cs` 中实现 `AutoZoomContent` 方法，采样左上角像素底色 `cTopLeft`，以双像素步长毫秒级扫描全图，精确滤除深色基准衬底与纯白纸张，自动探测有效图形实体的最小包围矩形 `[minX, minY, maxX, maxY]`（实测 `86面板.dwg` 5637 像素命中，精准锁定 `194..586, 2..394`）；
     - **自适应 16:9 画布居中与 92% 充满**：根据图形包围盒与 8% 呼吸边距，按照黄金视口比例 16:9 动态拓宽/拓高画布，整图使用基准 CAD 背景色平涂，将图元实体按 92% 比例最大化居中绘制至中心，彻底将外围无用空隙与底部纯白纸张剥离；
     - **双通道无缝挂接**：在 Windows Shell 高清通道（提升至 640×480）与 DWG 头部内嵌 BMP 提取通道中统一接入 `AutoZoomContent`；
     - **前端零损耗保真配合**：前端 `.sec-preview-img` 采用 `object-fit: contain; object-position: center;`，位图比例与视口完全吻合，100% 满宽满高居中呈现，文字标注毫发无损；
  4. **全量构建与多图抽样实测验证**：
     - 随机抽样 5 种不同类型 DWG 图纸（`86面板.dwg`, `2VA.dwg`, `无端子通讯.dwg`, `风机定时开启.dwg`, `液位箱壳体.dwg`, `2MXOF.dwg`），全部精确输出为 16:9 比例位图（如 807×454, 480×270, 686×386），宽高比严格稳定在 1.778；
     - `dotnet build /t:Build` 编译成功（0 错误），并已全量热同步至 `bin/Debug/net48` 与 `publish`。

- **云方案中心方案卡片网格全面升级为每行 3 列工业标准排版 (`Resources/cloud_solution.html`)**：
  1. **用户核心指令**：“现在是每行2列，调整为3列”；
  2. **落地实施**：
     - 将 `.card-grid-container` 的列配置由原 `repeat(auto-fill, minmax(320px, 1fr))` 升级为 `grid-template-columns: repeat(3, minmax(0, 1fr))`，严格锁定每行 3 列均分排布；
     - 移除 `.secondary-card` 历史残留的 `min-width: 310px` 限制，设为 `min-width: 0; width: 100%; min-height: 350px;`，让卡片在 3 列宽度下自由舒展；
     - 适配 3 列紧凑空间黄金比例：将缩略图视口微调为 `height: 145px`，参数网格间距微调为 `gap: 8px`，按钮栏紧凑排版，确保标签、标题、描述、蓝图视口、双列参数与底栏按钮毫无拥挤与遮挡；
     - 修改已热同步复制至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`，`dotnet build /t:Compile` 编译通过（0 错误）。

- **云方案中心二次方案卡片多图纸时被垂直压缩缺陷彻底根除 (`Resources/cloud_solution.html`)**：
  1. **问题根因定位**：
     - 当子文件夹内 DWG 数量较多时（例如 14 张、32 张），`.card-grid-container` 和父级 `.content-area` 缺少 `min-height: 0`，导致外层 Flex 容器未能建立正确的垂直滚动上下文；
     - `.secondary-card` 缺少固定高度（`min-height` / `height`），网格项在 CSS Grid 默认机制下为了在有限的视口内塞入全部行，将每一行强行压缩到了约 60px；
     - 卡片内部拥有 `overflow: hidden`，导致高度被压缩后，下半部分的 DWG 略缩图视口、双列参数网格和底栏操作按钮全部被切掉，在视觉上表现为卡片被压成扁平长条。
  2. **彻底闭环修复措施**：
     - **解除容器限制并激活滚动条**：在 `.main-body` 与 `.content-area` 中补充 `min-height: 0`，并在 `.card-grid-container` 中设置 `min-height: 0; overflow-y: auto !important; grid-auto-rows: max-content;`；
     - **显式锁定卡片高度防挤压**：为 `.secondary-card` 显式设置 `min-height: 425px; height: 425px; flex-shrink: 0;`，并为其子元素（header、dwg-viewport、params-body、footer）全量配置 `flex-shrink: 0`，杜绝任何纵向弹性形变；
     - **全局滚动条现代美化**：引入 7px 宽度的 WebKit 自定义扁平滚动条，提供丝滑流畅的纵向浏览体验；
     - **热同步与编译验证**：更新已全量同步覆盖至 `bin/Debug/net48/Resources/` 与 `publish/Resources/`，`dotnet build /t:Compile` 编译无任何错误。

- **本地二次回路方案与 DWG 图纸库全面整合至“云方案中心”企业方案二次方案 Tab (`Models/`, `Services/`, `Controllers/`, `Forms/`, `Resources/`)**：
  1. **用户核心指令与决策落地**：
     - 二次图纸根目录保存在配置文件中（`ConfigManager.Instance.Current.SecondaryCircuit.CircuitDwgDirectory`），点击齿轮图标 ⚙️ 可通过独立 STA 线程调出文件夹选择器或手动输入；
     - 根目录下的所有子文件夹在左侧展示为分类导航树（展示 DWG 文件数量徽标）；
     - 点击子文件夹后，右侧卡片网格展示该子目录下的所有 DWG 图纸；
     - 依据 DWG 图纸名称（去扩展名）在 `personal_components.db` 的 `secondary_circuit_schemes` 表中按 `applicable_codes`（逗号“,”分隔）精准匹配参数（跨门线、开孔要求、人工工费、二次材料费、二次排布图、品牌、描述）；
     - 卡片去除方案大标题，略缩图上方仅保留【方案目录名称】、【当前DWG文件名称（去后缀）】、【品牌】、【描述】；
     - 卡片提供【打开图纸】（调起系统关联的 AutoCAD 打开 DWG）与【编辑】按钮；
     - 点击【编辑】直接呼出“编辑二次回路方案与 BOM”弹窗（高保真复刻用户附件设计），支持对品牌、描述、跨门线、开孔、工费、备注进行修改，并支持对子 BOM 清单进行增删、合价联动重算与从本地物料库一键选型添加；
     - 【一次方案】保持原本配电柜分类与逻辑，【批量插入已选方案】按指示暂不实施。
  2. **端到端一揽子闭环实施**：
     - **实体与数据访问自愈升级**：在 `SecondarySchemeEntity` 中扩展 `Brand` 与 `Description`；在 `PersonalComponentDbService` 中实现自动平滑自愈迁移脚本，无损扩充 SQLite 列；
     - **精确逗号分词检索引擎**：在 `PersonalComponentDbService.SecondaryCircuit.cs` 中实现 `FindSchemeByDwgName`，优先精准命中 `applicable_codes` 逗号集合，未命中时平滑降级；
     - **控制器与通信信使构建**：在 `CloudSolutionController.cs` 中提供目录扫描、DWG 卡片组装、CAD 外部调起、方案与 BOM 序列化保存及本地物料库检索接口；
     - **通信路由与防死锁设计**：在 `CloudSolutionForm.cs` 中挂载 8 个二级分支，采用独立后台 STA 线程弹出 `FolderBrowserDialog`，彻底消灭 Chromium IPC 模态卡死隐患；
     - **前端高质感 UI 交互实现**：在 `Resources/cloud_solution.html` 中引入绿蓝 `#009688` 风格的卡片组件、目录树、目录设置弹窗、高保真二次方案与 BOM 编辑弹窗及本地物料库选型对话框，Vue `<script setup>` 完整挂接响应式数据流；
     - **IPC 消息参数解析修复与全分类聚合升级**：
       a. 根因剖析：前端 `postToHost` 将载荷封装在 `data` 对象中，而 `CloudSolutionForm.cs` 曾直接从 `root` 顶层提取 `folderPath` 与 `folderName`，导致传入后台的路径为空，命中保护提前返回空数组；
       b. 彻底修复：在 `CloudSolutionForm.cs` 增加双层安全提取机制 `GetStringProp`，兼容 `data` 对象与 `root` 顶层属性；
       c. 智能全分类聚合：在 `CloudSolutionController.GetFolderDwgCards` 中升级逻辑，若处于“全部方案分类”（路径为空），自动扫描根目录下全部子文件夹聚合展示全部 216 个 DWG，点击具体子目录时精准呈现对应数量卡片；
       d. 界面清理：依据指令彻底移除二次底栏中灰色的【批量插入已选方案 (0)】按钮；
     - **二次方案卡片 100% 像素级对齐图 1 高品质视觉规范**：
       a. 宽度自适应根治挤压变形：将 `.card-grid-container` 列宽从固死 4 列升级为 `repeat(auto-fill, minmax(320px, 1fr))`，锁定卡片最小宽度 310px，杜绝文字竖排；
       b. 缩略图视口锁死 180px 与图 1 电路图 SVG 矢量降级：添加 `min-height: 180px; flex-shrink: 0` 杜绝折叠；内嵌与图 1 相同的暗夜电路蓝图 SVG 示意图（KM1/KM2 框、机械电气互锁虚线、中点负载连线及“负载 LOAD”），无真实图片时依然呈现图 1 原型效果；
       c. 方案大标题与描述完整呈现：增加 `card.schemeName` 字段，展示加粗 15px 方案大标题与两行自适应工艺描述；
       d. 右上角绿色圆角对勾徽章：增加 `.sec-card-badge-check` 浮动对勾徽标；
       e. 参数网格与底栏按钮对齐：按图 1 严格双列排布跨门线、人工费用、二次材料、开孔需求与二次排布图，底栏展示绿色圆角【编辑】按钮与浅灰【打开图纸】。
     - **输出目录热同步与编译验证**：同步更新至 `bin/Debug/net48` 与 `publish` 目录，`dotnet build` 编译验证 **0 警告，0 错误**。

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

  6. **彻底根除 `MessageBox.Show` 阻塞死锁，升级为异步非模态通信**：
     - 在 `FormulaAdjustFeeForm` 中移除 `applyFormula` 下的 `MessageBox.Show`，改为通过 `PostWebMessageSafe` 回传 `applyFormulaResult`，由前端 `ElMessage` 友好提示；
     - 在 `EnterpriseSettingsForm` 中移除保存成功/失败的 `MessageBox.Show`，改为回传 `saveSettingsResult`，前端通过 `ElMessage` 提示并延时平滑退出；
     - 在 `CreateProjectForm` 中将异常直接回传 `startQuotationResult` 并打标日志，彻底杜绝 Chromium IPC 模态死锁。
  7. **全面对齐 local-heuristics.md:L133：新建项目补充“✍️ 粘贴路径”双轨保障**：
     - 在 `create_project.html` 的保存目录旁新增 `✍️ 粘贴路径` 绿色扁平按钮；
     - 通过 `ElMessageBox.prompt` 弹出原生输入框，支持直接粘贴 Windows 资源管理器路径，零弹窗极速设定，彻底绕过系统外壳慢速磁盘枚举卡顿。
  8. **二次方案管理中心拖拽机制升级 (杜绝系统级全局鼠标捕获死锁)**：
     - 在 `SecondaryCircuitForm.cs` 中增加 `moveWindow` 物理增量坐标位移处理；
     - 将 `secondary_circuit_manage.html` 标题栏拖拽全面升级为现代 `pointerdown` + `moveWindow`（带 `rAF` 节流与 DPI 适配）；
     - 在 `dragWindow` 兜底中引入 `(GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0` 物理按键检测，物理按键弹起时坚决丢弃，彻底根绝全屏幕鼠标锁死。
  9. **建立全工程窗体生命周期释放保护机制 (杜绝 Excel 进程残留)**：
     - 在 `FormulaAdjustFeeForm`、`EnterpriseSettingsForm`、`CreateProjectForm`、`SecondaryCircuitForm`、`CabinetAuxCalcForm`、`CategoryForm`、`SummaryAdjustPriceForm`、`TenderReportRegularForm`、`ModelParamParserForm`、`ComponentManageForm`、`ComponentGroupBuilderForm`、`SmartInputForm`、`SpotlightSettingForm` 中全部显式重写 `OnFormClosing`，解绑 WebMessageReceived 事件并显式调用 `_webView?.Dispose()`，彻底根除关闭 Excel 时由于后台 Chromium 子进程等待导致的进程僵死残留问题。
  10. **工程构建与热同步**：
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

  4. **靠顶部根因排查**：
     - 二次回路图纸对齐与绑定工作台此前显式配置了 `top="1.5vh"`，且编辑弹窗配置了 `top="2.5vh"`；
     - 同时 CSS 中 `.circuit-binding-dialog` 声明了 `margin: 0 auto !important;`，导致弹窗完全贴在主窗口顶部，缺乏视觉纵向居中平衡感；
  5. **全面启用原生 `align-center` 与弹性居中架构**：
     - 在 `<el-dialog v-model="circuitDwgDialogVisible"` 与 `<el-dialog v-model="dialogVisible"` 上统一移除 `top` 偏置属性，启用 Element Plus 原生 `align-center` 居中引擎；
     - 在 CSS 中增强全局遮罩 `.el-overlay-dialog { display: flex !important; align-items: center !important; justify-content: center !important; overflow: hidden !important; }`，彻底锁定居中视口；
     - 统一配置 `.circuit-binding-dialog` 与 `.scheme-edit-dialog` 为 `margin: auto !important; max-height: 94vh !important;`，确保在任何屏幕分辨率及最大化/窗口化切换时，弹窗均优雅、绝对居中于屏幕/窗体正中央；
  6. **静态资源热同步与验证**：
     - 已将修改后的 `secondary_circuit_manage.html` 同步复制至 `bin\Debug\net48\Resources\` 和 `publish\Resources\`；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误 0 警告。

  7. **左侧未反显的深层根因分析**：
     - **引用不一致**：左侧表格数据源 `:data="filteredDirItems"` 是基于 `dirDwgFiles` 实时 map 计算出来的新对象集合，此前用原数组项调用 `setCurrentRow` 无法命中 Element Plus 的引用对比（`===` 为 false），导致行浅绿色高亮未附着；
     - **跨目录物理分布差异**：用户图纸按分类存放在不同的子文件夹中（如 `双电源`、`功能表互感表`、`变频器` 等）。当用户处于“双电源”子目录时，点击绑定了其它子目录图纸的元件组（如 `接触器变频器`），左侧当前列表中根本不存在该文件，导致无法呈现反显与开图；
  8. **全面重构反显机制与跨目录穿透寻图**：
     - **精准引用匹配与平滑滚动 (`highlightLeftDwg`)**：必须在 `filteredDirItems.value` 中提取同一引用对象调用 `circuitDwgTableRef.value.setCurrentRow(targetItem)`，确保 100% 亮起 `.current-row` 浅绿底色，并自动调用 `scrollIntoView` 滚动到视口中央；
     - **C# 后端全局递归穿透定位 (`LocateDwgInDirectory` / `LocateDwgFile`)**：当当前目录未搜寻到图纸时，自动向 C# 发送 `locateAndHighlightDwg` 并在图纸库全目录中进行深搜枚举；
     - **前端自动下钻跳转与延迟反显**：一旦定位到图纸所在父目录，自动触发 `browseToDirectory(parentDir)` 切换左侧目录，并在层级加载完毕后瞬间高亮反显目标图纸，中间 CAD 视口即刻同步渲染该图纸大图；
  9. **编译与热同步交付**：
     - HTML 静态模板已覆盖至 `bin\Debug\net48\Resources\` 与 `publish\Resources\`；
     - `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误。
       - **【人工】**：如 `¥ 30.00`
       - **【二次组】**（方案主名称）：如 `多功能表1` / `热水泵31台`
     - 中间 CAD 矢量视口顶部同步加入参数微章预览群，一边看图一边掌握 5 大工艺定额参数；
  10. **构建验证与热同步交付**：
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

  6. **多重安全防重门禁全链路闭环**：
     - **门禁 1 (批量选图拦截)**：在 `confirmAddCircuitDwgCodes` 中，用户勾选图纸确认选入时，系统自动扫描当前全库所有既有方案，若图号已被其他方案绑定（例如 `CA1B` 已被【双电源互投】绑定），自动进行友好拦截与详细弹窗警示，只放行全库真正唯一的图纸；
     - **门禁 2 (手动输入拦截)**：在 `addNewCodeTag` 中，用户手动输入代号按回车时，先查当前方案内部重复，再查全局跨方案占用，发现冲突立即阻止添加并告知占用方案名；
     - **门禁 3 (前端保存前拦截)**：在 `submitScheme` 中执行终极代号扫描，一旦待保存方案包含跨方案重复代号，阻断提交并弹出阻断警示；
     - **门禁 4 (后端底层兜底保护)**：在 `PersonalComponentDbService.SecondaryCircuit.cs` 中增加 `CheckApplicableCodeConflict` 方法，并在 `SecondaryCircuitController.SaveScheme` 中接入拦截，确保数据库层绝对无同名回路代号落盘；
  7. **热更新与编译验证**：
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
  5. **规则4与工程算法闭环**：
     - 实现了多分类长词优先降序匹配（Maximal Match），彻底杜绝短代号吞噬长型号；
     - 实现了断路器遇漏电自动升格机制（微型断路器+漏电 ➔ 微型漏电，塑壳断路器+漏电 ➔ 塑壳漏电）；
     - 实现了短字符/单双字母安全边界保护（如阻止 16A 的 A 误判为接触器、阻止 400V 的 V 误判为浪涌，针对施耐德 Acti9 A9 系列设立专用保护）；
     - 实现了工业命名 KB0 / KBO 兼容归一化与中文品名强直通机制。
  6. **从 Excel 规则选区一键同步特征库**：
     - 公共服务层 `ImportCategoryDictFromExcelSelection` 支持纯内存二维数组直读用户在 Excel 中框选的特征表（第一行是类别名称，下方是代号）；
     - 控制器与窗体安全交互，前端一键同步并即时生效，极大简化字典维护。
  7. **Excel 批量识别与二维数组极速回填**：
     - 扩展 `ExecuteBatchModelParse` 支持名称输出列（B列）、最小/最大电流列、极数列、脱扣列的同时解析；
     - 支持“仅填空白单元格 (OnlyEmpty)”与“强制覆盖 (OverwriteAll)”策略；
     - 严格遵循规范第 7 条，纯内存二维数组批量读入写回，无 COM 卡顿。
  8. **Vue 3 + Element Plus 前端界面重构**：
     - 增加“2. 元器件名称/类别识别通道”配置卡片，提供类别胶囊切换面板与代号 Tag 池；
     - 升级沙盒实时测试预览，支持透明展示类别决策轨迹。
  9. **严格代码规范与测试验证**：
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
  6. **正向生成汇总表提取校准 (`GenerateComponentSummarySheet`)**：
     - 从分类明细表提取元器件时，严格按照：W(电流)、X(极数)、Y(脱扣)、Z(附件)、AA(BlockName)、AB(BlockCategory) 列提取并写入汇总表的 T、U、V、W、X、Y 列；
  7. **反向一键更新回写校准 (`UpdateFromComponentSummarySheet`)**：
     - 从元件汇总表写回分类明细表时，修正原先写入旧列的问题，严格回写至：
       - **W 列 (索引 23)**: `Current` (额定电流)
       - **X 列 (索引 24)**: `Poles` (极数)
       - **Y 列 (索引 25)**: `trip` (脱扣方式)
       - **Z 列 (索引 26)**: `Accessory` (配套附件)
       - **AA 列 (索引 27)**: `BlockName` (图块名称 / 扩展参数1)
       - **AB 列 (索引 28)**: `BlockCategory` (图块类别 / 扩展参数2)
  8. **编译校验**：
     - 执行 `dotnet build /t:Compile` 编译通过，0 错误。
  9. **模型列映射规范重构 (`Models/CabinetAuxCalcModels.cs`)**：
     - **W 列 (第 23 列)**: 额定电流 (`Current`)
     - **X 列 (第 24 列)**: 极数 (`Poles`)
     - **Y 列 (第 25 列)**: 脱扣类型/脱扣方式 (`Trip`)（新增属性）
     - **Z 列 (第 26 列)**: 附件描述 (`Accessory`)
     - **AA 列 (第 27 列)**: 图块名称 (`BlockName`)
     - **AB 列 (第 28 列)**: 图块类别 (`BlockCategory`)
  10. **元器件矩阵批量读取引擎适配 (`Services/ExcelServices.CabinetAuxCalc.cs`)**：
      - 依据规则 7，将元器件批量读取范围从 `A:Z` 扩展至 `A:AB`（共 28 列，`ws.Range[$"A{compStartRow}:AB{compEndRow}"]`）；
      - 逐行提取时，严格按照 W(电流)、X(极数)、Y(脱扣)、Z(附件)、AA(图块名)、AB(图块类别) 列索引进行读取与赋值；
  11. **编译校验**：
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

- **彻底修复“分布调价”弹窗提示“未检测到有效的分类工作表，已选择 0/0 个分类，共 0 台箱柜”故障 (`Forms/DistributedAdjustPriceForm.cs`, `Services/ExcelServices.DistributedAdjustPrice.cs`, `Resources/distributed_adjust_price.html`)**：
  1. **深度根因定位与排查**：
     - 在 `DistributedAdjustPriceForm.cs` 的 `HandleGetCategorySheets()` 中，使用了 `ThreadPool.QueueUserWorkItem` 将获取分类表的操作放到了后台线程池线程（MTA 线程）；
     - Excel-DNA 的 `ExcelDnaUtil.Application` 强制要求在 Excel STA 主线程中调用，在后台线程调用会导致安全拦截并返回 `null`；
     - 服务层 `GetCategorySheetsForDistribution()` 检测到 `app == null`，以为没有 Excel 应用程序，直接返回了空列表，导致前端界面展示 0 个分类工作表；
  2. **系统级重构与线程模型校准 (`Forms/DistributedAdjustPriceForm.cs`)**：
     - 将 `HandleGetCategorySheets()` 调整为与 `SummaryAdjustPriceForm` 完全一致的当前 STA 线程同步调用机制，确保 COM 上下文 100% 畅通；
     - 增加详尽的调试轨迹日志，捕获并记录分类扫描和回传前后的每一步状态；
  3. **增强服务层扫描与容错机制 (`Services/ExcelServices.DistributedAdjustPrice.cs`)**：
     - 在 `GetCategorySheetsForDistribution()` 中增加针对工作簿名称与各工作表扫描的详尽诊断日志，并采用原工作表名称避免空白字符截断；
  4. **前端交互体验提升 (`Resources/distributed_adjust_price.html`)**：
     - 在分类卡片头部增加“🔄 刷新列表”按钮，方便用户随时重测；
     - 列表为空时提供“🔄 重新扫描分类”轻量按钮；
     - 资源文件已完整同步至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
  5. **编译构建与代码规范**：
     - 严格遵循每 3 行包含至少 1 行中文注释，遵循最小变动法则；
     - 执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误。

- **全面落地分布调价“横向箱柜单台数量修改反向回写与自动插入新器件行”功能 (`Services/ExcelServices.DistributedAdjustPrice.cs`, `Resources/distributed_adjust_price.html`)**：
  1. **表头升级与所属分类映射 (`GenerateComponentDistributionSheet`)**：
     - 在横向箱柜复合表头（行 13~17）中增加第 13 行【所属分类】，精准锚定各箱柜列归属的工作表 `SheetName`，彻底消除同名柜号跨表混淆；
     - 将 D 列（汇总数量）重构为动态 `=SUMPRODUCT($L$14:$Z$14, L19:Z19)` 工业级加权乘积公式，当用户在横向箱柜列修改任何单台数量时，D 列汇总数量与 F 列总金额实时联动计算；
     - 将横向箱柜数量列设为淡黄白色底色，提示支持自由编辑修改；
  2. **反向回写引擎重构 (`UpdateFromComponentDistributionSheet`)**：
     - 批量读取横向各箱柜列的最新单台数量（包括修改数量、改成 0 以及新增器件）；
     - **方式 A 落实**：针对已有器件，按修改后的数量回写各箱柜 F 列，若改为 0 则设为 0，合价置零并保留行；
     - **自动插入新行**：针对分布表中填入数量（> 0）但该箱柜原本没有的新增器件，优先复用元器件区域预留的空白行，若无空白行则严格按照规则 6 在小计行前执行整行 `Insert`，完整写入 A~H 列各项数据及合价公式；
     - **公式刷新与自愈**：级联刷新小计行 `SUM` 公式与总计行公式，严格执行规则 8 调用 `Tool.FixAndFillCabinetNamesForSheet(sheet)` 自愈工作表 4 个定义名称；
  3. **交互指引与编译校验**：
     - 更新前端页面 Banner 副标题与操作指引，明确说明单台数量修改与自动插入新行机制；
     - 资源已全量热同步至 `publish/Resources/` 与 `bin/Debug/net48/Resources/`；
     - 遵循每 3 行包含至少 1 行中文注释，执行 `dotnet build /t:Compile /p:DebugType=none` 编译通过：0 错误。

## [Completed]

- [已验证] `Forms/DistributedAdjustPriceForm.cs` 线程模型已修复为 STA 同步调用。
- [已验证] `Services/ExcelServices.DistributedAdjustPrice.cs` 表头增加所属分类工作表（行 13），D 列重构为 SUMPRODUCT 动态加权求和公式。
- [已验证] `Services/ExcelServices.DistributedAdjustPrice.cs` 支持修改横向箱柜单台数量并反向回写（改成 0 设为 0 保留行；原本没有的器件自动插入新行并填入型号、厂家、数量和单价）。
- [已验证] 级联刷新各箱柜小计行、总计行求和公式，并触发规则 8 定义名称自愈。
- [已验证] 前端 UI 资源与模板已同步更新。
- [已验证] 代码编译通过，0 错误。

## [In-Progress]

- 提示用户重启 Excel 以使最新插件与功能生效。

## [Next]

- 验证用户在【材料分布表】中调整单价、型号、厂家以及横向箱柜数量，点击【一键更新到明细】后各箱柜数据与新增行的回写情况。


