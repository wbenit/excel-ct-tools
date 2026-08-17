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

3. **Excel-DNA 单文件打包 (-packed.xll) 依赖包含**：
   - NuGet 依赖（WebView2, System.Text.Json 等）须在 `.dna` 文件中声明 `<Reference Path="...dll" Pack="true" />` 方可打入 packed.xll 中。
   - 前端资源需在 csproj 配置 `CopyDependenciesToPublish` Target 目标发布，C# 层配合多级备用路径检索。

4. **原生 HTML 内嵌 Vue 模板中禁用 `/>` 自闭合语法（防按钮 DOM 被吞）**：
   - 在 HTML 文件挂载 Vue 3 时，浏览器 DOM 解析器会将 `<el-input ... />` 中的 `/>` 忽略，把紧随其后的同级 `<el-button>` 误解析为 `<el-input>` 的子元素，导致 Element Plus 渲染时直接抛弃该按钮节点。
   - **规则**：所有 Vue 自定义组件在 HTML 页面中必须使用显式闭合标签（如 `<el-input></el-input>`）。

5. **Excel COM 定义名称（Defined Names）与模板句柄生命周期**：
   - 从模板工作簿复制插入内容后需**立即关闭 `templateWb` 句柄**，防止名称误挂载在只读模板上被静默销毁。
   - `Names.Add` 的 `RefersTo` 参数必须传入以 `=` 开头的标准公式字符串（如 `$"='{sheetName}'!$A${row}"`）。

---

# 三、 最新进度与状态记录 🚀

- **已完成架构与功能**：
  1. **Ribbon 工具栏**：实现了全套“我的项目”、“①建项目→”、“②录元件”、“③调价格→”、“④出报表”等分组与按钮菜单。
  2. **全局 UI 主题规范**：根据规则要求统一升级所有 Web 窗口为**绿蓝相间主题，主色调为 `#009688`**。
  3. **“我的企业设置”窗口**：基于 C# + WebView2 + Vue 3 (Element Plus, `<script setup>`)，反显解析正常。
  4. **“新建项目”窗口与控制器**：支持项目创建、 SaveAs 模板清洗、 Win32 强力置顶与视口最大化。
  5. **“新建箱柜”功能与定义名称/超链接联动**：
# 一、 商业加壳与代码安全注意事项 🔒

## Ribbon 回调必须保留名称（Exclude Obfuscation）

- **规则**：Excel 菜单通过 XML 字符串（如 onAction="MyMethod"）利用反射机制调用 C# 方法。
- **注意**：在 RibbonController.cs 中新增的 Ribbon 回调函数，必须确保该类或方法拥有 `[Obfuscation(Exclude = true)]` 标记，避免加壳重命名导致找不到方法。
- **最佳实践**：将界面回调集中放在 RibbonController.cs，而将业务逻辑解耦在单独类中。

---

# 二、 启发式经验与踩坑结晶 💡

1. **Excel Ribbon CustomUI 严苛 Schema 防隐藏**：
   - 使用不存在的 `imageMso` 或空的 `group label=''` 会导致 Excel 静默抛弃整个自定义 Ribbon Tab。需使用全版本通用的标准 imageMso 名称。

2. **Excel-DNA Modeless 非模态窗体与线程安全机制**：
   - 保持弹窗时 Excel 可自由编辑：使用 `ShowModelessForm`，调用 `form.Show(new ExcelWin32Window(excelHwnd))` 挂载 Owner 为 Excel 主句柄，既避免窗口下沉，又保证 Excel 保持非阻塞可编辑状态。
   - 单例生命周期复用：维护窗体引用，已打开时调用 `BringToFront()` / `Activate()` 避免重复多实例。
   - **CoreWebView2 必须在 UI 线程访问**：异步方法（如 `await Task.Run(...)`）在非模态环境下恢复时可能处于 Worker 线程，所有 `PostWebMessageAsJson` 与 WinForms 控件交互必须通过 `SafeInvoke`（`this.Invoke` / `this.BeginInvoke`）切换至 UI 线程，防止 `CoreWebView2 can only be accessed from the UI thread` 异常。

3. **Excel-DNA 单文件打包 (-packed.xll) 依赖包含**：
   - NuGet 依赖（WebView2, System.Text.Json 等）须在 `.dna` 文件中声明 `<Reference Path="...dll" Pack="true" />` 方可打入 packed.xll 中。
   - 前端资源需在 csproj 配置 `CopyDependenciesToPublish` Target 目标发布，C# 层配合多级备用路径检索。

4. **原生 HTML 内嵌 Vue 模板中禁用 `/>` 自闭合语法（防按钮 DOM 被吞）**：
   - 在 HTML 文件挂载 Vue 3 时，浏览器 DOM 解析器会将 `<el-input ... />` 中的 `/>` 忽略，把紧随其后的同级 `<el-button>` 误解析为 `<el-input>` 的子元素，导致 Element Plus 渲染时直接抛弃该按钮节点。
   - **规则**：所有 Vue 自定义组件在 HTML 页面中必须使用显式闭合标签（如 `<el-input></el-input>`）。

5. **Excel COM 定义名称（Defined Names）与模板句柄生命周期**：
   - 从模板工作簿复制插入内容后需**立即关闭 `templateWb` 句柄**，防止名称误挂载在只读模板上被静默销毁。
   - `Names.Add` 的 `RefersTo` 参数必须传入以 `=` 开头的标准公式字符串（如 `$"='{sheetName}'!$A${row}"`）。

---

# 三、 最新进度与状态记录 🚀

- **已完成架构与功能**：
  1. **Ribbon 工具栏**：实现了全套“我的项目”、“①建项目→”、“②录元件”、“③调价格→”、“④出报表”等分组与按钮菜单。
  2. **全局 UI 主题规范**：根据规则要求统一升级所有 Web 窗口为**绿蓝相间主题，主色调为 `#009688`**。
  3. **“我的企业设置”窗口**：基于 C# + WebView2 + Vue 3 (Element Plus, `<script setup>`)，反显解析正常。
  4. **“新建项目”窗口与控制器**：支持项目创建、 SaveAs 模板清洗、 Win32 强力置顶与视口最大化。
      - **前端与宿主窗体**：创建 [Forms/CategoryForm.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Forms/CategoryForm.cs) 与 [Resources/category.html](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Resources/category.html)，基于 WebView2 + Vue 3 (`<script setup>`) + Element Plus 构建，包含实时重名提示、公式组选择与绿蓝相间主题（`#009688`），并修复了 `@mousedown.stop` 拖拽冒泡吞点击的隐患。
    - **最新构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。
   10. **分类工作表标准行号智能动态探测（彻底杜绝硬编码行号 +1 偏移与 #REF! 异常）**：
       - **原因定位**：不同模板（如 CabinetTemplate.xlsx 与动态生成的备用模板）的顶部汇总表头所在行数不同（Row 3 vs Row 6），且原有配置中 `CabTolsumRowIndex` (68 vs 71) 与硬编码的 `CabSumRowIndex` 不匹配，导致克隆新分类表时在 Row 5 错误写入超链接并保留了原模板 Row 4 的旧数据，同时公式引用越界产生 `#REF!`。
       - **修复方案**：
         1. 在 [Tool.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Tool.cs) 中新增 `FindStandardCategoryRowIndexes(dynamic sheet)`，通过内存二维数组扫描表头特征（“序号”、“柜号”、“设备”、“小计”、“总计”）智能计算并返回 `cabSumRow`、`cabDetRow`、`cabSubsumRow`、`cabTolsumRow`。
         2. 在 [Services/ExcelServices.Category.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Services/ExcelServices.Category.cs) 的 `InitializeCategorySheet` 中直接接入动态探测结果，精准回填首台箱柜与计费矩阵，自动适配任何模板布局。
         3. 修正 [appsettings.json](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/appsettings.json) 中 `CabTolsumRowIndex` 为标准 71。

    11. **“智能输入模式”与选择表联动（元器件去重词库与属性联动回填）**：
       - **业务与架构**：
         1. 在 Ribbon【辅助项】->【项目工具】下拉菜单中新增【智能输入】按钮（`btnSmartInput`），点击非模态弹出基于 WebView2 + Vue 3 的智能输入配置窗口。
         2. **刷新提取**：基于箱柜定义名称（`Cab_Det_X` 至 `Cab_Subsum_X`）与规则 6/7，采用内存二维数组一次性读取所有工作表的元器件插槽（`Cab_Det + 2` 至 `Cab_Subsum - 1`），以 C 列规格型号去重，将结构化数据存储于 `data/smart_components.json`。
         3. **数据源表多选**：界面支持多选数据源工作表（展示各表去重条目数，支持全选/清空），配置自动持久化于 `data/smart_input_config.json`。
         4. **回填字段范围设置**：支持用户多选 B 列(名称)、D 列(厂家)、E 列(单位)、G 列(单价) 的联动勾选框。
         5. **双模选择与联动**：
            - 支持一键为当前工作表各箱柜 C 列批量生成并绑定 Excel 原生下拉列表数据验证（引用自动维护的 `选择表` 独立工作表）；
            - 在 `ExcelEventManager.OnSheetChange` 中挂载 C 列单元格监听，当用户下拉选择或输入规格型号时，自动按勾选字段联动回填同行的 B/D/E/G 列；
            - 在配置窗口中提供物料候选模糊检索、双击一键填入当前活动行的录入助手面板。
       - **代码规范**：新增 `Models/SmartInputModels.cs`、`Controllers/SmartInputController.cs`、`Services/ExcelServices.SmartInput.cs`、`Forms/SmartInputForm.cs`、`Resources/smart_input.html`，严格遵循每 3 行包含一行中文注释与 `#009688` 绿蓝相间主题规范。

   9. **通用计算与底层辅助方法统一沉淀到 Tool.cs (internal)**：
      - **架构定位与分层**：
        1. [Tool.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Tool.cs) 声明为 `internal static class Tool`，定位为**通用工具与纯计算/数据转换层**。涵盖路径获取、名称提取与清洗（`ExtractCleanNameStr`, `ExtractIndexFromName`）、公式平移（`TransformFormulaRowOffset`）、计费矩阵二维数组转换（`BuildFeeMatrix`）、公式路径清洗（`CleanRangeFormulas`）、行对齐（`AlignRowRangeCount`）、定义名称字典构建（`BuildCabinetMap`）以及名称自动校准补齐等。通过 `internal` 访问级别，彻底隔离 Excel-DNA 自动扫描，避免函数误注册与 UDF 重名弹窗。
        2. [Services/ExcelServices.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/Services/ExcelServices.cs) 系列分部类定位为 **业务服务与 Excel COM 流程调度层**，通过直接调用 `Tool.***` 完成底层计算与转换，职责清晰解耦。
      - **代码规范**：所有新增与重构代码严格遵循每 3 行包含中文注释规范，无重复方法定义。
