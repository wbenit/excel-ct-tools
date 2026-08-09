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

---

# 三、 最新进度与状态记录 🚀

- **已完成架构与功能**：
  1. **Ribbon 工具栏**：实现了全套“我的项目”、“①建项目→”、“②录元件”、“③调价格→”、“④出报表”等分组与按钮菜单。
  2. **全局 UI 主题规范**：根据规则要求统一升级所有 Web 窗口为**绿蓝相间主题，主色调为 `#009688`**。
  3. **“我的企业设置”窗口**：
     - 基于 C# + WebView2 + Vue 3 (Element Plus, `<script setup>`)。
     - 完成 JSON 属性大小写与驼峰解析的完美双向反显绑定。
  4. **“新建项目”窗口与控制器**：
     - 新增 `CreateProjectForm.cs` 与 `ProjectController.cs`。
     - 前端 `Resources/create_project.html` 完美还原默认收起与展开“更多...”视图。
     - **保存目录右侧配置了专属 `[ 更改目录 ]` 按钮**（弹出原生 WinForms `FolderBrowserDialog` 对话框）。修复了原生 HTML 解析器因 `<el-input />` 自闭合语法将同级 `<el-button>` 误解析为子节点导致 Element Plus 忽略抛弃不渲染的问题（现已修改为显式闭合 `</el-input>` 标签）。
     - **支持无边框与防冒泡关闭**：将 `CreateProjectForm.cs` 修改为 `FormBorderStyle = FormBorderStyle.None` 沉浸式无边框窗口；引入 Windows 原生 Win32 `ReleaseCapture` 与 `SendMessage` 配合前端顶栏拖拽；对右上角关闭/刷新按钮添加 `@mousedown.stop.prevent` / `@click.stop.prevent` 与 `e.stopPropagation()` 彻底防止点击关闭时冒泡触发窗口拖拽。
     - **解决 Invoke 句柄未创建异常**：在 `CreateProjectForm.cs` 与 `EnterpriseSettingsForm.cs` 中引入 `SafeInvoke` 方法，在进行跨线程或 UI 调度前校验 `this.IsDisposed` 和 `this.IsHandleCreated` 并在非跨线程下直接同步调用，彻底修复了 WebView2 消息响应时抛出“在创建窗口句柄之前，不能在控件上调用 Invoke 或 BeginInvoke”的异常弹窗。
     - **终极解决焦点被操作系统强行复位回旧工作簿（Win32 原生置顶与最大化）**：分析出根本原因为 Excel 2013+ 采用 SDI 单文档架构，每个工作簿拥有独立的 Windows 顶栏 HWND，仅靠 COM 层的 `wb.Activate()` 无法触发操作系统级别的 Z-Order 窗口提拉。在 `ExcelServices.cs` 中引入 Win32 `SetForegroundWindow` 与 `ShowWindow(3)` (即 `SW_SHOWMAXIMIZED`) API，配合 `win.Hwnd` 句柄进行操作系统级别的硬置顶与最大化，彻底解决切不回新工作簿以及新窗口未全屏展现的问题。
     - **解决公式绝对路径与焦点切回旧工作簿**：将模板复制方式重构为直接 `SaveAs` 打开的 `CabinetTemplate.xlsx` 模板文件，并自动清洗 `[CabinetTemplate.xlsx]` 与外部文件路径前缀，确保公式与模板 100% 一致且不含外部绝对路径。
     - **解决 Excel 视口空白与焦点恢复**：在 `ProjectController.cs` 中精简移除了弹窗关闭前提前激活工作簿的冗余代码，统一由 `ExcelServices.cs` 在弹窗关闭后通过 `ActivateCreatedWorkbook` 显式设置 `newWb.Windows[1].Visible = true` 以及 `newWb.Windows[1].WindowState = -4137 (xlMaximized)`，确保新建工作簿窗口最大化高亮展现在用户面前。
     - 支持 `报价单号-项目名称-报价人` 的实时文件名联动与自动弹窗扩展。
  5. **箱柜下拉菜单（menuCabinet）**：
     - 已根据设计图更新 Ribbon XML 配置文件 [RibbonController.cs](file:///c:/Users/ADMIN/.gemini/antigravity/scratch/ExcelAddInCTtools/RibbonController.cs)，完整配置了包含 11 个子菜单项的箱柜功能列表（新建箱柜、新建无明细箱柜、批建箱柜、编辑箱柜信息、剪切箱柜、复制箱柜、插入复制的箱柜、删除箱柜、箱柜调序、导入箱柜BOM、智能导入箱柜BOM）。
     - 正确设置了对应的 imageMso 图标样式与缩进规范。
- **最新构建状态**：`dotnet build` 构建通过，生成 `0 个警告，0 个错误`。
