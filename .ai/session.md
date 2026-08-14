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
     - 顶部空行复用与序号 `=ROW()-ROW(A$6)` 公式保留。
     - 底部明细块（32 行插槽）特征文本搜索与整块插入。
     - **强类型定义名称与超链接**：直接在 `activeSheet` 内部复制模板行，彻底取消打开/关闭外部 `CabinetTemplate.xlsx` 文件，根除跨工作簿句柄清理导致定义名称丢失的致命隐患；以强类型 `Range` 对象直接调用 `targetWb.Names.Add`；带工作表前缀绑定超链接 `SubAddress`。
     - **GetNextCabinetIndex 序号动态增量**：将读取逻辑全面切换为 `Cell.Value2` 内存数据，配合 `ExtractIndexFromName` 清洗助手，实现 100% 递增 `maxK + 1` 与历史定义名称永久保留。
     - **SheetFollowHyperlink 视图定位**：监听超链接跳转，读取全局配置 `ScrollRowOffset` (如 `-3`) 修正 `win.ScrollRow`。
- **最新构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 警告 0 错误）。已彻底解决历史定义名称丢失问题。
- **当前任务状态 [新建箱柜重构]**：
  1. **配置驱动模板总行数 A**：在 `appsettings.json` 与 `AppConfig.cs` 的 `ExcelSettings` 中维护 `TemplateDetailTotalRows`（默认 32 行），彻底杜绝行数硬编码。
  2. **整块复制与行数动态对齐**：根据当前活动位置选择插队或末尾追加，复制本表已有箱柜明细整块（行数 B）；动态比较 B 与 A，在元器件区域自动删除多余行（$B > A$）或插入补齐行（$B < A$ 并复制格式），确保新箱柜总行数严格等于 A。
  3. **元器件区域智能清洗**：采用内存二维数组批量读取与写回（规则 7），精准识别并**原样保留以 '=' 开头的计算公式**（如序号公式、总价公式、成本总价公式等），仅对常量数据（元件名称、规格型号、厂家、数量、单价等）执行置空清洗。
  4. **4 个定义名称与超链接重新绑定**：严格遵循规则 6 架构，更新箱柜信息行名称，强类型注册 `Cab_Sum_K`, `Cab_Det_K`, `Cab_Subsum_K`, `Cab_Tolsum_K`，并完成双向超链接跳转与汇总行公式联动。
  5. **构建与打包**：`dotnet build` 编译 0 警告 0 错误通过，打包已生成最新的 `ExcelAddInDemo-AddIn64-packed.xll`。


