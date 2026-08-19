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
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。
  16. **每个箱柜明细块自带 3 行报价人信息体系重构（已落地）**：
      - **业务与架构**：
        1. 确立每个箱柜明细块从大标题（`Cab_Det - 3`）至总计行下方 3 行报价人信息（`Cab_Tolsum + 3`）的完整物理闭环；
        2. 在 [Services/ExcelServices.Cabinet.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.Cabinet.cs) 的 `CreateNewCabinet` 中将 `lastDetBlockEnd` 精准计算为 `maxTolsumRow + 3`，模板明细块复制源完整覆盖至 `templateTolsumRow + 3`，新总计行精准推导为 `newDetailStartRow + (templateTolsumRow - templateStartRow)`，使每个新创建的箱柜均完整自带专属 3 行报价人信息；
        3. 强化顶部汇总表插行判定（结合 A、B 列），并在插行触发时对下方明细行进行 `+1` 偏移补偿；
        4. 同步升级 [Services/ExcelServices.Cabinet.cs](file:///e:/Ace/excel-ct-tools/Services/ExcelServices.Cabinet.cs) 的 `DeleteCabinets`（删除区间为 `detRow - 3` 至 `tolsumRow + 3`）及 [Tool.cs](file:///e:/Ace/excel-ct-tools/Tool.cs) 的 `FindCabinetByRow`（命中区间扩展至 `tolR + 3`）；
        5. 修正汇总表 G 列单价公式为标准指向总计行的 `$"=H{newTolsumRow}"`。
      - **构建状态**：`ExcelAddInDemo.dll` 编译完全通过（0 错误）。

