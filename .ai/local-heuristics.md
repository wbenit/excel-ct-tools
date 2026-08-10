# 启发式经验与项目开发踩坑指南 (Local Heuristics)

### 1. Excel Ribbon UI 静默失败/不显示
- **现象**：修改 Ribbon XML 后，Excel 打开时选项卡完全不显示。
- **原因**：XML 中使用了无效/不存在的 `imageMso` 图标名称或空的 `group` `label=''` 属性，触发了 Office CustomUI 静默校验抛弃。
- **解决**：确保 `imageMso` 为 Office 通用标准图标名称，且标签非空，`xmlns` 建议使用全版本兼容的 `http://schemas.microsoft.com/office/2006/01/customui`。

### 2. WebView2 / Form 弹窗闪退与 0x800A03EC 硬禁用
- **现象**：点击按钮弹窗时 Excel 突然闪退崩溃，再次打开提示 `LoadComAddIn / TargetInvocationException / COMException 0x800A03EC`。
- **原因**：在独立的后台线程中调用 `Application.Run(form)`，遇到未捕获异常导致进程崩溃，Excel 机制将插件列入硬禁用列表。
- **解决**：
  1. 取消独立线程，获取 `ExcelDnaUtil.WindowHandle` 并使用 `form.ShowDialog(new ExcelWin32Window(excelHwnd))` 在主线程模态显示。
  2. 在 Excel 选项 -> 加载项 -> 禁用项目中将硬禁用的 Add-In 重新启用。

### 3. Excel-DNA 单文件打包 (-packed.xll) 缺少程序集 (FileNotFoundException)
- **现象**：提示 `未能加载文件或程序集 Microsoft.Web.WebView2.WinForms ...`。
- **原因**：`ExcelDnaPack` 默认只打包主程序集，不会打包 NuGet 依赖 DLL。
- **解决**：在 DNA 配置文件中显式添加 `<Reference Path="Microsoft.Web.WebView2.WinForms.dll" Pack="true" />`，指示打包器将依赖压入 packed.xll 中。

### 4. 发布（Publish）目录下缺失 Resources 资源文件
- **现象**：提示 `未找到界面资源文件: .../publish/Resources/enterprise_settings.html`。
- **原因**：编译时 Resources 文件夹仅复制到了 bin/Debug 目录，没有发布至 publish 输出子目录。
- **解决**：
  1. 在 csproj 项目中增加 `CopyDependenciesToPublish` 构建 Target 目标，在打包后将 Resources 自动同步至 publish/Resources。
  2. 在 C# 窗体代码中配置多重备用路径检索，按优先级自动寻找 HTML 文件。

### 5. 原生 HTML 内嵌 Vue 模板下，自定义组件禁用 `/>` 自闭合语法（DOM 节点被吞）
- **现象**：在 HTML 文件中，紧跟在 `<el-input ... />` 后的同级按钮 `<el-button>` 根本没有渲染，DevTools 审查元素发现该按钮 DOM 节点完全消失。
- **原因**：HTML5 规范规定自定义标签（如 `<el-input>`）不是 Void 元素，浏览器原生 DOM 解析器会忽略 `/>` 斜杠，误将后续同级节点（如 `<el-button>`）解析为 `<el-input>` 的内部子节点。由于 Element Plus `el-input` 没有默认插槽，挂载时直接丢弃了非插槽子节点。
- **解决**：在纯 HTML 模板中，所有 Vue 自定义组件必须使用显式闭合标签（如 `<el-input></el-input>`），绝不可使用 `/>` 自闭合语法。

### 6. 表单输入框固定宽度导致同级右侧按钮挤出视口
- **现象**：DOM 中存在按钮，但页面界面右侧看不到按钮。
- **原因**：输入框设置了 `width: 360px !important`，Label + 输入框 + 按钮总宽度超过窗口 ClientSize 宽度，被 `body { overflow: hidden }` 遮挡裁剪。
- **解决**：将输入框样式修改为 `flex: 1 !important; min-width: 0 !important;`，使其自适应剩余宽度，确保右侧按钮完整在视口内展示。

### 7. 复制 Sheet 导致的公式跨文件外部绝对路径问题
- **现象**：从模板 `CabinetTemplate.xlsx` 跨工作簿 `sheet.Copy()` 到新工作簿后，新工作簿里的公式带上了原模板文件的外部路径前缀（如 `='C:\...\CabinetTemplate.xlsx'!分类1!B2`）。
- **原因**：Excel COM `sheet.Copy()` 在跨 Workbook 拷贝时，会自动为引用了源工作簿其他 Sheet 的公式重写为外部文件链接。
- **解决**：改用 `app.Workbooks.Open(templatePath, ReadOnly: true)` 打开模板后直接调用 `SaveAs(targetFilePath)` 保存为新工作簿，并使用 `UsedRange.Replace` 自动清洗公式中残存的外部文件名与路径前缀。

### 8. Excel SDI 架构下 COM Activate 无法拉起前台窗口（Win32 操作系统级置顶解法）
- **现象**：创建新 Workbook 并调用 `wb.Activate()` 后，界面依然死死卡在旧工作簿窗口上。
- **原因**：Excel 2013+ 采用了 **SDI（单文档界面）架构**，每个工作簿在 Windows 操作系统中都是一个**独立的 top-level 视口 HWND 窗口**。仅仅在 COM 层调用 `wb.Activate()` 或 `win.Activate()` 只能改变 Excel 内部活动指针，**无法指挥 Windows 操作系统在 Z-Order 顶层切换 HWND 窗口置顶**。
- **解决**：从 `wb.Windows[1].Hwnd` 提取工作簿窗口的 Windows HWND 句柄，使用 Win32 API `ShowWindow(hwnd, 9)`（还原/展示）以及 `SetForegroundWindow(hwnd)` 强制指挥 Windows 操作系统将新工作簿 HWND 拉至 Desktop 最前台。

### 9. Excel MDI 视口窗口隐匿导致界面显示空白无工作簿
- **现象**：新建/激活工作簿后，Excel 主界面没有展示表格网格，而是显示一片无工作簿的空白/灰色背景画布。
- **原因**：通过 COM 接口打开或新建工作簿时，其视口窗口 `wb.Windows[1]` 的 `Visible` 属性未显式开启或处于最小化隐藏状态。
- **解决**：在 `wb.Activate()` 激活工作簿后，强力显式设置 `wb.Windows[1].Visible = true` 以及 `wb.Windows[1].WindowState = -4137` (即 `xlMaximized`)，确保工作簿窗口最大化高亮展现在 Excel 画布中央。

### 10. WinForms 控件在句柄未创建或窗体释放时调用 Invoke 抛出异常
- **现象**：提示 `处理消息发生错误: 在创建窗口句柄之前，不能在控件上调用 Invoke 或 BeginInvoke。`
- **原因**：在 WebView2 异步消息回调执行时直接使用 `this.Invoke`，若此时窗体未完成句柄创建（`IsHandleCreated == false`）或者已被 `this.Close()` 销毁（`IsDisposed == true`），WinForms 会强制抛出 `InvalidOperationException`。
- **解决**：封装 `SafeInvoke(Action action)` 方法，先判断 `!IsDisposed && IsHandleCreated`；再判断 `InvokeRequired`，若为主线程则直接同步运行 `action()`，避免盲目调用 `Control.Invoke`。

### 11. Excel-DNA XLL 插件中 SheetChange 等 COM 全局事件响应失效/未触发的结晶解法
- **现象**：在 `IExcelAddIn.AutoOpen()` 中通过 `((Application)ExcelDnaUtil.Application).SheetChange += ...` 注册全局事件后，修改 Excel 单元格没有任何响应，诊断发现事件根本没有被触发。
- **原因**：
  1. **COM 消息循环未就绪**：在 `AutoOpen()` 触发时刻，Excel 主线程的 COM 消息循环（Message Loop）与 COM 连接点（Connection Point）尚处于初始化未就绪状态，在此节点同步访问 `ExcelDnaUtil.Application` 并绑定事件，导致 COM 订阅接口被 Excel 内部静默丢弃。
  2. **过度防御布尔标志位的拦截陷阱**：在注册代码中使用 `if (!_isEventsInitialized)` 布尔变量防护，由于 `AutoOpen` 阶段已将其误设为 `true`，后续用户操作（如新建文件或运行业务逻辑）再次尝试补救注册时，被该标志位直接拦截返回，导致事件永远无法真正绑定。
  3. **Excel 系统级 `EnableEvents` 被隐式关闭**：在发生任何 COM 异常或在多线程/代码崩溃后，Excel 系统级的 `EnableEvents` 可能残留为 `false`，导致 Excel 引擎停止广播一切 `SheetChange` 事件。
- **结晶解法（终极规范）**：
  1. **延迟就绪注册**：在 `AutoOpen()` 中严禁直接挂载事件，必须使用 `ExcelAsyncUtil.QueueAsMacro(() => { ... })` 将事件注册延迟推迟到 Excel 消息循环完全准备完毕后再进行。
  2. **强制开启系统事件**：在注册函数入口处显式调用 `app.EnableEvents = true;`，强制拉高 Excel 事件引擎的开启状态。
  3. **解绑-重绑模式**：彻底废弃 `_isEventsInitialized` 布尔拦截，采用 `app.SheetChange -= OnSheetChange;` 先解绑旧委托，再 `app.SheetChange += OnSheetChange;` 重新强行挂载，确保 COM Connection Point 随时保持 100% 活着且唯一。
