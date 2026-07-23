# Excel C# 插件 (Excel-DNA) 实例与商业加壳使用指南

本项目为一个完整可编译运行的 Excel C# 插件实例，实现了在 Excel GUI 中增加自定义 Ribbon 选项卡，集成了 **按钮 (Button)**、**复选框 (CheckBox)** 和 **文本输入框 (EditBox)** 等组件及功能，并具备完美的 **商业加壳 (Commercial Obfuscation)** 兼容架构。

---

## 商业加壳加固工作流指南

为了保护你的核心代码不被反编译（如 dnSpy、ILSpy 逆向），你可以使用各类商业或开源 .NET 加壳/混淆工具（例如 **ConfuserEx**、**NetGuard**、**Dotfuscator**、**VMProtect**、**SmartAssembly** 等）。

### 加壳与防逆向步骤：

1. **执行编译**
   在项目根目录下运行编译指令：
   ```bash
   dotnet build -c Release
   ```
   编译产物位于 `bin/Release/net48/` 目录下的 `ExcelAddInDemo.dll`。

2. **加壳/混淆核心 DLL**
   - 使用你的商业加壳工具打开 `bin/Release/net48/ExcelAddInDemo.dll`。
   - **混淆配置规则**：
     - **核心业务逻辑**（如 `ExcelServices.cs` 中的算法、计算、数据处理逻辑）：开启 **100% 最高强度混淆**（包含控制流混淆、字符串加密、防调试、防篡改）。
     - **Ribbon 回调控制器**（`RibbonController.cs`）：代码中已预置 `[Obfuscation(Exclude = true, ApplyToMembers = true)]` 属性。在加壳工具中确保勾选 **"尊重/遵循 System.Reflection.Obfuscation 特性"**，加壳工具将自动跳过对 `RibbonController` 类名与 `OnInsertDataClicked` 等回调方法名的重命名，确保 Excel 反射调用正常工作。

3. **重新压制为单文件独立插件 (.xll)**
   加壳完成后，替换混淆后的 `ExcelAddInDemo.dll`，运行 Excel-DNA 压制打包：
   ```bash
   dotnet build -c Release
   ```
   在 `bin/Release/net48/publish/` 目录下将自动生成最终加壳保护的独立发布产物：
   - `ExcelAddInDemo-AddIn64-packed.xll` (适用于 64位 Office Excel)
   - `ExcelAddInDemo-AddIn-packed.xll` (适用于 32位 Office Excel)

---

## 功能说明与 Excel 测试方法

### 1. 安装与测试
- 双击或直接拖拽 `bin/Debug/net48/publish/ExcelAddInDemo-AddIn64-packed.xll` 到 Excel 窗口中（若使用 64位 Excel）。
- 在弹出的安全提示中点击 **“启用此嵌入的宏/插件”**。

### 2. Excel 界面功能
顶部菜单栏将新增 **【商业加壳插件演示】** 选项卡，包含以下分组和交互：

#### 分组一：数据写入操作
- **按钮：[插入当前时间]**
  - **功能**：在活动单元格中填入 `[测试数据] yyyy-MM-dd HH:mm:ss`。若勾选了自动高亮，单元格会自动加粗并设为淡黄背景。
- **按钮：[清空选中区域]**
  - **功能**：一键清空用户选中的单元格区域的内容与样式格式。

#### 分组二：交互与格式
- **复选框：[开启自动高亮]**
  - **功能**：控制后续数据写入时是否自动应用背景色与高亮格式（实时与 C# 后端属性同步）。
- **编辑框：[自定义内容:]**
  - **功能**：提供文本输入区（默认包含示例字符串），用户可随意修改输入文本。
- **按钮：[批量填充文本]**
  - **功能**：将编辑框中输入的自定义字符串批量填充到当前框选的所有 Excel 单元格中。

---

## 规范说明

- 所有 C# 源码均严格遵照要求：**至少每 3 行新增代码配备详细中文注释**。
- 加壳关键类与方法已添加标准 `[Obfuscation(Exclude = true)]` 特性标注。
