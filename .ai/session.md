一、 商业加壳与代码安全注意事项 🔒
Ribbon 回调必须保留名称（Exclude Obfuscation）

规则：Excel 菜单通过 XML 字符串（如 onAction="MyMethod"）利用反射机制调用 C# 方法。
一、 商业加壳与代码安全注意事项 🔒
Ribbon 回调必须保留名称（Exclude Obfuscation）

规则：Excel 菜单通过 XML 字符串（如 onAction="MyMethod"）利用反射机制调用 C# 方法。
注意：如果你在 RibbonController.cs 中新增了任何 Ribbon 回调函数（如新按钮事件、下拉框选择事件），必须确保该类或方法拥有 [Obfuscation(Exclude = true)] 标记，否则加壳工具将其重命名为 a()、b() 后，Excel 会报错“找不到回调方法”。
最佳实践：将界面回调函数集中放在 RibbonController.cs，而将核心算法、高价值逻辑、公式计算、加密校验放在单独的业务类（如 ExcelServices.cs 或扩展 DLL）中，对业务类开启 100% 最高强度混淆。
禁止硬编码密钥与敏感配置

商业插件中不要将数据库连接串、API 密钥、授权 Token 等直接硬编码为常量字符串，加壳工具虽然能加密字符串，但依然存在被内存 Dump 风险。建议结合动态对称加密或后端 API 授权校验。

## 二、 进度与状态记录
### 当前状态
- 已初始化本地 Git 仓库，并将代码成功推送到远程仓库 `https://github.com/wbenit/excel-ct-tools.git`。
- 实现了基于 C# (Excel-DNA) + WebView2 + Vue 3 (Element Plus, `<script setup>`) 的用户登录配置窗口。
  - 新增 `Resources/login.html`：采用 Vue 3 `<script setup>` 组合式 API + Element Plus 玻璃拟态主题。
  - 新增 `LoginForm.cs`：WinForms WebView2 容器并处理双向 JSON 消息。
  - 修复 `0x80070005 (E_ACCESSDENIED)` 拒绝访问异常：在 `LoginForm.cs` 中显式配置 `CoreWebView2Environment` 的 `userDataFolder` 为用户目录 `LocalApplicationData/ExcelAddInDemo/WebView2Data`，解决 Excel 试图向 Program Files 默认写入缓存导致的权限拒绝报错。
  - 新增 `Controllers/AuthController.cs`：.NET WebAPI 登录验证控制器。
  - 更新 `ExcelServices.cs` & `RibbonController.cs`：增加 Token 状态管理与 Ribbon 回调。
  - 新增 C# 代码严格遵循每 3 行包含 1 行中文注释的规范。
- **最新构建状态**：编译并成功重新打包（`0 个警告，0 个错误`）。
