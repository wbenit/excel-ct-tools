一、 商业加壳与代码安全注意事项 🔒
Ribbon 回调必须保留名称（Exclude Obfuscation）

规则：Excel 菜单通过 XML 字符串（如 onAction="MyMethod"）利用反射机制调用 C# 方法。
注意：如果你在 RibbonController.cs 中新增了任何 Ribbon 回调函数（如新按钮事件、下拉框选择事件），必须确保该类或方法拥有 [Obfuscation(Exclude = true)] 标记，否则加壳工具将其重命名为 a()、b() 后，Excel 会报错“找不到回调方法”。
最佳实践：将界面回调函数集中放在 RibbonController.cs，而将核心算法、高价值逻辑、公式计算、加密校验放在单独的业务类（如 ExcelServices.cs 或扩展 DLL）中，对业务类开启 100% 最高强度混淆。
禁止硬编码密钥与敏感配置

商业插件中不要将数据库连接串、API 密钥、授权 Token 等直接硬编码为常量字符串，加壳工具虽然能加密字符串，但依然存在被内存 Dump 风险。建议结合动态对称加密或后端 API 授权校验。