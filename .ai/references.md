# 核心架构模式与参考实现 (Architectural References)

本文档记录项目中反复使用的关键架构模式与代码模板，供跨功能模块直接复用。

---

## 1. WebView2 模态对话框异步 STA 线程解耦标准模板

在 WebView2 + WinForms 架构中，严禁在 `WebMessageReceived` 回调中直接同步调用 `ShowDialog()`，否则会导致 Chromium IPC 死锁与界面冻结。统一采用以下标准模板解耦：

```csharp
// 【标准范式】：WebMessageReceived 中的模态对话框异步化
case "selectDirectory":
    var dialogThread = new System.Threading.Thread(() =>
    {
        try
        {
            // 实例化对话框 (注意: .NET Framework 4.8 中 FolderBrowserDialog 无 AutoUpgradeEnabled)
            using var dialog = new FolderBrowserDialog
            {
                Description = "请选择目标文件夹",
                ShowNewFolderButton = true
            };

            // 优先恢复历史记录路径
            if (!string.IsNullOrWhiteSpace(lastPath) && Directory.Exists(lastPath))
            {
                dialog.SelectedPath = lastPath;
            }

            // 在独立的后台 STA 线程中运行模态循环，彻底解耦 UI 主线程与 Chromium IPC
            if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
            {
                string chosen = dialog.SelectedPath;
                // 线程安全切回主线程通知前端
                SafeInvoke(() =>
                {
                    PostWebMessageSafe(JsonSerializer.Serialize(new
                    {
                        action = "directorySelected",
                        path = chosen
                    }, JsonOptions));
                });
            }
        }
        catch (Exception ex)
        {
            LogHelper.WriteLog($"[DialogThread] 对话框异常: {ex.Message}");
        }
    });

    // 必须设定为 STA 单元模式
    dialogThread.SetApartmentState(System.Threading.ApartmentState.STA);
    dialogThread.IsBackground = true;
    dialogThread.Start();
    break;
```

---

## 2. 前端双轨路径设定标准交互 (弹窗浏览 + 极速粘贴)

为了兼顾不同用户习惯并彻底消除对弹窗的单一依赖，前端采用“浏览 + 粘贴”双轨模式：

```html
<!-- 轨 1: 浏览选择 -->
<el-button size="small" type="primary" plain @click="selectDir">📁 浏览选择</el-button>

<!-- 轨 2: 复制粘贴 (直接读取剪贴板或 Prompt 粘贴) -->
<el-button size="small" type="info" plain @click="inputDirPrompt">✍️ 粘贴路径</el-button>
```

```javascript
// JS Prompt 标准处理
const inputDirPrompt = () => {
  ElementPlus.ElMessageBox.prompt(
    '请输入或从 Windows 资源管理器地址栏粘贴目标目录路径:',
    '设置目标目录',
    {
      confirmButtonText: '确定绑定',
      cancelButtonText: '取消',
      inputValue: currentDir.value || '',
      inputPlaceholder: '例如: D:\\Drawings\\Projects'
    }
  ).then(({ value }) => {
    if (value && value.trim()) {
      // 传递至后端 setManual 接口
      postToCSharp({
        action: 'setDirectoryManual',
        path: value.trim()
      });
    }
  }).catch(() => {});
};
```

---

## 3. 免 AutoCAD 依赖的轻量 DWG 缩略图提取梯队

```csharp
// 调度 DwgPreviewService
var preview = DwgPreviewService.GetPreview(filePath, preferredType);
// 返回模型包含：
// - Base64Image (data:image/bmp;base64,... 或 data:image/png;base64,...)
// - IsPlaceholder (是否降级占位图)
// - FileSizeFormatted / LastModifiedFormatted (文件元数据)
// - ErrorMessage (异常信息)
```
