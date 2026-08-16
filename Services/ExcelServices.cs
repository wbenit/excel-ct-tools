using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using static ExcelAddInDemo.Tool;

namespace ExcelAddInDemo
{
    /// <summary>
    /// Excel 核心业务服务分部类：基础与通用功能
    /// </summary>
    public static partial class ExcelServices
    {
        // 记录用户是否开启了自动高亮格式的全局变量
        private static bool _isAutoHighlightEnabled = true;

        // 缓存编辑框中用户输入的自定义字符串
        private static string _customMessageText = "来自鑫壬成套服务的示例消息";

        // 保存用户认证通过后的 Token 授权密钥串
        private static string _currentToken = string.Empty;

        // 保存当前登录用户的显示名称
        private static string _currentUserDisplayName = "未登录";

        /// <summary>
        /// 获取或设置授权 Token 密钥串
        /// </summary>
        public static string CurrentToken
        {
            // 读取最新的 Token 凭据
            get => _currentToken;
            // 更新并保存 Token 凭据
            set => _currentToken = value ?? string.Empty;
        }

        /// <summary>
        /// 获取或设置当前登录用户显示的名称
        /// </summary>
        public static string CurrentUserDisplayName
        {
            // 读取用户显示的名称
            get => _currentUserDisplayName;
            // 更新用户显示名称
            set => _currentUserDisplayName = value ?? string.Empty;
        }

        /// <summary>
        /// Excel 主窗口 Win32 句柄包装类
        /// </summary>
        public class ExcelWin32Window : System.Windows.Forms.IWin32Window
        {
            // 存储 Win32 句柄
            public IntPtr Handle { get; }

            // 构造函数初始化句柄
            public ExcelWin32Window(IntPtr handle) => Handle = handle;
        }

        // 登录配置窗口静态单例引用 (可空)
        private static LoginForm? _loginForm;

        // 企业设置窗口静态单例引用 (可空)
        private static EnterpriseSettingsForm? _enterpriseSettingsForm;

        /// <summary>
        /// 安全展示非模态窗体，保证弹出时 Excel 依然处于可编辑交互状态
        /// </summary>
        /// <typeparam name="T">窗体类型</typeparam>
        /// <param name="formInstance">静态窗体引用</param>
        /// <param name="factory">窗体实例化工厂</param>
        public static void ShowModelessForm<T>(ref T? formInstance, Func<T> factory) where T : System.Windows.Forms.Form
        {
            // 启用 Windows 窗体视觉样式效果
            System.Windows.Forms.Application.EnableVisualStyles();

            // 若窗体尚未实例化或已被释放，则创建新实例
            if (formInstance == null || formInstance.IsDisposed)
            {
                // 创建窗体实例
                formInstance = factory();

                // 获取 Excel 主窗口 HWND 句柄
                IntPtr excelHwnd = ExcelDnaUtil.WindowHandle;

                // 判断句柄有效性并以非模态方式展示，保持 Excel 可直接编辑
                if (excelHwnd != IntPtr.Zero)
                {
                    // 设置 Owner 为 Excel 主窗口，防止沉入后台，且不阻塞 Excel 操作
                    formInstance.Show(new ExcelWin32Window(excelHwnd));
                }
                else
                {
                    // 无有效主句柄时独立非模态弹出
                    formInstance.Show();
                }
            }
            else
            {
                // 若窗体已处于最小化状态则恢复正常大小
                if (formInstance.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                {
                    // 还原窗体
                    formInstance.WindowState = System.Windows.Forms.FormWindowState.Normal;
                }

                // 将窗体推至最前
                formInstance.BringToFront();

                // 激活窗体获得焦点
                formInstance.Activate();
            }
        }

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的登录配置窗口 (非模态，可编辑 Excel)
        /// </summary>
        public static void ShowLoginDialog()
        {
            try
            {
                // 以非模态方式展示登录配置窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _loginForm, () => new LoginForm());
            }
            catch (Exception ex)
            {
                // 捕获弹窗异常防止 Excel 崩溃闪退
                System.Windows.Forms.MessageBox.Show($"弹出登录配置窗口发生异常: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 获取或设置自动高亮状态
        /// </summary>
        public static bool IsAutoHighlightEnabled
        {
            // 读取当前高亮开关状态
            get => _isAutoHighlightEnabled;
            // 写入并更新高亮开关状态
            set => _isAutoHighlightEnabled = value;
        }

        /// <summary>
        /// 获取或设置自定义消息文本
        /// </summary>
        public static string CustomMessageText
        {
            // 读取当前设置的文本内容
            get => _customMessageText;
            // 写入文本内容，防止为空
            set => _customMessageText = value ?? string.Empty;
        }

        /// <summary>
        /// 在 Excel 活动单元格写入时间戳和示例文本
        /// </summary>
        public static void InsertTimestampAndData()
        {
            // 获取 Excel 的 COM Application 对象
            dynamic app = ExcelDnaUtil.Application;
            // 获取当前选中的活动单元格
            dynamic activeCell = app.ActiveCell;

            // 校验单元格是否有效
            if (activeCell == null) return;

            // 写入包含当前系统精确时间的测试数据
            activeCell.Value2 = $"[测试数据] {DateTime.Now:yyyy-MM-dd HH:mm:ss}";

            // 判断是否启用了自动高亮选项
            if (_isAutoHighlightEnabled)
            {
                // 将单元格背景颜色设置为淡黄色
                activeCell.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 255, 204));
                // 将单元格文字字体加粗显示
                activeCell.Font.Bold = true;
            }
        }

        /// <summary>
        /// 清除选定区域的所有内容与格式
        /// </summary>
        public static void ClearActiveRange()
        {
            // 获取 Excel Application COM 引用
            dynamic app = ExcelDnaUtil.Application;
            // 获取当前用户框选的单元格区域
            dynamic selection = app.Selection;

            // 校验选中区域是否存在
            if (selection == null) return;

            // 调用 Excel 原生 API 清空数据与格式
            selection.Clear();
        }

        /// <summary>
        /// 将 EditBox 输入框的文本批量赋值给选中区域
        /// </summary>
        public static void ApplyCustomTextToSelection()
        {
            // 获取全局 Excel Application 实例
            dynamic app = ExcelDnaUtil.Application;
            // 获取选中的 Range 单元格集合
            dynamic selection = app.Selection;

            // 判断选中对象有效性
            if (selection == null) return;

            // 将用户输入的值批量赋值到选中的每个单元格
            selection.Value2 = _customMessageText;

            // 依据复选框状态决定是否应用突出格式
            if (_isAutoHighlightEnabled)
            {
                // 设置单元格填充背景为天蓝色
                selection.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(204, 229, 255));
                // 设置字体颜色为深蓝色以增强视效
                selection.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 51, 102));
            }
        }

        /// <summary>
        /// 启动并弹出基于 WebView2 + Vue 3 的“我的企业设置”窗口 (非模态，可编辑 Excel)
        /// </summary>
        public static void ShowEnterpriseSettingsDialog()
        {
            try
            {
                // 以非模态方式展示企业设置窗口，保持 Excel 处于可交互编辑状态
                ShowModelessForm(ref _enterpriseSettingsForm, () => new EnterpriseSettingsForm());
            }
            catch (Exception ex)
            {
                // 全局捕获异常防止程序闪退
                System.Windows.Forms.MessageBox.Show($"弹出企业设置窗口失败: {ex.Message}", "系统提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 激活并强力置顶新创建的工作簿视口
        /// </summary>
        public static void ActivateCreatedWorkbook(string targetPath)
        {
            try
            {
                // 获取 Excel Application COM 接口
                dynamic app = ExcelDnaUtil.Application;
                if (app == null || string.IsNullOrEmpty(targetPath)) return;

                // 遍历已打开的工作簿
                foreach (dynamic wb in app.Workbooks)
                {
                    if (string.Equals(wb.FullName, targetPath, StringComparison.OrdinalIgnoreCase))
                    {
                        // 激活目标工作簿
                        wb.Activate();
                        // 最大化当前窗口视口
                        app.ActiveWindow.WindowState = -4137; // xlMaximized
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录激活异常日志
                LogHelper.WriteLog($"激活工作簿异常: {ex.Message}");
            }
        }
    }
}
