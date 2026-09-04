using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json.Serialization;

namespace ExcelAddInDemo.Services
{
    /// <summary>
    /// DWG 文件基本信息实体模型
    /// 遵循规范：每 3 行代码至少包含 1 行中文注释，配置与硬编码显式标明
    /// </summary>
    public class DwgFileInfo
    {
        // DWG 文件名 (包含 .dwg 后缀)
        [JsonPropertyName("fileName")]
        public string FileName { get; set; } = string.Empty;

        // 不含扩展名的纯图纸名称 (如 "BDY" 或 "风机2DY")
        [JsonPropertyName("nameWithoutExt")]
        public string NameWithoutExt { get; set; } = string.Empty;

        // 物理全路径
        [JsonPropertyName("fullPath")]
        public string FullPath { get; set; } = string.Empty;

        // 文件大小 (字节数)
        [JsonPropertyName("fileSize")]
        public long FileSize { get; set; } = 0;

        // 格式化后的友好文件大小 (如 "1.25 MB")
        [JsonPropertyName("fileSizeFormatted")]
        public string FileSizeFormatted { get; set; } = string.Empty;

        // 最后修改时间字符串
        [JsonPropertyName("lastModified")]
        public string LastModified { get; set; } = string.Empty;
    }

    /// <summary>
    /// DWG 缩略图预览结果响应实体
    /// </summary>
    public class DwgPreviewResult
    {
        // 是否提取成功
        [JsonPropertyName("success")]
        public bool Success { get; set; } = false;

        // Base64 图片 Data URL 字符串 (如 "data:image/png;base64,...")
        [JsonPropertyName("base64Image")]
        public string Base64Image { get; set; } = string.Empty;

        // 异常或错误提示消息
        [JsonPropertyName("errorMessage")]
        public string ErrorMessage { get; set; } = string.Empty;

        // 预览的图纸名称
        [JsonPropertyName("fileName")]
        public string FileName { get; set; } = string.Empty;

        // 物理路径
        [JsonPropertyName("fullPath")]
        public string FullPath { get; set; } = string.Empty;

        // 格式化文件大小
        [JsonPropertyName("fileSizeFormatted")]
        public string FileSizeFormatted { get; set; } = string.Empty;

        // 最后修改时间
        [JsonPropertyName("lastModifiedFormatted")]
        public string LastModifiedFormatted { get; set; } = string.Empty;

        // 是否为系统内置的占位图 (当图纸内未包含内嵌缩略图时为 true)
        [JsonPropertyName("isPlaceholder")]
        public bool IsPlaceholder { get; set; } = false;
    }

    /// <summary>
    /// 子文件夹描述实体模型 (支持下钻展开)
    /// </summary>
    public class DwgDirectoryItem
    {
        // 文件夹名称
        [JsonPropertyName("name")]
        public string Name { get; set; } = string.Empty;

        // 文件夹完整物理路径
        [JsonPropertyName("fullPath")]
        public string FullPath { get; set; } = string.Empty;

        // 是否包含子项 (子文件夹或图纸)
        [JsonPropertyName("hasChildren")]
        public bool HasChildren { get; set; } = false;
    }

    /// <summary>
    /// 目录层级扫描综合结果模型
    /// </summary>
    public class DirectoryHierarchyResult
    {
        // 当前扫描目录的物理路径
        [JsonPropertyName("currentPath")]
        public string CurrentPath { get; set; } = string.Empty;

        // 父级目录物理路径 (便于前端一键返回上层)
        [JsonPropertyName("parentPath")]
        public string? ParentPath { get; set; } = null;

        // 当前目录下的子文件夹列表
        [JsonPropertyName("subDirectories")]
        public List<DwgDirectoryItem> SubDirectories { get; set; } = new List<DwgDirectoryItem>();

        // 当前目录下的 DWG 图纸文件列表
        [JsonPropertyName("dwgFiles")]
        public List<DwgFileInfo> DwgFiles { get; set; } = new List<DwgFileInfo>();
    }

    /// <summary>
    /// DWG 图纸扫描、缩略图提取与外部 CAD 调起专属服务类
    /// 核心特色：采用 Windows Shell 原生高清缩略图 + DWG 二进制 Header 提取双通道架构
    /// </summary>
    public static class DwgPreviewService
    {
        // ----------------- Windows Shell 原生 COM 互操作定义 -----------------
        // 导入 Windows 原生 SHCreateItemFromParsingName 接口函数
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out IShellItemImageFactory ppv);

        // 定义 Windows 经典尺寸结构体
        [StructLayout(LayoutKind.Sequential)]
        private struct NativeSize
        {
            public int Width;
            public int Height;
            public NativeSize(int width, int height)
            {
                Width = width;
                Height = height;
            }
        }

        // IShellItemImageFactory 获取图像标志位枚举
        [Flags]
        private enum SIIGBF
        {
            SIIGBF_RESIZETOFIT = 0x00,
            SIIGBF_BIGGERSIZEOK = 0x01,
            SIIGBF_MEMORYONLY = 0x02,
            SIIGBF_ICONONLY = 0x04,
            SIIGBF_THUMBNAILONLY = 0x08,
            SIIGBF_INCACHEONLY = 0x10
        }

        // IShellItemImageFactory 接口 COM 契约声明 (GUID: bcc18b79-ba16-442f-80c4-8a59c30c463b)
        [ComImport]
        [Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItemImageFactory
        {
            [PreserveSig]
            int GetImage(
                [In, MarshalAs(UnmanagedType.Struct)] NativeSize size,
                [In] SIIGBF flags,
                [Out] out IntPtr phbm);
        }

        // 导入 GDI 释放 HBITMAP 句柄原生函数
        [DllImport("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeleteObject(IntPtr hObject);

        /// <summary>
        /// 扫描指定本地目录下的所有 DWG 图纸文件
        /// </summary>
        /// <param name="directoryPath">本地文件夹绝对路径</param>
        /// <returns>DWG 文件信息列表</returns>
        public static List<DwgFileInfo> ScanDwgFiles(string directoryPath)
        {
            var resultList = new List<DwgFileInfo>();
            // 路径有效性与存在性校验
            if (string.IsNullOrWhiteSpace(directoryPath) || !Directory.Exists(directoryPath))
            {
                return resultList;
            }

            try
            {
                // 获取当前目录及一级子目录下的所有 .dwg 文件 (仅扫描顶级目录以确保响应极速)
                var filePaths = Directory.GetFiles(directoryPath, "*.dwg", SearchOption.TopDirectoryOnly);
                // 遍历每一个 DWG 文件实体
                foreach (var path in filePaths)
                {
                    try
                    {
                        var fileInfo = new FileInfo(path);
                        // 忽略隐藏或系统临时文件 (如以 ~$ 开头的临时锁文件)
                        if (fileInfo.Name.StartsWith("~$") || (fileInfo.Attributes & FileAttributes.Hidden) != 0)
                        {
                            continue;
                        }

                        // 组装 DWG 文件描述实体
                        var item = new DwgFileInfo
                        {
                            FileName = fileInfo.Name,
                            NameWithoutExt = Path.GetFileNameWithoutExtension(fileInfo.Name),
                            FullPath = fileInfo.FullName,
                            FileSize = fileInfo.Length,
                            FileSizeFormatted = FormatFileSize(fileInfo.Length),
                            LastModified = fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
                        };
                        resultList.Add(item);
                    }
                    catch (Exception ex)
                    {
                        // 记录单项读取异常但不中断整个扫描循环
                        LogHelper.WriteLog($"[DwgPreviewService] 读取文件信息异常 {path}: {ex.Message}");
                    }
                }

                // 按照文件名称自然拼音升序排列
                resultList.Sort((a, b) => string.Compare(a.FileName, b.FileName, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                // 记录目录扫描整体异常
                LogHelper.WriteLog($"[DwgPreviewService] ScanDwgFiles 异常: {ex.Message}");
            }

            return resultList;
        }

        /// <summary>
        /// 扫描指定目录下的子文件夹与 DWG 图纸文件，支持层级下钻与面包屑溯源
        /// </summary>
        /// <param name="dirPath">目标物理路径</param>
        /// <returns>包含子文件夹与图纸文件的综合数据模型</returns>
        public static DirectoryHierarchyResult ScanDirectoryHierarchy(string dirPath)
        {
            // 初始化返回实体并记录当前扫描路径
            var result = new DirectoryHierarchyResult { CurrentPath = dirPath };
            if (string.IsNullOrWhiteSpace(dirPath) || !Directory.Exists(dirPath))
            {
                // 路径无效直接返回空结构
                return result;
            }

            try
            {
                // 获取父级目录路径以支持返回上一级
                var parentDir = Directory.GetParent(dirPath);
                result.ParentPath = parentDir?.FullName;

                // 枚举当前目录下的所有直接子文件夹
                var subDirs = Directory.GetDirectories(dirPath);
                // 按文件夹名称字母顺序排序
                Array.Sort(subDirs, StringComparer.OrdinalIgnoreCase);

                foreach (var sub in subDirs)
                {
                    try
                    {
                        var di = new DirectoryInfo(sub);
                        // 过滤操作系统隐藏文件夹
                        if ((di.Attributes & FileAttributes.Hidden) != 0) continue;

                        // 探测该子目录下是否包含内容
                        bool hasSub = false;
                        try
                        {
                            // 尝试嗅探子文件夹或文件实体
                            hasSub = Directory.EnumerateFileSystemEntries(sub).GetEnumerator().MoveNext();
                        }
                        catch
                        {
                            // 忽略无权限探测异常
                            hasSub = false;
                        }

                        // 组装子目录描述项
                        result.SubDirectories.Add(new DwgDirectoryItem
                        {
                            Name = di.Name,
                            FullPath = di.FullName,
                            HasChildren = hasSub
                        });
                    }
                    catch (Exception exSub)
                    {
                        // 记录单个子目录读取异常
                        LogHelper.WriteLog($"[DwgPreviewService] 枚举子目录异常 {sub}: {exSub.Message}");
                    }
                }

                // 复用扫描当前目录下的所有 DWG 图纸文件
                result.DwgFiles = ScanDwgFiles(dirPath);
            }
            catch (Exception ex)
            {
                // 记录层级枚举整体异常
                LogHelper.WriteLog($"[DwgPreviewService] ScanDirectoryHierarchy 异常: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 全局递归定位指定回路图号或文件名的具体物理路径及其所在父文件夹
        /// 用于前端点击已绑定元件组时跨目录穿透反显图纸
        /// </summary>
        /// <param name="rootDir">搜索根目录 (通常为回路图纸根目录)</param>
        /// <param name="codeOrName">图号或文件名 (如 接触器变频器 或 接触器变频器.dwg)</param>
        /// <returns>(success, parentDir, fileInfo, message)</returns>
        public static (bool success, string parentDir, DwgFileInfo? fileInfo, string message) LocateDwgInDirectory(string rootDir, string codeOrName)
        {
            // 校验根目录有效性与搜索关键词非空
            if (string.IsNullOrWhiteSpace(rootDir) || !Directory.Exists(rootDir) || string.IsNullOrWhiteSpace(codeOrName))
            {
                return (false, string.Empty, null, "搜索根目录不存在或目标图号为空");
            }

            try
            {
                // 清理目标图号两端空白
                string targetClean = codeOrName.Trim();
                // 提取不带扩展名的纯图号
                string targetNoExt = Path.GetFileNameWithoutExtension(targetClean);

                // 全局枚举搜索所有 .dwg 文件 (支持所有层级递归遍历)
                var allDwgFiles = Directory.EnumerateFiles(rootDir, "*.dwg", SearchOption.AllDirectories);
                foreach (var file in allDwgFiles)
                {
                    // 提取当前物理文件的文件名
                    string fName = Path.GetFileName(file);
                    // 提取当前物理文件不带扩展名名称
                    string fNoExt = Path.GetFileNameWithoutExtension(file);

                    // 忽略隐藏或临时锁定的 dwg 文件
                    if (fName.StartsWith("~$")) continue;

                    // 匹配文件名或不带扩展名的名称 (忽略大小写)
                    if (string.Equals(fNoExt, targetNoExt, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(fName, targetClean, StringComparison.OrdinalIgnoreCase))
                    {
                        // 组装文件描述实体
                        var fi = new FileInfo(file);
                        var item = new DwgFileInfo
                        {
                            FileName = fi.Name,
                            NameWithoutExt = fNoExt,
                            FullPath = fi.FullName,
                            FileSize = fi.Length,
                            FileSizeFormatted = FormatFileSize(fi.Length),
                            LastModified = fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
                        };
                        // 获取父目录绝对路径
                        string parentPath = fi.DirectoryName ?? rootDir;
                        // 成功返回定位结果与实体
                        return (true, parentPath, item, "定位成功");
                    }
                }

                // 未匹配到图纸实体
                return (false, string.Empty, null, $"在图纸库目录中未检索到图纸: {codeOrName}");
            }
            catch (Exception ex)
            {
                // 记录检索异常日志
                LogHelper.WriteLog($"[DwgPreviewService] LocateDwgInDirectory 异常: {ex.Message}");
                return (false, string.Empty, null, $"检索异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 读取目标 DWG 文件的原始二进制数据并转换为 Base64 字符串
        /// 供前端 WebGL 矢量渲染器直接实例化 File/Blob 展开渲染
        /// </summary>
        /// <param name="filePath">DWG 文件物理全路径</param>
        /// <returns>(success: 是否成功, base64: 数据内容, error: 错误提示)</returns>
        public static (bool success, string base64, string error) ReadDwgBinaryBase64(string filePath)
        {
            // 校验文件路径有效性
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return (false, string.Empty, "图纸物理文件不存在或路径无效");
            }

            try
            {
                // 读取二进制字节数组 (大文件安全读取)
                byte[] bytes = File.ReadAllBytes(filePath);
                // 转换为 Base64 数据串
                string base64 = Convert.ToBase64String(bytes);
                return (true, base64, string.Empty);
            }
            catch (Exception ex)
            {
                // 记录读取异常
                LogHelper.WriteLog($"[DwgPreviewService] ReadDwgBinaryBase64 异常 {filePath}: {ex.Message}");
                return (false, string.Empty, ex.Message);
            }
        }

        /// <summary>
        /// 获取指定 DWG 文件的缩略图与详情预览 (双通道提取机制)
        /// </summary>
        /// <param name="filePath">DWG 物理文件绝对路径</param>
        /// <returns>预览结果实体</returns>
        public static DwgPreviewResult GetDwgPreview(string filePath)
        {
            var result = new DwgPreviewResult
            {
                FullPath = filePath,
                FileName = Path.GetFileName(filePath)
            };

            // 检查文件物理存在性
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                result.ErrorMessage = "指定的 DWG 文件不存在或已被移动！";
                return result;
            }

            try
            {
                var fi = new FileInfo(filePath);
                result.FileSizeFormatted = FormatFileSize(fi.Length);
                result.LastModifiedFormatted = fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss");

                // 通道 1：尝试通过 Windows 原生 Shell 提取系统缓存的高清 CAD 缩略图
                Bitmap? shellBitmap = TryExtractShellThumbnail(filePath, 320, 240);
                if (shellBitmap != null)
                {
                    using (shellBitmap)
                    {
                        result.Base64Image = BitmapToBase64DataUrl(shellBitmap);
                        result.Success = true;
                        result.IsPlaceholder = false;
                        return result;
                    }
                }

                // 通道 2：若 Shell 提取失败，尝试直接从 DWG 头部二进制流解析内嵌 BMP
                Bitmap? headerBitmap = TryExtractDwgHeaderBitmap(filePath);
                if (headerBitmap != null)
                {
                    using (headerBitmap)
                    {
                        result.Base64Image = BitmapToBase64DataUrl(headerBitmap);
                        result.Success = true;
                        result.IsPlaceholder = false;
                        return result;
                    }
                }

                // 通道 3：兜底机制，生成专业的 CAD 占位图
                using (var placeholderBitmap = GenerateCadPlaceholderBitmap(result.FileName, result.FileSizeFormatted))
                {
                    result.Base64Image = BitmapToBase64DataUrl(placeholderBitmap);
                    result.Success = true;
                    result.IsPlaceholder = true;
                    result.ErrorMessage = "该 DWG 文件未内嵌缩略图，可直接点击【在 CAD 中打开】查看详细图纸。";
                }
            }
            catch (Exception ex)
            {
                // 记录提取异常
                LogHelper.WriteLog($"[DwgPreviewService] GetDwgPreview 异常: {ex.Message}");
                result.ErrorMessage = $"提取图纸预览失败: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// 通道 1：调用 Windows 原生 IShellItemImageFactory 接口提取高清缩略图
        /// </summary>
        private static Bitmap? TryExtractShellThumbnail(string filePath, int width, int height)
        {
            try
            {
                // IShellItemImageFactory 接口 GUID
                Guid shellItemImageFactoryGuid = new Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b");
                // 调用 Windows API 创建 ShellItem
                int hr = SHCreateItemFromParsingName(filePath, IntPtr.Zero, ref shellItemImageFactoryGuid, out var factory);
                if (hr != 0 || factory == null)
                {
                    return null;
                }

                // 目标请求位图几何尺寸 (如 320x240) --硬编码-- 请求缩略图分辨率
                var nativeSize = new NativeSize(width, height);
                // 优先请求高清缩略图模式
                hr = factory.GetImage(nativeSize, SIIGBF.SIIGBF_BIGGERSIZEOK | SIIGBF.SIIGBF_THUMBNAILONLY, out IntPtr hBitmap);
                if (hr != 0 || hBitmap == IntPtr.Zero)
                {
                    // 若无严格缩略图，尝试允许缩放到适应尺寸
                    hr = factory.GetImage(nativeSize, SIIGBF.SIIGBF_RESIZETOFIT, out hBitmap);
                }

                if (hr == 0 && hBitmap != IntPtr.Zero)
                {
                    try
                    {
                        // 从 GDI HBITMAP 句柄创建托管 Bitmap 实例
                        using (var tempBmp = Image.FromHbitmap(hBitmap))
                        {
                            // 深度克隆脱离对原始 GDI 句柄的依赖
                            return new Bitmap(tempBmp);
                        }
                    }
                    finally
                    {
                        // 及时释放 GDI 非托管位图句柄以防内存泄露
                        DeleteObject(hBitmap);
                    }
                }
            }
            catch
            {
                // Shell 提取异常直接降级
            }

            return null;
        }

        /// <summary>
        /// 通道 2：纯二进制解析 DWG 文件头提取内嵌 BMP 缩略图
        /// 支持从 AutoCAD R13 (AC1012) 到 2018 (AC1030) 标准 DWG 头部
        /// </summary>
        private static Bitmap? TryExtractDwgHeaderBitmap(string filePath)
        {
            try
            {
                using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var br = new BinaryReader(fs);

                // 读取 DWG 前 6 字节版本签名 (如 "AC1027")
                byte[] versionBytes = br.ReadBytes(6);
                string version = System.Text.Encoding.ASCII.GetString(versionBytes);
                // 校验是否为合法的 AutoCAD DWG 签名
                if (!version.StartsWith("AC")) return null;

                // 定位到偏移量 0x0D (13) 处读取 4 字节的图像定位指针 (IMAGE_SEEK)
                fs.Seek(0x0D, SeekOrigin.Begin);
                int imageSeek = br.ReadInt32();
                if (imageSeek <= 0 || imageSeek >= fs.Length) return null;

                // 跳转至缩略图数据块起始位置
                fs.Seek(imageSeek, SeekOrigin.Begin);
                // 读取图像格式类型标记
                byte imageType = br.ReadByte();
                // 常见 1 或 2 表示内嵌 BMP
                if (imageType != 1 && imageType != 2 && imageType != 3)
                {
                    // 尝试向下探测 BMP 头部特征 "BM" (0x42 0x4D)
                    byte[] searchBuf = br.ReadBytes(256);
                    for (int i = 0; i < searchBuf.Length - 1; i++)
                    {
                        if (searchBuf[i] == 0x42 && searchBuf[i + 1] == 0x4D)
                        {
                            fs.Seek(imageSeek + 1 + i, SeekOrigin.Begin);
                            return new Bitmap(fs);
                        }
                    }
                    return null;
                }

                // 读取图像数据字节总长度
                int imageLen = br.ReadInt32();
                if (imageLen <= 14 || imageLen > 10 * 1024 * 1024) return null;

                // 读取完整的 BMP 二进制流
                byte[] bmpBytes = br.ReadBytes(imageLen);
                using var ms = new MemoryStream(bmpBytes);
                return new Bitmap(ms);
            }
            catch
            {
                // 二进制解析异常安全返回 null
                return null;
            }
        }

        /// <summary>
        /// 通道 3：动态生成优雅的 CAD 图纸占位图 (黑绿经典工业风格)
        /// </summary>
        private static Bitmap GenerateCadPlaceholderBitmap(string fileName, string fileSize)
        {
            // 创建 360x240 分辨率的标准预览底图 --硬编码-- 占位图基础规格
            var bmp = new Bitmap(360, 240);
            using (var g = Graphics.FromImage(bmp))
            {
                // 启用高质量抗锯齿排版渲染
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                // 填充经典 CAD 黑色深沉背景色 (#1e293b)
                g.Clear(Color.FromArgb(30, 41, 59));

                // 绘制 CAD 网格坐标参考浅线
                using (var gridPen = new Pen(Color.FromArgb(45, 58, 80), 1))
                {
                    for (int x = 20; x < 360; x += 30) g.DrawLine(gridPen, x, 0, x, 240);
                    for (int y = 20; y < 240; y += 30) g.DrawLine(gridPen, 0, y, 360, y);
                }

                // 绘制中心 CAD 风格矩形与十字交叉定位光标
                using (var borderPen = new Pen(Color.FromArgb(0, 150, 136), 2)) // #009688 蓝绿色
                {
                    g.DrawRectangle(borderPen, 40, 30, 280, 180);
                    // 绘制四角定位折角标
                    g.DrawLine(borderPen, 35, 30, 45, 30);
                    g.DrawLine(borderPen, 315, 30, 325, 30);
                    g.DrawLine(borderPen, 35, 210, 45, 210);
                    g.DrawLine(borderPen, 315, 210, 325, 210);
                }

                // 绘制图纸大标文字 "AutoCAD DWG"
                using (var fontTitle = new Font("Segoe UI", 16, FontStyle.Bold))
                using (var brushTitle = new SolidBrush(Color.FromArgb(241, 245, 249)))
                {
                    var sf = new StringFormat { Alignment = StringAlignment.Center };
                    g.DrawString("📐 AutoCAD 图纸", fontTitle, brushTitle, new PointF(180, 55), sf);
                }

                // 绘制文件名 (超长自动截断)
                using (var fontFile = new Font("Microsoft YaHei UI", 11, FontStyle.Regular))
                using (var brushFile = new SolidBrush(Color.FromArgb(0, 188, 212))) // 青色
                {
                    var sf = new StringFormat { Alignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
                    var rect = new RectangleF(50, 95, 260, 40);
                    g.DrawString(fileName, fontFile, brushFile, rect, sf);
                }

                // 绘制文件大小提示
                using (var fontMeta = new Font("Segoe UI", 9, FontStyle.Regular))
                using (var brushMeta = new SolidBrush(Color.FromArgb(148, 163, 184)))
                {
                    var sf = new StringFormat { Alignment = StringAlignment.Center };
                    g.DrawString($"大小: {fileSize} | 未内嵌位图", fontMeta, brushMeta, new PointF(180, 140), sf);
                    g.DrawString("点击下方【在 CAD 中打开】即可查看完整图纸", fontMeta, brushMeta, new PointF(180, 160), sf);
                }
            }

            return bmp;
        }

        /// <summary>
        /// 将 Bitmap 位图编码转换为 base64 data url 字符串
        /// </summary>
        private static string BitmapToBase64DataUrl(Bitmap bitmap)
        {
            using var ms = new MemoryStream();
            // 保存为高保真 PNG 格式
            bitmap.Save(ms, ImageFormat.Png);
            byte[] bytes = ms.ToArray();
            return "data:image/png;base64," + Convert.ToBase64String(bytes);
        }

        /// <summary>
        /// 格式化文件字节数为友好字符串 (KB / MB)
        /// </summary>
        private static string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            if (bytes < 1024 * 1024) return $"{(bytes / 1024.0):F1} KB";
            return $"{(bytes / (1024.0 * 1024.0)):F2} MB";
        }

        /// <summary>
        /// 调用 Windows 默认程序直接打开 DWG 文件 (优先调起 AutoCAD 等关联软件)
        /// </summary>
        /// <param name="filePath">DWG 物理全路径</param>
        /// <returns>操作成功与否及提示信息</returns>
        public static (bool Success, string Message) OpenInDefaultCad(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return (false, "指定的 DWG 文件路径不存在，无法打开！");
            }

            try
            {
                // 配置启动信息以调起 Windows 关联程序
                var psi = new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                };
                Process.Start(psi);
                return (true, $"已调用系统默认程序打开: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[DwgPreviewService] OpenInDefaultCad 失败: {ex.Message}");
                return (false, $"调起外部 CAD 失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 在 Windows 资源管理器中高亮定位并选中该文件
        /// </summary>
        /// <param name="filePath">文件全路径</param>
        public static (bool Success, string Message) OpenInExplorer(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return (false, "文件不存在，无法在文件夹中定位！");
            }

            try
            {
                // 调用 explorer.exe 附带 /select 参数
                string args = $"/select,\"{filePath}\"";
                Process.Start("explorer.exe", args);
                return (true, "已在资源管理器中定位图纸！");
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog($"[DwgPreviewService] OpenInExplorer 异常: {ex.Message}");
                return (false, $"打开资源管理器失败: {ex.Message}");
            }
        }
    }
}
