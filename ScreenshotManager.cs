using System.ComponentModel;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace DesktopAssistant;

public sealed record ScreenshotCapture(byte[] Bytes, Rectangle Bounds, CaptureWindow? Window, string? File);

public class ScreenshotManager
{
    public static readonly string SaveDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "DesktopAssistant", "Screenshots");
    public const int MaxRetainedCount = 10;
    private readonly object captureLock = new();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr OpenInputDesktop(uint flags, bool inherit, uint access);
    [DllImport("user32.dll")] private static extern bool CloseDesktop(IntPtr desktop);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool SetThreadDesktop(IntPtr desktop);
    [DllImport("user32.dll")] private static extern IntPtr GetThreadDesktop(int threadId);
    [DllImport("kernel32.dll")] private static extern int GetCurrentThreadId();

    private T OnInputDesktop<T>(Func<T> action)
    {
        lock (captureLock)
        {
            IntPtr oldDpi = WindowCapture.SetThreadDpiAwarenessContext(new IntPtr(-4));
            IntPtr oldDesktop = GetThreadDesktop(GetCurrentThreadId());
            IntPtr inputDesktop = OpenInputDesktop(0, false, 0x01FF);
            bool switched = false;
            try
            {
                if (inputDesktop == IntPtr.Zero)
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "无法访问当前桌面，请确认 Windows 已解锁。");
                switched = SetThreadDesktop(inputDesktop);
                if (!switched)
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "无法切换到当前桌面。");
                return action();
            }
            finally
            {
                if (switched) SetThreadDesktop(oldDesktop);
                if (inputDesktop != IntPtr.Zero) CloseDesktop(inputDesktop);
                if (oldDpi != IntPtr.Zero) WindowCapture.SetThreadDpiAwarenessContext(oldDpi);
            }
        }
    }

    public void Start()
    {
        Directory.CreateDirectory(SaveDirectory);
        Logger.Info("ScreenshotManager 启动 (按需全屏/窗口截图)");
    }

    public void Stop() => Logger.Info("ScreenshotManager 停止");

    public IReadOnlyList<CaptureWindow> ListWindows() => OnInputDesktop(WindowCapture.List);

    public ScreenshotCapture Capture(string? window = null, ImageFormat? format = null, bool save = false)
    {
        return OnInputDesktop(() =>
        {
            Rectangle desktopBounds = SystemInformation.VirtualScreen;
            CaptureWindow? target = string.IsNullOrWhiteSpace(window) ? null : WindowCapture.Select(window);
            Rectangle bounds = target == null ? desktopBounds : Rectangle.Intersect(desktopBounds, target.Bounds);
            if (bounds.Width <= 0 || bounds.Height <= 0)
                throw new WindowSelectionException("目标窗口不在可见屏幕范围内。", 409);

            // Crop the visible desktop. Do not focus/restore a window or replace obscured content.
            using var bitmap = new Bitmap(bounds.Width, bounds.Height, PixelFormat.Format32bppArgb);
            using (var graphics = Graphics.FromImage(bitmap))
                graphics.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy);
            ImageFormat imageFormat = format ?? ImageFormat.Png;
            using var stream = new MemoryStream();
            bitmap.Save(stream, imageFormat);
            byte[] bytes = stream.ToArray();
            string? path = null;
            if (save)
            {
                Directory.CreateDirectory(SaveDirectory);
                string extension = imageFormat.Guid == ImageFormat.Jpeg.Guid ? "jpg" : "png";
                path = Path.Combine(SaveDirectory, $"screenshot_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid():N}.{extension}");
                System.IO.File.WriteAllBytes(path, bytes);
                CleanOldScreenshots();
                Logger.Debug($"已保存截图 ({bounds.Width}x{bounds.Height}): {path}");
            }
            return new ScreenshotCapture(bytes, bounds, target, path);
        });
    }

    public byte[]? CaptureInMemory(ImageFormat? format = null) => Capture(format: format).Bytes;
    public string? CaptureNow() => Capture(save: true).File;

    public static string? GetLatestScreenshotPath()
    {
        try
        {
            return Directory.Exists(SaveDirectory)
                ? new DirectoryInfo(SaveDirectory).GetFiles("screenshot_*")
                    .Where(f => f.Extension is ".png" or ".jpg")
                    .OrderByDescending(f => f.LastWriteTimeUtc).FirstOrDefault()?.FullName
                : null;
        }
        catch (Exception ex)
        {
            Logger.Warn($"获取最新截图路径失败: {ex.Message}");
            return null;
        }
    }

    private static void CleanOldScreenshots()
    {
        try
        {
            foreach (var file in new DirectoryInfo(SaveDirectory).GetFiles("screenshot_*")
                         .Where(f => f.Extension is ".png" or ".jpg")
                         .OrderByDescending(f => f.LastWriteTimeUtc).Skip(MaxRetainedCount))
            {
                try { file.Delete(); }
                catch (Exception ex) { Logger.Warn($"删除旧截图失败: {file.Name}, {ex.Message}"); }
            }
        }
        catch (Exception ex) { Logger.Warn($"清理旧截图失败: {ex.Message}"); }
    }

    public static void OpenScreenshotFolder()
    {
        try
        {
            Directory.CreateDirectory(SaveDirectory);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = SaveDirectory, UseShellExecute = true
            });
        }
        catch (Exception ex) { Logger.Error("打开截图目录失败", ex); }
    }
}
