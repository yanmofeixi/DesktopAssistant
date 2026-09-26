using System.ComponentModel;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace DesktopAssistant;

/// <summary>Image pixel (x, y) maps to desktop (Bounds.Left + x / Scale, Bounds.Top + y / Scale).</summary>
public sealed record ScreenshotCapture(byte[] Bytes, Rectangle Bounds, double Scale, CaptureWindow? Window, string? File)
{
    public Point ToDesktop(int x, int y) =>
        new(Bounds.Left + (int)Math.Round(x / Scale), Bounds.Top + (int)Math.Round(y / Scale));
}

public class ScreenshotManager
{
    public static readonly string SaveDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "DesktopAssistant", "Screenshots");
    public const int MaxRetainedCount = 10;
    private readonly object captureLock = new();

    /// <summary>The last image returned to a client; clicks with shot=1 use its coordinate mapping.</summary>
    public ScreenshotCapture? LastCapture { get; private set; }

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

    public ScreenshotCapture Capture(string? window = null, Rectangle? region = null, int? maxWidth = null,
        ImageFormat? format = null, bool save = false)
    {
        var capture = OnInputDesktop(() =>
        {
            CaptureWindow? target = string.IsNullOrWhiteSpace(window) ? null : WindowCapture.Select(window);
            Rectangle bounds = VisibleBounds(target, region);
            using var bitmap = CopyScreen(bounds);
            double scale = maxWidth is > 0 && bounds.Width > maxWidth ? (double)maxWidth.Value / bounds.Width : 1;
            using var output = scale < 1 ? Resize(bitmap, scale) : null;
            ImageFormat imageFormat = format ?? ImageFormat.Png;
            using var stream = new MemoryStream();
            (output ?? bitmap).Save(stream, imageFormat);
            byte[] bytes = stream.ToArray();
            string? path = save ? SaveToDisk(bytes, imageFormat) : null;
            return new ScreenshotCapture(bytes, bounds, scale, target, path);
        });
        LastCapture = capture;
        return capture;
    }

    /// <summary>A small grayscale thumbnail used to detect screen changes cheaply.</summary>
    public byte[] SampleGray(string? window, Rectangle? region)
    {
        return OnInputDesktop(() =>
        {
            CaptureWindow? target = string.IsNullOrWhiteSpace(window) ? null : WindowCapture.Select(window);
            using var bitmap = CopyScreen(VisibleBounds(target, region));
            using var thumbnail = new Bitmap(bitmap, new Size(64, Math.Max(1, 64 * bitmap.Height / bitmap.Width)));
            var gray = new byte[thumbnail.Width * thumbnail.Height];
            for (int y = 0; y < thumbnail.Height; y++)
                for (int x = 0; x < thumbnail.Width; x++)
                {
                    Color c = thumbnail.GetPixel(x, y);
                    gray[y * thumbnail.Width + x] = (byte)((c.R * 299 + c.G * 587 + c.B * 114) / 1000);
                }
            return gray;
        });
    }

    private static Rectangle VisibleBounds(CaptureWindow? target, Rectangle? region)
    {
        Rectangle bounds = SystemInformation.VirtualScreen;
        if (target != null) bounds = Rectangle.Intersect(bounds, target.Bounds);
        if (region.HasValue) bounds = Rectangle.Intersect(bounds, region.Value);
        if (bounds.Width <= 0 || bounds.Height <= 0)
            throw new WindowSelectionException("截图区域不在可见屏幕范围内。", 409);
        return bounds;
    }

    // Copy the visible desktop. Do not focus/restore a window or replace obscured content.
    private static Bitmap CopyScreen(Rectangle bounds)
    {
        var bitmap = new Bitmap(bounds.Width, bounds.Height, PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy);
        return bitmap;
    }

    private static Bitmap Resize(Bitmap source, double scale)
    {
        var size = new Size(Math.Max(1, (int)Math.Round(source.Width * scale)), Math.Max(1, (int)Math.Round(source.Height * scale)));
        var resized = new Bitmap(size.Width, size.Height, PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(resized);
        graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
        graphics.DrawImage(source, new Rectangle(Point.Empty, size));
        return resized;
    }

    private static string SaveToDisk(byte[] bytes, ImageFormat format)
    {
        Directory.CreateDirectory(SaveDirectory);
        string extension = format.Guid == ImageFormat.Jpeg.Guid ? "jpg" : "png";
        string path = Path.Combine(SaveDirectory, $"screenshot_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid():N}.{extension}");
        System.IO.File.WriteAllBytes(path, bytes);
        CleanOldScreenshots();
        return path;
    }

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
