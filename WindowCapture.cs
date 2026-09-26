using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace DesktopAssistant;

public sealed record CaptureWindow(IntPtr Hwnd, string Title, string ProcessName, int ProcessId,
    bool Minimized, Rectangle Bounds)
{
    public object Metadata => new
    {
        handle = $"0x{Hwnd.ToInt64():X}", title = Title, process = ProcessName,
        processId = ProcessId, minimized = Minimized,
        left = Bounds.Left, top = Bounds.Top, width = Bounds.Width, height = Bounds.Height
    };

    public string Describe() => $"0x{Hwnd.ToInt64():X} {ProcessName} \"{Title}\" " +
        (Minimized ? "minimized" : $"{Bounds.Left},{Bounds.Top} {Bounds.Width}x{Bounds.Height}");
}

public sealed class WindowSelectionException(string message, int statusCode,
    IReadOnlyList<CaptureWindow>? matches = null) : Exception(message)
{
    public int StatusCode { get; } = statusCode;
    public IReadOnlyList<CaptureWindow> Matches { get; } = matches ?? [];
}

public static class WindowCapture
{
    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd, out NativeRect rect);
    [DllImport("dwmapi.dll", EntryPoint = "DwmGetWindowAttribute")]
    private static extern int GetFrameBounds(IntPtr hwnd, int attribute, out NativeRect rect, int size);
    [DllImport("dwmapi.dll", EntryPoint = "DwmGetWindowAttribute")]
    private static extern int GetCloaked(IntPtr hwnd, int attribute, out int cloaked, int size);
    [DllImport("user32.dll")] internal static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect { public int Left, Top, Right, Bottom; }

    // Callers run in a PerMonitorV2 thread context so these coordinates are physical pixels.
    public static IReadOnlyList<CaptureWindow> List()
    {
        var windows = new List<CaptureWindow>();
        EnumWindows((hwnd, _) =>
        {
            if (!IsWindowVisible(hwnd)) return true;
            if (GetCloaked(hwnd, 14, out int cloaked, sizeof(int)) == 0 && cloaked != 0) return true;
            var title = new StringBuilder(4096);
            if (GetWindowText(hwnd, title, title.Capacity) == 0) return true;
            bool minimized = IsIconic(hwnd);
            NativeRect rect;
            if (minimized || GetFrameBounds(hwnd, 9, out rect, Marshal.SizeOf<NativeRect>()) != 0)
            {
                if (!GetWindowRect(hwnd, out rect)) return true;
            }
            GetWindowThreadProcessId(hwnd, out uint processId);
            string processName = "";
            try
            {
                using var process = Process.GetProcessById((int)processId);
                processName = process.ProcessName;
            }
            catch (ArgumentException) { }
            catch (System.ComponentModel.Win32Exception) { }
            windows.Add(new CaptureWindow(hwnd, title.ToString(), processName, (int)processId,
                minimized, Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom)));
            return true;
        }, IntPtr.Zero);
        return windows;
    }

    public static CaptureWindow Select(string selector, bool allowMinimized = false)
    {
        selector = selector.Trim();
        var windows = List();
        CaptureWindow[] matches;
        if (selector.Equals("active", StringComparison.OrdinalIgnoreCase))
        {
            IntPtr foreground = GetForegroundWindow();
            matches = windows.Where(w => w.Hwnd == foreground).ToArray();
        }
        else if (selector.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
        {
            if (!long.TryParse(selector[2..], NumberStyles.AllowHexSpecifier,
                    CultureInfo.InvariantCulture, out long handle))
                throw new WindowSelectionException("窗口句柄必须是 0x 开头的十六进制数。", 400);
            matches = windows.Where(w => w.Hwnd.ToInt64() == handle).ToArray();
        }
        else
        {
            string processName = selector.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
                ? selector[..^4] : selector;
            matches = windows.Where(w => w.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length == 0)
                matches = windows.Where(w => w.Title.Equals(selector, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length == 0)
                matches = windows.Where(w => w.Title.Contains(selector, StringComparison.OrdinalIgnoreCase)).ToArray();
        }
        if (matches.Length == 0)
            throw new WindowSelectionException($"未找到窗口 '{selector}'；请用 /windows 查看可用窗口。", 404);
        if (matches.Length > 1)
            throw new WindowSelectionException($"'{selector}' 匹配多个窗口；请指定列表中的 handle。", 409, matches);
        if (matches[0].Minimized && !allowMinimized)
            throw new WindowSelectionException("目标窗口已最小化，请先恢复窗口。", 409, matches);
        return matches[0];
    }
}
