using System.ComponentModel;
using System.Diagnostics;
using System.Text;

namespace DesktopAssistant;

/// <summary>
/// Runs processes inside the logged-in desktop session. SSH key logons cannot unlock the user's DPAPI
/// keys or reach the desktop, so Credential Manager (git/GCM, gh) and GUI launches only work from here.
/// </summary>
public static class ProcessRunner
{
    public sealed record RunOutcome(int? ExitCode, string Stdout, string Stderr, bool TimedOut, int Pid, double Seconds);

    public sealed record LaunchOutcome(int Pid, string Path, int SessionId, string? ProductName, string? FileVersion,
        IntPtr MainWindow, string? MainWindowTitle);

    static ProcessRunner()
    {
        // Console tools on Chinese Windows print in code page 936 unless they choose UTF-8 themselves.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public static Encoding ResolveEncoding(string? name)
    {
        if (string.IsNullOrEmpty(name) || name.Equals("utf-8", StringComparison.OrdinalIgnoreCase)
            || name.Equals("utf8", StringComparison.OrdinalIgnoreCase))
            return new UTF8Encoding(false);
        try
        {
            return name.Equals("oem", StringComparison.OrdinalIgnoreCase)
                ? Encoding.GetEncoding(System.Globalization.CultureInfo.CurrentCulture.TextInfo.OEMCodePage)
                : int.TryParse(name, out int codePage) ? Encoding.GetEncoding(codePage) : Encoding.GetEncoding(name);
        }
        catch (ArgumentException)
        {
            throw new CommandException($"不支持的编码 '{name}'，可用 utf-8、oem、gbk 或代码页数字。");
        }
    }

    public static RunOutcome Run(ProcessStartInfo startInfo, string? stdin, TimeSpan timeout, int maxChars)
    {
        startInfo.UseShellExecute = false;
        startInfo.CreateNoWindow = true;
        startInfo.RedirectStandardInput = true;
        startInfo.RedirectStandardOutput = true;
        startInfo.RedirectStandardError = true;

        var elapsed = Stopwatch.StartNew();
        using var process = StartOrThrow(startInfo);
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        // Closing stdin makes prompts fail fast instead of hanging until the timeout.
        if (stdin != null) process.StandardInput.Write(stdin);
        process.StandardInput.Close();

        bool exited = process.WaitForExit(timeout);
        if (!exited)
        {
            try { process.Kill(entireProcessTree: true); } catch (InvalidOperationException) { }
            process.WaitForExit(5000);
        }
        Task.WaitAll([stdoutTask, stderrTask], 5000);
        return new RunOutcome(
            exited ? process.ExitCode : null,
            KeepTail(stdoutTask.IsCompletedSuccessfully ? stdoutTask.Result : "", maxChars),
            KeepTail(stderrTask.IsCompletedSuccessfully ? stderrTask.Result : "", maxChars),
            !exited, process.Id, Math.Round(elapsed.Elapsed.TotalSeconds, 2));
    }

    public static LaunchOutcome Launch(ProcessStartInfo startInfo, TimeSpan windowTimeout)
    {
        string path = startInfo.FileName;
        if (!System.IO.Path.IsPathFullyQualified(path) || !File.Exists(path))
            throw new CommandException($"launch 只接受存在的可执行文件绝对路径，收到 '{path}'。");
        startInfo.UseShellExecute = false;
        if (string.IsNullOrEmpty(startInfo.WorkingDirectory))
            startInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(path)!;

        var process = StartOrThrow(startInfo);
        var version = FileVersionInfo.GetVersionInfo(path);
        IntPtr mainWindow = WaitForMainWindow(process, windowTimeout);
        return new LaunchOutcome(process.Id, ReadActualPath(process) ?? path, process.SessionId,
            NullIfEmpty(version.ProductName), NullIfEmpty(version.FileVersion),
            mainWindow, mainWindow == IntPtr.Zero ? null : process.MainWindowTitle);
    }

    private static Process StartOrThrow(ProcessStartInfo startInfo)
    {
        try
        {
            return Process.Start(startInfo) ?? throw new CommandException($"无法启动 '{startInfo.FileName}'。", 500);
        }
        catch (Win32Exception ex)
        {
            throw new CommandException($"无法启动 '{startInfo.FileName}'：{ex.Message}");
        }
    }

    private static IntPtr WaitForMainWindow(Process process, TimeSpan timeout)
    {
        var elapsed = Stopwatch.StartNew();
        while (elapsed.Elapsed < timeout)
        {
            process.Refresh();
            if (process.HasExited) return IntPtr.Zero;
            if (process.MainWindowHandle != IntPtr.Zero) return process.MainWindowHandle;
            Thread.Sleep(200);
        }
        return IntPtr.Zero;
    }

    private static string? ReadActualPath(Process process)
    {
        try { return process.MainModule?.FileName; }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException) { return null; }
    }

    private static string KeepTail(string text, int maxChars) =>
        // Errors usually appear at the end, so long output keeps its tail.
        text.Length <= maxChars ? text : "…(已截断开头)\n" + text[^maxChars..];

    private static string? NullIfEmpty(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;
}
