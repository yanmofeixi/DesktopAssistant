using System.Diagnostics;
using System.Net;
using System.Text;
using System.Text.Json;
using DesktopAssistant;

internal static class PackageProbe
{
    [STAThread]
    private static int Main(string[] args)
    {
        var report = new Dictionary<string, object?>();
        RemoteInputServer? server = null;
        string? screenshot = null;
        try
        {
            var expected = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DesktopAssistant", "Screenshots");
            if (ScreenshotManager.SaveDirectory != expected) throw new Exception("Screenshot directory is not user-specific.");
            report["screenshotDirectory"] = expected;
            var core = Process.GetCurrentProcess().Modules.Cast<ProcessModule>().Single(m => m.ModuleName == "coreclr.dll").FileName;
            if (!core.StartsWith(AppContext.BaseDirectory, StringComparison.OrdinalIgnoreCase)) throw new Exception("Runtime was loaded from outside the published directory.");
            report["privateRuntime"] = core;
            report["sessionId"] = Process.GetCurrentProcess().SessionId;
            var manager = new ScreenshotManager();
            manager.Start();
            server = new RemoteInputServer(manager, 18889);
            server.Start();
            if (!server.IsRunning) throw new Exception("Probe listener failed to start.");
            using var client = new HttpClient(new HttpClientHandler { UseProxy = false }) { Timeout = TimeSpan.FromSeconds(15) };
            var status = client.GetStringAsync("http://127.0.0.1:18889/status").GetAwaiter().GetResult();
            using var statusJson = JsonDocument.Parse(status);
            if (statusJson.RootElement.GetProperty("status").GetString() != "ok") throw new Exception("Status failed.");
            report["statusApi"] = "ok";
            foreach (var text in new[] { "", "a\0b" })
            {
                using var response = client.PostAsync("http://127.0.0.1:18889/paste", new StringContent(text, Encoding.UTF8, "text/plain")).GetAwaiter().GetResult();
                if (response.StatusCode != HttpStatusCode.BadRequest) throw new Exception("Invalid paste was not rejected.");
            }
            report["invalidPasteRejected"] = true;
            // Use the real HTTP request thread, as the product does. A GUI-owning STA
            // thread cannot call SetThreadDesktop after Windows has attached hooks.
            var shotText = client.GetStringAsync("http://127.0.0.1:18889/screenshot?save=true").GetAwaiter().GetResult();
            using var shotJson = JsonDocument.Parse(shotText);
            screenshot = shotJson.RootElement.GetProperty("file").GetString();
            if (screenshot == null || Path.GetDirectoryName(screenshot) != expected || !File.Exists(screenshot)) throw new Exception("Screenshot was not saved under user data.");
            var bytes = File.ReadAllBytes(screenshot);
            if (!bytes.AsSpan(0, 8).SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 })) throw new Exception("Invalid PNG.");
            using var image = Image.FromStream(new MemoryStream(bytes));
            if (image.Width != shotJson.RootElement.GetProperty("width").GetInt32() || image.Height != shotJson.RootElement.GetProperty("height").GetInt32()) throw new Exception("Screenshot dimensions mismatch.");
            report["screenshot"] = new { image.Width, image.Height, bytes = bytes.Length };
            report["passed"] = true;
            return 0;
        }
        catch (Exception ex) { report["passed"] = false; report["error"] = ex.ToString(); return 1; }
        finally
        {
            server?.Stop();
            if (screenshot != null && File.Exists(screenshot)) File.Delete(screenshot);
            File.WriteAllText(args[0], JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
        }
    }
}

// The probe exercises production capture/input code without deleting the live app's logs.
namespace DesktopAssistant
{
    public static class Logger
    {
        public static void Info(string message) { }
        public static void Debug(string message) { }
        public static void Warn(string message) { }
        public static void Error(string message, Exception? exception = null) { }
        public static void Log(string message) { }
    }
}
