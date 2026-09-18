using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace DesktopAssistant;

/// <summary>Ctrl+F1 controls the upstairs study wall light through a dedicated HA webhook.</summary>
public sealed class StudyLightHotkey : Form
{
    private const int HotkeyId = 0x6301;
    private const int WmHotkey = 0x0312;
    private const uint ModControlNoRepeat = 0x0002 | 0x4000;
    private const uint VkF1 = 0x70;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint key);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr window, int id);

    private readonly HttpClient client = new(new HttpClientHandler
    {
        UseProxy = false,
        AllowAutoRedirect = false
    }) { Timeout = TimeSpan.FromSeconds(5) };

    private Uri? webhookUrl;
    private bool registered;
    private bool busy;
    private bool stopping;

    public event Action<string>? Error;

    public void Start()
    {
        if (registered || stopping) return;

        try
        {
            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DesktopAssistant", "HomeAssistant.json");
            using var config = JsonDocument.Parse(File.ReadAllText(path));
            var url = config.RootElement.GetProperty("WebhookUrl").GetString();
            if (!Uri.TryCreate(url, UriKind.Absolute, out webhookUrl)
                || (webhookUrl.Scheme != "http" && webhookUrl.Scheme != "https")
                || !webhookUrl.AbsolutePath.StartsWith("/api/webhook/", StringComparison.Ordinal))
            {
                throw new InvalidDataException("Invalid webhook configuration");
            }
        }
        catch (Exception ex)
        {
            // The URL contains a secret: do not include configuration or exception messages in logs.
            ReportError($"壁灯快捷键配置读取失败（{ex.GetType().Name}）。");
            return;
        }

        registered = RegisterHotKey(Handle, HotkeyId, ModControlNoRepeat, VkF1);
        if (!registered)
        {
            var error = Marshal.GetLastWin32Error();
            ReportError(error == 1409
                ? "Ctrl+F1 已被其他程序占用，壁灯快捷键未启用。"
                : $"无法注册壁灯快捷键 Ctrl+F1（错误 {error}）。");
            return;
        }

        Logger.Info("书房壁灯快捷键 Ctrl+F1 已注册（禁止长按重复触发）。");
    }

    protected override void WndProc(ref Message message)
    {
        if (message.Msg == WmHotkey && message.WParam.ToInt32() == HotkeyId)
        {
            if (!busy && !stopping) _ = ToggleAsync();
            return;
        }

        base.WndProc(ref message);
    }

    private async Task ToggleAsync()
    {
        busy = true;
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, webhookUrl);
            using var response = await client.SendAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                ReportError($"壁灯请求失败（HTTP {(int)response.StatusCode}），请检查 Home Assistant。");
                return;
            }

            // A webhook acknowledges receipt; HA makes the state decision and controls brightness.
            Logger.Info("已发送书房壁灯切换请求（开灯亮度 100%）。");
        }
        catch (OperationCanceledException) when (stopping) { }
        catch (OperationCanceledException)
        {
            ReportError("壁灯请求超时，结果未确认；请查看灯光状态后再按一次。");
        }
        catch (Exception ex)
        {
            if (!stopping) ReportError($"无法连接 Home Assistant（{ex.GetType().Name}）。");
        }
        finally
        {
            // No retries: repeating a toggle after a lost response could undo the first request.
            await Task.Delay(500);
            busy = false;
        }
    }

    private void ReportError(string message)
    {
        Logger.Warn(message);
        if (!stopping) Error?.Invoke(message);
    }

    public void Stop()
    {
        if (stopping) return;
        stopping = true;
        if (registered)
        {
            UnregisterHotKey(Handle, HotkeyId);
            registered = false;
        }
        client.Dispose();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) Stop();
        base.Dispose(disposing);
    }
}
