using System.Runtime.InteropServices;
using System.Text.Json;

namespace DesktopAssistant;

/// <summary>某一时刻的用户活动判断结果。</summary>
public sealed record UserActivityState(bool Active, double IdleSeconds, int RequiredIdleSeconds)
{
    public int RetryAfterSeconds => Active ? Math.Max(1, (int)Math.Ceiling(RequiredIdleSeconds - IdleSeconds)) : 0;
}

/// <summary>用户正在使用电脑时抛出，HTTP 层转成 423 和 Retry-After。</summary>
public sealed class UserActiveException(UserActivityState state, object? batchProgress = null)
    : Exception("The user is currently using this PC.")
{
    public UserActivityState State { get; } = state;
    public object? BatchProgress { get; } = batchProgress;
}

/// <summary>
/// 判断用户本人是否正在使用这台电脑：最近 N 秒内有不是本程序注入的键盘或鼠标输入就算在用（N 默认 10，0 表示关闭）。
/// 用户在用时，读写界面的远程操作一律不执行，避免和用户抢鼠标键盘，也避免截到用户正在看的窗口误导 agent。
/// GetLastInputInfo 只给出最后一次输入的时间，分不清是谁的：本程序的合成输入都经过 InputSimulator.OnInputDesktop，
/// 在那里登记注入的时间段，落在这个时间段里（含之后 0.25 秒）的输入算本程序的，其余算用户的；
/// 再用后台采样记住用户最后一次输入，免得被随后的合成输入覆盖。
/// </summary>
public static class UserActivityGuard
{
    private const int DefaultRequiredIdleSeconds = 10;
    // 合成输入在系统里登记的时间会略晚于注入调用返回
    private const int SyntheticGraceMs = 250;
    private const int SampleIntervalMs = 500;

    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DesktopAssistant", "settings.json");

    private static readonly object stateLock = new();
    private static readonly object settingsLock = new();
    private static bool syntheticActive;
    private static bool hasSyntheticWindow;
    private static uint syntheticStartTick;
    private static uint syntheticEndTick;
    private static bool hasUserInput;
    private static uint lastUserInputTick;
    private static System.Threading.Timer? sampler;
    private static DateTime settingsWriteTime = DateTime.MinValue;
    private static int requiredIdleSeconds = DefaultRequiredIdleSeconds;

    [DllImport("user32.dll")]
    private static extern bool GetLastInputInfo(ref LASTINPUTINFO info);

    [StructLayout(LayoutKind.Sequential)]
    private struct LASTINPUTINFO
    {
        public uint cbSize;
        public uint dwTime;
    }

    public static void Start()
    {
        Sample();
        sampler ??= new System.Threading.Timer(_ => Sample(), null, SampleIntervalMs, SampleIntervalMs);
    }

    public static void Stop()
    {
        sampler?.Dispose();
        sampler = null;
    }

    /// <summary>注入前调用：先记下此前用户的输入，再标记注入开始。</summary>
    public static void BeginSyntheticInput()
    {
        Sample();
        lock (stateLock)
        {
            syntheticActive = true;
            syntheticStartTick = NowTick();
        }
    }

    public static void EndSyntheticInput()
    {
        lock (stateLock)
        {
            syntheticActive = false;
            hasSyntheticWindow = true;
            syntheticEndTick = NowTick();
        }
    }

    public static UserActivityState Current()
    {
        Sample();
        int required = RequiredIdleSeconds();
        double idleSeconds;
        lock (stateLock)
        {
            idleSeconds = hasUserInput ? unchecked(NowTick() - lastUserInputTick) / 1000.0 : double.MaxValue;
        }
        double reportedIdle = Math.Round(Math.Min(idleSeconds, 86400 * 49), 1);
        return new UserActivityState(required > 0 && idleSeconds < required, reportedIdle, required);
    }

    /// <summary>用户在用时抛出 <see cref="UserActiveException"/>；force 为 true（用户授权）时跳过检查。</summary>
    public static void EnsureIdle(bool force)
    {
        if (force) return;
        var state = Current();
        if (state.Active) throw new UserActiveException(state);
    }

    /// <summary>最后一次合成输入距今的秒数，没有过合成输入时为 null。</summary>
    public static double? SecondsSinceSyntheticInput()
    {
        lock (stateLock)
        {
            if (syntheticActive) return 0;
            return hasSyntheticWindow ? Math.Round(unchecked(NowTick() - syntheticEndTick) / 1000.0, 1) : null;
        }
    }

    private static void Sample()
    {
        var info = new LASTINPUTINFO { cbSize = (uint)Marshal.SizeOf<LASTINPUTINFO>() };
        if (!GetLastInputInfo(ref info)) return;
        lock (stateLock)
        {
            if (IsSynthetic(info.dwTime)) return;
            if (!hasUserInput || TickDiff(info.dwTime, lastUserInputTick) > 0)
            {
                lastUserInputTick = info.dwTime;
                hasUserInput = true;
            }
        }
    }

    // 调用方持有 stateLock
    private static bool IsSynthetic(uint inputTick)
    {
        if (syntheticActive) return TickDiff(inputTick, syntheticStartTick) >= 0;
        return hasSyntheticWindow
            && TickDiff(inputTick, syntheticStartTick) >= 0
            && TickDiff(inputTick, syntheticEndTick) <= SyntheticGraceMs;
    }

    // GetLastInputInfo 和 Environment.TickCount 都是开机后的毫秒数，约 49.7 天回绕一次，用有符号差值比较
    private static int TickDiff(uint later, uint earlier) => unchecked((int)(later - earlier));

    private static uint NowTick() => unchecked((uint)Environment.TickCount);

    /// <summary>读取 settings.json 的 userIdleSeconds，文件改动后自动生效；缺失或无效时用默认值。</summary>
    private static int RequiredIdleSeconds()
    {
        lock (settingsLock) return ReadRequiredIdleSeconds();
    }

    private static int ReadRequiredIdleSeconds()
    {
        try
        {
            if (!File.Exists(SettingsPath)) return requiredIdleSeconds = DefaultRequiredIdleSeconds;
            DateTime writeTime = File.GetLastWriteTimeUtc(SettingsPath);
            if (writeTime == settingsWriteTime) return requiredIdleSeconds;
            settingsWriteTime = writeTime;
            using var document = JsonDocument.Parse(File.ReadAllText(SettingsPath));
            requiredIdleSeconds = document.RootElement.TryGetProperty("userIdleSeconds", out var value) && value.TryGetInt32(out int seconds)
                ? Math.Clamp(seconds, 0, 3600)
                : DefaultRequiredIdleSeconds;
        }
        catch (Exception ex)
        {
            Logger.Error($"读取 {SettingsPath} 失败，使用默认空闲秒数", ex);
            requiredIdleSeconds = DefaultRequiredIdleSeconds;
        }
        return requiredIdleSeconds;
    }
}
