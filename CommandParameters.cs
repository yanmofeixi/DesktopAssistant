using System.Collections.Specialized;
using System.Globalization;
using System.Text.Json;

namespace DesktopAssistant;

public sealed class CommandException(string message, int statusCode = 400, object? details = null) : Exception(message)
{
    public int StatusCode { get; } = statusCode;
    public object? Details { get; } = details;
}

public abstract record CommandResult;
public sealed record JsonResult(object Data) : CommandResult;
public sealed record TextResult(string Text) : CommandResult;
public sealed record ImageResult(ScreenshotCapture Capture, bool Jpeg) : CommandResult;

/// <summary>Case-insensitive command arguments from a query string, JSON object or batch step.</summary>
public sealed class CommandParameters
{
    private readonly Dictionary<string, string> values = new(StringComparer.OrdinalIgnoreCase);

    public static CommandParameters FromQuery(NameValueCollection query)
    {
        var parameters = new CommandParameters();
        foreach (string? key in query.AllKeys)
            if (key != null && query[key] is { } value) parameters.values[key] = value;
        return parameters;
    }

    public void MergeJson(JsonElement json)
    {
        if (json.ValueKind != JsonValueKind.Object)
            throw new CommandException("JSON 请求体必须是对象；批量操作请使用 /batch 并发送数组。");
        foreach (var property in json.EnumerateObject())
        {
            if (property.Value.ValueKind == JsonValueKind.Null) continue;
            values[property.Name] = property.Value.ValueKind == JsonValueKind.String
                ? property.Value.GetString()!
                : property.Value.GetRawText();
        }
    }

    public void SetIfMissing(string key, string value) => values.TryAdd(key, value);

    public string? Get(params string[] keys)
    {
        foreach (string key in keys)
            if (values.TryGetValue(key, out string? value) && value.Length > 0) return value;
        return null;
    }

    public string Require(string key, string hint) =>
        Get(key) ?? throw new CommandException($"缺少参数 {key}：{hint}");

    public int? Int(string key)
    {
        string? raw = Get(key);
        if (raw == null) return null;
        return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value)
            ? value
            : throw new CommandException($"参数 {key} 必须是整数，收到 '{raw}'。");
    }

    public int Int(string key, int fallback, int min, int max) => Math.Clamp(Int(key) ?? fallback, min, max);

    public double Double(string key, double fallback)
    {
        string? raw = Get(key);
        if (raw == null) return fallback;
        return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
            ? value
            : throw new CommandException($"参数 {key} 必须是数字，收到 '{raw}'。");
    }

    public bool Flag(string key) => Get(key)?.ToLowerInvariant() is "true" or "1" or "yes";

    /// <summary>Parses "x,y,width,height".</summary>
    public Rectangle? Rect(string key)
    {
        string? raw = Get(key);
        if (raw == null) return null;
        var parts = raw.Split(',', StringSplitOptions.TrimEntries);
        if (parts.Length == 4 && parts.All(p => int.TryParse(p, NumberStyles.Integer, CultureInfo.InvariantCulture, out _)))
        {
            int[] n = parts.Select(p => int.Parse(p, CultureInfo.InvariantCulture)).ToArray();
            if (n[2] > 0 && n[3] > 0) return new Rectangle(n[0], n[1], n[2], n[3]);
        }
        throw new CommandException($"参数 {key} 格式为 x,y,宽,高（桌面坐标，宽高为正），收到 '{raw}'。");
    }
}
