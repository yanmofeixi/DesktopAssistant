using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Automation;

namespace DesktopAssistant;

public sealed record UiElement(string Type, string Name, string Value, Point Center, bool Enabled, string? State = null)
{
    public string Describe()
    {
        string line = $"{Type} \"{Shorten(Name)}\"";
        if (Value.Length > 0 && Value != Name) line += $" value=\"{Shorten(Value)}\"";
        if (State != null) line += $" [{State}]";
        if (!Enabled) line += " disabled";
        return line + $" @{Center.X},{Center.Y}";
    }

    private static string Shorten(string text)
    {
        text = text.ReplaceLineEndings(" ");
        return text.Length <= 100 ? text : text[..100] + "…";
    }
}

/// <summary>Reads windows through UI Automation so callers get text instead of screenshots.</summary>
public static class UiAutomationReader
{
    private static readonly ControlType[] InteractiveTypes =
    [
        ControlType.Button, ControlType.Hyperlink, ControlType.Edit, ControlType.ComboBox,
        ControlType.CheckBox, ControlType.RadioButton, ControlType.TabItem, ControlType.MenuItem,
        ControlType.ListItem, ControlType.TreeItem, ControlType.SplitButton
    ];

    private static readonly Dictionary<string, ControlType> TypesByName = typeof(ControlType)
        .GetFields(BindingFlags.Public | BindingFlags.Static)
        .Where(f => f.FieldType == typeof(ControlType))
        .ToDictionary(f => f.Name.ToLowerInvariant(), f => (ControlType)f.GetValue(null)!);

    private static readonly Condition DocumentCondition =
        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document);

    public static CaptureWindow ResolveWindow(string? selector, bool allowMinimized = false) =>
        WithPhysicalPixels(() => WindowCapture.Select(string.IsNullOrWhiteSpace(selector) ? "active" : selector, allowMinimized));

    public static string ReadText(CaptureWindow window, bool includeOffscreen, int maxChars)
    {
        return WithPhysicalPixels(() =>
        {
            var root = AutomationElement.FromHandle(window.Hwnd);
            var parts = new List<string>();
            foreach (AutomationElement document in FindDocuments(root))
            {
                if (!document.TryGetCurrentPattern(TextPattern.Pattern, out object pattern)) continue;
                string text = Tidy(ReadDocument((TextPattern)pattern, includeOffscreen, maxChars));
                // Chrome exposes iframes as nested documents whose text may already be in the parent.
                if (text.Length == 0 || parts.Any(p => p.Contains(text, StringComparison.Ordinal))) continue;
                parts.Add(text);
            }
            string result = parts.Count > 0
                ? string.Join("\n\n", parts)
                : string.Join("\n", ListElements(root, null, "all", 500).Select(e => e.Describe()));
            return result.Length <= maxChars ? result : result[..maxChars] + "\n…（已截断）";
        });
    }

    public static IReadOnlyList<UiElement> FindElements(CaptureWindow window, string? name, string? types, int limit,
        bool exact = false, bool withState = false) =>
        WithPhysicalPixels(() => ListElements(AutomationElement.FromHandle(window.Hwnd), name, types, limit, exact, withState));

    /// <summary>
    /// Exact name matches win over substring matches (exact=true accepts only names equal
    /// to <paramref name="name"/>); ambiguity is an error unless index is given.
    /// </summary>
    public static UiElement FindOne(CaptureWindow window, string name, string? types, int? index, bool exact = false)
    {
        var matches = FindElements(window, name, types, 500);
        var exactMatches = matches.Where(e => IsExactName(e.Name, name)).ToList();
        var pool = exact || exactMatches.Count > 0 ? exactMatches : matches.ToList();
        if (pool.Count == 0)
            throw new CommandException(exact
                ? $"窗口中没有名称恰为 '{name}' 的可见元素；可用 /elements 查看。"
                : $"窗口中没有名称或值包含 '{name}' 的可见元素；可用 /elements 查看。", 404);
        if (index.HasValue)
            return index.Value >= 0 && index.Value < pool.Count
                ? pool[index.Value]
                : throw new CommandException($"index 超出范围：'{name}' 共匹配 {pool.Count} 个元素。");
        if (pool.Count > 1)
            throw new CommandException($"'{name}' 匹配 {pool.Count} 个元素，请加 type 或 index（从 0 开始）。", 409,
                pool.Take(20).Select(e => e.Describe()));
        return pool[0];
    }

    public static string ReadUrl(CaptureWindow window)
    {
        return WithPhysicalPixels(() =>
        {
            var root = AutomationElement.FromHandle(window.Hwnd);
            // Chromium exposes the full page URL as the document's value; the address bar
            // (which hides the scheme) is the fallback.
            var candidates = FindDocuments(root).Cast<AutomationElement>().Append(
                root.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit)));
            foreach (var element in candidates)
            {
                if (element?.TryGetCurrentPattern(ValuePattern.Pattern, out object pattern) == true &&
                    ((ValuePattern)pattern).Current.Value is { Length: > 0 } value)
                    return value;
            }
            throw new CommandException("该窗口没有可读取的网址。", 404);
        });
    }

    // Chromium builds the web page accessibility tree only after a UI Automation client
    // asks for it, so the first query on a fresh browser may see no documents yet.
    private static AutomationElementCollection FindDocuments(AutomationElement root)
    {
        var documents = root.FindAll(TreeScope.Subtree, DocumentCondition);
        if (documents.Count > 0) return documents;
        Thread.Sleep(1000);
        return root.FindAll(TreeScope.Subtree, DocumentCondition);
    }

    /// <summary>Names compare case-insensitively and ignoring surrounding whitespace, as listings show them.</summary>
    internal static bool IsExactName(string elementName, string name) =>
        elementName.Trim().Equals(name.Trim(), StringComparison.OrdinalIgnoreCase);

    private static List<UiElement> ListElements(AutomationElement root, string? name, string? types, int limit,
        bool exact = false, bool withState = false)
    {
        var cache = new CacheRequest { AutomationElementMode = AutomationElementMode.None };
        cache.Add(AutomationElement.NameProperty);
        cache.Add(AutomationElement.ControlTypeProperty);
        cache.Add(AutomationElement.BoundingRectangleProperty);
        cache.Add(AutomationElement.IsEnabledProperty);
        cache.Add(ValuePattern.ValueProperty);
        if (withState)
        {
            cache.Add(SelectionItemPattern.IsSelectedProperty);
            cache.Add(TogglePattern.ToggleStateProperty);
            cache.Add(ExpandCollapsePattern.ExpandCollapseStateProperty);
        }

        AutomationElementCollection found;
        using (cache.Activate())
            found = root.FindAll(TreeScope.Descendants, ElementCondition(types));

        var elements = new List<UiElement>();
        foreach (AutomationElement element in found)
        {
            var rect = element.Cached.BoundingRectangle;
            if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0) continue;
            string elementName = element.Cached.Name?.Trim() ?? "";
            string value = (element.GetCachedPropertyValue(ValuePattern.ValueProperty) as string)?.Trim() ?? "";
            if (elementName.Length == 0 && value.Length == 0) continue;
            if (name != null && exact && !IsExactName(elementName, name)) continue;
            if (name != null && !exact && !elementName.Contains(name, StringComparison.OrdinalIgnoreCase) &&
                !value.Contains(name, StringComparison.OrdinalIgnoreCase)) continue;
            string type = element.Cached.ControlType.ProgrammaticName.Replace("ControlType.", "").ToLowerInvariant();
            var center = new Point((int)(rect.Left + rect.Width / 2), (int)(rect.Top + rect.Height / 2));
            string? state = withState ? CachedState(element) : null;
            elements.Add(new UiElement(type, elementName, value, center, element.Cached.IsEnabled, state));
            if (elements.Count >= limit) break;
        }
        return elements;
    }

    private static string? CachedState(AutomationElement element)
    {
        if (element.GetCachedPropertyValue(SelectionItemPattern.IsSelectedProperty, true) is bool selected)
            return selected ? "selected" : "unselected";
        if (element.GetCachedPropertyValue(TogglePattern.ToggleStateProperty, true) is ToggleState toggle)
            return UiAutomationActions.Describe(toggle);
        if (element.GetCachedPropertyValue(ExpandCollapsePattern.ExpandCollapseStateProperty, true) is ExpandCollapseState expand)
            return UiAutomationActions.Describe(expand);
        return null;
    }

    internal static Condition ElementCondition(string? types)
    {
        Condition visible = new PropertyCondition(AutomationElement.IsOffscreenProperty, false);
        if (types?.Trim().ToLowerInvariant() == "all") return visible;
        ControlType[] selected = types == null
            ? InteractiveTypes
            : types.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(t => TypesByName.TryGetValue(t.ToLowerInvariant(), out var type) ? type
                    : throw new CommandException($"未知元素类型 '{t}'；常用 button、hyperlink、edit、tabitem、menuitem、listitem，或 all。"))
                .ToArray();
        Condition[] byType = selected.Select(t => (Condition)new PropertyCondition(AutomationElement.ControlTypeProperty, t)).ToArray();
        return new AndCondition(visible, byType.Length == 1 ? byType[0] : new OrCondition(byType));
    }

    private static string ReadDocument(TextPattern pattern, bool includeOffscreen, int maxChars)
    {
        if (includeOffscreen) return pattern.DocumentRange.GetText(maxChars);
        try
        {
            return string.Join("\n", pattern.GetVisibleRanges().Select(r => r.GetText(maxChars)));
        }
        catch (InvalidOperationException)
        {
            return pattern.DocumentRange.GetText(maxChars);
        }
    }

    private static string Tidy(string text)
    {
        // U+FFFC marks embedded objects (images, iframes); it carries no readable text.
        text = text.Replace("\uFFFC", "").ReplaceLineEndings("\n");
        text = Regex.Replace(text, @"[ \t\u00A0]+\n", "\n");
        return Regex.Replace(text, @"\n{3,}", "\n\n").Trim();
    }

    // Match screenshot and mouse coordinates, which are physical pixels.
    internal static T WithPhysicalPixels<T>(Func<T> action)
    {
        IntPtr previous = WindowCapture.SetThreadDpiAwarenessContext(new IntPtr(-4));
        try { return action(); }
        finally { if (previous != IntPtr.Zero) WindowCapture.SetThreadDpiAwarenessContext(previous); }
    }
}
