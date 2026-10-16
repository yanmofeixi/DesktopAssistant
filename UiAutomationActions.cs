using System.Windows.Automation;

namespace DesktopAssistant;

/// <summary>A live UI Automation element plus what listings show about it.</summary>
public sealed record UiTarget(AutomationElement Element, string Type, string Name, Point Center)
{
    public string Describe() => $"{Type} \"{Name}\" @{Center.X},{Center.Y}";
}

/// <summary>
/// Reads and drives single UI Automation elements through their patterns, so a button can be
/// pressed or an edit box filled without moving the mouse, using the clipboard or needing focus.
/// </summary>
public static class UiAutomationActions
{
    // A Win32 button that opens a modal dialog does not return from Invoke until the dialog closes.
    private static readonly TimeSpan InvokeWait = TimeSpan.FromSeconds(5);

    public static string Describe(ToggleState state) => state switch
    {
        ToggleState.On => "on",
        ToggleState.Off => "off",
        _ => "indeterminate"
    };

    public static string Describe(ExpandCollapseState state) => state switch
    {
        ExpandCollapseState.Expanded => "expanded",
        ExpandCollapseState.Collapsed => "collapsed",
        ExpandCollapseState.PartiallyExpanded => "partially_expanded",
        _ => "leaf"
    };

    /// <summary>Same matching rules as click by name: exact names first, then names or values containing it.</summary>
    public static UiTarget Find(CaptureWindow window, string name, string? types, int? index, bool exact) =>
        UiAutomationReader.WithPhysicalPixels(() =>
        {
            var root = AutomationElement.FromHandle(window.Hwnd);
            var matches = new List<UiTarget>();
            foreach (AutomationElement element in root.FindAll(TreeScope.Descendants, UiAutomationReader.ElementCondition(types)))
            {
                var rect = element.Current.BoundingRectangle;
                if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0) continue;
                string elementName = element.Current.Name?.Trim() ?? "";
                bool nameMatches = UiAutomationReader.IsExactName(elementName, name);
                if (!nameMatches && (exact || !(elementName.Contains(name, StringComparison.OrdinalIgnoreCase) ||
                                                ListedValue(element).Contains(name, StringComparison.OrdinalIgnoreCase))))
                    continue;
                string type = element.Current.ControlType.ProgrammaticName.Replace("ControlType.", "").ToLowerInvariant();
                matches.Add(new UiTarget(element, type, elementName,
                    new Point((int)(rect.Left + rect.Width / 2), (int)(rect.Top + rect.Height / 2))));
            }
            var exactMatches = matches.Where(m => UiAutomationReader.IsExactName(m.Name, name)).ToList();
            var pool = exactMatches.Count > 0 ? exactMatches : matches;
            if (pool.Count == 0)
                throw new CommandException($"窗口中没有{(exact ? "名称恰为" : "名称或值包含")} '{name}' 的可见元素；可用 /elements 查看。", 404);
            if (index.HasValue)
                return index.Value >= 0 && index.Value < pool.Count
                    ? pool[index.Value]
                    : throw new CommandException($"index 超出范围：'{name}' 共匹配 {pool.Count} 个元素。");
            if (pool.Count > 1)
                throw new CommandException($"'{name}' 匹配 {pool.Count} 个元素，请加 type 或 index（从 0 开始）。", 409,
                    pool.Take(20).Select(m => m.Describe()));
            return pool[0];
        });

    // Matching looks only at ValuePattern, like /elements; documents' full text would be slow.
    private static string ListedValue(AutomationElement element) =>
        element.TryGetCurrentPattern(ValuePattern.Pattern, out object value) ? ((ValuePattern)value).Current.Value ?? "" : "";

    /// <summary>The full value: ValuePattern first, then the text of a TextPattern document.</summary>
    public static string ReadValue(AutomationElement element, int maxChars = 1_000_000)
    {
        if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object value))
            return ((ValuePattern)value).Current.Value ?? "";
        if (element.TryGetCurrentPattern(TextPattern.Pattern, out object text))
            return ((TextPattern)text).DocumentRange.GetText(maxChars);
        return "";
    }

    public static string? ReadState(AutomationElement element)
    {
        if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object selection))
            return ((SelectionItemPattern)selection).Current.IsSelected ? "selected" : "unselected";
        if (element.TryGetCurrentPattern(TogglePattern.Pattern, out object toggle))
            return Describe(((TogglePattern)toggle).Current.ToggleState);
        if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out object expand))
            return Describe(((ExpandCollapsePattern)expand).Current.ExpandCollapseState);
        return null;
    }

    public static object Snapshot(AutomationElement element)
    {
        var rect = element.Current.BoundingRectangle;
        return new
        {
            type = element.Current.ControlType.ProgrammaticName.Replace("ControlType.", "").ToLowerInvariant(),
            name = element.Current.Name ?? "",
            value = ReadValue(element),
            state = ReadState(element),
            enabled = element.Current.IsEnabled,
            hasFocus = element.Current.HasKeyboardFocus,
            x = rect.IsEmpty ? (int?)null : (int)(rect.Left + rect.Width / 2),
            y = rect.IsEmpty ? (int?)null : (int)(rect.Top + rect.Height / 2)
        };
    }

    /// <summary>Presses the element with the first pattern it supports; returns the pattern used.</summary>
    public static (string Pattern, bool Completed) Invoke(AutomationElement element)
    {
        Action action;
        string pattern;
        if (element.TryGetCurrentPattern(InvokePattern.Pattern, out object invoke))
            (pattern, action) = ("invoke", ((InvokePattern)invoke).Invoke);
        else if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object selection))
            (pattern, action) = ("select", ((SelectionItemPattern)selection).Select);
        else if (element.TryGetCurrentPattern(TogglePattern.Pattern, out object toggle))
            (pattern, action) = ("toggle", ((TogglePattern)toggle).Toggle);
        else if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out object expand))
        {
            var expandCollapse = (ExpandCollapsePattern)expand;
            bool expanded = expandCollapse.Current.ExpandCollapseState == ExpandCollapseState.Expanded;
            pattern = expanded ? "collapse" : "expand";
            action = expanded ? expandCollapse.Collapse : expandCollapse.Expand;
        }
        else
            throw new CommandException("该元素不支持 Invoke/SelectionItem/Toggle/ExpandCollapse，只能用 click 按坐标点。", 409);

        if (!element.Current.IsEnabled) throw new CommandException("该元素当前不可用（disabled）。", 409);
        var task = Task.Run(action);
        bool completed = task.Wait(InvokeWait);
        if (task.IsFaulted) throw task.Exception!.InnerException ?? task.Exception;
        return (pattern, completed);
    }

    /// <summary>Replaces the element's value through ValuePattern and returns the value read back.</summary>
    public static string SetValue(AutomationElement element, string text)
    {
        if (!element.TryGetCurrentPattern(ValuePattern.Pattern, out object pattern))
            throw new CommandException("该元素不支持 ValuePattern，改用 click 后 paste。", 409);
        var valuePattern = (ValuePattern)pattern;
        if (valuePattern.Current.IsReadOnly) throw new CommandException("该元素是只读的。", 409);
        valuePattern.SetValue(text);
        return valuePattern.Current.Value ?? "";
    }

    public static AutomationElement Focused() =>
        AutomationElement.FocusedElement ?? throw new CommandException("当前没有焦点元素。", 404);
}
