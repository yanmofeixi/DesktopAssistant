using System.Windows.Automation;

namespace DesktopAssistant;

/// <summary>
/// The assert operation: every given check must hold, otherwise 412 with what was found.
/// Placed inside a batch it stops the following steps, so nothing is typed or submitted
/// into the wrong window or control.
/// </summary>
public static class UiAssertions
{
    private sealed record Failure(string check, string expected, string? actual);

    public static CommandResult Assert(CommandParameters p)
    {
        var failures = new List<Failure>();
        var passed = new List<string>();
        int checks = 0;

        if (p.Get("foreground") is { } foregroundSelector)
        {
            checks++;
            var expected = UiAutomationReader.ResolveWindow(foregroundSelector, allowMinimized: true);
            var actual = UiAutomationReader.ResolveWindow("active", allowMinimized: true);
            Record(failures, passed, "foreground", expected.Hwnd == actual.Hwnd, expected.Describe(), actual.Describe());
        }

        if (p.Get("focusedName", "focusedType", "focusedValue", "focusedValueContains") != null)
            checks += CheckFocused(p, failures, passed);

        if (p.Get("name") is { } name)
        {
            checks++;
            CheckElement(p, name, failures, passed);
        }

        if (p.Get("text") is { } text)
        {
            checks++;
            var window = UiAutomationReader.ResolveWindow(p.Get("window"));
            bool present = UiAutomationReader.ReadText(window, p.Flag("all"), 2_000_000)
                .Contains(text, StringComparison.OrdinalIgnoreCase);
            bool wantPresent = !p.Flag("textGone");
            Record(failures, passed, wantPresent ? "text" : "textGone", present == wantPresent,
                (wantPresent ? "contains " : "does not contain ") + text, present ? "present" : "absent");
        }

        if (checks == 0)
            throw new CommandException("assert 需要至少一项：foreground、focusedName/focusedType/focusedValue/focusedValueContains、name（可加 value/valueContains/state/gone）、text（可加 textGone=1）。");
        if (failures.Count > 0)
            throw new CommandException($"断言不成立：{string.Join("；", failures.Select(f => f.check))}", 412,
                new { failed = failures, passed });
        return new JsonResult(new { status = "ok", action = "assert", passed });
    }

    private static int CheckFocused(CommandParameters p, List<Failure> failures, List<string> passed)
    {
        var focused = UiAutomationActions.Focused();
        string type = focused.Current.ControlType.ProgrammaticName.Replace("ControlType.", "").ToLowerInvariant();
        string name = focused.Current.Name ?? "";
        string value = UiAutomationActions.ReadValue(focused);
        int checks = 0;
        if (p.Get("focusedType") is { } expectedType)
        {
            checks++;
            Record(failures, passed, "focusedType", type.Equals(expectedType, StringComparison.OrdinalIgnoreCase), expectedType, type);
        }
        if (p.Get("focusedName") is { } expectedName)
        {
            checks++;
            Record(failures, passed, "focusedName", UiAutomationReader.IsExactName(name, expectedName), expectedName, name);
        }
        if (p.Get("focusedValue") is { } expectedValue)
        {
            checks++;
            Record(failures, passed, "focusedValue", value == expectedValue, expectedValue, value);
        }
        if (p.Get("focusedValueContains") is { } fragment)
        {
            checks++;
            Record(failures, passed, "focusedValueContains", value.Contains(fragment, StringComparison.Ordinal), fragment, value);
        }
        return checks;
    }

    private static void CheckElement(CommandParameters p, string name, List<Failure> failures, List<string> passed)
    {
        var window = UiAutomationReader.ResolveWindow(p.Get("window"));
        bool exact = p.Flag("exact");
        if (p.Flag("gone"))
        {
            var found = UiAutomationReader.FindElements(window, name, p.Get("type"), 50, exact);
            var relevant = exact ? found : found.Where(e => UiAutomationReader.IsExactName(e.Name, name)).ToList();
            Record(failures, passed, "gone", relevant.Count == 0, $"no element named '{name}'",
                relevant.Count == 0 ? null : string.Join(" | ", relevant.Select(e => e.Describe())));
            return;
        }

        UiTarget target;
        try
        {
            target = UiAutomationActions.Find(window, name, p.Get("type"), p.Int("index"), exact);
        }
        catch (CommandException ex) when (ex.StatusCode is 404 or 409)
        {
            Record(failures, passed, "element", false, $"exactly one element '{name}'", ex.Message);
            return;
        }

        string value = UiAutomationActions.ReadValue(target.Element);
        string? state = UiAutomationActions.ReadState(target.Element);
        bool any = false;
        if (p.Get("value") is { } expectedValue)
        {
            any = true;
            Record(failures, passed, "value", value == expectedValue, expectedValue, value);
        }
        if (p.Get("valueContains") is { } fragment)
        {
            any = true;
            Record(failures, passed, "valueContains", value.Contains(fragment, StringComparison.Ordinal), fragment, value);
        }
        if (p.Get("state") is { } expectedState)
        {
            any = true;
            Record(failures, passed, "state", string.Equals(state, expectedState, StringComparison.OrdinalIgnoreCase), expectedState, state);
        }
        if (!any) passed.Add($"element {target.Describe()}");
    }

    private static void Record(List<Failure> failures, List<string> passed, string check, bool ok, string expected, string? actual)
    {
        if (ok) passed.Add(check);
        else failures.Add(new Failure(check, expected, actual is { Length: > 300 } ? actual[..300] + "…" : actual));
    }
}
