using System.ComponentModel;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace DesktopAssistant;

public static class InputSimulator
{
    private static readonly object inputLock = new();

    #region Win32 API Definitions

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr OpenInputDesktop(uint dwFlags, bool fInherit, uint dwDesiredAccess);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool CloseDesktop(IntPtr hDesktop);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool SetThreadDesktop(IntPtr hDesktop);
    [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr GetThreadDesktop(int dwThreadId);
    [DllImport("kernel32.dll")] private static extern int GetCurrentThreadId();
    [DllImport("user32.dll")] private static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] private static extern bool GetCursorPos(out Point point);
    [DllImport("user32.dll")] private static extern void mouse_event(uint flags, uint dx, uint dy, int data, UIntPtr extraInfo);
    [DllImport("user32.dll", SetLastError = true)] private static extern uint SendInput(uint count, INPUT[] inputs, int size);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hwnd, int command);
    [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hwnd);

    private const uint DESKTOP_MAXIMUM_ALLOWED = 0x01FF;
    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;
    private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    private const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
    private const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
    private const uint MOUSEEVENTF_WHEEL = 0x0800;
    private const uint MOUSEEVENTF_HWHEEL = 0x1000;
    private const uint INPUT_KEYBOARD = 1;
    private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const uint KEYEVENTF_UNICODE = 0x0004;
    private const int SW_RESTORE = 9;
    private const ushort VK_MENU = 0x12;

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint type;
        public InputUnion u;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)] public MOUSEINPUT mi;
        [FieldOffset(0)] public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }

    #endregion

    public static Point CursorPosition()
    {
        GetCursorPos(out Point point);
        return point;
    }

    public static void Click(Point? target, string button, bool isDouble)
    {
        (uint down, uint up) = button switch
        {
            "left" => (MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP),
            "right" => (MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP),
            "middle" => (MOUSEEVENTF_MIDDLEDOWN, MOUSEEVENTF_MIDDLEUP),
            _ => throw new CommandException($"button 必须为 left、right 或 middle，收到 '{button}'。")
        };
        OnInputDesktop(() =>
        {
            if (target.HasValue)
            {
                SetCursorPos(target.Value.X, target.Value.Y);
                Thread.Sleep(25);
            }
            for (int i = 0; i < (isDouble ? 2 : 1); i++)
            {
                if (i > 0) Thread.Sleep(80);
                mouse_event(down, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(30);
                mouse_event(up, 0, 0, 0, UIntPtr.Zero);
            }
        });
    }

    public static void Move(Point target) => OnInputDesktop(() => SetCursorPos(target.X, target.Y));

    public static void Drag(Point from, Point to, int durationMs, string button)
    {
        uint down = button == "right" ? MOUSEEVENTF_RIGHTDOWN : MOUSEEVENTF_LEFTDOWN;
        uint up = button == "right" ? MOUSEEVENTF_RIGHTUP : MOUSEEVENTF_LEFTUP;
        OnInputDesktop(() =>
        {
            SetCursorPos(from.X, from.Y);
            Thread.Sleep(30);
            mouse_event(down, 0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(30);
            int steps = Math.Max(10, durationMs / 15);
            for (int i = 1; i <= steps; i++)
            {
                double t = (double)i / steps;
                SetCursorPos((int)(from.X + (to.X - from.X) * t), (int)(from.Y + (to.Y - from.Y) * t));
                Thread.Sleep(durationMs / steps);
            }
            SetCursorPos(to.X, to.Y);
            Thread.Sleep(30);
            mouse_event(up, 0, 0, 0, UIntPtr.Zero);
        });
    }

    public static void Scroll(int delta, bool horizontal, Point? target)
    {
        OnInputDesktop(() =>
        {
            if (target.HasValue)
            {
                SetCursorPos(target.Value.X, target.Value.Y);
                Thread.Sleep(20);
            }
            mouse_event(horizontal ? MOUSEEVENTF_HWHEEL : MOUSEEVENTF_WHEEL, 0, 0, delta, UIntPtr.Zero);
        });
    }

    public static void TypeText(string text) => OnInputDesktop(() => SendUnicodeText(text));

    public static void PasteText(string text, string pasteKeys)
    {
        // Clipboard requires STA and must run in the logged-in desktop process,
        // not the SSH session's separate clipboard. Serialize with other input.
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                OnInputDesktop(() =>
                {
                    var data = new DataObject();
                    data.SetText(text, TextDataFormat.UnicodeText);
                    Clipboard.SetDataObject(data, true, 10, 50);
                    PressKeySequence(pasteKeys);
                    // Give the target time to consume the paste before the next input.
                    // Keep the clipboard intact for applications that read it later.
                    Thread.Sleep(200);
                });
            }
            catch (Exception ex) { failure = ex; }
        }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
    }

    /// <summary>Comma-separated combos run in order, e.g. "ctrl+m,m" for chord shortcuts.</summary>
    public static void PressKeys(string sequence) => OnInputDesktop(() => PressKeySequence(sequence));

    public static void FocusWindow(IntPtr hwnd)
    {
        OnInputDesktop(() =>
        {
            if (IsIconic(hwnd)) ShowWindow(hwnd, SW_RESTORE);
            // Windows only lets the process that received the last input change the
            // foreground window; a synthetic Alt tap satisfies that rule.
            SendChecked([MakeKeyInput(VK_MENU, false), MakeKeyInput(VK_MENU, true)]);
            SetForegroundWindow(hwnd);
        });
    }

    private static void OnInputDesktop(Action action)
    {
        lock (inputLock)
        {
            IntPtr oldDesktop = IntPtr.Zero;
            IntPtr inputDesktop = OpenInputDesktop(0, false, DESKTOP_MAXIMUM_ALLOWED);
            bool switched = false;
            try
            {
                if (inputDesktop != IntPtr.Zero)
                {
                    oldDesktop = GetThreadDesktop(GetCurrentThreadId());
                    switched = SetThreadDesktop(inputDesktop);
                }
                action();
            }
            finally
            {
                if (switched && oldDesktop != IntPtr.Zero) SetThreadDesktop(oldDesktop);
                if (inputDesktop != IntPtr.Zero) CloseDesktop(inputDesktop);
            }
        }
    }

    private static void PressKeySequence(string sequence)
    {
        var combos = sequence.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (combos.Length == 0) throw new CommandException("缺少按键，例如 enter、ctrl+c、ctrl+m,m。");
        for (int i = 0; i < combos.Length; i++)
        {
            if (i > 0) Thread.Sleep(60);
            PressCombo(combos[i]);
        }
    }

    private static void PressCombo(string combo)
    {
        var modifiers = new List<ushort>();
        var keys = new List<ushort>();
        // A '+' in a URL query decodes to a space, so spaces separate keys as well.
        foreach (string part in combo.Split(['+', ' '], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            string name = part.ToLowerInvariant();
            if (TryGetModifier(name, out ushort modifier)) modifiers.Add(modifier);
            else if (TryGetKey(name, out ushort key)) keys.Add(key);
            else throw new CommandException($"无法识别的按键名称: '{part}'");
        }

        var inputs = new List<INPUT>();
        inputs.AddRange(modifiers.Select(vk => MakeKeyInput(vk, false)));
        inputs.AddRange(keys.Select(vk => MakeKeyInput(vk, false)));
        inputs.AddRange(keys.Select(vk => MakeKeyInput(vk, true)));
        inputs.AddRange(Enumerable.Reverse(modifiers).Select(vk => MakeKeyInput(vk, true)));
        if (inputs.Count > 0) SendChecked(inputs.ToArray());
    }

    private static void SendUnicodeText(string text)
    {
        // Some modern editors coalesce a large burst of VK_PACKET events incorrectly.
        // Send one Unicode scalar at a time, keeping surrogate pairs together.
        for (int i = 0; i < text.Length; i++)
        {
            var inputs = new List<INPUT>(4);
            char c = text[i];
            if (c is '\r' or '\n' or '\t')
            {
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                ushort vk = c == '\t' ? (ushort)0x09 : (ushort)0x0D;
                inputs.Add(MakeKeyInput(vk, false));
                inputs.Add(MakeKeyInput(vk, true));
            }
            else
            {
                AddUnicodePair(inputs, c);
                if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                    AddUnicodePair(inputs, text[++i]);
            }
            SendChecked(inputs.ToArray());
            Thread.Sleep(50);
        }
    }

    private static void AddUnicodePair(List<INPUT> inputs, char character)
    {
        foreach (uint flags in new[] { KEYEVENTF_UNICODE, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP })
            inputs.Add(new INPUT
            {
                type = INPUT_KEYBOARD,
                u = new InputUnion { ki = new KEYBDINPUT { wScan = character, dwFlags = flags } }
            });
    }

    private static void SendChecked(INPUT[] inputs)
    {
        uint sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        if (sent != inputs.Length)
            throw new Win32Exception(Marshal.GetLastWin32Error(),
                $"SendInput 只发送了 {sent}/{inputs.Length} 个输入事件，请确认桌面已解锁且目标窗口可接收输入。");
    }

    private static INPUT MakeKeyInput(ushort vk, bool keyUp)
    {
        uint flags = keyUp ? KEYEVENTF_KEYUP : 0;
        if (IsExtendedKey(vk)) flags |= KEYEVENTF_EXTENDEDKEY;
        return new INPUT
        {
            type = INPUT_KEYBOARD,
            u = new InputUnion { ki = new KEYBDINPUT { wVk = vk, dwFlags = flags } }
        };
    }

    private static bool TryGetModifier(string name, out ushort vk)
    {
        vk = name switch
        {
            "ctrl" or "control" => 0x11,
            "alt" or "menu" => 0x12,
            "shift" => 0x10,
            "win" or "windows" or "meta" or "cmd" or "command" => 0x5B,
            _ => 0
        };
        return vk != 0;
    }

    private static bool TryGetKey(string name, out ushort vk)
    {
        vk = name switch
        {
            "enter" or "return" => 0x0D,
            "esc" or "escape" => 0x1B,
            "tab" => 0x09,
            "backspace" or "back" => 0x08,
            "delete" or "del" => 0x2E,
            "space" or "spacebar" => 0x20,
            "insert" or "ins" => 0x2D,
            "up" => 0x26,
            "down" => 0x28,
            "left" => 0x25,
            "right" => 0x27,
            "home" => 0x24,
            "end" => 0x23,
            "pageup" or "pgup" => 0x21,
            "pagedown" or "pgdn" => 0x22,
            "capslock" or "caps" => 0x14,
            "printscreen" or "prtsc" => 0x2C,
            // Punctuation keys have names because '+' and ',' are syntax in combos.
            "plus" or "equal" => 0xBB,
            "minus" => 0xBD,
            "comma" => 0xBC,
            "period" => 0xBE,
            "slash" => 0xBF,
            "semicolon" => 0xBA,
            "quote" => 0xDE,
            "backquote" => 0xC0,
            "bracketleft" => 0xDB,
            "bracketright" => 0xDD,
            "backslash" => 0xDC,
            _ => 0
        };
        if (vk != 0) return true;
        if (name.Length > 1 && name[0] == 'f' && int.TryParse(name[1..], out int f) && f is >= 1 and <= 24)
        {
            vk = (ushort)(0x70 + f - 1);
            return true;
        }
        if (name.Length == 1 && char.ToUpperInvariant(name[0]) is var c && (c is >= 'A' and <= 'Z' || c is >= '0' and <= '9'))
        {
            vk = c;
            return true;
        }
        return false;
    }

    private static bool IsExtendedKey(ushort vk) => vk is
        0x21 or 0x22 or 0x23 or 0x24 or 0x25 or 0x26 or 0x27 or 0x28 or // PgUp, PgDn, End, Home, arrows
        0x2D or 0x2E or // Insert, Delete
        0x5B or 0x5C; // LWin, RWin
}
