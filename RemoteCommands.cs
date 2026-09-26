using System.Diagnostics;
using System.Drawing.Imaging;
using System.Text.Json;
using System.Windows.Automation;

namespace DesktopAssistant;

/// <summary>Remote control operations shared by single HTTP requests and /batch steps.</summary>
public sealed class RemoteCommands(ScreenshotManager screenshots, int port)
{
    private readonly DateTime startTime = DateTime.Now;

    private const string HelpText = """
        DesktopAssistant 远程控制：http://127.0.0.1:18888/<op>
        参数可放查询串、JSON 对象请求体或 /batch 步骤；非 ASCII 参数建议用 JSON 请求体。
        坐标为桌面物理像素；加 shot=1 时 x/y 取自上一张返回的截图（自动换算偏移和缩放）。
        window 可写 active、程序名、标题片段或 0x句柄；省略时截图为全屏，其余操作为前台窗口。

        windows                                   列出窗口：句柄 程序 "标题" 位置 尺寸
        focus      window                         恢复最小化并切到前台
        screenshot window region=x,y,w,h maxWidth format=png|jpg save=1
                                                  默认返回图片；save=1 返回保存路径与坐标 JSON
        text       window all=1 max=20000         读窗口文字，默认只读屏幕上可见的部分
        elements   window name type=button,edit|all limit=200
                                                  列出可见元素：类型 "名称" value="值" @中心坐标
        url        window                         浏览器当前网址
        click      x y | name [type index window]  button=left|right|middle double=1 shot=1
        move       x y       drag x1 y1 x2 y2 duration       scroll delta [x y] horizontal=1
        paste      text keys=ctrl+v               经剪贴板一次粘贴；网页终端用 keys=ctrl+shift+v
        type       text                           逐字模拟按键（每字约 50ms）
        key        combo                          enter、ctrl+w、ctrl+m,m（逗号分隔依次按下）
        wait       text [gone=1 all=1 window] | change=1 [window region threshold=3]
                   timeout=300 interval=1000      等文字出现/消失或画面变化，超时返回 408
        sleep      ms
        batch      POST JSON 数组，如 [{"op":"click","name":"确定"},{"op":"screenshot","maxWidth":1280}]
                   依次执行，任一步失败即停止；最后一步是未 save 的截图时响应就是图片
        status     屏幕与光标信息
        """;

    public CommandResult Execute(string op, CommandParameters p, JsonElement? batchSteps = null) => op switch
    {
        "" or "help" => new TextResult(HelpText),
        "status" or "info" => Status(),
        "windows" => Windows(),
        "focus" => Focus(p),
        "screenshot" => Screenshot(p),
        "text" => Text(p),
        "elements" => Elements(p),
        "url" => new TextResult(UiAutomationReader.ReadUrl(UiAutomationReader.ResolveWindow(p.Get("window")))),
        "click" => Click(p),
        "move" => Move(p),
        "drag" => Drag(p),
        "scroll" => Scroll(p),
        "paste" => Paste(p),
        "type" => TypeText(p),
        "key" => Key(p),
        "wait" => Wait(p),
        "sleep" => Sleep(p),
        "batch" => Batch(batchSteps),
        _ => throw new CommandException($"未知操作 '{op}'；GET /help 查看全部操作。", 404)
    };

    private static JsonResult Ok(object data) => new(data);

    private CommandResult Status()
    {
        Point cursor = InputSimulator.CursorPosition();
        return Ok(new
        {
            status = "ok",
            port,
            uptimeSeconds = (int)(DateTime.Now - startTime).TotalSeconds,
            cursor = new { x = cursor.X, y = cursor.Y },
            screens = Screen.AllScreens.Select(s => new
            {
                primary = s.Primary, x = s.Bounds.X, y = s.Bounds.Y, width = s.Bounds.Width, height = s.Bounds.Height
            }),
            latestScreenshot = ScreenshotManager.GetLatestScreenshotPath()
        });
    }

    private CommandResult Windows()
    {
        return new TextResult(string.Join("\n", screenshots.ListWindows().Select(w => w.Describe())));
    }

    private static CommandResult Focus(CommandParameters p)
    {
        var window = UiAutomationReader.ResolveWindow(p.Require("window", "要切到前台的窗口"), allowMinimized: true);
        InputSimulator.FocusWindow(window.Hwnd);
        return Ok(new { status = "ok", action = "focus", window = window.Title });
    }

    private CommandResult Screenshot(CommandParameters p)
    {
        string format = (p.Get("format") ?? "png").ToLowerInvariant();
        if (format is not ("png" or "jpg" or "jpeg"))
            throw new CommandException("format 必须为 png 或 jpg。");
        bool jpeg = format != "png";
        bool save = p.Flag("save");
        var capture = screenshots.Capture(p.Get("window"), p.Rect("region"), p.Int("maxWidth"),
            jpeg ? ImageFormat.Jpeg : ImageFormat.Png, save);
        if (!save) return new ImageResult(capture, jpeg);
        return Ok(new
        {
            status = "ok", file = capture.File, fileName = Path.GetFileName(capture.File),
            left = capture.Bounds.Left, top = capture.Bounds.Top,
            width = capture.Bounds.Width, height = capture.Bounds.Height,
            scale = Math.Round(capture.Scale, 4), window = capture.Window?.Title
        });
    }

    private static CommandResult Text(CommandParameters p)
    {
        var window = UiAutomationReader.ResolveWindow(p.Get("window"));
        string text = UiAutomationReader.ReadText(window, p.Flag("all"), p.Int("max", 20000, 100, 1_000_000));
        return new TextResult($"[{window.ProcessName}] {window.Title}\n{text}");
    }

    private static CommandResult Elements(CommandParameters p)
    {
        var window = UiAutomationReader.ResolveWindow(p.Get("window"));
        var elements = UiAutomationReader.FindElements(window, p.Get("name"), p.Get("type"), p.Int("limit", 200, 1, 2000));
        return new TextResult($"[{window.ProcessName}] {window.Title}\n" + string.Join("\n", elements.Select(e => e.Describe())));
    }

    private CommandResult Click(CommandParameters p)
    {
        string button = (p.Get("button") ?? "left").ToLowerInvariant();
        bool isDouble = p.Flag("double");
        Point? target = OptionalPoint(p, "x", "y");
        UiElement? element = null;
        if (p.Get("name") is { } name)
        {
            element = UiAutomationReader.FindOne(UiAutomationReader.ResolveWindow(p.Get("window")), name, p.Get("type"), p.Int("index"));
            target = element.Center;
        }
        InputSimulator.Click(target, button, isDouble);
        Point at = InputSimulator.CursorPosition();
        return Ok(new { status = "ok", action = isDouble ? "double_click" : "click", button, x = at.X, y = at.Y, element = element?.Describe() });
    }

    private CommandResult Move(CommandParameters p)
    {
        Point target = RequirePoint(p, "x", "y");
        InputSimulator.Move(target);
        return Ok(new { status = "ok", action = "move", x = target.X, y = target.Y });
    }

    private CommandResult Drag(CommandParameters p)
    {
        Point from = RequirePoint(p, "x1", "y1");
        Point to = RequirePoint(p, "x2", "y2");
        int durationMs = p.Int("duration", 200, 50, 3000);
        InputSimulator.Drag(from, to, durationMs, (p.Get("button") ?? "left").ToLowerInvariant());
        return Ok(new { status = "ok", action = "drag", from = new { x = from.X, y = from.Y }, to = new { x = to.X, y = to.Y }, durationMs });
    }

    private CommandResult Scroll(CommandParameters p)
    {
        int delta = p.Int("delta") ?? throw new CommandException("缺少参数 delta：正数向上、负数向下，一格约 120。");
        bool horizontal = p.Flag("horizontal");
        InputSimulator.Scroll(delta, horizontal, OptionalPoint(p, "x", "y"));
        return Ok(new { status = "ok", action = "scroll", delta, horizontal });
    }

    private static CommandResult Paste(CommandParameters p)
    {
        string text = p.Require("text", "要粘贴的文字，也可作为 text/plain 请求体发送");
        if (text.Contains('\0')) throw new CommandException("粘贴文本不能包含 NUL 字符。");
        string keys = p.Get("keys") ?? "ctrl+v";
        InputSimulator.PasteText(text, keys);
        return Ok(new { status = "ok", action = "paste", length = text.Length, keys });
    }

    private static CommandResult TypeText(CommandParameters p)
    {
        string text = p.Require("text", "要逐字输入的文字");
        InputSimulator.TypeText(text);
        return Ok(new { status = "ok", action = "type", length = text.Length });
    }

    private static CommandResult Key(CommandParameters p)
    {
        string combo = p.Get("combo", "keys", "key", "text")
            ?? throw new CommandException("缺少参数 combo，例如 enter、ctrl+c、ctrl+m,m。");
        InputSimulator.PressKeys(combo);
        return Ok(new { status = "ok", action = "key", combo });
    }

    private CommandResult Wait(CommandParameters p)
    {
        int timeoutSeconds = p.Int("timeout", 300, 1, 3600);
        int intervalMs = p.Int("interval", 1000, 200, 60000);
        Func<bool> satisfied = p.Get("text") is { } text ? TextCondition(p, text)
            : p.Flag("change") ? ChangeCondition(p)
            : throw new CommandException("wait 需要 text=... 或 change=1。");

        var elapsed = Stopwatch.StartNew();
        while (!satisfied())
        {
            if (elapsed.Elapsed.TotalSeconds >= timeoutSeconds)
                throw new CommandException($"等待 {timeoutSeconds} 秒后条件仍未满足。", 408);
            Thread.Sleep(intervalMs);
        }
        return Ok(new { status = "ok", action = "wait", seconds = Math.Round(elapsed.Elapsed.TotalSeconds, 1) });
    }

    private static Func<bool> TextCondition(CommandParameters p, string text)
    {
        bool gone = p.Flag("gone");
        bool includeOffscreen = p.Flag("all");
        string? selector = p.Get("window");
        return () =>
        {
            try
            {
                var window = UiAutomationReader.ResolveWindow(selector);
                bool present = UiAutomationReader.ReadText(window, includeOffscreen, 2_000_000)
                    .Contains(text, StringComparison.OrdinalIgnoreCase);
                return present != gone;
            }
            catch (ElementNotAvailableException)
            {
                return false; // The page re-rendered mid-read; try again next interval.
            }
        };
    }

    private Func<bool> ChangeCondition(CommandParameters p)
    {
        string? window = p.Get("window");
        Rectangle? region = p.Rect("region");
        double threshold = p.Double("threshold", 3);
        byte[] baseline = screenshots.SampleGray(window, region);
        return () => MeanDifference(baseline, screenshots.SampleGray(window, region)) > threshold;
    }

    private static double MeanDifference(byte[] a, byte[] b) =>
        a.Length != b.Length ? double.MaxValue : a.Zip(b, (x, y) => Math.Abs(x - y)).Average();

    private static CommandResult Sleep(CommandParameters p)
    {
        int ms = p.Int("ms", 1000, 0, 60000);
        Thread.Sleep(ms);
        return Ok(new { status = "ok", action = "sleep", ms });
    }

    private CommandResult Batch(JsonElement? steps)
    {
        if (steps is not { ValueKind: JsonValueKind.Array } array)
            throw new CommandException("POST /batch 的请求体必须是 JSON 数组，例如 [{\"op\":\"key\",\"combo\":\"ctrl+t\"}]。");
        var results = new List<object>();
        int count = array.GetArrayLength();
        int index = 0;
        foreach (var step in array.EnumerateArray())
        {
            index++;
            var p = new CommandParameters();
            p.MergeJson(step);
            string op = p.Get("op")?.ToLowerInvariant() ?? throw new CommandException($"第 {index} 步缺少 op。");
            if (op == "batch") throw new CommandException("batch 不能嵌套。");
            CommandResult result;
            try
            {
                result = Execute(op, p);
            }
            catch (Exception ex)
            {
                int status = ex switch { CommandException c => c.StatusCode, WindowSelectionException w => w.StatusCode, _ => 500 };
                if (status == 500) Logger.Error($"batch 第 {index} 步 {op} 失败", ex);
                throw new CommandException($"第 {index} 步 {op} 失败：{ex.Message}", status, new { completed = results });
            }
            switch (result)
            {
                case ImageResult image when index == count:
                    return image;
                case ImageResult:
                    throw new CommandException($"第 {index} 步：图片只能作为最后一步返回，中间截图请加 save=1。", 400, new { completed = results });
                case TextResult text:
                    results.Add(new { op, text = text.Text });
                    break;
                case JsonResult json:
                    results.Add(json.Data);
                    break;
            }
        }
        return Ok(new { status = "ok", results });
    }

    private Point? OptionalPoint(CommandParameters p, string xKey, string yKey)
    {
        int? x = p.Int(xKey), y = p.Int(yKey);
        if (x == null && y == null) return null;
        if (x == null || y == null) throw new CommandException($"{xKey} 和 {yKey} 必须同时提供。");
        if (!p.Flag("shot")) return new Point(x.Value, y.Value);
        var last = screenshots.LastCapture ?? throw new CommandException("shot=1 需要先截一张图。");
        return last.ToDesktop(x.Value, y.Value);
    }

    private Point RequirePoint(CommandParameters p, string xKey, string yKey) =>
        OptionalPoint(p, xKey, yKey) ?? throw new CommandException($"缺少坐标 {xKey}、{yKey}。");
}
