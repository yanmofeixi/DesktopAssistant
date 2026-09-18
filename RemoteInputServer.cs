using System.Drawing.Imaging;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Text.Json;

namespace DesktopAssistant
{
    public class RemoteInputServer
    {
        public const int DefaultPort = 18888;
        private readonly int port;
        private HttpListener? listener;
        private CancellationTokenSource? cts;
        private Task? listenTask;
        private bool isRunning;
        private readonly ScreenshotManager screenshotManager;
        private readonly DateTime startTime = DateTime.Now;
        private static readonly object inputLock = new();

        public bool IsRunning => isRunning;
        public int Port => port;

        public RemoteInputServer(ScreenshotManager screenshotManager, int port = DefaultPort)
        {
            this.screenshotManager = screenshotManager;
            this.port = port;
        }

        #region Win32 API Definitions

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr OpenInputDesktop(uint dwFlags, bool fInherit, uint dwDesiredAccess);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool CloseDesktop(IntPtr hDesktop);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetThreadDesktop(IntPtr hDesktop);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetThreadDesktop(int dwThreadId);

        [DllImport("kernel32.dll")]
        private static extern int GetCurrentThreadId();

        private const uint DESKTOP_MAXIMUM_ALLOWED = 0x01FF;

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, UIntPtr dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
        private const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        private const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
        private const uint MOUSEEVENTF_WHEEL = 0x0800;
        private const uint MOUSEEVENTF_HWHEEL = 0x1000;

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public InputUnion u;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)]
            public MOUSEINPUT mi;
            [FieldOffset(0)]
            public KEYBDINPUT ki;
            [FieldOffset(0)]
            public HARDWAREINPUT hi;
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

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        private const uint INPUT_KEYBOARD = 1;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_UNICODE = 0x0004;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        #endregion

        public void Start()
        {
            if (isRunning) return;

            try
            {
                listener = new HttpListener();
                listener.Prefixes.Add($"http://127.0.0.1:{port}/");
                listener.Start();

                cts = new CancellationTokenSource();
                isRunning = true;
                listenTask = Task.Run(() => ListenLoop(cts.Token));

                Logger.Info($"RemoteInputServer 已启动，监听端口: {port}");
            }
            catch (Exception ex)
            {
                Logger.Error($"RemoteInputServer 启动失败 (端口 {port})", ex);
                isRunning = false;
            }
        }

        public void Stop()
        {
            if (!isRunning) return;

            try
            {
                isRunning = false;
                cts?.Cancel();
                listener?.Stop();
                listener?.Close();
                listener = null;
                Logger.Info("RemoteInputServer 已停止");
            }
            catch (Exception ex)
            {
                Logger.Error("RemoteInputServer 停止时出错", ex);
            }
        }

        private async Task ListenLoop(CancellationToken token)
        {
            while (!token.IsCancellationRequested && listener != null && listener.IsListening)
            {
                try
                {
                    var context = await listener.GetContextAsync();
                    _ = Task.Run(() => HandleRequest(context));
                }
                catch (HttpListenerException) when (!isRunning)
                {
                    break;
                }
                catch (ObjectDisposedException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    if (isRunning)
                    {
                        Logger.Error("RemoteInputServer 接受请求出错", ex);
                    }
                }
            }
        }

        private void HandleRequest(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;

            res.Headers.Add("Access-Control-Allow-Origin", "*");
            res.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
            res.Headers.Add("Access-Control-Allow-Headers", "Content-Type");

            if (req.HttpMethod.Equals("OPTIONS", StringComparison.OrdinalIgnoreCase))
            {
                res.StatusCode = (int)HttpStatusCode.NoContent;
                res.Close();
                return;
            }

            try
            {
                string path = req.Url?.AbsolutePath.TrimEnd('/').ToLowerInvariant() ?? "";

                switch (path)
                {
                    case "":
                    case "/help":
                        HandleHelp(res);
                        break;
                    case "/windows":
                        SendJson(res, new { status = "ok", windows = screenshotManager.ListWindows().Select(w => w.Metadata) });
                        break;
                    case "/info":
                    case "/status":
                        HandleInfo(res);
                        break;
                    case "/click":
                        HandleClick(req, res);
                        break;
                    case "/move":
                        HandleMove(req, res);
                        break;
                    case "/drag":
                        HandleDrag(req, res);
                        break;
                    case "/scroll":
                        HandleScroll(req, res);
                        break;
                    case "/type":
                        HandleType(req, res);
                        break;
                    case "/paste":
                        HandlePaste(req, res);
                        break;
                    case "/key":
                        HandleKey(req, res);
                        break;
                    case "/screenshot":
                    case "/screenshot/latest":
                    case "/screenshot/latest.png":
                        HandleScreenshot(req, res);
                        break;
                    default:
                        SendJson(res, new { status = "error", message = $"未找到端点: {path}" }, HttpStatusCode.NotFound);
                        break;
                }
            }
            catch (WindowSelectionException ex)
            {
                SendJson(res, new { status = "error", message = ex.Message, windows = ex.Matches.Select(w => w.Metadata) },
                    (HttpStatusCode)ex.StatusCode);
            }
            catch (Exception ex)
            {
                Logger.Error($"处理 HTTP 请求异常: {req.Url}", ex);
                SendJson(res, new { status = "error", message = ex.Message }, HttpStatusCode.InternalServerError);
            }
            finally
            {
                try { res.Close(); } catch { }
            }
        }

        #region Request Handlers

        private void HandleHelp(HttpListenerResponse res)
        {
            var help = new
            {
                service = "DesktopAssistant RemoteInputServer",
                port = port,
                endpoints = new[]
                {
                    new { method = "GET/POST", path = "/click", params_ = "x, y, button(left|right|middle), double(true|false)", desc = "鼠标点击" },
                    new { method = "GET/POST", path = "/move", params_ = "x, y", desc = "移动鼠标光标" },
                    new { method = "GET/POST", path = "/drag", params_ = "x1, y1, x2, y2, duration(ms), button", desc = "鼠标拖拽" },
                    new { method = "GET/POST", path = "/scroll", params_ = "delta(正往上/负往下), horizontal(true|false), x, y", desc = "鼠标滚轮" },
                    new { method = "POST", path = "/paste", params_ = "UTF-8 text/plain Body", desc = "默认文本输入：通过桌面剪贴板一次性粘贴长文本（覆盖剪贴板）" },
                    new { method = "GET/POST", path = "/type", params_ = "text (Query 或 POST Body)", desc = "逐字模拟 Unicode 按键，仅在需要按键事件或目标不支持粘贴时使用" },
                    new { method = "GET/POST", path = "/key", params_ = "combo (如 enter, esc, ctrl+c, alt+tab, win+d)", desc = "按键与组合快捷键" },
                    new { method = "GET", path = "/screenshot", params_ = "window(active|进程名|标题|0x句柄), save(true|false), format(png|jpg|json), base64(true|false)", desc = "截取可见全屏或指定窗口区域；默认返回 PNG，save=true 返回文件路径和坐标" },
                    new { method = "GET", path = "/screenshot/latest", params_ = "", desc = "直接返回最新一张截图 PNG 二进制流" },
                    new { method = "GET", path = "/windows", params_ = "", desc = "列出窗口句柄、标题、程序名和屏幕坐标" },
                    new { method = "GET", path = "/info", params_ = "", desc = "获取屏幕分辨率、光标当前坐标、服务状态" }
                }
            };
            SendJson(res, help);
        }

        private void HandleInfo(HttpListenerResponse res)
        {
            POINT pt;
            GetCursorPos(out pt);
            var virtualScreen = SystemInformation.VirtualScreen;
            var primaryScreen = Screen.PrimaryScreen?.Bounds ?? Rectangle.Empty;

            var info = new
            {
                status = "ok",
                uptimeSeconds = (int)(DateTime.Now - startTime).TotalSeconds,
                cursor = new { x = pt.X, y = pt.Y },
                screen = new
                {
                    virtualLeft = virtualScreen.Left,
                    virtualTop = virtualScreen.Top,
                    virtualWidth = virtualScreen.Width,
                    virtualHeight = virtualScreen.Height,
                    primaryWidth = primaryScreen.Width,
                    primaryHeight = primaryScreen.Height
                },
                screens = Screen.AllScreens.Select(s => new
                {
                    device = s.DeviceName,
                    primary = s.Primary,
                    bounds = new { x = s.Bounds.X, y = s.Bounds.Y, width = s.Bounds.Width, height = s.Bounds.Height }
                }),
                latestScreenshot = ScreenshotManager.GetLatestScreenshotPath()
            };
            SendJson(res, info);
        }

        private void HandleClick(HttpListenerRequest req, HttpListenerResponse res)
        {
            var q = req.QueryString;
            int? x = int.TryParse(q["x"], out int px) ? px : null;
            int? y = int.TryParse(q["y"], out int py) ? py : null;
            string button = q["button"]?.ToLowerInvariant() ?? "left";
            bool isDouble = q["double"]?.Equals("true", StringComparison.OrdinalIgnoreCase) == true || q["double"] == "1";

            EnsureInputDesktop(() =>
            {
                if (x.HasValue && y.HasValue)
                {
                    SetCursorPos(x.Value, y.Value);
                    Thread.Sleep(25);
                }

                uint downFlag = MOUSEEVENTF_LEFTDOWN;
                uint upFlag = MOUSEEVENTF_LEFTUP;

                if (button == "right")
                {
                    downFlag = MOUSEEVENTF_RIGHTDOWN;
                    upFlag = MOUSEEVENTF_RIGHTUP;
                }
                else if (button == "middle")
                {
                    downFlag = MOUSEEVENTF_MIDDLEDOWN;
                    upFlag = MOUSEEVENTF_MIDDLEUP;
                }

                mouse_event(downFlag, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(30);
                mouse_event(upFlag, 0, 0, 0, UIntPtr.Zero);

                if (isDouble)
                {
                    Thread.Sleep(80);
                    mouse_event(downFlag, 0, 0, 0, UIntPtr.Zero);
                    Thread.Sleep(30);
                    mouse_event(upFlag, 0, 0, 0, UIntPtr.Zero);
                }
            });

            GetCursorPos(out POINT current);
            SendJson(res, new { status = "ok", action = isDouble ? "double_click" : "click", button, x = current.X, y = current.Y });
        }

        private void HandleMove(HttpListenerRequest req, HttpListenerResponse res)
        {
            var q = req.QueryString;
            if (!int.TryParse(q["x"], out int x) || !int.TryParse(q["y"], out int y))
            {
                SendJson(res, new { status = "error", message = "缺少 x 或 y 坐标参数" }, HttpStatusCode.BadRequest);
                return;
            }

            EnsureInputDesktop(() =>
            {
                SetCursorPos(x, y);
            });

            SendJson(res, new { status = "ok", action = "move", x, y });
        }

        private void HandleDrag(HttpListenerRequest req, HttpListenerResponse res)
        {
            var q = req.QueryString;
            if (!int.TryParse(q["x1"], out int x1) || !int.TryParse(q["y1"], out int y1) ||
                !int.TryParse(q["x2"], out int x2) || !int.TryParse(q["y2"], out int y2))
            {
                SendJson(res, new { status = "error", message = "缺少 x1, y1, x2, y2 坐标参数" }, HttpStatusCode.BadRequest);
                return;
            }

            int durationMs = int.TryParse(q["duration"], out int d) ? Math.Clamp(d, 50, 3000) : 200;
            string button = q["button"]?.ToLowerInvariant() ?? "left";

            uint downFlag = button == "right" ? MOUSEEVENTF_RIGHTDOWN : MOUSEEVENTF_LEFTDOWN;
            uint upFlag = button == "right" ? MOUSEEVENTF_RIGHTUP : MOUSEEVENTF_LEFTUP;

            EnsureInputDesktop(() =>
            {
                SetCursorPos(x1, y1);
                Thread.Sleep(30);
                mouse_event(downFlag, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(30);

                int steps = Math.Max(10, durationMs / 15);
                int sleepPerStep = durationMs / steps;

                for (int i = 1; i <= steps; i++)
                {
                    double t = (double)i / steps;
                    int curX = (int)(x1 + (x2 - x1) * t);
                    int curY = (int)(y1 + (y2 - y1) * t);
                    SetCursorPos(curX, curY);
                    Thread.Sleep(sleepPerStep);
                }

                SetCursorPos(x2, y2);
                Thread.Sleep(30);
                mouse_event(upFlag, 0, 0, 0, UIntPtr.Zero);
            });

            SendJson(res, new { status = "ok", action = "drag", from = new { x = x1, y = y1 }, to = new { x = x2, y = y2 }, durationMs });
        }

        private void HandleScroll(HttpListenerRequest req, HttpListenerResponse res)
        {
            var q = req.QueryString;
            if (!int.TryParse(q["delta"], out int delta))
            {
                SendJson(res, new { status = "error", message = "缺少 delta 参数 (正数向上滚，负数向下滚，通常 ±120)" }, HttpStatusCode.BadRequest);
                return;
            }

            int? x = int.TryParse(q["x"], out int px) ? px : null;
            int? y = int.TryParse(q["y"], out int py) ? py : null;
            bool horizontal = q["horizontal"]?.Equals("true", StringComparison.OrdinalIgnoreCase) == true;

            EnsureInputDesktop(() =>
            {
                if (x.HasValue && y.HasValue)
                {
                    SetCursorPos(x.Value, y.Value);
                    Thread.Sleep(20);
                }

                uint flag = horizontal ? MOUSEEVENTF_HWHEEL : MOUSEEVENTF_WHEEL;
                mouse_event(flag, 0, 0, delta, UIntPtr.Zero);
            });

            SendJson(res, new { status = "ok", action = "scroll", delta, horizontal });
        }

        private void HandleType(HttpListenerRequest req, HttpListenerResponse res)
        {
            string text = "";

            if (req.HttpMethod.Equals("POST", StringComparison.OrdinalIgnoreCase))
            {
                using var reader = new StreamReader(req.InputStream, req.ContentEncoding ?? Encoding.UTF8);
                text = reader.ReadToEnd();
            }

            if (string.IsNullOrEmpty(text))
            {
                text = req.QueryString["text"] ?? "";
            }

            if (string.IsNullOrEmpty(text))
            {
                SendJson(res, new { status = "error", message = "未提供需要输入的文本" }, HttpStatusCode.BadRequest);
                return;
            }

            EnsureInputDesktop(() =>
            {
                SendUnicodeText(text);
            });

            SendJson(res, new { status = "ok", action = "type", length = text.Length });
        }

        private void HandlePaste(HttpListenerRequest req, HttpListenerResponse res)
        {
            if (!req.HttpMethod.Equals("POST", StringComparison.OrdinalIgnoreCase))
            {
                res.Headers["Allow"] = "POST";
                SendJson(res, new { status = "error", message = "请使用 POST 发送 UTF-8 文本请求体。" }, HttpStatusCode.MethodNotAllowed);
                return;
            }

            using var reader = new StreamReader(req.InputStream, req.ContentEncoding ?? Encoding.UTF8);
            string text = reader.ReadToEnd();
            if (string.IsNullOrEmpty(text) || text.Contains('\0'))
            {
                SendJson(res, new { status = "error", message = "粘贴文本不能为空或包含 NUL 字符。" }, HttpStatusCode.BadRequest);
                return;
            }

            // Clipboard requires STA and must run in the logged-in desktop process,
            // not the SSH session's separate clipboard. Serialize with other input.
            Exception? failure = null;
            var thread = new Thread(() =>
            {
                try
                {
                    EnsureInputDesktop(() =>
                    {
                        var data = new DataObject();
                        data.SetText(text, TextDataFormat.UnicodeText);
                        Clipboard.SetDataObject(data, true, 10, 50);
                        ExecuteKeyCombo("ctrl+v");
                        // Give the target time to consume Ctrl+V before the next input.
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

            SendJson(res, new { status = "ok", action = "paste", length = text.Length });
        }

        private void HandleKey(HttpListenerRequest req, HttpListenerResponse res)
        {
            string combo = req.QueryString["combo"] ?? req.QueryString["key"] ?? "";

            if (string.IsNullOrEmpty(combo) && req.HttpMethod.Equals("POST", StringComparison.OrdinalIgnoreCase))
            {
                using var reader = new StreamReader(req.InputStream, req.ContentEncoding ?? Encoding.UTF8);
                combo = reader.ReadToEnd().Trim();
            }

            if (string.IsNullOrEmpty(combo))
            {
                SendJson(res, new { status = "error", message = "缺少 combo 或 key 参数 (如 enter, ctrl+c, alt+tab, win+d)" }, HttpStatusCode.BadRequest);
                return;
            }

            bool success = false;
            string errorMsg = "";

            EnsureInputDesktop(() =>
            {
                try
                {
                    ExecuteKeyCombo(combo);
                    success = true;
                }
                catch (Exception ex)
                {
                    errorMsg = ex.Message;
                }
            });

            if (success)
            {
                SendJson(res, new { status = "ok", action = "key", combo });
            }
            else
            {
                SendJson(res, new { status = "error", message = errorMsg }, HttpStatusCode.BadRequest);
            }
        }

        private void HandleScreenshot(HttpListenerRequest req, HttpListenerResponse res)
        {
            var q = req.QueryString;
            bool save = q["save"]?.Equals("true", StringComparison.OrdinalIgnoreCase) == true || q["save"] == "1";
            string format = q["format"]?.ToLowerInvariant() ?? "png";
            if (format is not ("png" or "jpg" or "jpeg" or "json"))
            {
                SendJson(res, new { status = "error", message = "format 必须为 png、jpg、jpeg 或 json。" }, HttpStatusCode.BadRequest);
                return;
            }
            bool base64 = q["base64"]?.Equals("true", StringComparison.OrdinalIgnoreCase) == true || q["base64"] == "1" || format == "json";
            bool jpeg = format is "jpg" or "jpeg";
            var capture = screenshotManager.Capture(q["window"], jpeg ? ImageFormat.Jpeg : ImageFormat.Png, save);
            var bounds = capture.Bounds;
            if (save || base64)
            {
                SendJson(res, new
                {
                    status = "ok", file = capture.File,
                    fileName = capture.File == null ? null : Path.GetFileName(capture.File),
                    timestamp = DateTime.Now.ToString("s"), format = jpeg ? "jpg" : "png",
                    left = bounds.Left, top = bounds.Top, width = bounds.Width, height = bounds.Height,
                    window = capture.Window?.Metadata,
                    base64 = base64 ? Convert.ToBase64String(capture.Bytes) : null
                });
                return;
            }
            res.StatusCode = (int)HttpStatusCode.OK;
            res.ContentType = jpeg ? "image/jpeg" : "image/png";
            res.ContentLength64 = capture.Bytes.Length;
            res.Headers["X-Screenshot-Left"] = bounds.Left.ToString(System.Globalization.CultureInfo.InvariantCulture);
            res.Headers["X-Screenshot-Top"] = bounds.Top.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (!req.HttpMethod.Equals("HEAD", StringComparison.OrdinalIgnoreCase))
                res.OutputStream.Write(capture.Bytes, 0, capture.Bytes.Length);
        }

        #endregion

        #region Low-level Input Implementation

        private static void EnsureInputDesktop(Action action)
        {
            lock (inputLock) EnsureInputDesktopCore(action);
        }

        private static void EnsureInputDesktopCore(Action action)
        {
            IntPtr hOldDesktop = IntPtr.Zero;
            IntPtr hInputDesktop = IntPtr.Zero;
            bool switched = false;

            try
            {
                hInputDesktop = OpenInputDesktop(0, false, DESKTOP_MAXIMUM_ALLOWED);
                if (hInputDesktop != IntPtr.Zero)
                {
                    hOldDesktop = GetThreadDesktop(GetCurrentThreadId());
                    switched = SetThreadDesktop(hInputDesktop);
                }
                action();
            }
            finally
            {
                if (switched && hOldDesktop != IntPtr.Zero)
                {
                    SetThreadDesktop(hOldDesktop);
                }
                if (hInputDesktop != IntPtr.Zero)
                {
                    CloseDesktop(hInputDesktop);
                }
            }
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
                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(),
                    $"SendInput 只发送了 {sent}/{inputs.Length} 个输入事件，请确认桌面已解锁且目标窗口可接收输入。");
        }

        private static void ExecuteKeyCombo(string combo)
        {
            var parts = combo.Split(new[] { '+', '-' }, StringSplitOptions.RemoveEmptyEntries)
                             .Select(p => p.Trim().ToLowerInvariant())
                             .ToList();

            if (parts.Count == 0) return;

            var modifierVks = new List<ushort>();
            var regularVks = new List<ushort>();

            foreach (var part in parts)
            {
                if (IsModifier(part, out ushort modVk))
                {
                    modifierVks.Add(modVk);
                }
                else if (TryGetVk(part, out ushort vk))
                {
                    regularVks.Add(vk);
                }
                else
                {
                    throw new ArgumentException($"无法识别的按键名称: '{part}'");
                }
            }

            var inputs = new List<INPUT>();

            // 1. 按下所有修饰键
            foreach (var mod in modifierVks)
            {
                inputs.Add(MakeKeyInput(mod, false));
            }

            // 2. 按下并释放常规键
            foreach (var reg in regularVks)
            {
                inputs.Add(MakeKeyInput(reg, false));
            }
            foreach (var reg in regularVks)
            {
                inputs.Add(MakeKeyInput(reg, true));
            }

            // 3. 逆序释放所有修饰键
            for (int i = modifierVks.Count - 1; i >= 0; i--)
            {
                inputs.Add(MakeKeyInput(modifierVks[i], true));
            }

            if (inputs.Count > 0)
            {
                var arr = inputs.ToArray();
                SendChecked(arr);
            }
        }

        private static INPUT MakeKeyInput(ushort vk, bool keyUp)
        {
            uint flags = keyUp ? KEYEVENTF_KEYUP : 0;
            if (IsExtendedKey(vk))
            {
                flags |= KEYEVENTF_EXTENDEDKEY;
            }

            return new INPUT
            {
                type = INPUT_KEYBOARD,
                u = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = vk,
                        wScan = 0,
                        dwFlags = flags
                    }
                }
            };
        }

        private static bool IsModifier(string name, out ushort vk)
        {
            switch (name)
            {
                case "ctrl":
                case "control":
                    vk = 0x11; // VK_CONTROL
                    return true;
                case "alt":
                case "menu":
                    vk = 0x12; // VK_MENU
                    return true;
                case "shift":
                    vk = 0x10; // VK_SHIFT
                    return true;
                case "win":
                case "windows":
                case "meta":
                case "cmd":
                case "command":
                    vk = 0x5B; // VK_LWIN
                    return true;
                default:
                    vk = 0;
                    return false;
            }
        }

        private static bool TryGetVk(string name, out ushort vk)
        {
            switch (name)
            {
                case "enter":
                case "return":
                    vk = 0x0D; return true;
                case "esc":
                case "escape":
                    vk = 0x1B; return true;
                case "tab":
                    vk = 0x09; return true;
                case "backspace":
                case "back":
                    vk = 0x08; return true;
                case "delete":
                case "del":
                    vk = 0x2E; return true;
                case "space":
                case "spacebar":
                    vk = 0x20; return true;
                case "insert":
                case "ins":
                    vk = 0x2D; return true;
                case "up":
                    vk = 0x26; return true;
                case "down":
                    vk = 0x28; return true;
                case "left":
                    vk = 0x25; return true;
                case "right":
                    vk = 0x27; return true;
                case "home":
                    vk = 0x24; return true;
                case "end":
                    vk = 0x23; return true;
                case "pageup":
                case "pgup":
                    vk = 0x21; return true;
                case "pagedown":
                case "pgdn":
                    vk = 0x22; return true;
                case "capslock":
                case "caps":
                    vk = 0x14; return true;
                case "printscreen":
                case "prtsc":
                    vk = 0x2C; return true;
                default:
                    // F1 - F24
                    if (name.StartsWith("f") && int.TryParse(name[1..], out int fNum) && fNum >= 1 && fNum <= 24)
                    {
                        vk = (ushort)(0x70 + (fNum - 1)); // VK_F1 = 0x70
                        return true;
                    }
                    // 单个字母或数字
                    if (name.Length == 1)
                    {
                        char c = char.ToUpperInvariant(name[0]);
                        if ((c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9'))
                        {
                            vk = (ushort)c;
                            return true;
                        }
                    }
                    vk = 0;
                    return false;
            }
        }

        private static bool IsExtendedKey(ushort vk)
        {
            // 扩展键标志需要用于方向键、编辑键组、Windows键等
            return vk switch
            {
                0x21 or 0x22 or 0x23 or 0x24 or 0x25 or 0x26 or 0x27 or 0x28 or // PgUp, PgDn, End, Home, Left, Up, Right, Down
                0x2D or 0x2E or // Insert, Delete
                0x5B or 0x5C => true, // LWin, RWin
                _ => false
            };
        }

        #endregion

        #region Helper Output Methods

        private static void SendJson(HttpListenerResponse res, object data, HttpStatusCode code = HttpStatusCode.OK)
        {
            res.StatusCode = (int)code;
            res.ContentType = "application/json; charset=utf-8";
            string json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
            byte[] bytes = Encoding.UTF8.GetBytes(json);
            res.ContentLength64 = bytes.Length;
            res.OutputStream.Write(bytes, 0, bytes.Length);
        }

        private static void SendFile(HttpListenerResponse res, string filePath, string contentType)
        {
            res.StatusCode = (int)HttpStatusCode.OK;
            res.ContentType = contentType;
            using var fs = File.OpenRead(filePath);
            res.ContentLength64 = fs.Length;
            fs.CopyTo(res.OutputStream);
        }

        #endregion
    }
}
