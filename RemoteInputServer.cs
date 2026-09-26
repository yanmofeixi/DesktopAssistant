using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAssistant
{
    /// <summary>Loopback HTTP front end: GET/POST /&lt;op&gt; maps to <see cref="RemoteCommands"/>.</summary>
    public class RemoteInputServer
    {
        public const int DefaultPort = 18888;

        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            // Keep Chinese readable and omit nulls: responses are read by humans and LLM agents over SSH.
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };

        private readonly int port;
        private readonly RemoteCommands commands;
        private HttpListener? listener;
        private CancellationTokenSource? cts;
        private bool isRunning;

        public bool IsRunning => isRunning;
        public int Port => port;

        public RemoteInputServer(ScreenshotManager screenshotManager, int port = DefaultPort)
        {
            this.port = port;
            commands = new RemoteCommands(screenshotManager, port);
        }

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
                _ = Task.Run(() => ListenLoop(cts.Token));

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
                    if (isRunning) Logger.Error("RemoteInputServer 接受请求出错", ex);
                }
            }
        }

        private void HandleRequest(HttpListenerContext context)
        {
            var req = context.Request;
            var res = context.Response;
            try
            {
                // Any web page can send simple requests to localhost. Browsers tag them with
                // Origin / Sec-Fetch-* headers, which curl and PowerShell never send.
                if (req.Headers["Origin"] != null || req.Headers["Sec-Fetch-Mode"] != null)
                {
                    WriteJson(res, new { status = "error", message = "拒绝来自浏览器网页的请求。" }, HttpStatusCode.Forbidden);
                    return;
                }

                string op = (req.Url?.AbsolutePath.Trim('/') ?? "").ToLowerInvariant();
                var parameters = CommandParameters.FromQuery(req.QueryString);
                JsonElement? batchSteps = req.HasEntityBody ? ApplyBody(op, ReadBody(req), req.ContentType, parameters) : null;
                WriteResult(res, commands.Execute(op, parameters, batchSteps),
                    req.HttpMethod.Equals("HEAD", StringComparison.OrdinalIgnoreCase));
            }
            catch (CommandException ex)
            {
                WriteJson(res, new { status = "error", message = ex.Message, details = ex.Details }, (HttpStatusCode)ex.StatusCode);
            }
            catch (WindowSelectionException ex)
            {
                WriteJson(res, new { status = "error", message = ex.Message, windows = ex.Matches.Select(w => w.Describe()) },
                    (HttpStatusCode)ex.StatusCode);
            }
            catch (Exception ex)
            {
                Logger.Error($"处理 HTTP 请求异常: {req.Url}", ex);
                WriteJson(res, new { status = "error", message = ex.Message }, HttpStatusCode.InternalServerError);
            }
            finally
            {
                try { res.Close(); } catch { }
            }
        }

        private static string ReadBody(HttpListenerRequest req)
        {
            using var reader = new StreamReader(req.InputStream, req.ContentEncoding ?? Encoding.UTF8);
            return reader.ReadToEnd();
        }

        /// <summary>
        /// JSON objects become parameters and a JSON array is the /batch step list. Anything else is the
        /// "text" parameter; text/plain is never parsed as JSON so pasted code starting with '{' stays text.
        /// </summary>
        private static JsonElement? ApplyBody(string op, string body, string? contentType, CommandParameters parameters)
        {
            bool declaredText = contentType?.Contains("text/plain", StringComparison.OrdinalIgnoreCase) == true;
            bool declaredJson = contentType?.Contains("json", StringComparison.OrdinalIgnoreCase) == true;
            string trimmed = body.TrimStart();
            bool looksJson = trimmed.StartsWith('{') || trimmed.StartsWith('[');
            if (declaredText || !(declaredJson || looksJson))
            {
                parameters.SetIfMissing("text", body);
                return null;
            }

            JsonElement root;
            try
            {
                using var document = JsonDocument.Parse(body);
                root = document.RootElement.Clone();
            }
            catch (JsonException ex)
            {
                throw new CommandException($"JSON 请求体解析失败：{ex.Message}");
            }
            if (root.ValueKind == JsonValueKind.Array)
            {
                if (op != "batch") throw new CommandException("JSON 数组请求体只用于 /batch。");
                return root;
            }
            parameters.MergeJson(root);
            return null;
        }

        private static void WriteResult(HttpListenerResponse res, CommandResult result, bool headOnly)
        {
            switch (result)
            {
                case JsonResult json:
                    WriteJson(res, json.Data);
                    break;
                case TextResult text:
                    WriteBytes(res, Encoding.UTF8.GetBytes(text.Text.TrimEnd('\n') + "\n"), "text/plain; charset=utf-8", headOnly);
                    break;
                case ImageResult image:
                    var capture = image.Capture;
                    res.Headers["X-Screenshot-Left"] = capture.Bounds.Left.ToString(CultureInfo.InvariantCulture);
                    res.Headers["X-Screenshot-Top"] = capture.Bounds.Top.ToString(CultureInfo.InvariantCulture);
                    res.Headers["X-Screenshot-Scale"] = capture.Scale.ToString("0.####", CultureInfo.InvariantCulture);
                    WriteBytes(res, capture.Bytes, image.Jpeg ? "image/jpeg" : "image/png", headOnly);
                    break;
            }
        }

        private static void WriteJson(HttpListenerResponse res, object data, HttpStatusCode code = HttpStatusCode.OK)
        {
            res.StatusCode = (int)code;
            WriteBytes(res, Encoding.UTF8.GetBytes(JsonSerializer.Serialize(data, JsonOptions) + "\n"), "application/json; charset=utf-8", false);
        }

        private static void WriteBytes(HttpListenerResponse res, byte[] bytes, string contentType, bool headOnly)
        {
            res.ContentType = contentType;
            res.ContentLength64 = bytes.Length;
            if (!headOnly) res.OutputStream.Write(bytes, 0, bytes.Length);
        }
    }
}
