namespace DesktopAssistant
{
    /// <summary>
    /// 简单的日志系统，用于调试
    /// </summary>
    public static class Logger
    {
        private static readonly string logFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DesktopAssistant", "Logs");
        
        private static readonly object lockObj = new();
        private static string? currentLogFile;

        static Logger()
        {
            Directory.CreateDirectory(logFolder);
            currentLogFile = Path.Combine(logFolder, $"log_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
        }

        /// <summary>
        /// 获取日志文件夹路径
        /// </summary>
        public static string LogFolder => logFolder;

        /// <summary>
        /// 写入日志
        /// </summary>
        public static void Log(string message)
        {
            try
            {
                var logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
                lock (lockObj)
                {
                    File.AppendAllText(currentLogFile!, logLine + Environment.NewLine);
                }
                System.Diagnostics.Debug.WriteLine(logLine);
            }
            catch { }
        }

        /// <summary>
        /// 写入调试信息
        /// </summary>
        public static void Debug(string message) => Log($"[DEBUG] {message}");

        /// <summary>
        /// 写入信息
        /// </summary>
        public static void Info(string message) => Log($"[INFO] {message}");

        /// <summary>
        /// 写入警告
        /// </summary>
        public static void Warn(string message) => Log($"[WARN] {message}");

        /// <summary>
        /// 写入错误
        /// </summary>
        public static void Error(string message, Exception? ex = null)
        {
            if (ex != null)
                Log($"[ERROR] {message}: {ex.Message}\n{ex.StackTrace}");
            else
                Log($"[ERROR] {message}");
        }
    }
}
