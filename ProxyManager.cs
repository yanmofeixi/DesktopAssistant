using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace DesktopAssistant
{
    public static class ProxyManager
    {
        [DllImport("wininet.dll", SetLastError = true)]
        private static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int dwBufferLength);

        private const int INTERNET_OPTION_SETTINGS_CHANGED = 39;
        private const int INTERNET_OPTION_REFRESH = 37;

        private const string RegistryKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Internet Settings";

        public static bool IsProxyEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath);
                if (key != null)
                {
                    var result = key.GetValue("ProxyEnable");
                    return result != null && (int)result == 1;
                }
            }
            catch
            {
                // Ignore
            }
            return false;
        }

        public static void SetProxy(bool enable, string proxyServer = "127.0.0.1:8899")
        {
            try
            {
                if (enable)
                {
                    RunWhistleCommand("start");
                    SetWinHttpProxy(true, proxyServer);
                }
                else
                {
                    RunWhistleCommand("stop");
                    SetWinHttpProxy(false);
                }

                using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, true);
                if (key != null)
                {
                    key.SetValue("ProxyEnable", enable ? 1 : 0, RegistryValueKind.DWord);
                    if (enable)
                    {
                        key.SetValue("ProxyServer", proxyServer, RegistryValueKind.String);
                        // 强制排除本地地址，避免环回问题导致 Whistle 监控页面打不开
                        key.SetValue("ProxyOverride", "localhost;127.*;<local>", RegistryValueKind.String);
                    }
                }

                // 通知系统代理设置已更改 (WinINet)
                InternetSetOption(IntPtr.Zero, INTERNET_OPTION_SETTINGS_CHANGED, IntPtr.Zero, 0);
                InternetSetOption(IntPtr.Zero, INTERNET_OPTION_REFRESH, IntPtr.Zero, 0);
            }
            catch (Exception ex)
            {
                Logger.Log($"设置代理失败: {ex.Message}");
            }
        }

        private static void RunWhistleCommand(string command)
        {
            try
            {
                var processInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c w2 {command}",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                };
                System.Diagnostics.Process.Start(processInfo);
            }
            catch (Exception ex)
            {
                Logger.Log($"执行 whistle 命令 {command} 失败: {ex.Message}");
            }
        }

        private static void SetWinHttpProxy(bool enable, string proxyServer = "127.0.0.1:8899")
        {
            try
            {
                string args = enable 
                    ? $"winhttp set proxy \"{proxyServer}\" \"localhost;127.*;<local>\"" 
                    : "winhttp reset proxy";

                var processInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "netsh",
                    Arguments = args,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                };
                var process = System.Diagnostics.Process.Start(processInfo);
                process?.WaitForExit();
            }
            catch (Exception ex)
            {
                Logger.Log($"设置 WinHTTP 代理失败: {ex.Message}");
            }
        }
    }
}
