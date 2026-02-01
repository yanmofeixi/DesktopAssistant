using System.Reflection;
using System.Runtime.InteropServices;
using DesktopAssistant;

namespace DesktopAssistant
{
    internal class Program : Form
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private static HotkeyManager? hotkeyManager;
        private static ReminderManager? reminderManager;
        private static GameAhkManager? gameAhkManager;
        private static ToolStripMenuItem? startProxyMenuItem;
        private static ToolStripMenuItem? stopProxyMenuItem;

        private static readonly Form contextMenuOwner = new()

        {
            Visible = false,
            ShowInTaskbar = false,
            WindowState = FormWindowState.Minimized,
            FormBorderStyle = FormBorderStyle.None
        };

        private static readonly NotifyIcon icon = new()
        {
            Text = "DesktopAssistant",
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath),
            Visible = true
        };

        private static readonly ContextMenuStrip menu = new()
        {
            ShowCheckMargin = false,
            ShowImageMargin = false
        };

        private static void Main()
        {
            hotkeyManager = new HotkeyManager();
            reminderManager = new ReminderManager();
            gameAhkManager = new GameAhkManager();
            
            hotkeyManager.Start();
            reminderManager.Start();
            gameAhkManager.Start();
            SetUpTrayIcon();
            Application.Run();
        }

        private static void SetUpTrayIcon()
        {
            menu.Items.Add("关闭显示器", null, (_, _) => DisplayManager.TurnOff());
            menu.Items.Add("-");

            startProxyMenuItem = new ToolStripMenuItem("启动全局抓包代理", null, (_, _) => SetProxyState(true));
            stopProxyMenuItem = new ToolStripMenuItem("关闭全局抓包代理", null, (_, _) => SetProxyState(false));
            
            menu.Items.Add(startProxyMenuItem);
            menu.Items.Add(stopProxyMenuItem);

            // 根据初始状态设置可见性
            bool isEnabled = ProxyManager.IsProxyEnabled();
            startProxyMenuItem.Visible = !isEnabled;
            stopProxyMenuItem.Visible = isEnabled;

            menu.Items.Add("-");
            menu.Items.Add("启动实时翻译", null, (_, _) => StartRealtimeSubtitle());
            menu.Items.Add("-");
            menu.Items.Add("退出", null, (_, _) => { icon.Visible = false; hotkeyManager?.Stop(); gameAhkManager?.Stop(); Application.Exit(); });

            icon.MouseUp += (s, e) =>
            {
                if (e.Button == MouseButtons.Right || e.Button == MouseButtons.Left)
                {
                    // 动态更新一次状态，防止外部修改了代理设置
                    UpdateProxyMenuVisibility();
                    SetForegroundWindow(contextMenuOwner.Handle);
                    menu.Show(Cursor.Position);
                }
            };
        }

        private static void UpdateProxyMenuVisibility()
        {
            bool isEnabled = ProxyManager.IsProxyEnabled();
            if (startProxyMenuItem != null) startProxyMenuItem.Visible = !isEnabled;
            if (stopProxyMenuItem != null) stopProxyMenuItem.Visible = isEnabled;
        }

        private static void SetProxyState(bool enable)
        {
            ProxyManager.SetProxy(enable);
            UpdateProxyMenuVisibility();

            if (enable)
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "http://127.0.0.1:8899/#network",
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    Logger.Log($"无法打开浏览器: {ex.Message}");
                }
            }
        }

        private static void StartRealtimeSubtitle()
        {
            try
            {
                var batPath = @"C:\Code\RealtimeSubtitle\start.bat";
                if (System.IO.File.Exists(batPath))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = batPath,
                        WorkingDirectory = @"C:\Code\RealtimeSubtitle",
                        UseShellExecute = true,
                        WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                    });
                }
                else
                {
                    MessageBox.Show($"找不到文件: {batPath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}


