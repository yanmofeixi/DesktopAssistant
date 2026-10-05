using System.Diagnostics;

namespace DesktopAssistant
{
    /// <summary>
    /// 游戏 AHK 脚本管理器
    /// 使用轮询方式检测游戏进程，自动启动/关闭对应的 AHK 脚本
    /// 支持脚本中途 Reload（PID变更）以及外部启动场景的生命周期同步
    /// </summary>
    public class GameAhkManager
    {
        // AHK 脚本所在文件夹
        private static readonly string ahkFolder = @"C:\Code\AHK\";
        
        // 轮询间隔（毫秒）
        private static readonly int pollIntervalMs = 20 * 1000;
        
        // 游戏进程名 -> AHK脚本名的映射
        private static readonly Dictionary<string, string> gameAhkMapping = new()
        {
            { "GenshinImpact", "Genshin.ahk" },
            { "StarRail", "StarRail.ahk" },
            { "isaac-ng", "Isaac.ahk" }
        };

        // 常见 AutoHotkey 解释器路径（优先使用 64 位 v2 解释器，避免 UX Launcher 产生多余进程）
        private static readonly string[] ahkExeCandidates = new[]
        {
            @"C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe",
            @"C:\Program Files\AutoHotkey\v2\AutoHotkey.exe",
            @"C:\Program Files\AutoHotkey\AutoHotkey64.exe",
            @"C:\Program Files\AutoHotkey\AutoHotkey.exe"
        };

        // 记录哪些游戏的 AHK 已启动
        private readonly Dictionary<string, Process?> runningGames = new();
        
        // 轮询定时器
        private System.Threading.Timer? pollTimer;

        // 当前用户的 Session ID，用于过滤其他用户的进程
        private readonly int currentSessionId = Process.GetCurrentProcess().SessionId;

        /// <summary>
        /// 启动游戏 AHK 管理器
        /// </summary>
        public void Start()
        {
            Logger.Info($"GameAhkManager 启动, 当前用户 Session ID: {currentSessionId}");
            Logger.Info($"AHK 文件夹: {ahkFolder}");
            Logger.Info($"轮询间隔: {pollIntervalMs}ms");
            Logger.Info($"监控的游戏: {string.Join(", ", gameAhkMapping.Keys)}");
            
            // 启用系统 Debug 特权以稳定读取进程命令行
            ProcessCommandLineHelper.EnsureDebugPrivilege();

            // 先同步已有的 AHK 进程（程序重启后恢复状态）
            SyncExistingAhkProcesses();
            
            // 立即检查一次当前运行的游戏
            CheckAndSyncGameStatus();

            // 启动轮询定时器
            pollTimer = new System.Threading.Timer(
                _ => CheckAndSyncGameStatus(),
                null,
                pollIntervalMs,
                pollIntervalMs);
        }

        /// <summary>
        /// 停止游戏 AHK 管理器
        /// </summary>
        public void Stop()
        {
            // 停止定时器
            pollTimer?.Dispose();
            pollTimer = null;

            // 关闭所有监控的游戏 AHK 进程
            foreach (var game in gameAhkMapping.Keys)
            {
                StopAhkScript(game);
            }
            runningGames.Clear();
        }

        /// <summary>
        /// 检查并同步游戏状态：启动新游戏的 AHK，关闭已退出游戏的 AHK
        /// 直接基于系统真实进程状态判定，完美兼容 AHK 内部 Reload 行为
        /// </summary>
        private void CheckAndSyncGameStatus()
        {
            Logger.Debug("开始检查游戏状态...");
            
            foreach (var kvp in gameAhkMapping)
            {
                var game = kvp.Key;
                var ahkScript = kvp.Value;

                bool isGameRunning = IsGameRunning(game);
                bool isAhkRunning = IsAhkScriptRunning(ahkScript);

                Logger.Debug($"游戏 {game}: 游戏运行={isGameRunning}, AHK运行={isAhkRunning}");

                if (isGameRunning && !isAhkRunning)
                {
                    Logger.Info($"检测到游戏 {game} 在运行，但 AHK 没启动 -> 启动 AHK");
                    StartAhkScript(game);
                }
                else if (!isGameRunning && isAhkRunning)
                {
                    Logger.Info($"检测到游戏 {game} 已退出，但 AHK 仍在运行 -> 关闭 AHK");
                    StopAhkScript(game);
                }
                else if (isGameRunning && isAhkRunning)
                {
                    // 游戏与 AHK 均在运行，更新追踪字典以记录运行状态
                    if (!runningGames.ContainsKey(game))
                    {
                        runningGames[game] = null;
                    }
                }
            }
            
            Logger.Debug("游戏状态检查完成");
        }

        /// <summary>
        /// 同步已有的 AHK 进程到 runningGames 字典
        /// </summary>
        private void SyncExistingAhkProcesses()
        {
            Logger.Info("开始同步已有的 AHK 进程...");
            
            foreach (var kvp in gameAhkMapping)
            {
                var gameName = kvp.Key;
                var ahkScript = kvp.Value;
                
                if (IsAhkScriptRunning(ahkScript))
                {
                    Logger.Info($"检测到已有 AHK 进程: {ahkScript}，加入追踪列表");
                    runningGames[gameName] = null;
                }
            }
            
            Logger.Info($"AHK 进程同步完成，当前追踪: {string.Join(", ", runningGames.Keys)}");
        }

        /// <summary>
        /// 检查指定的 AHK 脚本是否在当前用户 Session 中运行
        /// </summary>
        private bool IsAhkScriptRunning(string ahkScript)
        {
            var matching = GetMatchingAhkProcesses(ahkScript);
            bool isRunning = matching.Count > 0;
            foreach (var proc in matching)
            {
                proc.Dispose();
            }
            return isRunning;
        }

        /// <summary>
        /// 获取所有在当前用户 Session 中运行且匹配指定脚本的 AHK 进程
        /// </summary>
        private List<Process> GetMatchingAhkProcesses(string ahkScript)
        {
            var result = new List<Process>();
            var scriptPath = Path.Combine(ahkFolder, ahkScript);
            var ahkProcessNames = new[] { "AutoHotkey", "AutoHotkeyUX", "AutoHotkey64", "AutoHotkey32" };
            
            foreach (var processName in ahkProcessNames)
            {
                Process[] processes;
                try
                {
                    processes = Process.GetProcessesByName(processName);
                }
                catch
                {
                    continue;
                }

                foreach (var process in processes)
                {
                    try
                    {
                        if (process.SessionId != currentSessionId)
                        {
                            process.Dispose();
                            continue;
                        }

                        string? commandLine = ProcessCommandLineHelper.GetCommandLine(process.Id);
                        if (!string.IsNullOrEmpty(commandLine) &&
                            (commandLine.Contains(ahkScript, StringComparison.OrdinalIgnoreCase) ||
                             commandLine.Contains(scriptPath, StringComparison.OrdinalIgnoreCase)))
                        {
                            result.Add(process);
                            continue; // 匹配成功，保留 process 对象供调用方使用
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug($"检查 AHK 进程 PID={process.Id} 命令行失败: {ex.Message}");
                    }

                    process.Dispose();
                }
            }

            return result;
        }

        /// <summary>
        /// 检查指定游戏是否在当前用户 Session 中运行
        /// </summary>
        private bool IsGameRunning(string gameName)
        {
            try
            {
                var processes = Process.GetProcessesByName(gameName);
                foreach (var process in processes)
                {
                    try
                    {
                        if (process.SessionId == currentSessionId)
                        {
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn($"访问进程 PID={process.Id} 时出错: {ex.Message}");
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"查找进程 '{gameName}' 时出错", ex);
            }
            return false;
        }

        /// <summary>
        /// 启动指定游戏的 AHK 脚本
        /// </summary>
        private void StartAhkScript(string gameName)
        {
            if (!gameAhkMapping.TryGetValue(gameName, out var ahkScript))
            {
                Logger.Warn($"StartAhkScript: 游戏 {gameName} 没有对应的 AHK 映射");
                return;
            }

            var ahkPath = Path.Combine(ahkFolder, ahkScript);
            if (!File.Exists(ahkPath))
            {
                Logger.Warn($"StartAhkScript: AHK 脚本不存在: {ahkPath}");
                return;
            }

            // 优先查找 AutoHotkey 解释器可执行文件直接启动，避免 UX Launcher 产生多余进程
            string? ahkExe = ahkExeCandidates.FirstOrDefault(File.Exists);

            try
            {
                Process? process;
                if (ahkExe != null)
                {
                    Logger.Info($"使用解释器启动 AHK 脚本: \"{ahkExe}\" \"{ahkPath}\"");
                    process = Process.Start(new ProcessStartInfo
                    {
                        FileName = ahkExe,
                        Arguments = $"\"{ahkPath}\"",
                        UseShellExecute = false
                    });
                }
                else
                {
                    Logger.Info($"使用系统默认关联启动 AHK 脚本: {ahkPath}");
                    process = Process.Start(new ProcessStartInfo
                    {
                        FileName = ahkPath,
                        UseShellExecute = true
                    });
                }

                runningGames[gameName] = process;
                Logger.Info($"AHK 脚本启动成功, PID={process?.Id}");
            }
            catch (Exception ex)
            {
                Logger.Error($"启动 AHK 脚本失败: {ahkPath}", ex);
            }
        }

        /// <summary>
        /// 停止指定游戏的所有 AHK 进程（包括中途 Reload 产生的进程）
        /// </summary>
        private void StopAhkScript(string gameName)
        {
            if (!gameAhkMapping.TryGetValue(gameName, out var ahkScript))
            {
                Logger.Warn($"StopAhkScript: 游戏 {gameName} 没有对应的 AHK 映射");
                return;
            }

            Logger.Info($"停止游戏 {gameName} 的 AHK 脚本: {ahkScript}");
            var matching = GetMatchingAhkProcesses(ahkScript);

            if (matching.Count == 0)
            {
                Logger.Info($"未发现运行中匹配 {ahkScript} 的 AHK 进程");
            }
            else
            {
                Logger.Info($"共发现 {matching.Count} 个匹配 {ahkScript} 的 AHK 进程，正在终止...");
                foreach (var proc in matching)
                {
                    try
                    {
                        Logger.Info($"正在杀死 AHK 进程: Name={proc.ProcessName}, PID={proc.Id}");
                        proc.Kill();
                        proc.WaitForExit(1000);
                        Logger.Info($"AHK 进程 PID={proc.Id} 已成功终止");
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn($"杀死 AHK 进程 PID={proc.Id} 失败: {ex.Message}");
                    }
                    finally
                    {
                        proc.Dispose();
                    }
                }
            }

            // 清理记录的 Process 对象
            if (runningGames.TryGetValue(gameName, out var savedProc))
            {
                savedProc?.Dispose();
            }
            runningGames.Remove(gameName);
        }
    }
}
