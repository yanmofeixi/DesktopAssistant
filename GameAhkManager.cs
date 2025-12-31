using System.Diagnostics;
using System.Management;

namespace DesktopAssistant
{
    /// <summary>
    /// 游戏 AHK 脚本管理器
    /// 使用轮询方式检测游戏进程，自动启动/关闭对应的 AHK 脚本
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
            { "StarRail", "StarRail.ahk" }
        };

        // 记录哪些游戏的AHK已启动，存储游戏名到对应AHK进程的映射
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

            // 关闭所有已启动的AHK进程
            foreach (var game in runningGames.Keys.ToList())
            {
                StopAhkScript(game);
            }
            runningGames.Clear();
        }

        /// <summary>
        /// 检查并同步游戏状态：启动新游戏的AHK，关闭已退出游戏的AHK
        /// </summary>
        private void CheckAndSyncGameStatus()
        {
            Logger.Debug($"开始检查游戏状态...");
            
            foreach (var game in gameAhkMapping.Keys)
            {
                bool isGameRunning = IsGameRunning(game);
                bool isAhkRunning = runningGames.ContainsKey(game);

                Logger.Debug($"游戏 {game}: 游戏运行={isGameRunning}, AHK运行={isAhkRunning}");

                if (isGameRunning && !isAhkRunning)
                {
                    Logger.Info($"检测到游戏 {game} 在运行，但 AHK 没启动 -> 启动 AHK");
                    StartAhkScript(game);
                }
                else if (!isGameRunning && isAhkRunning)
                {
                    Logger.Info($"检测到游戏 {game} 已退出，但 AHK 在运行 -> 关闭 AHK");
                    StopAhkScript(game);
                }
            }
            
            Logger.Debug($"游戏状态检查完成");
        }

        /// <summary>
        /// 同步已有的 AHK 进程到 runningGames 字典
        /// 用于程序重启后恢复对已运行 AHK 的追踪
        /// </summary>
        private void SyncExistingAhkProcesses()
        {
            Logger.Info("开始同步已有的 AHK 进程...");
            
            foreach (var kvp in gameAhkMapping)
            {
                var gameName = kvp.Key;
                var ahkScript = kvp.Value;
                
                // 如果已经在追踪中，跳过
                if (runningGames.ContainsKey(gameName)) continue;
                
                // 检查是否有对应的 AHK 进程在运行
                if (IsAhkScriptRunning(ahkScript))
                {
                    Logger.Info($"检测到已有 AHK 进程: {ahkScript}，加入追踪列表");
                    // 用 null 标记，表示这是外部启动的进程，需要通过命令行匹配关闭
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
            var scriptPath = Path.Combine(ahkFolder, ahkScript);
            var ahkProcessNames = new[] { "AutoHotkey", "AutoHotkeyUX", "AutoHotkey64", "AutoHotkey32" };
            
            foreach (var processName in ahkProcessNames)
            {
                try
                {
                    var processes = Process.GetProcessesByName(processName);
                    foreach (var process in processes)
                    {
                        try
                        {
                            if (process.SessionId != currentSessionId) continue;
                            
                            using var searcher = new ManagementObjectSearcher(
                                $"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {process.Id}");
                            
                            foreach (ManagementObject obj in searcher.Get())
                            {
                                var commandLine = obj["CommandLine"]?.ToString() ?? "";
                                if (commandLine.Contains(ahkScript, StringComparison.OrdinalIgnoreCase) ||
                                    commandLine.Contains(scriptPath, StringComparison.OrdinalIgnoreCase))
                                {
                                    return true;
                                }
                            }
                        }
                        catch { }
                        finally
                        {
                            process.Dispose();
                        }
                    }
                }
                catch { }
            }
            return false;
        }

        /// <summary>
        /// 检查指定游戏是否在当前用户 Session 中运行
        /// </summary>
        private bool IsGameRunning(string gameName)
        {
            try
            {
                var processes = Process.GetProcessesByName(gameName);
                Logger.Debug($"查找进程 '{gameName}': 找到 {processes.Length} 个进程");
                
                foreach (var process in processes)
                {
                    try
                    {
                        var processSessionId = process.SessionId;
                        Logger.Debug($"  进程 PID={process.Id}, SessionId={processSessionId}, 当前SessionId={currentSessionId}, 匹配={processSessionId == currentSessionId}");
                        
                        if (processSessionId == currentSessionId)
                        {
                            Logger.Debug($"  -> 找到匹配的进程，返回 true");
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn($"  访问进程 PID={process.Id} 时出错: {ex.Message}");
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
            
            // 如果已经在运行，不重复启动
            if (runningGames.ContainsKey(gameName))
            {
                Logger.Debug($"StartAhkScript: 游戏 {gameName} 的 AHK 已在运行列表中");
                return;
            }

            var ahkPath = Path.Combine(ahkFolder, ahkScript);
            if (!File.Exists(ahkPath))
            {
                Logger.Warn($"StartAhkScript: AHK 脚本不存在: {ahkPath}");
                return;
            }

            try
            {
                Logger.Info($"启动 AHK 脚本: {ahkPath}");
                var process = Process.Start(new ProcessStartInfo
                {
                    FileName = ahkPath,
                    UseShellExecute = true
                });
                runningGames[gameName] = process;
                Logger.Info($"AHK 脚本启动成功, PID={process?.Id}");
            }
            catch (Exception ex)
            {
                Logger.Error($"启动 AHK 脚本失败: {ahkPath}", ex);
            }
        }

        /// <summary>
        /// 停止指定游戏的 AHK 脚本
        /// </summary>
        private void StopAhkScript(string gameName)
        {
            if (!gameAhkMapping.TryGetValue(gameName, out var ahkScript))
            {
                Logger.Warn($"StopAhkScript: 游戏 {gameName} 没有对应的 AHK 映射");
                return;
            }
            if (!runningGames.TryGetValue(gameName, out var ahkProcess))
            {
                Logger.Debug($"StopAhkScript: 游戏 {gameName} 不在运行列表中");
                return;
            }

            // AHK v2 使用 launcher 机制，Process.Start 返回的是 launcher 进程
            // launcher 启动脚本后就退出，实际脚本运行在另一个进程中
            // 因此总是使用命令行匹配方式关闭
            Logger.Info($"停止 AHK 脚本: {ahkScript}");
            KillAhkByCommandLine(ahkScript);
            
            // 清理保存的 Process 对象
            ahkProcess?.Dispose();
            runningGames.Remove(gameName);
        }

        /// <summary>
        /// 通过命令行参数匹配来关闭 AHK 进程（备用方案）
        /// 用于程序重启后无法通过 Process 对象关闭的情况
        /// </summary>
        private void KillAhkByCommandLine(string ahkScript)
        {
            var scriptPath = Path.Combine(ahkFolder, ahkScript);
            
            // 查找所有 AutoHotkey 进程变体
            var ahkProcessNames = new[] { "AutoHotkey", "AutoHotkeyUX", "AutoHotkey64", "AutoHotkey32" };
            
            foreach (var processName in ahkProcessNames)
            {
                var processes = Process.GetProcessesByName(processName);
                foreach (var process in processes)
                {
                    try
                    {
                        // 确保只关闭当前用户 Session 启动的 AHK 进程
                        if (process.SessionId != currentSessionId) continue;

                        // 通过 WMI 查询命令行参数匹配脚本
                        using var searcher = new ManagementObjectSearcher(
                            $"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {process.Id}");
                        
                        foreach (ManagementObject obj in searcher.Get())
                        {
                            var commandLine = obj["CommandLine"]?.ToString() ?? "";
                            if (commandLine.Contains(ahkScript, StringComparison.OrdinalIgnoreCase) ||
                                commandLine.Contains(scriptPath, StringComparison.OrdinalIgnoreCase))
                            {
                                process.Kill();
                            }
                        }
                    }
                    catch { }
                    finally 
                    { 
                        process.Dispose(); 
                    }
                }
            }
        }
    }
}

