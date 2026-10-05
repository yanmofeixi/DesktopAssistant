using System;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;

namespace DesktopAssistant
{
    /// <summary>
    /// 跨进程命令行读取工具
    /// 优先使用 Native API 直接读取目标进程的 PEB，失败时回退到带特权的 WMI 查询
    /// </summary>
    public static class ProcessCommandLineHelper
    {
        private static bool privilegeEnabled = false;
        private static readonly object privilegeLock = new();

        #region P/Invoke Definitions

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool OpenProcessToken(IntPtr ProcessHandle, uint DesiredAccess, out IntPtr TokenHandle);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool LookupPrivilegeValue(string? lpSystemName, string lpName, out LUID lpLuid);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool AdjustTokenPrivileges(IntPtr TokenHandle, bool DisableAllPrivileges, ref TOKEN_PRIVILEGES NewState, int BufferLengthInBytes, IntPtr PreviousState, IntPtr ReturnLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(uint processAccess, bool bInheritHandle, int processId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        [DllImport("ntdll.dll")]
        private static extern int NtQueryInformationProcess(IntPtr processHandle, int processInformationClass, ref PROCESS_BASIC_INFORMATION processInformation, int processInformationLength, out int returnLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, [Out] byte[] lpBuffer, int dwSize, out IntPtr lpNumberOfBytesRead);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, out IntPtr lpBuffer, int dwSize, out IntPtr lpNumberOfBytesRead);

        [StructLayout(LayoutKind.Sequential)]
        private struct LUID
        {
            public uint LowPart;
            public int HighPart;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LUID_AND_ATTRIBUTES
        {
            public LUID Luid;
            public uint Attributes;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct TOKEN_PRIVILEGES
        {
            public uint PrivilegeCount;
            public LUID_AND_ATTRIBUTES Privileges;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct PROCESS_BASIC_INFORMATION
        {
            public IntPtr ExitStatus;
            public IntPtr PebBaseAddress;
            public IntPtr AffinityMask;
            public IntPtr BasePriority;
            public IntPtr UniqueProcessId;
            public IntPtr InheritedFromUniqueProcessId;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct UNICODE_STRING
        {
            public ushort Length;
            public ushort MaximumLength;
            public IntPtr Buffer;
        }

        private const uint TOKEN_ADJUST_PRIVILEGES = 0x0020;
        private const uint TOKEN_QUERY = 0x0008;
        private const uint SE_PRIVILEGE_ENABLED = 0x00000002;
        private const string SE_DEBUG_NAME = "SeDebugPrivilege";

        private const uint PROCESS_QUERY_INFORMATION = 0x0400;
        private const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;
        private const uint PROCESS_VM_READ = 0x0010;

        #endregion

        /// <summary>
        /// 启用当前进程的 SeDebugPrivilege 特权
        /// </summary>
        public static void EnsureDebugPrivilege()
        {
            if (privilegeEnabled) return;

            lock (privilegeLock)
            {
                if (privilegeEnabled) return;

                try
                {
                    IntPtr hToken;
                    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, out hToken))
                    {
                        return;
                    }

                    try
                    {
                        LUID luid;
                        if (!LookupPrivilegeValue(null, SE_DEBUG_NAME, out luid))
                        {
                            return;
                        }

                        TOKEN_PRIVILEGES tp = new()
                        {
                            PrivilegeCount = 1,
                            Privileges = new LUID_AND_ATTRIBUTES
                            {
                                Luid = luid,
                                Attributes = SE_PRIVILEGE_ENABLED
                            }
                        };

                        if (AdjustTokenPrivileges(hToken, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero))
                        {
                            privilegeEnabled = true;
                            Logger.Debug("SeDebugPrivilege 启用成功");
                        }
                    }
                    finally
                    {
                        CloseHandle(hToken);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Debug($"启用 SeDebugPrivilege 时出错: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 获取指定进程的命令行参数
        /// </summary>
        public static string? GetCommandLine(int processId)
        {
            EnsureDebugPrivilege();

            // 优先通过 PEB 读取
            string? cmdLine = GetCommandLineFromPeb(processId);
            if (!string.IsNullOrEmpty(cmdLine))
            {
                return cmdLine;
            }

            // 回退到 WMI 查询
            return GetCommandLineFromWmi(processId);
        }

        /// <summary>
        /// 通过 Native API 读取目标 64 位进程的 PEB 中的命令行
        /// </summary>
        private static string? GetCommandLineFromPeb(int processId)
        {
            IntPtr hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, processId);
            if (hProcess == IntPtr.Zero)
            {
                hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_VM_READ, false, processId);
            }

            if (hProcess == IntPtr.Zero)
            {
                return null;
            }

            try
            {
                PROCESS_BASIC_INFORMATION pbi = new();
                int returnLength;
                int status = NtQueryInformationProcess(hProcess, 0, ref pbi, Marshal.SizeOf(pbi), out returnLength);
                if (status != 0 || pbi.PebBaseAddress == IntPtr.Zero)
                {
                    return null;
                }

                // 64 位系统下：PEB.ProcessParameters 偏移为 0x20
                IntPtr processParametersAddr = new(pbi.PebBaseAddress.ToInt64() + 0x20);
                if (!ReadProcessMemory(hProcess, processParametersAddr, out IntPtr processParametersPtr, IntPtr.Size, out _))
                {
                    return null;
                }

                if (processParametersPtr == IntPtr.Zero)
                {
                    return null;
                }

                // 64 位系统下：RTL_USER_PROCESS_PARAMETERS.CommandLine 偏移为 0x70
                IntPtr cmdLineUnicodeStringAddr = new(processParametersPtr.ToInt64() + 0x70);
                byte[] unicodeStringBytes = new byte[Marshal.SizeOf<UNICODE_STRING>()];

                if (!ReadProcessMemory(hProcess, cmdLineUnicodeStringAddr, unicodeStringBytes, unicodeStringBytes.Length, out _))
                {
                    return null;
                }

                GCHandle handle = GCHandle.Alloc(unicodeStringBytes, GCHandleType.Pinned);
                UNICODE_STRING cmdLineString;
                try
                {
                    cmdLineString = Marshal.PtrToStructure<UNICODE_STRING>(handle.AddrOfPinnedObject());
                }
                finally
                {
                    handle.Free();
                }

                if (cmdLineString.Length == 0 || cmdLineString.Buffer == IntPtr.Zero)
                {
                    return "";
                }

                // 读取命令行宽字符内容
                byte[] cmdLineBytes = new byte[cmdLineString.Length];
                if (!ReadProcessMemory(hProcess, cmdLineString.Buffer, cmdLineBytes, cmdLineBytes.Length, out _))
                {
                    return null;
                }

                return Encoding.Unicode.GetString(cmdLineBytes);
            }
            catch
            {
                return null;
            }
            finally
            {
                CloseHandle(hProcess);
            }
        }

        /// <summary>
        /// 通过带特权的 WMI 查询命令行
        /// </summary>
        private static string? GetCommandLineFromWmi(int processId)
        {
            try
            {
                ConnectionOptions options = new()
                {
                    EnablePrivileges = true,
                    Impersonation = ImpersonationLevel.Impersonate
                };

                ManagementScope scope = new(@"\\.\root\cimv2", options);
                scope.Connect();

                using ManagementObjectSearcher searcher = new(scope,
                    new ObjectQuery($"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {processId}"));

                foreach (ManagementObject obj in searcher.Get())
                {
                    return obj["CommandLine"]?.ToString();
                }
            }
            catch
            {
            }

            return null;
        }
    }
}
