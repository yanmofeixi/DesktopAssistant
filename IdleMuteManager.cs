using System.Runtime.InteropServices;

namespace DesktopAssistant
{
    public class IdleMuteManager
    {
        private System.Threading.Timer? pollTimer;
        private bool isMutedByUs = false;
        private float savedVolume = -1f; // 静音前保存的原始音量
        private bool isEnabled = true; // 默认开启

        // 空闲超时时间，30分钟 = 1,800,000 ms
        private const uint IdleTimeoutMs = 30 * 60 * 1000;
        // 轮询时间间隔：30 秒
        private const int PollIntervalMs = 30 * 1000;

        public bool Enabled
        {
            get => isEnabled;
            set
            {
                if (isEnabled != value)
                {
                    isEnabled = value;
                    Logger.Info($"空闲静音功能已{(isEnabled ? "启用" : "禁用")}");
                    if (!isEnabled && isMutedByUs)
                    {
                        // 如果在禁用时当前处于我们设置的静音状态，则恢复音量
                        if (savedVolume > 0f)
                        {
                            SetSystemVolume(savedVolume);
                        }
                        isMutedByUs = false;
                        savedVolume = -1f;
                    }
                }
            }
        }

        public void Start()
        {
            Logger.Info("IdleMuteManager 启动");
            pollTimer = new System.Threading.Timer(
                _ => CheckIdleStatus(),
                null,
                PollIntervalMs,
                PollIntervalMs);
        }

        public void Stop()
        {
            pollTimer?.Dispose();
            pollTimer = null;
            if (isMutedByUs)
            {
                if (savedVolume > 0f)
                {
                    SetSystemVolume(savedVolume);
                }
                isMutedByUs = false;
                savedVolume = -1f;
            }
            Logger.Info("IdleMuteManager 停止");
        }

        private void CheckIdleStatus()
        {
            if (!isEnabled) return;

            try
            {
                uint idleMs = GetIdleTimeMs();
                float currentVolume = GetSystemVolume();

                Logger.Debug($"当前系统空闲时间: {idleMs / 1000} 秒, 系统音量: {currentVolume:F2}, isMutedByUs: {isMutedByUs}");

                // 如果我们记录为"由我们静音"，但用户已手动调高了音量，说明用户手动恢复了
                if (isMutedByUs && currentVolume > 0.01f)
                {
                    Logger.Info("检测到用户手动恢复了音量，重置自动静音状态追踪");
                    isMutedByUs = false;
                    savedVolume = -1f;
                }

                if (idleMs >= IdleTimeoutMs)
                {
                    if (!isMutedByUs && currentVolume > 0.01f)
                    {
                        Logger.Info($"检测到系统空闲时间超过 30 分钟，保存当前音量 {currentVolume:F2} 并设为 0");
                        savedVolume = currentVolume;
                        SetSystemVolume(0f);
                        isMutedByUs = true;
                    }
                }
                else
                {
                    // 用户恢复操作
                    if (isMutedByUs)
                    {
                        Logger.Info($"检测到用户活动恢复，还原音量为 {savedVolume:F2}");
                        if (savedVolume > 0f)
                        {
                            SetSystemVolume(savedVolume);
                        }
                        isMutedByUs = false;
                        savedVolume = -1f;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("CheckIdleStatus 出错", ex);
            }
        }

        private uint GetIdleTimeMs()
        {
            LASTINPUTINFO lii = new LASTINPUTINFO();
            lii.cbSize = (uint)Marshal.SizeOf(lii);
            if (GetLastInputInfo(ref lii))
            {
                uint currentTick = (uint)Environment.TickCount;
                return currentTick - lii.dwTime;
            }
            return 0;
        }

        /// <summary>
        /// 获取系统主音量 (0.0 ~ 1.0)
        /// </summary>
        private float GetSystemVolume()
        {
            IAudioEndpointVolume? volume = GetAudioEndpointVolume();
            if (volume != null)
            {
                try
                {
                    int hr = volume.GetMasterVolumeLevelScalar(out float level);
                    return hr == 0 ? level : -1f;
                }
                finally
                {
                    Marshal.ReleaseComObject(volume);
                }
            }
            return -1f;
        }

        /// <summary>
        /// 设置系统主音量 (0.0 ~ 1.0)
        /// </summary>
        private void SetSystemVolume(float level)
        {
            IAudioEndpointVolume? volume = GetAudioEndpointVolume();
            if (volume != null)
            {
                try
                {
                    Guid eventContext = Guid.NewGuid();
                    int hr = volume.SetMasterVolumeLevelScalar(level, ref eventContext);
                    if (hr == 0)
                    {
                        Logger.Info($"系统主音量已设置为: {level:F2}");
                    }
                    else
                    {
                        Logger.Warn($"设置系统音量失败，HRESULT: {hr}");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("设置系统音量失败", ex);
                }
                finally
                {
                    Marshal.ReleaseComObject(volume);
                }
            }
        }

        private IAudioEndpointVolume? GetAudioEndpointVolume()
        {
            try
            {
                var enumeratorType = Type.GetTypeFromCLSID(new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"));
                if (enumeratorType == null) return null;

                var enumerator = Activator.CreateInstance(enumeratorType) as IMMDeviceEnumerator;
                if (enumerator == null) return null;

                try
                {
                    // eRender = 0, eConsole = 0
                    int hr = enumerator.GetDefaultAudioEndpoint(0, 0, out IMMDevice device);
                    if (hr != 0 || device == null) return null;

                    try
                    {
                        Guid iid = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
                        hr = device.Activate(ref iid, 1, IntPtr.Zero, out object volumeObj);
                        if (hr != 0 || volumeObj == null) return null;
                        
                        return volumeObj as IAudioEndpointVolume;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(device);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(enumerator);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("获取 IAudioEndpointVolume 接口失败", ex);
                return null;
            }
        }

        #region Win32 / COM Imports

        [DllImport("user32.dll")]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO
        {
            public uint cbSize;
            public uint dwTime;
        }

        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDeviceEnumerator
        {
            [PreserveSig]
            int EnumAudioEndpoints(int dataFlow, int stateMask, out IntPtr ppDevices);
            [PreserveSig]
            int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppDevice);
            [PreserveSig]
            int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);
            [PreserveSig]
            int RegisterEndpointNotificationCallback(IntPtr pClient);
            [PreserveSig]
            int UnregisterEndpointNotificationCallback(IntPtr pClient);
        }

        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDevice
        {
            [PreserveSig]
            int Activate(ref Guid iid, int dwClsContext, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
            [PreserveSig]
            int OpenPropertyStore(int stgmAccess, out IntPtr ppProperties);
            [PreserveSig]
            int GetId(out IntPtr ppstrId);
            [PreserveSig]
            int GetState(out int pdwState);
        }

        [Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioEndpointVolume
        {
            [PreserveSig]
            int RegisterControlChangeNotify(IntPtr pNotify);
            [PreserveSig]
            int UnregisterControlChangeNotify(IntPtr pNotify);
            [PreserveSig]
            int GetChannelCount(out uint pnChannelCount);
            [PreserveSig]
            int SetMasterVolumeLevel(float fLevelDB, ref Guid pguidEventContext);
            [PreserveSig]
            int SetMasterVolumeLevelScalar(float fLevel, ref Guid pguidEventContext);
            [PreserveSig]
            int GetMasterVolumeLevel(out float pfLevelDB);
            [PreserveSig]
            int GetMasterVolumeLevelScalar(out float pfLevel);
            [PreserveSig]
            int SetChannelVolumeLevel(uint nChannel, float fLevelDB, ref Guid pguidEventContext);
            [PreserveSig]
            int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, ref Guid pguidEventContext);
            [PreserveSig]
            int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);
            [PreserveSig]
            int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);
            [PreserveSig]
            int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, ref Guid pguidEventContext);
            [PreserveSig]
            int GetMute([MarshalAs(UnmanagedType.Bool)] out bool pbMute);
            [PreserveSig]
            int GetVolumeStepInfo(out uint pnStep, out uint pnSteps);
            [PreserveSig]
            int VolumeStepUp(ref Guid pguidEventContext);
            [PreserveSig]
            int VolumeStepDown(ref Guid pguidEventContext);
            [PreserveSig]
            int QueryHardwareSupport(out uint pdwHardwareSupportMask);
            [PreserveSig]
            int GetVolumeRange(out float pfVolumeMinDB, out float pfVolumeMaxDB, out float pfVolumeIncrementDB);
        }

        #endregion
    }
}
