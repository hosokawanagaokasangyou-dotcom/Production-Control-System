using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PmAi.RdpRemoteLauncher;

internal static class RdpSessionDisconnecter
{
    private const int WtsCurrentSession = -1;

    internal static bool TryDisconnectCurrentSession(out string? errorMessage)
    {
        errorMessage = null;
        if (!WTSApi.WTSDisconnectSession(WTSApi.WtsCurrentServerHandle, WtsCurrentSession, false))
        {
            var win32 = Marshal.GetLastWin32Error();
            errorMessage = new Win32Exception(win32).Message + " (Win32=" + win32 + ")";
            return false;
        }

        return true;
    }

    private static class WTSApi
    {
        internal static readonly nint WtsCurrentServerHandle = nint.Zero;

        [DllImport("wtsapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool WTSDisconnectSession(nint hServer, int sessionId, bool bWait);
    }
}
