using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PmAi.RdpRemoteLauncher;

internal static class RdpSessionDisconnecter
{
    internal static bool TryDisconnectCurrentSession(out string? errorMessage)
    {
        errorMessage = null;

        if (!TryResolveCurrentSessionId(out var sessionId, out var resolveError))
        {
            errorMessage = resolveError ?? "セッション ID を取得できませんでした";
            return false;
        }

        LauncherLog.Info("RDP 切断対象セッション ID=" + sessionId + " (自プロセス PID=" + Environment.ProcessId + ")");

        if (WTSApi.WTSDisconnectSession(WTSApi.WtsCurrentServerHandle, (int)sessionId, false))
        {
            LauncherLog.Info("WTSDisconnectSession で RDP セッションを切断しました");
            return true;
        }

        errorMessage = FormatWin32Error("WTSDisconnectSession");
        return false;
    }

    private static bool TryResolveCurrentSessionId(out uint sessionId, out string? errorMessage)
    {
        sessionId = 0;
        errorMessage = null;
        var currentPid = Environment.ProcessId;
        if (Kernel32.ProcessIdToSessionId(currentPid, out sessionId))
        {
            return true;
        }

        var win32 = Marshal.GetLastWin32Error();
        errorMessage =
            "ProcessIdToSessionId 失敗 PID="
                + currentPid
                + ": "
                + new Win32Exception(win32).Message
                + " (Win32="
                + win32
                + ")";
        return false;
    }

    private static string FormatWin32Error(string apiName)
    {
        var win32 = Marshal.GetLastWin32Error();
        return apiName + ": " + new Win32Exception(win32).Message + " (Win32=" + win32 + ")";
    }

    private static class Kernel32
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ProcessIdToSessionId(int processId, out uint sessionId);
    }

    private static class WTSApi
    {
        internal static readonly nint WtsCurrentServerHandle = nint.Zero;

        [DllImport("wtsapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool WTSDisconnectSession(nint hServer, int sessionId, bool bWait);
    }
}
