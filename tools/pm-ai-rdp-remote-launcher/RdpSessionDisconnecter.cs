using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PmAi.RdpRemoteLauncher;

internal static class RdpSessionDisconnecter
{
    internal static bool TryDisconnectCurrentSession(out string? errorMessage)
    {
        errorMessage = null;

        if (!TryResolveCurrentSessionId(out var sessionId, out var resolveError))
        {
            LauncherLog.Warn(resolveError ?? "セッション ID を取得できませんでした");
            return TryShutdownLogoffFallback(out errorMessage);
        }

        LauncherLog.Info("RDP 切断対象セッション ID=" + sessionId + " (自プロセス PID=" + Environment.ProcessId + ")");

        if (WTSApi.WTSDisconnectSession(WTSApi.WtsCurrentServerHandle, (int)sessionId, false))
        {
            return true;
        }

        var disconnectError = FormatWin32Error("WTSDisconnectSession");
        LauncherLog.Warn("WTSDisconnectSession 失敗: " + disconnectError + " — WTSLogoffSession を試行");

        if (WTSApi.WTSLogoffSession(WTSApi.WtsCurrentServerHandle, (int)sessionId, false))
        {
            LauncherLog.Info("WTSLogoffSession でセッションを終了しました");
            return true;
        }

        var logoffError = FormatWin32Error("WTSLogoffSession");
        LauncherLog.Warn("WTSLogoffSession 失敗: " + logoffError + " — shutdown /l /f を試行");

        return TryShutdownLogoffFallback(out errorMessage);
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

    private static bool TryShutdownLogoffFallback(out string? errorMessage)
    {
        errorMessage = null;
        try
        {
            var systemRoot = Environment.GetEnvironmentVariable("SystemRoot");
            var shutdownPath = string.IsNullOrWhiteSpace(systemRoot)
                ? "shutdown.exe"
                : Path.Combine(systemRoot, "System32", "shutdown.exe");

            LauncherLog.Info("shutdown.exe /l /f を起動してログオフを試行: " + shutdownPath);
            using var shutdown = Process.Start(
                new ProcessStartInfo
                {
                    FileName = shutdownPath,
                    Arguments = "/l /f",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                });
            if (shutdown == null)
            {
                errorMessage = "shutdown.exe の起動に失敗しました";
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            errorMessage = "shutdown /l /f 失敗: " + ex.Message;
            return false;
        }
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

        [DllImport("wtsapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool WTSLogoffSession(nint hServer, int sessionId, bool bWait);
    }
}
