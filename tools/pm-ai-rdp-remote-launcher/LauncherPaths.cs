using System.Diagnostics;

namespace PmAi.RdpRemoteLauncher;

internal static class LauncherPaths
{
    /// <summary>
    /// 単一ファイル publish 時は <see cref="AppContext.BaseDirectory"/> が展開先 TEMP になるため、
    /// 実 exe（UNC 上の PmAiRdpRemoteLauncher.exe）のパスを使う。
    /// </summary>
    internal static string? ResolveExecutablePath()
    {
        var processPath = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(processPath))
        {
            return processPath.Trim();
        }

        try
        {
            return Process.GetCurrentProcess().MainModule?.FileName;
        }
        catch
        {
            return null;
        }
    }

    internal static string? ResolveExecutableDirectory()
    {
        var exePath = ResolveExecutablePath();
        if (string.IsNullOrWhiteSpace(exePath))
        {
            return null;
        }

        var dir = Path.GetDirectoryName(exePath);
        return string.IsNullOrWhiteSpace(dir) ? null : dir;
    }
}
