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

    /// <summary>ini や UI 由来で残った外側の引用符を除去する。</summary>
    internal static string NormalizeExecutablePath(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return string.Empty;
        }

        var path = raw.Trim();
        while (path.Length >= 2 && path[0] == '"' && path[^1] == '"')
        {
            path = path[1..^1].Trim();
        }

        return path.Replace("\"\"", "\"");
    }
}
