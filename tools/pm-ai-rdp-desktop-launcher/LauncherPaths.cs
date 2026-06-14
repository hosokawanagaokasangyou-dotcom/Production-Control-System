using System.Diagnostics;

namespace PmAi.RdpDesktopLauncher;

internal static class LauncherPaths
{
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
