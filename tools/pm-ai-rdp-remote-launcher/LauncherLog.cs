using System.Text;

namespace PmAi.RdpRemoteLauncher;

internal static class LauncherLog
{
    private static readonly object Gate = new();
    private static string? mirrorDirectory;

    internal static void SetMirrorDirectory(string? directory)
    {
        mirrorDirectory = string.IsNullOrWhiteSpace(directory) ? null : directory.Trim();
    }

    internal static void Info(string message)
    {
        Write("INFO", message);
    }

    internal static void Warn(string message)
    {
        Write("WARN", message);
    }

    internal static void Error(string message)
    {
        Write("ERROR", message);
    }

    private static void Write(string level, string message)
    {
        var line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " [" + level + "] " + message;
        lock (Gate)
        {
            Console.Error.WriteLine(line);
            AppendToDailyLog(Path.Combine(Path.GetTempPath(), "PM-AI-RDP-Launcher"), line);
            if (!string.IsNullOrWhiteSpace(mirrorDirectory))
            {
                AppendToDailyLog(mirrorDirectory, line);
            }
        }
    }

    private static void AppendToDailyLog(string directory, string line)
    {
        try
        {
            Directory.CreateDirectory(directory);
            var file = Path.Combine(directory, "launcher-" + DateTime.Now.ToString("yyyyMMdd") + ".log");
            File.AppendAllText(file, line + Environment.NewLine, Encoding.UTF8);
        }
        catch
        {
            // ログ書込失敗は無視
        }
    }
}
