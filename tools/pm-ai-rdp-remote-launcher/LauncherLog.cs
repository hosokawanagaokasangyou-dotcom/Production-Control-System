using System.Text;

namespace PmAi.RdpRemoteLauncher;

internal static class LauncherLog
{
    private static readonly object Gate = new();

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
            try
            {
                var dir = Path.Combine(Path.GetTempPath(), "PM-AI-RDP-Launcher");
                Directory.CreateDirectory(dir);
                var file = Path.Combine(dir, "launcher-" + DateTime.Now.ToString("yyyyMMdd") + ".log");
                File.AppendAllText(file, line + Environment.NewLine, Encoding.UTF8);
            }
            catch
            {
                // ログ書込失敗は無視
            }
        }
    }
}
