using System.Diagnostics;
using System.Text;

namespace PmAi.RdpRemoteLauncher;

internal static class Program
{
    private const int ExitOk = 0;
    private const int ExitError = 1;
    private const int ExitMissingIni = 2;

    private static int Main(string[] args)
    {
        try
        {
            var iniPath = ResolveIniPath(args);
            if (string.IsNullOrWhiteSpace(iniPath))
            {
                LauncherLog.Error("ini パスが未指定です。--ini または PM_AI_RDP_LAUNCHER_INI、または exe 同階層の RAP設定.ini を指定してください。");
                return ExitMissingIni;
            }

            if (!File.Exists(iniPath))
            {
                LauncherLog.Error("ini が見つかりません: " + iniPath);
                return ExitMissingIni;
            }

            var ini = LauncherIni.Load(iniPath);
            var commandLine = ini.ResolveSelectedCommand();
            if (string.IsNullOrWhiteSpace(commandLine))
            {
                LauncherLog.Error("起動プログラム番号 " + ini.SelectedSlot + " に対応するスロットが ini にありません: " + iniPath);
                return ExitError;
            }

            ParsedCommand parsed;
            try
            {
                parsed = CommandLineParser.Parse(commandLine);
            }
            catch (FormatException ex)
            {
                LauncherLog.Error(ex.Message);
                return ExitError;
            }

            if (ProcessRunningChecker.IsAlreadyRunning(parsed))
            {
                LauncherLog.Info("既に起動済みのためスキップ: " + commandLine);
                return ExitOk;
            }

            var workingDirectory = Path.GetDirectoryName(parsed.Executable);
            if (string.IsNullOrWhiteSpace(workingDirectory))
            {
                workingDirectory = Environment.CurrentDirectory;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = parsed.Executable,
                Arguments = parsed.Arguments,
                WorkingDirectory = workingDirectory,
                UseShellExecute = true,
            };

            Process.Start(startInfo);
            LauncherLog.Info("起動しました: " + commandLine);
            return ExitOk;
        }
        catch (Exception ex)
        {
            LauncherLog.Error(ex.Message);
            return ExitError;
        }
    }

    private static string? ResolveIniPath(string[] args)
    {
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg == "--ini" && i + 1 < args.Length)
            {
                return args[i + 1];
            }

            if (arg.StartsWith("--ini=", StringComparison.OrdinalIgnoreCase))
            {
                return arg["--ini=".Length..];
            }
        }

        var fromEnv = Environment.GetEnvironmentVariable(LauncherIni.IniPathEnvVar);
        if (!string.IsNullOrWhiteSpace(fromEnv))
        {
            return fromEnv.Trim();
        }

        var exeDir = AppContext.BaseDirectory;
        if (string.IsNullOrWhiteSpace(exeDir))
        {
            exeDir = Environment.CurrentDirectory;
        }

        return Path.Combine(exeDir, LauncherIni.DefaultIniFileName);
    }
}
