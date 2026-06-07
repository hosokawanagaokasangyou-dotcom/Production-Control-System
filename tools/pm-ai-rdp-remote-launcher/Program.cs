using System.Diagnostics;

namespace PmAi.RdpRemoteLauncher;

internal static class Program
{
    private const int ExitOk = 0;
    private const int ExitError = 1;
    private const int ExitMissingIni = 2;

    private static int Main(string[] args)
    {
        var exeDir = LauncherPaths.ResolveExecutableDirectory();
        LauncherLog.SetMirrorDirectory(exeDir);
        LauncherLog.Info(
            "PmAiRdpRemoteLauncher 開始"
                + " ProcessPath="
                + (LauncherPaths.ResolveExecutablePath() ?? "(不明)")
                + " BaseDirectory="
                + AppContext.BaseDirectory);

        try
        {
            var iniPath = ResolveIniPath(args, exeDir);
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

            LauncherLog.Info("ini パス: " + iniPath);

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

            var disconnectOnChildExit = ini.ResolveDisconnectOnChildExit();
            LauncherLog.Info("終了時RDP切断: " + (disconnectOnChildExit ? "有効" : "無効"));

            Process? child = null;
            var existingProcessId = ProcessRunningChecker.TryFindRunningProcessId(parsed);
            if (existingProcessId.HasValue)
            {
                LauncherLog.Info("既に起動済みのため監視のみ: PID=" + existingProcessId.Value + " | " + commandLine);
                child = TryOpenProcess(existingProcessId.Value);
                if (child == null)
                {
                    LauncherLog.Error("既存プロセスを開けませんでした PID=" + existingProcessId.Value);
                    return ExitError;
                }
            }
            else
            {
                var workingDirectory = Path.GetDirectoryName(parsed.Executable);
                if (string.IsNullOrWhiteSpace(workingDirectory))
                {
                    workingDirectory = Environment.CurrentDirectory;
                }

                var argumentTokens = WindowsArgumentFormatter.TokenizeForProcess(parsed.Arguments);
                var startInfo = new ProcessStartInfo
                {
                    FileName = parsed.Executable,
                    UseShellExecute = false,
                    WorkingDirectory = workingDirectory,
                };
                foreach (var token in argumentTokens)
                {
                    startInfo.ArgumentList.Add(token);
                }

                LauncherLog.Info(
                    "起動コマンド: exe="
                        + parsed.Executable
                        + " | args="
                        + (argumentTokens.Count == 0
                            ? "(なし)"
                            : "[" + string.Join("] [", argumentTokens) + "]")
                        + " | cwd="
                        + workingDirectory);

                child = Process.Start(startInfo);
                if (child == null)
                {
                    LauncherLog.Error("子プロセスの起動に失敗しました: " + commandLine);
                    return ExitError;
                }

                LauncherLog.Info("起動しました PID=" + child.Id + " (ini 行): " + commandLine);
            }

            using (child)
            {
                child.WaitForExit();
                LauncherLog.Info("子プロセス終了 PID=" + child.Id + " ExitCode=" + child.ExitCode);
            }

            if (disconnectOnChildExit)
            {
                if (RdpSessionDisconnecter.TryDisconnectCurrentSession(out var disconnectError))
                {
                    LauncherLog.Info("RDP セッションを切断しました");
                }
                else
                {
                    LauncherLog.Error("RDP 切断失敗: " + disconnectError);
                }
            }

            return ExitOk;
        }
        catch (Exception ex)
        {
            LauncherLog.Error(ex.Message);
            return ExitError;
        }
    }

    private static Process? TryOpenProcess(int processId)
    {
        try
        {
            var process = Process.GetProcessById(processId);
            if (process.HasExited)
            {
                process.Dispose();
                return null;
            }

            return process;
        }
        catch (ArgumentException)
        {
            return null;
        }
    }

    private static string? ResolveIniPath(string[] args, string? exeDir)
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

        var directory = exeDir;
        if (string.IsNullOrWhiteSpace(directory))
        {
            directory = AppContext.BaseDirectory;
        }

        if (string.IsNullOrWhiteSpace(directory))
        {
            directory = Environment.CurrentDirectory;
        }

        return Path.Combine(directory, LauncherIni.DefaultIniFileName);
    }
}
