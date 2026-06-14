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
                + " version="
                + ResolveLauncherVersionLabel()
                + " ProcessPath="
                + (LauncherPaths.ResolveExecutablePath() ?? "(不明)")
                + " BaseDirectory="
                + AppContext.BaseDirectory);

        string? suppressIniPath = null;
        var consumedSlot = 0;
        try
        {
            var iniPath = ResolveIniPath(args, exeDir);
            if (string.IsNullOrWhiteSpace(iniPath))
            {
                LauncherLog.Error(
                    "ini パスが未指定です。--ini、PM_AI_RDP_LAUNCHER_INI、"
                        + "操作者名を第1引数（例: PmAiRdpRemoteLauncher.exe 細川 → 細川_RPA設定.ini）、"
                        + "または exe 同階層の RPA設定.ini を指定してください。");
                return ExitWithLog(ExitMissingIni, "ini パス未指定");
            }

            if (!File.Exists(iniPath))
            {
                LauncherLog.Error("ini が見つかりません: " + iniPath);
                return ExitWithLog(ExitMissingIni, "ini 不在");
            }

            LauncherLog.Info("ini パス: " + iniPath);
            suppressIniPath = iniPath;

            var ini = LauncherIni.Load(iniPath);
            consumedSlot = ini.SelectedSlot;
            LauncherLog.Info("起動プログラム番号=" + consumedSlot);
            if (ini.IsLauncherDisabled)
            {
                LauncherLog.Info(
                    "起動プログラム番号="
                        + LauncherIni.DisabledSlot
                        + " のため何もしません（RPA 起動・サインアウトなし）。");
                return ExitWithLog(ExitOk, "抑止（起動プログラム番号=0）");
            }

            var startedSlot = consumedSlot;
            var commandLine = ini.ResolveSelectedCommand();
            if (string.IsNullOrWhiteSpace(commandLine))
            {
                LauncherLog.Error("起動プログラム番号 " + ini.SelectedSlot + " に対応するスロットが ini にありません: " + iniPath);
                return ExitWithLog(ExitError, "スロット未定義");
            }

            ParsedCommand parsed;
            try
            {
                parsed = CommandLineParser.Parse(commandLine);
            }
            catch (FormatException ex)
            {
                LauncherLog.Error(ex.Message);
                return ExitWithLog(ExitError, "コマンド行解析失敗");
            }

            var rawExecutable = parsed.Executable;
            var rawArguments = parsed.Arguments;
            parsed = new ParsedCommand(
                UncPathSegmentRepair.Repair(parsed.Executable),
                RpaScenarioArgumentSupport.RepairScenarioArguments(parsed.Arguments));
            if (!string.Equals(rawExecutable, parsed.Executable, StringComparison.Ordinal)
                || !string.Equals(rawArguments, parsed.Arguments, StringComparison.Ordinal))
            {
                LauncherLog.Info(
                    "UNC パスを修復しました: exe="
                        + parsed.Executable
                        + " | args="
                        + parsed.Arguments);
            }

            foreach (var scenarioPath in RpaScenarioArgumentSupport.ExtractScenarioPaths(parsed.Arguments))
            {
                var repairedScenario = UncPathSegmentRepair.Repair(scenarioPath);
                if (File.Exists(repairedScenario))
                {
                    continue;
                }

                LauncherLog.Error("シナリオファイルが見つかりません: " + repairedScenario);
                return ExitWithLog(ExitError, "シナリオ不在");
            }

            var sessionEndAction = ini.ResolveSessionEndAction();
            LauncherLog.Info("終了時セッション操作: " + FormatSessionEndAction(sessionEndAction));
            LauncherLog.Info(
                "操作者="
                    + (string.IsNullOrWhiteSpace(ini.OperatorName) ? "(未設定)" : ini.OperatorName));

            var credentials = OperatorAladdinCredentialsStore.Resolve(iniPath, ini.OperatorName);
            if (credentials == null)
            {
                LauncherLog.Error(
                    "アラジン資格情報が未設定のため RPA を起動しません。"
                        + " PM-AI リモートデスクトップタブで資格情報を保存してから接続してください。");
                return ExitWithLog(ExitError, "アラジン資格情報未設定");
            }

            var iniArgumentTokens = WindowsArgumentFormatter.TokenizeForProcess(parsed.Arguments);
            if (AladdinRpaArgumentAppender.WouldStripEternalForScenario(iniArgumentTokens))
            {
                LauncherLog.Info(
                    "シナリオ指定のため --eternal を除去します（終了検知とセッション終了操作のため）。");
            }

            var argumentTokens = AladdinRpaArgumentAppender.AppendCredentials(
                iniArgumentTokens,
                credentials);
            var launchArguments = WindowsArgumentFormatter.FormatArgumentString(
                string.Join(" ", argumentTokens));
            var launchCommand = new ParsedCommand(parsed.Executable, launchArguments);

            if (!File.Exists(parsed.Executable))
            {
                LauncherLog.Error("起動プログラムが見つかりません: " + parsed.Executable);
                return ExitWithLog(ExitError, "起動プログラム不在");
            }

            Process? child = null;
            var existingProcessId = ProcessRunningChecker.TryFindRunningProcessId(
                launchCommand,
                credentials.LoginId);
            if (existingProcessId.HasValue)
            {
                LauncherLog.Info(
                    "既に起動済みのため監視のみ: PID="
                        + existingProcessId.Value
                        + " | "
                        + launchCommand.Executable
                        + " "
                        + launchArguments);
                child = TryOpenProcess(existingProcessId.Value);
                if (child == null)
                {
                    LauncherLog.Error("既存プロセスを開けませんでした PID=" + existingProcessId.Value);
                    return ExitWithLog(ExitError, "既存 PID オープン失敗");
                }
            }
            else
            {
                var workingDirectory = Path.GetDirectoryName(parsed.Executable);
                if (string.IsNullOrWhiteSpace(workingDirectory))
                {
                    workingDirectory = Environment.CurrentDirectory;
                }

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
                    LauncherLog.Error("子プロセスの起動に失敗しました: " + launchCommand.Executable + " " + launchArguments);
                    return ExitWithLog(ExitError, "子プロセス起動失敗");
                }

                LauncherLog.Info(
                    "起動しました PID="
                        + child.Id
                        + " | "
                        + launchCommand.Executable
                        + " "
                        + launchArguments);
            }

            TrySuppressIniSlot(iniPath, startedSlot);

            using (child)
            {
                var monitor = new ProcessTreeMonitor(child, launchCommand, credentials.LoginId);
                monitor.WaitUntilFinished();
                LauncherLog.Info("子プロセス終了 PID=" + child.Id + " ExitCode=" + FormatExitCode(child));
            }

            if (sessionEndAction != SessionEndAction.None)
            {
                var sessionEndSucceeded = sessionEndAction switch
                {
                    SessionEndAction.Disconnect =>
                        RdpSessionDisconnecter.TryDisconnectCurrentSession(out var disconnectError)
                            ? LogSessionEndSuccess("RDP セッションを切断しました")
                            : LogSessionEndFailure("RDP 切断失敗: " + disconnectError),
                    SessionEndAction.SignOut =>
                        RdpSessionSignOuter.TrySignOutCurrentSession(out var signOutError)
                            ? LogSessionEndSuccess("RDP セッションをサインアウトしました")
                            : LogSessionEndFailure("サインアウト失敗: " + signOutError),
                    _ => true,
                };
                _ = sessionEndSucceeded;
            }

            return ExitWithLog(ExitOk, "正常終了");
        }
        catch (Exception ex)
        {
            LauncherLog.Error(ex.Message);
            return ExitWithLog(ExitError, "例外: " + ex.Message);
        }
        finally
        {
            if (consumedSlot > LauncherIni.DisabledSlot && !string.IsNullOrWhiteSpace(suppressIniPath))
            {
                TrySuppressIniSlot(suppressIniPath, consumedSlot);
            }
        }
    }

    private static void TrySuppressIniSlot(string iniPath, int startedSlot)
    {
        try
        {
            LauncherIni.WriteSelectedSlot(iniPath, LauncherIni.DisabledSlot);
            LauncherLog.Info(
                "起動プログラム番号を 0 に設定しました（"
                    + startedSlot
                    + " → 0）。タスクスケジューラ再実行時の二重起動を抑止します。");
        }
        catch (Exception ex)
        {
            LauncherLog.Error("起動プログラム番号の 0 設定に失敗: " + ex.Message);
        }
    }

    private static bool LogSessionEndSuccess(string message)
    {
        LauncherLog.Info(message);
        return true;
    }

    private static bool LogSessionEndFailure(string message)
    {
        LauncherLog.Error(message);
        return false;
    }

    private static string FormatSessionEndAction(SessionEndAction action)
    {
        return action switch
        {
            SessionEndAction.None => "なし",
            SessionEndAction.Disconnect => "切断",
            SessionEndAction.SignOut => "サインアウト",
            _ => action.ToString(),
        };
    }

    private static int ExitWithLog(int exitCode, string reason)
    {
        LauncherLog.Info("PmAiRdpRemoteLauncher 終了 exitCode=" + exitCode + " reason=" + reason);
        return exitCode;
    }

    private static string FormatExitCode(Process process)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Refresh();
            }

            return process.HasExited ? process.ExitCode.ToString() : "(実行中)";
        }
        catch
        {
            return "(不明)";
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

    private static string ResolveLauncherVersionLabel()
    {
        try
        {
            var exeDir = LauncherPaths.ResolveExecutableDirectory();
            if (string.IsNullOrWhiteSpace(exeDir))
            {
                return "(不明)";
            }

            var versionPath = Path.Combine(exeDir, "PmAiRdpRemoteLauncher.version.txt");
            if (!File.Exists(versionPath))
            {
                return "(version.txt なし)";
            }

            var line = File.ReadLines(versionPath).FirstOrDefault()?.Trim();
            return string.IsNullOrWhiteSpace(line) ? "(空)" : line;
        }
        catch
        {
            return "(読取失敗)";
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

        var operatorFromArgs = LauncherIni.TryParseOperatorArgument(args);
        if (!string.IsNullOrWhiteSpace(operatorFromArgs))
        {
            LauncherLog.Info("操作者名引数: " + operatorFromArgs);
            return LauncherIni.ResolveIniPathInDeployLayout(exeDir, operatorFromArgs);
        }

        return LauncherIni.ResolveIniPathInDeployLayout(exeDir, null);
    }
}
