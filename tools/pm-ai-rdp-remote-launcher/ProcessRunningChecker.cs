using System.Management;

namespace PmAi.RdpRemoteLauncher;

internal static class ProcessRunningChecker
{
    internal static bool IsAlreadyRunning(ParsedCommand command, string? loginId = null)
    {
        return TryFindRunningProcessId(command, loginId).HasValue;
    }

    internal static int? TryFindRunningProcessId(ParsedCommand command, string? loginId = null)
    {
        var executable = NormalizePath(LauncherPaths.NormalizeExecutablePath(command.Executable));
        var processName = Path.GetFileNameWithoutExtension(executable);
        if (string.IsNullOrWhiteSpace(processName))
        {
            return null;
        }

        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT ProcessId, CommandLine, ExecutablePath FROM Win32_Process WHERE Name = '"
                    + EscapeWmi(processName + ".exe")
                    + "'");
            foreach (ManagementObject obj in searcher.Get())
            {
                using (obj)
                {
                    var commandLine = obj["CommandLine"]?.ToString() ?? string.Empty;
                    var executablePath = obj["ExecutablePath"]?.ToString() ?? string.Empty;
                    if (!Matches(command, loginId, commandLine, executablePath, executable))
                    {
                        continue;
                    }

                    var processIdRaw = obj["ProcessId"];
                    if (processIdRaw != null && uint.TryParse(processIdRaw.ToString(), out var processId))
                    {
                        return (int)processId;
                    }

                    return null;
                }
            }
        }
        catch (ManagementException)
        {
            // WMI 不可時は重複判定をスキップし起動を試みる。
            return null;
        }

        return null;
    }

    private static bool Matches(
        ParsedCommand command,
        string? loginId,
        string commandLine,
        string executablePath,
        string executable)
    {
        if (!ExecutableMatches(executable, executablePath, commandLine))
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(command.Arguments))
        {
            // 引数なし起動は常に新規起動する（--id 付与は起動直前のため既存判定しない）。
            return false;
        }

        if (string.IsNullOrWhiteSpace(loginId))
        {
            return false;
        }

        if (!commandLine.Contains(AladdinRpaLaunchArgs.IdFlag, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!commandLine.Contains(loginId.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        foreach (var token in WindowsArgumentFormatter.Tokenize(command.Arguments))
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                continue;
            }

            if (IsCredentialFlag(token) || string.Equals(token, loginId.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!commandLine.Contains(token, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        return true;
    }

    private static bool ExecutableMatches(string executable, string executablePath, string commandLine)
    {
        if (!string.IsNullOrWhiteSpace(executablePath)
            && string.Equals(NormalizePath(executablePath), executable, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (string.IsNullOrWhiteSpace(commandLine))
        {
            return false;
        }

        return commandLine.Contains(executable, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsCredentialFlag(string token)
    {
        return string.Equals(token, AladdinRpaLaunchArgs.IdFlag, StringComparison.OrdinalIgnoreCase)
            || string.Equals(token, AladdinRpaLaunchArgs.PasswordFlag, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizePath(string path)
    {
        try
        {
            return Path.GetFullPath(path.Trim().Trim('"'));
        }
        catch
        {
            return path.Trim().Trim('"');
        }
    }

    private static string EscapeWmi(string value)
    {
        return value.Replace("\\", "\\\\").Replace("'", "\\'");
    }
}
