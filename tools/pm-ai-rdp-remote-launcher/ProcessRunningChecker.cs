using System.Management;
using System.Text;

namespace PmAi.RdpRemoteLauncher;

internal static class ProcessRunningChecker
{
    internal static bool IsAlreadyRunning(ParsedCommand command)
    {
        return TryFindRunningProcessId(command).HasValue;
    }

    internal static int? TryFindRunningProcessId(ParsedCommand command)
    {
        var executable = NormalizePath(command.Executable);
        var processName = Path.GetFileNameWithoutExtension(executable);
        if (string.IsNullOrWhiteSpace(processName))
        {
            return null;
        }

        var signature = BuildMatchSignature(executable, command.Arguments);
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
                    if (Matches(executable, command.Arguments, signature, commandLine, executablePath))
                    {
                        var processIdRaw = obj["ProcessId"];
                        if (processIdRaw != null && uint.TryParse(processIdRaw.ToString(), out var processId))
                        {
                            return (int)processId;
                        }

                        return null;
                    }
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
        string executable,
        string arguments,
        string signature,
        string commandLine,
        string executablePath)
    {
        if (!string.IsNullOrWhiteSpace(executablePath)
            && string.Equals(NormalizePath(executablePath), executable, StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(arguments))
            {
                return true;
            }

            if (commandLine.Contains(arguments, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        if (string.IsNullOrWhiteSpace(commandLine))
        {
            return false;
        }

        return commandLine.Contains(signature, StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildMatchSignature(string executable, string arguments)
    {
        if (!string.IsNullOrWhiteSpace(arguments))
        {
            var argToken = ExtractArgumentToken(arguments);
            if (!string.IsNullOrWhiteSpace(argToken))
            {
                return argToken;
            }
        }

        return executable;
    }

    private static string ExtractArgumentToken(string arguments)
    {
        var trimmed = arguments.Trim();
        if (trimmed.Length == 0)
        {
            return string.Empty;
        }

        if (trimmed[0] == '"')
        {
            var end = trimmed.IndexOf('"', 1);
            return end > 0 ? trimmed[1..end] : trimmed.Trim('"');
        }

        var space = trimmed.IndexOf(' ');
        return space < 0 ? trimmed : trimmed[..space];
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
