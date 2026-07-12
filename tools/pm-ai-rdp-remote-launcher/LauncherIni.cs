using System.Text;

namespace PmAi.RdpRemoteLauncher;

internal sealed class LauncherIni
{
    internal const string SelectedSlotKey = "起動プログラム番号";
    internal const string OperatorKey = "操作者";
    internal const string DisconnectOnChildExitKey = "終了時RDP切断";
    internal const string SessionEndActionKey = "終了時セッション操作";
    /** 後方互換: 旧方式の接続時サインアウトフラグ（新方式は 99=--signout）。 */
    internal const string SignOutOnConnectKey = "接続時サインアウト";
    internal const string SignOutLauncherArgs = "--signout";
    internal const string DefaultIniFileName = "RPA設定.ini";
    internal const string LegacyIniFileName = "RAP設定.ini";
    internal const string IniPathEnvVar = "PM_AI_RDP_LAUNCHER_INI";
    internal const string DisconnectOnChildExitEnvVar = "PM_AI_RDP_DISCONNECT_ON_CHILD_EXIT";
    internal const string SessionEndActionEnvVar = "PM_AI_RDP_SESSION_END_ACTION";

    /** 接続先サインアウト専用スロット（ini 起動プログラム番号・スロット定義とも 99）。 */
    internal const int SignOutSlot = 99;
    /** タスクスケジューラ RPA 抑止専用（サインアウトしない）。 */
    internal const int LegacySignOutSlot = 0;
    /** @deprecated SignOutSlot を使用 */
    internal const int DisabledSlot = SignOutSlot;

    internal int SelectedSlot { get; set; } = 1;

    /** 接続直前に PM-AI が書くセッション操作者名。 */
    internal string OperatorName { get; set; } = "";

    /** 子プロセス終了後に RDP セッション操作を行う（後方互換。SessionEndAction が正）。 */
    internal bool DisconnectOnChildExit { get; set; } = true;

    /** 子プロセス終了後のセッション操作（既定サインアウト）。 */
    internal SessionEndAction SessionEndAction { get; set; } = SessionEndAction.SignOut;

    /** 旧方式: {@link SignOutOnConnectKey}=1（後方互換読取のみ）。 */
    internal bool SignOutOnConnectRequested { get; set; }

    internal Dictionary<int, string> Slots { get; } = new();

    internal static string BuildIniFileNameForOperator(string? operatorName)
    {
        var trimmed = operatorName?.Trim() ?? "";
        if (string.IsNullOrEmpty(trimmed))
        {
            return DefaultIniFileName;
        }

        return SanitizeOperatorForIniFilename(trimmed) + "_" + DefaultIniFileName;
    }

    internal static string BuildLegacyIniFileNameForOperator(string? operatorName)
    {
        var trimmed = operatorName?.Trim() ?? "";
        if (string.IsNullOrEmpty(trimmed))
        {
            return LegacyIniFileName;
        }

        return SanitizeOperatorForIniFilename(trimmed) + "_" + LegacyIniFileName;
    }

    internal static string SanitizeOperatorForIniFilename(string operatorName)
    {
        if (string.IsNullOrWhiteSpace(operatorName))
        {
            return "operator";
        }

        var s = operatorName.Trim();
        foreach (var ch in new[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' })
        {
            s = s.Replace(ch, '_');
        }

        while (s.EndsWith('.') || s.EndsWith(' '))
        {
            s = s[..^1].TrimEnd();
        }

        return string.IsNullOrEmpty(s) ? "operator" : s;
    }

    /// exe と同階層の ini を解決する（操作者名あり: {操作者}_RPA設定.ini、なし: RPA設定.ini。レガシー RAP設定.ini / DATA も試行）。
    internal static string ResolveIniPathInDeployLayout(string? exeDir, string? operatorName)
    {
        var directory = exeDir;
        if (string.IsNullOrWhiteSpace(directory))
        {
            directory = AppContext.BaseDirectory;
        }

        if (string.IsNullOrWhiteSpace(directory))
        {
            directory = Environment.CurrentDirectory;
        }

        var op = operatorName?.Trim() ?? "";
        if (!string.IsNullOrEmpty(op))
        {
            var perUser = Path.Combine(directory, BuildIniFileNameForOperator(op));
            if (File.Exists(perUser))
            {
                return perUser;
            }

            var perUserLegacy = Path.Combine(directory, BuildLegacyIniFileNameForOperator(op));
            if (File.Exists(perUserLegacy))
            {
                return perUserLegacy;
            }

            var dataPerUser = Path.Combine(directory, "DATA", BuildIniFileNameForOperator(op));
            if (File.Exists(dataPerUser))
            {
                return dataPerUser;
            }

            var dataPerUserLegacy =
                Path.Combine(directory, "DATA", BuildLegacyIniFileNameForOperator(op));
            if (File.Exists(dataPerUserLegacy))
            {
                return dataPerUserLegacy;
            }

            return perUser;
        }

        var sameDirRpa = Path.Combine(directory, DefaultIniFileName);
        if (File.Exists(sameDirRpa))
        {
            return sameDirRpa;
        }

        var sameDirRap = Path.Combine(directory, LegacyIniFileName);
        if (File.Exists(sameDirRap))
        {
            return sameDirRap;
        }

        var dataRpa = Path.Combine(directory, "DATA", DefaultIniFileName);
        if (File.Exists(dataRpa))
        {
            return dataRpa;
        }

        var dataRap = Path.Combine(directory, "DATA", LegacyIniFileName);
        if (File.Exists(dataRap))
        {
            return dataRap;
        }

        return sameDirRpa;
    }

    internal static string? TryParseOperatorArgument(string[] args)
    {
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg == "--ini" && i + 1 < args.Length)
            {
                i++;
                continue;
            }

            if (arg.StartsWith("--ini=", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (arg.StartsWith("--", StringComparison.Ordinal))
            {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(arg))
            {
                return arg.Trim();
            }
        }

        return null;
    }

    internal static LauncherIni Load(string path)
    {
        var ini = new LauncherIni();
        var sawSessionEndActionKey = false;
        foreach (var rawLine in File.ReadAllLines(path, Encoding.UTF8))
        {
            var line = rawLine.Trim();
            if (line.Length == 0 || line.StartsWith('#') || line.StartsWith(';'))
            {
                continue;
            }

            var eq = line.IndexOf('=');
            if (eq <= 0)
            {
                continue;
            }

            var key = line[..eq].Trim();
            var value = line[(eq + 1)..].Trim();
            if (key == SelectedSlotKey)
            {
                if (int.TryParse(value, out var slot) && slot >= LegacySignOutSlot)
                {
                    ini.SelectedSlot = slot;
                }
                continue;
            }

            if (key == DisconnectOnChildExitKey)
            {
                ini.DisconnectOnChildExit = ParseBoolean(value, defaultValue: true);
                continue;
            }

            if (key == SessionEndActionKey)
            {
                ini.SessionEndAction = ParseSessionEndAction(value, ini.SessionEndAction);
                sawSessionEndActionKey = true;
                continue;
            }

            if (key == OperatorKey)
            {
                ini.OperatorName = value;
                continue;
            }

            if (key == SignOutOnConnectKey)
            {
                ini.SignOutOnConnectRequested = ParseBoolean(value, defaultValue: false);
                continue;
            }

            if (int.TryParse(key, out var slotNumber)
                && (slotNumber >= 1 || slotNumber == SignOutSlot)
                && !string.IsNullOrWhiteSpace(value))
            {
                ini.Slots[slotNumber] = value;
            }
        }

        if (!sawSessionEndActionKey)
        {
            ini.SessionEndAction = ini.DisconnectOnChildExit
                ? SessionEndAction.SignOut
                : SessionEndAction.None;
        }

        return ini;
    }

    internal bool IsSuppressOnly => SelectedSlot == LegacySignOutSlot;

    internal bool IsSignOutSlotSelected => SelectedSlot == SignOutSlot;

    internal bool IsSignOutOnly => IsSuppressOnly || IsSignOutSlotSelected;

    internal bool IsLauncherDisabled => IsSuppressOnly;

    internal static bool IsSignOutSlotCommand(string? commandLine)
    {
        if (string.IsNullOrWhiteSpace(commandLine))
        {
            return false;
        }

        try
        {
            var parsed = CommandLineParser.Parse(commandLine);
            return string.Equals(
                parsed.Executable,
                SignOutLauncherArgs,
                StringComparison.OrdinalIgnoreCase);
        }
        catch (FormatException)
        {
            return string.Equals(
                commandLine.Trim(),
                SignOutLauncherArgs,
                StringComparison.OrdinalIgnoreCase);
        }
    }

    internal string? ResolveSelectedCommand()
    {
        if (IsSuppressOnly)
        {
            return null;
        }

        if (Slots.TryGetValue(SelectedSlot, out var command) && !string.IsNullOrWhiteSpace(command))
        {
            return command;
        }

        if (IsSignOutSlotSelected)
        {
            return SignOutLauncherArgs;
        }

        return null;
    }

    internal bool ResolveDisconnectOnChildExit()
    {
        return ResolveSessionEndAction() != SessionEndAction.None;
    }

    internal SessionEndAction ResolveSessionEndAction()
    {
        var fromEnv = Environment.GetEnvironmentVariable(SessionEndActionEnvVar);
        if (!string.IsNullOrWhiteSpace(fromEnv))
        {
            return ParseSessionEndAction(fromEnv, SessionEndAction);
        }

        var legacyEnv = Environment.GetEnvironmentVariable(DisconnectOnChildExitEnvVar);
        if (!string.IsNullOrWhiteSpace(legacyEnv) && !ParseBoolean(legacyEnv, DisconnectOnChildExit))
        {
            return SessionEndAction.None;
        }

        return SessionEndAction;
    }

    /// <summary>
    /// 起動プログラム番号行のみ更新する。スロット定義行は保持する。
    /// </summary>
    internal static void WriteSelectedSlot(string path, int selectedSlot)
    {
        if (selectedSlot < LegacySignOutSlot)
        {
            throw new ArgumentOutOfRangeException(nameof(selectedSlot), selectedSlot, "slot must be >= 0");
        }

        var lines = File.ReadAllLines(path, Encoding.UTF8).ToList();
        var keyPrefix = SelectedSlotKey + "=";
        var replaced = false;
        for (var i = 0; i < lines.Count; i++)
        {
            var trimmed = lines[i].Trim();
            if (trimmed.Length == 0 || trimmed.StartsWith('#') || trimmed.StartsWith(';'))
            {
                continue;
            }

            if (!trimmed.StartsWith(keyPrefix, StringComparison.Ordinal))
            {
                continue;
            }

            lines[i] = SelectedSlotKey + "=" + selectedSlot;
            replaced = true;
            break;
        }

        if (!replaced)
        {
            lines.Insert(0, SelectedSlotKey + "=" + selectedSlot);
        }

        File.WriteAllLines(path, lines, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    /// <summary>
    /// 抑止用に ini ファイルへ 0 をそのまま書く。
    /// </summary>
    internal static void WriteSuppressSlotLiteral(string path)
    {
        var lines = File.ReadAllLines(path, Encoding.UTF8).ToList();
        var keyPrefix = SelectedSlotKey + "=";
        var replaced = false;
        for (var i = 0; i < lines.Count; i++)
        {
            var trimmed = lines[i].Trim();
            if (trimmed.Length == 0 || trimmed.StartsWith('#') || trimmed.StartsWith(';'))
            {
                continue;
            }

            if (!trimmed.StartsWith(keyPrefix, StringComparison.Ordinal))
            {
                continue;
            }

            lines[i] = SelectedSlotKey + "=" + LegacySignOutSlot;
            replaced = true;
            break;
        }

        if (!replaced)
        {
            lines.Insert(0, SelectedSlotKey + "=" + LegacySignOutSlot);
        }

        File.WriteAllLines(path, lines, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    internal static void ClearSignOutOnConnectRequest(string path)
    {
        MergeIniScalarKey(path, SignOutOnConnectKey, "0");
    }

    private static void MergeIniScalarKey(string path, string key, string value)
    {
        var lines = File.ReadAllLines(path, Encoding.UTF8).ToList();
        var keyPrefix = key + "=";
        var replaced = false;
        for (var i = 0; i < lines.Count; i++)
        {
            var trimmed = lines[i].Trim();
            if (trimmed.Length == 0 || trimmed.StartsWith('#') || trimmed.StartsWith(';'))
            {
                continue;
            }

            if (!trimmed.StartsWith(keyPrefix, StringComparison.Ordinal))
            {
                continue;
            }

            lines[i] = key + "=" + value;
            replaced = true;
            break;
        }

        if (!replaced)
        {
            lines.Add(key + "=" + value);
        }

        File.WriteAllLines(path, lines, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    private static SessionEndAction ParseSessionEndAction(string raw, SessionEndAction defaultValue)
    {
        return SessionEndActionParser.Parse(raw, defaultValue);
    }

    private static bool ParseBoolean(string raw, bool defaultValue)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return defaultValue;
        }

        var v = raw.Trim().ToLowerInvariant();
        return v switch
        {
            "1" or "true" or "on" or "yes" => true,
            "0" or "false" or "off" or "no" => false,
            _ => defaultValue,
        };
    }
}
