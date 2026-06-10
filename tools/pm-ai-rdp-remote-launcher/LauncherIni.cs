using System.Text;

namespace PmAi.RdpRemoteLauncher;

internal sealed class LauncherIni
{
    internal const string SelectedSlotKey = "起動プログラム番号";
    internal const string OperatorKey = "操作者";
    internal const string DisconnectOnChildExitKey = "終了時RDP切断";
    internal const string SessionEndActionKey = "終了時セッション操作";
    internal const string DefaultIniFileName = "RAP設定.ini";
    internal const string IniPathEnvVar = "PM_AI_RDP_LAUNCHER_INI";
    internal const string DisconnectOnChildExitEnvVar = "PM_AI_RDP_DISCONNECT_ON_CHILD_EXIT";
    internal const string SessionEndActionEnvVar = "PM_AI_RDP_SESSION_END_ACTION";

    /** タスクスケジューラ経由の自動起動を抑止（RPA 起動・サインアウトともに行わない）。 */
    internal const int DisabledSlot = 0;

    internal int SelectedSlot { get; set; } = 1;

    /** 接続直前に PM-AI が書くセッション操作者名。 */
    internal string OperatorName { get; set; } = "";

    /** 子プロセス終了後に RDP セッション操作を行う（後方互換。SessionEndAction が正）。 */
    internal bool DisconnectOnChildExit { get; set; } = true;

    /** 子プロセス終了後のセッション操作（既定サインアウト）。 */
    internal SessionEndAction SessionEndAction { get; set; } = SessionEndAction.SignOut;

    internal Dictionary<int, string> Slots { get; } = new();

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
                if (int.TryParse(value, out var slot) && slot >= DisabledSlot)
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

            if (int.TryParse(key, out var slotNumber) && slotNumber >= 1 && !string.IsNullOrWhiteSpace(value))
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

    internal bool IsLauncherDisabled => SelectedSlot == DisabledSlot;

    internal string? ResolveSelectedCommand()
    {
        if (IsLauncherDisabled)
        {
            return null;
        }

        return Slots.TryGetValue(SelectedSlot, out var command) ? command : null;
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
        if (selectedSlot < DisabledSlot)
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

    private static SessionEndAction ParseSessionEndAction(string raw, SessionEndAction defaultValue)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return defaultValue;
        }

        var normalized = raw.Trim().ToLowerInvariant();
        return normalized switch
        {
            "なし" or "none" or "off" or "0" => SessionEndAction.None,
            "切断" or "disconnect" or "rdp切断" or "rdp_disconnect" => SessionEndAction.Disconnect,
            "サインアウト" or "signout" or "sign_out" or "logoff" or "ログオフ" => SessionEndAction.SignOut,
            _ => defaultValue,
        };
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
