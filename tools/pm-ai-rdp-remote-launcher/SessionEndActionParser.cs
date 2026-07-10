namespace PmAi.RdpRemoteLauncher;

internal static class SessionEndActionParser
{
    internal static SessionEndAction Parse(string raw, SessionEndAction defaultValue)
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

    internal static string Format(SessionEndAction action)
    {
        return action switch
        {
            SessionEndAction.None => "なし",
            SessionEndAction.Disconnect => "切断",
            SessionEndAction.SignOut => "サインアウト",
            _ => action.ToString(),
        };
    }
}
