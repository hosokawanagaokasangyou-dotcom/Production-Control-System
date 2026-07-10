namespace PmAi.RdpRemoteLauncher;

internal static class SessionEndExecutor
{
    internal static bool TryExecute(SessionEndAction action, out string? errorMessage)
    {
        errorMessage = null;
        if (action == SessionEndAction.None)
        {
            LauncherLog.Info("セッション終了操作は「なし」のためスキップします");
            return true;
        }

        return action switch
        {
            SessionEndAction.Disconnect =>
                RdpSessionDisconnecter.TryDisconnectCurrentSession(out errorMessage)
                    ? LogSuccess("RDP セッションを切断しました")
                    : LogFailure("RDP 切断失敗: " + errorMessage),
            SessionEndAction.SignOut =>
                RdpSessionSignOuter.TrySignOutCurrentSession(out errorMessage)
                    ? LogSuccess("RDP セッションをサインアウトしました")
                    : LogFailure("サインアウト失敗: " + errorMessage),
            _ => true,
        };
    }

    private static bool LogSuccess(string message)
    {
        LauncherLog.Info(message);
        return true;
    }

    private static bool LogFailure(string message)
    {
        LauncherLog.Error(message);
        return false;
    }
}
