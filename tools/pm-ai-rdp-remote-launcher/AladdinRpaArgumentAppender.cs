namespace PmAi.RdpRemoteLauncher;

internal static class AladdinRpaArgumentAppender
{
    internal static List<string> AppendCredentials(
        IReadOnlyList<string> tokens,
        OperatorAladdinCredentials credentials)
    {
        var result = new List<string>(tokens);
        RemoveExistingCredentialFlags(result);
        result.Add(AladdinRpaLaunchArgs.IdFlag);
        result.Add(credentials.LoginId);
        result.Add(AladdinRpaLaunchArgs.PasswordFlag);
        result.Add(credentials.Password);
        return result;
    }

    private static void RemoveExistingCredentialFlags(List<string> tokens)
    {
        for (var i = 0; i < tokens.Count; i++)
        {
            if (!IsCredentialFlag(tokens[i]))
            {
                continue;
            }

            tokens.RemoveAt(i);
            if (i < tokens.Count)
            {
                tokens.RemoveAt(i);
            }

            i = Math.Max(-1, i - 1);
        }
    }

    private static bool IsCredentialFlag(string token)
    {
        return string.Equals(token, AladdinRpaLaunchArgs.IdFlag, StringComparison.OrdinalIgnoreCase)
            || string.Equals(token, AladdinRpaLaunchArgs.PasswordFlag, StringComparison.OrdinalIgnoreCase);
    }
}
