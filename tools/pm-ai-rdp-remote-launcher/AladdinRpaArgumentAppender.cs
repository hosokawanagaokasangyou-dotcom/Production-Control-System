namespace PmAi.RdpRemoteLauncher;

internal static class AladdinRpaArgumentAppender
{
    /// <summary>
    /// Aladdin RPA 起動引数を組み立てる。
    /// 資格情報（--id / --password）を先頭、--eternal は末尾（シナリオなし起動時のログイン失敗回避）。
    /// </summary>
    internal static List<string> AppendCredentials(
        IReadOnlyList<string> tokens,
        OperatorAladdinCredentials credentials)
    {
        var others = new List<string>(tokens);
        RemoveExistingCredentialFlags(others);
        var hasEternal = RemoveFlag(others, AladdinRpaLaunchArgs.EternalFlag);
        NormalizeScenarioTokens(others);

        var result = new List<string>
        {
            AladdinRpaLaunchArgs.IdFlag,
            credentials.LoginId,
            AladdinRpaLaunchArgs.PasswordFlag,
            credentials.Password,
        };
        result.AddRange(others);
        if (hasEternal)
        {
            result.Add(AladdinRpaLaunchArgs.EternalFlag);
        }

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

    private static bool RemoveFlag(List<string> tokens, string flag)
    {
        var found = false;
        for (var i = tokens.Count - 1; i >= 0; i--)
        {
            if (!string.Equals(tokens[i], flag, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            tokens.RemoveAt(i);
            found = true;
        }

        return found;
    }

    private static bool IsCredentialFlag(string token)
    {
        return string.Equals(token, AladdinRpaLaunchArgs.IdFlag, StringComparison.OrdinalIgnoreCase)
            || string.Equals(token, AladdinRpaLaunchArgs.PasswordFlag, StringComparison.OrdinalIgnoreCase);
    }

    /** 旧 ini（.ardrpa パスのみ）を {@code --scenario path} に揃える。 */
    private static void NormalizeScenarioTokens(List<string> tokens)
    {
        var scenarioPaths = new List<string>();
        var others = new List<string>();
        for (var i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];
            if (string.Equals(token, AladdinRpaLaunchArgs.ScenarioFlag, StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 < tokens.Count)
                {
                    var (path, nextIndex) = CollectScenarioPathAfterFlag(tokens, i + 1);
                    scenarioPaths.Add(path);
                    i = nextIndex;
                }

                continue;
            }

            if (LooksLikeScenarioPath(token))
            {
                scenarioPaths.Add(token);
            }
            else if (!string.IsNullOrWhiteSpace(token))
            {
                others.Add(token);
            }
        }

        tokens.Clear();
        tokens.AddRange(others);
        foreach (var path in scenarioPaths)
        {
            tokens.Add(AladdinRpaLaunchArgs.ScenarioFlag);
            tokens.Add(path);
        }
    }

    private static (string Path, int Index) CollectScenarioPathAfterFlag(
        IReadOnlyList<string> tokens,
        int startIndex)
    {
        var parts = new List<string> { tokens[startIndex] };
        var index = startIndex;
        while (index + 1 < tokens.Count && !LooksLikeScenarioPath(string.Join(" ", parts)))
        {
            index++;
            parts.Add(tokens[index]);
        }

        return (string.Join(" ", parts), index);
    }

    private static bool LooksLikeScenarioPath(string token)
    {
        return token.EndsWith(".ardrpa", StringComparison.OrdinalIgnoreCase);
    }
}
