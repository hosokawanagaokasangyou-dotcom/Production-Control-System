namespace PmAi.RdpRemoteLauncher;

/// <summary>
/// RPA 起動引数内の {@code --scenario} パスを、空白トークン分割なしで抽出・正規化する。
/// </summary>
internal static class RpaScenarioArgumentSupport
{
    internal static string NormalizeScenarioArguments(string? arguments)
    {
        return RebuildScenarioArguments(arguments, stripEternalFlag: true);
    }

    internal static string RepairScenarioArguments(string? arguments)
    {
        return RebuildScenarioArguments(arguments, stripEternalFlag: false);
    }

    internal static IReadOnlyList<string> ExtractScenarioPaths(string? arguments)
    {
        if (string.IsNullOrWhiteSpace(arguments))
        {
            return Array.Empty<string>();
        }

        return ParseScenarioAndOtherArguments(arguments.Trim(), stripEternalFlag: false).ScenarioPaths;
    }

    private static string RebuildScenarioArguments(string? arguments, bool stripEternalFlag)
    {
        if (string.IsNullOrWhiteSpace(arguments))
        {
            return string.Empty;
        }

        var parsed = ParseScenarioAndOtherArguments(arguments.Trim(), stripEternalFlag);
        var normalized = new List<string>(parsed.OtherTokens);
        foreach (var path in parsed.ScenarioPaths)
        {
            normalized.Add(AladdinRpaLaunchArgs.ScenarioFlag);
            normalized.Add(UncPathSegmentRepair.Repair(path));
        }

        if (normalized.Count == 0)
        {
            return string.Empty;
        }

        return WindowsArgumentFormatter.FormatArgumentTokens(normalized);
    }

    private static ParsedArguments ParseScenarioAndOtherArguments(string input)
    {
        return ParseScenarioAndOtherArguments(input, stripEternalFlag: true);
    }

    private static ParsedArguments ParseScenarioAndOtherArguments(
        string input,
        bool stripEternalFlag)
    {
        var scenarioPaths = new List<string>();
        var otherTokens = new List<string>();
        var index = 0;
        while (index < input.Length)
        {
            index = SkipWhitespace(input, index);
            if (index >= input.Length)
            {
                break;
            }

            if (StartsWithFlag(input, index, AladdinRpaLaunchArgs.ScenarioFlag))
            {
                index += AladdinRpaLaunchArgs.ScenarioFlag.Length;
                index = SkipWhitespace(input, index);
                var extracted = ExtractScenarioPath(input, index);
                if (!string.IsNullOrWhiteSpace(extracted.Path))
                {
                    scenarioPaths.Add(extracted.Path);
                }

                index = extracted.EndIndex;
                continue;
            }

            if (TryExtractBareScenarioPath(input, index, out var barePath, out var bareEnd))
            {
                scenarioPaths.Add(barePath);
                index = bareEnd;
                continue;
            }

            if (StartsWithFlag(input, index, AladdinRpaLaunchArgs.IdFlag)
                || StartsWithFlag(input, index, AladdinRpaLaunchArgs.PasswordFlag))
            {
                index = SkipFlagWithValue(input, index);
                continue;
            }

            var tokenExtracted = ReadNextToken(input, index);
            if (!string.IsNullOrEmpty(tokenExtracted.Token))
            {
                if (stripEternalFlag
                    && string.Equals(
                        tokenExtracted.Token,
                        AladdinRpaLaunchArgs.EternalFlag,
                        StringComparison.OrdinalIgnoreCase))
                {
                    index = tokenExtracted.EndIndex;
                    continue;
                }

                otherTokens.Add(tokenExtracted.Token);
            }

            index = tokenExtracted.EndIndex;
        }

        return new ParsedArguments(scenarioPaths, otherTokens);
    }

    private static ExtractedPath ExtractScenarioPath(string input, int startIndex)
    {
        if (startIndex >= input.Length)
        {
            return new ExtractedPath(string.Empty, startIndex);
        }

        if (input[startIndex] == '"')
        {
            var end = FindClosingQuote(input, startIndex + 1);
            if (end < 0)
            {
                return new ExtractedPath(string.Empty, startIndex);
            }

            var path = input[(startIndex + 1)..end].Replace("\"\"", "\"");
            return new ExtractedPath(path, end + 1);
        }

        var ardrpaIndex = IndexOfIgnoreCase(input, ".ardrpa", startIndex);
        if (ardrpaIndex < 0)
        {
            var token = ReadNextToken(input, startIndex);
            return new ExtractedPath(token.Token, token.EndIndex);
        }

        var endIndex = ardrpaIndex + ".ardrpa".Length;
        return new ExtractedPath(input[startIndex..endIndex], endIndex);
    }

    private static bool TryExtractBareScenarioPath(
        string input,
        int startIndex,
        out string path,
        out int endIndex)
    {
        path = string.Empty;
        endIndex = startIndex;
        if (startIndex >= input.Length)
        {
            return false;
        }

        if (input[startIndex] == '"')
        {
            return false;
        }

        if (!LooksLikePathStart(input, startIndex))
        {
            return false;
        }

        var ardrpaIndex = IndexOfIgnoreCase(input, ".ardrpa", startIndex);
        if (ardrpaIndex < 0)
        {
            return false;
        }

        endIndex = ardrpaIndex + ".ardrpa".Length;
        path = input[startIndex..endIndex];
        return true;
    }

    private static bool LooksLikePathStart(string input, int startIndex)
    {
        if (startIndex + 1 < input.Length && input[startIndex] == '\\' && input[startIndex + 1] == '\\')
        {
            return true;
        }

        return startIndex + 2 < input.Length
               && char.IsLetter(input[startIndex])
               && input[startIndex + 1] == ':';
    }

    private static TokenExtract ReadNextToken(string input, int startIndex)
    {
        if (startIndex >= input.Length)
        {
            return new TokenExtract(string.Empty, startIndex);
        }

        if (input[startIndex] == '"')
        {
            var end = FindClosingQuote(input, startIndex + 1);
            if (end < 0)
            {
                return new TokenExtract(input[startIndex..], input.Length);
            }

            return new TokenExtract(
                input[(startIndex + 1)..end].Replace("\"\"", "\""),
                end + 1);
        }

        var index = startIndex;
        while (index < input.Length && !char.IsWhiteSpace(input[index]))
        {
            index++;
        }

        return new TokenExtract(input[startIndex..index], index);
    }

    private static int SkipFlagWithValue(string input, int startIndex)
    {
        var token = ReadNextToken(input, startIndex);
        var index = SkipWhitespace(input, token.EndIndex);
        if (index >= input.Length || input[index] == '-')
        {
            return index;
        }

        return ReadNextToken(input, index).EndIndex;
    }

    private static int SkipWhitespace(string input, int startIndex)
    {
        var index = startIndex;
        while (index < input.Length && char.IsWhiteSpace(input[index]))
        {
            index++;
        }

        return index;
    }

    private static bool StartsWithFlag(string input, int startIndex, string flag)
    {
        if (startIndex + flag.Length > input.Length)
        {
            return false;
        }

        return string.Compare(
            input,
            startIndex,
            flag,
            0,
            flag.Length,
            StringComparison.OrdinalIgnoreCase) == 0;
    }

    private static int FindClosingQuote(string text, int fromIndex)
    {
        for (var i = fromIndex; i < text.Length; i++)
        {
            if (text[i] != '"')
            {
                continue;
            }

            if (i + 1 < text.Length && text[i + 1] == '"')
            {
                i++;
                continue;
            }

            return i;
        }

        return -1;
    }

    private static int IndexOfIgnoreCase(string text, string value, int startIndex)
    {
        return text.IndexOf(value, startIndex, StringComparison.OrdinalIgnoreCase);
    }

    private readonly record struct ExtractedPath(string Path, int EndIndex);

    private readonly record struct TokenExtract(string Token, int EndIndex);

    private readonly record struct ParsedArguments(
        List<string> ScenarioPaths,
        List<string> OtherTokens);
}
