namespace PmAi.RdpRemoteLauncher;

internal readonly record struct ParsedCommand(string Executable, string Arguments);

internal static class CommandLineParser
{
    internal static ParsedCommand Parse(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            throw new FormatException("コマンド行が空です。");
        }

        var trimmed = line.Trim();
        if (trimmed.Length == 0)
        {
            throw new FormatException("コマンド行が空です。");
        }

        if (trimmed[0] == '"')
        {
            var end = FindClosingQuote(trimmed, 1);
            if (end < 0)
            {
                throw new FormatException("引用符が閉じられていません: " + line);
            }

            var executable = trimmed[1..end].Replace("\"\"", "\"");
            var arguments = end + 1 < trimmed.Length ? trimmed[(end + 1)..].TrimStart() : string.Empty;
            return new ParsedCommand(executable, arguments);
        }

        var space = trimmed.IndexOf(' ');
        if (space < 0)
        {
            return new ParsedCommand(trimmed, string.Empty);
        }

        return new ParsedCommand(trimmed[..space], trimmed[(space + 1)..].Trim());
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
}
