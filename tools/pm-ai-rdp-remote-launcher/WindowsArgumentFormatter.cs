using System.Text;

namespace PmAi.RdpRemoteLauncher;

/// <summary>
/// Windows プロセス起動向けに、空白を含む引数トークンへ {@code "..."} を付与する。
/// </summary>
internal static class WindowsArgumentFormatter
{
    internal static List<string> TokenizeForProcess(string? arguments)
    {
        if (string.IsNullOrWhiteSpace(arguments))
        {
            return new List<string>();
        }

        var trimmed = arguments.Trim();
        if (!trimmed.StartsWith('"') && LooksLikeSinglePathWithSpaces(trimmed))
        {
            return new List<string> { trimmed };
        }

        return Tokenize(trimmed);
    }

    internal static string FormatArgumentString(string? arguments)
    {
        if (string.IsNullOrWhiteSpace(arguments))
        {
            return string.Empty;
        }

        var trimmed = arguments.Trim();
        if (!trimmed.StartsWith('"') && LooksLikeSinglePathWithSpaces(trimmed))
        {
            return QuoteIfNeeded(trimmed);
        }

        var tokens = Tokenize(trimmed);
        if (tokens.Count == 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        foreach (var token in tokens)
        {
            if (string.IsNullOrEmpty(token))
            {
                continue;
            }

            if (builder.Length > 0)
            {
                builder.Append(' ');
            }

            builder.Append(QuoteIfNeeded(token));
        }

        return builder.ToString();
    }

    internal static List<string> Tokenize(string arguments)
    {
        var tokens = new List<string>();
        if (string.IsNullOrEmpty(arguments))
        {
            return tokens;
        }

        var current = new StringBuilder();
        var inQuotes = false;
        for (var i = 0; i < arguments.Length; i++)
        {
            var c = arguments[i];
            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < arguments.Length && arguments[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    current.Append(c);
                }
            }
            else if (c == '"')
            {
                inQuotes = true;
            }
            else if (char.IsWhiteSpace(c))
            {
                if (current.Length > 0)
                {
                    tokens.Add(current.ToString());
                    current.Clear();
                }
            }
            else
            {
                current.Append(c);
            }
        }

        if (current.Length > 0)
        {
            tokens.Add(current.ToString());
        }

        return tokens;
    }

    private static string QuoteIfNeeded(string token)
    {
        if (token.IndexOf(' ') >= 0 || token.IndexOf('\t') >= 0)
        {
            return "\"" + token.Replace("\"", "\"\"") + "\"";
        }

        return token;
    }

    private static bool LooksLikeSinglePathWithSpaces(string value)
    {
        if (value.IndexOf(' ') < 0)
        {
            return false;
        }

        if (value.StartsWith(@"\\"))
        {
            return true;
        }

        return value.Length >= 2
               && value[1] == ':'
               && char.IsLetter(value[0]);
    }
}
