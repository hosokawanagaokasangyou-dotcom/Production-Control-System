using System.Text;
using System.Text.RegularExpressions;

namespace PmAi.RdpRemoteLauncher;

/// <summary>
/// UNC パス内の既知セグメント名を実フォルダ表記へ揃える。
/// </summary>
internal static class UncPathSegmentRepair
{
    internal const string Konan002KakogSegment = "002  加工G";

    /** {@code \002} 直後の空白 run を実フォルダ名（スペース2つ）へ。パターン先頭に {@code \002} を置かない（8進解釈回避）。 */
    private static readonly Regex Konan002KakogPattern =
        new(@"(\\)002\s+加工G", RegexOptions.CultureInvariant);

    internal static string Repair(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return path ?? string.Empty;
        }

        return Konan002KakogPattern.Replace(
            path,
            m => m.Groups[1].Value + Konan002KakogSegment);
    }
}
