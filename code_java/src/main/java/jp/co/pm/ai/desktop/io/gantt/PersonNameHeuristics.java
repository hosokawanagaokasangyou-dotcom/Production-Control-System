package jp.co.pm.ai.desktop.io.gantt;

import java.util.regex.Pattern;

/** 担当者名らしさの簡易判定（計画表の機械記号・数量文字列の除外）。 */
public final class PersonNameHeuristics {

    private static final Pattern QTY_WITH_UNIT =
            Pattern.compile("^[\\d,.]+\\s*[mMｍ]$");
    private static final Pattern HAS_JAPANESE_NAME_SCRIPT =
            Pattern.compile("[\\p{Script=Han}\\p{Script=Hiragana}\\p{Script=Katakana}ー・]");
    private static final Pattern MACHINE_OR_LOT_CODE =
            Pattern.compile("^[\\[\\(【].*");
    private static final Pattern ASCII_TOKEN_CODE =
            Pattern.compile("^[A-Za-z0-9][A-Za-z0-9\\-_.·/]*$");

    private PersonNameHeuristics() {}

    /** {@code true} のとき担当 OP／バッジ表示に使ってよい人名らしい文字列。 */
    public static boolean looksLikePersonName(String raw) {
        if (raw == null) {
            return false;
        }
        String t = raw.strip();
        if (t.isEmpty()) {
            return false;
        }
        String lower = t.toLowerCase(java.util.Locale.ROOT);
        if ("nan".equals(lower) || "none".equals(lower) || "null".equals(lower)) {
            return false;
        }
        if (MACHINE_OR_LOT_CODE.matcher(t).matches()) {
            return false;
        }
        if (looksLikeNumericQty(t) || QTY_WITH_UNIT.matcher(t).matches()) {
            return false;
        }
        if (HAS_JAPANESE_NAME_SCRIPT.matcher(t).find()) {
            return true;
        }
        if (ASCII_TOKEN_CODE.matcher(t).matches()) {
            return false;
        }
        return false;
    }

    private static boolean looksLikeNumericQty(String raw) {
        if (raw == null || raw.isBlank()) {
            return false;
        }
        try {
            Double.parseDouble(raw.strip().replace(",", ""));
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }
}
