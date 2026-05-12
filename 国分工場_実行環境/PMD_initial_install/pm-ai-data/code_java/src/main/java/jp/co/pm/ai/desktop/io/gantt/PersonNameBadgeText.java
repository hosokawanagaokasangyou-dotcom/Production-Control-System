package jp.co.pm.ai.desktop.io.gantt;

import java.text.Normalizer;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.regex.Pattern;

/**
 * 設備ガント担当バッジ用：氏名から姓を解釈し、表示は姓の先頭最大2コードポイント。
 *
 * <p>Python {@code _split_person_sei_mei} / {@code _gantt_member_label_surname_only} /
 * {@code _normalize_sei_for_match} に概ね整合する。
 */
public final class PersonNameBadgeText {

    /** 同一スロット内の複数バッジを連結するときの区切り（描画側で分割）。 */
    public static final String UNIT_SEPARATOR = "\u001f";

    private static final Pattern TRAILING_HONORIFIC =
            Pattern.compile("(さん|様|殿)$");

    private PersonNameBadgeText() {}

    /**
     * 姓のみ（正規化済み・フル）。Python の {@code _gantt_member_label_surname_only} 相当。
     */
    public static String surnameLabelOnly(String raw) {
        String[] sm = splitSeiMei(raw);
        String sei = sm[0];
        if (sei.isEmpty()) {
            return "";
        }
        String n = normalizeSeiForMatch(sei);
        return n.isEmpty() ? sei : n;
    }

    /**
     * バッジに載せる文字列（姓の先頭最大2コードポイント）。サロゲートペア安全。
     */
    public static String badgeTwoFromRawName(String raw) {
        String seiFull = surnameLabelOnly(raw);
        return firstNCodePoints(seiFull, 2);
    }

    /**
     * イベント {@code op} / {@code sub} から、重複除去済みのバッジ文字列（各2コードポイントまで）を出現順で返す。
     *
     * @param subSplitStartup {@code true} のとき {@code sub} は {@code [,、]} で分割（始業系）。{@code false} のとき {@code
     *     [,、」]} で分割（加工セグメント系）。
     */
    public static List<String> badgeListFromOpSub(String op, String sub, boolean subSplitStartup) {
        List<String> rawOrdered = orderedRawNames(op, sub, subSplitStartup);
        LinkedHashSet<String> seenLabs = new LinkedHashSet<>();
        List<String> out = new ArrayList<>();
        for (String raw : rawOrdered) {
            String lab = surnameLabelOnly(raw);
            if (lab.isEmpty()) {
                continue;
            }
            if (!seenLabs.add(lab)) {
                continue;
            }
            String badge = firstNCodePoints(lab, 2);
            if (!badge.isEmpty()) {
                out.add(badge);
            }
        }
        return List.copyOf(out);
    }

    static String joinBadgeCells(List<String> parts) {
        if (parts == null || parts.isEmpty()) {
            return "";
        }
        return String.join(UNIT_SEPARATOR, parts);
    }

    /** 連続スロット区間で最初の非空バッジセルを返す（同一ランの代表）。 */
    public static String firstNonEmptyInSlotRange(List<String> badgeSlots, int fromSlot, int toSlot) {
        if (badgeSlots == null || badgeSlots.isEmpty()) {
            return "";
        }
        int to = Math.min(toSlot, badgeSlots.size() - 1);
        for (int i = fromSlot; i <= to; i++) {
            String s = badgeSlots.get(i);
            if (s != null && !s.isBlank()) {
                return s.strip();
            }
        }
        return "";
    }

    public static List<String> splitBadgeCell(String cell) {
        if (cell == null || cell.isEmpty()) {
            return List.of();
        }
        String[] a = cell.split(Pattern.quote(UNIT_SEPARATOR), -1);
        List<String> out = new ArrayList<>();
        for (String s : a) {
            if (s != null && !s.isEmpty()) {
                out.add(s);
            }
        }
        return out;
    }

    /** Python {@code _split_person_sei_mei} 相当：{@code [0]} 姓、{@code [1]} 名。 */
    static String[] splitSeiMei(String s) {
        if (s == null) {
            return new String[] {"", ""};
        }
        String t = Normalizer.normalize(s.strip(), Normalizer.Form.NFKC);
        if (t.isEmpty()) {
            return new String[] {"", ""};
        }
        String lower = t.toLowerCase(Locale.ROOT);
        if ("nan".equals(lower) || "none".equals(lower) || "null".equals(lower)) {
            return new String[] {"", ""};
        }
        t = TRAILING_HONORIFIC.matcher(t).replaceFirst("");
        for (int i = 0; i < t.length(); ) {
            int cp = t.codePointAt(i);
            if (cp == ' ' || cp == '\u3000') {
                String sei = t.substring(0, i).strip();
                String rest = t.substring(i + Character.charCount(cp));
                String mei = rest.strip().replaceAll("[\\s　]+", "");
                return new String[] {sei, mei};
            }
            i += Character.charCount(cp);
        }
        return new String[] {t.strip(), ""};
    }

    static String normalizeSeiForMatch(String sei) {
        if (sei == null || sei.isEmpty()) {
            return "";
        }
        String t = Normalizer.normalize(sei.strip(), Normalizer.Form.NFKC);
        if (t.contains("富田")) {
            t = t.replace("富田", "冨田");
        }
        return t.replaceAll("[\\s　]+", "");
    }

    private static List<String> orderedRawNames(String op, String sub, boolean subSplitStartup) {
        List<String> rawNames = new ArrayList<>();
        LinkedHashSet<String> seenRaw = new LinkedHashSet<>();
        String opN = normalizeSpaces(op);
        if (!opN.isEmpty() && seenRaw.add(opN)) {
            rawNames.add(opN);
        }
        String subRaw = sub != null ? sub.strip() : "";
        if (subRaw.isEmpty()) {
            return rawNames;
        }
        Pattern splitPat = subSplitStartup ? Pattern.compile("[,、]") : Pattern.compile("[,、」]");
        for (String seg : splitPat.split(subRaw)) {
            String t = seg.strip();
            if (!t.isEmpty() && seenRaw.add(t)) {
                rawNames.add(t);
            }
        }
        return rawNames;
    }

    private static String normalizeSpaces(String s) {
        if (s == null) {
            return "";
        }
        return String.join(" ", s.strip().split("\\s+"));
    }

    /** テスト・呼び出し元からも利用可。 */
    public static String firstNCodePoints(String s, int maxCp) {
        if (s == null || s.isEmpty() || maxCp <= 0) {
            return "";
        }
        int n = 0;
        int i = 0;
        while (i < s.length() && n < maxCp) {
            int cp = s.codePointAt(i);
            i += Character.charCount(cp);
            n++;
        }
        return s.substring(0, i);
    }
}
