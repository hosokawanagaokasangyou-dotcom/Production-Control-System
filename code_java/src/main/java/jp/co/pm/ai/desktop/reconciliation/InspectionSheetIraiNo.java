package jp.co.pm.ai.desktop.reconciliation;

import java.util.Locale;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** 後加工検査表の依頼NO正規化・ファイル名／セルからの抽出。 */
public final class InspectionSheetIraiNo {

    private static final Pattern TPI = Pattern.compile("(?i)TPI\\s*\\d+-\\d+");
    private static final Pattern STANDARD = Pattern.compile("[A-Za-z]{1,4}\\d{1,2}-\\d{1,3}");

    private InspectionSheetIraiNo() {}

    public static Optional<String> extractFromFileName(String fileName) {
        return extractFromCellText(fileName);
    }

    public static Optional<String> extractFromCellText(String text) {
        if (text == null || text.isBlank()) {
            return Optional.empty();
        }
        String half = toHalfWidth(text);
        Matcher tpi = TPI.matcher(half);
        if (tpi.find()) {
            return Optional.of(canonicalTpi(tpi.group()));
        }
        Matcher std = STANDARD.matcher(half);
        if (std.find()) {
            return Optional.of(std.group().toUpperCase(Locale.ROOT));
        }
        return Optional.empty();
    }

    public static String normalize(String raw) {
        if (raw == null) {
            return "";
        }
        return toHalfWidth(raw).strip().replaceAll("\\s+", "").toUpperCase(Locale.ROOT);
    }

    public static boolean matches(String left, String right) {
        String a = normalize(left);
        String b = normalize(right);
        return !a.isEmpty() && a.equals(b);
    }

    static String toHalfWidth(String raw) {
        StringBuilder sb = new StringBuilder(raw.length());
        for (int i = 0; i < raw.length(); i++) {
            char c = raw.charAt(i);
            if (c >= 'Ａ' && c <= 'Ｚ') {
                c = (char) (c - 'Ａ' + 'A');
            } else if (c >= 'ａ' && c <= 'ｚ') {
                c = (char) (c - 'ａ' + 'a');
            } else if (c >= '０' && c <= '９') {
                c = (char) (c - '０' + '0');
            } else if (c == '－' || c == '―' || c == '‐' || c == 'ー') {
                c = '-';
            }
            sb.append(c);
        }
        return sb.toString();
    }

    private static String canonicalTpi(String matched) {
        String upper = matched.toUpperCase(Locale.ROOT).replaceAll("\\s+", "");
        return upper.replaceFirst("^TPI", "TPI ");
    }
}
