package jp.co.pm.ai.desktop.reconciliation;

import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.regex.Pattern;

/**
 * 依頼書原本 ↔ 受注ファイル転記照合の値正規化。
 * {@link ReconciliationApp} の private 正規化と同等。
 */
public final class JuchuTransferValueNormalizer {

    private static final Pattern NUMERIC_PATTERN = Pattern.compile("[-+]?\\d*\\.\\d+|\\d+");

    private static final List<String> DATE_FORMATS =
            Arrays.asList(
                    "yyyy-MM-dd HH:mm:ss",
                    "yyyy/MM/dd HH:mm:ss",
                    "yyyy-MM-dd",
                    "yyyy/MM/dd");

    private JuchuTransferValueNormalizer() {}

    public static String normalizeKey(String val) {
        if (val == null) {
            return "";
        }
        String text = val.strip().toUpperCase(Locale.ROOT);
        text = java.text.Normalizer.normalize(text, java.text.Normalizer.Form.NFKC);
        text = text.replaceAll("\\s+", "");
        text = normalizeDashVariants(text);
        return text;
    }

    public static String normalizeText(String val) {
        if (val == null) {
            return "";
        }
        String text = val.strip();
        text = java.text.Normalizer.normalize(text, java.text.Normalizer.Form.NFKC);
        text = text.replaceAll("\\s+", "");
        text = normalizeDashVariants(text);
        return text.toUpperCase(Locale.ROOT);
    }

    public static double normalizeNumeric(String val) {
        if (val == null || val.isEmpty()) {
            return 0.0;
        }
        var m = NUMERIC_PATTERN.matcher(val.strip());
        if (m.find()) {
            return Double.parseDouble(m.group());
        }
        return 0.0;
    }

    public static String normalizeDateVal(String val) {
        if (val == null) {
            return "";
        }
        String text = val.strip();
        for (String fmt : DATE_FORMATS) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(fmt);
                Date d = sdf.parse(text);
                SimpleDateFormat outSdf = new SimpleDateFormat("yyyy-MM-dd");
                return outSdf.format(d);
            } catch (Exception ignored) {
                // try next format
            }
        }
        return normalizeText(val);
    }

    public static boolean isBlank(String val) {
        return val == null || val.strip().isEmpty();
    }

    /**
     * 受注・原本の日付文字列を {@link LocalDate} に変換。解釈不能時は {@code null}。
     */
    public static LocalDate parseLocalDate(String val) {
        if (val == null || val.isBlank()) {
            return null;
        }
        String normalized = normalizeDateVal(val);
        if (normalized.isBlank()) {
            return null;
        }
        try {
            return LocalDate.parse(normalized, DateTimeFormatter.ISO_LOCAL_DATE);
        } catch (DateTimeParseException ex) {
            // fall through
        }
        for (String fmt : DATE_FORMATS) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(fmt);
                Date d = sdf.parse(val.strip());
                return d.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            } catch (Exception ignored) {
                // try next format
            }
        }
        return null;
    }

    private static String normalizeDashVariants(String text) {
        return text.replace("－", "-")
                .replace("ー", "-")
                .replace("―", "-")
                .replace("‐", "-");
    }
}
