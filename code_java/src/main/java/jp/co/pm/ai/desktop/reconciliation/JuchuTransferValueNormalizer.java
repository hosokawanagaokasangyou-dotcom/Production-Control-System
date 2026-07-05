package jp.co.pm.ai.desktop.reconciliation;

import java.text.SimpleDateFormat;
import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoUnit;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 依頼書原本 ↔ 受注ファイル転記照合の値正規化。
 * {@link ReconciliationApp} の private 正規化と同等。
 */
public final class JuchuTransferValueNormalizer {

    private static final Pattern NUMERIC_PATTERN = Pattern.compile("[-+]?\\d*\\.\\d+|\\d+");

    /** 原本の短い日付（例: {@code 7/15}、{@code 6/10（水）}）。年省略時は参照日から補完。 */
    private static final Pattern MONTH_DAY =
            Pattern.compile(
                    "^(\\d{1,2})[/／](\\d{1,2})(?:[/／](\\d{2,4}))?(?:[（(][^）)]*[）)])?$");

    /** Excel 等の {@code 7月7日}、{@code 6月10日（水）}。年省略時は参照日から補完。 */
    private static final Pattern JAPANESE_MONTH_DAY =
            Pattern.compile(
                    "^(\\d{1,2})月(\\d{1,2})日(?:[（(][^）)]*[）)])?$");

    /** {@code 2026年7月7日} 形式。 */
    private static final Pattern JAPANESE_FULL_DATE =
            Pattern.compile(
                    "^(\\d{4})年(\\d{1,2})月(\\d{1,2})日(?:[（(][^）)]*[）)])?$");

    /** 年省略 M/D の年推定で参照日から許容する最大日数差（約6か月）。 */
    private static final long SHORT_DATE_YEAR_INFERENCE_MAX_DAYS = 183L;

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
        LocalDate parsed = parseLocalDate(val);
        if (parsed != null) {
            return parsed.format(DateTimeFormatter.ISO_LOCAL_DATE);
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
        return parseLocalDate(val, LocalDate.now());
    }

    /**
     * 年省略の M/D を解釈するときの参照日（通常は受注側の完全日付または当日）。
     */
    public static LocalDate parseLocalDate(String val, LocalDate yearReference) {
        if (val == null || val.isBlank()) {
            return null;
        }
        LocalDate ref = yearReference != null ? yearReference : LocalDate.now();
        String text = val.strip();
        for (String fmt : DATE_FORMATS) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(fmt);
                Date d = sdf.parse(text);
                return d.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            } catch (Exception ignored) {
                // try next format
            }
        }
        try {
            return LocalDate.parse(text.split("\\s+", 2)[0], DateTimeFormatter.ISO_LOCAL_DATE);
        } catch (DateTimeParseException ex) {
            // fall through
        }
        for (String fmt : List.of("yyyy/MM/dd", "yyyyMMdd")) {
            try {
                String dateOnly = text.split("\\s+", 2)[0];
                return LocalDate.parse(dateOnly, DateTimeFormatter.ofPattern(fmt));
            } catch (Exception ignored) {
                // try next format
            }
        }
        String dateLine = text.split("\\n", 2)[0].strip();
        LocalDate jpFull = parseJapaneseFullDate(dateLine);
        if (jpFull != null) {
            return jpFull;
        }
        return parseMonthDayValue(dateLine, ref);
    }

    private static LocalDate parseJapaneseFullDate(String val) {
        if (val == null || val.isBlank()) {
            return null;
        }
        String text =
                java.text.Normalizer.normalize(val.strip(), java.text.Normalizer.Form.NFKC);
        Matcher m = JAPANESE_FULL_DATE.matcher(text);
        if (!m.matches()) {
            return null;
        }
        return safeLocalDate(
                Integer.parseInt(m.group(1)),
                Integer.parseInt(m.group(2)),
                Integer.parseInt(m.group(3)));
    }

    static LocalDate parseMonthDayValue(String val, LocalDate yearReference) {
        if (val == null || val.isBlank()) {
            return null;
        }
        String text =
                java.text.Normalizer.normalize(val.strip(), java.text.Normalizer.Form.NFKC);
        Matcher slash = MONTH_DAY.matcher(text);
        if (slash.matches()) {
            return resolveMonthDay(
                    Integer.parseInt(slash.group(1)),
                    Integer.parseInt(slash.group(2)),
                    slash.group(3),
                    yearReference);
        }
        Matcher jp = JAPANESE_MONTH_DAY.matcher(text);
        if (jp.matches()) {
            return resolveMonthDay(
                    Integer.parseInt(jp.group(1)),
                    Integer.parseInt(jp.group(2)),
                    null,
                    yearReference);
        }
        return null;
    }

    private static LocalDate resolveMonthDay(
            int month, int day, String yearGroup, LocalDate yearReference) {
        if (yearGroup != null && !yearGroup.isBlank()) {
            int year =
                    yearGroup.length() <= 2
                            ? 2000 + Integer.parseInt(yearGroup)
                            : Integer.parseInt(yearGroup);
            return safeLocalDate(year, month, day);
        }
        LocalDate ref = yearReference != null ? yearReference : LocalDate.now();
        LocalDate candidate = safeLocalDate(ref.getYear(), month, day);
        if (candidate == null) {
            return null;
        }
        long daysFromRef = ChronoUnit.DAYS.between(ref, candidate);
        if (daysFromRef > SHORT_DATE_YEAR_INFERENCE_MAX_DAYS) {
            LocalDate priorYear = safeLocalDate(ref.getYear() - 1, month, day);
            if (priorYear != null) {
                return priorYear;
            }
        } else if (daysFromRef < -SHORT_DATE_YEAR_INFERENCE_MAX_DAYS) {
            LocalDate nextYear = safeLocalDate(ref.getYear() + 1, month, day);
            if (nextYear != null) {
                return nextYear;
            }
        }
        return candidate;
    }

    private static LocalDate safeLocalDate(int year, int month, int day) {
        try {
            return LocalDate.of(year, month, day);
        } catch (DateTimeException ex) {
            return null;
        }
    }

    private static String normalizeDashVariants(String text) {
        return text.replace("－", "-")
                .replace("ー", "-")
                .replace("―", "-")
                .replace("‐", "-");
    }
}
