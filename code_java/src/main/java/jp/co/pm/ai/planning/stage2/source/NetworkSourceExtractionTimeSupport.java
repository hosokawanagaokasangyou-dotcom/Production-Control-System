package jp.co.pm.ai.planning.stage2.source;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** ネットワークソースファイル名・メタ行・mtime から取得時刻を解決する。 */
public final class NetworkSourceExtractionTimeSupport {

    /** 計画・日報の取得時刻差がこの分数を超えると UI 警告（ブロックしない）。 */
    public static final int PAIR_DELTA_WARN_MINUTES = 60;

    private static final Pattern FILENAME_TIMESTAMP =
            Pattern.compile("_(\\d{8})_(\\d{6})(?:\\.[^.]+)?$", Pattern.CASE_INSENSITIVE);

    private static final DateTimeFormatter[] META_DATE_TIME_FORMATS =
            new DateTimeFormatter[] {
                DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"),
                DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"),
                DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm"),
                DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm"),
            };

    private NetworkSourceExtractionTimeSupport() {}

    public record ResolvedExtractionTime(
            LocalDateTime dateTime, SourceKind sourceKind, String detail) {}

    public enum SourceKind {
        SHEET_COLUMN,
        FILENAME,
        CSV_META,
        FILE_MTIME
    }

    public static Optional<LocalDateTime> parseFilenameTimestamp(String fileName) {
        if (fileName == null || fileName.isBlank()) {
            return Optional.empty();
        }
        Matcher m = FILENAME_TIMESTAMP.matcher(fileName.strip());
        if (!m.find()) {
            return Optional.empty();
        }
        try {
            LocalDate d = LocalDate.parse(m.group(1), DateTimeFormatter.BASIC_ISO_DATE);
            LocalTime t = LocalTime.parse(m.group(2), DateTimeFormatter.ofPattern("HHmmss"));
            return Optional.of(LocalDateTime.of(d, t));
        } catch (DateTimeParseException ex) {
            return Optional.empty();
        }
    }

    public static Optional<LocalDateTime> fromFileMtime(Path path) {
        if (path == null || !Files.isRegularFile(path)) {
            return Optional.empty();
        }
        try {
            FileTime ft = Files.getLastModifiedTime(path);
            Instant instant = ft.toInstant();
            return Optional.of(LocalDateTime.ofInstant(instant, ZoneId.systemDefault()));
        } catch (IOException ex) {
            return Optional.empty();
        }
    }

    /** 加工日報 CSV 先頭メタ行（最大3行）から日時を探す。 */
    public static Optional<LocalDateTime> parseDailyReportCsvMeta(Path csvPath) {
        if (csvPath == null || !Files.isRegularFile(csvPath)) {
            return Optional.empty();
        }
        try {
            List<String> lines = Files.readAllLines(csvPath, StandardCharsets.UTF_8);
            int limit = Math.min(3, lines.size());
            for (int i = 0; i < limit; i++) {
                Optional<LocalDateTime> parsed = parseLooseDateTimeText(lines.get(i));
                if (parsed.isPresent()) {
                    return parsed;
                }
            }
        } catch (IOException ignored) {
            return Optional.empty();
        }
        return Optional.empty();
    }

    public static Optional<LocalDateTime> parseLooseDateTimeText(String text) {
        if (text == null || text.isBlank()) {
            return Optional.empty();
        }
        String s = text.strip();
        Matcher embedded = Pattern.compile("(\\d{4}[/-]\\d{2}[/-]\\d{2}[ T]\\d{2}:\\d{2}(:\\d{2})?)")
                .matcher(s);
        if (embedded.find()) {
            s = embedded.group(1).replace('T', ' ');
        }
        for (DateTimeFormatter fmt : META_DATE_TIME_FORMATS) {
            try {
                return Optional.of(LocalDateTime.parse(s, fmt));
            } catch (DateTimeParseException ignored) {
                // next
            }
        }
        try {
            LocalDate d = LocalDate.parse(s, DateTimeFormatter.ofPattern("yyyy/MM/dd"));
            return Optional.of(d.atStartOfDay());
        } catch (DateTimeParseException ignored) {
            // fall through
        }
        try {
            LocalDate d = LocalDate.parse(s, DateTimeFormatter.ISO_LOCAL_DATE);
            return Optional.of(d.atStartOfDay());
        } catch (DateTimeParseException ignored) {
            return Optional.empty();
        }
    }

    public static long deltaMinutes(LocalDateTime planAt, LocalDateTime dailyAt) {
        if (planAt == null || dailyAt == null) {
            return Long.MAX_VALUE;
        }
        return Math.abs(DurationSafe.betweenMinutes(planAt, dailyAt));
    }

    public static boolean isLargePairDelta(long deltaMinutes) {
        return deltaMinutes > PAIR_DELTA_WARN_MINUTES;
    }

    public static String displayTime(LocalDateTime dt) {
        if (dt == null) {
            return "—";
        }
        return dt.format(DateTimeFormatter.ofPattern("HH:mm", Locale.JAPAN));
    }

    public static String displayDateTime(LocalDateTime dt) {
        if (dt == null) {
            return "—";
        }
        return dt.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm", Locale.JAPAN));
    }

    /** {@link java.time.Duration#toMinutes()} の符号付き差分の絶対値（分）。 */
    private static final class DurationSafe {
        private DurationSafe() {}

        static long betweenMinutes(LocalDateTime a, LocalDateTime b) {
            return java.time.Duration.between(a, b).toMinutes();
        }
    }
}
