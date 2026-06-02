package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Optional;
import java.util.Set;

import jp.co.pm.ai.desktop.io.ExcelCellReadSupport;

/**
 * 配台計画_タスク入力タブの日付系列判定とセル文字列の parse / format。
 */
public final class PlanInputDateColumnSupport {

    private static final Set<String> EXACT_DATE_COLUMNS =
            Set.of(
                    "回答納期",
                    "指定納期",
                    "原反投入日",
                    "計画基準納期",
                    "加工開始日",
                    "受注日",
                    "データ抽出日");

    private static final DateTimeFormatter[] PARSE_FORMATTERS =
            new DateTimeFormatter[] {
                DateTimeFormatter.ofPattern("yyyy/M/d"),
                DateTimeFormatter.ofPattern("yyyy/MM/dd"),
                DateTimeFormatter.ofPattern("yyyy-M-d"),
                DateTimeFormatter.ofPattern("yyyy-MM-dd"),
            };

    /** Excel / 計画シート表示に合わせた出力（{@code 2026/6/15} 形式）。 */
    private static final DateTimeFormatter OUTPUT_FORMAT = DateTimeFormatter.ofPattern("yyyy/M/d");

    /** 段階2.0 配台開始下限に使う日時列（時刻まで保持）。 */
    private static final Set<String> EXACT_DATETIME_COLUMNS = Set.of("配台可能日時");

    private static final DateTimeFormatter[] DATETIME_PARSE_FORMATTERS =
            new DateTimeFormatter[] {
                DateTimeFormatter.ofPattern("yyyy/M/d H:mm"),
                DateTimeFormatter.ofPattern("yyyy/M/d HH:mm"),
                DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm"),
                DateTimeFormatter.ofPattern("yyyy-M-d H:mm"),
                DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm"),
            };

    /** 配台可能日時の出力（{@code 2026/6/15 12:45} 形式）。 */
    private static final DateTimeFormatter DATETIME_OUTPUT_FORMAT =
            DateTimeFormatter.ofPattern("yyyy/M/d H:mm");

    private PlanInputDateColumnSupport() {}

    /**
     * ダブルクリックでカレンダー編集する日付列か。
     *
     * <p>（元）… や …_試行前 など参照専用列は除外する。
     */
    public static boolean isEditableDateColumn(String columnTitle) {
        if (columnTitle == null) {
            return false;
        }
        String h = columnTitle.strip();
        if (h.isEmpty()) {
            return false;
        }
        if (h.startsWith("(元)") || h.startsWith("（元）")) {
            return false;
        }
        if (h.contains("試行前")) {
            return false;
        }
        if (EXACT_DATE_COLUMNS.contains(h)) {
            return true;
        }
        return isDateColumnBase(h);
    }

    private static boolean isDateColumnBase(String h) {
        if (h == null || h.isBlank()) {
            return false;
        }
        return TabularCellHighlight.isStage1DateColumnHeader(h);
    }

    /** セル文字列を {@link LocalDate} に解釈する（空・解釈不能は empty）。 */
    public static Optional<LocalDate> parseCellValue(String raw) {
        if (raw == null) {
            return Optional.empty();
        }
        String t = raw.strip();
        if (t.isEmpty()) {
            return Optional.empty();
        }
        int space = t.indexOf(' ');
        if (space > 0) {
            String timePart = t.substring(space + 1).strip();
            if (timePart.startsWith("00:00")) {
                t = t.substring(0, space).strip();
            }
        }
        t = ExcelCellReadSupport.stripMidnightDateTimeSuffix(t).strip();
        if (t.isEmpty()) {
            return Optional.empty();
        }
        for (DateTimeFormatter fmt : PARSE_FORMATTERS) {
            try {
                return Optional.of(LocalDate.parse(t, fmt));
            } catch (DateTimeParseException ignored) {
            }
        }
        return Optional.empty();
    }

    /** カレンダー選択結果を計画シート向け文字列へ。 */
    public static String formatCellValue(LocalDate date) {
        if (date == null) {
            return "";
        }
        return OUTPUT_FORMAT.format(date);
    }

    /**
     * 日時として編集する列か（配台可能日時）。
     *
     * <p>（元）… 参照列は除外する。日付列（{@link #isEditableDateColumn}）とは排他。
     */
    public static boolean isEditableDateTimeColumn(String columnTitle) {
        if (columnTitle == null) {
            return false;
        }
        String h = columnTitle.strip();
        if (h.isEmpty()) {
            return false;
        }
        if (h.startsWith("(元)") || h.startsWith("（元）")) {
            return false;
        }
        return EXACT_DATETIME_COLUMNS.contains(h);
    }

    /** セル文字列を {@link LocalDateTime} に解釈する（時刻欠落時は {@code 12:45} 補完）。 */
    public static Optional<LocalDateTime> parseDateTimeCellValue(String raw) {
        if (raw == null) {
            return Optional.empty();
        }
        String t = raw.strip();
        if (t.isEmpty()) {
            return Optional.empty();
        }
        for (DateTimeFormatter fmt : DATETIME_PARSE_FORMATTERS) {
            try {
                return Optional.of(LocalDateTime.parse(t, fmt));
            } catch (DateTimeParseException ignored) {
            }
        }
        // 時刻が無い日付だけの入力は配台開始既定時刻 12:45 を補う。
        return parseCellValue(t).map(d -> d.atTime(LocalTime.of(12, 45)));
    }

    /** 編集結果を計画シート向け文字列へ（{@code 2026/6/15 12:45} 形式）。 */
    public static String formatDateTimeCellValue(LocalDateTime dateTime) {
        if (dateTime == null) {
            return "";
        }
        return DATETIME_OUTPUT_FORMAT.format(dateTime);
    }
}
