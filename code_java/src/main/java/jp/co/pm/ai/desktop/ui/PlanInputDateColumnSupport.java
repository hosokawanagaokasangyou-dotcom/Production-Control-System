package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
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
                    "原反投入日_上書き",
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
        if (h.endsWith("_上書き")) {
            String base = h.substring(0, h.length() - "_上書き".length());
            return isDateColumnBase(base);
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
}
