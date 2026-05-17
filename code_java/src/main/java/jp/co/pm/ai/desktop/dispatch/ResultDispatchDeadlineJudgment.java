package jp.co.pm.ai.desktop.dispatch;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Locale;
import java.util.Map;

/**
 * 納期管理ビュー「配台結果」タブ向けの納期判定（OK / NG）。
 *
 * <p>判定規則は Python {@code _result_task_plan_end_within_answer_or_spec_16_label} と整合する。
 * 計画終了は {@code 加工終了日時}（無ければ {@code 加工完了日} の暦日開始）を用い、納期日は
 * {@code 回答納期} を優先し空なら {@code 指定納期} とする。
 */
public final class ResultDispatchDeadlineJudgment {

    public static final String COL_TITLE = "納期判定";

    public static final String OK = "OK";
    public static final String NG = "NG";

    private static final LocalTime DUE_DAY_COMPLETION_TIME = LocalTime.of(16, 0);

    private static final DateTimeFormatter DISPATCH_SLASH_DT =
            DateTimeFormatter.ofPattern("yyyy/M/d H:mm", Locale.JAPAN);

    private static final DateTimeFormatter ISO_DT =
            DateTimeFormatter.ofPattern("yyyy-M-d H:mm", Locale.ROOT);

    private ResultDispatchDeadlineJudgment() {}

    /**
     * @return {@link #OK} / {@link #NG}。計画終了または納期日が取れないときは空文字。
     */
    public static String judgmentOkNg(Map<String, String> row) {
        if (row == null) {
            return "";
        }
        LocalDateTime planEnd = resolvePlanEnd(row);
        if (planEnd == null) {
            return "";
        }
        LocalDate dueDay = resolveDueDay(row);
        if (dueDay == null) {
            return "";
        }
        if (isVPrefixTaskId(row.get("依頼NO"))) {
            LocalDateTime deadline = LocalDateTime.of(dueDay, DUE_DAY_COMPLETION_TIME);
            return planEnd.isAfter(deadline) ? NG : OK;
        }
        LocalDateTime startOfDue = dueDay.atStartOfDay();
        return planEnd.isBefore(startOfDue) ? OK : NG;
    }

    static LocalDate resolveDueDay(Map<String, String> row) {
        LocalDate answer = ResultDispatchPivot.parseIsoDate(nz(row.get("回答納期")));
        if (answer != null) {
            return answer;
        }
        return ResultDispatchPivot.parseIsoDate(nz(row.get("指定納期")));
    }

    static LocalDateTime resolvePlanEnd(Map<String, String> row) {
        LocalDateTime dt = parseDispatchDateTime(nz(row.get("加工終了日時")));
        if (dt != null) {
            return dt;
        }
        LocalDate d = ResultDispatchPivot.parseIsoDate(nz(row.get("加工完了日")));
        return d != null ? d.atStartOfDay() : null;
    }

    static LocalDateTime parseDispatchDateTime(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String t = raw.strip();
        int space = t.indexOf(' ');
        if (space > 0 && space + 1 < t.length()) {
            String datePart = t.substring(0, space);
            String timePart = t.substring(space + 1).trim();
            LocalDate d = ResultDispatchPivot.parseIsoDate(datePart);
            if (d == null) {
                return null;
            }
            try {
                if (timePart.length() >= 4 && timePart.charAt(2) == ':') {
                    int h = Integer.parseInt(timePart.substring(0, 2));
                    int m = Integer.parseInt(timePart.substring(3, 5));
                    return LocalDateTime.of(d, LocalTime.of(h, m));
                }
            } catch (NumberFormatException ignored) {
                return null;
            }
        }
        if (t.length() >= 16 && t.charAt(10) == ' ') {
            try {
                return LocalDateTime.parse(t.substring(0, 16), DISPATCH_SLASH_DT);
            } catch (DateTimeParseException ignored) {
                // fall through
            }
            try {
                return LocalDateTime.parse(t.substring(0, 16), ISO_DT);
            } catch (DateTimeParseException ignored) {
                // fall through
            }
        }
        LocalDate only = ResultDispatchPivot.parseIsoDate(t);
        return only != null ? only.atStartOfDay() : null;
    }

    static boolean isVPrefixTaskId(String taskId) {
        if (taskId == null) {
            return false;
        }
        String t = taskId.strip();
        return !t.isEmpty() && t.toUpperCase(Locale.ROOT).startsWith("V");
    }

    private static String nz(String v) {
        return v != null ? v : "";
    }
}
