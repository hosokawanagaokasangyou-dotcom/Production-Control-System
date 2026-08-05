package jp.co.pm.ai.desktop.dispatch;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/** 段階3.5 ウィザード向け勤怠プレビュー（Python {@code attendance_overtime_preview.py} の JSON）。 */
public final class AttendanceOvertimePreview {

    private static final ObjectMapper JSON = new ObjectMapper();

    /** 段階3.5 ウィザード: 暦日表示は当日からこの日数後まで（当日を含む）。 */
    public static final int OVERTIME_SIM_DATE_WINDOW_DAYS_AFTER_TODAY = 30;

    private AttendanceOvertimePreview() {}

    public record CellInfo(
            boolean working,
            boolean eligibleForAssignment,
            int overtimeMinutes,
            boolean weekend) {}

    public record Preview(
            boolean ok,
            String error,
            List<String> members,
            List<LocalDate> dates,
            Map<LocalDate, Map<String, CellInfo>> cells) {}

    public static Preview parseJson(String raw) throws Exception {
        String payload = MasterReadSummaryJson.extractLastJsonLine(raw);
        JsonNode root = JSON.readTree(payload);
        boolean ok = root.path("ok").asBoolean(false);
        String error = root.path("error").asText(null);
        List<String> members = new ArrayList<>();
        for (JsonNode n : root.path("members")) {
            String m = n.asText("").trim();
            if (!m.isEmpty()) {
                members.add(m);
            }
        }
        List<LocalDate> dates = new ArrayList<>();
        for (JsonNode n : root.path("dates")) {
            LocalDate d = LocalDate.parse(n.asText());
            dates.add(d);
        }
        Map<LocalDate, Map<String, CellInfo>> cells = new LinkedHashMap<>();
        JsonNode cellsNode = root.path("cells");
        for (LocalDate d : dates) {
            JsonNode day = cellsNode.path(d.toString());
            Map<String, CellInfo> row = new LinkedHashMap<>();
            for (String m : members) {
                JsonNode c = day.path(m);
                row.put(
                        m,
                        new CellInfo(
                                c.path("is_working").asBoolean(false),
                                c.path("eligible_for_assignment").asBoolean(false),
                                c.path("overtime_minutes").asInt(0),
                                c.path("weekend").asBoolean(d.getDayOfWeek().getValue() >= 6)));
            }
            cells.put(d, row);
        }
        return new Preview(ok, error, members, dates, cells);
    }

    /**
     * 段階3.5 ウィザード向け: {@code fromInclusive} 〜 {@code toInclusive} の暦日だけ残す。
     */
    public static Preview limitToDateWindow(
            Preview preview, LocalDate fromInclusive, LocalDate toInclusive) {
        if (preview == null) {
            return new Preview(false, "preview is null", List.of(), List.of(), Map.of());
        }
        if (fromInclusive == null || toInclusive == null || fromInclusive.isAfter(toInclusive)) {
            return preview;
        }
        List<LocalDate> dates = new ArrayList<>();
        for (LocalDate d : preview.dates()) {
            if (!d.isBefore(fromInclusive) && !d.isAfter(toInclusive)) {
                dates.add(d);
            }
        }
        Map<LocalDate, Map<String, CellInfo>> cells = new LinkedHashMap<>();
        for (LocalDate d : dates) {
            Map<String, CellInfo> row = preview.cells().get(d);
            if (row != null) {
                cells.put(d, row);
            }
        }
        return new Preview(preview.ok(), preview.error(), preview.members(), dates, cells);
    }

    /** 当日 〜 当日+{@link #OVERTIME_SIM_DATE_WINDOW_DAYS_AFTER_TODAY}（両端含む）。 */
    public static Preview limitToDefaultOvertimeSimWindow(Preview preview) {
        LocalDate today = LocalDate.now();
        return limitToDateWindow(
                preview, today, today.plusDays(OVERTIME_SIM_DATE_WINDOW_DAYS_AFTER_TODAY));
    }

    public static String formatDateHeader(LocalDate d) {
        String dow =
                switch (d.getDayOfWeek()) {
                    case SATURDAY -> "土";
                    case SUNDAY -> "日";
                    case MONDAY -> "月";
                    case TUESDAY -> "火";
                    case WEDNESDAY -> "水";
                    case THURSDAY -> "木";
                    case FRIDAY -> "金";
                };
        return d.getYear() + "/" + d.getMonthValue() + "/" + d.getDayOfMonth() + "(" + dow + ")";
    }

    /** master_read_summary と同趣旨の stdout から最終 JSON 行を抽出。 */
    public static final class MasterReadSummaryJson {
        private MasterReadSummaryJson() {}

        public static String extractLastJsonLine(String merged) {
            if (merged == null || merged.isBlank()) {
                return "{}";
            }
            String[] lines = merged.split("\n");
            for (int i = lines.length - 1; i >= 0; i--) {
                String t = lines[i].trim();
                if (t.startsWith("{") && t.endsWith("}")) {
                    return t;
                }
            }
            return merged.trim();
        }
    }
}
