package jp.co.pm.ai.desktop.io.gantt;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeSet;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.io.JsonTableIo;

/**
 * {@code *_equipment_gantt_contract.json}（設備ガント描画契約）から、
 * {@link jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane} が期待する
 * 「日付・機械名・工程名・タスク概覝・HH:MM 列…」形式の 1 シート分を組み立てる。
 *
 * <p>Excel シートと完全一致は目指さず、タイムラインイベントからスロットを充填する近似。
 * xlsx 不要でグラフィック表示に使える。
 *
 * <p>{@code sorted_dates} は計画ホライズン全体を含むことがあり、イベントがまだ無い暦日が先頭に並ぶ。
 * その日はタイムラインがすべて空になるため、グラフィック用の暦日は {@code timeline_events} に現れる日付に限定する。
 */
public final class EquipmentGanttContractSheetTableBuilder {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String COL_DATE = "日付";
    private static final String COL_MACH = "機械名";
    private static final String COL_PROC = "工程名";
    private static final String COL_TASK = "タスク概覝";

    private static final int SLOT_MINUTES = 10;
    private static final LocalTime FALLBACK_DAY_START = LocalTime.of(8, 0);
    private static final LocalTime FALLBACK_DAY_END = LocalTime.of(21, 0);

    private EquipmentGanttContractSheetTableBuilder() {}

    public static JsonTableIo.SheetTable buildFromContractPath(Path contractPath) throws IOException {
        JsonNode root = JSON.readTree(Files.readString(contractPath, StandardCharsets.UTF_8));
        JsonNode packed = root.get("kwargs_packed");
        if (packed == null || !packed.isObject()) {
            throw new IOException("契約 JSON に kwargs_packed がありません: " + contractPath);
        }
        JsonNode eventsNode = packed.get("timeline_events");
        JsonNode equipNode = packed.get("equipment_list");
        JsonNode datesNode = packed.get("sorted_dates");
        if (eventsNode == null || !eventsNode.isArray()) {
            throw new IOException("契約に timeline_events がありません");
        }
        if (equipNode == null || !equipNode.isArray() || equipNode.isEmpty()) {
            throw new IOException("契約に equipment_list がありません");
        }
        if (datesNode == null || !datesNode.isArray() || datesNode.isEmpty()) {
            throw new IOException("契約に sorted_dates がありません");
        }

        List<TimelineEvent> events = new ArrayList<>();
        for (JsonNode en : eventsNode) {
            TimelineEvent te = TimelineEvent.from(en);
            if (te != null) {
                events.add(te);
            }
        }

        List<LocalDate> sortedDates = new ArrayList<>();
        for (JsonNode dn : datesNode) {
            Object d = GanttContractValueDecoder.decodeValue(dn);
            LocalDate ld = GanttContractValueDecoder.toLocalDate(d);
            if (ld != null) {
                sortedDates.add(ld);
            }
        }

        List<String> equipmentLines = new ArrayList<>();
        for (JsonNode eq : equipNode) {
            if (eq != null && eq.isTextual()) {
                equipmentLines.add(eq.asText());
            }
        }

        List<LocalTime> slotStarts = computeSlotTimes(events);
        List<String> columns = new ArrayList<>();
        columns.add(COL_DATE);
        columns.add(COL_MACH);
        columns.add(COL_PROC);
        columns.add(COL_TASK);
        for (LocalTime t : slotStarts) {
            columns.add(formatSlotColumn(t));
        }

        List<Map<String, String>> rows = new ArrayList<>();

        List<LocalDate> graphicDays = graphicCalendarDates(events, sortedDates);
        for (LocalDate day : graphicDays) {
            Map<String, String> section = new LinkedHashMap<>();
            for (String col : columns) {
                section.put(col, col.equals(COL_DATE) ? formatSectionBanner(day) : "");
            }
            rows.add(section);

            for (String equipLine : equipmentLines) {
                String[] split = splitEquipmentLine(equipLine);
                String proc = split[0];
                String mach = split[1];
                Map<String, String> row = new LinkedHashMap<>();
                row.put(COL_DATE, "");
                row.put(COL_MACH, mach);
                row.put(COL_PROC, proc);
                row.put(COL_TASK, "—");

                for (LocalTime slotStart : slotStarts) {
                    String col = formatSlotColumn(slotStart);
                    LocalDateTime winStart = LocalDateTime.of(day, slotStart);
                    LocalDateTime winEnd = winStart.plusMinutes(SLOT_MINUTES);
                    String cell = "";
                    for (TimelineEvent ev : events) {
                        if (!eventTouchesCalendarDay(ev, day)) {
                            continue;
                        }
                        if (!equipLine.equals(ev.machine)) {
                            continue;
                        }
                        if (!rangesOverlap(ev.start, ev.end, winStart, winEnd)) {
                            continue;
                        }
                        if (ev.isInBreaks(winStart, winEnd)) {
                            continue;
                        }
                        cell = ev.cellLabel();
                        break;
                    }
                    row.put(col, cell);
                }
                rows.add(row);
            }
        }

        return new JsonTableIo.SheetTable(columns, rows);
    }

    /**
     * タイムラインにイベントが存在する暦日のみ（イベントの {@code date} および {@code start_dt} の日付）。
     * イベントが 1 件も無いときだけフォールバックで {@code sorted_dates} をそのまま使う。
     */
    private static List<LocalDate> graphicCalendarDates(
            List<TimelineEvent> events, List<LocalDate> sortedDatesFallback) {
        TreeSet<LocalDate> days = new TreeSet<>();
        for (TimelineEvent ev : events) {
            if (ev.date != null) {
                days.add(ev.date);
            }
            if (ev.start != null) {
                days.add(ev.start.toLocalDate());
            }
        }
        if (!days.isEmpty()) {
            return new ArrayList<>(days);
        }
        return new ArrayList<>(sortedDatesFallback);
    }

    /** 契約の {@code date} または開始／終了時刻が、その暦日と関係するイベントを対象にする。 */
    private static boolean eventTouchesCalendarDay(TimelineEvent ev, LocalDate day) {
        if (ev.date != null && ev.date.equals(day)) {
            return true;
        }
        if (ev.start != null && ev.start.toLocalDate().equals(day)) {
            return true;
        }
        if (ev.end != null && ev.end.toLocalDate().equals(day)) {
            return true;
        }
        return false;
    }

    /** "工程+機械…" を最初の '+' で分割（無ければ全体を機械名）。 */
    static String[] splitEquipmentLine(String line) {
        if (line == null || line.isEmpty()) {
            return new String[] {"", ""};
        }
        int p = line.indexOf('+');
        if (p < 0) {
            return new String[] {"", line};
        }
        return new String[] {line.substring(0, p).strip(), line.substring(p + 1).strip()};
    }

    static String formatSectionBanner(LocalDate day) {
        return "【"
                + day.getYear()
                + "/"
                + String.format("%02d", day.getMonthValue())
                + "/"
                + String.format("%02d", day.getDayOfMonth())
                + "】";
    }

    static String formatSlotColumn(LocalTime t) {
        return t.getHour() + ":" + String.format("%02d", t.getMinute());
    }

    static boolean rangesOverlap(
            LocalDateTime a0, LocalDateTime a1, LocalDateTime b0, LocalDateTime b1) {
        return a0.isBefore(b1) && a1.isAfter(b0);
    }

    /**
     * イベントの発生時刻を包含するよう 10 分刻みスロット開始時刻を列挙。
     * 無イベント時は 8:00〜21:00（21:00 排他的終端）。
     */
    static List<LocalTime> computeSlotTimes(List<TimelineEvent> events) {
        LocalTime start = FALLBACK_DAY_START;
        LocalTime endCap = FALLBACK_DAY_END;
        for (TimelineEvent ev : events) {
            LocalTime fs = floorToSlot(ev.start.toLocalTime());
            LocalTime ce = ceilToSlot(ev.end.toLocalTime());
            if (fs.isBefore(start)) {
                start = fs;
            }
            if (ce.isAfter(endCap)) {
                endCap = ce;
            }
        }
        List<LocalTime> out = new ArrayList<>();
        LocalTime t = start;
        while (t.isBefore(endCap)) {
            out.add(t);
            t = t.plusMinutes(SLOT_MINUTES);
            if (out.size() > 5000) {
                break;
            }
        }
        return out;
    }

    static LocalTime floorToSlot(LocalTime t) {
        int m = t.getHour() * 60 + t.getMinute();
        m = (m / SLOT_MINUTES) * SLOT_MINUTES;
        return LocalTime.of(m / 60, m % 60);
    }

    static LocalTime ceilToSlot(LocalTime t) {
        int m = t.getHour() * 60 + t.getMinute();
        int rem = m % SLOT_MINUTES;
        if (rem != 0) {
            m += SLOT_MINUTES - rem;
        }
        return LocalTime.of(m / 60, m % 60);
    }

    static final class TimelineEvent {
        final LocalDate date;
        final String machine;
        final String taskId;
        final String eventKind;
        final LocalDateTime start;
        final LocalDateTime end;
        final Long unitM;
        final List<List<LocalDateTime>> breaks;

        TimelineEvent(
                LocalDate date,
                String machine,
                String taskId,
                String eventKind,
                LocalDateTime start,
                LocalDateTime end,
                Long unitM,
                List<List<LocalDateTime>> breaks) {
            this.date = date;
            this.machine = machine;
            this.taskId = taskId;
            this.eventKind = eventKind;
            this.start = start;
            this.end = end;
            this.unitM = unitM;
            this.breaks = breaks;
        }

        static TimelineEvent from(JsonNode n) {
            if (n == null || !n.isObject()) {
                return null;
            }
            Object d = GanttContractValueDecoder.decodeValue(n.get("date"));
            LocalDate date = GanttContractValueDecoder.toLocalDate(d);
            Object sdt = GanttContractValueDecoder.decodeValue(n.get("start_dt"));
            Object edt = GanttContractValueDecoder.decodeValue(n.get("end_dt"));
            LocalDateTime start = GanttContractValueDecoder.toLocalDateTime(sdt);
            LocalDateTime end = GanttContractValueDecoder.toLocalDateTime(edt);
            if (date == null || start == null || end == null) {
                return null;
            }
            String machine = text(n, "machine");
            String taskId = text(n, "task_id");
            String eventKind = text(n, "event_kind");
            Long unitM = null;
            if (n.has("unit_m") && n.get("unit_m").isNumber()) {
                unitM = n.get("unit_m").longValue();
            }
            List<List<LocalDateTime>> breaks = new ArrayList<>();
            JsonNode bn = n.get("breaks");
            if (bn != null && bn.isArray()) {
                for (JsonNode b : bn) {
                    Object tup = GanttContractValueDecoder.decodeValue(b);
                    if (tup instanceof List<?> list && list.size() >= 2) {
                        LocalDateTime b0 =
                                GanttContractValueDecoder.toLocalDateTime(list.get(0));
                        LocalDateTime b1 =
                                GanttContractValueDecoder.toLocalDateTime(list.get(1));
                        if (b0 != null && b1 != null) {
                            List<LocalDateTime> pair = new ArrayList<>(2);
                            pair.add(b0);
                            pair.add(b1);
                            breaks.add(pair);
                        }
                    }
                }
            }
            return new TimelineEvent(date, machine, taskId, eventKind, start, end, unitM, breaks);
        }

        static String text(JsonNode n, String field) {
            JsonNode x = n.get(field);
            return x != null && x.isTextual() ? x.asText() : "";
        }

        boolean isInBreaks(LocalDateTime winStart, LocalDateTime winEnd) {
            for (List<LocalDateTime> br : breaks) {
                if (br.size() < 2) {
                    continue;
                }
                LocalDateTime b0 = br.get(0);
                LocalDateTime b1 = br.get(1);
                if (rangesOverlap(b0, b1, winStart, winEnd)) {
                    return true;
                }
            }
            return false;
        }

        String cellLabel() {
            if ("machine_daily_startup".equals(eventKind)) {
                return "日次始業準備";
            }
            if ("machine_daily_inspection".equals(eventKind) || "daily_inspection".equals(eventKind)) {
                return "日次点検";
            }
            StringBuilder sb = new StringBuilder();
            if (taskId != null && !taskId.isEmpty()) {
                sb.append(taskId);
            }
            if (unitM != null && unitM > 0) {
                if (sb.length() > 0) {
                    sb.append(" ");
                }
                sb.append(unitM).append("m");
            }
            if (sb.length() == 0 && eventKind != null && !eventKind.isEmpty()) {
                return eventKind;
            }
            return sb.toString();
        }
    }
}
