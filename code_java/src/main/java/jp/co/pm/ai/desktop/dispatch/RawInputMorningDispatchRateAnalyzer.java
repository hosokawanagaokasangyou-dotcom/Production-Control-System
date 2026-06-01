package jp.co.pm.ai.desktop.dispatch;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.io.gantt.GanttContractValueDecoder;

/**
 * 設備ガント契約 {@code timeline_events} と原反投入日から、暦日ごとの午前配台率を算出する。
 *
 * <p>午前帯は Python {@code same_day_raw_start_limit}（12:45）と揃え {@code 08:45}～{@code 12:45}。
 * 原反投入日＝当暦日の依頼が同日 12:45 以降にしか加工できないことが、当該日の午前配台率低下の原因とみなす。
 */
public final class RawInputMorningDispatchRateAnalyzer {

    /** Python {@code DEFAULT_START_TIME} と同一。 */
    public static final LocalTime MORNING_WINDOW_START = LocalTime.of(8, 45);

    /** Python {@code same_day_raw_start_limit}（DISPATCHABLE_FROM_TIME=12:45）と同一。 */
    public static final LocalTime MORNING_WINDOW_END = LocalTime.of(12, 45);

    public static final double RATE_THRESHOLD = 0.5;

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String EVENT_MACHINING = "machining";

    private RawInputMorningDispatchRateAnalyzer() {}

    public record DayLowRate(
            LocalDate date,
            double morningRate,
            long morningUsedMinutes,
            long morningCapacityMinutes,
            int rawInputSameDayTaskCount,
            List<String> rawInputSameDayTaskIds) {}

    public record AnalysisResult(List<DayLowRate> lowRateDays) {
        public boolean hasWarnings() {
            return lowRateDays != null && !lowRateDays.isEmpty();
        }
    }

    public static AnalysisResult analyze(Path equipmentGanttContractPath, Map<String, LocalDate> rawInputByTaskId)
            throws IOException {
        if (equipmentGanttContractPath == null
                || !Files.isRegularFile(equipmentGanttContractPath)
                || rawInputByTaskId == null
                || rawInputByTaskId.isEmpty()) {
            return new AnalysisResult(List.of());
        }
        JsonNode root = JSON.readTree(Files.readString(equipmentGanttContractPath, StandardCharsets.UTF_8));
        JsonNode packed = root.get("kwargs_packed");
        if (packed == null || !packed.isObject()) {
            return new AnalysisResult(List.of());
        }
        JsonNode eventsNode = packed.get("timeline_events");
        if (eventsNode == null || !eventsNode.isArray()) {
            return new AnalysisResult(List.of());
        }

        Map<LocalDate, Map<String, Long>> usedByDayMachine = new LinkedHashMap<>();
        Map<LocalDate, Set<String>> activeMachinesByDay = new LinkedHashMap<>();
        Map<LocalDate, Map<String, LocalDateTime>> earliestMachiningStartByDayTask =
                new LinkedHashMap<>();

        for (JsonNode en : eventsNode) {
            if (en == null || !en.isObject()) {
                continue;
            }
            ParsedEvent ev = ParsedEvent.from(en);
            if (ev == null) {
                continue;
            }
            String machineKey = ev.machineKey();
            activeMachinesByDay
                    .computeIfAbsent(ev.day(), d -> new LinkedHashSet<>())
                    .add(machineKey);

            if (!EVENT_MACHINING.equals(ev.eventKind())) {
                continue;
            }
            long overlap = overlapMinutes(ev.start(), ev.end(), ev.day());
            if (overlap > 0) {
                usedByDayMachine
                        .computeIfAbsent(ev.day(), d -> new LinkedHashMap<>())
                        .merge(machineKey, overlap, Long::sum);
            }
            String tid = ev.taskId();
            if (tid != null && !tid.isBlank()) {
                earliestMachiningStartByDayTask
                        .computeIfAbsent(ev.day(), d -> new LinkedHashMap<>())
                        .merge(
                                tid.strip(),
                                ev.start(),
                                (a, b) -> a.isBefore(b) ? a : b);
            }
        }

        long morningCapacityPerMachine =
                Duration.between(MORNING_WINDOW_START, MORNING_WINDOW_END).toMinutes();
        List<DayLowRate> warnings = new ArrayList<>();

        for (Map.Entry<LocalDate, Set<String>> dayEn : activeMachinesByDay.entrySet()) {
            LocalDate day = dayEn.getKey();
            Set<String> machines = dayEn.getValue();
            if (machines.isEmpty()) {
                continue;
            }
            long capacity = morningCapacityPerMachine * machines.size();
            long used =
                    usedByDayMachine.getOrDefault(day, Map.of()).values().stream()
                            .mapToLong(Long::longValue)
                            .sum();
            double rate = capacity > 0 ? (double) used / capacity : 0.0;
            if (rate >= RATE_THRESHOLD - 1e-9) {
                continue;
            }

            List<String> rawSameDayTasks = new ArrayList<>();
            Map<String, LocalDateTime> taskStarts =
                    earliestMachiningStartByDayTask.getOrDefault(day, Map.of());
            for (Map.Entry<String, LocalDate> rawEn : rawInputByTaskId.entrySet()) {
                if (!day.equals(rawEn.getValue())) {
                    continue;
                }
                String taskId = rawEn.getKey();
                LocalDateTime earliest = taskStarts.get(taskId);
                if (earliest == null) {
                    continue;
                }
                if (!earliest.toLocalDate().equals(day)) {
                    continue;
                }
                if (!earliest.toLocalTime().isBefore(MORNING_WINDOW_END)) {
                    rawSameDayTasks.add(taskId);
                }
            }
            if (rawSameDayTasks.isEmpty()) {
                continue;
            }
            warnings.add(
                    new DayLowRate(
                            day,
                            rate,
                            used,
                            capacity,
                            rawSameDayTasks.size(),
                            List.copyOf(rawSameDayTasks)));
        }
        return new AnalysisResult(List.copyOf(warnings));
    }

    static long overlapMinutes(
            LocalDateTime eventStart, LocalDateTime eventEnd, LocalDate day) {
        LocalDateTime winStart = LocalDateTime.of(day, MORNING_WINDOW_START);
        LocalDateTime winEnd = LocalDateTime.of(day, MORNING_WINDOW_END);
        LocalDateTime s = eventStart.isBefore(winStart) ? winStart : eventStart;
        LocalDateTime e = eventEnd.isAfter(winEnd) ? winEnd : eventEnd;
        if (!e.isAfter(s)) {
            return 0L;
        }
        return Duration.between(s, e).toMinutes();
    }

    static LocalDateTime parseDateTime(JsonNode n) {
        if (n == null || n.isNull()) {
            return null;
        }
        Object decoded = GanttContractValueDecoder.decodeValue(n);
        LocalDateTime ldt = GanttContractValueDecoder.toLocalDateTime(decoded);
        if (ldt != null) {
            return ldt;
        }
        if (decoded instanceof String s) {
            return parseDateTimeString(s);
        }
        if (n.isTextual()) {
            return parseDateTimeString(n.asText());
        }
        return null;
    }

    private static LocalDateTime parseDateTimeString(String s) {
        if (s == null) {
            return null;
        }
        String t = s.strip();
        if (t.isEmpty()) {
            return null;
        }
        if (t.contains("T")) {
            return LocalDateTime.parse(t.length() >= 19 ? t.substring(0, 19) : t);
        }
        if (t.length() >= 19 && t.charAt(10) == ' ') {
            return LocalDateTime.parse(t.substring(0, 19));
        }
        return null;
    }

    private record ParsedEvent(
            LocalDate day,
            String machineKey,
            String taskId,
            String eventKind,
            LocalDateTime start,
            LocalDateTime end) {

        static ParsedEvent from(JsonNode n) {
            Object d = GanttContractValueDecoder.decodeValue(n.get("date"));
            LocalDate day = GanttContractValueDecoder.toLocalDate(d);
            if (day == null && d instanceof String ds) {
                day = LocalDate.parse(ds.strip());
            }
            LocalDateTime start = parseDateTime(n.get("start_dt"));
            LocalDateTime end = parseDateTime(n.get("end_dt"));
            if (day == null || start == null || end == null) {
                return null;
            }
            String occ = text(n, "machine_occupancy_key");
            String machine = text(n, "machine");
            String machineKey = !occ.isBlank() ? occ.strip() : machine.strip();
            if (machineKey.isEmpty()) {
                return null;
            }
            return new ParsedEvent(
                    day,
                    machineKey,
                    text(n, "task_id"),
                    text(n, "event_kind"),
                    start,
                    end);
        }

        private static String text(JsonNode n, String field) {
            JsonNode x = n.get(field);
            return x != null && x.isTextual() ? x.asText() : "";
        }
    }
}
