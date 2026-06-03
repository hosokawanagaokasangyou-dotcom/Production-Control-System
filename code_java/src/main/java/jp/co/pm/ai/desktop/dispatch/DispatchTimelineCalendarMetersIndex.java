package jp.co.pm.ai.desktop.dispatch;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.Normalizer;
import java.time.LocalDate;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.io.Stage2EquipmentGanttContractPaths;
import jp.co.pm.ai.desktop.io.gantt.GanttContractValueDecoder;

/**
 * 設備ガント契約 {@code timeline_events} から、依頼×工程×機械ごとの暦日別加工量(m)を引く。
 * 配台表 JSON が配台日1行に集約されていても、タスク×日付表示をガントと揃える。
 */
public final class DispatchTimelineCalendarMetersIndex {

    private static final ObjectMapper JSON = new ObjectMapper();

    private final Map<String, Map<LocalDate, Double>> metersByProfileKey;

    private DispatchTimelineCalendarMetersIndex(Map<String, Map<LocalDate, Double>> metersByProfileKey) {
        this.metersByProfileKey = metersByProfileKey;
    }

    public static DispatchTimelineCalendarMetersIndex empty() {
        return new DispatchTimelineCalendarMetersIndex(Map.of());
    }

    public boolean isLoaded() {
        return !metersByProfileKey.isEmpty();
    }

    public static DispatchTimelineCalendarMetersIndex tryLoadNearResultDispatchJson(Path resultDispatchJson) {
        Path contract = Stage2EquipmentGanttContractPaths.resolveNearResultDispatchJson(resultDispatchJson);
        if (contract == null) {
            return empty();
        }
        try {
            return loadFromContractPath(contract);
        } catch (IOException e) {
            return empty();
        }
    }

    public static DispatchTimelineCalendarMetersIndex loadFromContractPath(Path contractPath)
            throws IOException {
        JsonNode root = JSON.readTree(Files.readString(contractPath, StandardCharsets.UTF_8));
        JsonNode packed = root.get("kwargs_packed");
        if (packed == null || !packed.isObject()) {
            return empty();
        }
        JsonNode eventsNode = packed.get("timeline_events");
        if (eventsNode == null || !eventsNode.isArray()) {
            return empty();
        }
        Map<String, Map<LocalDate, Double>> acc = new LinkedHashMap<>();
        for (JsonNode en : eventsNode) {
            if (en == null || !en.isObject()) {
                continue;
            }
            String kind = text(en, "event_kind");
            if (!"machining".equals(kind)) {
                continue;
            }
            String taskId = text(en, "task_id").strip();
            if (taskId.isEmpty()) {
                continue;
            }
            Object d = GanttContractValueDecoder.decodeValue(en.get("date"));
            LocalDate day = GanttContractValueDecoder.toLocalDate(d);
            if (day == null) {
                continue;
            }
            String machineLine = text(en, "machine");
            String[] split = splitEquipmentLine(machineLine);
            String process = split[0];
            String machine = split[1];
            double unitM = number(en, "unit_m");
            double unitsDone = number(en, "units_done");
            if (unitM <= 1e-12 || unitsDone <= 1e-12) {
                continue;
            }
            double meters = unitsDone * unitM;
            String key = profileKey(taskId, process, machine);
            acc.computeIfAbsent(key, k -> new LinkedHashMap<>())
                    .merge(day, meters, Double::sum);
        }
        return new DispatchTimelineCalendarMetersIndex(Map.copyOf(acc));
    }

    /**
     * タイムラインに当該プロファイルの加工イベントがあるとき、その暦日の m。無いプロファイルは empty（表データへフォールバック）。
     */
    public Optional<Double> metersForTaskProfile(
            String taskId, String process, String machine, LocalDate day) {
        if (day == null || metersByProfileKey.isEmpty()) {
            return Optional.empty();
        }
        String procNorm = normalizeEquipmentMatchKey(process);
        String machNorm = normalizeEquipmentMatchKey(machine);
        double sum = 0.0;
        boolean hit = false;
        for (Map.Entry<String, Map<LocalDate, Double>> en : metersByProfileKey.entrySet()) {
            String[] parts = en.getKey().split("\\|", 3);
            if (parts.length < 3) {
                continue;
            }
            if (!procNorm.equals(parts[1]) || !machNorm.equals(parts[2])) {
                continue;
            }
            if (!taskIdMatchesFamily(parts[0], taskId)) {
                continue;
            }
            hit = true;
            sum += en.getValue().getOrDefault(day, 0.0);
        }
        return hit ? Optional.of(sum) : Optional.empty();
    }

    /** 枝番依頼NO（例 V6-2-01）を親依頼NO（V6-2）へ集約してガント契約と手動修正ワイド表を揃える。 */
    static boolean taskIdMatchesFamily(String eventTaskId, String queryTaskId) {
        String ev = normalizeEquipmentMatchKey(eventTaskId);
        String q = normalizeEquipmentMatchKey(queryTaskId);
        if (ev.isEmpty() || q.isEmpty()) {
            return false;
        }
        if (ev.equals(q)) {
            return true;
        }
        if (!ev.startsWith(q) || ev.length() <= q.length()) {
            return false;
        }
        return ev.charAt(q.length()) == '-';
    }

    /** 工程×機械×暦日で全依頼のタイムライン加工量を合算（工程＋機械×日付用）。 */
    public Optional<Double> metersForProcessMachine(String process, String machine, LocalDate day) {
        if (day == null || metersByProfileKey.isEmpty()) {
            return Optional.empty();
        }
        String procNorm = normalizeEquipmentMatchKey(process);
        String machNorm = normalizeEquipmentMatchKey(machine);
        double sum = 0.0;
        boolean hit = false;
        for (Map.Entry<String, Map<LocalDate, Double>> en : metersByProfileKey.entrySet()) {
            String[] parts = en.getKey().split("\\|", 3);
            if (parts.length < 3) {
                continue;
            }
            if (!procNorm.equals(parts[1]) || !machNorm.equals(parts[2])) {
                continue;
            }
            hit = true;
            sum += en.getValue().getOrDefault(day, 0.0);
        }
        return hit ? Optional.of(sum) : Optional.empty();
    }

    static String profileKey(String taskId, String process, String machine) {
        return normalizeEquipmentMatchKey(taskId)
                + "|"
                + normalizeEquipmentMatchKey(process)
                + "|"
                + normalizeEquipmentMatchKey(machine);
    }

    static String[] splitEquipmentLine(String line) {
        if (line == null || line.isEmpty()) {
            return new String[] {"", ""};
        }
        int p = line.indexOf('+');
        if (p < 0) {
            return new String[] {"", line.strip()};
        }
        return new String[] {line.substring(0, p).strip(), line.substring(p + 1).strip()};
    }

    static String normalizeEquipmentMatchKey(String val) {
        if (val == null || val.isBlank()) {
            return "";
        }
        String t = Normalizer.normalize(val, Normalizer.Form.NFKC);
        t = t.replace('\u00a0', ' ').replace('\u3000', ' ');
        t = t.replaceAll("[\u200b\u200c\u200d\ufeff]", "");
        return t.replaceAll("\\s+", " ").strip();
    }

    private static String text(JsonNode n, String field) {
        JsonNode x = n.get(field);
        return x != null && x.isTextual() ? x.asText() : "";
    }

    private static double number(JsonNode n, String field) {
        JsonNode x = n.get(field);
        return x != null && x.isNumber() ? x.doubleValue() : 0.0;
    }
}
