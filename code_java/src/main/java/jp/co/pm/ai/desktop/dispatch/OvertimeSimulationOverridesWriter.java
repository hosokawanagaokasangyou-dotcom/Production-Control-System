package jp.co.pm.ai.desktop.dispatch;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.LinkedHashMap;
import java.util.Map;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

/** 段階3.5 シミュレーション上書き JSON を書き出す。 */
public final class OvertimeSimulationOverridesWriter {

    private static final ObjectMapper JSON = new ObjectMapper();

    private OvertimeSimulationOverridesWriter() {}

    public record OverridesPayload(
            Map<LocalDate, Map<String, Boolean>> workingOverrides,
            Map<LocalDate, Map<String, Integer>> overtimeMinutes) {}

    public static void write(Path target, OverridesPayload payload) throws Exception {
        ObjectNode root = JSON.createObjectNode();
        root.put("format_version", 1);
        ObjectNode working = JSON.createObjectNode();
        for (Map.Entry<LocalDate, Map<String, Boolean>> e :
                payload.workingOverrides().entrySet()) {
            ObjectNode mem = JSON.createObjectNode();
            for (Map.Entry<String, Boolean> me : e.getValue().entrySet()) {
                mem.put(me.getKey(), me.getValue());
            }
            working.set(e.getKey().toString(), mem);
        }
        root.set("working_overrides", working);
        ObjectNode ot = JSON.createObjectNode();
        for (Map.Entry<LocalDate, Map<String, Integer>> e :
                payload.overtimeMinutes().entrySet()) {
            ObjectNode mem = JSON.createObjectNode();
            for (Map.Entry<String, Integer> me : e.getValue().entrySet()) {
                mem.put(me.getKey(), me.getValue());
            }
            ot.set(e.getKey().toString(), mem);
        }
        root.set("overtime_minutes", ot);
        Files.writeString(target, JSON.writerWithDefaultPrettyPrinter().writeValueAsString(root) + "\n", StandardCharsets.UTF_8);
    }

    public static OverridesPayload buildFromEditState(OvertimeSimulationEditState state) {
        Map<LocalDate, Map<String, Boolean>> working = new LinkedHashMap<>();
        Map<LocalDate, Map<String, Integer>> overtime = new LinkedHashMap<>();
        for (LocalDate d : state.dates()) {
            for (String m : state.members()) {
                OvertimeSimulationEditState.CellState cs = state.cell(d, m);
                if (cs.currentWorking() != cs.baselineWorking()) {
                    working.computeIfAbsent(d, k -> new LinkedHashMap<>()).put(m, cs.currentWorking());
                }
                if (cs.overtimeEdited() && cs.currentWorking()) {
                    overtime.computeIfAbsent(d, k -> new LinkedHashMap<>()).put(m, cs.currentOvertimeMinutes());
                }
            }
        }
        return new OverridesPayload(working, overtime);
    }
}
