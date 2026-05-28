package jp.co.pm.ai.desktop.dispatch;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/** 段階3.5 残業シミュレーション上書き JSON の読取り。 */
public final class OvertimeSimulationOverridesReader {

    private static final ObjectMapper JSON = new ObjectMapper();

    private OvertimeSimulationOverridesReader() {}

    public static Stage35BaselineActualSnapshotStore.OverrideSummary summarize(Path overridesJson) {
        if (overridesJson == null || !Files.isRegularFile(overridesJson)) {
            return Stage35BaselineActualSnapshotStore.OverrideSummary.empty();
        }
        try {
            JsonNode root =
                    JSON.readTree(Files.readString(overridesJson, StandardCharsets.UTF_8));
            int workOn = countMemberDateCells(root.path("working_overrides"), true);
            int workOff = countMemberDateCells(root.path("working_overrides"), false);
            int overtime = countOvertimeCells(root.path("overtime_minutes"));
            return new Stage35BaselineActualSnapshotStore.OverrideSummary(workOn, workOff, overtime);
        } catch (Exception ignored) {
            return Stage35BaselineActualSnapshotStore.OverrideSummary.empty();
        }
    }

    private static int countMemberDateCells(JsonNode workingOverrides, boolean flag) {
        if (workingOverrides == null || !workingOverrides.isObject()) {
            return 0;
        }
        int count = 0;
        for (JsonNode memMap : workingOverrides) {
            if (memMap == null || !memMap.isObject()) {
                continue;
            }
            for (JsonNode v : memMap) {
                if (v != null && v.isBoolean() && v.booleanValue() == flag) {
                    count++;
                }
            }
        }
        return count;
    }

    private static int countOvertimeCells(JsonNode overtimeMinutes) {
        if (overtimeMinutes == null || !overtimeMinutes.isObject()) {
            return 0;
        }
        int count = 0;
        for (JsonNode memMap : overtimeMinutes) {
            if (memMap == null || !memMap.isObject()) {
                continue;
            }
            count += memMap.size();
        }
        return count;
    }
}
