package jp.co.pm.ai.desktop.dispatch;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/** 残業シミュレーション JSON の変更セル数サマリ。 */
public final class OvertimeSimulationOverridesReader {

    private static final ObjectMapper JSON = new ObjectMapper();

    private OvertimeSimulationOverridesReader() {}

    public static Stage21TrialSnapshotStore.OverrideSummary summarize(Path overridesJson) {
        if (overridesJson == null || !Files.isRegularFile(overridesJson)) {
            return Stage21TrialSnapshotStore.OverrideSummary.empty();
        }
        try {
            JsonNode root = JSON.readTree(Files.readString(overridesJson, StandardCharsets.UTF_8));
            int workOn = 0;
            int workOff = 0;
            JsonNode working = root.path("working_overrides");
            if (working.isObject()) {
                for (JsonNode dateNode : working) {
                    if (!dateNode.isObject()) {
                        continue;
                    }
                    var it = dateNode.fields();
                    while (it.hasNext()) {
                        var f = it.next();
                        if (f.getValue().asBoolean(false)) {
                            workOn++;
                        } else {
                            workOff++;
                        }
                    }
                }
            }
            int overtime = 0;
            JsonNode ot = root.path("overtime_minutes");
            if (ot.isObject()) {
                for (JsonNode dateNode : ot) {
                    if (dateNode.isObject()) {
                        overtime += dateNode.size();
                    }
                }
            }
            return new Stage21TrialSnapshotStore.OverrideSummary(workOn, workOff, overtime);
        } catch (Exception ignored) {
            return Stage21TrialSnapshotStore.OverrideSummary.empty();
        }
    }
}
