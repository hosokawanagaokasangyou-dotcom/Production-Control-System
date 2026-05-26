package jp.co.pm.ai.desktop.dispatch.rules.trace;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths;

/** Load dispatch_rule_applications.json sidecar. */
public final class DispatchRuleTraceLoader {

    public record ApplicationEvent(
            String taskId,
            String ruleId,
            int applyOrder,
            String phase,
            String effect,
            String summaryJa) {}

    private static final ObjectMapper JSON = new ObjectMapper();

    private DispatchRuleTraceLoader() {}

    public static Path sidecarPath(Map<String, String> ui) {
        return DispatchRulePaths.workDirectory(ui).resolve("dispatch_rule_applications.json");
    }

    public static List<ApplicationEvent> loadFromWorkDir(Map<String, String> ui) throws IOException {
        Path p = sidecarPath(ui);
        if (!Files.isRegularFile(p)) {
            return List.of();
        }
        JsonNode root = JSON.readTree(Files.readString(p, StandardCharsets.UTF_8));
        List<ApplicationEvent> out = new ArrayList<>();
        for (JsonNode n : root.withArray("events")) {
            out.add(
                    new ApplicationEvent(
                            n.path("task_id").asText(),
                            n.path("rule_id").asText(),
                            n.path("apply_order").asInt(),
                            n.path("phase").asText(),
                            n.path("effect").asText(),
                            n.path("summary_ja").asText()));
        }
        return out;
    }

    public static void reloadFromLatestStage2(Map<String, String> ui) {
        // sidecar is written by Python trace_recorder during stage2
    }
}
