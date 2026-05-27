package jp.co.pm.ai.desktop.dispatch.rules.simulation;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleDocument;

/** Runs {@code tools/simulate_dispatch_rules.py} for the rule test lab. */
public final class DispatchRuleSimulationService {

    private static final ObjectMapper JSON =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    private DispatchRuleSimulationService() {}

    public static DispatchRuleSimulationResult simulate(
            Path pythonExe,
            Map<String, String> uiEnv,
            DispatchRuleDocument document,
            Map<String, String> taskRow,
            Map<String, String> secTaskRow,
            String ruleId,
            Map<String, Object> contextOverrides,
            boolean allRolls,
            Consumer<String> logLine)
            throws Exception {
        Path pyDir = AppPaths.resolvePythonScriptDir(uiEnv);
        Path script = pyDir.resolve("tools/simulate_dispatch_rules.py");
        if (!Files.isRegularFile(script)) {
            script = pyDir.getParent().resolve("tools/simulate_dispatch_rules.py");
        }
        if (!Files.isRegularFile(script)) {
            return DispatchRuleSimulationResult.empty("simulate_dispatch_rules.py が見つかりません");
        }
        Map<String, Object> request = new LinkedHashMap<>();
        request.put("document", document);
        request.put("task_row", taskRow != null ? taskRow : Map.of());
        if (secTaskRow != null && !secTaskRow.isEmpty()) {
            request.put("sec_task_row", secTaskRow);
        }
        if (ruleId != null && !ruleId.isBlank()) {
            request.put("rule_id", ruleId);
        }
        if (contextOverrides != null && !contextOverrides.isEmpty()) {
            request.put("context_overrides", contextOverrides);
        }
        request.put("all_rolls", allRolls);
        String payload = JSON.writeValueAsString(request);
        ProcessBuilder pb =
                new ProcessBuilder(pythonExe.toString(), script.toAbsolutePath().toString(), "--stdin");
        pb.directory(pyDir.toFile());
        pb.redirectErrorStream(true);
        PythonProcessRunner.mergeUiEnvIntoProcess(pb, uiEnv, pyDir);
        Process p = pb.start();
        p.getOutputStream().write(payload.getBytes(StandardCharsets.UTF_8));
        p.getOutputStream().close();
        StringBuilder merged = new StringBuilder();
        try (BufferedReader br =
                new BufferedReader(
                        new InputStreamReader(p.getInputStream(), StandardCharsets.UTF_8))) {
            String line;
            while ((line = br.readLine()) != null) {
                if (logLine != null) {
                    logLine.accept(line);
                }
                if (!line.startsWith("[") && !merged.isEmpty()) {
                    merged.append('\n');
                }
                if (!line.startsWith("[")) {
                    merged.append(line);
                }
            }
        }
        boolean finished = p.waitFor(120, TimeUnit.SECONDS);
        if (!finished) {
            p.destroyForcibly();
            throw new IllegalStateException("simulate_dispatch_rules.py timeout");
        }
        if (p.exitValue() != 0) {
            throw new IllegalStateException(
                    "simulate_dispatch_rules.py exit " + p.exitValue() + ": " + merged);
        }
        JsonNode root = JSON.readTree(extractLastJsonObjectLine(merged.toString()));
        List<DispatchRuleSimulationStep> steps = new ArrayList<>();
        for (JsonNode n : root.withArray("steps")) {
            steps.add(parseStep(n));
        }
        return new DispatchRuleSimulationResult(
                root.path("final_blocked").asBoolean(false),
                root.path("summary_ja").asText(""),
                steps,
                root.path("roll_total").asInt(0),
                root.path("blocked_at_roll").asInt(0));
    }

    private static DispatchRuleSimulationStep parseStep(JsonNode n) {
        JsonNode m = n.path("metrics");
        return new DispatchRuleSimulationStep(
                n.path("sequence").asInt(),
                n.path("phase").asText(),
                n.path("rule_id").asText(),
                n.path("node_id").asText(),
                n.path("node_type").asText(),
                n.path("edge_from").asText(null),
                n.path("edge_to").asText(null),
                n.path("effect").asText(null),
                n.path("summary_ja").asText(""),
                n.path("roll_index").asInt(0),
                n.path("roll_total").asInt(0),
                n.path("wip_count").asDouble(0),
                n.path("animation_kind").asText(""),
                m.path("pre_input_raw_rolls").asInt(0),
                m.path("connection_rolls").asInt(m.path("connection_machine_rolls").asInt(0)),
                m.path("sec_before_wip_rolls").asInt(m.path("sec_before_rolls").asInt(0)),
                m.path("sec_complete_rolls").asInt(m.path("sec_machine_rolls").asInt(0)),
                n.path("flow_phase").asText(""));
    }

    /**
     * Python bootstrap の stdout ログ（{@code 2026-01-01 00:00:00,000 - INFO - …}）混在時も、
     * 最終行の JSON オブジェクトだけを返す（{@code attendance_overtime_preview} と同趣旨）。
     */
    static String extractLastJsonObjectLine(String merged) {
        if (merged == null || merged.isBlank()) {
            return "{}";
        }
        String[] lines = merged.split("\\R", -1);
        for (int i = lines.length - 1; i >= 0; i--) {
            String t = lines[i].trim();
            if (t.startsWith("{") && t.endsWith("}")) {
                return t;
            }
        }
        return merged.trim();
    }
}
