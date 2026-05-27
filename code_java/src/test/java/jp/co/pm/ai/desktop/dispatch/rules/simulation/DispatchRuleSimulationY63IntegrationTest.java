package jp.co.pm.ai.desktop.dispatch.rules.simulation;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleDocument;
import jp.co.pm.ai.desktop.dispatch.rules.planinput.DispatchRulePlanInputTaskSource;

/** End-to-end simulate_dispatch_rules.py for Y6-3 / L13 (requires Python 3.14+). */
class DispatchRuleSimulationY63IntegrationTest {

    private static final ObjectMapper JSON =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    @Test
    void simulate_y63_l13_viaPythonChild() throws Exception {
        Path py = MainShellController.defaultPythonPathWhenShellMissing();
        String pyName = py.getFileName().toString().toLowerCase();
        org.junit.jupiter.api.Assumptions.assumeTrue(
                pyName.contains("3.14") || Files.isRegularFile(Path.of("/usr/bin/python3.14")),
                "Python 3.14 required for planning_core");

        Path plan = AppPaths.defaultStage1PlanTasksPath(Map.of());
        org.junit.jupiter.api.Assumptions.assumeTrue(
                java.nio.file.Files.isRegularFile(plan), "missing " + plan);

        DispatchRulePlanInputTaskSource tasks = new DispatchRulePlanInputTaskSource();
        tasks.reload(Map.of(), null);
        org.junit.jupiter.api.Assumptions.assumeFalse(tasks.labels().isEmpty(), tasks.sourceDescription());

        String label =
                tasks.labels().stream()
                        .filter(l -> l.startsWith("Y6-3 / 接続"))
                        .findFirst()
                        .orElse(
                                tasks.labels().stream()
                                        .filter(l -> l.startsWith("Y6-3 /"))
                                        .findFirst()
                                        .orElse(tasks.labels().get(0)));
        Map<String, String> taskRow =
                tasks.findRowByLabel(label).orElseThrow();
        Map<String, String> secTaskRow =
                tasks.findSecRowForRequest(taskRow.getOrDefault("依頼NO", "")).orElse(Map.of());

        Path rulesJson =
                AppPaths.resolveRepoRoot(Map.of())
                        .resolve("code")
                        .resolve("json")
                        .resolve("dispatch_special_rules")
                        .resolve("dispatch_special_rules.json");
        org.junit.jupiter.api.Assumptions.assumeTrue(
                java.nio.file.Files.isRegularFile(rulesJson), "missing rules json");

        DispatchRuleDocument doc = JSON.readValue(rulesJson.toFile(), DispatchRuleDocument.class);
        for (var rule : doc.rules) {
            if ("L13".equals(rule.id)) {
                rule.executionMode = "dsl";
            }
        }

        Path pythonExe =
                java.nio.file.Files.isRegularFile(Path.of("/usr/bin/python3.14"))
                        ? Path.of("/usr/bin/python3.14")
                        : py;

        DispatchRuleSimulationResult result =
                DispatchRuleSimulationService.simulate(
                        pythonExe,
                        Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, AppPaths.resolveRepoRoot(Map.of()).toString()),
                        doc,
                        new LinkedHashMap<>(taskRow),
                        new LinkedHashMap<>(secTaskRow),
                        "L13",
                        Map.of("metrics", Map.of("initial_wip", 5, "request_roll_diff", 10)),
                        true,
                        line -> {});

        assertFalse(result.steps().isEmpty(), "expected simulation steps");
        assertTrue(result.finalBlocked(), "Y6-3 L13 should block when WIP reaches threshold");
        assertTrue(result.rollTotal() >= 1, "expected roll_total");
        assertTrue(
                result.steps().stream().anyMatch(DispatchRuleSimulationStep::rollAccumulateStep),
                "expected roll accumulation steps");
        if (!secTaskRow.isEmpty() && !result.finalBlocked()) {
            assertTrue(
                    result.steps().stream().anyMatch(s -> "sec".equals(s.flowPhase())),
                    "expected SEC phase steps when pipeline completes");
        }
    }
}
