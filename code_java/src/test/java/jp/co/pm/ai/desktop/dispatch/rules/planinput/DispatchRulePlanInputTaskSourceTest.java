package jp.co.pm.ai.desktop.dispatch.rules.planinput;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.AppPaths;

/** {@link DispatchRulePlanInputTaskSource} loads tasks from repo {@code output/plan_input_tasks.xlsx}. */
class DispatchRulePlanInputTaskSourceTest {

    @Test
    void reload_readsRepoPlanInputWhenPresent() {
        Path plan = AppPaths.defaultStage1PlanTasksPath(Map.of());
        if (!java.nio.file.Files.isRegularFile(plan)) {
            return;
        }
        DispatchRulePlanInputTaskSource src = new DispatchRulePlanInputTaskSource();
        src.reload(Map.of(), null);
        assertFalse(
                src.labels().isEmpty(),
                () -> "expected tasks from " + plan + " but got: " + src.sourceDescription());
        assertTrue(
                src.findRowByLabel(src.labels().get(0)).isPresent(),
                "first label should resolve to a row map");
    }
}
