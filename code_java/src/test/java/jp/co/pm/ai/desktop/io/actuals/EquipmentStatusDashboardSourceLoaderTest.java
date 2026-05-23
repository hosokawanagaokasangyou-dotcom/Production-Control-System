package jp.co.pm.ai.desktop.io.actuals;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.ReloadDecision;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.SourceFingerprint;

class EquipmentStatusDashboardSourceLoaderTest {

    private static Map<String, String> uiForDir(Path dir) {
        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR, dir.toString());
        ui.put(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR, dir.resolve("actual-empty").toString());
        return ui;
    }

    @Test
    void loadIfChanged_skipsWhenFingerprintMatches(@TempDir Path dir) throws Exception {
        Files.createDirectories(dir.resolve("actual-empty"));
        Path taskInputDir = dir.resolve("task-input");
        Files.createDirectories(taskInputDir);
        Files.writeString(taskInputDir.resolve("plan.csv"), "機械名\n");
        Files.writeString(
                dir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME),
                "{\"columns\":[],\"rows\":[]}");

        Map<String, String> ui = uiForDir(dir);
        ui.put(AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR, taskInputDir.toString());
        SourceFingerprint fp = EquipmentStatusDashboardSourceLoader.fingerprint(ui);

        ReloadDecision skip =
                EquipmentStatusDashboardSourceLoader.loadIfChanged(ui, fp, true);
        Assertions.assertTrue(skip.sourcesUnchanged());
        Assertions.assertNull(skip.sources());
    }

    @Test
    void fingerprint_detectsFileChange(@TempDir Path dir) throws Exception {
        Files.createDirectories(dir.resolve("actual-empty"));
        Path taskInputDir = dir.resolve("task-input");
        Files.createDirectories(taskInputDir);
        Path planCsv = taskInputDir.resolve("plan.csv");
        Files.writeString(planCsv, "機械名\nM1\n");
        Files.writeString(
                dir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME),
                "{\"columns\":[],\"rows\":[]}");

        Map<String, String> ui = uiForDir(dir);
        ui.put(AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR, taskInputDir.toString());
        SourceFingerprint fp1 = EquipmentStatusDashboardSourceLoader.fingerprint(ui);

        Thread.sleep(20);
        Files.writeString(planCsv, "機械名,依頼NO\nM1,R1\n");
        SourceFingerprint fp2 = EquipmentStatusDashboardSourceLoader.fingerprint(ui);
        Assertions.assertNotEquals(fp1, fp2);
    }

    @Test
    void fingerprint_prefersTaskInputOverShapedAladdinJson(@TempDir Path dir) throws Exception {
        Files.createDirectories(dir.resolve("actual-empty"));
        Path taskInputDir = dir.resolve("task-input");
        Files.createDirectories(taskInputDir);
        Files.writeString(
                dir.resolve(AppPaths.SHAPED_ALADDIN_PLAN_JSON_BASENAME),
                "{\"columns\":[\"機械名\"],\"rows\":[[\"M1\"]]}");
        Path planCsv = taskInputDir.resolve("plan.csv");
        Files.writeString(planCsv, "機械名,依頼NO\nM1,R1\n");

        Map<String, String> ui = uiForDir(dir);
        ui.put(AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR, taskInputDir.toString());

        SourceFingerprint fp = EquipmentStatusDashboardSourceLoader.fingerprint(ui);
        Assertions.assertTrue(fp.aladdinKey().contains("plan.csv"));
        Assertions.assertFalse(fp.aladdinKey().contains("shaped"));
    }
}
