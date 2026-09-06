package jp.co.pm.ai.desktop.io.actuals;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.io.NetworkSourceFileReloadCache;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.LoadedSources;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.ReloadDecision;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardSourceLoader.SourceFingerprint;

class EquipmentStatusDashboardSourceLoaderTest {

    private static Map<String, String> uiForDir(Path dir) {
        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_REPO_ROOT, dir.toString());
        ui.put(AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR, dir.toString());
        ui.put(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR, dir.resolve("actual-empty").toString());
        return ui;
    }

    @Test
    void loadIfChanged_skipsWhenFingerprintMatches(@TempDir Path dir) throws Exception {
        Files.createDirectories(dir.resolve("actual-empty"));
        Path taskInputDir = dir.resolve("task-input");
        Files.createDirectories(taskInputDir);
        Files.writeString(taskInputDir.resolve("plan.xlsx"), "機械名\n");
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
        Path planXlsx = taskInputDir.resolve("plan.xlsx");
        Files.writeString(planXlsx, "機械名\nM1\n");
        Files.writeString(
                dir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME),
                "{\"columns\":[],\"rows\":[]}");

        Map<String, String> ui = uiForDir(dir);
        ui.put(AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR, taskInputDir.toString());
        SourceFingerprint fp1 = EquipmentStatusDashboardSourceLoader.fingerprint(ui);

        Thread.sleep(20);
        Files.writeString(planXlsx, "機械名,依頼NO\nM1,R1\n");
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
        Path planXlsx = taskInputDir.resolve("plan.xlsx");
        Files.writeString(planXlsx, "機械名,依頼NO\nM1,R1\n");

        Map<String, String> ui = uiForDir(dir);
        ui.put(AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR, taskInputDir.toString());

        SourceFingerprint fp = EquipmentStatusDashboardSourceLoader.fingerprint(ui);
        Assertions.assertTrue(fp.aladdinKey().contains("task-input-newest"));
        Assertions.assertFalse(fp.aladdinKey().contains("shaped"));
    }

    @Test
    void load_doesNotRejectActualFileOverDefaultUiLimit(@TempDir Path dir) throws Exception {
        Path actualDir = dir.resolve("actual");
        Files.createDirectories(actualDir);
        Path taskInputDir = dir.resolve("task-input");
        Files.createDirectories(taskInputDir);
        Path big = actualDir.resolve("actual.xlsx");
        Files.write(
                big,
                new byte[(int) (AppPaths.DEFAULT_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES + 1024)]);
        Files.writeString(taskInputDir.resolve("plan.csv"), "機械名,依頼NO\nM1,R1\n");
        Files.writeString(
                dir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME),
                "{\"columns\":[],\"rows\":[]}");

        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_REPO_ROOT, dir.toString());
        ui.put(AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR, dir.toString());
        ui.put(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR, actualDir.toString());
        ui.put(AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR, taskInputDir.toString());

        LoadedSources loaded = EquipmentStatusDashboardSourceLoader.load(ui);
        Assertions.assertNotNull(loaded);
        Assertions.assertTrue(loaded.loadNotice().contains("大きい"));
        Assertions.assertTrue(loaded.actuals().rows().isEmpty());
        Assertions.assertTrue(loaded.loadStats().totalSourceBytes() > 0L);
        Assertions.assertTrue(loaded.loadStats().loadDurationMs() >= 0L);
    }

    @Test
    void formatLoadStatsSummary_andByteSize() {
        EquipmentStatusDashboardSourceLoader.LoadStats stats =
                new EquipmentStatusDashboardSourceLoader.LoadStats(
                        1_572_864L, 850L, 12034, 560, 89);
        String summary = EquipmentStatusDashboardSourceLoader.formatLoadStatsSummary(stats);
        Assertions.assertTrue(summary.contains("1.5 MiB"));
        Assertions.assertTrue(summary.contains("850 ms"));
        Assertions.assertTrue(summary.contains("12,034"));
        Assertions.assertEquals("512 B", EquipmentStatusDashboardSourceLoader.formatByteSize(512L));
        Assertions.assertEquals("1.0 KiB", EquipmentStatusDashboardSourceLoader.formatByteSize(1024L));
        Assertions.assertEquals("1.50 s", EquipmentStatusDashboardSourceLoader.formatLoadDuration(1500L));
    }

    @Test
    void load_prefersShapedActualsJsonWhenPresent(@TempDir Path dir) throws Exception {
        Path actualDir = dir.resolve("actual");
        Files.createDirectories(actualDir);
        Path big = actualDir.resolve("actual.xlsx");
        Files.write(
                big,
                new byte[(int) (AppPaths.DEFAULT_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES + 1024)]);
        Files.writeString(
                dir.resolve(AppPaths.SHAPED_PROCESSING_ACTUALS_JSON_BASENAME),
                "{\"columns\":[\"機械名\",\"依頼NO\",\"工程名\",\"加工日\"],"
                        + "\"rows\":[[\"M1\",\"R1\",\"P1\",\"2026/05/25\"]]}");
        Files.writeString(
                dir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME),
                "{\"columns\":[],\"rows\":[]}");

        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_REPO_ROOT, dir.toString());
        ui.put(AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR, dir.toString());
        ui.put(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR, actualDir.toString());

        LoadedSources loaded = EquipmentStatusDashboardSourceLoader.load(ui);
        Assertions.assertEquals(1, loaded.actuals().rows().size());
        Assertions.assertTrue(loaded.loadNotice().isBlank());
    }

    @Test
    void projectShapedForCache_filtersToEssentialHeaders() {
        List<String> headers = List.of("倉庫名", "機械名", "工程名", "依頼NO", "不要列1", "実加工数", "加工日");
        List<List<String>> rows = List.of(List.of("W1", "M1", "P1", "R1", "X", "100", "2026/08/01"));
        PlanInputTabularIo.TabularSheet in = new PlanInputTabularIo.TabularSheet(headers, rows);

        PlanInputTabularIo.TabularSheet projected =
                EquipmentStatusDashboardSourceLoader.projectShapedForCache(in);

        Assertions.assertEquals(
                List.of("機械名", "工程名", "依頼NO", "実加工数", "加工日"),
                projected.headers());
        Assertions.assertEquals(1, projected.rows().size());
        Assertions.assertEquals(
                List.of("M1", "P1", "R1", "100", "2026/08/01"),
                projected.rows().getFirst());
    }

    @Test
    void load_savesShapedActualsJsonAndPopulatesMemoryCache(@TempDir Path dir) throws Exception {
        NetworkSourceFileReloadCache.clearAll();
        Path actualDir = dir.resolve("actual");
        Files.createDirectories(actualDir);
        // CSV formatted actual detail (readRaw supports CSV as well as Excel)
        Path actualCsv = actualDir.resolve("actual.csv");
        Files.writeString(
                actualCsv,
                "工程名,機械名,依頼NO,加工日,実加工数\n"
                        + "P1,M1,R1,2026/08/01,500\n");
        Files.writeString(
                dir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME),
                "{\"columns\":[],\"rows\":[]}");

        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_REPO_ROOT, dir.toString());
        ui.put(AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR, dir.toString());
        ui.put(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK, actualCsv.toString());

        LoadedSources loaded = EquipmentStatusDashboardSourceLoader.load(ui);
        Assertions.assertNotNull(loaded);
        Assertions.assertEquals(1, loaded.actuals().rows().size());

        // Verify memory cache hit (NetworkSourceDirResolver copies live file to actual-detail-newest.*)
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(ui);
        Assertions.assertTrue(
                NetworkSourceFileReloadCache.matchActuals(r.actualDetailPath().orElseThrow()).isPresent(),
                "Memory cache should be populated for resolved actual detail path");

        // Verify shaped JSON file was saved to disk
        Path shapedJson = AppPaths.resolveShapedProcessingActualsJsonPath(ui);
        Assertions.assertTrue(
                Files.isRegularFile(shapedJson),
                "shaped_processing_actuals.json should be saved");
        String jsonContent = Files.readString(shapedJson);
        Assertions.assertTrue(jsonContent.contains("実加工数"));
        Assertions.assertTrue(jsonContent.contains("500"));
    }
}
