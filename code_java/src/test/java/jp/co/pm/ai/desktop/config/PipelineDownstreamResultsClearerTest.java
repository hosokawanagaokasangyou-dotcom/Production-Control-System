package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.PlanInputStage3TabController;
import jp.co.pm.ai.desktop.dispatch.Stage21TrialSnapshotStore;
import jp.co.pm.ai.desktop.dispatch.Stage3PlanningMetaStore;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

class PipelineDownstreamResultsClearerTest {

    @TempDir Path temp;

    private Map<String, String> ui(Path repo) {
        Path python = repo.resolve("code").resolve("python");
        return Map.of(
                AppPaths.KEY_PM_AI_REPO_ROOT,
                repo.toString(),
                AppPaths.KEY_PM_AI_CODE_PYTHON_DIR,
                python.toString());
    }

    @Test
    void clearStage2ThroughStage32_removesDispatchSidecarsAndStage3SheetRows() throws Exception {
        Path repo = temp.resolve("repo");
        Path output = repo.resolve("code").resolve("output");
        Path python = repo.resolve("code").resolve("python");
        Files.createDirectories(output);
        Files.createDirectories(python);
        Files.writeString(python.resolve("task_extract_stage1.py"), "# stub\n");

        Map<String, String> ui = ui(repo);
        Path dispatchJson = output.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME);
        Files.writeString(dispatchJson, "{\"rows\":[]}");
        Files.writeString(
                Stage21TrialSnapshotStore.sidecarPathFor(dispatchJson),
                "{\"stage21_applied\":true}");
        Files.writeString(
                Stage3PlanningMetaStore.sidecarPath(dispatchJson),
                "{\"variant\":\"3.1\"}");
        Files.writeString(output.resolve("shaped_aladdin_plan.json"), "{}");
        Files.writeString(output.resolve("計画2606011200000001.xlsx"), "stub");
        Files.createDirectories(AppPaths.resolveStage21OutputDir(ui));
        Files.writeString(
                AppPaths.resolveStage21OutputDir(ui).resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME),
                "{}");

        Path workbook = output.resolve(AppPaths.STAGE1_PLAN_TASKS_FILENAME);
        PlanInputTabularIo.write(
                workbook,
                "配台計画_タスク入力",
                new PlanInputTabularIo.TabularSheet(
                        java.util.List.of("依頼NO"), java.util.List.of(java.util.List.of("Y1"))));
        PlanInputTabularIo.writeExcelSheetPreservingOthers(
                workbook,
                PlanInputStage3TabController.STAGE3_SHEET_NAME,
                new PlanInputTabularIo.TabularSheet(
                        java.util.List.of("依頼NO", "元依頼NO"),
                        java.util.List.of(java.util.List.of("Y1-01", "Y1"))));

        PipelineDownstreamResultsClearer.ClearResult result =
                PipelineDownstreamResultsClearer.clearStage2ThroughStage32(ui);

        assertTrue(result.anyDeleted());
        assertFalse(Files.exists(dispatchJson));
        assertFalse(Files.isRegularFile(Stage21TrialSnapshotStore.sidecarPathFor(dispatchJson)));
        assertFalse(Files.isRegularFile(Stage3PlanningMetaStore.sidecarPath(dispatchJson)));
        assertFalse(Files.exists(output.resolve("shaped_aladdin_plan.json")));
        assertFalse(Files.exists(output.resolve("計画2606011200000001.xlsx")));
        assertFalse(
                Files.isRegularFile(
                        AppPaths.resolveStage21OutputDir(ui)
                                .resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME)));

        PlanInputTabularIo.TabularSheet stage3 =
                PlanInputTabularIo.read(workbook, PlanInputStage3TabController.STAGE3_SHEET_NAME);
        assertTrue(stage3.rows().isEmpty());
        assertFalse(stage3.headers().isEmpty());
    }
}
