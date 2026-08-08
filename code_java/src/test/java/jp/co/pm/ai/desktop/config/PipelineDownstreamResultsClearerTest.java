package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.dispatch.Stage21TrialSnapshotStore;

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
    void clearStage2Downstream_removesDispatchSidecarsAndStage2Artifacts() throws Exception {
        Path repo = temp.resolve("repo");
        Path python = repo.resolve("code").resolve("python");
        Files.createDirectories(python);
        Files.writeString(python.resolve("task_extract_stage1.py"), "# stub\n");

        Map<String, String> ui = ui(repo);
        Path output = AppPaths.resolveResultDispatchTableDir(ui);
        Files.createDirectories(output);
        Path dispatchJson = output.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME);
        Files.writeString(dispatchJson, "{\"rows\":[]}");
        Files.writeString(
                Stage21TrialSnapshotStore.sidecarPathFor(dispatchJson),
                "{\"stage21_applied\":true}");
        Files.writeString(output.resolve("shaped_aladdin_plan.json"), "{}");
        Files.writeString(output.resolve("計画2606011200000001.xlsx"), "stub");
        Files.createDirectories(AppPaths.resolveStage21OutputDir(ui));
        Files.writeString(
                AppPaths.resolveStage21OutputDir(ui).resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME),
                "{}");

        PipelineDownstreamResultsClearer.ClearResult result =
                PipelineDownstreamResultsClearer.clearStage2Downstream(ui);

        assertTrue(result.anyDeleted());
        assertFalse(Files.exists(dispatchJson));
        assertFalse(Files.isRegularFile(Stage21TrialSnapshotStore.sidecarPathFor(dispatchJson)));
        assertFalse(Files.exists(output.resolve("shaped_aladdin_plan.json")));
        assertFalse(Files.exists(output.resolve("計画2606011200000001.xlsx")));
        assertFalse(
                Files.isRegularFile(
                        AppPaths.resolveStage21OutputDir(ui)
                                .resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME)));
    }

    @Test
    void clearStage2Downstream_canPreserveTodayDispatchSourceBundle() throws Exception {
        Path repo = temp.resolve("repo2");
        Path python = repo.resolve("code").resolve("python");
        Files.createDirectories(python);
        Files.writeString(python.resolve("task_extract_stage1.py"), "# stub\n");

        Map<String, String> ui = ui(repo);
        Files.createDirectories(AppPaths.resolveResultDispatchTableDir(ui));
        Path bundle =
                jp.co.pm.ai.planning.stage2.source.Stage1SourceBundleIo.defaultCachePath(ui);
        Files.createDirectories(bundle.getParent());
        Files.writeString(bundle, "{\"version\":1}");

        PipelineDownstreamResultsClearer.clearStage2Downstream(ui, true);

        assertTrue(Files.isRegularFile(bundle));
    }
}
