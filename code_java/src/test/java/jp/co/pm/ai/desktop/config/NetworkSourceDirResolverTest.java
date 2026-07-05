package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class NetworkSourceDirResolverTest {

    @Test
    void resolve_taskInputFromCache_falseWhenLiveNetworkFileFound(@TempDir Path fakeRepo) throws Exception {
        Path src = fakeRepo.resolve("in").resolve("src");
        Files.createDirectories(src);
        Path live = src.resolve("plan.csv");
        Files.writeString(live, "x");
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        fakeRepo.toString(),
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        src.toString());
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(ui);
        assertTrue(r.taskInputPath().isPresent());
        assertFalse(
                r.taskInputFromCache(),
                "network dir reachable and live file found → not a cache fallback");
    }

    @Test
    void taskInputSourceDir_reachable_whenUnderRepo(@TempDir Path fakeRepo) throws Exception {
        Path src = fakeRepo.resolve("in").resolve("src");
        Files.createDirectories(src);
        Files.writeString(src.resolve("a.csv"), "x");
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        fakeRepo.toString(),
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        src.toString());
        assertTrue(NetworkSourceDirResolver.isTaskInputSourceDirReachable(ui));
    }

    @Test
    void taskInputSourceDir_unreachable_whenDirMissing() {
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR, "/no/such/path/pm_ai_task_input_" + System.nanoTime());
        assertFalse(NetworkSourceDirResolver.isTaskInputSourceDirReachable(ui));
    }

    @Test
    void requestFormOriginalDir_reachable_whenPresent(@TempDir Path dir) throws Exception {
        Files.writeString(dir.resolve("sample.xlsm"), "x");
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR, dir.toString());
        assertTrue(NetworkSourceDirResolver.isRequestFormOriginalDirReachable(ui));
    }

    @Test
    void requestFormOriginalDir_unreachable_whenMissing() {
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR,
                        "/no/such/request-form-original-" + System.nanoTime());
        assertFalse(NetworkSourceDirResolver.isRequestFormOriginalDirReachable(ui));
    }

    @Test
    void requestFormTpiPdfDir_unreachable_whenUnset() {
        GlobalInitSettingTarget.save(FactorySite.KOKUBU);
        assertFalse(NetworkSourceDirResolver.isRequestFormTpiPdfDirReachable(Map.of()));
    }

    @Test
    void requestFormTpiPdfDir_reachable_whenPresent(@TempDir Path dir) throws Exception {
        Files.writeString(dir.resolve("sample.pdf"), "x");
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_REQUEST_FORM_TPI_PDF_DIR, dir.toString());
        assertTrue(NetworkSourceDirResolver.isRequestFormTpiPdfDirReachable(ui));
    }

    @Test
    void pruneSiblingCacheFiles_removesOldExtension(@TempDir Path root) throws Exception {
        Files.createDirectories(root);
        Path staleCsv = root.resolve("task-input-newest.csv");
        Path keepXlsx = root.resolve("task-input-newest.xlsx");
        Files.writeString(staleCsv, "old");
        Files.writeString(keepXlsx, "new");

        NetworkSourceDirResolver.pruneSiblingCacheFiles(
                root, "task-input-newest", keepXlsx.getFileName().toString());

        assertFalse(Files.exists(staleCsv));
        assertTrue(Files.exists(keepXlsx));
    }
}
