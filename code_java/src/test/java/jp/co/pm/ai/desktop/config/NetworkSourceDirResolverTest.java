package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class NetworkSourceDirResolverTest {

    private String priorAppHome;
    private String priorUserHome;

    @BeforeEach
    void setUp(@TempDir Path tmp) {
        priorAppHome = AppPaths.desktopAppHomeDirName();
        priorUserHome = System.getProperty("user.home");
        System.setProperty("user.home", tmp.toString());
        AppPaths.setDesktopAppHomeDirName(".pm-ai-desktop-test");
        GlobalInitSettingTarget.save(FactorySite.KONAN);
    }

    @AfterEach
    void tearDown() {
        AppPaths.setDesktopAppHomeDirName(priorAppHome);
        System.setProperty("user.home", priorUserHome);
    }

    @Test
    void resolve_taskInputFromCache_falseWhenLiveNetworkFileFound(@TempDir Path fakeRepo) throws Exception {
        Path src = fakeRepo.resolve("in").resolve("src");
        Files.createDirectories(src);
        Path live = src.resolve("plan.xlsx");
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
        Files.writeString(src.resolve("a.xlsx"), "x");
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
        GlobalInitSettingTarget.save(FactorySite.KONAN);
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

    @Test
    void copyLiveFileToCache_rewritesOfficeZipWithReadableCentralDirectory(@TempDir Path root)
            throws Exception {
        Path live = root.resolve("live.xlsx");
        Path dest = root.resolve("task-input-newest.xlsx");
        try (OutputStream out = Files.newOutputStream(live);
                ZipOutputStream zout = new ZipOutputStream(out)) {
            zout.putNextEntry(new ZipEntry("[Content_Types].xml"));
            zout.write("<Types/>".getBytes(StandardCharsets.UTF_8));
            zout.closeEntry();
            zout.putNextEntry(new ZipEntry("xl/workbook.xml"));
            zout.write("<workbook/>".getBytes(StandardCharsets.UTF_8));
            zout.closeEntry();
        }

        List<String> logs = new ArrayList<>();
        NetworkSourceDirResolver.copyLiveFileToCache(live, dest, logs);

        assertTrue(Files.isRegularFile(dest));
        try (ZipFile zip = new ZipFile(dest.toFile())) {
            assertEquals(2, zip.size());
            assertTrue(zip.getEntry("[Content_Types].xml") != null);
            assertTrue(zip.getEntry("xl/workbook.xml") != null);
        }
        assertTrue(logs.stream().noneMatch(line -> line.contains("素コピー")));
    }
}
