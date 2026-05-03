package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class AppPathsTest {

    @Test
    void outputDir_isFolderPathKey() {
        assertTrue(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_OUTPUT_DIR));
        assertFalse(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_OUTPUT_DIR));
    }

    @Test
    void geminiCredentialsJson_usesFilePickerNotFolder() {
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_GEMINI_CREDENTIALS_JSON));
        assertFalse(AppPaths.isFolderPathEnvKey(AppPaths.KEY_GEMINI_CREDENTIALS_JSON));
    }

    @Test
    void excludeRulesJson_usesFilePickerNotFolder() {
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON));
        assertFalse(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON));
    }

    @Test
    void masterAndRelatedPaths_useFilePickerKinds() {
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_MASTER_WORKBOOK));
        assertTrue(AppPaths.isExcelWorkbookPathEnvKey(AppPaths.KEY_PM_AI_MASTER_WORKBOOK));
        assertTrue(AppPaths.isExcelWorkbookPathEnvKey(AppPaths.KEY_PM_AI_COLUMN_CONFIG_WORKBOOK));
        assertTrue(AppPaths.isCsvFilePathEnvKey(AppPaths.KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV));
        assertFalse(AppPaths.isJsonFilePathEnvKey(AppPaths.KEY_PM_AI_MASTER_WORKBOOK));
    }

    @Test
    void resolveMasterWorkbookCandidate_prefersPlanMaster(@TempDir Path fakeRepo) throws Exception {
        Path planMaster = fakeRepo.resolve("plan").resolve("master.xlsm");
        Files.createDirectories(planMaster.getParent());
        Files.createFile(planMaster);
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());
        assertEquals(planMaster.toAbsolutePath().normalize(), AppPaths.resolveMasterWorkbookCandidate(ui).get());
    }

    @Test
    void taskInputSourceDir_defaultMatchesPqAUncSuffix() {
        Path p = AppPaths.resolveTaskInputSourceDir(Map.of());
        String s = p.toString().replace('\\', '/');
        assertTrue(s.contains("192.168.0.101"), "host: " + p);
        assertTrue(s.endsWith("\u751f\u7523\u8a08\u753b\u554f\u5408\u305b"), "suffix: " + p);
    }

    @Test
    void actualDetailSourceDir_defaultMatchesPq02UncSuffix() {
        Path p = AppPaths.resolveActualDetailSourceDir(Map.of());
        String s = p.toString().replace('\\', '/');
        assertTrue(s.contains("192.168.0.101"), "host: " + p);
        assertTrue(s.contains("002"), "segment 002  \u52a0\u5de5G: " + p);
        assertTrue(
                s.endsWith("\u52a0\u5de5\u5b9f\u7e3e\u660e\u7d30DATA"),
                "suffix: " + p);
    }

    @Test
    void resolveDefaultOutputDir_defaultsToRepoOutput(@TempDir Path fakeRepo) throws Exception {
        Path code = fakeRepo.resolve("code").resolve("python");
        Files.createDirectories(code);
        Files.createFile(code.resolve("task_extract_stage1.py"));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());
        assertEquals(
                fakeRepo.resolve("output").toAbsolutePath().normalize(),
                AppPaths.resolveDefaultOutputDir(ui));
    }

    @Test
    void resolveDefaultOutputDir_respectsOverride(@TempDir Path fakeRepo, @TempDir Path out) throws Exception {
        Path code = fakeRepo.resolve("code").resolve("python");
        Files.createDirectories(code);
        Files.createFile(code.resolve("task_extract_stage1.py"));
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        fakeRepo.toString(),
                        AppPaths.KEY_PM_AI_OUTPUT_DIR,
                        out.toString());
        assertEquals(out.toAbsolutePath().normalize(), AppPaths.resolveDefaultOutputDir(ui));
    }

    @Test
    void resultDispatchTableDir_defaultsToRepoCode(@TempDir Path fakeRepo) throws Exception {
        Path code = fakeRepo.resolve("Production-Control-System").resolve("code").resolve("python");
        Files.createDirectories(code);
        Files.createFile(code.resolve("task_extract_stage1.py"));
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.resolve("Production-Control-System").toString());
        Path expected =
                fakeRepo.resolve("Production-Control-System").resolve("code").toAbsolutePath().normalize();
        assertEquals(expected, AppPaths.resolveResultDispatchTableDir(ui));
    }

    @Test
    void resultDispatchTableJsonPath_joinsBasename(@TempDir Path fakeRepo) throws Exception {
        Path code = fakeRepo.resolve("Production-Control-System").resolve("code").resolve("python");
        Files.createDirectories(code);
        Files.createFile(code.resolve("task_extract_stage1.py"));
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.resolve("Production-Control-System").toString());
        Path dir = AppPaths.resolveResultDispatchTableDir(ui);
        Path json = AppPaths.resolveResultDispatchTableJsonPath(ui);
        assertEquals(dir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME), json);
    }

    @Test
    void pickMacroWorkbook_singleFile(@TempDir Path dir) throws Exception {
        Path wb = dir.resolve("only.xlsm");
        Files.createFile(wb);
        assertEquals(Optional.of(wb), AppPaths.pickMacroWorkbook(dir));
    }

    @Test
    void pickMacroWorkbook_prefersNameContainingHaitai(@TempDir Path dir) throws Exception {
        Files.createFile(dir.resolve("other.xlsm"));
        Path preferred = dir.resolve("\u751f\u7523\u7ba1\u7406_AI\u914d\u53f0_V2.xlsm");
        Files.createFile(preferred);
        assertEquals(Optional.of(preferred), AppPaths.pickMacroWorkbook(dir));
    }

    @Test
    void resolveMasterWorkbookPathResolved_usesPmAiMasterWhenFileExists(@TempDir Path tmp) throws Exception {
        Path master = tmp.resolve("m.xlsm");
        Files.createFile(master);
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_MASTER_WORKBOOK, master.toString());
        assertEquals(
                master.toAbsolutePath().normalize(),
                AppPaths.resolveMasterWorkbookPathResolved(ui, ""));
    }

    @Test
    void resolveMasterWorkbookPathResolved_relativeUsesCodeFolder(@TempDir Path tmp) throws Exception {
        Path code = tmp.resolve("code");
        Path py = code.resolve("python");
        Files.createDirectories(py);
        Files.createFile(py.resolve("task_extract_stage1.py"));
        Path master = code.resolve("master.xlsm");
        Files.createFile(master);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        tmp.toString(),
                        AppPaths.KEY_MASTER_WORKBOOK_FILE,
                        "master.xlsm");
        assertEquals(
                master.toAbsolutePath().normalize(),
                AppPaths.resolveMasterWorkbookPathResolved(ui, ""));
    }
}
