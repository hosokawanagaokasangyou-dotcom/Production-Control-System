package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
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
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK));
        assertTrue(AppPaths.isExcelWorkbookPathEnvKey(AppPaths.KEY_PM_AI_MASTER_WORKBOOK));
        assertTrue(AppPaths.isExcelWorkbookPathEnvKey(AppPaths.KEY_PM_AI_COLUMN_CONFIG_WORKBOOK));
        assertTrue(AppPaths.isExcelWorkbookPathEnvKey(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK));
        assertTrue(AppPaths.isExcelWorkbookPathEnvKey(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK));
        assertTrue(AppPaths.isCsvFilePathEnvKey(AppPaths.KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV));
        assertFalse(AppPaths.isJsonFilePathEnvKey(AppPaths.KEY_PM_AI_MASTER_WORKBOOK));
    }

    @Test
    void planInputAndSidecarPaths_useFilePickerNotFolder() {
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH));
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH));
        assertFalse(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH));
        assertTrue(AppPaths.isPlanInputPathEnvKey(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH));
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH));
        assertFalse(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH));
        assertTrue(AppPaths.isJsonFilePathEnvKey(AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH));
    }

    @Test
    void actualDetailWorkbook_usesFilePickerNotFolder() {
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK));
        assertFalse(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK));
        assertTrue(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR));
    }

    @Test
    void tabularMasterTablePaths_useFilePickerNotFolder() {
        assertTrue(AppPaths.isTabularDataTablePathEnvKey("RAW_FABRIC_WIDTH_TABLE_PATH"));
        assertTrue(AppPaths.isFilePathEnvKey("PRODUCT_THICKNESS_TABLE_PATH"));
        assertFalse(AppPaths.isFolderPathEnvKey("ROLL_UNIT_BY_USED_RAW_TABLE_PATH"));
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
        assertTrue(s.endsWith("生産計画問合せ"), "suffix: " + p);
    }

    @Test
    void actualDetailSourceDir_defaultMatchesPq02UncSuffix() {
        Path p = AppPaths.resolveActualDetailSourceDir(Map.of());
        String s = p.toString().replace('\\', '/');
        assertTrue(s.contains("192.168.0.101"), "host: " + p);
        assertTrue(s.contains("002"), "segment 002  加工G: " + p);
        assertTrue(
                s.endsWith("加工実績明細DATA"),
                "suffix: " + p);
    }

    @Test
    void actualDetailRawMaxBytes_defaultsToTwentyMiB() {
        assertEquals(20L * 1024 * 1024, AppPaths.resolveActualDetailRawMaxBytes(Map.of()));
    }

    @Test
    void actualDetailRawMaxBytes_acceptsSuffixAndZeroUnlimited() {
        assertEquals(
                16L * 1024 * 1024,
                AppPaths.resolveActualDetailRawMaxBytes(
                        Map.of(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES, "16M")));
        assertEquals(
                0L,
                AppPaths.resolveActualDetailRawMaxBytes(
                        Map.of(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES, "0")));
    }

    @Test
    void actualDetailRawMaxBytes_invalidFallsBackToDefault() {
        assertEquals(
                AppPaths.DEFAULT_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES,
                AppPaths.resolveActualDetailRawMaxBytes(
                        Map.of(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES, "not-a-number")));
    }

    @Test
    void parseEnvByteCountToLong_rejectsInvalid() {
        assertTrue(AppPaths.parseEnvByteCountToLong("??") < 0);
    }

    @Test
    void ensureActualDetailRawFileWithinLimit_throwsWhenTooLarge(@TempDir Path dir) throws Exception {
        Path f = dir.resolve("huge.xlsx");
        Files.write(f, new byte[500]);
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES, "100");
        IOException ex =
                assertThrows(IOException.class, () -> AppPaths.ensureActualDetailRawFileWithinLimit(f, ui));
        assertTrue(ex.getMessage().contains("上限"), ex.getMessage());
    }

    @Test
    void ensureActualDetailRawFileWithinLimit_skipsWhenMaxZero(@TempDir Path dir) throws Exception {
        Path f = dir.resolve("any.xlsx");
        Files.write(f, new byte[500]);
        AppPaths.ensureActualDetailRawFileWithinLimit(
                f, Map.of(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES, "0"));
    }

    @Test
    void resolveDefaultExcludeRulesJsonPath_prefersPrimaryThenStage1(@TempDir Path fakeRepo) throws Exception {
        Path code = fakeRepo.resolve("code");
        Files.createDirectories(code.resolve("python"));
        Files.createFile(code.resolve("python").resolve("task_extract_stage1.py"));
        Path jsonDir = code.resolve("json");
        Files.createDirectories(jsonDir);
        Path stage1 = jsonDir.resolve(AppPaths.STAGE1_EXCLUDE_RULES_JSON_FILENAME);
        Files.createFile(stage1);
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());
        assertEquals(
                stage1.toAbsolutePath().normalize(),
                AppPaths.resolveDefaultExcludeRulesJsonPath(ui).get());

        Path primary = code.resolve("exclude_rules.json");
        Files.createFile(primary);
        assertEquals(
                primary.toAbsolutePath().normalize(),
                AppPaths.resolveDefaultExcludeRulesJsonPath(ui).get());
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
                fakeRepo.resolve("Production-Control-System")
                        .resolve("code")
                        .resolve("output")
                        .toAbsolutePath()
                        .normalize();
        assertEquals(expected, AppPaths.resolveResultDispatchTableDir(ui));
    }

    @Test
    void resultDispatchTableDir_usesPlanInputParentWhenExcel(@TempDir Path tmp) throws Exception {
        Path out = tmp.resolve("output");
        Files.createDirectories(out);
        Path xlsm = out.resolve("task.xlsm");
        Files.createFile(xlsm);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        tmp.toString(),
                        AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                        xlsm.toString());
        assertEquals(out.toAbsolutePath().normalize(), AppPaths.resolveResultDispatchTableDir(ui));
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
        Path preferred = dir.resolve("生産管理_AI配台_V2.xlsm");
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
    void summaryAiDispatchXlsmPath_defaultsUnderCode(@TempDir Path fakeRepo) throws Exception {
        Path code = fakeRepo.resolve("code").resolve("python");
        Files.createDirectories(code);
        Files.createFile(code.resolve("task_extract_stage1.py"));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());
        Path expected =
                fakeRepo.resolve("code")
                        .resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSM)
                        .normalize()
                        .toAbsolutePath();
        assertEquals(expected, AppPaths.summaryAiDispatchXlsmPath(ui));
    }

    @Test
    void summaryAiDispatchXlsmPath_respectsOverrideAbsolute(@TempDir Path tmp) throws Exception {
        Path custom = tmp.resolve("custom.xlsm");
        Files.createFile(custom);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, custom.toString());
        assertEquals(custom.toAbsolutePath().normalize(), AppPaths.summaryAiDispatchXlsmPath(ui));
    }

    @Test
    void summaryAiDispatchXlsmPath_respectsOverrideRelativeToCode(@TempDir Path fakeRepo) throws Exception {
        Path code = fakeRepo.resolve("code");
        Files.createDirectories(code.resolve("python"));
        Files.createFile(code.resolve("python").resolve("task_extract_stage1.py"));
        Path alt = code.resolve("alt.xlsm");
        Files.createFile(alt);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        fakeRepo.toString(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        "alt.xlsm");
        assertEquals(alt.toAbsolutePath().normalize(), AppPaths.summaryAiDispatchXlsmPath(ui));
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

    @Test
    void normalizeFolderEnvValue_relativeUnderRepo_becomesAbsolute(@TempDir Path repo) throws Exception {
        Path py = repo.resolve("code").resolve("python");
        Files.createDirectories(py);
        Files.createFile(py.resolve("task_extract_stage1.py"));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repo.toString());
        Optional<String> n =
                AppPaths.normalizeFolderEnvValue(ui, AppPaths.KEY_PM_AI_OUTPUT_DIR, "output");
        assertTrue(n.isPresent());
        assertEquals(repo.resolve("output").normalize().toString(), n.get());
    }

    @Test
    void normalizedFolderEnv_relocatesOldAbsoluteClone(@TempDir Path tmp) throws Exception {
        Path repoNew = tmp.resolve("PCS");
        Files.createDirectories(repoNew.resolve("code").resolve("python"));
        Path legacyAbs = tmp.resolve("somewhere").resolve("PCS").resolve("code").resolve("python");
        Files.createDirectories(legacyAbs);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        repoNew.toString(),
                        AppPaths.KEY_PM_AI_CODE_PYTHON_DIR,
                        legacyAbs.toString());
        Map<String, String> o = AppPaths.normalizedFolderEnvOverrides(ui);
        assertEquals(
                repoNew.resolve("code").resolve("python").normalize().toString(),
                o.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR));
    }

    @Test
    void normalizeFolderEnvValue_escapingRelativeUnchanged(@TempDir Path repo) throws Exception {
        Path py = repo.resolve("code").resolve("python");
        Files.createDirectories(py);
        Files.createFile(py.resolve("task_extract_stage1.py"));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repo.toString());
        Optional<String> n =
                AppPaths.normalizeFolderEnvValue(ui, AppPaths.KEY_PM_AI_OUTPUT_DIR, "../outside");
        assertFalse(n.isPresent());
    }

    @Test
    void normalizePmAiPythonExecutable_folderResolvesToPythonExe(@TempDir Path tmp) throws IOException {
        Path embed = tmp.resolve("python-embed");
        Files.createDirectories(embed);
        Path exe = embed.resolve("python.exe");
        Files.createFile(exe);
        assertEquals(
                exe.toAbsolutePath().normalize().toString(),
                AppPaths.normalizePmAiPythonExecutable(embed.toString()));
    }

    @Test
    void normalizePmAiPythonExecutable_plainExePathUnchanged(@TempDir Path tmp) throws IOException {
        Path exe = tmp.resolve("python.exe");
        Files.createFile(exe);
        String s = exe.toAbsolutePath().normalize().toString();
        assertEquals(s, AppPaths.normalizePmAiPythonExecutable(s));
    }

    @Test
    void normalizePmAiPythonExecutable_folderWithoutInterpreterReturnsEmpty(@TempDir Path tmp)
            throws IOException {
        Path embed = tmp.resolve("python-embed");
        Files.createDirectories(embed);
        assertEquals("", AppPaths.normalizePmAiPythonExecutable(embed.toString()));
    }
}
