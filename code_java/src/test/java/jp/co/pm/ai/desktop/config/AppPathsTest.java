package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
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
    void requestFormJuchuFile_usesFilePickerNotFolder() {
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE));
        assertTrue(AppPaths.isExcelWorkbookPathEnvKey(AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE));
        assertFalse(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE));
    }

    @Test
    void requestFormOriginalDir_usesFolderPickerNotFile() {
        assertTrue(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR));
        assertFalse(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR));
    }

    @Test
    void requestFormRdpProfile_usesFilePickerNotFolder(@TempDir Path tmp) throws Exception {
        assertTrue(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE));
        assertFalse(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE));
        Path rdp = tmp.resolve("remote.rdp");
        Files.writeString(rdp, "screen mode id:i:2");
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_REQUEST_FORM_RDP_PROFILE, rdp.toString());
        assertEquals(
                rdp.toAbsolutePath().normalize(),
                AppPaths.resolveRequestFormRdpProfile(ui).orElseThrow());
    }

    @Test
    void resolveWindowsDefaultRdpProfile_prefersOneDriveJapaneseDocuments(@TempDir Path fakeHome)
            throws Exception {
        Path plainDocs = fakeHome.resolve("Documents");
        Files.createDirectories(plainDocs);
        Path plainDefault = plainDocs.resolve(AppPaths.WINDOWS_DEFAULT_RDP_FILENAME);
        Files.writeString(plainDefault, "screen mode id:i:2");

        Path oneDriveJa =
                fakeHome.resolve("OneDrive").resolve("ドキュメント");
        Files.createDirectories(oneDriveJa);
        Path oneDriveDefault = oneDriveJa.resolve(AppPaths.WINDOWS_DEFAULT_RDP_FILENAME);
        Files.writeString(oneDriveDefault, "screen mode id:i:2");

        assertEquals(
                plainDefault.toAbsolutePath().normalize(),
                AppPaths.resolveWindowsDefaultRdpProfileUnder(fakeHome).orElseThrow());

        Files.delete(plainDefault);
        assertEquals(
                oneDriveDefault.toAbsolutePath().normalize(),
                AppPaths.resolveWindowsDefaultRdpProfileUnder(fakeHome).orElseThrow());
    }

    @Test
    void resolveRequestFormRdpProfile_fallsBackToWindowsDefaultWhenEnvEmpty(@TempDir Path fakeHome)
            throws Exception {
        Path docs = fakeHome.resolve("Documents");
        Files.createDirectories(docs);
        Path rdp = docs.resolve(AppPaths.WINDOWS_DEFAULT_RDP_FILENAME);
        Files.writeString(rdp, "screen mode id:i:2");

        assertEquals(
                rdp.toAbsolutePath().normalize(),
                AppPaths.resolveWindowsDefaultRdpProfileUnder(fakeHome).orElseThrow());
    }

    @Test
    void rdpCompanionProgram_isPlainEnvKeyNotFilePicker() {
        assertFalse(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM));
        assertFalse(AppPaths.isFolderPathEnvKey(AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM));
        assertFalse(AppPaths.isFilePathEnvKey(AppPaths.KEY_PM_AI_RDP_COMPANION_PROGRAM_ARGS));
    }

    @Test
    void resolveRequestFormOriginalDir_usesEnvOverrideOrFactoryDefault(@TempDir Path tmp) throws Exception {
        Path custom = tmp.resolve("original-forms");
        Files.createDirectories(custom);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR, custom.toString());
        assertEquals(
                custom.toAbsolutePath().normalize(),
                AppPaths.resolveRequestFormOriginalDir(ui));

        GlobalInitSettingTarget.save(FactorySite.KONAN);
        assertEquals(
                Path.of(AppPaths.defaultRequestFormOriginalDirForFactory(FactorySite.KONAN))
                        .toAbsolutePath()
                        .normalize(),
                AppPaths.resolveRequestFormOriginalDir(Map.of()));

        GlobalInitSettingTarget.save(FactorySite.KOKUBU);
        assertEquals(
                Path.of(AppPaths.defaultRequestFormOriginalDirForFactory(FactorySite.KOKUBU))
                        .toAbsolutePath()
                        .normalize(),
                AppPaths.resolveRequestFormOriginalDir(Map.of()));
    }

    @Test
    void resolveRequestFormJuchuFile_usesEnvOverrideOrFactoryDefault(@TempDir Path tmp) throws Exception {
        Path book = tmp.resolve("加工依頼書入力.xlsm");
        Files.createFile(book);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE, book.toString());
        assertEquals(book.toAbsolutePath().normalize(), AppPaths.resolveRequestFormJuchuFile(ui).get());

        GlobalInitSettingTarget.save(FactorySite.KONAN);
        assertEquals(
                AppPaths.DEFAULT_PM_AI_REQUEST_FORM_JUCHU_FILE_KONAN,
                AppPaths.resolveRequestFormJuchuFile(Map.of()).get().toString());

        GlobalInitSettingTarget.save(FactorySite.KOKUBU);
        assertEquals(
                AppPaths.DEFAULT_PM_AI_REQUEST_FORM_JUCHU_FILE_KOKUBU,
                AppPaths.resolveRequestFormJuchuFile(Map.of()).get().toString());
    }

    @Test
    void resolveAladdinMasterDir_usesEnvOverrideOrFactoryDefault(@TempDir Path fakeRepo) {
        Path custom = fakeRepo.resolve("custom-aladdin");
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        fakeRepo.toString(),
                        AppPaths.KEY_PM_AI_ALADDIN_MASTER_DIR,
                        custom.toString());
        assertEquals(custom.toAbsolutePath().normalize(), AppPaths.resolveAladdinMasterDir(ui));

        GlobalInitSettingTarget.save(FactorySite.KONAN);
        assertEquals(
                AppPaths.DEFAULT_PM_AI_ALADDIN_MASTER_DIR_KONAN,
                AppPaths.resolveAladdinMasterDir(Map.of()).toString());

        GlobalInitSettingTarget.save(FactorySite.KOKUBU);
        assertEquals(
                AppPaths.DEFAULT_PM_AI_ALADDIN_MASTER_DIR_KOKUBU,
                AppPaths.resolveAladdinMasterDir(Map.of()).toString());
    }

    @Test
    void resolveRdpLauncherPaths_sameDirAsSummary(@TempDir Path fakeRepo) throws IOException {
        Path summary = fakeRepo.resolve("code").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createDirectories(summary.getParent());
        Files.writeString(summary, "x");
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());
        Path deployDir = AppPaths.resolveRdpLauncherDeployDir(ui);
        assertEquals(summary.getParent().normalize(), deployDir.normalize());
        assertEquals(
                deployDir.resolve(AppPaths.RDP_LAUNCHER_INI_BASENAME).normalize(),
                AppPaths.resolveRdpLauncherIni(ui).normalize());
        assertEquals(
                deployDir.resolve(AppPaths.RDP_LAUNCHER_EXE_BASENAME).normalize(),
                AppPaths.resolveRdpLauncherExe(ui).normalize());
    }

    @Test
    void hasSavedRdpLaunchProfileNumber_detectsValidSessionValue() {
        assertFalse(AppPaths.hasSavedRdpLaunchProfileNumber(Map.of()));
        assertFalse(
                AppPaths.hasSavedRdpLaunchProfileNumber(
                        Map.of(AppPaths.KEY_PM_AI_RDP_LAUNCH_PROFILE_NUMBER, "")));
        assertFalse(
                AppPaths.hasSavedRdpLaunchProfileNumber(
                        Map.of(AppPaths.KEY_PM_AI_RDP_LAUNCH_PROFILE_NUMBER, "0")));
        assertTrue(
                AppPaths.hasSavedRdpLaunchProfileNumber(
                        Map.of(AppPaths.KEY_PM_AI_RDP_LAUNCH_PROFILE_NUMBER, "3")));
    }

    @Test
    void resolveRdpLaunchProfileNumber_defaultsWhenMissingOrInvalid() {
        assertEquals(1, AppPaths.resolveRdpLaunchProfileNumber(Map.of()));
        assertEquals(
                1,
                AppPaths.resolveRdpLaunchProfileNumber(
                        Map.of(AppPaths.KEY_PM_AI_RDP_LAUNCH_PROFILE_NUMBER, "x")));
        assertEquals(
                5,
                AppPaths.resolveRdpLaunchProfileNumber(
                        Map.of(AppPaths.KEY_PM_AI_RDP_LAUNCH_PROFILE_NUMBER, "5")));
    }

    @Test
    void defaultAladdinMasterDir_pathsMatchFactorySharedData() {
        assertEquals(
                AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR + "\\" + AppPaths.ALADDIN_MASTER_DIR_LEAF_NAME,
                AppPaths.DEFAULT_PM_AI_ALADDIN_MASTER_DIR_KONAN);
        assertEquals(
                AppPaths.DEFAULT_KOKUBU_DATA_DIR + "\\" + AppPaths.ALADDIN_MASTER_DIR_LEAF_NAME,
                AppPaths.DEFAULT_PM_AI_ALADDIN_MASTER_DIR_KOKUBU);
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
    void resolveDefaultExcludeRulesJsonPath_copiesBundledToSummarySibling(@TempDir Path fakeRepo)
            throws Exception {
        Path code = fakeRepo.resolve("code");
        Files.createDirectories(code.resolve("python"));
        Files.createFile(code.resolve("python").resolve("task_extract_stage1.py"));
        Path jsonDir = code.resolve("json");
        Files.createDirectories(jsonDir);
        Path bundled = jsonDir.resolve(AppPaths.STAGE1_EXCLUDE_RULES_JSON_FILENAME);
        Files.writeString(bundled, "{\"rules\":[]}");
        Path summaryDir = fakeRepo.resolve("shared");
        Files.createDirectories(summaryDir);
        Path summary = summaryDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summary);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        fakeRepo.toString(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        summary.toString());
        Path expected =
                summaryDir
                        .resolve(AppPaths.STAGE1_EXCLUDE_RULES_JSON_FILENAME)
                        .toAbsolutePath()
                        .normalize();
        assertEquals(expected, AppPaths.resolveDefaultExcludeRulesJsonPath(ui).get());
        assertTrue(Files.isRegularFile(expected));
    }

    @Test
    void stage1ExcludeRulesJsonPath_usesSummaryWorkbookParent(@TempDir Path tmp) {
        Path custom = tmp.resolve("shared").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, custom.toString());
        Path expected =
                custom.getParent()
                        .resolve(AppPaths.STAGE1_EXCLUDE_RULES_JSON_FILENAME)
                        .toAbsolutePath()
                        .normalize();
        assertEquals(expected, AppPaths.stage1ExcludeRulesJsonPath(ui));
    }

    @Test
    void ensureStage1ExcludeRulesJsonFromRepoIfMissing_copiesWhenAbsent(@TempDir Path fakeRepo)
            throws Exception {
        Path code = fakeRepo.resolve("code");
        Files.createDirectories(code.resolve("python"));
        Files.createFile(code.resolve("python").resolve("task_extract_stage1.py"));
        Path jsonDir = code.resolve("json");
        Files.createDirectories(jsonDir);
        Path bundled = jsonDir.resolve(AppPaths.STAGE1_EXCLUDE_RULES_JSON_FILENAME);
        Files.writeString(bundled, "{\"rules\":[]}");
        Path summaryDir = fakeRepo.resolve("work");
        Files.createDirectories(summaryDir);
        Path summary = summaryDir.resolve("custom_summary.xlsx");
        Files.createFile(summary);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        fakeRepo.toString(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        summary.toString());
        Path target = AppPaths.stage1ExcludeRulesJsonPath(ui);
        assertFalse(Files.isRegularFile(target));
        assertTrue(AppPaths.ensureStage1ExcludeRulesJsonFromRepoIfMissing(ui));
        assertTrue(Files.isRegularFile(target));
    }

    @Test
    void dispatchLookupTablePath_usesSummaryWorkbookParent(@TempDir Path tmp) {
        Path custom = tmp.resolve("shared").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, custom.toString());
        Path expected =
                custom.getParent()
                        .resolve(AppPaths.DISPATCH_LOOKUP_PRODUCT_THICK)
                        .toAbsolutePath()
                        .normalize();
        assertEquals(expected, AppPaths.dispatchLookupTablePath(ui, AppPaths.DISPATCH_LOOKUP_PRODUCT_THICK));
    }

    @Test
    void ensureDispatchLookupTableFromRepoIfMissing_copiesWhenAbsent(@TempDir Path fakeRepo)
            throws Exception {
        Path code = fakeRepo.resolve("code");
        Files.createDirectories(code);
        Files.writeString(code.resolve(AppPaths.DISPATCH_LOOKUP_PRODUCT_THICK), "製品名,製品厚み\n");
        Path summaryDir = fakeRepo.resolve("work");
        Files.createDirectories(summaryDir);
        Path summary = summaryDir.resolve("custom_summary.xlsx");
        Files.createFile(summary);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        fakeRepo.toString(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        summary.toString());
        Path target = AppPaths.dispatchLookupTablePath(ui, AppPaths.DISPATCH_LOOKUP_PRODUCT_THICK);
        assertFalse(Files.isRegularFile(target));
        assertTrue(
                AppPaths.ensureDispatchLookupTableFromRepoIfMissing(
                        ui, AppPaths.DISPATCH_LOOKUP_PRODUCT_THICK));
        assertTrue(Files.isRegularFile(target));
    }

    @Test
    void overwriteDispatchLookupTableFromRepo_replacesExisting(@TempDir Path fakeRepo) throws Exception {
        Path code = fakeRepo.resolve("code");
        Files.createDirectories(code);
        Files.writeString(code.resolve(AppPaths.DISPATCH_LOOKUP_PRODUCT_THICK), "製品名,製品厚み\nrepo\n");
        Path summaryDir = fakeRepo.resolve("work");
        Files.createDirectories(summaryDir);
        Path summary = summaryDir.resolve("custom_summary.xlsx");
        Files.createFile(summary);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        fakeRepo.toString(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        summary.toString(),
                        AppPaths.KEY_PM_AI_CODE_DIR,
                        code.toString());
        Path target = AppPaths.dispatchLookupTablePath(ui, AppPaths.DISPATCH_LOOKUP_PRODUCT_THICK);
        Files.writeString(target, "製品名,製品厚み\nlocal\n");
        AppPaths.DispatchLookupTableOverwriteResult r =
                AppPaths.overwriteDispatchLookupTableFromRepo(ui, AppPaths.DISPATCH_LOOKUP_PRODUCT_THICK);
        assertTrue(r.success());
        assertEquals("製品名,製品厚み\nrepo\n", Files.readString(target));
    }

    @Test
    void dispatchLookupTableFilenames_listsSixTables() {
        assertEquals(6, AppPaths.dispatchLookupTableFilenames().size());
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
        Files.createDirectories(dir);
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
    void resolveMasterWorkbookPathResolved_pmAiMasterWinsOverOtherBasenames(@TempDir Path tmp)
            throws Exception {
        Path code = tmp.resolve("code");
        Path py = code.resolve("python");
        Files.createDirectories(py);
        Files.createFile(py.resolve("task_extract_stage1.py"));
        Path masterDefault = code.resolve("master.xlsm");
        Path kokubu = code.resolve("国分master.xlsm");
        Files.createFile(masterDefault);
        Files.createFile(kokubu);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        tmp.toString(),
                        AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
                        masterDefault.toString());
        Path p = AppPaths.resolveMasterWorkbookPathResolved(ui, "");
        assertEquals(masterDefault.toAbsolutePath().normalize(), p);
    }

    @Test
    void resolveMasterWorkbookPathForDesktopOpen_findsCodeWhenTaskInputInOutput(@TempDir Path tmp)
            throws Exception {
        Path code = tmp.resolve("code");
        Path py = code.resolve("python");
        Files.createDirectories(py);
        Files.createFile(py.resolve("task_extract_stage1.py"));
        Path out = tmp.resolve("output");
        Files.createDirectories(out);
        Path planInOutput = out.resolve("計画2605122.xlsx");
        Files.createFile(planInOutput);
        Path kokubu = code.resolve("国分master.xlsm");
        Files.createFile(kokubu);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        tmp.toString(),
                        AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
                        "国分master.xlsm");
        Path wrong = AppPaths.resolveMasterWorkbookPathResolved(ui, planInOutput.toString());
        assertEquals(out.resolve("国分master.xlsm").normalize().toAbsolutePath(), wrong);
        assertFalse(Files.isRegularFile(wrong));
        Path fixed =
                AppPaths.resolveMasterWorkbookPathForDesktopOpen(ui, planInOutput.toString());
        assertEquals(kokubu.toAbsolutePath().normalize(), fixed);
        assertTrue(Files.isRegularFile(fixed));
    }

    @Test
    void summaryAiDispatchXlsmPath_defaultsUnderCode(@TempDir Path fakeRepo) throws Exception {
        Path code = fakeRepo.resolve("code").resolve("python");
        Files.createDirectories(code);
        Files.createFile(code.resolve("task_extract_stage1.py"));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());
        Path expected =
                fakeRepo.resolve("code")
                        .resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX)
                        .normalize()
                        .toAbsolutePath();
        assertEquals(expected, AppPaths.summaryAiDispatchXlsxPath(ui));
    }

    @Test
    void summaryAiDispatchXlsmPath_respectsOverrideAbsolute(@TempDir Path tmp) throws Exception {
        Path custom = tmp.resolve("custom.xlsm");
        Files.createFile(custom);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, custom.toString());
        assertEquals(custom.toAbsolutePath().normalize(), AppPaths.summaryAiDispatchXlsxPath(ui));
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
        assertEquals(alt.toAbsolutePath().normalize(), AppPaths.summaryAiDispatchXlsxPath(ui));
    }

    @Test
    void equipmentGanttPdfPath_usesSummaryWorkbookParent(@TempDir Path tmp) {
        Path custom = tmp.resolve("shared").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, custom.toString());
        Path expected =
                custom.getParent()
                        .resolve(AppPaths.EQUIPMENT_GANTT_PDF)
                        .toAbsolutePath()
                        .normalize();
        assertEquals(expected, AppPaths.equipmentGanttPdfPath(ui));
    }

    @Test
    void pipelineExecutionTimingHistoryPath_usesSummaryWorkbookParent(@TempDir Path tmp) throws Exception {
        Path custom = tmp.resolve("shared").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, custom.toString());
        Path expected =
                custom.getParent()
                        .resolve(AppPaths.PIPELINE_EXECUTION_TIMING_HISTORY_JSON)
                        .toAbsolutePath()
                        .normalize();
        assertEquals(expected, AppPaths.pipelineExecutionTimingHistoryPath(ui));
    }

    @Test
    void summaryAiDispatchXlsxPathForFactory_usesFactoryDefaultWhenEnvPointsToOtherFactory(@TempDir Path tmp)
            throws Exception {
        Path konanSummary =
                tmp.resolve("湖南工場").resolve("共有DATA").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createDirectories(konanSummary.getParent());
        Files.createFile(konanSummary);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, konanSummary.toString());
        Path kokubuSummary =
                AppPaths.summaryAiDispatchXlsxPathForFactory(ui, FactorySite.KOKUBU);
        assertTrue(
                kokubuSummary.toString().contains("国分"),
                "国分既定サマリへ切替: " + kokubuSummary);
        assertEquals(
                konanSummary.toAbsolutePath().normalize(),
                AppPaths.summaryAiDispatchXlsxPathForFactory(ui, FactorySite.KONAN));
    }

    @Test
    void factoryOperatorUsersStorePath_usesEffectiveFactoryDataDirWhenSummaryMismatch(@TempDir Path tmp)
            throws Exception {
        Path konanSummary =
                tmp.resolve("湖南工場").resolve("共有DATA").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createDirectories(konanSummary.getParent());
        Files.createFile(konanSummary);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, konanSummary.toString());
        Path store = AppPaths.factoryOperatorUsersStorePath(ui, FactorySite.KOKUBU);
        Path konanBin =
                konanSummary.getParent()
                        .resolve(AppPaths.FACTORY_OPERATOR_USERS_BIN)
                        .toAbsolutePath()
                        .normalize();
        assertNotEquals(konanBin, store, "湖南サマリ配下ではなく国分側 bin を参照");
        assertEquals(AppPaths.FACTORY_OPERATOR_USERS_BIN, store.getFileName().toString());
    }

    @Test
    void factoryOperatorUsersStorePath_usesSummaryWorkbookParent(@TempDir Path tmp) throws Exception {
        Path custom = tmp.resolve("shared").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, custom.toString());
        Path expected =
                custom.getParent()
                        .resolve(AppPaths.FACTORY_OPERATOR_USERS_BIN)
                        .toAbsolutePath()
                        .normalize();
        assertEquals(expected, AppPaths.factoryOperatorUsersStorePath(ui));
    }

    @Test
    void factoryOperatorUsersBackupsRoot_usesSummaryWorkbookParent(@TempDir Path tmp) throws Exception {
        Path custom = tmp.resolve("shared").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, custom.toString());
        Path expected =
                custom.getParent()
                        .resolve(AppPaths.FACTORY_OPERATOR_USERS_BACKUPS_DIR)
                        .toAbsolutePath()
                        .normalize();
        assertEquals(expected, AppPaths.factoryOperatorUsersBackupsRoot(ui));
    }

    @Test
    void factoryOperatorUsersPdfPath_usesSummaryWorkbookParent(@TempDir Path tmp) {
        Path custom = tmp.resolve("shared").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, custom.toString());
        Path expected =
                custom.getParent()
                        .resolve(AppPaths.factoryOperatorUsersPdfFileName(FactorySite.KONAN))
                        .toAbsolutePath()
                        .normalize();
        assertEquals(expected, AppPaths.factoryOperatorUsersPdfPath(ui, FactorySite.KONAN));
    }

    @Test
    void migrateLegacyMasterWorkbookFileToPmAi_resolvesRelativeBasename(@TempDir Path tmp) throws Exception {
        Path code = tmp.resolve("code");
        Path py = code.resolve("python");
        Files.createDirectories(py);
        Files.createFile(py.resolve("task_extract_stage1.py"));
        Path kokubu = code.resolve("国分master.xlsm");
        Files.createFile(kokubu);
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, tmp.toString());
        Optional<String> migrated =
                AppPaths.migrateLegacyMasterWorkbookFileToPmAi(ui, "国分master.xlsm");
        assertTrue(migrated.isPresent());
        assertEquals(kokubu.toAbsolutePath().normalize().toString(), migrated.get());
    }

    @Test
    void migrateLegacyMasterWorkbookFileToPmAi_skipsWhenPmAiAlreadySet(@TempDir Path tmp) throws Exception {
        Path master = tmp.resolve("m.xlsm");
        Files.createFile(master);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
                        master.toString(),
                        AppPaths.KEY_MASTER_WORKBOOK_FILE,
                        "国分master.xlsm");
        assertTrue(AppPaths.migrateLegacyMasterWorkbookFileToPmAi(ui, "国分master.xlsm").isEmpty());
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
                        AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
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

    @Test
    void findPortablePythonEmbedExecutable_walksUpFromNestedDir(@TempDir Path root) throws IOException {
        Path install = root.resolve("PortableApp");
        Path exe =
                install.resolve("pm-ai-data")
                        .resolve("runtime")
                        .resolve("python-embed")
                        .resolve("python.exe");
        Files.createDirectories(exe.getParent());
        Files.createFile(exe);
        Path nested = install.resolve("launcher").resolve("bin");
        Files.createDirectories(nested);
        assertEquals(
                exe.toAbsolutePath().normalize(),
                AppPaths.findPortablePythonEmbedExecutable(nested).orElseThrow());
    }

    /** 親が 12 段あるケース（既定の上位探索幅ギリギリ）。 */
    @Test
    void findPortablePythonEmbedExecutable_walksUpElevenNestedDirs(@TempDir Path root) throws IOException {
        Path install = root.resolve("bundleRoot");
        Path exe =
                install.resolve("pm-ai-data")
                        .resolve("runtime")
                        .resolve("python-embed")
                        .resolve("python.exe");
        Files.createDirectories(exe.getParent());
        Files.createFile(exe);
        Path deep = install;
        for (int i = 0; i < 11; i++) {
            deep = deep.resolve("n" + i);
        }
        Files.createDirectories(deep);
        assertEquals(
                exe.toAbsolutePath().normalize(),
                AppPaths.findPortablePythonEmbedExecutable(deep).orElseThrow());
    }

    @Test
    void findPortablePythonEmbedExecutable_missingReturnsEmpty(@TempDir Path tmp) {
        assertTrue(AppPaths.findPortablePythonEmbedExecutable(tmp.resolve("no_embed_here")).isEmpty());
    }

    @Test
    void resolveManualIndexHtml_underRepoRoot(@TempDir Path fakeRepo) {
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());
        assertEquals(
                fakeRepo.resolve(AppPaths.MANUAL_INDEX_HTML_REL).toAbsolutePath().normalize(),
                AppPaths.resolveManualIndexHtml(ui));
    }

    @Test
    void resolveDispatchUsageGuideDocx_underRepoRoot(@TempDir Path fakeRepo) {
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());
        assertEquals(
                fakeRepo.resolve(AppPaths.DISPATCH_USAGE_GUIDE_DOCX).toAbsolutePath().normalize(),
                AppPaths.resolveDispatchUsageGuideDocx(ui));
    }

    @Test
    void resolveDispatchRulesHtml_underRepoRoot(@TempDir Path fakeRepo) {
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());
        assertEquals(
                fakeRepo.resolve(AppPaths.DISPATCH_RULES_HTML_REL).toAbsolutePath().normalize(),
                AppPaths.resolveDispatchRulesHtml(ui));
    }

    @Test
    void resolveRequestFormPreviewPdfCjkScale_clampsAndDefaults() {
        assertEquals(
                0.72f,
                AppPaths.resolveRequestFormPreviewPdfCjkScale(Map.of()),
                0.001f);
        assertEquals(
                0.68f,
                AppPaths.resolveRequestFormPreviewPdfCjkScale(
                        Map.of(AppPaths.KEY_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE, "0.68")),
                0.001f);
        assertEquals(
                0.50f,
                AppPaths.resolveRequestFormPreviewPdfCjkScale(
                        Map.of(AppPaths.KEY_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE, "0.2")),
                0.001f);
        assertEquals(
                1.00f,
                AppPaths.resolveRequestFormPreviewPdfCjkScale(
                        Map.of(AppPaths.KEY_PM_AI_REQUEST_FORM_PREVIEW_PDF_CJK_SCALE, "1.5")),
                0.001f);
    }
}
