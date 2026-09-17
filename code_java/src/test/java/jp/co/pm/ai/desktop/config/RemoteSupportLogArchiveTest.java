package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RemoteSupportLogArchiveTest {

    @Test
    void stageIdForMainShellScript_mapsKnownStages() {
        assertEquals(
                "stage1",
                RemoteSupportLogArchive.stageIdForMainShellScript(
                        "task_extract_stage1.py",
                        "task_extract_stage1.py",
                        "plan_simulation_stage2.py",
                        "plan_simulation_stage2_1.py"));
        assertEquals(
                "stage2",
                RemoteSupportLogArchive.stageIdForMainShellScript(
                        "plan_simulation_stage2.py",
                        "task_extract_stage1.py",
                        "plan_simulation_stage2.py",
                        "plan_simulation_stage2_1.py"));
        assertEquals(
                "stage2.1",
                RemoteSupportLogArchive.stageIdForMainShellScript(
                        "plan_simulation_stage2_1.py",
                        "task_extract_stage1.py",
                        "plan_simulation_stage2.py",
                        "plan_simulation_stage2_1.py"));
        assertEquals(
                null,
                RemoteSupportLogArchive.stageIdForMainShellScript(
                        "other.py",
                        "task_extract_stage1.py",
                        "plan_simulation_stage2.py",
                        "plan_simulation_stage2_1.py"));
    }

    @Test
    void isGenerationExpired_byFolderDatePrefix() {
        LocalDate today = LocalDate.of(2026, 7, 17);
        assertFalse(
                RemoteSupportLogArchive.isGenerationExpired(
                        "20260715-120000_stage2", today, 3));
        assertTrue(
                RemoteSupportLogArchive.isGenerationExpired(
                        "20260713-235959_stage1", today, 3));
        assertFalse(
                RemoteSupportLogArchive.isGenerationExpired("not-a-generation", today, 3));
    }

    @Test
    void archiveAfterStage_writesUiAndMetaAndPrunes(@TempDir Path temp) throws Exception {
        Path summary = temp.resolve("サマリ_AI配台.xlsx");
        Files.writeString(summary, "x", StandardCharsets.UTF_8);
        Path codeLog = temp.resolve("code").resolve("log");
        Files.createDirectories(codeLog);
        Path execLog = codeLog.resolve(AppPaths.EXECUTION_LOG_TXT);
        Files.writeString(execLog, "python-log-line\n", StandardCharsets.UTF_8);

        Path out = temp.resolve("output");
        Files.createDirectories(out);
        Path dispatchJson = out.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME);
        Files.writeString(dispatchJson, "{\"rows\":[]}\n", StandardCharsets.UTF_8);

        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summary.toString());
        ui.put(AppPaths.KEY_PM_AI_REPO_ROOT, temp.toString());
        ui.put(AppPaths.KEY_PM_AI_CODE_DIR, temp.resolve("code").toString());
        ui.put(AppPaths.KEY_PM_AI_OUTPUT_DIR, out.toString());
        ui.put(AppPaths.KEY_PM_AI_REMOTE_LOG, "1");

        Path keep =
                RemoteSupportLogArchive.archiveAfterStage(
                        ui,
                        "細川",
                        "stage2",
                        0,
                        null,
                        "ui line 1\nui line 2\n",
                        LocalDateTime.of(2026, 7, 17, 10, 41, 8));
        assertTrue(Files.isDirectory(keep));
        assertTrue(Files.isRegularFile(keep.resolve(RemoteSupportLogArchive.UI_RUN_LOG_FILENAME)));
        assertTrue(Files.isRegularFile(keep.resolve(AppPaths.EXECUTION_LOG_TXT)));
        assertTrue(Files.isRegularFile(keep.resolve(RemoteSupportLogArchive.META_JSON_FILENAME)));
        assertTrue(
                Files.isRegularFile(keep.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME)));
        String uiText =
                Files.readString(
                        keep.resolve(RemoteSupportLogArchive.UI_RUN_LOG_FILENAME),
                        StandardCharsets.UTF_8);
        assertTrue(uiText.contains("ui line 1"));
        String execText =
                Files.readString(
                        keep.resolve(AppPaths.EXECUTION_LOG_TXT), StandardCharsets.UTF_8);
        assertTrue(execText.contains("python-log-line"));

        Path userDir = keep.getParent();
        Path oldGen = userDir.resolve("20260710-090000_stage1");
        Files.createDirectories(oldGen);
        Files.writeString(oldGen.resolve("old.txt"), "old", StandardCharsets.UTF_8);
        List<Path> removed =
                RemoteSupportLogArchive.pruneExpiredGenerations(
                        userDir, LocalDate.of(2026, 7, 17), 3);
        assertEquals(1, removed.size());
        assertFalse(Files.exists(oldGen));
        assertTrue(Files.isDirectory(keep));
    }

    @Test
    void isEnabled_respectsOffFlag() {
        assertTrue(RemoteSupportLogArchive.isEnabled(Map.of()));
        assertFalse(
                RemoteSupportLogArchive.isEnabled(
                        Map.of(AppPaths.KEY_PM_AI_REMOTE_LOG, "0")));
        assertFalse(
                RemoteSupportLogArchive.isEnabled(
                        Map.of(AppPaths.KEY_PM_AI_REMOTE_LOG, "off")));
    }

    @Test
    void resolveRemoteLogRoot_usesKonanSharedDataWhenFactoryIsKonan() {
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_FACTORY_SITE,
                        FactorySite.KONAN.name(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR);
        assertEquals(
                Path.of(AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR, AppPaths.REMOTE_LOG_DIR_NAME)
                        .toAbsolutePath()
                        .normalize(),
                AppPaths.resolveRemoteLogRoot(ui));
        assertEquals(
                Path.of(
                                AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR,
                                AppPaths.OPERATOR_ACTION_LOG_DIR_NAME)
                        .toAbsolutePath()
                        .normalize(),
                AppPaths.resolveOperatorActionLogRoot(ui));
    }

    @Test
    void resolveRemoteLogRoot_usesKokubuDataWhenFactoryIsKokubu() {
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_FACTORY_SITE,
                        FactorySite.KOKUBU.name(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        AppPaths.DEFAULT_KOKUBU_DATA_DIR);
        assertEquals(
                Path.of(AppPaths.DEFAULT_KOKUBU_DATA_DIR, AppPaths.REMOTE_LOG_DIR_NAME)
                        .toAbsolutePath()
                        .normalize(),
                AppPaths.resolveRemoteLogRoot(ui));
        assertEquals(
                Path.of(AppPaths.DEFAULT_KOKUBU_DATA_DIR, AppPaths.OPERATOR_ACTION_LOG_DIR_NAME)
                        .toAbsolutePath()
                        .normalize(),
                AppPaths.resolveOperatorActionLogRoot(ui));
    }

    @Test
    void resolveRemoteLogRoot_kokubuFactoryIgnoresKonanSummaryOverride() {
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_FACTORY_SITE,
                        FactorySite.KOKUBU.name(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR);
        assertEquals(
                Path.of(AppPaths.DEFAULT_KOKUBU_DATA_DIR, AppPaths.REMOTE_LOG_DIR_NAME)
                        .toAbsolutePath()
                        .normalize(),
                AppPaths.resolveRemoteLogRoot(ui));
    }

    @Test
    void copiesPipelineArtifacts_onlyForDispatchStages() {
        assertTrue(RemoteSupportLogArchive.copiesPipelineArtifacts("stage1"));
        assertTrue(RemoteSupportLogArchive.copiesPipelineArtifacts("stage2"));
        assertTrue(RemoteSupportLogArchive.copiesPipelineArtifacts("stage2.1"));
        assertFalse(RemoteSupportLogArchive.copiesPipelineArtifacts("kouchin"));
        assertFalse(RemoteSupportLogArchive.copiesPipelineArtifacts("session"));
        assertFalse(RemoteSupportLogArchive.copiesPipelineArtifacts("runtime"));
    }

    @Test
    void isDiagnosticRuntimeLine_detectsKouchinAndFailures() {
        assertTrue(
                RemoteSupportLogArchive.isDiagnosticRuntimeLine(
                        "[kouchin] ドロップされたファイルがありません（Outlookの添付はファイルとしてドロップしてください）"));
        assertTrue(
                RemoteSupportLogArchive.isDiagnosticRuntimeLine(
                        "[kouchin] 検証C 月次実績① 読取不可"));
        assertTrue(
                RemoteSupportLogArchive.isDiagnosticRuntimeLine(
                        "[kouchin-trend] 失敗: ロック"));
        assertTrue(
                RemoteSupportLogArchive.isDiagnosticRuntimeLine(
                        "java.io.IOException: Access is denied"));
        assertFalse(
                RemoteSupportLogArchive.isDiagnosticRuntimeLine(
                        "[remote_log] 保存: \\\\share\\remote_log\\古家"));
        assertFalse(RemoteSupportLogArchive.isDiagnosticRuntimeLine("段階2を開始します"));
    }

    @Test
    void archiveAfterStage_kouchinSkipsPipelineArtifactsAndWritesMeta(@TempDir Path temp)
            throws Exception {
        Map<String, String> ui = archiveUi(temp);
        Path keep =
                RemoteSupportLogArchive.archiveAfterStage(
                        ui,
                        "古家",
                        "kouchin",
                        1,
                        "ドロップ空",
                        "[kouchin] ドロップされたファイルがありません\n",
                        LocalDateTime.of(2026, 9, 18, 7, 30, 0));
        assertTrue(Files.isDirectory(keep));
        assertTrue(Files.isRegularFile(keep.resolve(RemoteSupportLogArchive.UI_RUN_LOG_FILENAME)));
        assertFalse(Files.exists(keep.resolve(AppPaths.EXECUTION_LOG_TXT)));
        assertFalse(Files.exists(keep.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME)));
        String meta =
                Files.readString(
                        keep.resolve(RemoteSupportLogArchive.META_JSON_FILENAME),
                        StandardCharsets.UTF_8);
        assertTrue(meta.contains("\"stage_id\" : \"kouchin\""));
        assertTrue(meta.contains("java_io_tmpdir"));
        assertTrue(meta.contains("os_user"));
        assertTrue(meta.contains("app_version"));
    }

    @Test
    void appendDailyUiLog_writesPerOperatorFileAndPrunes(@TempDir Path temp) throws Exception {
        Map<String, String> ui = archiveUi(temp);
        Path file =
                RemoteSupportLogArchive.appendDailyUiLog(
                        ui,
                        "古家",
                        "[kouchin] ドロップされたファイルがありません\n",
                        LocalDate.of(2026, 9, 18));
        assertTrue(Files.isRegularFile(file));
        assertEquals(
                AppPaths.resolveRemoteLogRoot(ui)
                        .resolve(OperatorUserPaths.sanitizeOperatorDirName("古家"))
                        .resolve(RemoteSupportLogArchive.UI_DAILY_DIR_NAME)
                        .resolve("2026-09-18.txt"),
                file.toAbsolutePath().normalize());
        String text = Files.readString(file, StandardCharsets.UTF_8);
        assertTrue(text.contains("factory="));
        assertTrue(text.contains("[kouchin] ドロップされたファイルがありません"));

        Path userDaily = file.getParent();
        Path oldFile = userDaily.resolve("2026-09-14.txt");
        Files.writeString(oldFile, "old\n", StandardCharsets.UTF_8);
        List<Path> removed =
                RemoteSupportLogArchive.pruneExpiredDailyUiLogs(
                        userDaily, LocalDate.of(2026, 9, 18), 3);
        assertEquals(1, removed.size());
        assertFalse(Files.exists(oldFile));
        assertTrue(Files.isRegularFile(file));
    }

    private static Map<String, String> archiveUi(Path temp) throws Exception {
        Path summary = temp.resolve("サマリ_AI配台.xlsx");
        Files.writeString(summary, "x", StandardCharsets.UTF_8);
        Path codeLog = temp.resolve("code").resolve("log");
        Files.createDirectories(codeLog);
        Files.writeString(
                codeLog.resolve(AppPaths.EXECUTION_LOG_TXT),
                "python-log-line\n",
                StandardCharsets.UTF_8);
        Path out = temp.resolve("output");
        Files.createDirectories(out);
        Files.writeString(
                out.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME),
                "{\"rows\":[]}\n",
                StandardCharsets.UTF_8);
        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summary.toString());
        ui.put(AppPaths.KEY_PM_AI_REPO_ROOT, temp.toString());
        ui.put(AppPaths.KEY_PM_AI_CODE_DIR, temp.resolve("code").toString());
        ui.put(AppPaths.KEY_PM_AI_OUTPUT_DIR, out.toString());
        ui.put(AppPaths.KEY_PM_AI_REMOTE_LOG, "1");
        return ui;
    }
}
