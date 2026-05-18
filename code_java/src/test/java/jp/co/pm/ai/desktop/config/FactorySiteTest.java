package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.assertFalse;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class FactorySiteTest {

    @Test
    void konanMatchesAppPathsDefaults() {
        assertEquals(AppPaths.DEFAULT_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR, FactorySite.KONAN.portableBundleSourceDir());
        assertEquals(AppPaths.DEFAULT_PM_AI_TASK_INPUT_SOURCE_DIR, FactorySite.KONAN.taskInputSourceDir());
        assertEquals(AppPaths.DEFAULT_PM_AI_ACTUAL_DETAIL_SOURCE_DIR, FactorySite.KONAN.actualDetailSourceDir());
        assertEquals("", FactorySite.KONAN.masterWorkbookFileBasename());
    }

    @Test
    void kokubuUsesKokubuNetworkLayout() {
        assertTrue(FactorySite.KOKUBU.portableBundleSourceDir().contains("国分工場"));
        assertEquals(AppPaths.DEFAULT_PM_AI_PORTABLE_BUNDLE_RELEASE_DIR_KOKUBU, FactorySite.KOKUBU.portableBundleSourceDir());
        assertTrue(FactorySite.KOKUBU.portableBundleSourceDir().endsWith("pm-ai-package-release"));
        assertTrue(FactorySite.KOKUBU.taskInputSourceDir().contains("DATA\\計画"));
        assertTrue(FactorySite.KOKUBU.actualDetailSourceDir().contains("DATA\\実績"));
        assertEquals("国分master.xlsm", FactorySite.KOKUBU.masterWorkbookFileBasename());
    }

    @Test
    void toStringUsesJapaneseLabel() {
        assertEquals("湖南工場", FactorySite.KONAN.toString());
        assertEquals("国分工場", FactorySite.KOKUBU.toString());
    }

    @Test
    void konanPmAiMasterWorkbookEnvValueUsesSharedDataUnc() {
        assertEquals(
                AppPaths.DEFAULT_PM_AI_MASTER_WORKBOOK_KONAN,
                FactorySite.KONAN.pmAiMasterWorkbookEnvValue(Map.of()));
    }

    @Test
    void kokubuPmAiMasterWorkbookEnvValueUnderRepoCode(@TempDir Path repo) throws Exception {
        Files.createDirectories(repo.resolve("code"));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repo.toString());
        assertEquals(
                repo.resolve("code")
                        .resolve("国分master.xlsm")
                        .normalize()
                        .toAbsolutePath()
                        .toString(),
                FactorySite.KOKUBU.pmAiMasterWorkbookEnvValue(ui));
    }

    @Test
    void konanPmAiSummaryDispatchUsesSharedDataUnc() {
        assertEquals(
                AppPaths.DEFAULT_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK_KONAN,
                FactorySite.KONAN.pmAiSummaryAiDispatchWorkbookEnvValue(Map.of()));
    }

    @Test
    void kokubuPmAiSummaryDispatchUnderRepoCode(@TempDir Path repo) throws Exception {
        Files.createDirectories(repo.resolve("code"));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repo.toString());
        assertEquals(
                repo.resolve("code")
                        .resolve("国分サマリ_AI配台.xlsx")
                        .normalize()
                        .toAbsolutePath()
                        .toString(),
                FactorySite.KOKUBU.pmAiSummaryAiDispatchWorkbookEnvValue(ui));
    }

    @Test
    void kokubuSummaryFilenameConstant() {
        assertEquals("国分サマリ_AI配台.xlsx", AppPaths.KOKUBU_SUMMARY_AI_DISPATCH_WORKBOOK_XLSX);
    }

    @Test
    void inferFromPortableBundleSourceValue_detectsKokubuAndKonan() {
        assertEquals(
                FactorySite.KOKUBU,
                FactorySite.inferFromPortableBundleSourceValue(
                                "\\\\host\\国分工場\\国分共有\\pm-ai-package-release")
                        .orElseThrow());
        assertEquals(
                FactorySite.KOKUBU,
                FactorySite.inferFromPortableBundleSourceValue(
                                "\\\\host\\国分工場\\国分共有\\pm-ai-package-release\\PMD_version_upgrade.zip")
                        .orElseThrow());
        assertEquals(
                FactorySite.KONAN,
                FactorySite.inferFromPortableBundleSourceValue(
                                "\\\\host\\湖南工場\\湖南共有\\pm-ai-package-release")
                        .orElseThrow());
        assertFalse(FactorySite.inferFromPortableBundleSourceValue("").isPresent());
    }

    @Test
    void inferFromPortableBundleInitSetting_readsBundledSessionDefaults(@TempDir Path installRoot)
            throws Exception {
        Path initDir = installRoot.resolve("pm-ai-data").resolve("init_setting");
        Files.createDirectories(initDir);
        String json =
                """
                [
                  {
                    "name": "PM_AI_PORTABLE_BUNDLE_SOURCE_DIR",
                    "value": "\\\\\\\\host\\\\国分工場\\\\国分共有\\\\pm-ai-package-release",
                    "description": ""
                  }
                ]
                """;
        Files.writeString(initDir.resolve(InitSettingPaths.SESSION_DEFAULTS_FILE), json);
        assertEquals(FactorySite.KOKUBU, FactorySite.inferFromPortableBundleInitSetting(installRoot).orElseThrow());
    }
}
