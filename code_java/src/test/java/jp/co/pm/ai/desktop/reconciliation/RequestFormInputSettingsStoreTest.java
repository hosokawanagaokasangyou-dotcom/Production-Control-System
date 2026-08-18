package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.DesktopSessionStateStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;

class RequestFormInputSettingsStoreTest {

    @TempDir
    Path tempDir;

    @Test
    void saveAndLoad_besideSummaryWorkbook() throws Exception {
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summaryXlsx.toString());

        RequestFormComboChoices combo =
                RequestFormComboChoices.of(
                        Map.of(RequestFormComboChoices.KEY_INPUT_KBN, List.of("通常入力", "例外入力")),
                        Map.of(RequestFormComboChoices.KEY_INPUT_KBN, "通常入力"));
        RequestFormInputSettingsStore.save(
                ui, combo, "C:\\work\\originals", "C:\\work\\juchu.xlsm");

        Path storePath = RequestFormInputSettingsStore.resolveStorePath(ui);
        assertEquals(summaryXlsx.getParent(), storePath.getParent());
        assertTrue(Files.isRegularFile(storePath));

        RequestFormInputSettingsStore.Settings loaded =
                RequestFormInputSettingsStore.load(ui).orElseThrow();
        assertEquals("通常入力", loaded.comboChoices().effectiveDefaultFor(RequestFormComboChoices.KEY_INPUT_KBN));
        assertEquals(2, loaded.comboChoices().optionsFor(RequestFormComboChoices.KEY_INPUT_KBN).size());
        assertEquals("C:\\work\\originals", loaded.paths().targetFolder());
        assertEquals("C:\\work\\juchu.xlsm", loaded.paths().juchuFilePath());
    }

    @Test
    void factoryShipmentComboChoices_fallBackToBundledWhenNoInitSetting() {
        RequestFormComboChoices factory =
                DesktopSessionStateStore.factoryShipmentRequestFormComboChoices(
                        Map.of(), FactorySite.KONAN);
        assertEquals(
                "通常入力",
                factory.effectiveDefaultFor(RequestFormComboChoices.KEY_INPUT_KBN));
        assertEquals(
                "後加工",
                factory.effectiveDefaultFor(RequestFormComboChoices.KEY_KAKO_KBN));
        assertTrue(
                factory.optionsFor(RequestFormComboChoices.KEY_INPUT_KBN)
                        .contains("例外入力"));
    }

    @Test
    void saveAndLoad_masterCandidatePrefixFilters() throws Exception {
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summaryXlsx.toString());

        RequestFormComboChoices combo =
                RequestFormComboChoices.of(
                        Map.of(
                                RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_PRODUCT,
                                List.of("A2", "B1"),
                                RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_RAW,
                                List.of("G1")),
                        Map.of());
        RequestFormInputSettingsStore.save(ui, combo, "", "");

        Path storePath = RequestFormInputSettingsStore.resolveStorePath(ui);
        assertEquals(summaryXlsx.getParent(), storePath.getParent());
        assertTrue(Files.readString(storePath).contains("masterCandidatePrefixProduct"));
        assertTrue(Files.readString(storePath).contains("masterCandidatePrefixRaw"));

        RequestFormComboChoices loaded =
                RequestFormInputSettingsStore.loadComboChoices(
                        ui, jp.co.pm.ai.desktop.config.GlobalInitSettingTarget.load());
        assertEquals(
                List.of("A2", "B1"),
                loaded.optionsFor(RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_PRODUCT));
        assertEquals(
                List.of("G1"),
                loaded.optionsFor(RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_RAW));
    }

    @Test
    void load_readsFlatPrefixKeysAtSettingsRoot() throws Exception {
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summaryXlsx.toString());

        Path storePath = RequestFormInputSettingsStore.resolveStorePath(ui);
        Files.writeString(
                storePath,
                """
                {
                  "masterCandidatePrefixProduct": ["LEG"],
                  "masterCandidatePrefixRaw": ["RAW"]
                }
                """);

        RequestFormComboChoices loaded =
                RequestFormInputSettingsStore.loadComboChoices(
                        ui, GlobalInitSettingTarget.load());
        assertEquals(
                List.of("LEG"),
                loaded.optionsFor(RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_PRODUCT));
        assertEquals(
                List.of("RAW"),
                loaded.optionsFor(RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_RAW));
    }

    @Test
    void resolveEffectiveJuchuFilePath_envWinsOverSettingsJsonOtherFactory() throws Exception {
        FactorySite prior = GlobalInitSettingTarget.load();
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        summaryXlsx.toString(),
                        AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE,
                        AppPaths.DEFAULT_PM_AI_REQUEST_FORM_JUCHU_FILE_KOKUBU);
        try {
            RequestFormInputSettingsStore.save(
                    ui,
                    RequestFormComboChoices.empty(),
                    "",
                    AppPaths.DEFAULT_PM_AI_REQUEST_FORM_JUCHU_FILE_KONAN);

            String resolved = RequestFormInputSettingsStore.resolveEffectiveJuchuFilePath(ui);
            assertEquals(
                    Path.of(AppPaths.DEFAULT_PM_AI_REQUEST_FORM_JUCHU_FILE_KOKUBU)
                            .toAbsolutePath()
                            .normalize()
                            .toString(),
                    resolved);
        } finally {
            GlobalInitSettingTarget.save(prior);
        }
    }

    @Test
    void resolveEffectiveJuchuFilePath_skipsSettingsJsonWhenItConflictsWithFactory()
            throws Exception {
        FactorySite prior = GlobalInitSettingTarget.load();
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        summaryXlsx.toString(),
                        AppPaths.KEY_PM_AI_FACTORY_SITE,
                        FactorySite.KOKUBU.name());
        try {
            RequestFormInputSettingsStore.save(
                    ui,
                    RequestFormComboChoices.empty(),
                    "",
                    AppPaths.DEFAULT_PM_AI_REQUEST_FORM_JUCHU_FILE_KONAN);

            String resolved = RequestFormInputSettingsStore.resolveEffectiveJuchuFilePath(ui);
            assertEquals(
                    Path.of(AppPaths.DEFAULT_PM_AI_REQUEST_FORM_JUCHU_FILE_KOKUBU)
                            .toAbsolutePath()
                            .normalize()
                            .toString(),
                    resolved);
        } finally {
            GlobalInitSettingTarget.save(prior);
        }
    }
}
