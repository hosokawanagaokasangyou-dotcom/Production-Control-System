package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
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

    @Test
    void readTextForEditor_prettyPrintsValidObject() throws Exception {
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summaryXlsx.toString());
        Path storePath = RequestFormInputSettingsStore.resolveStorePath(ui);
        Files.writeString(storePath, "{\"inputKbn\":[\"通常入力\"]}");

        String text = RequestFormInputSettingsStore.readTextForEditor(ui);
        assertTrue(text.contains("\"inputKbn\""));
        assertTrue(text.contains("通常入力"));
        assertTrue(text.contains("\n"));
    }

    @Test
    void readTextForEditor_returnsRawWhenJsonIsBroken() throws Exception {
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summaryXlsx.toString());
        Path storePath = RequestFormInputSettingsStore.resolveStorePath(ui);
        String broken = "{ \"inputKbn\": [";
        Files.writeString(storePath, broken);

        assertEquals(broken, RequestFormInputSettingsStore.readTextForEditor(ui));
    }

    @Test
    void savePrettyJson_roundTripPreservesChoicesAndPaths() throws Exception {
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summaryXlsx.toString());

        RequestFormInputSettingsStore.Settings saved =
                RequestFormInputSettingsStore.savePrettyJson(
                        ui,
                        """
                        {
                          "inputKbn": ["通常入力", "例外入力"],
                          "fieldDefaults": { "inputKbn": "通常入力" },
                          "targetFolder": "C:\\\\orig",
                          "juchuFilePath": "C:\\\\juchu.xlsm"
                        }
                        """);
        assertEquals(
                List.of("通常入力", "例外入力"),
                saved.comboChoices().optionsFor(RequestFormComboChoices.KEY_INPUT_KBN));
        assertEquals("C:\\orig", saved.paths().targetFolder());
        assertEquals("C:\\juchu.xlsm", saved.paths().juchuFilePath());

        RequestFormInputSettingsStore.Settings loaded =
                RequestFormInputSettingsStore.load(ui).orElseThrow();
        assertEquals("C:\\orig", loaded.paths().targetFolder());
        assertEquals(
                "通常入力",
                loaded.comboChoices().effectiveDefaultFor(RequestFormComboChoices.KEY_INPUT_KBN));
    }

    @Test
    void save_emptyComboDoesNotWipeExistingComboChoices() throws Exception {
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summaryXlsx.toString());

        RequestFormComboChoices combo =
                RequestFormComboChoices.of(
                        Map.of(RequestFormComboChoices.KEY_YOTO, List.of("W（自動車）", "独自用途")));
        RequestFormInputSettingsStore.save(ui, combo, "C:\\orig", "C:\\juchu.xlsm");

        RequestFormInputSettingsStore.save(
                ui, RequestFormComboChoices.empty(), "C:\\orig2", "C:\\juchu2.xlsm");

        RequestFormInputSettingsStore.Settings loaded =
                RequestFormInputSettingsStore.load(ui).orElseThrow();
        assertEquals(
                List.of("W（自動車）", "独自用途"),
                loaded.comboChoices().optionsFor(RequestFormComboChoices.KEY_YOTO));
        assertEquals("C:\\orig2", loaded.paths().targetFolder());
        assertEquals("C:\\juchu2.xlsm", loaded.paths().juchuFilePath());
    }

    @Test
    void save_partialComboDoesNotDropOmittedFeedLoc() throws Exception {
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summaryXlsx.toString());

        RequestFormComboChoices full =
                RequestFormComboChoices.of(
                        Map.of(
                                RequestFormComboChoices.KEY_FEED_LOC,
                                List.of("EC", "SEC", "ｽﾗｲｽ", "ｽﾘｯﾄ", "ｴﾝﾎﾞｽ", "検反", "融着"),
                                RequestFormComboChoices.KEY_USER,
                                List.of("自動転記", "ｵｶﾓﾄ")));
        RequestFormInputSettingsStore.save(ui, full, "", "");

        RequestFormComboChoices partial =
                RequestFormComboChoices.of(
                        Map.of(RequestFormComboChoices.KEY_USER, List.of("自動転記", "ｵｶﾓﾄ", "追加ユーザー")));
        RequestFormInputSettingsStore.save(ui, partial, "", "");

        RequestFormInputSettingsStore.Settings loaded =
                RequestFormInputSettingsStore.load(ui).orElseThrow();
        assertTrue(
                loaded.comboChoices().asMap().containsKey(RequestFormComboChoices.KEY_FEED_LOC),
                "部分保存で投入場所キーが消えてはならない");
        assertEquals(
                List.of("EC", "SEC", "ｽﾗｲｽ", "ｽﾘｯﾄ", "ｴﾝﾎﾞｽ", "検反", "融着"),
                loaded.comboChoices().asMap().get(RequestFormComboChoices.KEY_FEED_LOC));
        assertEquals(
                List.of("自動転記", "ｵｶﾓﾄ", "追加ユーザー"),
                loaded.comboChoices().optionsFor(RequestFormComboChoices.KEY_USER));
        String raw = Files.readString(RequestFormInputSettingsStore.resolveStorePath(ui));
        assertTrue(raw.contains("\"feedLoc\""));
        assertTrue(raw.contains("ｽﾗｲｽ"));
    }

    @Test
    void saveAndLoad_deletedBundledOptionDoesNotReappear() throws Exception {
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summaryXlsx.toString());

        RequestFormComboChoices combo =
                RequestFormComboChoices.of(
                        Map.of(
                                RequestFormComboChoices.KEY_YOTO,
                                List.of("W（自動車）", "B（輸出）", "Y（工材）")));
        RequestFormInputSettingsStore.save(ui, combo, "", "");

        RequestFormComboChoices loaded =
                RequestFormInputSettingsStore.loadComboChoices(ui, GlobalInitSettingTarget.load());
        assertEquals(
                List.of("W（自動車）", "B（輸出）", "Y（工材）"),
                loaded.optionsFor(RequestFormComboChoices.KEY_YOTO));
        assertFalse(loaded.optionsFor(RequestFormComboChoices.KEY_YOTO).contains("小口加工"));
    }

    @Test
    void save_throwsWhenDestinationIsADirectory() throws Exception {
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summaryXlsx.toString());
        Path storePath = RequestFormInputSettingsStore.resolveStorePath(ui);
        Files.createDirectories(storePath);

        assertThrows(
                IOException.class,
                () ->
                        RequestFormInputSettingsStore.save(
                                ui,
                                RequestFormComboChoices.of(
                                        Map.of(
                                                RequestFormComboChoices.KEY_INPUT_KBN,
                                                List.of("通常入力"))),
                                "",
                                ""));
    }

    @Test
    void confirmWriteMessage_distinguishesCreateAndOverwrite(@TempDir Path tmp) throws Exception {
        Path missing = tmp.resolve("request_form_input_settings.json");
        String create = RequestFormInputSettingsStore.confirmWriteMessage(missing);
        assertTrue(create.contains("新規作成"));
        assertTrue(create.contains(AppPaths.REQUEST_FORM_INPUT_SETTINGS_JSON_FILENAME));
        assertTrue(create.contains(missing.toString()));

        Files.writeString(missing, "{}");
        String overwrite = RequestFormInputSettingsStore.confirmWriteMessage(missing);
        assertTrue(overwrite.contains("上書き"));
        assertTrue(overwrite.contains(missing.toString()));
        assertFalse(overwrite.contains("新規作成"));
    }

    @Test
    void savePrettyJson_rejectsInvalidSyntaxAndArrayRoot() throws Exception {
        Path summaryXlsx = tempDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summaryXlsx);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summaryXlsx.toString());

        IOException syntax =
                assertThrows(
                        IOException.class,
                        () -> RequestFormInputSettingsStore.savePrettyJson(ui, "{ not json"));
        assertTrue(syntax.getMessage().contains("構文"));

        IOException arrayRoot =
                assertThrows(
                        IOException.class,
                        () -> RequestFormInputSettingsStore.savePrettyJson(ui, "[1,2]"));
        assertTrue(arrayRoot.getMessage().contains("オブジェクト"));
    }
}
