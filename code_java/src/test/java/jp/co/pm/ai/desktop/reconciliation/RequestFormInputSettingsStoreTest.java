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
}
