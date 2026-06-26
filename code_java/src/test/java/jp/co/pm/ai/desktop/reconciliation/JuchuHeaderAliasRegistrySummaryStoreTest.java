package jp.co.pm.ai.desktop.reconciliation;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class JuchuHeaderAliasRegistrySummaryStoreTest {

    @TempDir Path tempDir;

    @Test
    void saveToDisk_writesJsonWhenStorePathEndsWithJson() throws Exception {
        Path store = tempDir.resolve("juchu_header_aliases.json");
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry(store);
        var col = JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT;
        String juchuPath = "C:\\test\\juchu.xlsm";

        registry.setExpectedOverride(juchuPath, col, "商品(製品)");
        registry.setExpectedPickLabel(juchuPath, col, "BU列: 商品(製品)");
        registry.saveToDisk();

        assertTrue(Files.isRegularFile(store));
        String text = Files.readString(store);
        assertTrue(text.contains("expectedPick"));
        assertTrue(text.contains("BU列: 商品(製品)"));

        JuchuHeaderAliasRegistry reloaded = new JuchuHeaderAliasRegistry(store);
        reloaded.reloadFromDisk();
        assertEquals(
                "BU列: 商品(製品)",
                reloaded.expectedPickLabelFor(juchuPath, col).orElse(""));
    }

    @Test
    void resolveStorePath_usesSummaryJsonWhenUiProvided(@TempDir Path tmp) throws Exception {
        Path summaryWorkbook = tmp.resolve("shared").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createDirectories(summaryWorkbook.getParent());
        Files.createFile(summaryWorkbook);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summaryWorkbook.toString());
        assertEquals(
                AppPaths.juchuHeaderAliasesJsonPath(ui),
                JuchuHeaderAliasRegistry.resolveStorePath(FactorySite.KONAN, ui));
    }

    @Test
    void resolveStorePath_fallsBackToLegacyHomeWhenUiEmpty() {
        assertEquals(
                AppPaths.juchuHeaderAliasesLegacyHomePath(FactorySite.KONAN),
                JuchuHeaderAliasRegistry.resolveStorePath(FactorySite.KONAN, Map.of()));
    }
}
