package jp.co.pm.ai.desktop.reconciliation;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class JuchuExpectedPickLabelTest {

    private static JuchuSheetColumnLayout.ExcelHeaderPick pick(
            String letter, int index, String header) {
        return new JuchuSheetColumnLayout.ExcelHeaderPick(letter, index, header);
    }

    @Test
    void expectedPickLabel_restoresBuWhenApHasSameHeaderText() {
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry();
        String path = "C:\\test\\juchu.xlsm";
        var col = JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT;

        registry.setExpectedOverride(path, col, "商品(製品)");
        registry.setExpectedPickLabel(path, col, "BU列: 商品(製品)");

        var picks =
                List.of(
                        pick("AP", 41, "商品(製品)"),
                        pick("BU", 72, "商品(製品)"));

        var mismatch = new JuchuHeaderMismatch(col, "商品(製品)", "タイプ", false);

        assertEquals(
                "BU列: 商品(製品)",
                JuchuSheetHeaderRepairWizard.defaultSelectedPickLabel(
                        mismatch, picks, registry, path));
    }

    @Test
    void suggestPickLabel_findsMatchingColumnAcrossSheet() {
        var col = JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT;
        var picks =
                List.of(
                        pick("AP", 41, "タイプ"),
                        pick("BU", 72, "商品(製品)"));
        var mismatch = new JuchuHeaderMismatch(col, "masterBase商品(製品)", "タイプ", false);

        var best =
                JuchuSheetHeaderRepairWizard.findBestMatchingPick(
                        mismatch, picks, null, "C:\\test\\juchu.xlsm");
        assertEquals("BU", best.columnLetter());
        assertEquals(
                "BU列: 商品(製品)",
                JuchuSheetHeaderRepairWizard.defaultSelectedPickLabel(
                        mismatch, picks, null, "C:\\test\\juchu.xlsm"));
    }

    @Test
    void needsAdoptionPersist_crossColumnPickWithoutSave() {
        var col = JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_PRODUCT;
        var picks =
                List.of(
                        pick("AP", 41, "タイプ"),
                        pick("BU", 72, "商品(製品)"));
        var mismatch = new JuchuHeaderMismatch(col, "masterBase商品(製品)", "タイプ", false);
        var row =
                new JuchuSheetHeaderRepairWizard.KnownRow(
                        mismatch,
                        JuchuSheetHeaderRepairWizard.FixAction.SKIP,
                        "BU列: 商品(製品)",
                        "商品(製品)");

        assertTrue(
                JuchuSheetHeaderRepairWizard.needsAdoptionPersist(
                        row, new JuchuHeaderAliasRegistry(), "C:\\test\\juchu.xlsm", picks));
        assertEquals(
                1,
                JuchuSheetHeaderRepairWizard.promoteRowsNeedingAdoptionPersist(
                        List.of(row),
                        new JuchuHeaderAliasRegistry(),
                        "C:\\test\\juchu.xlsm",
                        picks));
        assertEquals(JuchuSheetHeaderRepairWizard.FixAction.REDEFINE, row.getAction());
    }

    @Test
    void expectedPickLabel_persistsOnDisk(@TempDir Path tempDir) throws Exception {
        Path store = tempDir.resolve("aliases.properties");
        String path = "C:\\test\\juchu.xlsm";
        var col = JuchuSheetColumnLayout.Col.MASTER_BASE_SHOHIN_RAW;

        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry(store);
        registry.setExpectedPickLabel(path, col, "BV列: 商品(原反)");
        registry.saveToDisk();

        JuchuHeaderAliasRegistry reloaded = new JuchuHeaderAliasRegistry(store);
        reloaded.reloadFromDisk();
        assertEquals("BV列: 商品(原反)", reloaded.expectedPickLabelFor(path, col).orElse(""));
    }
}
