package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;

class InspectionSheetDirPickerTest {

    @Test
    void matchesKensa_trimLowercaseOnly() {
        assertTrue(InspectionSheetDirPicker.matchesKensa("kensa"));
        assertTrue(InspectionSheetDirPicker.matchesKensa(" kensa "));
        assertFalse(InspectionSheetDirPicker.matchesKensa("KENSA"));
        assertFalse(InspectionSheetDirPicker.matchesKensa(""));
        assertFalse(InspectionSheetDirPicker.matchesKensa(null));
        assertFalse(InspectionSheetDirPicker.matchesKensa("kens"));
    }

    @Test
    void initialDirectoryCandidates_factoryLeafFirst() {
        var konan = InspectionSheetDirPicker.initialDirectoryCandidates(FactorySite.KONAN);
        assertEquals(Path.of(AppPaths.DEFAULT_PM_AI_INSPECTION_SHEET_DIR_KONAN), konan.get(0));
        var kokubu = InspectionSheetDirPicker.initialDirectoryCandidates(FactorySite.KOKUBU);
        assertTrue(kokubu.get(0).endsWith(Path.of("後加工検査表", "国分工場")));
    }

    @Test
    void looksLikeInspectionSheetFile_acceptsKnownInspectionNames() {
        assertTrue(
                InspectionSheetDirPicker.looksLikeInspectionSheetFile(
                        Path.of("2026_C8-9(SEC済)完了.xlsx")));
        assertTrue(
                InspectionSheetDirPicker.looksLikeInspectionSheetFile(
                        Path.of("C1-4#product(EC.xlsx")));
        assertTrue(
                InspectionSheetDirPicker.looksLikeInspectionSheetFile(
                        Path.of("GB60804.xlsx")));
    }

    @Test
    void looksLikeInspectionSheetFile_rejectsRequestFormAndUnrelatedExcel() {
        assertFalse(
                InspectionSheetDirPicker.looksLikeInspectionSheetFile(
                        Path.of("Y-9月（2026年）加工依頼書（国分Y）.xlsm")));
        assertFalse(
                InspectionSheetDirPicker.looksLikeInspectionSheetFile(Path.of("report.xlsx")));
        assertFalse(InspectionSheetDirPicker.looksLikeInspectionSheetFile(Path.of("notes.csv")));
        assertFalse(InspectionSheetDirPicker.looksLikeInspectionSheetFile(null));
    }

    @Test
    void containsInspectionSheets_trueWhenNestedInspectionExcelExists(@TempDir Path tmp)
            throws Exception {
        Path month = tmp.resolve("2026年").resolve("9月");
        Files.createDirectories(month);
        Files.writeString(month.resolve("2026_C8-9(SEC済)完了.xlsx"), "x");
        assertTrue(InspectionSheetDirPicker.containsInspectionSheets(tmp));
        assertTrue(InspectionSheetDirPicker.validateInspectionSheetDir(tmp).isEmpty());
    }

    @Test
    void containsInspectionSheets_falseForEmptyOrRequestFormFolder(@TempDir Path tmp)
            throws Exception {
        Path empty = tmp.resolve("empty");
        Files.createDirectories(empty);
        assertFalse(InspectionSheetDirPicker.containsInspectionSheets(empty));
        assertTrue(InspectionSheetDirPicker.validateInspectionSheetDir(empty).isPresent());

        Path originals = tmp.resolve("原本");
        Files.createDirectories(originals);
        Files.writeString(originals.resolve("Y-9月（2026年）加工依頼書（国分Y）.xlsm"), "x");
        assertFalse(InspectionSheetDirPicker.containsInspectionSheets(originals));

        assertFalse(InspectionSheetDirPicker.containsInspectionSheets(tmp.resolve("missing")));
        assertFalse(InspectionSheetDirPicker.containsInspectionSheets(null));
    }

    @Test
    void looksLikeInspectionSheetFile_acceptsExcelUnderInspectionTree(@TempDir Path tmp)
            throws Exception {
        Path tree = tmp.resolve("後加工検査表").resolve("国分工場");
        Files.createDirectories(tree);
        Path odd = tree.resolve("backup.xlsx");
        Files.writeString(odd, "x");
        assertTrue(InspectionSheetDirPicker.looksLikeInspectionSheetFile(odd));
        assertTrue(InspectionSheetDirPicker.containsInspectionSheets(tree));
    }
}
