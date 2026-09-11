package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class InspectionSheetIraiNoTest {

    @Test
    void extractFromFileName_stripsYearPrefix() {
        assertEquals(
                "C8-9",
                InspectionSheetIraiNo.extractFromFileName("2026_C8-9(SEC済)完了.xlsx").orElseThrow());
    }

    @Test
    void extractFromFileName_oldKonanName() {
        assertEquals(
                "C10-10",
                InspectionSheetIraiNo.extractFromFileName("C10-10（SEC済-ｽﾘｯﾄ済）完了.xlsx")
                        .orElseThrow());
    }

    @Test
    void extractFromFileName_stripsNumericDatePrefixBeforeHash() {
        assertEquals(
                "C4-46",
                InspectionSheetIraiNo.extractFromFileName(
                                "510C4-46#5Z002-NA00-1060x500QABK0A(LAC.xlsx")
                        .orElseThrow());
    }

    @Test
    void extractFromFileName_kokubuHashName() {
        assertEquals(
                "C1-4",
                InspectionSheetIraiNo.extractFromFileName(
                                "C1-4#15011-NY00-1240x300FAWH1V(EC.xlsx")
                        .orElseThrow());
    }

    @Test
    void extractFromFileName_tpiKeepsDigits() {
        assertEquals(
                "TPI 1-1",
                InspectionSheetIraiNo.extractFromFileName(
                                "TPI 1-1#30500-DG20- 960X100Z-AGA0U(融着.xlsx")
                        .orElseThrow());
    }

    @Test
    void normalize_collapsesTpiSpaceAndFullwidth() {
        assertEquals("TPI1-1", InspectionSheetIraiNo.normalize("TPI 1-1"));
        assertEquals("C8-9", InspectionSheetIraiNo.normalize("Ｃ８－９"));
        assertEquals("C8-9", InspectionSheetIraiNo.normalize(" c8-9 "));
    }

    @Test
    void matches_ignoresCaseAndWidth() {
        assertTrue(InspectionSheetIraiNo.matches("C8-9", "ｃ８－９"));
        assertTrue(InspectionSheetIraiNo.matches("TPI 1-1", "TPI1-1"));
    }

    @Test
    void extractFromCellText_usesFirstIraiToken() {
        assertEquals("C1-4", InspectionSheetIraiNo.extractFromCellText("C1-4").orElseThrow());
        assertEquals(
                "TPI 1-1", InspectionSheetIraiNo.extractFromCellText("TPI 1-1").orElseThrow());
    }
}
