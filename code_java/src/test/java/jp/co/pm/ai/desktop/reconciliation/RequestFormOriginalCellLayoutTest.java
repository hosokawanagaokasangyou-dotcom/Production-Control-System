package jp.co.pm.ai.desktop.reconciliation;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RequestFormOriginalCellLayoutTest {

    @Test
    void columnLetterToIndex_matchesExcelRefs() {
        assertEquals(17, RequestFormOriginalCellLayout.columnLetterToIndex("R"));
        assertEquals(30, RequestFormOriginalCellLayout.columnLetterToIndex("AE"));
        assertEquals(23, RequestFormOriginalCellLayout.columnLetterToIndex("X"));
        assertEquals(31, RequestFormOriginalCellLayout.columnLetterToIndex("AF"));
    }

    @Test
    void basicField_cells_matchSpec() {
        RequestFormOriginalCellLayout.CellAddress irai =
                RequestFormOriginalCellLayout.BasicField.IRAI_NO.cell();
        assertEquals("R5", irai.excelRef());
        assertEquals(4, irai.rowIndex());
        assertEquals(17, irai.columnIndex());

        RequestFormOriginalCellLayout.CellAddress charge =
                RequestFormOriginalCellLayout.BasicField.KAKOCHIN.cell();
        assertEquals("AE20", charge.excelRef());
    }

    @Test
    void excelRef_roundTrip() {
        assertEquals("X14", RequestFormOriginalCellLayout.excelRef(13, 23));
        assertEquals("H23", RequestFormOriginalCellLayout.excelRef(22, 7));
    }

    @Test
    void joinNonBlankParts_skipsEmpty() {
        assertEquals(
                "line1 line2",
                RequestFormOriginalCellLayout.joinNonBlankParts(java.util.List.of("line1", "", " line2 ")));
    }

    @Test
    void formExtractKeys_includeTokkiSeparately() {
        assertTrue(RequestFormOriginalCellLayout.FORM_EXTRACT_RAW_KEYS.contains("特記事項1"));
        assertTrue(RequestFormOriginalCellLayout.FORM_EXTRACT_RAW_KEYS.contains("特記事項2"));
    }
}
