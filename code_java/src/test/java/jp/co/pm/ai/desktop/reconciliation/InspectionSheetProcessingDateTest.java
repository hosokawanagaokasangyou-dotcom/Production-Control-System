package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.Optional;

import org.junit.jupiter.api.Test;

class InspectionSheetProcessingDateTest {

    @Test
    void parse_singleDateInLabel() {
        assertEquals(
                LocalDate.of(2026, 1, 8),
                InspectionSheetProcessingDate.parse("加工日： 2026/1/8").orElseThrow());
    }

    @Test
    void parse_rangeUsesStartDate() {
        assertEquals(
                LocalDate.of(2026, 1, 9),
                InspectionSheetProcessingDate.parse("加工日： 2026/1/9 ～ 1/13").orElseThrow());
    }

    @Test
    void parse_excelSerial() {
        assertEquals(
                LocalDate.of(2026, 8, 27),
                InspectionSheetProcessingDate.fromExcelSerial(46261).orElseThrow());
        assertEquals(
                LocalDate.of(2026, 8, 18),
                InspectionSheetProcessingDate.fromExcelSerial(46252).orElseThrow());
    }

    @Test
    void yearMonth_fromDate() {
        assertEquals("2026-08", InspectionSheetProcessingDate.yearMonth(LocalDate.of(2026, 8, 18)));
        assertEquals("", InspectionSheetProcessingDate.yearMonth(null));
    }

    @Test
    void parse_empty() {
        assertEquals(Optional.empty(), InspectionSheetProcessingDate.parse(""));
        assertEquals(Optional.empty(), InspectionSheetProcessingDate.parse("加工日カコウビ"));
        assertFalse(InspectionSheetProcessingDate.fromExcelSerial(0).isPresent());
    }

    @Test
    void isProcessingDateLabel_detectsFuriganaConcat() {
        assertTrue(InspectionSheetProcessingDate.isProcessingDateLabel("加工日カコウビ"));
        assertTrue(InspectionSheetProcessingDate.isProcessingDateLabel("加工日： 2026/1/8"));
        assertFalse(InspectionSheetProcessingDate.isProcessingDateLabel("投入日"));
    }
}
