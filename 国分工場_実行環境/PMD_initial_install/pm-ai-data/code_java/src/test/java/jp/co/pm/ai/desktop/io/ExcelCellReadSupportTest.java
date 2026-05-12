package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class ExcelCellReadSupportTest {

    @Test
    void normalizeCommaDigitArtifacts_collapsesPerDigitCommas() {
        assertEquals("3000", ExcelCellReadSupport.normalizeCommaDigitArtifacts("3,0,0,0"));
    }

    @Test
    void normalizeCommaDigitArtifacts_leavesThousandsGrouping() {
        assertEquals("1,234", ExcelCellReadSupport.normalizeCommaDigitArtifacts("1,234"));
    }

    @Test
    void normalizeCommaDigitArtifacts_collapsesIntegerBeforeDecimal() {
        assertEquals("8000.00", ExcelCellReadSupport.normalizeCommaDigitArtifacts("8,0,0,0.00"));
        assertEquals("400.00", ExcelCellReadSupport.normalizeCommaDigitArtifacts("4,0,0.00"));
        assertEquals("-18000.00", ExcelCellReadSupport.normalizeCommaDigitArtifacts("-1,8,0,0,0.00"));
    }
}
