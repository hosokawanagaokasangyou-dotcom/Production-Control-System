package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;

import org.junit.jupiter.api.Test;

class AladdinEntryIdentityCheckDiffTableTest {

    @Test
    void toCsv_usesCommaAndNewlineWithHeader() {
        AladdinEntryDispatchPlanIdentityCheck.Diff d =
                new AladdinEntryDispatchPlanIdentityCheck.Diff(
                        "EC機 湖南", "W8-9", "EC", LocalDate.of(2026, 8, 18), 1500, 0);

        String csv = AladdinEntryIdentityCheckDiffTable.toCsv(List.of(d));

        assertEquals(
                "機械,依頼NO,工程,日付,シス計,加工計画\nEC機 湖南,W8-9,EC,2026-08-18,1500,0",
                csv);
    }

    @Test
    void toCsv_quotesCellContainingComma() {
        AladdinEntryDispatchPlanIdentityCheck.Diff d =
                new AladdinEntryDispatchPlanIdentityCheck.Diff(
                        "A,B", "T001", "工程A", LocalDate.of(2026, 7, 7), 10, 99);

        String csv = AladdinEntryIdentityCheckDiffTable.toCsv(List.of(d));

        assertTrue(csv.contains("\"A,B\""), csv);
    }

    @Test
    void toHtmlTable_containsHeaderAndCells() {
        AladdinEntryDispatchPlanIdentityCheck.Diff d =
                new AladdinEntryDispatchPlanIdentityCheck.Diff(
                        "M1", "T001", "工程A", LocalDate.of(2026, 7, 7), 10, 99);

        String html = AladdinEntryIdentityCheckDiffTable.toHtmlTable(List.of(d));

        assertTrue(html.contains("<table"), html);
        assertTrue(html.contains("機械"), html);
        assertTrue(html.contains("T001"), html);
        assertTrue(html.contains("99"), html);
    }

    @Test
    void toCsv_emptyWhenNoDiffs() {
        assertEquals("", AladdinEntryIdentityCheckDiffTable.toCsv(List.of()));
        assertEquals("", AladdinEntryIdentityCheckDiffTable.toCsv(null));
    }
}
