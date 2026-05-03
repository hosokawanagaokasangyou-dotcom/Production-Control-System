package jp.co.pm.ai.desktop.print;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.time.LocalDate;
import java.util.List;

import org.junit.jupiter.api.Test;

class OperatorCardDocumentBuilderTest {

    @Test
    void mergeTimeRange_joins_first_start_and_last_end() {
        assertEquals(
                "08:00-09:30",
                OperatorCardDocumentBuilder.mergeTimeRange(
                        "08:00-08:10", "09:20-09:30"));
    }

    @Test
    void resolveThreeDayColumns_finds_mm_dd_headers() {
        List<String> cols =
                List.of(
                        "\u6642\u9593\u5e2f",
                        "05/07 (Thu)",
                        "05/08 (Fri)",
                        "05/09 (Sat)");
        LocalDate start = LocalDate.of(2026, 5, 7);
        List<String> three = OperatorCardDocumentBuilder.resolveThreeDayColumns(cols, start);
        assertEquals(3, three.size());
        assertEquals("05/07 (Thu)", three.get(0));
        assertEquals("05/08 (Fri)", three.get(1));
        assertEquals("05/09 (Sat)", three.get(2));
    }

    @Test
    void parseColumnDate_reads_month_day() {
        LocalDate d =
                OperatorCardDocumentBuilder.parseColumnDate("05/08 (Fri)", 2026);
        assertNotNull(d);
        assertEquals(LocalDate.of(2026, 5, 8), d);
    }
}
