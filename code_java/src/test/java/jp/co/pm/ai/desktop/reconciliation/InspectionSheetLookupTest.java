package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;

import org.junit.jupiter.api.Test;

class InspectionSheetLookupTest {

    @Test
    void findsNormalizedIrai() {
        InspectionSheetIndexStore.Row row =
                new InspectionSheetIndexStore.Row(
                        "C8-9", LocalDate.of(2026, 8, 18), "2026-08", "p", "f.xlsx", 1, 1, "t");
        List<InspectionSheetIndexStore.Row> hits =
                InspectionSheetLookup.find(List.of(row), "Ｃ８－９");
        assertEquals(1, hits.size());
        assertEquals("C8-9", hits.get(0).iraiNo());
        assertTrue(InspectionSheetLookup.find(List.of(row), "C9-1").isEmpty());
    }
}
