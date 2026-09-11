package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class InspectionSheetIndexStoreTest {

    @Test
    void roundTrip_quotesCommaInPath(@TempDir Path tmp) throws Exception {
        Path csv = tmp.resolve("KONAN.csv");
        InspectionSheetIndexStore.Row row =
                new InspectionSheetIndexStore.Row(
                        "C8-9",
                        LocalDate.of(2026, 8, 18),
                        "2026-08",
                        tmp.resolve("dir,with,comma").resolve("a.xlsx").toString(),
                        "a.xlsx",
                        123L,
                        456L,
                        "2026-09-11T00:00:00Z");
        InspectionSheetIndexStore.save(csv, List.of(row));
        List<InspectionSheetIndexStore.Row> loaded = InspectionSheetIndexStore.load(csv);
        assertEquals(1, loaded.size());
        assertEquals(row, loaded.get(0));
    }
}
