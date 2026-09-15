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

    @Test
    void loadMerged_overlaysPartialByPath_keepsCompleteUntilCleared(@TempDir Path tmp)
            throws Exception {
        Path csv = tmp.resolve("KONAN.csv");
        InspectionSheetIndexStore.Row complete =
                row("OLD", tmp.resolve("a.xlsx").toString(), 10L, 100L);
        InspectionSheetIndexStore.Row other =
                row("KEEP", tmp.resolve("b.xlsx").toString(), 20L, 200L);
        InspectionSheetIndexStore.Row partial =
                row("NEW", tmp.resolve("a.xlsx").toString(), 11L, 101L);
        InspectionSheetIndexStore.save(csv, List.of(complete, other));
        InspectionSheetIndexStore.savePartial(csv, List.of(partial));

        assertEquals(List.of(complete, other), InspectionSheetIndexStore.load(csv));
        List<InspectionSheetIndexStore.Row> merged = InspectionSheetIndexStore.loadMerged(csv);
        assertEquals(2, merged.size());
        assertEquals("NEW", merged.get(0).iraiNo());
        assertEquals("KEEP", merged.get(1).iraiNo());

        InspectionSheetIndexStore.clearPartial(csv);
        assertEquals(List.of(complete, other), InspectionSheetIndexStore.loadMerged(csv));
    }

    private static InspectionSheetIndexStore.Row row(
            String irai, String path, long mtime, long size) {
        return new InspectionSheetIndexStore.Row(
                irai, LocalDate.of(2026, 8, 18), "2026-08", path, "a.xlsx", mtime, size, "t");
    }
}
