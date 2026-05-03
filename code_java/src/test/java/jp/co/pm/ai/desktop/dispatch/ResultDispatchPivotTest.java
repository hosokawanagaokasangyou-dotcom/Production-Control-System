package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class ResultDispatchPivotTest {

    @Test
    void normalizePreservesTotalQty() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        List<Map<String, String>> rows = new ArrayList<>();
        LinkedHashMap<String, String> a = new LinkedHashMap<>();
        for (String c : cols) {
            a.put(c, "");
        }
        a.put(ResultDispatchSchema.COL_PROCESS, "P");
        a.put(ResultDispatchSchema.COL_MACHINE, "M");
        a.put(ResultDispatchSchema.COL_DISPATCH_DATE, "2026-05-10");
        a.put(ResultDispatchSchema.COL_DISPATCH_QTY, "50");
        rows.add(a);
        LinkedHashMap<String, String> b = new LinkedHashMap<>(a);
        b.put(ResultDispatchSchema.COL_DISPATCH_QTY, "50");
        rows.add(b);

        double sumBefore =
                rows.stream()
                        .mapToDouble(
                                r ->
                                        ResultDispatchNormalizer.parseDouble(
                                                r.get(ResultDispatchSchema.COL_DISPATCH_QTY)))
                        .sum();
        ResultDispatchNormalizer.normalizeInPlace(cols, rows);
        double sumAfter =
                rows.stream()
                        .mapToDouble(
                                r ->
                                        ResultDispatchNormalizer.parseDouble(
                                                r.get(ResultDispatchSchema.COL_DISPATCH_QTY)))
                        .sum();
        assertEquals(sumBefore, sumAfter, 1e-6);
        assertEquals(1, rows.size());
        assertEquals(100.0, ResultDispatchNormalizer.parseDouble(rows.getFirst().get(ResultDispatchSchema.COL_DISPATCH_QTY)), 1e-9);
    }

    @Test
    void upsertAllocationChangesQuantity() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        List<Map<String, String>> rows = new ArrayList<>();
        LinkedHashMap<String, String> base = new LinkedHashMap<>();
        for (String c : cols) {
            base.put(c, "");
        }
        base.put(ResultDispatchSchema.COL_PROCESS, "P");
        base.put(ResultDispatchSchema.COL_MACHINE, "M");
        base.put("\u4f9d\u983cNO", "X1");
        base.put(ResultDispatchSchema.COL_DISPATCH_DATE, "2026-05-10");
        base.put(ResultDispatchSchema.COL_DISPATCH_QTY, "100");
        rows.add(base);
        ResultDispatchNormalizer.normalizeInPlace(cols, rows);

        Map<String, String> profile =
                ResultDispatchPivot.distinctTaskProfiles(cols, rows).getFirst();
        ResultDispatchPivot.upsertAllocation(
                cols, rows, profile, LocalDate.of(2026, 5, 10), 40);
        assertEquals(1, rows.size());
        assertEquals(
                40.0,
                ResultDispatchNormalizer.parseDouble(
                        rows.getFirst().get(ResultDispatchSchema.COL_DISPATCH_QTY)),
                1e-9);
    }

    @Test
    void machineCalendarJsonParses() {
        String json = "{\"blocks\":{\"M1\":[\"2026-05-09\"]}}";
        MachineCalendarBlockIndex idx = MachineCalendarBlockIndex.parseStdoutJson(json);
        assertEquals(true, idx.isBlockedDay("P", "M1", LocalDate.of(2026, 5, 9)));
    }
}
