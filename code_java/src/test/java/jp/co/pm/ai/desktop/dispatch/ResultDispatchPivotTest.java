package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Comparator;
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
    void matchesTaskProfileExceptTrialOrder_ignoresTrialColumn() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        Map<String, String> profile = new LinkedHashMap<>();
        Map<String, String> row = new LinkedHashMap<>();
        for (String c : cols) {
            profile.put(c, "");
            row.put(c, "");
        }
        profile.put(ResultDispatchSchema.COL_PROCESS, "P");
        profile.put(ResultDispatchSchema.COL_MACHINE, "M");
        profile.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, "2");
        row.put(ResultDispatchSchema.COL_PROCESS, "P");
        row.put(ResultDispatchSchema.COL_MACHINE, "M");
        row.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, "1");
        assertTrue(ResultDispatchPivot.matchesTaskProfileExceptTrialOrder(cols, profile, row));
        assertFalse(ResultDispatchPivot.matchesTaskProfile(cols, profile, row));
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
        base.put("依頼NO", "X1");
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

    @Test
    void parseLoadOutcome_surfacesPythonJsonError() {
        String json =
                "{\"error\": \"No module named 'pandas'\", \"blocks\": {}}";
        MachineCalendarBlockIndex.LoadOutcome lo = MachineCalendarBlockIndex.parseLoadOutcome(json);
        assertTrue(lo.index().isEmpty());
        assertEquals("No module named 'pandas'", lo.pythonJsonError());
        assertEquals(null, lo.pythonDiagnosticsJson());
    }

    @Test
    void parseLoadOutcome_errorOnly_noBlocksKey() {
        String json = "{\"error\": \"file_not_found\", \"path\": \"/x\"}";
        MachineCalendarBlockIndex.LoadOutcome lo = MachineCalendarBlockIndex.parseLoadOutcome(json);
        assertTrue(lo.index().isEmpty());
        assertEquals("file_not_found", lo.pythonJsonError());
        assertEquals(null, lo.pythonDiagnosticsJson());
    }

    @Test
    void parseLoadOutcome_includesDiagnosticsJson() {
        String json =
                "{\"blocks\":{},\"diagnostics\":{\"has_machine_calendar_sheet\":true,\"equipment_list_len\":3}}";
        MachineCalendarBlockIndex.LoadOutcome lo = MachineCalendarBlockIndex.parseLoadOutcome(json);
        assertTrue(lo.index().isEmpty());
        assertEquals(null, lo.pythonJsonError());
        assertTrue(lo.pythonDiagnosticsJson().contains("has_machine_calendar_sheet"));
        assertTrue(lo.pythonDiagnosticsJson().contains("equipment_list_len"));
    }

    @Test
    void distinctProfiles_sortByStaticGroupKey_isStableWhenRowOrderDiffers() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        List<Map<String, String>> rowsA = new ArrayList<>();
        List<Map<String, String>> rowsB = new ArrayList<>();
        LinkedHashMap<String, String> rowP = new LinkedHashMap<>();
        for (String c : cols) {
            rowP.put(c, "");
        }
        rowP.put(ResultDispatchSchema.COL_PROCESS, "BProc");
        rowP.put(ResultDispatchSchema.COL_MACHINE, "M");
        rowP.put(ResultDispatchSchema.COL_DISPATCH_DATE, "2026-05-10");
        rowP.put(ResultDispatchSchema.COL_DISPATCH_QTY, "1");
        Map<String, String> rowQ = new LinkedHashMap<>(rowP);
        rowQ.put(ResultDispatchSchema.COL_PROCESS, "AProc");
        rowsA.add(rowP);
        rowsA.add(rowQ);
        rowsB.add(rowQ);
        rowsB.add(rowP);
        List<Map<String, String>> pa = ResultDispatchPivot.distinctTaskProfiles(cols, rowsA);
        List<Map<String, String>> pb = ResultDispatchPivot.distinctTaskProfiles(cols, rowsB);
        assertEquals("BProc", pa.getFirst().get(ResultDispatchSchema.COL_PROCESS));
        assertEquals("AProc", pb.getFirst().get(ResultDispatchSchema.COL_PROCESS));
        Comparator<Map<String, String>> cmp =
                Comparator.comparing(p -> ResultDispatchNormalizer.staticGroupKey(cols, p));
        pa.sort(cmp);
        pb.sort(cmp);
        assertEquals(pa.getFirst().get(ResultDispatchSchema.COL_PROCESS), pb.getFirst().get(ResultDispatchSchema.COL_PROCESS));
    }

    @Test
    void pickJsonPayload_skipsLeadingLogLines() {
        String stdout =
                "[planning_core] warning line\n"
                        + "{\"blocks\":{\"M1\":[\"2026-05-09\"]}}\n";
        String payload = MachineCalendarBlockIndex.pickJsonPayload(stdout);
        assertTrue(payload.startsWith("{"));
        MachineCalendarBlockIndex idx = MachineCalendarBlockIndex.parseStdoutJson(payload);
        assertEquals(true, idx.isBlockedDay("P", "M1", LocalDate.of(2026, 5, 9)));
    }

    @Test
    void parseIsoDate_acceptsSlashAndDotYmd() {
        LocalDate d = LocalDate.of(2026, 5, 4);
        assertEquals(d, ResultDispatchPivot.parseIsoDate("2026-05-04"));
        assertEquals(d, ResultDispatchPivot.parseIsoDate("2026/05/04"));
        assertEquals(d, ResultDispatchPivot.parseIsoDate("2026.05.04"));
    }
}
