package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class DispatchAladdinEntrySheetBuilderTest {

    private static final LocalDate TODAY = LocalDate.of(2026, 7, 7);

    private static final List<String> COLUMNS =
            List.of(
                    "依頼NO",
                    ResultDispatchSchema.COL_PROCESS,
                    ResultDispatchSchema.COL_MACHINE,
                    "原反投入日",
                    "回答納期",
                    "換算数量",
                    "実加工数",
                    ResultDispatchSchema.COL_DISPATCH_DATE,
                    ResultDispatchSchema.COL_DISPATCH_QTY);

    @Test
    void dateAxisSpansTodayToLatestDispatchDate() {
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("W1", "巻返し", "M1", "2026-07-08", "100"));
        rows.add(row("W1", "巻返し", "M1", "2026-07-10", "50"));

        List<LocalDate> dates = DispatchAladdinEntrySheetBuilder.dateAxis(rows, TODAY);

        assertEquals(4, dates.size());
        assertEquals(TODAY, dates.getFirst());
        assertEquals(LocalDate.of(2026, 7, 10), dates.getLast());
    }

    @Test
    void dateAxisFallsBackToTodayOnlyWhenAllDatesPast() {
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("W1", "巻返し", "M1", "2026-07-01", "100"));

        List<LocalDate> dates = DispatchAladdinEntrySheetBuilder.dateAxis(rows, TODAY);

        assertEquals(List.of(TODAY), dates);
    }

    @Test
    void buildGroupsByMachineAndTaskAndPivotsSystemQty() {
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("W1", "巻返し", "M1", "2026-07-08", "100"));
        rows.add(row("W1", "巻返し", "M1", "2026-07-09", "200"));
        rows.add(row("W2", "スリット", "M2", "2026-07-08", "80"));

        DispatchAladdinEntrySheetBuilder.EntryWorkbook wb =
                DispatchAladdinEntrySheetBuilder.build(COLUMNS, rows, Map.of(), Map.of(), TODAY);

        assertEquals(2, wb.sheets().size());
        assertEquals("M1", wb.sheets().getFirst().machineName());
        assertEquals("M2", wb.sheets().get(1).machineName());

        DispatchAladdinEntrySheetBuilder.EntryRow w1 = wb.sheets().getFirst().rows().getFirst();
        assertEquals("W1", w1.taskId());
        assertEquals(300d, w1.dispatchTotal(), 1e-9);
        assertEquals(100d, w1.cells().get(LocalDate.of(2026, 7, 8)).systemQty(), 1e-9);
        assertEquals(200d, w1.cells().get(LocalDate.of(2026, 7, 9)).systemQty(), 1e-9);
        assertEquals(LocalDate.of(2026, 7, 8), w1.earliestDispatchDate());
    }

    @Test
    void quantityCheckOkWhenDispatchTotalMatchesConversionQty() {
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> r = row("W1", "巻返し", "M1", "2026-07-08", "300");
        r.put("換算数量", "300");
        rows.add(r);

        DispatchAladdinEntrySheetBuilder.EntryRow out = buildSingleRow(rows);

        assertTrue(out.quantityOk());
        assertEquals("OK", out.quantityCheckText());
    }

    @Test
    void quantityCheckOkWhenDispatchTotalPlusCompletedMatchesConversionQty() {
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> r = row("W7-3", "EC", "M1", "2026-07-08", "1500");
        r.put("換算数量", "6000");
        r.put("実加工数", "4500");
        rows.add(r);

        DispatchAladdinEntrySheetBuilder.EntryRow out = buildSingleRow(rows);

        assertTrue(out.quantityOk());
        assertEquals("OK", out.quantityCheckText());
    }

    @Test
    void quantityCheckNgShowsSignedDifference() {
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> r = row("W1", "巻返し", "M1", "2026-07-08", "100");
        r.put("換算数量", "300");
        rows.add(r);

        DispatchAladdinEntrySheetBuilder.EntryRow out = buildSingleRow(rows);

        assertFalse(out.quantityOk());
        assertEquals("NG (差 -200)", out.quantityCheckText());
    }

    @Test
    void aladdinQtyIsLookedUpPerMachineTaskDateProcess() {
        List<String> shapedHeaders =
                List.of("機械名", "依頼NO", "工程名", "2026/07/08", "2026/07/09");
        List<List<String>> shapedRows =
                List.of(List.of("M1", "W1", "巻返し", "120", "0"));
        Map<String, Map<String, Map<String, Map<String, Double>>>> lookup =
                AladdinShapedPlanQtyLookup.buildLookup(shapedHeaders, shapedRows);

        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("W1", "巻返し", "M1", "2026-07-08", "100"));

        DispatchAladdinEntrySheetBuilder.EntryWorkbook wb =
                DispatchAladdinEntrySheetBuilder.build(COLUMNS, rows, lookup, Map.of(), TODAY);
        DispatchAladdinEntrySheetBuilder.EntryCell cell =
                wb.sheets().getFirst().rows().getFirst().cells().get(LocalDate.of(2026, 7, 8));

        assertNotNull(cell);
        assertEquals(120d, cell.aladdinQty(), 1e-9);
        assertEquals(100d, cell.systemQty(), 1e-9);
        assertTrue(cell.mismatch());
        assertEquals("（現アラ計）120\n（シス計）100", cell.cellText());
    }

    @Test
    void cellWithoutQuantitiesIsEmpty() {
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("W1", "巻返し", "M1", "2026-07-08", "100"));

        DispatchAladdinEntrySheetBuilder.EntryCell cell =
                buildSingleRow(rows).cells().get(TODAY);

        assertNotNull(cell);
        assertTrue(cell.isEmpty());
        assertEquals("", cell.cellText());
        assertFalse(cell.mismatch());
    }

    @Test
    void indexInfoOverridesKaitoNokiAndSuppliesContractNo() {
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> r = row("W1", "巻返し", "M1", "2026-07-08", "100");
        r.put("回答納期", "2026-07-30");
        rows.add(r);
        Map<String, DispatchAladdinEntrySheetBuilder.IndexInfo> index =
                Map.of(
                        AladdinShapedPlanQtyLookup.normalizeTaskIdKey("W1"),
                        new DispatchAladdinEntrySheetBuilder.IndexInfo("2026/07/25", "K-123"));

        DispatchAladdinEntrySheetBuilder.EntryWorkbook wb =
                DispatchAladdinEntrySheetBuilder.build(COLUMNS, rows, Map.of(), index, TODAY);
        DispatchAladdinEntrySheetBuilder.EntryRow out = wb.sheets().getFirst().rows().getFirst();

        assertEquals("2026/07/25", out.kaitoNoki());
        assertEquals("K-123", out.contractNo());
    }

    @Test
    void kaitoNokiFallsBackToDispatchRowWhenIndexMissing() {
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> r = row("W1", "巻返し", "M1", "2026-07-08", "100");
        r.put("回答納期", "2026-07-30");
        rows.add(r);

        DispatchAladdinEntrySheetBuilder.EntryRow out = buildSingleRow(rows);

        assertEquals("2026-07-30", out.kaitoNoki());
        assertEquals("", out.contractNo());
    }

    @Test
    void rowsSortByEarliestDispatchDateThenTaskId() {
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("W9", "巻返し", "M1", "2026-07-10", "10"));
        rows.add(row("W1", "巻返し", "M1", "2026-07-08", "10"));

        DispatchAladdinEntrySheetBuilder.EntryWorkbook wb =
                DispatchAladdinEntrySheetBuilder.build(COLUMNS, rows, Map.of(), Map.of(), TODAY);
        List<DispatchAladdinEntrySheetBuilder.EntryRow> out = wb.sheets().getFirst().rows();

        assertEquals("W1", out.getFirst().taskId());
        assertEquals("W9", out.get(1).taskId());
    }

    private static DispatchAladdinEntrySheetBuilder.EntryRow buildSingleRow(
            List<Map<String, String>> rows) {
        DispatchAladdinEntrySheetBuilder.EntryWorkbook wb =
                DispatchAladdinEntrySheetBuilder.build(COLUMNS, rows, Map.of(), Map.of(), TODAY);
        return wb.sheets().getFirst().rows().getFirst();
    }

    private static Map<String, String> row(
            String tid, String proc, String machine, String dispatchDate, String qty) {
        Map<String, String> m = new LinkedHashMap<>();
        m.put("依頼NO", tid);
        m.put(ResultDispatchSchema.COL_PROCESS, proc);
        m.put(ResultDispatchSchema.COL_MACHINE, machine);
        m.put("原反投入日", "2026-07-05");
        m.put("回答納期", "");
        m.put("換算数量", "300");
        m.put("実加工数", "0");
        m.put(ResultDispatchSchema.COL_DISPATCH_DATE, dispatchDate);
        m.put(ResultDispatchSchema.COL_DISPATCH_QTY, qty);
        return m;
    }
}
