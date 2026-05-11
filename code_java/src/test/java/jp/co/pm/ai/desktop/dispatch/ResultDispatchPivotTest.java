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
    void parseIsoDate_acceptsSlashAndDotYmd() {
        LocalDate d = LocalDate.of(2026, 5, 4);
        assertEquals(d, ResultDispatchPivot.parseIsoDate("2026-05-04"));
        assertEquals(d, ResultDispatchPivot.parseIsoDate("2026/05/04"));
        assertEquals(d, ResultDispatchPivot.parseIsoDate("2026.05.04"));
    }

    @Test
    void distinctWideTaskProfiles_unifiesSameTaskSplitByTrialOrMeta() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        List<Map<String, String>> rows = new ArrayList<>();
        LinkedHashMap<String, String> r1 = new LinkedHashMap<>();
        for (String c : cols) {
            r1.put(c, "");
        }
        r1.put(ResultDispatchSchema.COL_PROCESS, "SEC");
        r1.put(ResultDispatchSchema.COL_MACHINE, "SEC機");
        r1.put("加工内容", "Coat");
        r1.put("依頼NO", "Y5-25");
        r1.put("換算数量", "6300");
        r1.put("計画合計", "6300");
        r1.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, "35");
        r1.put("メンバー名", "甲");
        r1.put(ResultDispatchSchema.COL_DISPATCH_DATE, "2026-05-18");
        r1.put(ResultDispatchSchema.COL_DISPATCH_QTY, "5700");
        LinkedHashMap<String, String> r2 = new LinkedHashMap<>(r1);
        r2.put(ResultDispatchSchema.COL_DISPATCH_TRIAL_ORDER, "36");
        r2.put("メンバー名", "乙");
        r2.put(ResultDispatchSchema.COL_DISPATCH_DATE, "2026-05-19");
        r2.put(ResultDispatchSchema.COL_DISPATCH_QTY, "600");
        rows.add(r1);
        rows.add(r2);

        List<Map<String, String>> profiles =
                ResultDispatchPivot.distinctWideTaskProfiles(
                        cols, rows, ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        assertEquals(1, profiles.size());
        Map<String, String> p0 = profiles.getFirst();
        assertEquals(
                5700.0,
                ResultDispatchPivot.sumQuantityForProfileAndDateForWideMerge(
                        rows,
                        p0,
                        LocalDate.of(2026, 5, 18),
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS),
                1e-9);
        assertEquals(
                600.0,
                ResultDispatchPivot.sumQuantityForProfileAndDateForWideMerge(
                        rows,
                        p0,
                        LocalDate.of(2026, 5, 19),
                        ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS),
                1e-9);

        ResultDispatchPivot.mergeDispatchRowsByWideIdentity(
                cols, rows, ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        assertEquals(2, rows.size());
        ResultDispatchNormalizer.normalizeInPlace(cols, rows);
    }

    @Test
    void mergeDispatchRowsByWideIdentity_joinsSameDayDuplicates() {
        List<String> cols = new ArrayList<>(ResultDispatchSchema.canonicalColumnOrder());
        List<Map<String, String>> rows = new ArrayList<>();
        LinkedHashMap<String, String> r1 = new LinkedHashMap<>();
        for (String c : cols) {
            r1.put(c, "");
        }
        r1.put(ResultDispatchSchema.COL_PROCESS, "P");
        r1.put(ResultDispatchSchema.COL_MACHINE, "M");
        r1.put("加工内容", "x");
        r1.put("依頼NO", "T1");
        r1.put("換算数量", "100");
        r1.put("計画合計", "100");
        r1.put("メンバー名", "a");
        r1.put(ResultDispatchSchema.COL_DISPATCH_DATE, "2026-05-10");
        r1.put(ResultDispatchSchema.COL_DISPATCH_QTY, "40");
        LinkedHashMap<String, String> r2 = new LinkedHashMap<>(r1);
        r2.put("メンバー名", "b");
        r2.put(ResultDispatchSchema.COL_DISPATCH_QTY, "60");
        rows.add(r1);
        rows.add(r2);
        ResultDispatchPivot.mergeDispatchRowsByWideIdentity(
                cols, rows, ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        assertEquals(1, rows.size());
        assertEquals(
                100.0,
                ResultDispatchNormalizer.parseDouble(
                        rows.getFirst().get(ResultDispatchSchema.COL_DISPATCH_QTY)),
                1e-9);
    }
}
