package jp.co.pm.ai.desktop.dispatch;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 結果_配台表の日別行を、依頼NO×工程名×機械名ごとに 1 行へ集約する。
 *
 * <p>加工開始日時は最早、加工終了日時は最遅。配台日は空にし、当日配台数量は合算する。
 */
public final class ResultDispatchTaskSummaryConsolidator {

    private static final String COL_TID = "依頼NO";
    private static final String COL_START = "加工開始日時";
    private static final String COL_END = "加工終了日時";

    private ResultDispatchTaskSummaryConsolidator() {}

    /** 依頼NO×工程名×機械名のタスク識別キー（集約・日別索引で共通）。 */
    public static String taskGroupKey(Map<String, String> row) {
        return groupKey(
                row,
                COL_TID,
                ResultDispatchSchema.COL_PROCESS,
                ResultDispatchSchema.COL_MACHINE);
    }

    /** 日別行をタスクキーごとに索引する。 */
    public static Map<String, List<Map<String, String>>> indexDailyRowsByTaskGroup(
            List<Map<String, String>> dailyRows) {
        Map<String, List<Map<String, String>>> out = new LinkedHashMap<>();
        if (dailyRows == null) {
            return out;
        }
        for (Map<String, String> row : dailyRows) {
            out.computeIfAbsent(taskGroupKey(row), k -> new ArrayList<>()).add(row);
        }
        return out;
    }

    /** 配台日昇順（同日は加工開始日時）で日別行を並べ替える。 */
    public static List<Map<String, String>> sortedDailyScheduleRows(
            List<Map<String, String>> groupRows) {
        if (groupRows == null || groupRows.isEmpty()) {
            return List.of();
        }
        List<Map<String, String>> copy = new ArrayList<>(groupRows);
        copy.sort(
                Comparator.comparing(
                                (Map<String, String> r) ->
                                        parseDispatchDate(r.get(ResultDispatchSchema.COL_DISPATCH_DATE)),
                                Comparator.nullsLast(Comparator.naturalOrder()))
                        .thenComparing(
                                r ->
                                        ResultDispatchDeadlineJudgment.parseDispatchDateTime(
                                                nz(r.get(COL_START))),
                                Comparator.nullsLast(Comparator.naturalOrder())));
        return copy;
    }

    public static List<Map<String, String>> consolidate(
            List<String> columns, List<Map<String, String>> rows) {
        if (rows == null || rows.isEmpty()) {
            return List.of();
        }
        String procCol = ResultDispatchSchema.COL_PROCESS;
        String machCol = ResultDispatchSchema.COL_MACHINE;
        Map<String, List<Map<String, String>>> groups = new LinkedHashMap<>();
        for (Map<String, String> row : rows) {
            String gk = groupKey(row, COL_TID, procCol, machCol);
            groups.computeIfAbsent(gk, k -> new ArrayList<>()).add(row);
        }
        List<Map<String, String>> out = new ArrayList<>(groups.size());
        for (List<Map<String, String>> group : groups.values()) {
            out.add(mergeGroup(group, columns));
        }
        return out;
    }

    private static Map<String, String> mergeGroup(
            List<Map<String, String>> group, List<String> columns) {
        Map<String, String> base = new LinkedHashMap<>(group.getFirst());
        LocalDateTime minStart = null;
        String minStartRaw = "";
        LocalDateTime maxEnd = null;
        String maxEndRaw = "";
        double qtySum = 0;
        String qtyCol = ResultDispatchSchema.COL_DISPATCH_QTY;
        String actualCol = ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL;
        boolean hasActual = columns != null && columns.contains(actualCol);
        double actualSum = 0;

        for (Map<String, String> row : group) {
            String startRaw = nz(row.get(COL_START));
            LocalDateTime start = ResultDispatchDeadlineJudgment.parseDispatchDateTime(startRaw);
            if (start != null && (minStart == null || start.isBefore(minStart))) {
                minStart = start;
                minStartRaw = startRaw;
            }
            String endRaw = nz(row.get(COL_END));
            LocalDateTime end = ResultDispatchDeadlineJudgment.parseDispatchDateTime(endRaw);
            if (end != null && (maxEnd == null || end.isAfter(maxEnd))) {
                maxEnd = end;
                maxEndRaw = endRaw;
            }
            qtySum += ResultDispatchNormalizer.parseDouble(row.get(qtyCol));
            if (hasActual) {
                actualSum += ResultDispatchNormalizer.parseDouble(row.get(actualCol));
            }
        }

        base.put(COL_START, minStartRaw);
        base.put(COL_END, maxEndRaw);
        base.put(qtyCol, ResultDispatchNormalizer.formatQty(qtySum));
        if (hasActual) {
            base.put(actualCol, ResultDispatchNormalizer.formatQty(actualSum));
        }
        base.put(ResultDispatchSchema.COL_DISPATCH_DATE, "");
        return base;
    }

    private static String groupKey(
            Map<String, String> row, String tidCol, String procCol, String machCol) {
        return nz(row.get(tidCol))
                + "\u0000"
                + nz(row.get(procCol))
                + "\u0000"
                + nz(row.get(machCol));
    }

    private static LocalDate parseDispatchDate(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        return ResultDispatchPivot.parseIsoDate(raw.strip());
    }

    private static String nz(String s) {
        return s != null ? s : "";
    }
}
