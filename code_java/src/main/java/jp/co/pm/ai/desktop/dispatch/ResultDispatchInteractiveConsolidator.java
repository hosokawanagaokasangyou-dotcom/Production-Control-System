package jp.co.pm.ai.desktop.dispatch;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 段階3試行後の結果_配台表（配台結果タブ表示用）で、編集目標行とタイムライン実績行の分断・重複を解消する。
 */
public final class ResultDispatchInteractiveConsolidator {

    private static final double EPS = 1e-6;

    private ResultDispatchInteractiveConsolidator() {}

    /**
     * {@link ResultDispatchSchema#COL_DISPATCH_QTY_ACTUAL} 列があるときのみ実行。
     *
     * <ul>
     *   <li>編集目標と時刻が同一行にあるとき、当日=0 の別暦日実績行は除く（Y5-15 型）。
     *   <li>編集目標が時刻無し孤立行のみのとき、実績暦日行へ当日を実配台から揃え孤立目標を除く（W5-13 型）。
     *   <li>時刻の無い孤立目標行は、実績行と共存するとき除く。
     * </ul>
     */
    public static void consolidatePlanAndTimelineRowsInPlace(
            List<String> columns, List<Map<String, String>> rows) {
        if (rows.isEmpty() || columns == null || columns.isEmpty()) {
            return;
        }
        if (!columns.contains(ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL)) {
            return;
        }
        String qtyCol = ResultDispatchSchema.COL_DISPATCH_QTY;
        String actualCol = ResultDispatchSchema.COL_DISPATCH_QTY_ACTUAL;
        String tidCol = "依頼NO";
        String procCol = ResultDispatchSchema.COL_PROCESS;
        String machCol = ResultDispatchSchema.COL_MACHINE;
        String startCol = "加工開始日時";

        Map<String, List<Integer>> groups = new LinkedHashMap<>();
        for (int i = 0; i < rows.size(); i++) {
            Map<String, String> row = rows.get(i);
            String gk = groupKey(row, tidCol, procCol, machCol);
            groups.computeIfAbsent(gk, k -> new ArrayList<>()).add(i);
        }

        List<Integer> toRemove = new ArrayList<>();
        int syncedPlanFromActual = 0;
        int removedTimelineOverflow = 0;
        for (List<Integer> idxs : groups.values()) {
            List<Integer> anchoredPlanIdx = new ArrayList<>();
            List<Integer> timelineOnlyIdx = new ArrayList<>();
            List<Integer> orphanPlanIdx = new ArrayList<>();
            boolean anyTimelineMeta = false;

            for (int i : idxs) {
                Map<String, String> row = rows.get(i);
                double plan = ResultDispatchNormalizer.parseDouble(row.get(qtyCol));
                double actual = ResultDispatchNormalizer.parseDouble(row.get(actualCol));
                boolean hasStart = !nz(row.get(startCol)).isEmpty();
                boolean hasTimelineMeta = actual > EPS || hasStart;
                if (hasTimelineMeta) {
                    anyTimelineMeta = true;
                }
                if (plan > EPS) {
                    if (hasTimelineMeta) {
                        anchoredPlanIdx.add(i);
                    } else {
                        orphanPlanIdx.add(i);
                    }
                } else if (hasTimelineMeta) {
                    timelineOnlyIdx.add(i);
                }
            }

            if (!anchoredPlanIdx.isEmpty() && !timelineOnlyIdx.isEmpty()) {
                toRemove.addAll(timelineOnlyIdx);
                removedTimelineOverflow += timelineOnlyIdx.size();
            } else if (anchoredPlanIdx.isEmpty() && !timelineOnlyIdx.isEmpty()) {
                for (int i : timelineOnlyIdx) {
                    Map<String, String> row = rows.get(i);
                    double actual = ResultDispatchNormalizer.parseDouble(row.get(actualCol));
                    double plan = ResultDispatchNormalizer.parseDouble(row.get(qtyCol));
                    if (actual > EPS && plan <= EPS) {
                        row.put(qtyCol, ResultDispatchNormalizer.formatQty(actual));
                        syncedPlanFromActual++;
                    }
                }
            }

            if (anyTimelineMeta && !orphanPlanIdx.isEmpty()) {
                toRemove.addAll(orphanPlanIdx);
            }
        }
        // #region agent log
        if (syncedPlanFromActual > 0 || removedTimelineOverflow > 0 || !toRemove.isEmpty()) {
            List<Map<String, Object>> y515 = sampleY515Slice(rows, tidCol, procCol, qtyCol, actualCol, startCol);
            if (!y515.isEmpty()) {
                jp.co.pm.ai.desktop.debug.AgentDebugLog.appendStructured(
                        Map.of(),
                        "1002c2",
                        "H2",
                        "ResultDispatchInteractiveConsolidator",
                        "consolidate done",
                        Map.of(
                                "syncedPlanFromActual",
                                syncedPlanFromActual,
                                "removedTimelineOverflow",
                                removedTimelineOverflow,
                                "removedOrphans",
                                toRemove.size(),
                                "y5_15_after",
                                y515));
            }
        }
        // #endregion
        if (toRemove.isEmpty()) {
            return;
        }
        toRemove.sort((a, b) -> Integer.compare(b, a));
        for (int i : toRemove) {
            rows.remove(i);
        }
    }

    private static List<Map<String, Object>> sampleY515Slice(
            List<Map<String, String>> rows,
            String tidCol,
            String procCol,
            String qtyCol,
            String actualCol,
            String startCol) {
        List<Map<String, Object>> y515 = new ArrayList<>();
        for (Map<String, String> row : rows) {
            String tid = nz(row.get(tidCol));
            if (!tid.contains("Y5-15") && !tid.contains("5-15")) {
                continue;
            }
            if (!"スライス".equals(nz(row.get(procCol)))) {
                continue;
            }
            y515.add(
                    Map.of(
                            "配台日",
                            nz(row.get(ResultDispatchSchema.COL_DISPATCH_DATE)),
                            "当日",
                            nz(row.get(qtyCol)),
                            "実配台",
                            nz(row.get(actualCol)),
                            "開始",
                            nz(row.get(startCol))));
        }
        return y515;
    }

    private static String groupKey(
            Map<String, String> row, String tidCol, String procCol, String machCol) {
        return nz(row.get(tidCol))
                + "\u0000"
                + nz(row.get(procCol))
                + "\u0000"
                + nz(row.get(machCol));
    }

    private static String nz(String s) {
        return s != null ? s : "";
    }
}
