package jp.co.pm.ai.desktop.dispatch;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 段階2（Python {@code planning_core._core.RESULT_DISPATCH_TABLE_STATIC_HEADERS} + 配台日／当日配台数量）が
 * 前提とする列が {@link ResultDispatchDocument} に揃っているか検査し、欠けていれば列を追加して順序を正規化する。
 */
public final class ResultDispatchStage2ColumnSupport {

    private ResultDispatchStage2ColumnSupport() {}

    /**
     * 必須列を補完し、列リストを正規順（静的列＋配台日＋当日配台数量）に揃えたうえで、正に無い列は末尾に維持する。
     *
     * @return 列リスト・行マップに変更があったとき {@code true}
     */
    public static boolean ensureStage2RequiredColumns(ResultDispatchDocument doc) {
        if (doc == null) {
            return false;
        }
        return ensureStage2RequiredColumns(doc.columns(), doc.rows());
    }

    /**
     * @return 変更があったとき {@code true}
     */
    public static boolean ensureStage2RequiredColumns(List<String> columns, List<Map<String, String>> rows) {
        if (columns == null || rows == null) {
            return false;
        }
        List<String> canonical = ResultDispatchSchema.canonicalColumnOrder();
        Set<String> canonicalSet = new LinkedHashSet<>(canonical);
        Set<String> oldHeaderSet = new HashSet<>();
        for (String c : columns) {
            if (c != null && !c.isBlank()) {
                oldHeaderSet.add(c.strip());
            }
        }
        int missingFromCanonical = 0;
        for (String c : canonical) {
            if (!oldHeaderSet.contains(c)) {
                missingFromCanonical++;
            }
        }
        List<String> newOrder = new ArrayList<>(canonical.size() + columns.size());
        newOrder.addAll(canonical);
        for (String c : columns) {
            if (c == null || c.isBlank()) {
                continue;
            }
            String t = c.strip();
            if (!canonicalSet.contains(t)) {
                newOrder.add(t);
            }
        }
        boolean reorder = !newOrder.equals(columns);
        if (reorder) {
            columns.clear();
            columns.addAll(newOrder);
        }
        boolean rowPatched = false;
        for (Map<String, String> row : rows) {
            for (String col : newOrder) {
                if (!row.containsKey(col)) {
                    row.put(col, "");
                    rowPatched = true;
                }
            }
        }
        boolean changed = reorder || missingFromCanonical > 0 || rowPatched;
        if (changed) {
            ResultDispatchNormalizer.normalizeInPlace(columns, rows);
        }
        return changed;
    }
}
