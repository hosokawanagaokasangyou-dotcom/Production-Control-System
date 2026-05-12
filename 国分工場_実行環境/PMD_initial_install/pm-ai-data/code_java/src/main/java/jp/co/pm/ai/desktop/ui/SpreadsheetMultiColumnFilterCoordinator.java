package jp.co.pm.ai.desktop.ui;

import java.util.BitSet;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.WeakHashMap;

import org.controlsfx.control.spreadsheet.Grid;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * ControlsFX {@link SpreadsheetView} の列フィルタを Excel と同様に複数列 AND で合成する。
 *
 * <p>各列の「表示する値」の集合を保持し、{@link SpreadsheetView#setHiddenRows} で行の表示を更新する。
 */
public final class SpreadsheetMultiColumnFilterCoordinator {

    private static final WeakHashMap<SpreadsheetView, Map<Integer, Set<String>>> COLUMN_ALLOWED =
            new WeakHashMap<>();

    /** 計画結果ビューアなど、フィルタ確定後にストレージへ保存するフック（1 ビュー 1 つ）。 */
    private static final WeakHashMap<SpreadsheetView, Runnable> COLUMN_FILTER_COMMIT_HOOK =
            new WeakHashMap<>();

    private SpreadsheetMultiColumnFilterCoordinator() {}

    /**
     * 列フィルタ適用（{@link #commitColumnSelection}）のたびに呼ぶデバウンス用フック。
     * {@link #clear} では呼ばない（再構築時の clear が保存で空上書きしないため）。
     */
    public static void setColumnFilterCommitHook(SpreadsheetView spv, Runnable hook) {
        if (spv == null) {
            return;
        }
        if (hook == null) {
            COLUMN_FILTER_COMMIT_HOOK.remove(spv);
        } else {
            COLUMN_FILTER_COMMIT_HOOK.put(spv, hook);
        }
    }

    private static void runColumnFilterCommitHook(SpreadsheetView spv) {
        Runnable r = COLUMN_FILTER_COMMIT_HOOK.get(spv);
        if (r != null) {
            r.run();
        }
    }

    /**
     * 永続化用に、列インデックスごとの許容値集合のコピーを返す（空なら空マップ）。
     */
    public static Map<Integer, Set<String>> copyColumnAllowedByIndex(SpreadsheetView spv) {
        Map<Integer, Set<String>> map = COLUMN_ALLOWED.get(spv);
        if (map == null || map.isEmpty()) {
            return Map.of();
        }
        Map<Integer, Set<String>> out = new HashMap<>();
        for (Map.Entry<Integer, Set<String>> e : map.entrySet()) {
            out.put(e.getKey(), new HashSet<>(e.getValue()));
        }
        return out;
    }

    public static void clear(SpreadsheetView spv) {
        if (spv != null) {
            COLUMN_ALLOWED.remove(spv);
        }
    }

    /**
     * 1 列ぶんのチェック状態を反映し、全列の AND に基づいて非表示行を再計算する。
     *
     * @param selectedValues その列で表示を許可するセル文字列（チェック済み集合）
     */
    public static void commitColumnSelection(SpreadsheetView spv, int column, Set<String> selectedValues) {
        Objects.requireNonNull(spv);
        Objects.requireNonNull(selectedValues);
        Set<String> universe = distinctValuesForColumnRespectingOtherFilters(spv, column);
        Map<Integer, Set<String>> map = COLUMN_ALLOWED.computeIfAbsent(spv, k -> new HashMap<>());
        if (universe.isEmpty()) {
            map.remove(column);
        } else if (universe.equals(selectedValues)) {
            map.remove(column);
        } else {
            map.put(column, new HashSet<>(selectedValues));
        }
        if (map.isEmpty()) {
            COLUMN_ALLOWED.remove(spv);
        }
        recomputeHiddenRows(spv);
        runColumnFilterCommitHook(spv);
    }

    /**
     * 指定列以外の列フィルタのみを満たす行について、当該列のセル値の集合を返す（フィルタ UI の候補用）。
     */
    public static Set<String> distinctValuesForColumnRespectingOtherFilters(SpreadsheetView spv, int column) {
        Set<String> out = new HashSet<>();
        Grid grid = spv.getGrid();
        if (grid == null || grid.getRows() == null) {
            return out;
        }
        int first = spv.getFilteredRow() + 1;
        int n = grid.getRowCount();
        for (int i = first; i < n; i++) {
            if (!rowPassesFiltersExcept(spv, i, column)) {
                continue;
            }
            SpreadsheetCell cell = grid.getRows().get(i).get(column);
            out.add(cell.getText());
        }
        return out;
    }

    static boolean rowPassesFiltersExcept(SpreadsheetView spv, int gridRow, int exceptColumn) {
        Map<Integer, Set<String>> map = COLUMN_ALLOWED.get(spv);
        if (map == null || map.isEmpty()) {
            return true;
        }
        Grid grid = spv.getGrid();
        if (grid == null || grid.getRows() == null) {
            return true;
        }
        int colCount = spv.getColumns().size();
        for (Map.Entry<Integer, Set<String>> e : map.entrySet()) {
            int col = e.getKey();
            if (col == exceptColumn) {
                continue;
            }
            if (col < 0 || col >= colCount) {
                continue;
            }
            Set<String> allowed = e.getValue();
            if (allowed == null || allowed.isEmpty()) {
                return false;
            }
            String txt = grid.getRows().get(gridRow).get(col).getText();
            if (!allowed.contains(txt)) {
                return false;
            }
        }
        return true;
    }

    static void recomputeHiddenRows(SpreadsheetView spv) {
        Map<Integer, Set<String>> map = COLUMN_ALLOWED.get(spv);
        Grid grid = spv.getGrid();
        if (grid == null || grid.getRows() == null) {
            return;
        }
        int first = spv.getFilteredRow() + 1;
        int n = grid.getRowCount();
        BitSet hidden = new BitSet(Math.max(n, spv.getHiddenRows().size()));
        if (map == null || map.isEmpty()) {
            spv.setHiddenRows(hidden);
            return;
        }
        int colCount = spv.getColumns().size();
        for (int i = first; i < n; i++) {
            boolean hide = false;
            for (Map.Entry<Integer, Set<String>> e : map.entrySet()) {
                int col = e.getKey();
                if (col < 0 || col >= colCount) {
                    continue;
                }
                Set<String> allowed = e.getValue();
                if (allowed == null || allowed.isEmpty()) {
                    hide = true;
                    break;
                }
                String txt = grid.getRows().get(i).get(col).getText();
                if (!allowed.contains(txt)) {
                    hide = true;
                    break;
                }
            }
            hidden.set(i, hide);
        }
        spv.setHiddenRows(hidden);
    }
}
