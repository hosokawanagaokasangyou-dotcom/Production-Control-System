package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.BitSet;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
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
 * {@link #setRowTextSearchQuery} による行検索（いずれかの列への部分一致）が有効なときは、列フィルタと AND で合成する。
 * {@link #setAdditionalHiddenRows} も同じ AND に載せる。
 */
public final class SpreadsheetMultiColumnFilterCoordinator {

    private static final WeakHashMap<SpreadsheetView, Map<Integer, Set<String>>> COLUMN_ALLOWED =
            new WeakHashMap<>();

    /**
     * 行単位の全文検索（いずれかの列のセル文字列に部分一致する行だけ表示）。列フィルタと AND で合成する。
     */
    private static final WeakHashMap<SpreadsheetView, String> ROW_TEXT_SEARCH = new WeakHashMap<>();

    /** 計画結果ビューアなど、フィルタ確定後にストレージへ保存するフック（1 ビュー 1 つ）。 */
    private static final WeakHashMap<SpreadsheetView, Runnable> COLUMN_FILTER_COMMIT_HOOK =
            new WeakHashMap<>();

    /**
     * 列フィルタ以外で隠す行（配台マスタ組み合わせ表の工程+機械絞り込みなど）。{@link #recomputeHiddenRows} と AND。
     */
    private static final WeakHashMap<SpreadsheetView, BitSet> ADDITIONAL_HIDDEN_ROWS = new WeakHashMap<>();

    /**
     * この行番号以下（フィルタ行を含む）は列フィルタ・行検索でも隠さない。配台マスタの工程名・機械名行。
     */
    private static final WeakHashMap<SpreadsheetView, Integer> ALWAYS_VISIBLE_THROUGH = new WeakHashMap<>();

    private SpreadsheetMultiColumnFilterCoordinator() {}

    /**
     * 列フィルタ・行検索と AND する追加の非表示行。空または null なら解除する。
     */
    public static void setAdditionalHiddenRows(SpreadsheetView spv, BitSet extra) {
        if (spv == null) {
            return;
        }
        if (extra == null || extra.isEmpty()) {
            ADDITIONAL_HIDDEN_ROWS.remove(spv);
        } else {
            ADDITIONAL_HIDDEN_ROWS.put(spv, (BitSet) extra.clone());
        }
        recomputeHiddenRows(spv);
    }

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
     * 行検索クエリのみを解除する（列フィルタの許容値集合はそのまま）。
     *
     * <p>{@link SpreadsheetTabularSupport#clearAllFiltersAndSort} からも呼ばれる。
     */
    public static void clearRowTextSearch(SpreadsheetView spv) {
        if (spv != null) {
            ROW_TEXT_SEARCH.remove(spv);
        }
    }

    /**
     * 行検索を設定する。{@code query} が null・空白のときは解除し、列フィルタのみの表示に戻す。
     */
    public static void setRowTextSearchQuery(SpreadsheetView spv, String query) {
        Objects.requireNonNull(spv);
        String q = query == null ? "" : query.trim();
        if (q.isEmpty()) {
            ROW_TEXT_SEARCH.remove(spv);
        } else {
            ROW_TEXT_SEARCH.put(spv, q);
        }
        recomputeHiddenRows(spv);
        runColumnFilterCommitHook(spv);
    }

    /**
     * {@link #copyColumnAllowedByIndex} で得たスナップショットを再適用する。
     *
     * <p>{@link SpreadsheetTabularSupport#applyColumnFiltersWithDialog} 直後など、コーディネータが空になった
     * あとで呼び、行の並べ替えなどでグリッドだけ差し替えた場合に列フィルタを維持する。
     */
    public static void restoreColumnAllowedSnapshot(
            SpreadsheetView spv, Map<Integer, Set<String>> snapshot) {
        if (spv == null || snapshot == null || snapshot.isEmpty()) {
            return;
        }
        Map<Integer, Set<String>> map = COLUMN_ALLOWED.computeIfAbsent(spv, k -> new HashMap<>());
        map.clear();
        for (Map.Entry<Integer, Set<String>> e : snapshot.entrySet()) {
            Set<String> vals = e.getValue();
            if (vals != null && !vals.isEmpty()) {
                map.put(e.getKey(), new HashSet<>(vals));
            }
        }
        if (map.isEmpty()) {
            COLUMN_ALLOWED.remove(spv);
        }
        recomputeHiddenRows(spv);
        runColumnFilterCommitHook(spv);
    }

    /**
     * 1 列ぶんのチェック状態を反映し、全列の AND に基づいて非表示行を再計算する。
     *
     * @param selectedValues その列で表示を許可するセル文字列（チェック済み集合）
     */
    /**
     * 指定列の空欄・空白のみのセル値を持つ行を非表示にする（初期表示用）。
     *
     * <p>非空の値が 1 つも無い、または列に空欄が無いときは何もしない。
     */
    public static void applyDefaultExcludeBlankCellValues(SpreadsheetView spv, int column) {
        if (spv == null || column < 0) {
            return;
        }
        Set<String> universe = distinctValuesForColumnRespectingOtherFilters(spv, column);
        if (universe.isEmpty()) {
            return;
        }
        Set<String> allowed = new HashSet<>();
        for (String v : universe) {
            if (v != null && !v.isBlank()) {
                allowed.add(v);
            }
        }
        if (allowed.isEmpty() || allowed.size() >= universe.size()) {
            return;
        }
        commitColumnSelection(spv, column, allowed);
    }

    public static void commitColumnSelection(SpreadsheetView spv, int column, Set<String> selectedValues) {
        Objects.requireNonNull(spv);
        Objects.requireNonNull(selectedValues);
        Set<String> selected = new HashSet<>(selectedValues);
        selected.addAll(lockedColumnFilterValues(spv, column));
        Set<String> universe = distinctValuesForColumnRespectingOtherFilters(spv, column);
        Map<Integer, Set<String>> map = COLUMN_ALLOWED.computeIfAbsent(spv, k -> new HashMap<>());
        if (universe.isEmpty()) {
            map.remove(column);
        } else if (universe.equals(selected)) {
            map.remove(column);
        } else {
            map.put(column, selected);
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
        if (gridRow <= lastAlwaysVisibleGridRow(spv)) {
            return true;
        }
        if (!rowMatchesTextSearch(spv, gridRow)) {
            return false;
        }
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

    private static boolean rowMatchesTextSearch(SpreadsheetView spv, int gridRow) {
        String needle = ROW_TEXT_SEARCH.get(spv);
        if (needle == null || needle.isEmpty()) {
            return true;
        }
        Grid grid = spv.getGrid();
        if (grid == null || grid.getRows() == null) {
            return true;
        }
        int colCount = spv.getColumns().size();
        var rowCells = grid.getRows().get(gridRow);
        for (int c = 0; c < colCount && c < rowCells.size(); c++) {
            String txt = rowCells.get(c).getText();
            if (txt != null && txt.contains(needle)) {
                return true;
            }
        }
        return false;
    }

    public static void recomputeHiddenRows(SpreadsheetView spv) {
        Map<Integer, Set<String>> map = COLUMN_ALLOWED.get(spv);
        Grid grid = spv.getGrid();
        if (grid == null || grid.getRows() == null) {
            return;
        }
        Integer filtered = spv.getFilteredRow();
        int first = filtered == null ? 1 : filtered + 1;
        int lastAlways = lastAlwaysVisibleGridRow(spv);
        int n = grid.getRowCount();
        BitSet hidden = new BitSet(Math.max(n, spv.getHiddenRows().size()));
        boolean hasColumnFilters = map != null && !map.isEmpty();
        BitSet extra = ADDITIONAL_HIDDEN_ROWS.get(spv);
        int colCount = spv.getColumns().size();
        for (int i = first; i < n; i++) {
            boolean failsColumn = false;
            if (hasColumnFilters) {
                for (Map.Entry<Integer, Set<String>> e : map.entrySet()) {
                    int col = e.getKey();
                    if (col < 0 || col >= colCount) {
                        continue;
                    }
                    Set<String> allowed = e.getValue();
                    if (allowed == null || allowed.isEmpty()) {
                        failsColumn = true;
                        break;
                    }
                    String txt = grid.getRows().get(i).get(col).getText();
                    if (!allowed.contains(txt)) {
                        failsColumn = true;
                        break;
                    }
                }
            }
            boolean failsSearch = !rowMatchesTextSearch(spv, i);
            boolean extraHide = extra != null && extra.get(i);
            hidden.set(
                    i, shouldHideGridRow(i, lastAlways, failsColumn, failsSearch, extraHide));
        }
        spv.setHiddenRows(hidden);
    }

    /**
     * 値検索で候補一覧が絞られているとき、一覧に出ていない値を選択集合から除外する。
     * 例: 「W7-7」で検索し当該行だけが見えてチェックされている状態で OK したとき、
     * 非表示の他値が選択集合に残らないようにする。
     */
    static void retainSelectionToSearchVisible(
            Set<String> copySet, Iterable<String> displayedItems, String searchQuery) {
        retainSelectionToSearchVisible(copySet, displayedItems, searchQuery, Set.of());
    }

    /**
     * {@link #retainSelectionToSearchVisible(Set, Iterable, String)} に加え、タイトル行由来の値は検索で一覧から
     * 外れても選択集合に残す（常時表示行のチェックを消さない）。
     */
    static void retainSelectionToSearchVisible(
            Set<String> copySet,
            Iterable<String> displayedItems,
            String searchQuery,
            Set<String> lockedValues) {
        if (copySet == null || displayedItems == null) {
            return;
        }
        String q = searchQuery != null ? searchQuery.trim() : "";
        if (!q.isEmpty()) {
            Set<String> visible = new HashSet<>();
            for (String item : displayedItems) {
                if (item != null) {
                    visible.add(item);
                }
            }
            copySet.removeIf(item -> !visible.contains(item));
        }
        if (lockedValues != null) {
            copySet.addAll(lockedValues);
        }
    }

    /**
     * フィルタ行の直後から連続して固定されている行の末尾（フィルタ行自身を含む）。
     * 配台マスタの工程名・機械名行はここに入る。
     */
    static int lastAlwaysVisibleGridRow(int filteredRow, Iterable<Integer> fixedRows) {
        int last = filteredRow;
        if (fixedRows == null) {
            return last;
        }
        Set<Integer> set = new HashSet<>();
        for (Integer r : fixedRows) {
            if (r != null) {
                set.add(r);
            }
        }
        int expect = filteredRow + 1;
        while (set.contains(expect)) {
            last = expect;
            expect++;
        }
        return last;
    }

    static int lastAlwaysVisibleGridRow(SpreadsheetView spv) {
        if (spv == null) {
            return 0;
        }
        Integer filtered = spv.getFilteredRow();
        int first = filtered == null ? 0 : filtered;
        int fromFixed = lastAlwaysVisibleGridRow(first, spv.getFixedRows());
        Integer explicit = ALWAYS_VISIBLE_THROUGH.get(spv);
        if (explicit == null) {
            return fromFixed;
        }
        return Math.max(fromFixed, explicit);
    }

    /**
     * フィルタ行の次から {@code lastAlwaysVisible} までの非空白セル値（チェックを外せない候補）。
     */
    static Set<String> lockedColumnFilterValues(
            List<String> cellTextsByGridRow, int filteredRow, int lastAlwaysVisible) {
        Set<String> out = new HashSet<>();
        if (cellTextsByGridRow == null) {
            return out;
        }
        int from = filteredRow + 1;
        int to = Math.min(lastAlwaysVisible, cellTextsByGridRow.size() - 1);
        for (int i = from; i <= to; i++) {
            String t = cellTextsByGridRow.get(i);
            if (t != null && !t.isBlank()) {
                out.add(t);
            }
        }
        return out;
    }

    static Set<String> lockedColumnFilterValues(SpreadsheetView spv, int column) {
        Set<String> out = new HashSet<>();
        if (spv == null || column < 0) {
            return out;
        }
        Grid grid = spv.getGrid();
        if (grid == null || grid.getRows() == null) {
            return out;
        }
        Integer filtered = spv.getFilteredRow();
        int first = filtered == null ? 0 : filtered;
        int last = lastAlwaysVisibleGridRow(spv);
        int n = grid.getRowCount();
        int colCount = grid.getColumnCount();
        if (column >= colCount) {
            return out;
        }
        ArrayList<String> col = new ArrayList<>(n);
        for (int i = 0; i < n; i++) {
            var row = grid.getRows().get(i);
            String txt = "";
            if (row != null && column < row.size() && row.get(column) != null) {
                String t = row.get(column).getText();
                txt = t != null ? t : "";
            }
            col.add(txt);
        }
        out.addAll(lockedColumnFilterValues(col, first, last));
        return out;
    }

    static boolean shouldHideGridRow(
            int gridRow,
            int lastAlwaysVisible,
            boolean failsColumnFilter,
            boolean failsTextSearch,
            boolean extraHidden) {
        if (gridRow <= lastAlwaysVisible) {
            return false;
        }
        return failsColumnFilter || failsTextSearch || extraHidden;
    }

    /**
     * 列フィルタでも隠さない最後のグリッド行（フィルタ行を含む）。配台マスタのタイトル行用。
     * {@link #clear} では消さない（フィルタ再インストールと独立）。
     */
    public static void setAlwaysVisibleThroughGridRow(SpreadsheetView spv, int lastInclusive) {
        if (spv == null) {
            return;
        }
        if (lastInclusive < 0) {
            ALWAYS_VISIBLE_THROUGH.remove(spv);
            return;
        }
        ALWAYS_VISIBLE_THROUGH.put(spv, lastInclusive);
    }
}
