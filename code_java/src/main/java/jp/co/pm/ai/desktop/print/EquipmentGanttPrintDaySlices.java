package jp.co.pm.ai.desktop.print;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.regex.Pattern;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

/**
 * 設備ガント（グラフィック）の印刷用に、1 ページあたりの行インデックス束を求める。
 *
 * <p>契約 JSON 由来の表では暦日ごとにセクション行（区切り）が挿入されるため、セクション行を境界として
 * 塊に分ける（1 塊＝A3 横 1 ページ）。セクション行が 1 行も無いときは表全体を 1 塊とする。
 */
public final class EquipmentGanttPrintDaySlices {

    private static final Pattern BRACKETED_PLAIN_DATE_LABEL =
            Pattern.compile("^\\s*【\\s*\\d{4}[/\\-]\\d{1,2}[/\\-]\\d{1,2}\\s*】\\s*$");

    private EquipmentGanttPrintDaySlices() {}

    /**
     * {@link jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane} の {@code isSectionRow} と同一判定。
     */
    public static boolean isSectionLikeRow(ObservableList<String> row) {
        if (row == null || row.isEmpty()) {
            return false;
        }
        for (int i = 0; i < Math.min(4, row.size()); i++) {
            String s = row.get(i) != null ? row.get(i) : "";
            if (s.contains("■") || s.contains("▪")) {
                return true;
            }
            if (s.contains("【")) {
                if (BRACKETED_PLAIN_DATE_LABEL.matcher(s.strip()).matches()) {
                    continue;
                }
                return true;
            }
        }
        return false;
    }

    /**
     * 各行が属する「印刷ページ」の行インデックスリスト（先頭からの順）。
     *
     * @param rows シート行（列順は呼び出し側の表と一致）
     */
    public static List<List<Integer>> rowIndexGroupsOnePagePerDay(
            ObservableList<ObservableList<String>> rows) {
        List<List<Integer>> out = new ArrayList<>();
        if (rows == null || rows.isEmpty()) {
            return out;
        }
        boolean anySection = false;
        for (ObservableList<String> row : rows) {
            if (row != null && isSectionLikeRow(row)) {
                anySection = true;
                break;
            }
        }
        if (!anySection) {
            List<Integer> all = new ArrayList<>(rows.size());
            for (int i = 0; i < rows.size(); i++) {
                all.add(i);
            }
            out.add(all);
            return out;
        }
        List<Integer> cur = new ArrayList<>();
        for (int i = 0; i < rows.size(); i++) {
            ObservableList<String> row = rows.get(i);
            if (row != null && isSectionLikeRow(row)) {
                if (!cur.isEmpty()) {
                    out.add(cur);
                    cur = new ArrayList<>();
                }
                cur.add(i);
            } else {
                cur.add(i);
            }
        }
        if (!cur.isEmpty()) {
            out.add(cur);
        }
        return out;
    }

    public static ObservableList<ObservableList<String>> sliceRowsByIndices(
            ObservableList<ObservableList<String>> full, List<Integer> indices) {
        ObservableList<ObservableList<String>> sub = FXCollections.observableArrayList();
        if (full == null || indices == null) {
            return sub;
        }
        for (int i : indices) {
            if (i >= 0 && i < full.size()) {
                sub.add(full.get(i));
            }
        }
        return sub;
    }

    /**
     * 元の行インデックスに対応する担当バッジ行を切り出す。要素数は {@code indices} と同じ。
     *
     * @param badgeAll 元表と同じ行数のバッジグリッド（null 可）
     * @param indices 切り出す行インデックス
     * @param slotColumns スロット列数（バッジ行のパディング幅）
     */
    public static List<List<String>> sliceBadgeRowsAligned(
            List<List<String>> badgeAll, List<Integer> indices, int slotColumns) {
        int nSlots = Math.max(0, slotColumns);
        if (indices == null || indices.isEmpty()) {
            return List.of();
        }
        if (badgeAll == null) {
            return null;
        }
        List<List<String>> out = new ArrayList<>(indices.size());
        for (int ri : indices) {
            List<String> raw =
                    ri >= 0 && ri < badgeAll.size() && badgeAll.get(ri) != null
                            ? badgeAll.get(ri)
                            : List.of();
            out.add(padBadgeRow(raw, nSlots));
        }
        return out;
    }

    private static List<String> padBadgeRow(List<String> raw, int nSlots) {
        if (nSlots <= 0) {
            return raw == null ? List.of() : List.copyOf(raw);
        }
        List<String> row = new ArrayList<>(nSlots);
        for (int i = 0; i < nSlots; i++) {
            String s =
                    raw != null && i < raw.size() && raw.get(i) != null ? raw.get(i).strip() : "";
            row.add(s);
        }
        return Collections.unmodifiableList(row);
    }
}
