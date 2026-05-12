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
 * <p>契約 JSON 由来の表では暦日ごとに「日付」列に {@code 【yyyy/M/d】} 形式のバナー行が挿入される。
 * Excel 由来で ■ 等のセクション行がある場合も境界とする。境界ごとに 1 塊＝A3 横 1 ページ。
 * 境界行が 1 行も無いときは表全体を 1 塊とする。
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
     * 暦日ごとの印刷ページの境界行。{@link jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane} のセクション行に加え、
     * 設備ガント契約 JSON が挿入する「日付」列のみの {@code 【yyyy/M/d】} バナー行も境界とする。
     */
    public static boolean isPrintDayBoundaryRow(
            List<String> columns, ObservableList<String> row) {
        if (isSectionLikeRow(row)) {
            return true;
        }
        int dateCol = columns != null ? columns.indexOf("日付") : -1;
        if (dateCol < 0 || row == null || dateCol >= row.size()) {
            return false;
        }
        String dv = row.get(dateCol) != null ? row.get(dateCol).strip() : "";
        return BRACKETED_PLAIN_DATE_LABEL.matcher(dv).matches();
    }

    /**
     * 各行が属する「印刷ページ」の行インデックスリスト（先頭からの順）。
     *
     * @param columns シート列見出し（「日付」列で暦日バナーを検出する）
     * @param rows シート行（列順は {@code columns} と一致）
     */
    public static List<List<Integer>> rowIndexGroupsOnePagePerDay(
            List<String> columns, ObservableList<ObservableList<String>> rows) {
        List<List<Integer>> out = new ArrayList<>();
        if (rows == null || rows.isEmpty()) {
            return out;
        }
        boolean anyBoundary = false;
        for (ObservableList<String> row : rows) {
            if (isPrintDayBoundaryRow(columns, row)) {
                anyBoundary = true;
                break;
            }
        }
        if (!anyBoundary) {
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
            if (isPrintDayBoundaryRow(columns, row)) {
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
