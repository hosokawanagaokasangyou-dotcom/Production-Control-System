package jp.co.pm.ai.desktop.print;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

/**
 * 設備ガント（グラフィック）の印刷用に、1 ページあたりの行インデックス束を求める。
 *
 * <p>契約 JSON 由来の表では暦日ごとに「日付」列に {@code 【yyyy/M/d】} 形式のバナー行が挿入される。
 * Excel 由来で ■ 等のセクション行がある場合も境界とする。境界ごとに 1 塊＝A3 横 1 ページ。
 * 境界行が無いときは「日付」列の繰り上がり（暦日キー変化）で塊に分ける。いずれも無ければ表全体を 1 塊。
 */
public final class EquipmentGanttPrintDaySlices {

    private static final Pattern BRACKETED_PLAIN_DATE_LABEL =
            Pattern.compile(
                    "^\\s*【\\s*(\\d{4})\\s*[/\\-]\\s*(\\d{1,2})\\s*[/\\-]\\s*(\\d{1,2})\\s*】\\s*$");

    /** Excel 等で括弧無しの {@code yyyy/M/d} を日付列に入れる場合の抽出用 */
    private static final Pattern LOOSE_YMD =
            Pattern.compile("(\\d{4})[/\\-.](\\d{1,2})[/\\-.](\\d{1,2})");

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
        if (row == null) {
            return false;
        }
        for (int c = 0; c < Math.min(4, row.size()); c++) {
            String s = row.get(c) != null ? row.get(c).strip() : "";
            if (BRACKETED_PLAIN_DATE_LABEL.matcher(s).matches()) {
                return true;
            }
        }
        int dateCol = columns != null ? columns.indexOf("日付") : -1;
        if (dateCol < 0 || dateCol >= row.size()) {
            return false;
        }
        String dv = row.get(dateCol) != null ? row.get(dateCol).strip() : "";
        if (BRACKETED_PLAIN_DATE_LABEL.matcher(dv).matches()) {
            return true;
        }
        if (!dv.isEmpty() && normalizeYmdKey(dv) != null && rowLooksLikeDayBannerRow(row, dateCol)) {
            return true;
        }
        return false;
    }

    /**
     * 日付列以外にデータがほとんど無い行を、括弧無しの暦日見出しとみなす。
     */
    private static boolean rowLooksLikeDayBannerRow(ObservableList<String> row, int dateCol) {
        int nonEmptyOthers = 0;
        for (int c = 0; c < Math.min(row.size(), 12); c++) {
            if (c == dateCol) {
                continue;
            }
            String s = row.get(c) != null ? row.get(c).strip() : "";
            if (s.isEmpty() || "—".equals(s) || "-".equals(s)) {
                continue;
            }
            nonEmptyOthers++;
        }
        return nonEmptyOthers <= 1;
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
            return groupsByCarriedCalendarChange(columns, rows);
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

    private static String[] carriedAtEachRow(
            List<String> columns, ObservableList<ObservableList<String>> rows) {
        int dateCol = columns != null ? columns.indexOf("日付") : -1;
        String[] out = new String[rows.size()];
        String cd = "";
        for (int i = 0; i < rows.size(); i++) {
            ObservableList<String> row = rows.get(i);
            if (row == null) {
                out[i] = cd;
                continue;
            }
            if (isSectionLikeRow(row)) {
                out[i] = cd;
                continue;
            }
            if (dateCol >= 0 && row.size() > dateCol) {
                String dv = row.get(dateCol) != null ? row.get(dateCol).strip() : "";
                if (!dv.isEmpty()) {
                    cd = dv;
                }
            }
            out[i] = cd;
        }
        return out;
    }

    /** 抽出できなければ null（空文字列は null） */
    private static String normalizeYmdKey(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String s = raw.strip();
        if (s.startsWith("【")) {
            int end = s.indexOf('】');
            if (end > 1) {
                s = s.substring(1, end).strip();
            }
        }
        Matcher m = LOOSE_YMD.matcher(s);
        if (m.find()) {
            try {
                int y = Integer.parseInt(m.group(1));
                int mo = Integer.parseInt(m.group(2));
                int d = Integer.parseInt(m.group(3));
                return LocalDate.of(y, mo, d).toString();
            } catch (Exception ignored) {
                return null;
            }
        }
        return null;
    }

    /**
     * 【…】や ■ 境界が無い表で、繰り上がり日付（正規化した暦日）が変わる位置でページを分ける。
     */
    private static List<List<Integer>> groupsByCarriedCalendarChange(
            List<String> columns, ObservableList<ObservableList<String>> rows) {
        String[] carried = carriedAtEachRow(columns, rows);
        List<List<Integer>> groups = new ArrayList<>();
        List<Integer> cur = new ArrayList<>();
        String prevKey = null;
        for (int i = 0; i < rows.size(); i++) {
            ObservableList<String> row = rows.get(i);
            if (isSectionLikeRow(row)) {
                if (!cur.isEmpty()) {
                    groups.add(new ArrayList<>(cur));
                    cur.clear();
                }
                prevKey = null;
                cur.add(i);
                continue;
            }
            String key = normalizeYmdKey(carried[i]);
            if (!cur.isEmpty()
                    && key != null
                    && prevKey != null
                    && !prevKey.equals(key)) {
                groups.add(new ArrayList<>(cur));
                cur.clear();
            }
            cur.add(i);
            if (key != null) {
                prevKey = key;
            }
        }
        if (!cur.isEmpty()) {
            groups.add(cur);
        }
        if (groups.isEmpty()) {
            List<Integer> all = new ArrayList<>(rows.size());
            for (int i = 0; i < rows.size(); i++) {
                all.add(i);
            }
            groups.add(all);
        }
        return groups;
    }
}
