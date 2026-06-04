package jp.co.pm.ai.desktop.ui;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javafx.collections.ObservableList;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * 配台計画タスク入力: 「元の値から書き換えたセル」のマークを JSON sidecar に保存し、着色する。
 *
 * <p>旧来の {@code *_上書き} / {@code （元）*_上書き} 列の代替。基底列を直接編集し、編集したセルのみ
 * (行キー × 列見出し) でマークする。行キーは {@code 依頼NO|工程名|機械名}。
 */
public final class PlanInputEditedCellMarks {

    /** sidecar JSON のキー連結子（セル値には現れない制御文字）。 */
    public static final String SEP = "\u0001";

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final String SIDECAR_SUFFIX = ".editmarks.json";
    private static final String FIELD_MARKS = "marks";

    /** 行同定に使う列（存在するものだけ連結）。 */
    private static final List<String> ROW_KEY_COLUMNS = List.of("依頼NO", "工程名", "機械名");

    private PlanInputEditedCellMarks() {}

    public static Path sidecarPath(Path planInput) {
        return sidecarPath(planInput, "");
    }

    /**
     * @param namespace 同一ファイル内で別シートのマークを分けるための接尾辞（空可）。
     */
    public static Path sidecarPath(Path planInput, String namespace) {
        if (planInput == null) {
            return null;
        }
        String name = planInput.getFileName() != null ? planInput.getFileName().toString() : "";
        String ns = namespace != null && !namespace.isBlank() ? "." + namespace.strip() : "";
        return planInput.resolveSibling(name + ns + SIDECAR_SUFFIX);
    }

    /** 行同定キー（同定列がすべて空のときは空文字＝マーク対象外）。 */
    public static String rowKey(List<String> headers, List<String> row) {
        if (headers == null || row == null) {
            return "";
        }
        List<String> parts = new ArrayList<>();
        boolean anyValue = false;
        for (String col : ROW_KEY_COLUMNS) {
            int idx = headers.indexOf(col);
            String v = idx >= 0 ? cellAt(row, idx) : "";
            parts.add(v);
            if (!v.isBlank()) {
                anyValue = true;
            }
        }
        return anyValue ? String.join(SEP, parts) : "";
    }

    public static String markKey(String rowKey, String columnTitle) {
        return rowKey + SEP + (columnTitle != null ? columnTitle : "");
    }

    /** 読込直後の全セル値を markKey→値で記録（差分判定の基準）。 */
    public static Map<String, String> captureBaseline(
            List<String> headers, ObservableList<ObservableList<String>> rows) {
        Map<String, String> baseline = new LinkedHashMap<>();
        if (headers == null || rows == null) {
            return baseline;
        }
        for (ObservableList<String> row : rows) {
            String rk = rowKey(headers, row);
            if (rk.isEmpty()) {
                continue;
            }
            for (int c = 0; c < headers.size(); c++) {
                baseline.put(markKey(rk, headers.get(c)), cellAt(row, c));
            }
        }
        return baseline;
    }

    /**
     * 現在の表を走査し、基準値と異なるセルを {@code editedMarks} に加える。
     * 基準値に戻ったセルは、読込時 JSON に無かったものだけ外す（保存済みマークは保持）。
     */
    public static void recompute(
            List<String> headers,
            ObservableList<ObservableList<String>> rows,
            Map<String, String> baseline,
            Set<String> persistedAtLoad,
            Set<String> editedMarks) {
        if (headers == null || rows == null || baseline == null || editedMarks == null) {
            return;
        }
        for (ObservableList<String> row : rows) {
            String rk = rowKey(headers, row);
            if (rk.isEmpty()) {
                continue;
            }
            for (int c = 0; c < headers.size(); c++) {
                String key = markKey(rk, headers.get(c));
                if (!baseline.containsKey(key)) {
                    continue;
                }
                String base = baseline.get(key);
                String cur = cellAt(row, c);
                if (!equalsNorm(cur, base)) {
                    editedMarks.add(key);
                } else if (persistedAtLoad == null || !persistedAtLoad.contains(key)) {
                    editedMarks.remove(key);
                }
            }
        }
    }

    /** 現在の表に存在する行のマークだけ残す（消えた行のマークは捨てる）。 */
    public static Set<String> filterToPresentRows(
            List<String> headers, ObservableList<ObservableList<String>> rows, Set<String> marks) {
        Set<String> out = new LinkedHashSet<>();
        if (marks == null || marks.isEmpty() || headers == null || rows == null) {
            return out;
        }
        Set<String> presentRowKeys = new LinkedHashSet<>();
        for (ObservableList<String> row : rows) {
            String rk = rowKey(headers, row);
            if (!rk.isEmpty()) {
                presentRowKeys.add(rk);
            }
        }
        for (String mk : marks) {
            int sep = mk.lastIndexOf(SEP);
            String rk = sep >= 0 ? mk.substring(0, sep) : mk;
            if (presentRowKeys.contains(rk)) {
                out.add(mk);
            }
        }
        return out;
    }

    /** 編集済みセルの背景を着色する。 */
    public static void applyHighlights(
            GridBase grid,
            List<String> headers,
            ObservableList<ObservableList<String>> rows,
            int firstDataRowIndex,
            Set<String> editedMarks) {
        if (grid == null || headers == null || rows == null || editedMarks == null
                || editedMarks.isEmpty()) {
            return;
        }
        var gridRows = grid.getRows();
        for (int r = 0; r < rows.size(); r++) {
            int gridRow = firstDataRowIndex + r;
            if (gridRow < 0 || gridRow >= gridRows.size()) {
                continue;
            }
            ObservableList<String> row = rows.get(r);
            String rk = rowKey(headers, row);
            if (rk.isEmpty()) {
                continue;
            }
            var rowCells = gridRows.get(gridRow);
            for (int c = 0; c < headers.size() && c < rowCells.size(); c++) {
                if (!editedMarks.contains(markKey(rk, headers.get(c)))) {
                    continue;
                }
                String columnTitle = headers.get(c);
                if (shouldPreservePlanInputExcludeYesStyle(columnTitle, cellAt(row, c))) {
                    continue;
                }
                SpreadsheetCell cell = rowCells.get(c);
                if (cell != null) {
                    cell.setStyle(TabularCellHighlight.PLAN_INPUT_EDITED_CELL_STYLE);
                }
            }
        }
    }

    public static Set<String> load(Path planInput) {
        return load(planInput, "");
    }

    public static Set<String> load(Path planInput, String namespace) {
        Set<String> out = new LinkedHashSet<>();
        Path sidecar = sidecarPath(planInput, namespace);
        if (sidecar == null || !Files.isRegularFile(sidecar)) {
            return out;
        }
        try {
            JsonNode root = JSON.readTree(sidecar.toFile());
            JsonNode marks = root.get(FIELD_MARKS);
            if (marks != null && marks.isArray()) {
                for (JsonNode m : marks) {
                    String s = m.asText("");
                    if (!s.isEmpty()) {
                        out.add(s);
                    }
                }
            }
        } catch (Exception ignored) {
            // 壊れた sidecar は無視（マーク無し扱い）
        }
        return out;
    }

    public static void save(Path planInput, Set<String> marks) {
        save(planInput, marks, "");
    }

    public static void save(Path planInput, Set<String> marks, String namespace) {
        Path sidecar = sidecarPath(planInput, namespace);
        if (sidecar == null) {
            return;
        }
        try {
            if (marks == null || marks.isEmpty()) {
                Files.deleteIfExists(sidecar);
                return;
            }
            ObjectNode root = JSON.createObjectNode();
            ArrayNode arr = root.putArray(FIELD_MARKS);
            for (String mk : marks) {
                arr.add(mk);
            }
            if (sidecar.getParent() != null) {
                Files.createDirectories(sidecar.getParent());
            }
            JSON.writerWithDefaultPrettyPrinter().writeValue(sidecar.toFile(), root);
        } catch (Exception ignored) {
            // sidecar 失敗で表編集は止めない
        }
    }

    /**
     * 編集マークの薄黄は {@link TabularCellHighlight#PLAN_INPUT_EXCLUDE_YES_STYLE} より弱い。
     * 「配台不要」オンセルは buildPlanInputGrid の赤を維持する。
     */
    static boolean shouldPreservePlanInputExcludeYesStyle(String columnTitle, String cellValue) {
        return "配台不要".equals(columnTitle)
                && TabularCellHighlight.planInputExcludeFromAssignmentIsOn(cellValue);
    }

    private static boolean equalsNorm(String a, String b) {
        return nz(a).equals(nz(b));
    }

    private static String nz(String s) {
        return s != null ? s.strip() : "";
    }

    private static String cellAt(List<String> row, int colIndex) {
        if (row == null || colIndex < 0 || colIndex >= row.size()) {
            return "";
        }
        String v = row.get(colIndex);
        return v != null ? v.strip() : "";
    }
}
