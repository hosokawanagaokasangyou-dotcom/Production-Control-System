package jp.co.pm.ai.desktop.io;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Transforms raw tabular rows for the Aladdin processing-plan view after {@link TaskInputSourceRawGridIo#readRaw}.
 * Documentation uses Excel-style <strong>1-based</strong> sheet row numbers; {@code rows.get(0)} is sheet row 1.
 */
public final class AladdinProcessingPlanRawSheetTransforms {

    private AladdinProcessingPlanRawSheetTransforms() {}

    /** Step 1: remove sheet rows 2-5 (indices 1-4). */
    public static void deleteSheetRows2Through5(List<? extends List<String>> rows) {
        if (rows.size() <= 5) {
            return;
        }
        rows.subList(1, 5).clear();
    }

    /**
     * Step 2: for each column, if row 6 (index 5) is blank, copy from row 7 (index 6).
     */
    public static void copyRow7IntoBlankCellsOfRow6(List<? extends List<String>> rows) {
        if (rows.size() <= 6) {
            return;
        }
        List<String> line6 = rows.get(5);
        List<String> line7 = rows.get(6);
        int cols = Math.max(line6.size(), line7.size());
        ensureWidth(line6, cols);
        ensureWidth(line7, cols);
        for (int j = 0; j < cols; j++) {
            String a = line6.get(j);
            String b = line7.get(j);
            if (isBlank(a) && !isBlank(b)) {
                line6.set(j, b);
            }
        }
    }

    /**
     * Step 3: delete every column whose row 7 (index 6) cell is exactly the processing-speed or time label,
     * then renumber headers to sequential Japanese column titles (same pattern as raw import).
     */
    public static void removeColumnsWhereRow7IsSpeedOrTime(
            List<String> headersRef, List<? extends List<String>> rows) {
        if (rows.size() <= 6) {
            return;
        }
        List<String> seventh = rows.get(6);
        List<Integer> remove = new ArrayList<>();
        for (int j = 0; j < seventh.size(); j++) {
            if (isProcessingSpeedOrTimeLabel(seventh.get(j))) {
                remove.add(j);
            }
        }
        remove.sort(Collections.reverseOrder());
        for (int j : remove) {
            for (List<String> row : rows) {
                if (j < row.size()) {
                    row.remove(j);
                }
            }
        }
        int width = rows.stream().mapToInt(List::size).max().orElse(0);
        regenerateSequentialColumnHeaders(headersRef, width);
        padAllRowsToUniformColumnCount(rows);
    }

    /** Step 4: remove sheet row 7 (index 6). */
    public static void deleteSheetRow7(List<? extends List<String>> rows) {
        if (rows.size() <= 6) {
            return;
        }
        rows.remove(6);
    }

    /** Pad every row to the same width (max column count). */
    public static void padAllRowsToUniformColumnCount(List<? extends List<String>> rows) {
        int max = rows.stream().mapToInt(List::size).max().orElse(0);
        for (List<String> row : rows) {
            ensureWidth(row, max);
        }
    }

    private static void regenerateSequentialColumnHeaders(List<String> headersRef, int columnCount) {
        headersRef.clear();
        for (int i = 0; i < columnCount; i++) {
            headersRef.add("\u5217" + (i + 1));
        }
    }

    private static void ensureWidth(List<String> row, int width) {
        while (row.size() < width) {
            row.add("");
        }
    }

    private static boolean isBlank(String s) {
        return s == null || s.trim().isEmpty();
    }

    private static boolean isProcessingSpeedOrTimeLabel(String cell) {
        if (cell == null) {
            return false;
        }
        String t = cell.trim();
        return "\u52a0\u5de5\u901f\u5ea6".equals(t) || "\u52a0\u5de5\u6642\u9593".equals(t);
    }
}
