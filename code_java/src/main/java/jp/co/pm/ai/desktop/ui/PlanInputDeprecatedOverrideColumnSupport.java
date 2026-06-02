package jp.co.pm.ai.desktop.ui;

import java.text.Normalizer;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

import javafx.collections.ObservableList;

/**
 * 配台計画タスク入力: 廃止した {@code *_上書き} / {@code （元）…} 参照列の読込マージと UI からの除去。
 *
 * <p>基底列（例: {@code 原反投入日}）を直接編集する運用に統一する。
 */
public final class PlanInputDeprecatedOverrideColumnSupport {

    private static final String OVERRIDE_SUFFIX = "_上書き";

    private static final Map<String, String> OVERRIDE_TO_BASE =
            Map.of(
                    "加工速度_上書き", "加工速度",
                    "原反投入日_上書き", "原反投入日",
                    "配台可能日時_上書き", "配台可能日時");

    private PlanInputDeprecatedOverrideColumnSupport() {}

    /** {@code （元）加工速度_上書き} 等の参照列（{@link #isOriginalReferenceColumn} の部分集合）。 */
    public static boolean isDeprecatedReferenceOverrideColumn(String columnTitle) {
        if (columnTitle == null) {
            return false;
        }
        String h = normalize(columnTitle);
        return isOriginalReferenceColumn(h) && h.endsWith(OVERRIDE_SUFFIX);
    }

    /** {@code （元）担当OP_指定} 等、見出しが {@code （元）} / {@code (元)} で始まる参照列。 */
    public static boolean isOriginalReferenceColumn(String columnTitle) {
        if (columnTitle == null) {
            return false;
        }
        String h = normalize(columnTitle);
        return h.startsWith("(元)") || h.startsWith("（元）");
    }

    /** {@code 加工速度_上書き} 等の上書き列（参照列は {@link #isDeprecatedReferenceOverrideColumn}）。 */
    public static boolean isDeprecatedOverrideColumn(String columnTitle) {
        if (columnTitle == null) {
            return false;
        }
        String h = normalize(columnTitle);
        if (isDeprecatedReferenceOverrideColumn(h)) {
            return false;
        }
        return h.endsWith(OVERRIDE_SUFFIX) && OVERRIDE_TO_BASE.containsKey(h);
    }

    public static boolean isDeprecatedOverrideOrReferenceColumn(String columnTitle) {
        return isDeprecatedOverrideColumn(columnTitle)
                || isOriginalReferenceColumn(columnTitle);
    }

    /**
     * 上書き列の非空値を対応する基底列へ移し、廃止列を headers / rows から削除する。
     *
     * @return 削除した列数
     */
    public static int migrateAndDropDeprecatedOverrideColumns(
            List<String> headers, ObservableList<ObservableList<String>> rows) {
        if (headers == null || headers.isEmpty() || rows == null) {
            return 0;
        }
        mergeOverrideValuesIntoBaseColumns(headers, rows);
        List<Integer> dropIdx = new ArrayList<>();
        for (int c = 0; c < headers.size(); c++) {
            if (isDeprecatedOverrideOrReferenceColumn(headers.get(c))) {
                dropIdx.add(c);
            }
        }
        if (dropIdx.isEmpty()) {
            return 0;
        }
        dropIdx.sort(Comparator.reverseOrder());
        for (int col : dropIdx) {
            headers.remove(col);
            for (ObservableList<String> row : rows) {
                if (col < row.size()) {
                    row.remove(col);
                }
            }
        }
        return dropIdx.size();
    }

    private static void mergeOverrideValuesIntoBaseColumns(
            List<String> headers, ObservableList<ObservableList<String>> rows) {
        for (int c = 0; c < headers.size(); c++) {
            String title = headers.get(c);
            if (!isDeprecatedOverrideColumn(title)) {
                continue;
            }
            String baseTitle = OVERRIDE_TO_BASE.get(normalize(title));
            if (baseTitle == null) {
                continue;
            }
            int baseIdx = headers.indexOf(baseTitle);
            if (baseIdx < 0) {
                continue;
            }
            for (ObservableList<String> row : rows) {
                String overrideVal = cellAt(row, c);
                if (overrideVal.isBlank()) {
                    continue;
                }
                ensureSize(row, Math.max(c, baseIdx) + 1);
                String baseVal = cellAt(row, baseIdx);
                if (baseVal.isBlank()) {
                    row.set(baseIdx, overrideVal);
                }
            }
        }
    }

    private static String normalize(String title) {
        return Normalizer.normalize(title.strip(), Normalizer.Form.NFKC);
    }

    private static String cellAt(ObservableList<String> row, int colIndex) {
        if (row == null || colIndex < 0 || colIndex >= row.size()) {
            return "";
        }
        String v = row.get(colIndex);
        return v != null ? v.strip() : "";
    }

    private static void ensureSize(ObservableList<String> row, int size) {
        while (row.size() < size) {
            row.add("");
        }
    }
}
