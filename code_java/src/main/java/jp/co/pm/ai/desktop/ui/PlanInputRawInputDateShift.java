package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.util.List;
import java.util.Optional;

import javafx.collections.ObservableList;

/**
 * 配台計画_タスク入力: 原反投入日_上書き列の一括操作（前倒し・クリア）。
 */
public final class PlanInputRawInputDateShift {

    public static final String COL_RAW_INPUT_DATE = "原反投入日";
    public static final String COL_RAW_INPUT_DATE_OVERRIDE = "原反投入日_上書き";

    /** 上書き列が無いときの戻り値（{@link #applyMinusOneDayToAllOverrides} / {@link #clearAllOverrides}）。 */
    public static final int MISSING_OVERRIDE_COLUMN = -1;

    private PlanInputRawInputDateShift() {}

    /**
     * 全行について実効原反投入日（上書きがあればそれ、無ければ {@link #COL_RAW_INPUT_DATE}）を 1 暦日前にし、
     * {@link #COL_RAW_INPUT_DATE_OVERRIDE} に書き込む。
     *
     * @return 更新した行数。上書き列が無いときは {@link #MISSING_OVERRIDE_COLUMN}
     */
    public static int applyMinusOneDayToAllOverrides(
            List<String> headers, ObservableList<ObservableList<String>> rows) {
        if (headers == null || rows == null) {
            return 0;
        }
        int idxOverride = headers.indexOf(COL_RAW_INPUT_DATE_OVERRIDE);
        if (idxOverride < 0) {
            return MISSING_OVERRIDE_COLUMN;
        }
        int idxBase = headers.indexOf(COL_RAW_INPUT_DATE);
        int updated = 0;
        for (ObservableList<String> row : rows) {
            String base = cellAt(row, idxBase);
            String override = cellAt(row, idxOverride);
            String source = !override.isBlank() ? override : base;
            Optional<LocalDate> effective = PlanInputDateColumnSupport.parseCellValue(source);
            if (effective.isEmpty()) {
                continue;
            }
            ensureSize(row, idxOverride + 1);
            row.set(
                    idxOverride,
                    PlanInputDateColumnSupport.formatCellValue(effective.get().minusDays(1)));
            updated++;
        }
        return updated;
    }

    /**
     * 全行の {@link #COL_RAW_INPUT_DATE_OVERRIDE} を空にする（既に空の行はカウントしない）。
     *
     * @return クリアした行数。上書き列が無いときは {@link #MISSING_OVERRIDE_COLUMN}
     */
    public static int clearAllOverrides(
            List<String> headers, ObservableList<ObservableList<String>> rows) {
        if (headers == null || rows == null) {
            return 0;
        }
        int idxOverride = headers.indexOf(COL_RAW_INPUT_DATE_OVERRIDE);
        if (idxOverride < 0) {
            return MISSING_OVERRIDE_COLUMN;
        }
        int cleared = 0;
        for (ObservableList<String> row : rows) {
            String override = cellAt(row, idxOverride);
            if (override.isBlank()) {
                continue;
            }
            ensureSize(row, idxOverride + 1);
            row.set(idxOverride, "");
            cleared++;
        }
        return cleared;
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
