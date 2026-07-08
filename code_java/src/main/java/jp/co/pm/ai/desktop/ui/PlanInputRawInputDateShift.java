package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.util.List;
import java.util.Optional;

import javafx.collections.ObservableList;

/**
 * 配台計画_タスク入力: {@link #COL_RAW_INPUT_DATE} 列の一括操作（前倒し）。
 */
public final class PlanInputRawInputDateShift {

    public static final String COL_RAW_INPUT_DATE = "原反投入日";

    public static final String COL_TASK_ID = "依頼NO";

    /** 列が無いときの戻り値（{@link #applyMinusOneDayToAllRows}）。 */
    public static final int MISSING_RAW_INPUT_DATE_COLUMN = -1;

    /** @deprecated {@link #COL_RAW_INPUT_DATE} を使用。 */
    @Deprecated
    public static final String COL_RAW_INPUT_DATE_OVERRIDE = "原反投入日_上書き";

    /** @deprecated {@link #MISSING_RAW_INPUT_DATE_COLUMN} を使用。 */
    @Deprecated
    public static final int MISSING_OVERRIDE_COLUMN = MISSING_RAW_INPUT_DATE_COLUMN;

    private PlanInputRawInputDateShift() {}

    /**
     * 全行の {@link #COL_RAW_INPUT_DATE} を解釈できた行のみ 1 暦日前に更新する。
     *
     * @return 更新した行数。列が無いときは {@link #MISSING_RAW_INPUT_DATE_COLUMN}
     */
    public static int applyMinusOneDayToAllRows(
            List<String> headers, ObservableList<ObservableList<String>> rows) {
        if (headers == null || rows == null) {
            return 0;
        }
        int idxBase = headers.indexOf(COL_RAW_INPUT_DATE);
        if (idxBase < 0) {
            return MISSING_RAW_INPUT_DATE_COLUMN;
        }
        int updated = 0;
        for (ObservableList<String> row : rows) {
            String base = cellAt(row, idxBase);
            Optional<LocalDate> effective = PlanInputDateColumnSupport.parseCellValue(base);
            if (effective.isEmpty()) {
                continue;
            }
            ensureSize(row, idxBase + 1);
            row.set(
                    idxBase,
                    PlanInputDateColumnSupport.formatCellValue(effective.get().minusDays(1)));
            updated++;
        }
        return updated;
    }

    /**
     * {@code editedRowIndex} 行と同じ依頼NOの全行の {@link #COL_RAW_INPUT_DATE} を {@code newValue} に揃える。
     * 依頼NO が空のときは編集行のみ更新する。
     *
     * @return 値を書き換えた行数
     */
    public static int propagateRawInputDateToSameTaskIdRows(
            List<String> headers,
            ObservableList<ObservableList<String>> rows,
            int editedRowIndex,
            String newValue) {
        if (headers == null || rows == null || editedRowIndex < 0 || editedRowIndex >= rows.size()) {
            return 0;
        }
        int idxTid = headers.indexOf(COL_TASK_ID);
        int idxDate = headers.indexOf(COL_RAW_INPUT_DATE);
        if (idxDate < 0) {
            return 0;
        }
        String normalizedNew = newValue != null ? newValue : "";
        ObservableList<String> editedRow = rows.get(editedRowIndex);
        String taskId = idxTid >= 0 ? cellAt(editedRow, idxTid) : "";
        if (taskId.isBlank()) {
            ensureSize(editedRow, idxDate + 1);
            if (cellAt(editedRow, idxDate).equals(normalizedNew)) {
                return 0;
            }
            editedRow.set(idxDate, normalizedNew);
            return 1;
        }
        int updated = 0;
        for (ObservableList<String> row : rows) {
            if (idxTid >= 0 && !taskId.equals(cellAt(row, idxTid))) {
                continue;
            }
            ensureSize(row, idxDate + 1);
            if (cellAt(row, idxDate).equals(normalizedNew)) {
                continue;
            }
            row.set(idxDate, normalizedNew);
            updated++;
        }
        return updated;
    }

    /** @deprecated {@link #applyMinusOneDayToAllRows} を使用。 */
    @Deprecated
    public static int applyMinusOneDayToAllOverrides(
            List<String> headers, ObservableList<ObservableList<String>> rows) {
        return applyMinusOneDayToAllRows(headers, rows);
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
