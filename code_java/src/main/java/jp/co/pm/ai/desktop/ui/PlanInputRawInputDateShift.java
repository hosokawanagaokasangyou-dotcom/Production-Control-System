package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.List;
import java.util.Optional;
import java.util.Set;

import javafx.collections.ObservableList;

/**
 * 配台計画_タスク入力: {@link #COL_RAW_INPUT_DATE} 列の一括操作（前倒し）。
 *
 * <p>段階2は原反投入日＋在庫場所由来の配台可能時刻を正とするため、原反投入日を変更したときは
 * {@link #COL_DISPATCHABLE_DATETIME} も追従させる（湖南工場内在庫 K/湖南は 8:45、他は 12:45）。
 */
public final class PlanInputRawInputDateShift {

    public static final String COL_RAW_INPUT_DATE = "原反投入日";

    public static final String COL_DISPATCHABLE_DATETIME = "配台可能日時";

    public static final String COL_STOCK_LOCATION = "在庫場所";

    public static final String COL_TASK_ID = "依頼NO";

    /** 列が無いときの戻り値（{@link #applyMinusOneDayToAllRows}）。 */
    public static final int MISSING_RAW_INPUT_DATE_COLUMN = -1;

    private static final LocalTime DEFAULT_DISPATCHABLE_TIME = LocalTime.of(12, 45);

    /** 湖南工場内在庫（在庫場所 K / 湖南）の同日配台開始。Python DISPATCHABLE_FROM_TIME_KONAN_STOCK 既定と揃える。 */
    private static final LocalTime KONAN_LOCAL_STOCK_DISPATCHABLE_TIME = LocalTime.of(8, 45);

    /** @deprecated {@link #COL_RAW_INPUT_DATE} を使用。 */
    @Deprecated
    public static final String COL_RAW_INPUT_DATE_OVERRIDE = "原反投入日_上書き";

    /** @deprecated {@link #MISSING_RAW_INPUT_DATE_COLUMN} を使用。 */
    @Deprecated
    public static final int MISSING_OVERRIDE_COLUMN = MISSING_RAW_INPUT_DATE_COLUMN;

    private PlanInputRawInputDateShift() {}

    /**
     * 全行の {@link #COL_RAW_INPUT_DATE} を解釈できた行のみ 1 暦日前に更新する。
     * {@link #COL_DISPATCHABLE_DATETIME} がある行は同日へ追従する。
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
            LocalDate newDate = effective.get().minusDays(1);
            ensureSize(row, idxBase + 1);
            row.set(idxBase, PlanInputDateColumnSupport.formatCellValue(newDate));
            syncDispatchableDatetimeOnRow(headers, row, newDate);
            updated++;
        }
        return updated;
    }

    /**
     * {@code editedRowIndex} 行と同じ依頼NOの全行の {@link #COL_RAW_INPUT_DATE} を {@code newValue} に揃える。
     * 依頼NO が空のときは編集行のみ更新する。{@link #COL_DISPATCHABLE_DATETIME} も追従する。
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
        Optional<LocalDate> newDate = PlanInputDateColumnSupport.parseCellValue(normalizedNew);
        ObservableList<String> editedRow = rows.get(editedRowIndex);
        String taskId = idxTid >= 0 ? cellAt(editedRow, idxTid) : "";
        if (taskId.isBlank()) {
            ensureSize(editedRow, idxDate + 1);
            boolean rawChanged = !cellAt(editedRow, idxDate).equals(normalizedNew);
            if (rawChanged) {
                editedRow.set(idxDate, normalizedNew);
            }
            boolean synced =
                    newDate
                            .map(d -> syncDispatchableDatetimeOnRow(headers, editedRow, d))
                            .orElse(false);
            return (rawChanged || synced) ? 1 : 0;
        }
        int updated = 0;
        for (ObservableList<String> row : rows) {
            if (idxTid >= 0 && !taskId.equals(cellAt(row, idxTid))) {
                continue;
            }
            ensureSize(row, idxDate + 1);
            boolean rawChanged = !cellAt(row, idxDate).equals(normalizedNew);
            if (rawChanged) {
                row.set(idxDate, normalizedNew);
            }
            boolean synced =
                    newDate
                            .map(d -> syncDispatchableDatetimeOnRow(headers, row, d))
                            .orElse(false);
            if (rawChanged || synced) {
                updated++;
            }
        }
        return updated;
    }

    /**
     * 原反投入日だけ編集マークがあり配台可能日時が暦日ずれしている行を、原反投入日へ追従させる。
     *
     * <p>配台可能日時側に編集マークがある行は意図的な遅延とみなし触らない。
     *
     * @return 更新した行数
     */
    public static int resyncStaleDispatchableDatetimeFromRawInput(
            List<String> headers,
            ObservableList<ObservableList<String>> rows,
            Set<String> editedMarks) {
        if (headers == null || rows == null || editedMarks == null || editedMarks.isEmpty()) {
            return 0;
        }
        int idxRaw = headers.indexOf(COL_RAW_INPUT_DATE);
        int idxDisp = headers.indexOf(COL_DISPATCHABLE_DATETIME);
        if (idxRaw < 0 || idxDisp < 0) {
            return 0;
        }
        int updated = 0;
        for (ObservableList<String> row : rows) {
            String rk = PlanInputEditedCellMarks.rowKey(headers, row);
            if (rk.isEmpty()) {
                continue;
            }
            String rawMark = PlanInputEditedCellMarks.markKey(rk, COL_RAW_INPUT_DATE);
            String dispMark = PlanInputEditedCellMarks.markKey(rk, COL_DISPATCHABLE_DATETIME);
            if (!editedMarks.contains(rawMark) || editedMarks.contains(dispMark)) {
                continue;
            }
            Optional<LocalDate> rawDate =
                    PlanInputDateColumnSupport.parseCellValue(cellAt(row, idxRaw));
            if (rawDate.isEmpty()) {
                continue;
            }
            Optional<LocalDateTime> disp =
                    PlanInputDateColumnSupport.parseDateTimeCellValue(cellAt(row, idxDisp));
            if (disp.isPresent() && disp.get().toLocalDate().equals(rawDate.get())) {
                continue;
            }
            syncDispatchableDatetimeOnRow(headers, row, rawDate.get());
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

    /**
     * @return セルを書き換えたとき {@code true}
     */
    static boolean syncDispatchableDatetimeOnRow(
            List<String> headers, ObservableList<String> row, LocalDate rawDate) {
        if (headers == null || row == null || rawDate == null) {
            return false;
        }
        int idxDisp = headers.indexOf(COL_DISPATCHABLE_DATETIME);
        if (idxDisp < 0) {
            return false;
        }
        LocalTime time = dispatchableTimeForStock(headers, row);
        String next =
                PlanInputDateColumnSupport.formatDateTimeCellValue(rawDate.atTime(time));
        ensureSize(row, idxDisp + 1);
        if (cellAt(row, idxDisp).equals(next.strip())) {
            return false;
        }
        row.set(idxDisp, next);
        return true;
    }

    /** 湖南工場内在庫（在庫場所に「湖南」、またはコード K）なら 8:45、それ以外は 12:45。 */
    static LocalTime dispatchableTimeForStock(List<String> headers, ObservableList<String> row) {
        if (headers == null || row == null) {
            return DEFAULT_DISPATCHABLE_TIME;
        }
        int idxStock = headers.indexOf(COL_STOCK_LOCATION);
        if (idxStock >= 0 && isKonanLocalStock(cellAt(row, idxStock))) {
            return KONAN_LOCAL_STOCK_DISPATCHABLE_TIME;
        }
        return DEFAULT_DISPATCHABLE_TIME;
    }

    static boolean isKonanLocalStock(String stockLocation) {
        if (stockLocation == null) {
            return false;
        }
        String s = stockLocation.strip();
        if (s.isEmpty()) {
            return false;
        }
        if (s.contains("湖南")) {
            return true;
        }
        return "K".equalsIgnoreCase(s);
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
