package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.List;
import java.util.Optional;

import javafx.collections.ObservableList;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;

/**
 * 入力3表: 原反投入日 + {@link #RAW_INPUT_SAME_DAY_START} が {@code 配台可能日時} より後かを判定する。
 *
 * <p>Python {@code DISPATCHABLE_FROM_TIME}（12:45）と整合。
 */
public final class PlanInputStage3DispatchableViolationSupport {

    public static final String COL_DISPATCHABLE_DATETIME = "配台可能日時";

    /** 原反投入日と同一暦日の配台開始下限（Python DISPATCHABLE_FROM_TIME）。 */
    public static final LocalTime RAW_INPUT_SAME_DAY_START = LocalTime.of(12, 45);

    private PlanInputStage3DispatchableViolationSupport() {}

    /** 原反投入日（列 {@link PlanInputRawInputDateShift#COL_RAW_INPUT_DATE}）。 */
    public static Optional<LocalDate> effectiveRawInputDate(List<String> headers, List<String> row) {
        if (headers == null || row == null) {
            return Optional.empty();
        }
        int idxBase = headers.indexOf(PlanInputRawInputDateShift.COL_RAW_INPUT_DATE);
        return PlanInputDateColumnSupport.parseCellValue(cellAt(row, idxBase));
    }

    /**
     * {@code 原反投入日+12:45 > 配台可能日時} か（両方解釈できる行のみ）。
     */
    public static boolean isDispatchableBeforeRawInputLimit(List<String> headers, List<String> row) {
        Optional<LocalDate> rawDate = effectiveRawInputDate(headers, row);
        if (rawDate.isEmpty()) {
            return false;
        }
        int idxDispatchable = headers.indexOf(COL_DISPATCHABLE_DATETIME);
        if (idxDispatchable < 0) {
            return false;
        }
        Optional<LocalDateTime> dispatchable =
                PlanInputDateColumnSupport.parseDateTimeCellValue(cellAt(row, idxDispatchable));
        if (dispatchable.isEmpty()) {
            return false;
        }
        LocalDateTime rawLimit = rawDate.get().atTime(RAW_INPUT_SAME_DAY_START);
        return rawLimit.isAfter(dispatchable.get());
    }

    public static int countViolations(
            List<String> headers, ObservableList<ObservableList<String>> rows) {
        if (headers == null || rows == null || rows.isEmpty()) {
            return 0;
        }
        int count = 0;
        for (ObservableList<String> row : rows) {
            if (isDispatchableBeforeRawInputLimit(headers, row)) {
                count++;
            }
        }
        return count;
    }

    public static String warningMessage(int violationCount) {
        if (violationCount <= 0) {
            return "";
        }
        return violationCount
                + " 件の枝番タスクで「原反投入日+12:45」が「配台可能日時」より後になっています。"
                + "配台可能日時（手動修正表の配台日+定常開始）より早く開始できないため、原反投入日の見直しを検討してください。";
    }

    /** 違反行の {@code 配台可能日時} セル背景を赤系で強調する。 */
    public static void applyViolationHighlights(
            GridBase grid,
            List<String> headers,
            ObservableList<ObservableList<String>> rows,
            int firstDataRowIndex) {
        if (grid == null || headers == null || rows == null || rows.isEmpty()) {
            return;
        }
        int idxDispatchable = headers.indexOf(COL_DISPATCHABLE_DATETIME);
        if (idxDispatchable < 0) {
            return;
        }
        var gridRows = grid.getRows();
        for (int r = 0; r < rows.size(); r++) {
            int gridRow = firstDataRowIndex + r;
            if (gridRow < 0 || gridRow >= gridRows.size()) {
                continue;
            }
            var rowCells = gridRows.get(gridRow);
            if (idxDispatchable >= rowCells.size()) {
                continue;
            }
            SpreadsheetCell cell = rowCells.get(idxDispatchable);
            if (cell == null) {
                continue;
            }
            if (isDispatchableBeforeRawInputLimit(headers, rows.get(r))) {
                cell.setStyle(TabularCellHighlight.PLAN_INPUT_DISPATCHABLE_DATETIME_VIOLATION_STYLE);
            }
        }
    }

    private static String cellAt(List<String> row, int colIndex) {
        if (row == null || colIndex < 0 || colIndex >= row.size()) {
            return "";
        }
        String v = row.get(colIndex);
        return v != null ? v.strip() : "";
    }
}
