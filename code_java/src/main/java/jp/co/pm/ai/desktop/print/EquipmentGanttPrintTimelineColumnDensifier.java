package jp.co.pm.ai.desktop.print;

import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.regex.Pattern;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

/**
 * 印刷向けに HH:MM タイムライン列を半開区間 {@code [start, end)} で絞り込む。
 * 左固定列・進度列は維持し、担当バッジ行はスロット列と同じインデックスで切り出す。
 */
public final class EquipmentGanttPrintTimelineColumnDensifier {

    private static final Pattern TIME_SLOT_HEADER =
            Pattern.compile("^\\s*(\\d{1,2}):(\\d{2})\\s*$");

    private EquipmentGanttPrintTimelineColumnDensifier() {}

    /**
     * 時刻範囲が無効（{@code null} または start≧end）のときは入力をそのまま返す。
     */
    public static EquipmentGanttPrintTableData densify(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            List<List<String>> badgeSlotRows,
            LocalTime rangeStartInclusive,
            LocalTime rangeEndExclusive) {
        if (columns == null || columns.isEmpty()) {
            return new EquipmentGanttPrintTableData(
                    List.of(),
                    rows != null ? rows : FXCollections.observableArrayList(),
                    badgeSlotRows);
        }
        if (rangeStartInclusive == null
                || rangeEndExclusive == null
                || !rangeStartInclusive.isBefore(rangeEndExclusive)) {
            return new EquipmentGanttPrintTableData(columns, rows, badgeSlotRows);
        }

        List<Integer> slotSourceIndices = new ArrayList<>();
        List<Integer> keptSlotPositions = new ArrayList<>();
        List<Integer> keepColumnIndices = new ArrayList<>();

        for (int c = 0; c < columns.size(); c++) {
            LocalTime slotStart = parseTimeHeader(columns.get(c));
            if (slotStart == null) {
                keepColumnIndices.add(c);
                continue;
            }
            int slotPos = slotSourceIndices.size();
            slotSourceIndices.add(c);
            if (slotInHalfOpenRange(slotStart, rangeStartInclusive, rangeEndExclusive)) {
                keepColumnIndices.add(c);
                keptSlotPositions.add(slotPos);
            }
        }

        if (keptSlotPositions.size() == slotSourceIndices.size()) {
            return new EquipmentGanttPrintTableData(columns, rows, badgeSlotRows);
        }

        List<String> newCols = new ArrayList<>(keepColumnIndices.size());
        for (int c : keepColumnIndices) {
            newCols.add(columns.get(c));
        }

        ObservableList<ObservableList<String>> newRows = FXCollections.observableArrayList();
        if (rows != null) {
            for (ObservableList<String> row : rows) {
                newRows.add(sliceRow(row, keepColumnIndices));
            }
        }

        List<List<String>> newBadges = filterBadgeSlotRows(badgeSlotRows, keptSlotPositions);

        return new EquipmentGanttPrintTableData(newCols, newRows, newBadges);
    }

    static boolean slotInHalfOpenRange(
            LocalTime slotStart, LocalTime rangeStartInclusive, LocalTime rangeEndExclusive) {
        if (slotStart == null || rangeStartInclusive == null || rangeEndExclusive == null) {
            return false;
        }
        return !slotStart.isBefore(rangeStartInclusive) && slotStart.isBefore(rangeEndExclusive);
    }

    static LocalTime parseTimeHeader(String col) {
        if (col == null) {
            return null;
        }
        var m = TIME_SLOT_HEADER.matcher(col.strip());
        if (!m.matches()) {
            return null;
        }
        int hh = Integer.parseInt(m.group(1));
        int mm = Integer.parseInt(m.group(2));
        try {
            return LocalTime.of(hh, mm);
        } catch (Exception e) {
            return null;
        }
    }

    private static ObservableList<String> sliceRow(
            ObservableList<String> row, List<Integer> keepColumnIndices) {
        if (row == null || row.isEmpty()) {
            return FXCollections.observableArrayList();
        }
        ObservableList<String> out = FXCollections.observableArrayList();
        for (int c : keepColumnIndices) {
            if (c >= 0 && c < row.size()) {
                String v = row.get(c);
                out.add(v != null ? v : "");
            } else {
                out.add("");
            }
        }
        return out;
    }

    private static List<List<String>> filterBadgeSlotRows(
            List<List<String>> badgeSlotRows, List<Integer> keptSlotPositions) {
        if (badgeSlotRows == null) {
            return null;
        }
        if (keptSlotPositions.isEmpty()) {
            return List.of();
        }
        List<List<String>> out = new ArrayList<>(badgeSlotRows.size());
        for (List<String> raw : badgeSlotRows) {
            List<String> row = new ArrayList<>(keptSlotPositions.size());
            for (int si : keptSlotPositions) {
                String s =
                        raw != null && si >= 0 && si < raw.size() && raw.get(si) != null
                                ? raw.get(si).strip()
                                : "";
                row.add(s);
            }
            out.add(Collections.unmodifiableList(row));
        }
        return out;
    }
}
