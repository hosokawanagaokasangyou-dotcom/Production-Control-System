package jp.co.pm.ai.desktop.io.gantt;

import java.time.LocalTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

/** 設備ガント表の HH:MM 時刻列インデックス（{@link jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane} と同一規則）。 */
public final class EquipmentGanttTimelineColumns {

    private static final Pattern TIME_HEADER =
            Pattern.compile("^\\s*(\\d{1,2}):(\\d{2})\\s*$");

    private EquipmentGanttTimelineColumns() {}

    public static LocalTime parseTimeHeader(String col) {
        if (col == null) {
            return null;
        }
        var m = TIME_HEADER.matcher(col.strip());
        if (!m.matches()) {
            return null;
        }
        try {
            return LocalTime.of(Integer.parseInt(m.group(1)), Integer.parseInt(m.group(2)));
        } catch (RuntimeException e) {
            return null;
        }
    }

    /** 列見出しが HH:MM の列インデックスを左から列挙する。 */
    public static List<Integer> timeSlotColumnIndices(List<String> columns) {
        List<Integer> out = new ArrayList<>();
        if (columns == null) {
            return out;
        }
        for (int c = 0; c < columns.size(); c++) {
            if (parseTimeHeader(columns.get(c)) != null) {
                out.add(c);
            }
        }
        return out;
    }

    /** 1 行分のタイムラインスロット文言（時刻列のみ、UI の cellsInSlots と同じ並び）。 */
    public static List<String> timelineSlotTexts(Map<String, String> row, List<String> columns) {
        List<Integer> slotCols = timeSlotColumnIndices(columns);
        List<String> out = new ArrayList<>(slotCols.size());
        for (int c : slotCols) {
            String colName = columns.get(c);
            out.add(row != null && colName != null ? row.getOrDefault(colName, "") : "");
        }
        return out;
    }
}
