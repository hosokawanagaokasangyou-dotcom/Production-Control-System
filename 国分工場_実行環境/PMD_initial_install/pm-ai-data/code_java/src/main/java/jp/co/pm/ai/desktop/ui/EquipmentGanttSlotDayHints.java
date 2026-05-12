package jp.co.pm.ai.desktop.ui;

import java.time.LocalTime;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

/**
 * Heuristics to estimate how many time-slot columns form one calendar day (for day-based tiling / scheduling).
 */
public final class EquipmentGanttSlotDayHints {

    private static final Pattern TIME_HEADER =
            Pattern.compile("^\\s*(\\d{1,2}):(\\d{2})\\s*$");

    private EquipmentGanttSlotDayHints() {}

    /**
     * From ordered HH:MM column headers, find the first index where time jumps backward or by a large gap
     * (treated as a new day). Returns that index as slots-per-day, or {@code <= 0} if unknown.
     */
    public static int estimateSlotsPerCalendarDay(List<String> columnHeaders, int slotMinutes) {
        if (columnHeaders == null || columnHeaders.isEmpty() || slotMinutes < 1) {
            return -1;
        }
        List<LocalTime> times = new ArrayList<>();
        for (String h : columnHeaders) {
            LocalTime t = parseTimeHeaderLoose(h);
            if (t != null) {
                times.add(t);
            }
        }
        if (times.size() < 4) {
            return -1;
        }
        int maxGapSlots = Math.max(3, (14 * 60) / slotMinutes);
        for (int i = 1; i < times.size(); i++) {
            LocalTime prev = times.get(i - 1);
            LocalTime cur = times.get(i);
            int pm = prev.getHour() * 60 + prev.getMinute();
            int cm = cur.getHour() * 60 + cur.getMinute();
            if (cm < pm || cm - pm > maxGapSlots * slotMinutes) {
                return i;
            }
        }
        return -1;
    }

    private static LocalTime parseTimeHeaderLoose(String col) {
        if (col == null) {
            return null;
        }
        var m = TIME_HEADER.matcher(col.strip());
        if (!m.matches()) {
            return null;
        }
        try {
            return LocalTime.of(Integer.parseInt(m.group(1)), Integer.parseInt(m.group(2)));
        } catch (Exception e) {
            return null;
        }
    }
}
