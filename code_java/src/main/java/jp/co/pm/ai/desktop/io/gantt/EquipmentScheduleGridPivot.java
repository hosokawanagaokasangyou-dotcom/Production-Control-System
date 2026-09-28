package jp.co.pm.ai.desktop.io.gantt;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.io.JsonTableIo;

/**
 * 「結果_設備毎の時間割」は行が時刻帯・列が設備である。
 * グラフィックガントは行が設備・列が HH:MM なので、契約 JSON が無いときにこの表へ組み替える。
 */
public final class EquipmentScheduleGridPivot {

    private static final Pattern SECTION_DATE =
            Pattern.compile("■\\s*(\\d{4})[/\\-.](\\d{1,2})[/\\-.](\\d{1,2})");

    private static final Pattern TIME_BAND =
            Pattern.compile("^(\\d{1,2}):(\\d{2})\\s*-\\s*(\\d{1,2}):(\\d{2})$");

    private static final Pattern TRAILING_PROCESS =
            Pattern.compile("^(.+)[（(]([^）)]+)[）)]\\s*$");

    private static final Pattern PRIMARY_OP = Pattern.compile("主[:：]\\s*(.+?)\\s*(?=補[:：]|$)");

    private static final Pattern ASSIST_OP = Pattern.compile("補[:：]\\s*(.+)$");

    private static final String COL_DATE = "日付";
    private static final String COL_MACH = "機械名";
    private static final String COL_PROC = "工程名";
    private static final String COL_TASK = "タスク概覝";

    private EquipmentScheduleGridPivot() {}

    public record Result(JsonTableIo.SheetTable table, List<List<String>> badgeSlotRows) {}

    /**
     * 先頭列が「日時帯」の時間割を、設備ガント用の横軸表にする。変換できる非空セルが無ければ null。
     */
    public static Result pivot(JsonTableIo.SheetTable schedule) {
        if (schedule == null || schedule.columns() == null || schedule.columns().isEmpty()) {
            return null;
        }
        if (!"日時帯".equals(schedule.columns().get(0))) {
            return null;
        }
        List<MachineCol> machines = machineColumns(schedule.columns());
        if (machines.isEmpty() || schedule.rows() == null) {
            return null;
        }

        Map<String, DayBlock> days = new LinkedHashMap<>();
        TreeSet<Integer> slotMinutes = new TreeSet<>();
        String currentDay = null;
        for (Map<String, String> row : schedule.rows()) {
            if (row == null) {
                continue;
            }
            String band = text(row.get("日時帯"));
            Matcher section = SECTION_DATE.matcher(band);
            if (section.find()) {
                currentDay = sectionDayKey(section);
                days.computeIfAbsent(currentDay, DayBlock::new);
                continue;
            }
            Matcher time = TIME_BAND.matcher(band);
            if (!time.matches() || currentDay == null) {
                continue;
            }
            int minute = Integer.parseInt(time.group(1)) * 60 + Integer.parseInt(time.group(2));
            DayBlock day = days.computeIfAbsent(currentDay, DayBlock::new);
            for (MachineCol machine : machines) {
                String cell = text(row.get(machine.header()));
                if (cell.isEmpty()) {
                    continue;
                }
                day.text.computeIfAbsent(machine.header(), k -> new LinkedHashMap<>()).put(minute, cell);
                day.badges
                        .computeIfAbsent(machine.header(), k -> new LinkedHashMap<>())
                        .put(minute, badgeFragment(cell));
                slotMinutes.add(minute);
            }
        }
        if (slotMinutes.isEmpty()) {
            return null;
        }

        List<Integer> slots = List.copyOf(slotMinutes);
        List<String> columns = new ArrayList<>();
        columns.add(COL_DATE);
        columns.add(COL_MACH);
        columns.add(COL_PROC);
        columns.add(COL_TASK);
        for (int minute : slots) {
            columns.add(formatSlot(minute));
        }

        List<Map<String, String>> outRows = new ArrayList<>();
        List<List<String>> badgeRows = new ArrayList<>();
        for (DayBlock day : days.values()) {
            boolean any = false;
            for (MachineCol machine : machines) {
                if (day.text.containsKey(machine.header())) {
                    any = true;
                    break;
                }
            }
            if (!any) {
                continue;
            }
            outRows.add(sectionRow(columns, day.banner()));
            badgeRows.add(emptyBadges(slots.size()));
            for (MachineCol machine : machines) {
                Map<Integer, String> cells = day.text.get(machine.header());
                if (cells == null || cells.isEmpty()) {
                    continue;
                }
                Map<Integer, String> badges = day.badges.getOrDefault(machine.header(), Map.of());
                outRows.add(dataRow(columns, slots, day.dateCell(), machine, cells));
                badgeRows.add(badgeRow(slots, badges));
            }
        }
        if (outRows.isEmpty()) {
            return null;
        }
        return new Result(new JsonTableIo.SheetTable(columns, outRows), badgeRows);
    }

    private static List<MachineCol> machineColumns(List<String> columns) {
        List<MachineCol> machines = new ArrayList<>();
        for (String header : columns) {
            if (header == null || header.isBlank() || "日時帯".equals(header) || header.endsWith("進度")) {
                continue;
            }
            String h = header.strip();
            Matcher m = TRAILING_PROCESS.matcher(h);
            if (m.matches() && !m.group(1).isBlank()) {
                machines.add(new MachineCol(header, m.group(1).strip(), m.group(2).strip()));
            } else {
                machines.add(new MachineCol(header, h, ""));
            }
        }
        return machines;
    }

    private static String sectionDayKey(Matcher section) {
        int y = Integer.parseInt(section.group(1));
        int mo = Integer.parseInt(section.group(2));
        int d = Integer.parseInt(section.group(3));
        return LocalDate.of(y, mo, d).toString();
    }

    private static Map<String, String> sectionRow(List<String> columns, String banner) {
        Map<String, String> row = new LinkedHashMap<>();
        for (String col : columns) {
            row.put(col, COL_DATE.equals(col) ? banner : "");
        }
        return row;
    }

    private static Map<String, String> dataRow(
            List<String> columns,
            List<Integer> slots,
            String dateCell,
            MachineCol machine,
            Map<Integer, String> cells) {
        Map<String, String> row = new LinkedHashMap<>();
        row.put(COL_DATE, dateCell);
        row.put(COL_MACH, machine.machine());
        row.put(COL_PROC, machine.process());
        row.put(COL_TASK, "—");
        for (int i = 0; i < slots.size(); i++) {
            String col = columns.get(4 + i);
            row.put(col, cells.getOrDefault(slots.get(i), ""));
        }
        return row;
    }

    private static List<String> badgeRow(List<Integer> slots, Map<Integer, String> badges) {
        List<String> row = new ArrayList<>();
        for (int minute : slots) {
            String b = badges.get(minute);
            row.add(b != null ? b : "");
        }
        return row;
    }

    private static List<String> emptyBadges(int n) {
        List<String> row = new ArrayList<>();
        for (int i = 0; i < n; i++) {
            row.add("");
        }
        return row;
    }

    private static String formatSlot(int minuteOfDay) {
        int hh = minuteOfDay / 60;
        int mm = minuteOfDay % 60;
        return hh + ":" + String.format("%02d", mm);
    }

    static String badgeFragment(String cell) {
        String primary = "";
        String assist = "";
        Matcher primaryMatch = PRIMARY_OP.matcher(cell);
        if (primaryMatch.find()) {
            primary = primaryMatch.group(1).strip();
        }
        Matcher assistMatch = ASSIST_OP.matcher(cell);
        if (assistMatch.find()) {
            assist = assistMatch.group(1).strip();
        }
        if (primary.isEmpty() && assist.isEmpty()) {
            return "";
        }
        return PersonNameBadgeText.joinBadgeCells(
                PersonNameBadgeText.badgeListFromOpSub(primary, assist, false));
    }

    private static String text(String raw) {
        return raw == null ? "" : raw.strip();
    }

    private record MachineCol(String header, String machine, String process) {}

    private static final class DayBlock {
        private final String dayKey;
        private final Map<String, Map<Integer, String>> text = new LinkedHashMap<>();
        private final Map<String, Map<Integer, String>> badges = new LinkedHashMap<>();

        private DayBlock(String dayKey) {
            this.dayKey = dayKey;
        }

        private String banner() {
            LocalDate day = LocalDate.parse(dayKey);
            return "■ "
                    + day.getYear()
                    + "/"
                    + String.format("%02d", day.getMonthValue())
                    + "/"
                    + String.format("%02d", day.getDayOfMonth())
                    + " ■";
        }

        private String dateCell() {
            LocalDate day = LocalDate.parse(dayKey);
            return "【"
                    + day.getYear()
                    + "/"
                    + String.format("%02d", day.getMonthValue())
                    + "/"
                    + String.format("%02d", day.getDayOfMonth())
                    + "】";
        }
    }
}
