package jp.co.pm.ai.desktop.io.gantt;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.io.JsonTableIo;

class EquipmentScheduleGridPivotTest {

    @Test
    void pivotsTimeBandRowsIntoMachineRowsAndSkipsProgressColumns() {
        JsonTableIo.SheetTable schedule =
                new JsonTableIo.SheetTable(
                        List.of("日時帯", "EC機　湖南（EC）", "EC機　湖南（EC）進度", "改装"),
                        List.of(
                                row("■ 2026/09/28 (Mon) ■", "", "", ""),
                                row("08:40-08:50", "日次始業準備", "", ""),
                                row(
                                        "08:50-09:00",
                                        "[C9-16] 主:宮島　　剛 補:森岡　真由美, 菅沼　めぐみ",
                                        "1/2R",
                                        ""),
                                row("■ 2026/09/29 (Tue) ■", "", "", ""),
                                row("08:40-08:50", "", "", "")));

        EquipmentScheduleGridPivot.Result pivoted = EquipmentScheduleGridPivot.pivot(schedule);

        assertEquals(
                List.of("日付", "機械名", "工程名", "タスク概覝", "8:40", "8:50"),
                pivoted.table().columns());
        assertEquals(2, pivoted.table().rows().size());
        Map<String, String> section = pivoted.table().rows().get(0);
        assertTrue(section.get("日付").contains("2026/09/28"));
        assertTrue(section.get("日付").contains("■"));
        Map<String, String> data = pivoted.table().rows().get(1);
        assertEquals("【2026/09/28】", data.get("日付"));
        assertEquals("EC機　湖南", data.get("機械名"));
        assertEquals("EC", data.get("工程名"));
        assertEquals("日次始業準備", data.get("8:40"));
        assertEquals("[C9-16] 主:宮島　　剛 補:森岡　真由美, 菅沼　めぐみ", data.get("8:50"));
        assertEquals("", pivoted.badgeSlotRows().get(1).get(0));
        String badge = pivoted.badgeSlotRows().get(1).get(1);
        assertTrue(badge.contains("宮島"), badge);
        assertTrue(badge.contains("森岡"), badge);
        assertTrue(badge.contains("菅沼"), badge);
    }

    @Test
    void returnsNullWhenTimelineCellsAreEmpty() {
        JsonTableIo.SheetTable schedule =
                new JsonTableIo.SheetTable(
                        List.of("日時帯", "EC機　湖南（EC）"),
                        List.of(row("■ 2026/09/28 (Mon) ■", ""), row("08:00-08:10", "")));
        assertNull(EquipmentScheduleGridPivot.pivot(schedule));
    }

    @Test
    void returnsNullForHorizontalGanttSheet() {
        JsonTableIo.SheetTable gantt =
                new JsonTableIo.SheetTable(
                        List.of("日付", "機械名", "8:00"),
                        List.of(Map.of("日付", "【2026/09/28】", "機械名", "EC機", "8:00", "A")));
        assertNull(EquipmentScheduleGridPivot.pivot(gantt));
    }

    private static Map<String, String> row(String band, String ec, String progress, String remodel) {
        Map<String, String> row = new LinkedHashMap<>();
        row.put("日時帯", band);
        row.put("EC機　湖南（EC）", ec);
        row.put("EC機　湖南（EC）進度", progress);
        row.put("改装", remodel);
        return row;
    }

    private static Map<String, String> row(String band, String ec) {
        Map<String, String> row = new LinkedHashMap<>();
        row.put("日時帯", band);
        row.put("EC機　湖南（EC）", ec);
        return row;
    }
}
