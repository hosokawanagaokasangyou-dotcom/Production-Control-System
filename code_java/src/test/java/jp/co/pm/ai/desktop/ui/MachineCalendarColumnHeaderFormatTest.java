package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class MachineCalendarColumnHeaderFormatTest {

    @Test
    void stripsCommonSuffixAndRedundantProcess() {
        List<EditableMachineCalendarGridPane.ColumnDef> cols =
                List.of(
                        new EditableMachineCalendarGridPane.ColumnDef(
                                "EC+EC機　湖南", "EC", "EC機　湖南"),
                        new EditableMachineCalendarGridPane.ColumnDef(
                                "SEC+SEC機　湖南", "SEC", "SEC機　湖南"),
                        new EditableMachineCalendarGridPane.ColumnDef(
                                "スリット+スリット機1　湖南", "スリット", "スリット機1　湖南"));
        var displays = MachineCalendarColumnHeaderFormat.formatAll(cols);
        assertEquals(3, displays.size());
        assertEquals("EC機", displays.get(0).text());
        assertEquals("SEC機", displays.get(1).text());
        assertEquals("スリット機1", displays.get(2).text());
    }

    @Test
    void disambiguatesSameMachineUnderDifferentProcesses() {
        List<EditableMachineCalendarGridPane.ColumnDef> cols =
                List.of(
                        new EditableMachineCalendarGridPane.ColumnDef(
                                "接続+熱融着機　湖南", "接続", "熱融着機　湖南"),
                        new EditableMachineCalendarGridPane.ColumnDef(
                                "熱融着+熱融着機　湖南", "熱融着", "熱融着機　湖南"),
                        new EditableMachineCalendarGridPane.ColumnDef(
                                "検査+熱融着機　湖南", "検査", "熱融着機　湖南"));
        var displays = MachineCalendarColumnHeaderFormat.formatAll(cols);
        assertEquals(3, displays.size());
        assertEquals("接続·熱融着機", displays.get(0).text());
        assertEquals("熱融着·熱融着機", displays.get(1).text());
        assertEquals("検査·熱融着機", displays.get(2).text());
    }

    @Test
    void tooltipKeepsFullNames() {
        List<EditableMachineCalendarGridPane.ColumnDef> cols =
                List.of(
                        new EditableMachineCalendarGridPane.ColumnDef(
                                "EC+EC機　湖南", "EC", "EC機　湖南"));
        var displays = MachineCalendarColumnHeaderFormat.formatAll(cols);
        assertTrue(displays.get(0).tooltip().contains("EC / EC機"));
        assertTrue(displays.get(0).tooltip().contains("EC+EC機"));
    }
}
