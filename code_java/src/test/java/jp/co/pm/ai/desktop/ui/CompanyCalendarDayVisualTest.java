package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.time.LocalDate;

import org.junit.jupiter.api.Test;

class CompanyCalendarDayVisualTest {

    @Test
    void resolveKind_usesSpecialHolidayFromEntry() {
        LocalDate d = LocalDate.of(2026, 8, 14);
        CompanyCalendarDayVisual.DayInfo entry =
                new CompanyCalendarDayVisual.DayInfo(
                        CompanyCalendarDayVisual.KIND_SPECIAL, "夏季休暇", true, null);
        assertEquals(
                CompanyCalendarDayVisual.KIND_SPECIAL,
                CompanyCalendarDayVisual.resolveKind(d, entry));
        assertEquals(
                "14特",
                CompanyCalendarDayVisual.shortLabel(
                        14, CompanyCalendarDayVisual.KIND_SPECIAL, entry));
    }

    @Test
    void resolveKind_blankKindFallsBackToWeekdayDefault() {
        LocalDate weekday = LocalDate.of(2026, 8, 12);
        CompanyCalendarDayVisual.DayInfo blank =
                new CompanyCalendarDayVisual.DayInfo("", "", false, null);
        assertEquals(
                CompanyCalendarDayVisual.KIND_WORKING,
                CompanyCalendarDayVisual.resolveKind(weekday, blank));
    }
}
