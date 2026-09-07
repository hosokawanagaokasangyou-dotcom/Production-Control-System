package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.time.LocalDate;
import org.junit.jupiter.api.Test;

class ProcessingTrendPeriodPresetTest {

    @Test
    void past3Months_spansInclusiveThreeCalendarMonths() {
        LocalDate today = LocalDate.of(2026, 9, 7);
        LocalDate[] range =
                ProcessingTrendTabController.presetRange(
                        ProcessingTrendTabController.PeriodPreset.PAST_3_MONTHS, today);
        assertNotNull(range);
        assertEquals(LocalDate.of(2026, 7, 1), range[0]);
        assertEquals(LocalDate.of(2026, 9, 30), range[1]);
    }

    @Test
    void past6Months_spansInclusiveSixCalendarMonths() {
        LocalDate today = LocalDate.of(2026, 9, 7);
        LocalDate[] range =
                ProcessingTrendTabController.presetRange(
                        ProcessingTrendTabController.PeriodPreset.PAST_6_MONTHS, today);
        assertNotNull(range);
        assertEquals(LocalDate.of(2026, 4, 1), range[0]);
        assertEquals(LocalDate.of(2026, 9, 30), range[1]);
    }
}
