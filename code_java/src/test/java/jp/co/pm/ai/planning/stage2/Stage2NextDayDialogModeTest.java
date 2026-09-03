package jp.co.pm.ai.planning.stage2;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class Stage2NextDayDialogModeTest {

    @Test
    void effectiveWhenSkipToday_preservesRequestedMode() {
        for (Stage2NextDayDialogMode requested : Stage2NextDayDialogMode.values()) {
            assertEquals(
                    requested,
                    Stage2NextDayDialogMode.effectiveForTodayDispatch(requested, true));
        }
    }

    @Test
    void effectiveWhenTodayDispatch_suppressesAllNextDayDialogs() {
        for (Stage2NextDayDialogMode requested : Stage2NextDayDialogMode.values()) {
            assertEquals(
                    Stage2NextDayDialogMode.NONE,
                    Stage2NextDayDialogMode.effectiveForTodayDispatch(requested, false),
                    "requested=" + requested);
        }
    }
}
