package jp.co.pm.ai.planning.stage2.source;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

import org.junit.jupiter.api.Test;

class Stage2SkipTodayDispatchPolicyTest {

    @Test
    void todayBeforeShiftStart_isTodayDispatch() {
        LocalDateTime extraction = LocalDateTime.of(2026, 7, 10, 7, 5);
        assertFalse(
                Stage2SkipTodayDispatchPolicy.shouldSkipTodayDispatch(
                        extraction, LocalDate.of(2026, 7, 10), LocalTime.of(8, 45)));
    }

    @Test
    void todayAfterShiftStart_skipsToday() {
        LocalDateTime extraction = LocalDateTime.of(2026, 7, 10, 9, 0);
        assertTrue(
                Stage2SkipTodayDispatchPolicy.shouldSkipTodayDispatch(
                        extraction, LocalDate.of(2026, 7, 10), LocalTime.of(8, 45)));
    }

    @Test
    void otherDay_skipsToday() {
        LocalDateTime extraction = LocalDateTime.of(2026, 7, 9, 7, 5);
        assertTrue(
                Stage2SkipTodayDispatchPolicy.shouldSkipTodayDispatch(
                        extraction, LocalDate.of(2026, 7, 10), LocalTime.of(8, 45)));
    }
}
