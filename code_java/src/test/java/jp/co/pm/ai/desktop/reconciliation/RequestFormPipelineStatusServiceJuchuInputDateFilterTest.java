package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.Map;

import org.junit.jupiter.api.Test;

class RequestFormPipelineStatusServiceJuchuInputDateFilterTest {

    private static final int DAYS = RequestFormPipelineStatusService.DEFAULT_JUCHU_INPUT_DATE_HIDE_DAYS;

    @Test
    void shouldHide_whenInputDateIs31DaysAgo() {
        LocalDate old = LocalDate.now().minusDays(31);
        Map<String, String> juchu = Map.of("入力日", old.toString());
        assertTrue(RequestFormPipelineStatusService.shouldHideByJuchuInputDate(juchu, DAYS));
    }

    @Test
    void shouldHide_whenInputDateIsExactly30DaysAgo() {
        LocalDate boundary = LocalDate.now().minusDays(30);
        Map<String, String> juchu = Map.of("入力日", boundary + " 00:00:00");
        assertTrue(RequestFormPipelineStatusService.shouldHideByJuchuInputDate(juchu, DAYS));
    }

    @Test
    void shouldNotHide_whenInputDateIs29DaysAgo() {
        LocalDate recent = LocalDate.now().minusDays(29);
        Map<String, String> juchu =
                Map.of(
                        "入力日",
                        recent.format(java.time.format.DateTimeFormatter.ofPattern("yyyy/MM/dd")));
        assertFalse(RequestFormPipelineStatusService.shouldHideByJuchuInputDate(juchu, DAYS));
    }

    @Test
    void shouldNotHide_whenJuchuMissingOrInputDateBlank() {
        assertFalse(
                RequestFormPipelineStatusService.shouldHideByJuchuInputDate(
                        (Map<String, String>) null, DAYS));
        assertFalse(RequestFormPipelineStatusService.shouldHideByJuchuInputDate(Map.of(), DAYS));
        assertFalse(
                RequestFormPipelineStatusService.shouldHideByJuchuInputDate(
                        Map.of("入力日", ""), DAYS));
    }

    @Test
    void shouldNotHide_whenExcludeDaysIsZero() {
        LocalDate old = LocalDate.now().minusDays(100);
        assertFalse(
                RequestFormPipelineStatusService.shouldHideByJuchuInputDate(old, 0));
    }

    @Test
    void formatJuchuInputOperatorDisplay_readsNyuryokuTanto() {
        Map<String, String> juchu = Map.of("入力担当", "細川");
        assertEquals(
                "細川",
                RequestFormPipelineStatusService.formatJuchuInputOperatorDisplay(juchu));
    }

    @Test
    void formatJuchuInputOperatorDisplay_acceptsNyuryokushaAliasKey() {
        Map<String, String> juchu = Map.of("入力者", "田中");
        assertEquals(
                "田中",
                RequestFormPipelineStatusService.formatJuchuInputOperatorDisplay(juchu));
    }

    @Test
    void shouldHideAdjustDelivery_whenBeforeTodayOrMissing() {
        assertTrue(
                RequestFormPipelineStatusService.shouldHideByAdjustDeliveryBeforeToday(null));
        assertTrue(
                RequestFormPipelineStatusService.shouldHideByAdjustDeliveryBeforeToday(
                        LocalDate.now().minusDays(1)));
        assertFalse(
                RequestFormPipelineStatusService.shouldHideByAdjustDeliveryBeforeToday(
                        LocalDate.now()));
        assertFalse(
                RequestFormPipelineStatusService.shouldHideByAdjustDeliveryBeforeToday(
                        LocalDate.now().plusDays(1)));
    }
}
