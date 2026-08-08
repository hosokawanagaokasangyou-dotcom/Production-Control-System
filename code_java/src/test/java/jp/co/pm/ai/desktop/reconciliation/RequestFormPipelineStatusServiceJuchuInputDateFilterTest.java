package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class RequestFormPipelineStatusServiceJuchuInputDateFilterTest {

    private static final int DAYS = RequestFormPipelineStatusService.DEFAULT_JUCHU_INPUT_DATE_HIDE_DAYS;

    @Test
    void shouldSkipJuchuRowDuringScan_whenHideDaysEnabled() {
        LocalDate old = LocalDate.now().minusDays(31);
        Map<String, String> juchu = Map.of("入力日", old.toString());
        assertTrue(
                RequestFormPipelineStatusService.shouldSkipJuchuRowDuringScan(juchu, DAYS));
        assertFalse(RequestFormPipelineStatusService.shouldSkipJuchuRowDuringScan(juchu, 0));
    }

    @Test
    void indexExcelRawByIraiKey_skipsTpiAndIndexesExcel() {
        List<Map<String, String>> raw =
                List.of(
                        Map.of(
                                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND,
                                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF,
                                "依頼Ｎｏ",
                                "GB60604"),
                        Map.of("依頼Ｎｏ", "Y8-99", "_sourceFileName", "book.xlsm"));
        Map<String, Map<String, String>> index =
                RequestFormPipelineStatusService.indexExcelRawByIraiKey(raw);
        assertEquals(1, index.size());
        assertEquals("Y8-99", index.get(JuchuTransferValueNormalizer.normalizeKey("Y8-99")).get("依頼Ｎｏ"));
    }

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

    @Test
    void isAdjustDeliveryOnOrAfterToday_whenTodayOrFuture() {
        assertFalse(RequestFormPipelineStatusService.isAdjustDeliveryOnOrAfterToday(null));
        assertFalse(
                RequestFormPipelineStatusService.isAdjustDeliveryOnOrAfterToday(
                        LocalDate.now().minusDays(1)));
        assertTrue(
                RequestFormPipelineStatusService.isAdjustDeliveryOnOrAfterToday(LocalDate.now()));
        assertTrue(
                RequestFormPipelineStatusService.isAdjustDeliveryOnOrAfterToday(
                        LocalDate.now().plusDays(1)));
    }

    @Test
    void resolveAdjustDeliveryLocalDate_prefersParsedValue() {
        LocalDate adjust = LocalDate.of(2026, 8, 17);
        var row =
                new RequestFormPipelineStatusService.PipelineStatusRow(
                        "X",
                        "",
                        false,
                        "",
                        false,
                        "",
                        0,
                        0,
                        "",
                        "",
                        false,
                        java.util.List.of(),
                        null,
                        java.util.List.of(),
                        null,
                        "",
                        "",
                        adjust,
                        "2026/7/1",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        null);
        assertEquals(
                adjust, RequestFormPipelineStatusService.resolveAdjustDeliveryLocalDate(row));
    }

    @Test
    void resolveAdjustDeliveryLocalDate_fallsBackToDisplayWhenParsedMissing() {
        var row =
                new RequestFormPipelineStatusService.PipelineStatusRow(
                        "X",
                        "",
                        false,
                        "",
                        false,
                        "",
                        0,
                        0,
                        "",
                        "",
                        false,
                        java.util.List.of(),
                        null,
                        java.util.List.of(),
                        null,
                        "",
                        "",
                        null,
                        "2026/8/17",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        null);
        assertEquals(
                LocalDate.of(2026, 8, 17),
                RequestFormPipelineStatusService.resolveAdjustDeliveryLocalDate(row));
    }
}
