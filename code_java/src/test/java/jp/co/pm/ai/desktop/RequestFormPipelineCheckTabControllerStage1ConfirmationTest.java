package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.RequestFormPipelineCheckTabController.MainRow;
import jp.co.pm.ai.desktop.reconciliation.JuchuTransferCoverageCheck.CoverageResult;
import jp.co.pm.ai.desktop.reconciliation.RawInputDateCrossSourceCheck;
import jp.co.pm.ai.desktop.reconciliation.RawInputDateCrossSourceCheck.CrossSourceResult;
import jp.co.pm.ai.desktop.reconciliation.RawInputDateCrossSourceCheck.SourceValues;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineStatusService.PipelineStatusRow;

class RequestFormPipelineCheckTabControllerStage1ConfirmationTest {

    @Test
    void requiresStage1Confirmation_trueWhenAllThreeConditionsMet() {
        MainRow row =
                sampleRow(
                        LocalDate.now().plusDays(3),
                        "2026/8/17",
                        "―",
                        RawInputDateCrossSourceCheck.STATUS_MISMATCH,
                        true,
                        true,
                        "",
                        false,
                        "なし");
        assertTrue(RequestFormPipelineCheckTabController.requiresStage1Confirmation(row));
    }

    @Test
    void requiresStage1Confirmation_falseWhenAdjustDeliveryBeforeToday() {
        MainRow row =
                sampleRow(
                        LocalDate.now().minusDays(1),
                        "2026/7/28",
                        "―",
                        RawInputDateCrossSourceCheck.STATUS_MISMATCH,
                        true,
                        true,
                        "",
                        true,
                        "あり");
        assertFalse(RequestFormPipelineCheckTabController.requiresStage1Confirmation(row));
    }

    @Test
    void requiresStage1Confirmation_falseWhenAdjustDeliveryMissingForTriplet() {
        MainRow row =
                sampleRow(
                        null,
                        "",
                        "―",
                        RawInputDateCrossSourceCheck.STATUS_MISMATCH,
                        true,
                        true,
                        "",
                        true,
                        "あり");
        assertFalse(RequestFormPipelineCheckTabController.requiresStage1Confirmation(row));
    }

    @Test
    void requiresStage1Confirmation_falseWhenDailyReportNotAbsent() {
        MainRow row =
                sampleRow(
                        LocalDate.now().plusDays(1),
                        "2026/8/1",
                        "未了",
                        RawInputDateCrossSourceCheck.STATUS_MISMATCH,
                        true,
                        true,
                        "",
                        true,
                        "あり");
        assertFalse(RequestFormPipelineCheckTabController.requiresStage1Confirmation(row));
    }

    @Test
    void requiresStage1Confirmation_falseWhenRawInputDateMatches() {
        MainRow row =
                sampleRow(
                        LocalDate.now().plusDays(1),
                        "2026/8/1",
                        "―",
                        RawInputDateCrossSourceCheck.STATUS_MATCH,
                        true,
                        true,
                        "",
                        true,
                        "あり");
        assertFalse(RequestFormPipelineCheckTabController.requiresStage1Confirmation(row));
    }

    @Test
    void requiresStage1Confirmation_usesDisplayWhenParsedDateMissing() {
        MainRow row =
                sampleRow(
                        null,
                        "2026/8/17",
                        "―",
                        RawInputDateCrossSourceCheck.STATUS_MISMATCH,
                        true,
                        true,
                        "",
                        false,
                        "なし");
        assertTrue(RequestFormPipelineCheckTabController.requiresStage1Confirmation(row));
    }

    @Test
    void requiresStage1Confirmation_falseForInHouseSelfProcessingIraiNo() {
        LocalDate future = LocalDate.now().plusDays(3);
        MainRow row =
                sampleRow(
                        future,
                        "2026/8/17",
                        "―",
                        RawInputDateCrossSourceCheck.STATUS_MISMATCH,
                        true,
                        true,
                        "",
                        false,
                        "なし");
        row.setIraiNo("2125-02-16");
        assertFalse(RequestFormPipelineCheckTabController.requiresStage1Confirmation(row));
    }

    @Test
    void requiresStage1Confirmation_trueWhenAladdinMissingAndJuchuAdjustDeliveryOnOrAfterToday() {
        LocalDate future = LocalDate.now().plusDays(5);
        MainRow row =
                sampleRow(
                        future,
                        future.getYear() + "/" + future.getMonthValue() + "/" + future.getDayOfMonth(),
                        "未了",
                        RawInputDateCrossSourceCheck.STATUS_MATCH,
                        true,
                        true,
                        "",
                        false,
                        "なし");
        assertTrue(RequestFormPipelineCheckTabController.requiresStage1Confirmation(row));
    }

    @Test
    void requiresStage1Confirmation_trueWhenAladdinMissingAndIndexDeliveryOnOrAfterToday() {
        LocalDate future = LocalDate.now().plusDays(3);
        MainRow row =
                sampleRow(
                        null,
                        "",
                        "未了",
                        RawInputDateCrossSourceCheck.STATUS_MATCH,
                        true,
                        false,
                        future.getYear() + "/" + future.getMonthValue() + "/" + future.getDayOfMonth(),
                        false,
                        "なし");
        assertTrue(RequestFormPipelineCheckTabController.requiresStage1Confirmation(row));
    }

    @Test
    void requiresStage1Confirmation_falseWhenDailyReportComplete() {
        LocalDate future = LocalDate.now().plusDays(5);
        MainRow row =
                sampleRow(
                        future,
                        future.getYear() + "/" + future.getMonthValue() + "/" + future.getDayOfMonth(),
                        "完了",
                        RawInputDateCrossSourceCheck.STATUS_MISMATCH,
                        true,
                        true,
                        "",
                        false,
                        "なし");
        assertFalse(RequestFormPipelineCheckTabController.requiresStage1Confirmation(row));
    }

    @Test
    void requiresStage1Confirmation_falseWhenAladdinMissingButDeliveryBeforeToday() {
        MainRow row =
                sampleRow(
                        LocalDate.now().minusDays(2),
                        "2026/7/1",
                        "―",
                        RawInputDateCrossSourceCheck.STATUS_MATCH,
                        true,
                        true,
                        "2026/6/1",
                        false,
                        "なし");
        assertFalse(RequestFormPipelineCheckTabController.requiresStage1Confirmation(row));
    }

    @Test
    void requiresStage1Confirmation_falseWhenAladdinJsonUnread() {
        MainRow row =
                sampleRow(
                        LocalDate.now().plusDays(1),
                        "2026/8/1",
                        "―",
                        RawInputDateCrossSourceCheck.STATUS_MATCH,
                        true,
                        true,
                        "",
                        false,
                        "未確認");
        assertFalse(RequestFormPipelineCheckTabController.requiresStage1Confirmation(row));
    }

    @Test
    void confirmAllRequiringConfirmation_checksOnlyUnconfirmedRequiringRows() {
        LocalDate future = LocalDate.now().plusDays(5);
        MainRow requiring =
                sampleRow(
                        future,
                        future.getYear() + "/" + future.getMonthValue() + "/" + future.getDayOfMonth(),
                        "―",
                        RawInputDateCrossSourceCheck.STATUS_MISMATCH,
                        true,
                        true,
                        "",
                        false,
                        "なし");
        MainRow exempt =
                sampleRow(
                        future,
                        future.getYear() + "/" + future.getMonthValue() + "/" + future.getDayOfMonth(),
                        "完了",
                        RawInputDateCrossSourceCheck.STATUS_MISMATCH,
                        true,
                        true,
                        "",
                        false,
                        "なし");

        int confirmed =
                RequestFormPipelineCheckTabController.confirmAllRequiringConfirmation(
                        List.of(requiring, exempt));

        assertEquals(1, confirmed);
        assertTrue(requiring.isIssueConfirmed());
        assertFalse(exempt.isIssueConfirmed());
    }

    @Test
    void confirmAllRequiringConfirmation_skipsAlreadyConfirmedRows() {
        LocalDate future = LocalDate.now().plusDays(5);
        MainRow row =
                sampleRow(
                        future,
                        future.getYear() + "/" + future.getMonthValue() + "/" + future.getDayOfMonth(),
                        "―",
                        RawInputDateCrossSourceCheck.STATUS_MISMATCH,
                        true,
                        true,
                        "",
                        false,
                        "なし");
        row.issueConfirmedProperty().set(true);

        int confirmed =
                RequestFormPipelineCheckTabController.confirmAllRequiringConfirmation(List.of(row));

        assertEquals(0, confirmed);
        assertTrue(row.isIssueConfirmed());
    }

    private static MainRow sampleRow(
            LocalDate parsedAdjustDelivery,
            String displayAdjustDelivery,
            String dailyReportStatus,
            String rawInputMatchStatus,
            boolean originalPresent,
            boolean juchuRegistered,
            String indexDeliveryDate,
            boolean aladdinPresent,
            String aladdinStatus) {
        CoverageResult coverage =
                new CoverageResult(juchuRegistered, 4, 4, 100.0, List.of());
        CrossSourceResult cross =
                new CrossSourceResult(
                        rawInputMatchStatus, new SourceValues("", "", "", ""), "");
        PipelineStatusRow source =
                new PipelineStatusRow(
                        "TEST",
                        originalPresent ? "book.xlsm" : "",
                        originalPresent,
                        "user",
                        juchuRegistered,
                        coverage.rateDisplay(),
                        coverage.ratePercent(),
                        coverage.mismatchCount(),
                        "",
                        "",
                        aladdinPresent,
                        List.of(),
                        coverage,
                        List.of(),
                        null,
                        "",
                        "",
                        parsedAdjustDelivery,
                        displayAdjustDelivery,
                        "",
                        "",
                        "",
                        indexDeliveryDate,
                        "",
                        "",
                        "",
                        cross);
        MainRow row = new MainRow();
        row.setSource(source);
        row.setJuchuAdjustDeliveryDate(displayAdjustDelivery);
        row.setIndexDeliveryDate(indexDeliveryDate);
        row.setDailyReportOrderStatus(dailyReportStatus);
        row.setRawInputDateMatchStatus(rawInputMatchStatus);
        row.setAladdinStatus(aladdinStatus);
        row.setHasIssues(true);
        return row;
    }
}
