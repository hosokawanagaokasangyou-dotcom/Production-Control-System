package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.reconciliation.JuchuTransferCoverageCheck.CoverageResult;
import jp.co.pm.ai.desktop.reconciliation.RawInputDateCrossSourceCheck.CrossSourceResult;
import jp.co.pm.ai.desktop.reconciliation.RawInputDateCrossSourceCheck.SourceValues;
import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineStatusService.PipelineStatusRow;

class RequestFormPipelineIssueCheckTest {

    @Test
    void detect_flagsRawInputMismatchTransferAndAladdinMissing() {
        PipelineStatusRow row = sampleRow(true, 2, false, "186932F", true, "不一致");
        List<RequestFormPipelineIssueCheck.IssueKind> issues =
                RequestFormPipelineIssueCheck.detect(row, true);
        assertTrue(issues.contains(RequestFormPipelineIssueCheck.IssueKind.RAW_INPUT_DATE_MISMATCH));
        assertTrue(issues.contains(RequestFormPipelineIssueCheck.IssueKind.TRANSFER_MISMATCH));
        assertTrue(issues.contains(RequestFormPipelineIssueCheck.IssueKind.ALADDIN_MISSING));
        assertFalse(issues.contains(RequestFormPipelineIssueCheck.IssueKind.CONTRACT_NO_MISSING));
    }

    @Test
    void detect_flagsContractNoMissingAndNoOriginal() {
        PipelineStatusRow row = sampleRow(false, 0, false, "未入力", true, "―");
        List<RequestFormPipelineIssueCheck.IssueKind> issues =
                RequestFormPipelineIssueCheck.detect(row, true);
        assertTrue(issues.contains(RequestFormPipelineIssueCheck.IssueKind.CONTRACT_NO_MISSING));
        assertTrue(issues.contains(RequestFormPipelineIssueCheck.IssueKind.NO_ORIGINAL));
        assertTrue(issues.contains(RequestFormPipelineIssueCheck.IssueKind.ALADDIN_MISSING));
    }

    @Test
    void formatSummary_joinsLabels() {
        String summary =
                RequestFormPipelineIssueCheck.formatSummary(
                        List.of(
                                RequestFormPipelineIssueCheck.IssueKind.ALADDIN_MISSING,
                                RequestFormPipelineIssueCheck.IssueKind.TRANSFER_MISMATCH));
        assertEquals("アラジン計画なし・転記未一致", summary);
    }

    @Test
    void formatConfirmedDisplay_reflectsState() {
        assertEquals("―", RequestFormPipelineIssueCheck.formatConfirmedDisplay(false, false));
        assertEquals("未", RequestFormPipelineIssueCheck.formatConfirmedDisplay(true, false));
        assertEquals("済", RequestFormPipelineIssueCheck.formatConfirmedDisplay(true, true));
    }

    private static PipelineStatusRow sampleRow(
            boolean originalPresent,
            int mismatchCount,
            boolean aladdinPresent,
            String contractNoStatus,
            boolean juchuRegistered,
            String rawInputStatus) {
        CoverageResult coverage =
                new CoverageResult(juchuRegistered, 4, 4 - mismatchCount, 50.0, List.of());
        CrossSourceResult cross =
                new CrossSourceResult(
                        rawInputStatus,
                        new SourceValues("2026/7/6", "2026/7/7", "", ""),
                        "");
        return new PipelineStatusRow(
                "W7-14",
                originalPresent ? "book.xlsm" : "",
                originalPresent,
                "user",
                juchuRegistered,
                coverage.rateDisplay(),
                coverage.ratePercent(),
                coverage.mismatchCount(),
                "186932F",
                contractNoStatus,
                aladdinPresent,
                List.of(),
                coverage,
                List.of(),
                LocalDate.of(2026, 7, 1),
                "2026/7/1",
                "担当",
                LocalDate.of(2026, 7, 8),
                "2026/7/8",
                "2026/7/6",
                "",
                "",
                "",
                "",
                "",
                "",
                cross);
    }
}
