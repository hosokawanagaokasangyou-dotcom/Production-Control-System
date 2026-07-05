package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.RequestFormPipelineCheckTabController.PlanCopyHeader;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup.PlanEntry;

class RequestFormPipelineCheckTabControllerTest {

    private static PlanCopyHeader sampleHeader() {
        return new PlanCopyHeader("GB6064", "187735W", "2026/7/5");
    }

    @Test
    void formatPlanTableTsv_includesRequestHeaderAndRows() {
        var row =
                new RequestFormPipelineCheckTabController.PlanRow(
                        new PlanEntry("スリット機1　湖南", "スリット", "2026/07/01", 3600));
        String tsv =
                RequestFormPipelineCheckTabController.formatPlanTableTsv(
                        sampleHeader(), List.of(row));
        assertEquals(
                "依頼No\tGB6064\n"
                        + "契約NO\t187735W\n"
                        + "原反投入日\t2026/7/5\n"
                        + "機械名\t工程名\t日付\t計画m\n"
                        + "スリット機1　湖南\tスリット\t2026/07/01\t3600",
                tsv);
    }

    @Test
    void formatPlanTableHtml_includesRequestHeaderAndTableCells() {
        var row =
                new RequestFormPipelineCheckTabController.PlanRow(
                        new PlanEntry("EC機　湖南", "EC", "2026/07/03", 3600));
        String html =
                RequestFormPipelineCheckTabController.formatPlanTableHtml(
                        sampleHeader(), List.of(row));
        assertTrue(html.contains("依頼No"));
        assertTrue(html.contains("GB6064"));
        assertTrue(html.contains("EC機　湖南"));
        assertTrue(html.contains("2026/07/03"));
    }
}
