package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.RequestFormPipelineCheckTabController.MainRow;
import jp.co.pm.ai.desktop.RequestFormPipelineCheckTabController.PlanCopyHeader;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup.PlanEntry;

class RequestFormPipelineCheckTabControllerTest {

    private static PlanCopyHeader sampleHeader() {
        return new PlanCopyHeader("GB6064", "187735W", "2026/7/5");
    }

    @Test
    void formatMainTableTsv_includesVisibleColumns() {
        MainRow row = sampleMainRow();

        String tsv = RequestFormPipelineCheckTabController.formatMainTableTsv(List.of(row));
        assertTrue(tsv.startsWith("依頼No\t原本\t回答日"));
        assertTrue(tsv.contains("W7-14"));
        assertTrue(tsv.contains("6/9"));
        assertTrue(tsv.contains("納期備考A"));
        assertTrue(tsv.contains("79% (15/19)"));
        assertTrue(tsv.contains("7/3 100m"));
    }

    @Test
    void formatMainTableTsv_includesRawInputDateMatchColumn() {
        MainRow row = sampleMainRow();
        row.setRawInputDateMatchStatus("不一致");

        String tsv = RequestFormPipelineCheckTabController.formatMainTableTsv(List.of(row));
        assertTrue(tsv.contains("投入日一致"));
        assertTrue(tsv.contains("不一致"));
    }

    @Test
    void formatMainTableCsv_quotesCommaInCell() {
        MainRow row = sampleMainRow();
        row.setIndexDeliveryRemarks("A,B");

        String csv = RequestFormPipelineCheckTabController.formatMainTableCsv(List.of(row));
        assertTrue(csv.startsWith("依頼No,原本,回答日"));
        assertTrue(csv.contains("\"A,B\""));
    }

    @Test
    void formatMainTableHtml_includesTableCells() {
        MainRow row = new MainRow();
        row.setIraiNo("JR260701");
        row.setIndexResponseDate("7/1");
        row.setRateDisplay("100% (20/20)");
        row.setPlanDayValues(List.of("7/6 1940m"));

        String html = RequestFormPipelineCheckTabController.formatMainTableHtml(List.of(row));
        assertTrue(html.contains("JR260701"));
        assertTrue(html.contains("回答日"));
        assertTrue(html.contains("100% (20/20)"));
        assertTrue(html.contains("7/6 1940m"));
    }

    private static MainRow sampleMainRow() {
        MainRow row = new MainRow();
        row.setIraiNo("W7-14");
        row.setOriginalFile("W-7月加工依頼書.xlsm");
        row.setIndexResponseDate("6/9");
        row.setIndexInputDate("6/10");
        row.setIndexDeliveryDate("6/22");
        row.setIndexDeliveryRemarks("納期備考A");
        row.setIndexContractNo("186932F");
        row.setIndexContractRemarks("契約備考B");
        row.setJuchuInputDate("2024/6/29");
        row.setJuchuInputOperator("砂田");
        row.setJuchuAdjustDeliveryDate("2024/7/8");
        row.setRateDisplay("79% (15/19)");
        row.setMismatchCount("4");
        row.setOriginalContractNo("186932F");
        row.setContractNoStatus("186932F");
        row.setAladdinStatus("なし");
        row.setPlanDayValues(List.of("7/3 100m", "", "", "", "", "", ""));
        return row;
    }

    @Test
    void formatCrossSourceTableTsv_includesHeaderAndRows() {
        var rows =
                List.of(
                        new RequestFormPipelineCheckTabController.CrossSourceRow(
                                "アラジン", "2026/7/6", "不一致"),
                        new RequestFormPipelineCheckTabController.CrossSourceRow(
                                "受注ファイル", "2026/7/6", "不一致"));
        var header =
                new RequestFormPipelineCheckTabController.CrossSourceCopyHeader("W7-14", "不一致");
        String tsv =
                RequestFormPipelineCheckTabController.formatCrossSourceTableTsv(header, rows);
        assertEquals(
                "依頼No\tW7-14\n"
                        + "投入日一致\t不一致\n"
                        + "ソース\t原反投入日\t照合\n"
                        + "アラジン\t2026/7/6\t不一致\n"
                        + "受注ファイル\t2026/7/6\t不一致",
                tsv);
    }

    @Test
    void formatPlanTableTsv_includesRequestHeaderAndRows() {
        var row =
                new RequestFormPipelineCheckTabController.PlanRow(
                        new PlanEntry("スリット機1　湖南", "スリット", "2026/07/01", 3600), "未了");
        String tsv =
                RequestFormPipelineCheckTabController.formatPlanTableTsv(
                        sampleHeader(), List.of(row));
        assertEquals(
                "依頼No\tGB6064\n"
                        + "契約NO\t187735W\n"
                        + "原反投入日\t2026/7/5\n"
                        + "機械名\t工程名\t日付\t計画m\t日報完了\n"
                        + "スリット機1　湖南\tスリット\t2026/07/01\t3600\t未了",
                tsv);
    }

    @Test
    void formatPlanTableHtml_includesRequestHeaderAndTableCells() {
        var row =
                new RequestFormPipelineCheckTabController.PlanRow(
                        new PlanEntry("EC機　湖南", "EC", "2026/07/03", 3600), "完了");
        String html =
                RequestFormPipelineCheckTabController.formatPlanTableHtml(
                        sampleHeader(), List.of(row));
        assertTrue(html.contains("依頼No"));
        assertTrue(html.contains("GB6064"));
        assertTrue(html.contains("EC機　湖南"));
        assertTrue(html.contains("2026/07/03"));
    }
}
