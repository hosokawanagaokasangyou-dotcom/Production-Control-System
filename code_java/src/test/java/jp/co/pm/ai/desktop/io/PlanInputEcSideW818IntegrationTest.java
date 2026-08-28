package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.reconciliation.EcSideClassification;

/** リポジトリ output/plan_input_tasks.xlsx の W8-18 EC面区分が Java POI で読めることを確認。 */
class PlanInputEcSideW818IntegrationTest {

    @Test
    void readW818EcSideFromFixture(@TempDir Path tempDir) throws Exception {
        Path plan = tempDir.resolve(AppPaths.STAGE1_PLAN_TASKS_FILENAME);
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("依頼NO");
            header.createCell(1).setCellValue(EcSideClassification.COLUMN_TITLE);
            Row data = sheet.createRow(1);
            data.createCell(0).setCellValue("W8-18");
            data.createCell(1).setCellValue(EcSideClassification.DOUBLE_SIDED);
            try (var out = Files.newOutputStream(plan)) {
                workbook.write(out);
            }
        }
        PlanInputTabularIo.TabularRead tr =
                PlanInputTabularIo.readWithResolvedSheet(plan, AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
        List<String> headers = tr.tabular().headers();
        int iTask = headers.indexOf("依頼NO");
        int iEc = headers.indexOf(EcSideClassification.COLUMN_TITLE);
        if (iTask < 0 || iEc < 0) {
            throw new AssertionError("missing columns task=" + iTask + " ec=" + iEc);
        }
        String ecVal = "";
        for (List<String> row : tr.tabular().rows()) {
            String tid = iTask < row.size() && row.get(iTask) != null ? row.get(iTask).strip() : "";
            if ("W8-18".equalsIgnoreCase(tid)) {
                ecVal = iEc < row.size() && row.get(iEc) != null ? row.get(iEc).strip() : "";
                break;
            }
        }
        assertEquals(
                EcSideClassification.DOUBLE_SIDED,
                ecVal,
                "W8-18 EC面区分 in " + plan);
    }
}
