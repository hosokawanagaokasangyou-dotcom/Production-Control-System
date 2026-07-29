package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

class PlanTasksMissingSkillsColumnPromptTest {

    @TempDir Path tmp;

    @Test
    void detectsPlanPairsMissingFromSkillsSheet() throws Exception {
        Path code = tmp.resolve("code");
        Files.createDirectories(code);
        Path master = code.resolve("master.xlsm");
        Path plan = code.resolve("output").resolve("plan_input_tasks.xlsx");
        Files.createDirectories(plan.getParent());

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("skills");
            var r0 = sh.createRow(0);
            r0.createCell(1).setCellValue("スライス");
            var r1 = sh.createRow(1);
            r1.createCell(1).setCellValue("スライス機3");
            var r2 = sh.createRow(2);
            r2.createCell(0).setCellValue("山田");
            r2.createCell(1).setCellValue("OP1");
            try (OutputStream os = Files.newOutputStream(master)) {
                wb.write(os);
            }
        }

        PlanInputTabularIo.write(
                plan,
                AppPaths.STAGE1_PLAN_OUTPUT_SHEET,
                new PlanInputTabularIo.TabularSheet(
                        List.of("依頼NO", "工程名", "機械名", "配台不要"),
                        List.of(
                                List.of("W7-22", "検反", "フィルム挿入機(間紙)", ""),
                                List.of("Y5-186", "スライス", "スライス機3", ""))));

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        tmp.toString(),
                        AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
                        master.toString(),
                        AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                        plan.toString());

        var bundle = PlanTasksMissingSkillsColumnPrompt.collectMissingPairs(ui);
        assertFalse(bundle.empty());
        assertEquals(1, bundle.pairs().size());
        assertEquals("検反", bundle.pairs().getFirst().process());
        assertEquals("フィルム挿入機(間紙)", bundle.pairs().getFirst().machine());
        assertEquals("W7-22", bundle.pairs().getFirst().sampleTaskId());
    }

    @Test
    void skipsExcludedPlanRows() throws Exception {
        Path code = tmp.resolve("code");
        Files.createDirectories(code);
        Path master = code.resolve("master.xlsm");
        Path plan = code.resolve("output").resolve("plan_input_tasks.xlsx");
        Files.createDirectories(plan.getParent());

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("skills");
            var r0 = sh.createRow(0);
            r0.createCell(1).setCellValue("スライス");
            var r1 = sh.createRow(1);
            r1.createCell(1).setCellValue("スライス機3");
            try (OutputStream os = Files.newOutputStream(master)) {
                wb.write(os);
            }
        }

        PlanInputTabularIo.write(
                plan,
                AppPaths.STAGE1_PLAN_OUTPUT_SHEET,
                new PlanInputTabularIo.TabularSheet(
                        List.of("依頼NO", "工程名", "機械名", "配台不要"),
                        List.of(List.of("W7-22", "検反", "フィルム挿入機(間紙)", "yes"))));

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        tmp.toString(),
                        AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
                        master.toString(),
                        AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                        plan.toString());

        var bundle = PlanTasksMissingSkillsColumnPrompt.collectMissingPairs(ui);
        assertTrue(bundle.empty());
    }
}
