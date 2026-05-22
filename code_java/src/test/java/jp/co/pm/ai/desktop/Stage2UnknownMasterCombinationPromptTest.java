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
import jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

class Stage2UnknownMasterCombinationPromptTest {

    @TempDir Path tmp;

    @Test
    void detectsPlanPairsMissingFromMasterCombinationSheet() throws Exception {
        Path code = tmp.resolve("code");
        Files.createDirectories(code);
        Path master = code.resolve("master.xlsm");
        Path plan = code.resolve("output").resolve("plan_input_tasks.xlsx");
        Files.createDirectories(plan.getParent());

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet(MasterTeamCombinationTableReader.SHEET_NAME);
            var h = sh.createRow(0);
            h.createCell(1).setCellValue("工程名");
            h.createCell(2).setCellValue("機械名");
            var r1 = sh.createRow(1);
            r1.createCell(1).setCellValue("スライス");
            r1.createCell(2).setCellValue("スライス機3");
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
                                List.of("W5-6", "巻返し", "フィルム挿入機(間紙)", ""),
                                List.of("Y5-186", "スライス", "スライス機3", ""))));

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        tmp.toString(),
                        AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
                        master.toString(),
                        AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                        plan.toString());

        var bundle = Stage2UnknownMasterCombinationPrompt.collectUnknownPairs(ui);
        assertFalse(bundle.empty());
        assertEquals(1, bundle.pairs().size());
        assertEquals("巻返し", bundle.pairs().getFirst().process());
        assertEquals("フィルム挿入機(間紙)", bundle.pairs().getFirst().machine());
    }

    @Test
    void applyUpdatesExcludeJsonAndPlanRows() throws Exception {
        Path code = tmp.resolve("code");
        Path jsonDir = code.resolve("json");
        Files.createDirectories(jsonDir);
        Path json = jsonDir.resolve("stage1_exclude_rules.json");
        Files.writeString(json, "{\"rules\":[]}\n", StandardCharsets.UTF_8);
        Path plan = code.resolve("output").resolve("plan_input_tasks.xlsx");
        Files.createDirectories(plan.getParent());
        PlanInputTabularIo.write(
                plan,
                AppPaths.STAGE1_PLAN_OUTPUT_SHEET,
                new PlanInputTabularIo.TabularSheet(
                        List.of("依頼NO", "工程名", "機械名", "配台不要"),
                        List.of(List.of("W5-6", "巻返し", "フィルム挿入機(間紙)", ""))));

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        tmp.toString(),
                        AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON,
                        json.toString(),
                        AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                        plan.toString());

        var applied =
                Stage2UnknownMasterCombinationPrompt.applyExcludeSelections(
                        ui,
                        List.of(
                                new Stage2UnknownMasterCombinationPrompt.UnknownPair(
                                        "巻返し", "フィルム挿入機(間紙)", "W5-6")));
        assertTrue(applied.excludeRulesUpdated() >= 1);
        assertEquals(1, applied.planRowsUpdated());
        String jsonText = Files.readString(json, StandardCharsets.UTF_8);
        assertTrue(jsonText.contains("フィルム挿入機(間紙)"));
        assertTrue(jsonText.contains("\"yes\""));
    }
}
