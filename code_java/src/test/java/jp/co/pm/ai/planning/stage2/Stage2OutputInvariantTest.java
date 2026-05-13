package jp.co.pm.ai.planning.stage2;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.planning.stage2.output.Stage2PythonishPlanWorkbookLayout;

/**
 * PassThrough 経路でも将来の Python 同一検証に向けた不変条件（スキーマ・シート数・空タイムライン）を固定する。
 */
class Stage2OutputInvariantTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Test
    void javaEngine_outputsPlanJsonEquipmentContractAndSheets(@TempDir Path root) throws Exception {
        Path out = root.resolve("out");
        Files.createDirectories(out);
        Path master = root.resolve("master.xlsx");
        writeMinimalMaster(master);
        Path plan = root.resolve("plan.csv");
        Files.writeString(plan, "依頼NO,工程名\nT1,加工A\n", StandardCharsets.UTF_8);
        Path pyDir = root.resolve("code").resolve("python");
        Files.createDirectories(pyDir);
        Files.writeString(pyDir.resolve("task_extract_stage1.py"), "# stub\n", StandardCharsets.UTF_8);

        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_OUTPUT_DIR, out.toString());
        ui.put(AppPaths.KEY_PM_AI_REPO_ROOT, root.toString());
        ui.put(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR, pyDir.toString());
        ui.put(AppPaths.KEY_PM_AI_MASTER_WORKBOOK, master.toString());
        ui.put(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH, plan.toString());
        ui.put(AppPaths.KEY_PM_AI_PLAN_WORKBOOK_JSON, "1");
        ui.put(AppPaths.KEY_PM_AI_MEMBER_SCHEDULE_JSON, "1");
        ui.put(AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL, "1");

        List<String> log = new ArrayList<>();
        Stage2RunContext ctx = new Stage2RunContext(ui, "", log::add);
        int code = Stage2JavaEngine.run(ctx);
        assertEquals(0, code, String.join("\n", log));

        Path contract =
                Files.list(out)
                        .filter(
                                p -> {
                                    String n = p.getFileName().toString();
                                    return n.startsWith("計画") && n.endsWith("設.json");
                                })
                        .findFirst()
                        .orElseThrow();
        JsonNode croot = MAPPER.readTree(Files.readString(contract, StandardCharsets.UTF_8));
        assertEquals("equipment_gantt", croot.get("kind").asText());
        assertTrue(croot.get("kwargs_packed").get("timeline_events").isArray());

        Path planXlsx =
                Files.list(out)
                        .filter(
                                p -> {
                                    String n = p.getFileName().toString();
                                    return n.startsWith("計画") && n.endsWith(".xlsx") && !n.contains("設");
                                })
                        .findFirst()
                        .orElseThrow();
        try (var wb = new XSSFWorkbook(Files.newInputStream(planXlsx))) {
            assertEquals(Stage2PythonishPlanWorkbookLayout.PLAN_SHEET_ORDER.size(), wb.getNumberOfSheets());
            for (int i = 0; i < Stage2PythonishPlanWorkbookLayout.PLAN_SHEET_ORDER.size(); i++) {
                assertEquals(
                        Stage2PythonishPlanWorkbookLayout.PLAN_SHEET_ORDER.get(i),
                        wb.getSheetAt(i).getSheetName());
            }
        }
    }

    private static void writeMinimalMaster(Path path) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var skills = wb.createSheet("skills");
            skills.createRow(0).createCell(0).setCellValue("メンバー");
            skills.createRow(0).createCell(1).setCellValue("工程テスト");
            skills.createRow(1).createCell(1).setCellValue("機械テスト");
            skills.createRow(2).createCell(0).setCellValue("テストオペレータ");
            var main = wb.createSheet("メイン");
            var r12 = main.createRow(11);
            r12.createCell(0).setCellValue("08:00");
            r12.createCell(1).setCellValue("17:00");
            try (OutputStream os = Files.newOutputStream(path)) {
                wb.write(os);
            }
        }
    }
}
