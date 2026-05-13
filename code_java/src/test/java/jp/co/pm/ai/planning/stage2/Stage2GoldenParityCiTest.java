package jp.co.pm.ai.planning.stage2;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
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
import org.junit.jupiter.api.condition.EnabledIfEnvironmentVariable;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * CI や手動で {@code PM_AI_STAGE2_GOLDEN_CI=1} を付けたときのみ実行。Python 二重起動は含まない（環境依存が大きいため）。
 * 将来: ここに golden 全件の同一検証ヘッドレス化を追加する。
 */
@EnabledIfEnvironmentVariable(named = "PM_AI_STAGE2_GOLDEN_CI", matches = "1")
class Stage2GoldenParityCiTest {

    @Test
    void javaEngineMinimalRun_forCiGate(@TempDir Path root) throws Exception {
        Path out = root.resolve("out");
        Files.createDirectories(out);
        Path master = root.resolve("master.xlsx");
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
            try (OutputStream os = Files.newOutputStream(master)) {
                wb.write(os);
            }
        }
        Path plan = root.resolve("plan.csv");
        Files.writeString(plan, "依頼NO,工程名\nT9,加工A\n", StandardCharsets.UTF_8);
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
        ui.put(AppPaths.KEY_PM_AI_MEMBER_SCHEDULE_JSON, "0");
        ui.put(AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL, "1");

        List<String> log = new ArrayList<>();
        Stage2RunContext ctx = new Stage2RunContext(ui, "", log::add);
        int code = Stage2JavaEngine.run(ctx);
        assertEquals(0, code);
        assertFalse(log.isEmpty());
        assertTrue(log.stream().anyMatch(s -> s.contains("Stage2JavaDispatchEngine")), () -> String.join("\n", log));
    }
}
