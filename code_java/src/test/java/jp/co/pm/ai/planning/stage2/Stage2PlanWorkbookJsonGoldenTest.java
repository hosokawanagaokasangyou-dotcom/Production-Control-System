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
import org.junit.jupiter.api.io.TempDir;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.planning.stage2.parity.Stage2PlanJsonGoldenNormalizer;

/**
 * Java 段階2（PassThrough 経路）が出す計画ブック JSON のうち、{@code 結果_タスク一覧} シートがゴールデンと一致するか検証する。
 *
 * <p>{@code 結果_タスク一覧} は Python {@code default_result_task_sheet_column_order(0)} と同一の 28 列見出し（履歴列なし）で
 * PassThrough が転写する。値の厳密一致（配台状況メモ・数値型など）を Python 正本と取る場合は同入力で Python を実行し
 * {@code golden/case_minimal/result_task_list_sheet_expected.json} を差し替える（README_GOLDEN 参照）。
 */
class Stage2PlanWorkbookJsonGoldenTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Test
    void case_minimal_resultTaskListSheetMatchesGolden(@TempDir Path root) throws Exception {
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

        Path planJson =
                Files.list(out)
                        .filter(
                                p -> {
                                    String n = p.getFileName().toString();
                                    return n.startsWith("計画")
                                            && n.endsWith(".json")
                                            && !n.endsWith("設.json");
                                })
                        .findFirst()
                        .orElseThrow();

        JsonNode rootNode = MAPPER.readTree(Files.readString(planJson, StandardCharsets.UTF_8));
        JsonNode normalized = Stage2PlanJsonGoldenNormalizer.stripTopLevelVolatile(rootNode);
        JsonNode actualTaskSheet =
                Stage2PlanJsonGoldenNormalizer.sheetSubtree(normalized, "結果_タスク一覧");

        JsonNode expected;
        try (var in =
                Stage2PlanWorkbookJsonGoldenTest.class.getResourceAsStream(
                        "/stage2-parity/golden/case_minimal/result_task_list_sheet_expected.json")) {
            assertTrue(in != null);
            expected = MAPPER.readTree(in.readAllBytes());
        }

        assertFalse(actualTaskSheet.isMissingNode());
        assertEquals(expected, actualTaskSheet, "結果_タスク一覧 がゴールデンと一致しません（Python 正本へ寄せるときは golden を更新）");
    }
}
