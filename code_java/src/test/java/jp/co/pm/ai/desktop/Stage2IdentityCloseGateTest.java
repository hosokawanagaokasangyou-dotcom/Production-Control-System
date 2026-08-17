package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.DispatchAladdinEntrySheetBuilder;
import jp.co.pm.ai.desktop.io.DispatchAladdinEntryWorkbookExporter;

class Stage2IdentityCloseGateTest {

    private static final String PLAN_CSV =
            "列1,列2,列3,列4\n"
                    + "上段1,,,\n"
                    + "上段2,,,\n"
                    + "上段3,,,\n"
                    + "機械名,依頼NO,工程名,2026/07/07\n"
                    + ",,,\n"
                    + "M1,T001,工程A,10\n";

    @Test
    void decide_skipsWhenStage2NotCompleted(@TempDir Path tempDir) {
        Stage2IdentityCloseGate gate = new Stage2IdentityCloseGate();

        Stage2IdentityCloseGate.Decision d = gate.decide(testUi(tempDir));

        assertFalse(d.required());
    }

    @Test
    void decide_requiresChallengeWhenExcelMissing(@TempDir Path tempDir) {
        Stage2IdentityCloseGate gate = new Stage2IdentityCloseGate();
        gate.markStage2Completed();

        Stage2IdentityCloseGate.Decision d = gate.decide(testUi(tempDir));

        assertTrue(d.required());
        assertFalse(d.detail().isBlank());
    }

    @Test
    void decide_skipsWhenLocalExcelMatchesPlan(@TempDir Path tempDir) throws Exception {
        Map<String, String> ui = testUi(tempDir);
        Path sourceDir = AppPaths.resolveTaskInputSourceDir(ui);
        Files.createDirectories(sourceDir);
        Files.writeString(sourceDir.resolve("aladdin-plan.csv"), PLAN_CSV);
        Files.createDirectories(Path.of(ui.get(AppPaths.KEY_PM_AI_REPO_ROOT)).resolve("code"));
        LocalDate d = LocalDate.of(2026, 7, 7);
        DispatchAladdinEntryWorkbookExporter.write(ui, matchingWorkbook(d, 10));

        Stage2IdentityCloseGate gate = new Stage2IdentityCloseGate();
        gate.markStage2Completed();

        Stage2IdentityCloseGate.Decision decision = gate.decide(ui);

        assertFalse(decision.required(), decision.detail());
    }

    @Test
    void decide_requiresChallengeWhenExcelExportFailedEvenIfOldLocalMatches(
            @TempDir Path tempDir) throws Exception {
        Map<String, String> ui = testUi(tempDir);
        Path sourceDir = AppPaths.resolveTaskInputSourceDir(ui);
        Files.createDirectories(sourceDir);
        Files.writeString(sourceDir.resolve("aladdin-plan.csv"), PLAN_CSV);
        Files.createDirectories(Path.of(ui.get(AppPaths.KEY_PM_AI_REPO_ROOT)).resolve("code"));
        LocalDate d = LocalDate.of(2026, 7, 7);
        DispatchAladdinEntryWorkbookExporter.write(ui, matchingWorkbook(d, 10));

        Stage2IdentityCloseGate gate = new Stage2IdentityCloseGate();
        gate.markStage2Completed(false);

        Stage2IdentityCloseGate.Decision decision = gate.decide(ui);

        assertTrue(decision.required(), "古い一致xlsxがあっても出力失敗ならゲート必須");
        assertEquals("Excel出力失敗", decision.detail());
    }

    @Test
    void decide_requiresChallengeWhenQtyMismatches(@TempDir Path tempDir) throws Exception {
        Map<String, String> ui = testUi(tempDir);
        Path sourceDir = AppPaths.resolveTaskInputSourceDir(ui);
        Files.createDirectories(sourceDir);
        Files.writeString(sourceDir.resolve("aladdin-plan.csv"), PLAN_CSV);
        Files.createDirectories(Path.of(ui.get(AppPaths.KEY_PM_AI_REPO_ROOT)).resolve("code"));
        LocalDate d = LocalDate.of(2026, 7, 7);
        DispatchAladdinEntryWorkbookExporter.write(ui, matchingWorkbook(d, 99));

        Stage2IdentityCloseGate gate = new Stage2IdentityCloseGate();
        gate.markStage2Completed();

        Stage2IdentityCloseGate.Decision decision = gate.decide(ui);

        assertTrue(decision.required());
        assertEquals("差異 1件", decision.detail());
    }

    private static DispatchAladdinEntrySheetBuilder.EntryWorkbook matchingWorkbook(
            LocalDate d, double systemQty) {
        return new DispatchAladdinEntrySheetBuilder.EntryWorkbook(
                List.of(d),
                List.of(
                        new DispatchAladdinEntrySheetBuilder.MachineSheet(
                                "M1",
                                List.of(
                                        new DispatchAladdinEntrySheetBuilder.EntryRow(
                                                "T001",
                                                "",
                                                "工程A",
                                                "",
                                                "",
                                                "",
                                                systemQty,
                                                0,
                                                systemQty,
                                                Map.of(
                                                        d,
                                                        new DispatchAladdinEntrySheetBuilder
                                                                .EntryCell(0, systemQty)),
                                                d,
                                                2026)))));
    }

    private static Map<String, String> testUi(Path tempDir) {
        Path repo = tempDir.resolve("repo");
        Path sourceDir = tempDir.resolve("task-input");
        Path outputDir = tempDir.resolve("output");
        Path shared = tempDir.resolve("shared");
        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_REPO_ROOT, repo.toString());
        ui.put(AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR, sourceDir.toString());
        ui.put(AppPaths.KEY_PM_AI_OUTPUT_DIR, outputDir.toString());
        ui.put(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, shared.toString());
        ui.put(AppPaths.KEY_PM_AI_OPERATOR_USER, "テスト太郎");
        return ui;
    }
}
