package jp.co.pm.ai.desktop.io;

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
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.DispatchAladdinEntrySheetBuilder;
import jp.co.pm.ai.desktop.reconciliation.PostProcessingPlanMachineLookup;

class AladdinEntryDispatchPlanIdentityCheckTest {

    private static final String PLAN_CSV =
            "列1,列2,列3,列4\n"
                    + "上段1,,,\n"
                    + "上段2,,,\n"
                    + "上段3,,,\n"
                    + "機械名,依頼NO,工程名,2026/07/07\n"
                    + ",,,\n"
                    + "M1,T001,工程A,10\n";

    @Test
    void compare_identicalWhenSystemQtyMatchesPlan() {
        LocalDate d = LocalDate.of(2026, 7, 7);
        List<AladdinEntryDispatchPlanIdentityCheck.SystemQty> system =
                List.of(
                        new AladdinEntryDispatchPlanIdentityCheck.SystemQty(
                                "M1", "T001", "工程A", d, 10));
        Map<String, Map<String, Map<String, Map<String, Double>>>> lookup =
                AladdinShapedPlanQtyLookup.buildLookup(
                        List.of("機械名", "依頼NO", "工程名", "2026/07/07"),
                        List.of(List.of("M1", "T001", "工程A", "10")));

        AladdinEntryDispatchPlanIdentityCheck.Result result =
                AladdinEntryDispatchPlanIdentityCheck.compare(system, lookup);

        assertTrue(result.identical());
        assertTrue(result.diffs().isEmpty());
        assertEquals("配台計画と加工計画は同一", result.badgeText());
    }

    @Test
    void compare_mismatchWhenSystemQtyDiffersFromPlan() {
        LocalDate d = LocalDate.of(2026, 7, 7);
        List<AladdinEntryDispatchPlanIdentityCheck.SystemQty> system =
                List.of(
                        new AladdinEntryDispatchPlanIdentityCheck.SystemQty(
                                "M1", "T001", "工程A", d, 10));
        Map<String, Map<String, Map<String, Map<String, Double>>>> lookup =
                AladdinShapedPlanQtyLookup.buildLookup(
                        List.of("機械名", "依頼NO", "工程名", "2026/07/07"),
                        List.of(List.of("M1", "T001", "工程A", "99")));

        AladdinEntryDispatchPlanIdentityCheck.Result result =
                AladdinEntryDispatchPlanIdentityCheck.compare(system, lookup);

        assertFalse(result.identical());
        assertEquals(1, result.diffs().size());
        assertEquals("差異 1件", result.badgeText());
        assertEquals(10d, result.diffs().getFirst().systemQty(), 1e-9);
        assertEquals(99d, result.diffs().getFirst().planQty(), 1e-9);
    }

    @Test
    void compare_mismatchWhenPlanHasExtraQty() {
        LocalDate d = LocalDate.of(2026, 7, 7);
        List<AladdinEntryDispatchPlanIdentityCheck.SystemQty> system =
                List.of(
                        new AladdinEntryDispatchPlanIdentityCheck.SystemQty(
                                "M1", "T001", "工程A", d, 10));
        Map<String, Map<String, Map<String, Map<String, Double>>>> lookup =
                AladdinShapedPlanQtyLookup.buildLookup(
                        List.of("機械名", "依頼NO", "工程名", "2026/07/07"),
                        List.of(
                                List.of("M1", "T001", "工程A", "10"),
                                List.of("M1", "T002", "工程A", "5")));

        AladdinEntryDispatchPlanIdentityCheck.Result result =
                AladdinEntryDispatchPlanIdentityCheck.compare(system, lookup);

        assertFalse(result.identical());
        assertEquals(1, result.diffs().size());
        assertEquals("T002", result.diffs().getFirst().taskId());
        assertEquals(0d, result.diffs().getFirst().systemQty(), 1e-9);
        assertEquals(5d, result.diffs().getFirst().planQty(), 1e-9);
    }

    @Test
    void evaluate_errorWhenOperatorHasNoGeneration(@TempDir Path tempDir) throws Exception {
        Map<String, String> ui = testUi(tempDir);
        Files.createDirectories(AppPaths.resolveTaskInputSourceDir(ui));

        AladdinEntryDispatchPlanIdentityCheck.Result result =
                AladdinEntryDispatchPlanIdentityCheck.evaluate(ui);

        assertTrue(result.error());
        assertFalse(result.identical());
        assertEquals(
                AladdinEntryDispatchPlanIdentityCheck.ERROR_NO_GENERATION, result.message());
    }

    @Test
    void evaluate_identicalForOperatorGenerationAndMatchingPlan(@TempDir Path tempDir)
            throws Exception {
        Map<String, String> ui = testUi(tempDir);
        Path sourceDir = AppPaths.resolveTaskInputSourceDir(ui);
        Files.createDirectories(sourceDir);
        Files.writeString(sourceDir.resolve("aladdin-plan.csv"), PLAN_CSV);
        Files.createDirectories(Path.of(ui.get(AppPaths.KEY_PM_AI_REPO_ROOT)).resolve("code"));

        LocalDate d = LocalDate.of(2026, 7, 7);
        DispatchAladdinEntrySheetBuilder.EntryWorkbook model =
                new DispatchAladdinEntrySheetBuilder.EntryWorkbook(
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
                                                        10,
                                                        0,
                                                        10,
                                                        Map.of(
                                                                d,
                                                                new DispatchAladdinEntrySheetBuilder
                                                                        .EntryCell(0, 10)),
                                                        d,
                                                        2026)))));
        DispatchAladdinEntryWorkbookExporter.write(ui, model);

        Path shapedBefore = AppPaths.resolveShapedAladdinPlanJsonPath(ui);
        boolean shapedExisted = Files.isRegularFile(shapedBefore);

        AladdinEntryDispatchPlanIdentityCheck.Result result =
                AladdinEntryDispatchPlanIdentityCheck.evaluate(ui);

        assertFalse(result.error(), result.message());
        assertTrue(result.identical(), result.dialogBody());
        assertTrue(result.excelPath().isPresent());
        assertTrue(result.planSourcePath().isPresent());
        assertEquals(shapedExisted, Files.isRegularFile(shapedBefore));
    }

    @Test
    void readSystemQtys_readsSystemLineFromExportedWorkbook(@TempDir Path tempDir) throws Exception {
        Map<String, String> ui = testUi(tempDir);
        Files.createDirectories(Path.of(ui.get(AppPaths.KEY_PM_AI_REPO_ROOT)).resolve("code"));
        LocalDate d = LocalDate.of(2026, 7, 7);
        DispatchAladdinEntrySheetBuilder.EntryWorkbook model =
                new DispatchAladdinEntrySheetBuilder.EntryWorkbook(
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
                                                        10,
                                                        0,
                                                        10,
                                                        Map.of(
                                                                d,
                                                                new DispatchAladdinEntrySheetBuilder
                                                                        .EntryCell(3, 10)),
                                                        d,
                                                        2026)))));
        DispatchAladdinEntryWorkbookExporter.ExportResult exported =
                DispatchAladdinEntryWorkbookExporter.write(ui, model);

        List<AladdinEntryDispatchPlanIdentityCheck.SystemQty> qtys =
                AladdinEntryDispatchPlanWorkbookReader.readSystemQtys(
                        exported.generationPath(), d);

        assertEquals(1, qtys.size());
        assertEquals("T001", qtys.getFirst().taskId());
        assertEquals("工程A", qtys.getFirst().processName());
        assertEquals(d, qtys.getFirst().date());
        assertEquals(10d, qtys.getFirst().qty(), 1e-9);
    }

    @Test
    void resolveMachineNameFromSheet_mapsDisplayLabelToMachineName() {
        PostProcessingPlanMachineLookup.Snapshot snap =
                new PostProcessingPlanMachineLookup.Snapshot(
                        Path.of(""),
                        -1L,
                        true,
                        true,
                        Map.of("C01", "M1"),
                        Map.of("M1", "C01"),
                        List.of("C01 M1"));

        assertEquals(
                "M1",
                AladdinEntryDispatchPlanIdentityCheck.resolveMachineNameFromSheet("C01 M1", snap));
        assertEquals(
                "M1", AladdinEntryDispatchPlanIdentityCheck.resolveMachineNameFromSheet("M1", snap));
    }

    @Test
    void parseSystemQty_readsLowerLine() {
        assertEquals(
                200d,
                AladdinEntryDispatchPlanWorkbookReader.parseSystemQty("（現アラ計）110\n（シス計）200"),
                1e-9);
        assertEquals(0d, AladdinEntryDispatchPlanWorkbookReader.parseSystemQty(""), 1e-9);
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
