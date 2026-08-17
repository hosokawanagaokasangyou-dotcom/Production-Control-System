package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.reconciliation.EcSideClassification;
import jp.co.pm.ai.desktop.ui.Stage1EcSideUnknownDialogResult;

class Stage1EcSideUnknownPromptTest {

    @TempDir Path tmp;

    @Test
    void collectsUnknownEcRowsByIraiNo() throws Exception {
        Path plan = tmp.resolve("output").resolve("plan_input_tasks.xlsx");
        Files.createDirectories(plan.getParent());
        PlanInputTabularIo.write(
                plan,
                AppPaths.STAGE1_PLAN_OUTPUT_SHEET,
                new PlanInputTabularIo.TabularSheet(
                        List.of("依頼NO", "工程名", "EC面区分"),
                        List.of(
                                List.of("W1-1", "EC", EcSideClassification.UNKNOWN),
                                List.of("W1-1", "スリット", ""),
                                List.of("W2-3", "EC", EcSideClassification.UNKNOWN),
                                List.of("W3-4", "EC", EcSideClassification.DOUBLE_SIDED))));

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        tmp.toString(),
                        AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                        plan.toString());

        var bundle = Stage1EcSideUnknownPrompt.collectUnknownIraiNos(ui);
        assertFalse(bundle.empty());
        assertEquals(2, bundle.items().size());
        assertEquals("W1-1", bundle.items().get(0).iraiNo());
        assertEquals("W2-3", bundle.items().get(1).iraiNo());
    }

    @Test
    void applyUpdatesUnknownEcRows() throws Exception {
        Path plan = tmp.resolve("output").resolve("plan_input_tasks.xlsx");
        Files.createDirectories(plan.getParent());
        PlanInputTabularIo.write(
                plan,
                AppPaths.STAGE1_PLAN_OUTPUT_SHEET,
                new PlanInputTabularIo.TabularSheet(
                        List.of("依頼NO", "工程名", "EC面区分"),
                        List.of(
                                List.of("W1-1", "EC", EcSideClassification.UNKNOWN),
                                List.of("W2-3", "EC", EcSideClassification.UNKNOWN))));

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        tmp.toString(),
                        AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                        plan.toString());

        var applied =
                Stage1EcSideUnknownPrompt.applySelections(
                        ui,
                        List.of(
                                new Stage1EcSideUnknownDialogResult.Selection(
                                        "W1-1", EcSideClassification.DOUBLE_SIDED),
                                new Stage1EcSideUnknownDialogResult.Selection(
                                        "W2-3", EcSideClassification.SINGLE_SIDED)));
        assertEquals(2, applied.rowsUpdated());

        PlanInputTabularIo.TabularRead tr =
                PlanInputTabularIo.readWithResolvedSheet(plan, AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
        List<String> headers = tr.tabular().headers();
        int iEc = headers.indexOf(EcSideClassification.COLUMN_TITLE);
        List<List<String>> rows = tr.tabular().rows();
        assertEquals(EcSideClassification.DOUBLE_SIDED, rows.get(0).get(iEc));
        assertEquals(EcSideClassification.SINGLE_SIDED, rows.get(1).get(iEc));
    }

    @Test
    void emptyWhenNoUnknownRows() throws Exception {
        Path plan = tmp.resolve("output").resolve("plan_input_tasks.xlsx");
        Files.createDirectories(plan.getParent());
        PlanInputTabularIo.write(
                plan,
                AppPaths.STAGE1_PLAN_OUTPUT_SHEET,
                new PlanInputTabularIo.TabularSheet(
                        List.of("依頼NO", "工程名", "EC面区分"),
                        List.of(List.of("W1-1", "EC", EcSideClassification.DOUBLE_SIDED))));

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        tmp.toString(),
                        AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                        plan.toString());

        assertTrue(Stage1EcSideUnknownPrompt.collectUnknownIraiNos(ui).empty());
    }
}
