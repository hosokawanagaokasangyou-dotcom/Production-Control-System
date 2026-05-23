package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.ui.TabularCellHighlight;

class Stage1DevMarkAllExcludeAfterRunTest {

    @TempDir Path tmp;

    @Test
    void marksAllTaskRowsExcludeYes() throws Exception {
        Path plan = tmp.resolve("output").resolve("plan_input_tasks.xlsx");
        Files.createDirectories(plan.getParent());
        PlanInputTabularIo.write(
                plan,
                AppPaths.STAGE1_PLAN_OUTPUT_SHEET,
                new PlanInputTabularIo.TabularSheet(
                        List.of("依頼NO", "工程名", "機械名", "配台不要"),
                        List.of(
                                List.of("W5-6", "巻返し", "フィルム挿入機(間紙)", ""),
                                List.of("Y5-186", "スライス", "スライス機3", "いいえ"),
                                List.of("Y5-187", "スライス", "スライス機3", "yes"))));

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_REPO_ROOT,
                        tmp.toString(),
                        AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                        plan.toString());

        var summary = Stage1DevMarkAllExcludeAfterRun.applyToPlanInput(ui);
        assertEquals(3, summary.totalRows());
        assertEquals(2, summary.updatedRows());

        PlanInputTabularIo.TabularRead tr =
                PlanInputTabularIo.readWithResolvedSheet(plan, AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
        int iEx = tr.tabular().headers().indexOf("配台不要");
        for (List<String> row : tr.tabular().rows()) {
            assertTrue(
                    TabularCellHighlight.planInputExcludeFromAssignmentIsOn(row.get(iEx)),
                    "row should be exclude=yes: " + row);
        }
    }
}
