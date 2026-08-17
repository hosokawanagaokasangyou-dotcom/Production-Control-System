package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.reconciliation.EcSideClassification;

/** リポジトリ output/plan_input_tasks.xlsx の W8-18 EC面区分が Java POI で読めることを確認。 */
class PlanInputEcSideW818IntegrationTest {

    @Test
    void readW818EcSideFromRepoOutputIfPresent() throws Exception {
        Path repoRoot = Path.of("..").toAbsolutePath().normalize();
        Path plan = repoRoot.resolve("output").resolve(AppPaths.STAGE1_PLAN_TASKS_FILENAME);
        if (!Files.isRegularFile(plan)) {
            return;
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
