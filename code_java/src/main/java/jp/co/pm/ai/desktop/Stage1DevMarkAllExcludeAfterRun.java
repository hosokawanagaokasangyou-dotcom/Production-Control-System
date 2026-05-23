package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.ui.TabularCellHighlight;

/**
 * 段階1正常終了後（開発用）: 配台計画_タスク入力の全行「配台不要」を yes にする。
 * 段階2はシート上の「配台不要」列をそのまま参照するため、ここでの更新が配台対象の正本になる。
 */
public final class Stage1DevMarkAllExcludeAfterRun {

    private static final String COL_TASK = "依頼NO";
    private static final String COL_PROCESS = "工程名";
    private static final String COL_MACHINE = "機械名";
    private static final String COL_EXCLUDE = "配台不要";

    public record ApplySummary(int totalRows, int updatedRows, Path planPath) {}

    private Stage1DevMarkAllExcludeAfterRun() {}

    public static ApplySummary applyToPlanInput(Map<String, String> ui) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path plan = Stage2UnknownMasterCombinationPrompt.resolvePlanInputPath(u);
        if (!Files.isRegularFile(plan)) {
            throw new IOException("計画タスク入力が見つかりません: " + plan);
        }
        PlanInputTabularIo.TabularRead tr =
                PlanInputTabularIo.readWithResolvedSheet(plan, AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
        List<String> headers = new ArrayList<>(tr.tabular().headers());
        List<List<String>> rows = new ArrayList<>();
        for (List<String> src : tr.tabular().rows()) {
            rows.add(new ArrayList<>(src));
        }
        int iTask = headers.indexOf(COL_TASK);
        int iProc = headers.indexOf(COL_PROCESS);
        int iMach = headers.indexOf(COL_MACHINE);
        int iEx = headers.indexOf(COL_EXCLUDE);
        if (iEx < 0) {
            headers.add(COL_EXCLUDE);
            iEx = headers.size() - 1;
            for (List<String> row : rows) {
                while (row.size() < headers.size()) {
                    row.add("");
                }
            }
        }
        int total = 0;
        int updated = 0;
        for (List<String> row : rows) {
            while (row.size() < headers.size()) {
                row.add("");
            }
            if (!isTaskRow(row, iTask, iProc, iMach)) {
                continue;
            }
            total++;
            if (TabularCellHighlight.planInputExcludeFromAssignmentIsOn(cell(row, iEx))) {
                continue;
            }
            row.set(iEx, "yes");
            updated++;
        }
        String sheet =
                tr.resolvedSheetName() != null && !tr.resolvedSheetName().isBlank()
                        ? tr.resolvedSheetName()
                        : AppPaths.STAGE1_PLAN_OUTPUT_SHEET;
        PlanInputTabularIo.write(
                plan, sheet, new PlanInputTabularIo.TabularSheet(headers, rows));
        return new ApplySummary(total, updated, plan.toAbsolutePath().normalize());
    }

    /** 依頼NO または 工程+機械 が埋まっている行をタスク行とみなす。 */
    private static boolean isTaskRow(List<String> row, int iTask, int iProc, int iMach) {
        if (row == null || row.isEmpty()) {
            return false;
        }
        if (iTask >= 0 && !cell(row, iTask).isBlank()) {
            return true;
        }
        if (iProc >= 0 && iMach >= 0) {
            return !cell(row, iProc).isBlank() && !cell(row, iMach).isBlank();
        }
        return row.stream().anyMatch(c -> c != null && !c.isBlank());
    }

    private static String cell(List<String> row, int col) {
        if (row == null || col < 0 || col >= row.size()) {
            return "";
        }
        String v = row.get(col);
        return v != null ? v.strip() : "";
    }
}
