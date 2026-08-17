package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.reconciliation.EcSideClassification;
import jp.co.pm.ai.desktop.ui.Stage1EcSideUnknownDialogResult;

/**
 * 段階1完了後: 配台計画タスク入力の EC面区分が「不明」の依頼NOを収集し、
 * ユーザー選択に応じて xlsx を更新する。
 */
public final class Stage1EcSideUnknownPrompt {

    private static final String COL_PROCESS = "工程名";
    private static final String COL_TASK = "依頼NO";

    public record UnknownIrai(String iraiNo) {}

    public record PromptBundle(List<UnknownIrai> items) {
        public boolean empty() {
            return items == null || items.isEmpty();
        }
    }

    public record ApplySummary(int rowsUpdated) {}

    private Stage1EcSideUnknownPrompt() {}

    private static boolean isEcSideClassificationProcess(String process) {
        return "EC".equals(process) || "SEC".equals(process);
    }

    public static PromptBundle collectUnknownIraiNos(Map<String, String> ui) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path plan = resolvePlanInputPath(u);
        if (!Files.isRegularFile(plan)) {
            return new PromptBundle(List.of());
        }
        PlanInputTabularIo.TabularRead tr =
                PlanInputTabularIo.readWithResolvedSheet(plan, AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
        List<String> headers = tr.tabular().headers();
        List<List<String>> rows = tr.tabular().rows();
        if (headers == null || headers.isEmpty() || rows == null) {
            return new PromptBundle(List.of());
        }
        int iProc = headers.indexOf(COL_PROCESS);
        int iTask = headers.indexOf(COL_TASK);
        int iEc = headers.indexOf(EcSideClassification.COLUMN_TITLE);
        if (iProc < 0 || iTask < 0 || iEc < 0) {
            return new PromptBundle(List.of());
        }

        LinkedHashSet<String> unknown = new LinkedHashSet<>();
        for (List<String> row : rows) {
            if (!isEcSideClassificationProcess(cell(row, iProc))) {
                continue;
            }
            if (!EcSideClassification.UNKNOWN.equals(cell(row, iEc))) {
                continue;
            }
            String tid = cell(row, iTask);
            if (!tid.isBlank()) {
                unknown.add(tid.strip());
            }
        }
        List<UnknownIrai> items = new ArrayList<>();
        for (String tid : unknown) {
            items.add(new UnknownIrai(tid));
        }
        return new PromptBundle(List.copyOf(items));
    }

    public static ApplySummary applySelections(
            Map<String, String> ui, List<Stage1EcSideUnknownDialogResult.Selection> selections)
            throws IOException {
        if (selections == null || selections.isEmpty()) {
            return new ApplySummary(0);
        }
        Map<String, String> u = ui != null ? ui : Map.of();
        Map<String, String> byIrai = new LinkedHashMap<>();
        for (Stage1EcSideUnknownDialogResult.Selection sel : selections) {
            if (sel == null || sel.iraiNo() == null || sel.iraiNo().isBlank()) {
                continue;
            }
            String choice = sel.ecSideClass() != null ? sel.ecSideClass().strip() : "";
            if (!EcSideClassification.DOUBLE_SIDED.equals(choice)
                    && !EcSideClassification.SINGLE_SIDED.equals(choice)) {
                continue;
            }
            byIrai.put(sel.iraiNo().strip(), choice);
        }
        if (byIrai.isEmpty()) {
            return new ApplySummary(0);
        }

        Path plan = resolvePlanInputPath(u);
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
        int iProc = headers.indexOf(COL_PROCESS);
        int iTask = headers.indexOf(COL_TASK);
        int iEc = headers.indexOf(EcSideClassification.COLUMN_TITLE);
        if (iProc < 0 || iTask < 0 || iEc < 0) {
            return new ApplySummary(0);
        }

        int updated = 0;
        for (List<String> row : rows) {
            while (row.size() < headers.size()) {
                row.add("");
            }
            if (!isEcSideClassificationProcess(cell(row, iProc))) {
                continue;
            }
            if (!EcSideClassification.UNKNOWN.equals(cell(row, iEc))) {
                continue;
            }
            String tid = cell(row, iTask);
            String chosen = byIrai.get(tid);
            if (chosen == null) {
                continue;
            }
            row.set(iEc, chosen);
            updated++;
        }

        String sheet =
                tr.resolvedSheetName() != null && !tr.resolvedSheetName().isBlank()
                        ? tr.resolvedSheetName()
                        : AppPaths.STAGE1_PLAN_OUTPUT_SHEET;
        PlanInputTabularIo.write(
                plan, sheet, new PlanInputTabularIo.TabularSheet(headers, rows));
        return new ApplySummary(updated);
    }

    static Path resolvePlanInputPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String pip = u.get(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH);
        if (pip != null && !pip.isBlank()) {
            Path p = Path.of(pip.strip());
            if (Files.isRegularFile(p)) {
                return p.toAbsolutePath().normalize();
            }
        }
        return AppPaths.defaultStage1PlanTasksPath(u).toAbsolutePath().normalize();
    }

    private static String cell(List<String> row, int col) {
        if (row == null || col < 0 || col >= row.size()) {
            return "";
        }
        String v = row.get(col);
        return v != null ? v.strip() : "";
    }
}
