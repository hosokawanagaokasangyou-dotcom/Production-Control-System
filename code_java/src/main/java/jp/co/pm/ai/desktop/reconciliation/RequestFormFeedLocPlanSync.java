package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

/**
 * 段階1出力 {@code plan_input_tasks.xlsx} の「原反投入場所」列から、依頼書入力フォームの投入場所 ComboBox 候補を補完する。
 */
public final class RequestFormFeedLocPlanSync {

    /** 段階1タスク一覧・結果_配台表など計画データ上の列名。 */
    public static final String COL_RAW_FEED_LOCATION = "原反投入場所";

    private RequestFormFeedLocPlanSync() {}

    /**
     * 計画データ（段階1出力）から非空の原反投入場所を重複なく収集する（出現順を保持）。
     */
    public static List<String> collectDistinctFeedLocations(Map<String, String> ui) throws IOException {
        Path plan = resolvePlanInputPath(ui);
        if (!Files.isRegularFile(plan)) {
            return List.of();
        }
        PlanInputTabularIo.TabularRead tr =
                PlanInputTabularIo.readWithResolvedSheet(plan, AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
        List<String> headers = tr.tabular().headers();
        List<List<String>> rows = tr.tabular().rows();
        if (headers == null || headers.isEmpty() || rows == null || rows.isEmpty()) {
            return List.of();
        }
        int col = headers.indexOf(COL_RAW_FEED_LOCATION);
        if (col < 0) {
            return List.of();
        }
        Set<String> seen = new LinkedHashSet<>();
        List<String> out = new ArrayList<>();
        for (List<String> row : rows) {
            String value = cellAt(row, col);
            if (value.isEmpty() || seen.contains(value)) {
                continue;
            }
            seen.add(value);
            out.add(value);
        }
        return List.copyOf(out);
    }

    /** {@code base} に {@code additions} を末尾追記（既存値は維持、空白は無視）。 */
    public static List<String> mergeDistinctFeedLocations(List<String> base, List<String> additions) {
        LinkedHashSet<String> merged = new LinkedHashSet<>();
        if (base != null) {
            for (String value : base) {
                addIfNonBlank(merged, value);
            }
        }
        if (additions != null) {
            for (String value : additions) {
                addIfNonBlank(merged, value);
            }
        }
        return List.copyOf(merged);
    }

    public static int countNewValues(List<String> before, List<String> after) {
        if (after == null || after.isEmpty()) {
            return 0;
        }
        Set<String> prior = new LinkedHashSet<>();
        if (before != null) {
            for (String value : before) {
                addIfNonBlank(prior, value);
            }
        }
        int added = 0;
        for (String value : after) {
            if (value == null) {
                continue;
            }
            String text = value.strip();
            if (text.isEmpty() || prior.contains(text)) {
                continue;
            }
            prior.add(text);
            added++;
        }
        return added;
    }

    private static void addIfNonBlank(Set<String> target, String raw) {
        if (raw == null) {
            return;
        }
        String text = raw.strip();
        if (!text.isEmpty()) {
            target.add(text);
        }
    }

    private static String cellAt(List<String> row, int col) {
        if (row == null || col < 0 || col >= row.size()) {
            return "";
        }
        String value = row.get(col);
        return value != null ? value.strip() : "";
    }

    private static Path resolvePlanInputPath(Map<String, String> ui) {
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
}
