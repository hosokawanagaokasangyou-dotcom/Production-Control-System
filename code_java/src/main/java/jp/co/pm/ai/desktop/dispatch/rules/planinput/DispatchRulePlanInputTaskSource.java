package jp.co.pm.ai.desktop.dispatch.rules.planinput;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import jp.co.pm.ai.desktop.PlanInputTabController;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

/** Loads plan_input task rows for the rule test lab (memory first, then disk). */
public final class DispatchRulePlanInputTaskSource {

    private final Map<String, Map<String, String>> rowsByLabel = new LinkedHashMap<>();
    private List<String> labels = List.of();
    private String sourceDescription = "";

    public void reload(Map<String, String> ui, PlanInputTabController planInputOrNull) {
        rowsByLabel.clear();
        labels = List.of();
        sourceDescription = "";
        if (planInputOrNull != null) {
            List<String> mem = planInputOrNull.listPlanInputTaskLabels();
            if (!mem.isEmpty()) {
                labels = List.copyOf(mem);
                for (String label : mem) {
                    planInputOrNull.findPlanRowMapByLabel(label).ifPresent(row -> rowsByLabel.put(label, row));
                }
                sourceDescription = "配台計画_タスク入力タブ（メモリ）";
                return;
            }
        }
        Path path = resolvePlanInputPath(ui);
        if (path == null || !Files.isRegularFile(path)) {
            sourceDescription =
                    "plan_input 未読込（"
                            + (path != null ? path : "パス未設定")
                            + " が見つかりません）";
            return;
        }
        String sheet = resolveSheetName(ui, planInputOrNull);
        try {
            PlanInputTabularIo.TabularRead read = PlanInputTabularIo.readWithResolvedSheet(path, sheet);
            PlanInputTabularIo.TabularSheet tabular = read.tabular();
            List<String> headers = tabular.headers();
            if (headers.isEmpty()) {
                sourceDescription = "ヘッダなし: " + path;
                return;
            }
            int colTask = headers.indexOf("依頼NO");
            int colProcess = headers.indexOf("工程名");
            int colMachine = headers.indexOf("機械名");
            List<String> built = new ArrayList<>();
            for (List<String> cells : tabular.rows()) {
                String tid = cellAt(cells, colTask);
                if (tid.isEmpty()) {
                    continue;
                }
                String proc = cellAt(cells, colProcess);
                String mach = cellAt(cells, colMachine);
                String label = tid + " / " + proc + " / " + mach;
                LinkedHashMap<String, String> row = new LinkedHashMap<>();
                for (int c = 0; c < headers.size(); c++) {
                    row.put(headers.get(c), cellAt(cells, c));
                }
                row.putIfAbsent("task_id", tid);
                rowsByLabel.put(label, row);
                built.add(label);
            }
            labels = List.copyOf(built);
            sourceDescription =
                    path.getFileName()
                            + (read.resolvedSheetName().isBlank()
                                    ? ""
                                    : " [" + read.resolvedSheetName() + "]");
        } catch (IOException ex) {
            sourceDescription = "読込失敗: " + ex.getMessage();
        }
    }

    public List<String> labels() {
        return labels;
    }

    public String sourceDescription() {
        return sourceDescription;
    }

    public Optional<Map<String, String>> findRowByLabel(String label) {
        if (label == null || label.isBlank()) {
            return Optional.empty();
        }
        Map<String, String> row = rowsByLabel.get(label);
        if (row != null) {
            return Optional.of(new LinkedHashMap<>(row));
        }
        return Optional.empty();
    }

    /** 同一依頼NOの SEC 行（接続→SEC 試走用）。 */
    public Optional<Map<String, String>> findSecRowForRequest(String requestNo) {
        if (requestNo == null || requestNo.isBlank()) {
            return Optional.empty();
        }
        for (Map<String, String> row : rowsByLabel.values()) {
            if (requestNo.equals(trim(row.get("依頼NO"))) && "SEC".equals(trim(row.get("工程名")))) {
                return Optional.of(new LinkedHashMap<>(row));
            }
        }
        return Optional.empty();
    }

    public static boolean isConnectionProcess(Map<String, String> row) {
        return row != null && "接続".equals(trim(row.get("工程名")));
    }

    /** plan_input 行の配台ロール数（試走ラボ用）。 */
    public static int parseRollCount(Map<String, String> row) {
        if (row == null || row.isEmpty()) {
            return 1;
        }
        String raw = row.getOrDefault("配台ロール数", row.getOrDefault("dispatch_roll_count", "1"));
        try {
            return Math.max(1, (int) Double.parseDouble(raw.strip()));
        } catch (NumberFormatException ex) {
            return 1;
        }
    }

    private static Path resolvePlanInputPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String pip = trim(u.get(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH));
        if (!pip.isEmpty()) {
            Path p = Path.of(pip).toAbsolutePath().normalize();
            if (Files.isRegularFile(p)) {
                return p;
            }
        }
        Path def = AppPaths.defaultStage1PlanTasksPath(u);
        if (Files.isRegularFile(def)) {
            return def;
        }
        Path repoOut =
                AppPaths.resolveRepoRoot(u).resolve("output").resolve(AppPaths.STAGE1_PLAN_TASKS_FILENAME);
        if (Files.isRegularFile(repoOut)) {
            return repoOut.toAbsolutePath().normalize();
        }
        return def;
    }

    private static String resolveSheetName(Map<String, String> ui, PlanInputTabController planInputOrNull) {
        if (planInputOrNull != null) {
            String s = trim(planInputOrNull.snapshotPlanInputSheet());
            if (!s.isEmpty()) {
                return s;
            }
        }
        Map<String, String> u = ui != null ? ui : Map.of();
        String envSheet = trim(u.get(PlanInputTabController.ENV_TASK_PLAN_SHEET));
        if (!envSheet.isEmpty()) {
            return envSheet;
        }
        return AppPaths.STAGE1_PLAN_OUTPUT_SHEET;
    }

    private static String cellAt(List<String> cells, int col) {
        if (col < 0 || col >= cells.size()) {
            return "";
        }
        String v = cells.get(col);
        return v != null ? v.strip() : "";
    }

    private static String trim(String s) {
        return s != null ? s.strip() : "";
    }
}
