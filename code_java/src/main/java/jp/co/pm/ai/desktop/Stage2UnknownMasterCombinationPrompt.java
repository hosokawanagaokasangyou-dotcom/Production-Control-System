package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.ui.TabularCellHighlight;

/**
 * 段階2前: 計画タスクの工程+機械が master「組み合わせ表」に無い組み合わせを検出し、
 * ユーザー選択に応じて配台不要 JSON と plan_input の「配台不要」列を更新する。
 */
public final class Stage2UnknownMasterCombinationPrompt {

    private static final String COL_PROCESS = "工程名";
    private static final String COL_MACHINE = "機械名";
    private static final String COL_TASK = "依頼NO";
    private static final String COL_EXCLUDE = "配台不要";

    private static final String JSON_PROCESS = "工程名";
    private static final String JSON_MACHINE = "機械名";
    private static final String JSON_FLAG = "配台不要";

    private static final ObjectMapper JSON =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    public record UnknownPair(String process, String machine, String sampleTaskId) {}

    public record PromptBundle(List<UnknownPair> pairs) {
        public boolean empty() {
            return pairs == null || pairs.isEmpty();
        }
    }

    public record ApplySummary(int excludeRulesUpdated, int planRowsUpdated) {}

    private Stage2UnknownMasterCombinationPrompt() {}

    public static PromptBundle collectUnknownPairs(Map<String, String> ui) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path master = resolveMasterWorkbook(u);
        if (master == null || !Files.isRegularFile(master)) {
            return new PromptBundle(List.of());
        }
        Set<String> masterKeys = MasterTeamCombinationTableReader.readNormalizedComboKeys(master);
        if (masterKeys.isEmpty()) {
            return new PromptBundle(List.of());
        }
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
        int iMach = headers.indexOf(COL_MACHINE);
        int iTask = headers.indexOf(COL_TASK);
        int iEx = headers.indexOf(COL_EXCLUDE);
        if (iProc < 0 || iMach < 0) {
            return new PromptBundle(List.of());
        }

        LinkedHashMap<String, UnknownPair> unknown = new LinkedHashMap<>();
        for (List<String> row : rows) {
            if (TabularCellHighlight.planInputExcludeFromAssignmentIsOn(cell(row, iEx))) {
                continue;
            }
            String proc = cell(row, iProc);
            String mach = cell(row, iMach);
            if (proc.isBlank() || mach.isBlank()) {
                continue;
            }
            String nk = MasterTeamCombinationTableReader.normalizedComboKey(proc, mach);
            if (nk.isEmpty() || masterKeys.contains(nk)) {
                continue;
            }
            unknown.putIfAbsent(
                    nk,
                    new UnknownPair(
                            proc.strip(),
                            mach.strip(),
                            iTask >= 0 ? cell(row, iTask) : ""));
        }
        return new PromptBundle(List.copyOf(unknown.values()));
    }

    public static ApplySummary applyExcludeSelections(
            Map<String, String> ui, List<UnknownPair> markExclude) throws IOException {
        if (markExclude == null || markExclude.isEmpty()) {
            return new ApplySummary(0, 0);
        }
        Map<String, String> u = ui != null ? ui : Map.of();
        int jsonN = updateExcludeRulesJson(u, markExclude);
        int planN = updatePlanInputExcludeColumn(u, markExclude);
        return new ApplySummary(jsonN, planN);
    }

    public static Optional<Path> resolveExcludeRulesJsonPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String env = u.get(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON);
        if (env != null && !env.isBlank()) {
            Path p = Path.of(env.strip());
            if (Files.isRegularFile(p)) {
                return Optional.of(p.toAbsolutePath().normalize());
            }
        }
        AppPaths.ensureStage1ExcludeRulesJsonFromRepoIfMissing(u);
        return AppPaths.resolveDefaultExcludeRulesJsonPath(u).map(Path::toAbsolutePath).map(Path::normalize);
    }

    private static int updateExcludeRulesJson(Map<String, String> ui, List<UnknownPair> pairs)
            throws IOException {
        Optional<Path> pathOpt = resolveExcludeRulesJsonPath(ui);
        if (pathOpt.isEmpty()) {
            throw new IOException("配台不要ルール JSON のパスが解決できません（PM_AI_EXCLUDE_RULES_JSON）。");
        }
        Path path = pathOpt.get();
        JsonNode root;
        if (Files.isRegularFile(path)) {
            root = JSON.readTree(Files.readString(path, StandardCharsets.UTF_8));
        } else {
            ObjectNode wrapper = JSON.createObjectNode();
            wrapper.set("rules", JSON.createArrayNode());
            root = wrapper;
            Path parent = path.getParent();
            if (parent != null) {
                Files.createDirectories(parent);
            }
        }

        ArrayNode rules;
        ObjectNode top;
        if (root.isArray()) {
            rules = (ArrayNode) root;
            top = JSON.createObjectNode();
            top.set("rules", rules);
        } else if (root.isObject()) {
            top = (ObjectNode) root;
            JsonNode r = top.get("rules");
            if (r instanceof ArrayNode arr) {
                rules = arr;
            } else {
                rules = JSON.createArrayNode();
                top.set("rules", rules);
            }
        } else {
            throw new IOException("配台不要ルール JSON の形式が不正です: " + path);
        }

        Set<String> existing = new LinkedHashSet<>();
        for (JsonNode row : rules) {
            if (!row.isObject()) {
                continue;
            }
            String proc = text(row.get(JSON_PROCESS));
            String mach = text(row.get(JSON_MACHINE));
            if (proc.isEmpty()) {
                continue;
            }
            existing.add(pairKey(proc, mach));
        }

        int updated = 0;
        for (UnknownPair p : pairs) {
            String key = pairKey(p.process(), p.machine());
            if (key.isEmpty()) {
                continue;
            }
            ObjectNode rec = findRuleRow(rules, p.process(), p.machine());
            if (rec != null) {
                rec.put(JSON_FLAG, "yes");
                updated++;
                continue;
            }
            if (existing.contains(key)) {
                continue;
            }
            ObjectNode added = JSON.createObjectNode();
            added.put(JSON_PROCESS, p.process());
            added.put(JSON_MACHINE, p.machine());
            added.put(JSON_FLAG, "yes");
            added.putNull("配台不要ロジック");
            added.putNull("ロジック式");
            rules.add(added);
            existing.add(key);
            updated++;
        }

        Files.writeString(
                path,
                JSON.writerWithDefaultPrettyPrinter().writeValueAsString(top) + "\n",
                StandardCharsets.UTF_8);
        return updated;
    }

    private static ObjectNode findRuleRow(ArrayNode rules, String process, String machine) {
        String want = pairKey(process, machine);
        for (JsonNode row : rules) {
            if (!row.isObject()) {
                continue;
            }
            String proc = text(row.get(JSON_PROCESS));
            if (proc.isEmpty()) {
                continue;
            }
            String mach = text(row.get(JSON_MACHINE));
            if (want.equals(pairKey(proc, mach))) {
                return (ObjectNode) row;
            }
        }
        return null;
    }

    private static int updatePlanInputExcludeColumn(Map<String, String> ui, List<UnknownPair> pairs)
            throws IOException {
        Path plan = resolvePlanInputPath(ui);
        if (!Files.isRegularFile(plan)) {
            throw new IOException("計画タスク入力が見つかりません: " + plan);
        }
        Set<String> markKeys = new LinkedHashSet<>();
        for (UnknownPair p : pairs) {
            String k = pairKey(p.process(), p.machine());
            if (!k.isEmpty()) {
                markKeys.add(k);
            }
        }
        PlanInputTabularIo.TabularRead tr =
                PlanInputTabularIo.readWithResolvedSheet(plan, AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
        List<String> headers = new ArrayList<>(tr.tabular().headers());
        List<List<String>> rows = new ArrayList<>();
        for (List<String> src : tr.tabular().rows()) {
            rows.add(new ArrayList<>(src));
        }
        int iProc = headers.indexOf(COL_PROCESS);
        int iMach = headers.indexOf(COL_MACHINE);
        int iEx = headers.indexOf(COL_EXCLUDE);
        if (iProc < 0 || iMach < 0) {
            return 0;
        }
        if (iEx < 0) {
            headers.add(COL_EXCLUDE);
            iEx = headers.size() - 1;
            for (List<String> row : rows) {
                while (row.size() < headers.size()) {
                    row.add("");
                }
            }
        }
        int updated = 0;
        for (List<String> row : rows) {
            while (row.size() < headers.size()) {
                row.add("");
            }
            String proc = cell(row, iProc);
            String mach = cell(row, iMach);
            String key = pairKey(proc, mach);
            if (!markKeys.contains(key)) {
                continue;
            }
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
        return updated;
    }

    private static String pairKey(String process, String machine) {
        return MasterTeamCombinationTableReader.normalizedComboKey(process, machine);
    }

    private static String text(JsonNode n) {
        if (n == null || n.isNull()) {
            return "";
        }
        return n.asText("").strip();
    }

    private static String cell(List<String> row, int col) {
        if (row == null || col < 0 || col >= row.size()) {
            return "";
        }
        String v = row.get(col);
        return v != null ? v.strip() : "";
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

    static Path resolveMasterWorkbook(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String masterEnv = u.get(AppPaths.KEY_PM_AI_MASTER_WORKBOOK);
        if (masterEnv != null && !masterEnv.isBlank()) {
            Path p = Path.of(masterEnv.strip());
            if (Files.isRegularFile(p)) {
                return p.toAbsolutePath().normalize();
            }
        }
        return AppPaths.resolveMasterWorkbookPathForDesktopOpen(u, "").toAbsolutePath().normalize();
    }
}
