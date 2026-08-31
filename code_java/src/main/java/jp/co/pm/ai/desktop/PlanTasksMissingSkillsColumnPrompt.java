package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.SkillsSheetEquipmentListReader;
import jp.co.pm.ai.desktop.ui.MasterDispatchSheetEditRules;
import jp.co.pm.ai.desktop.ui.TabularCellHighlight;

/**
 * 計画タスクの工程+機械が master「skills」シートに列として無い組み合わせを検出する。
 *
 * <p>該当行は OP/AS スキルが割り当てられず段階2で配台できない。
 */
public final class PlanTasksMissingSkillsColumnPrompt {

    private static final String COL_PROCESS = "工程名";
    private static final String COL_MACHINE = "機械名";
    private static final String COL_TASK = "依頼NO";
    private static final String COL_EXCLUDE = "配台不要";

    public record MissingPair(String process, String machine, String sampleTaskId) {}

    public record PromptBundle(List<MissingPair> pairs) {
        public boolean empty() {
            return pairs == null || pairs.isEmpty();
        }

        /** ユーザー向け要約（最大 {@code maxLines} 件まで列挙）。 */
        public String summaryJa(int maxLines) {
            if (empty()) {
                return "";
            }
            int limit = Math.max(1, maxLines);
            StringBuilder sb = new StringBuilder();
            int n = 0;
            for (MissingPair p : pairs) {
                if (n >= limit) {
                    sb.append("… 他 ").append(pairs.size() - limit).append(" 件");
                    break;
                }
                if (n > 0) {
                    sb.append('\n');
                }
                sb.append("・").append(p.process()).append(" × ").append(p.machine());
                if (p.sampleTaskId() != null && !p.sampleTaskId().isBlank()) {
                    sb.append("（例: ").append(p.sampleTaskId()).append("）");
                }
                n++;
            }
            return sb.toString();
        }
    }

    private PlanTasksMissingSkillsColumnPrompt() {}

    public static Set<String> normalizedSkillsKeys(List<List<String>> skillsRows) {
        LinkedHashSet<String> keys = new LinkedHashSet<>();
        for (String[] pm : MasterDispatchSheetEditRules.skillsEquipmentPairs(skillsRows)) {
            String k = MasterTeamCombinationTableReader.normalizedComboKey(pm[0], pm[1]);
            if (!k.isEmpty()) {
                keys.add(k);
            }
        }
        return Set.copyOf(keys);
    }

    public static PromptBundle collectMissingPairs(Map<String, String> ui) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path master = resolveMasterWorkbook(u);
        if (master == null || !Files.isRegularFile(master)) {
            return new PromptBundle(List.of());
        }
        var skillsKeys = SkillsSheetEquipmentListReader.readNormalizedComboKeys(master);
        return collectMissingAgainstSkillsKeys(u, skillsKeys);
    }

    /**
     * 計画タスクにあって {@code skillsKeys} に無い工程+機械。空のキー集合なら未登録なしとみなす。
     */
    public static PromptBundle collectMissingAgainstSkillsKeys(
            Map<String, String> ui, Set<String> skillsKeys) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        if (skillsKeys == null || skillsKeys.isEmpty()) {
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

        LinkedHashMap<String, MissingPair> missing = new LinkedHashMap<>();
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
            if (nk.isEmpty() || skillsKeys.contains(nk)) {
                continue;
            }
            missing.putIfAbsent(
                    nk,
                    new MissingPair(
                            proc.strip(),
                            mach.strip(),
                            iTask >= 0 ? cell(row, iTask) : ""));
        }
        return new PromptBundle(List.copyOf(missing.values()));
    }

    private static String cell(List<String> row, int col) {
        if (row == null || col < 0 || col >= row.size()) {
            return "";
        }
        String v = row.get(col);
        return v != null ? v.strip() : "";
    }

    static Path resolvePlanInputPath(Map<String, String> ui) {
        return Stage2UnknownMasterCombinationPrompt.resolvePlanInputPath(ui);
    }

    static Path resolveMasterWorkbook(Map<String, String> ui) {
        return Stage2UnknownMasterCombinationPrompt.resolveMasterWorkbook(ui);
    }
}
