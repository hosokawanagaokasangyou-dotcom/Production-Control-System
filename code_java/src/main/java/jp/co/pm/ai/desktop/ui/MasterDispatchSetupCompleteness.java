package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.EnumSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.OptionalInt;
import java.util.Set;

import jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader;

/**
 * 配台マスタ設定ウィザード用: 計画タスクの工程+機械について
 * skills → need → 組み合わせ表 → 加工速度の完了判定。
 */
public final class MasterDispatchSetupCompleteness {

    public enum Step {
        SKILLS("資格（skills）"),
        NEED("必要人数（need）"),
        COMBINATIONS("組み合わせ表"),
        SPEED("加工速度（speed）");

        private final String labelJa;

        Step(String labelJa) {
            this.labelJa = labelJa;
        }

        public String labelJa() {
            return labelJa;
        }
    }

    public record EquipmentRef(String process, String machine, String sampleTaskId) {
        public String display() {
            return process + " × " + machine;
        }

        public String normalizedKey() {
            return MasterTeamCombinationTableReader.normalizedComboKey(
                    process != null ? process : "", machine != null ? machine : "");
        }
    }

    public record EquipmentStatus(EquipmentRef equipment, EnumSet<Step> incompleteSteps) {
        public boolean complete() {
            return incompleteSteps == null || incompleteSteps.isEmpty();
        }

        public Step firstIncomplete() {
            if (complete()) {
                return null;
            }
            for (Step s : Step.values()) {
                if (incompleteSteps.contains(s)) {
                    return s;
                }
            }
            return null;
        }
    }

    public record Evaluation(List<EquipmentStatus> statuses) {
        public boolean allComplete() {
            if (statuses == null || statuses.isEmpty()) {
                return true;
            }
            for (EquipmentStatus s : statuses) {
                if (!s.complete()) {
                    return false;
                }
            }
            return true;
        }

        public List<EquipmentStatus> incomplete() {
            List<EquipmentStatus> out = new ArrayList<>();
            if (statuses == null) {
                return List.of();
            }
            for (EquipmentStatus s : statuses) {
                if (!s.complete()) {
                    out.add(s);
                }
            }
            return List.copyOf(out);
        }

        public String summaryJa(int maxLines) {
            List<EquipmentStatus> bad = incomplete();
            if (bad.isEmpty()) {
                return "";
            }
            int limit = Math.max(1, maxLines);
            StringBuilder sb = new StringBuilder();
            int n = 0;
            for (EquipmentStatus s : bad) {
                if (n >= limit) {
                    sb.append("… 他 ").append(bad.size() - limit).append(" 件");
                    break;
                }
                if (n > 0) {
                    sb.append('\n');
                }
                sb.append("・").append(s.equipment().display()).append(" — 未完了: ");
                boolean first = true;
                for (Step step : Step.values()) {
                    if (s.incompleteSteps().contains(step)) {
                        if (!first) {
                            sb.append("、");
                        }
                        sb.append(step.labelJa());
                        first = false;
                    }
                }
                n++;
            }
            return sb.toString();
        }
    }

    private MasterDispatchSetupCompleteness() {}

    public static Evaluation evaluate(
            List<EquipmentRef> equipment,
            List<List<String>> skills,
            List<List<String>> need,
            List<List<String>> combinations,
            List<List<String>> speed) {
        List<EquipmentRef> eqs = equipment != null ? equipment : List.of();
        LinkedHashMap<String, EquipmentStatus> byKey = new LinkedHashMap<>();
        for (EquipmentRef eq : eqs) {
            if (eq == null) {
                continue;
            }
            String key = eq.normalizedKey();
            if (key.isEmpty() || byKey.containsKey(key)) {
                continue;
            }
            EnumSet<Step> incomplete = EnumSet.noneOf(Step.class);
            if (!skillsComplete(skills, eq.process(), eq.machine())) {
                incomplete.add(Step.SKILLS);
            }
            OptionalInt needK = readBaseRequiredHeadcount(need, eq.process(), eq.machine());
            if (needK.isEmpty() || needK.getAsInt() < 1) {
                incomplete.add(Step.NEED);
            }
            int k = needK.isPresent() && needK.getAsInt() >= 1 ? needK.getAsInt() : 1;
            if (!combinationsComplete(combinations, eq.process(), eq.machine(), k)) {
                incomplete.add(Step.COMBINATIONS);
            }
            if (!speedComplete(speed, eq.process(), eq.machine())) {
                incomplete.add(Step.SPEED);
            }
            byKey.put(key, new EquipmentStatus(eq, incomplete));
        }
        return new Evaluation(List.copyOf(byKey.values()));
    }

    public static boolean skillsComplete(List<List<String>> skills, String process, String machine) {
        return !skilledOpAsMembers(skills, process, machine).isEmpty();
    }

    /** 当該設備列で OP/AS が入っているメンバー表示（「OP 氏名」形式）。 */
    public static List<String> skilledOpAsMembers(
            List<List<String>> skills, String process, String machine) {
        return MasterDispatchSheetEditRules.skilledMembersForEquipment(skills, process, machine);
    }

    /**
     * need の基本必要人数。列が無い・空・非数・1 未満は empty。
     */
    public static OptionalInt readBaseRequiredHeadcount(
            List<List<String>> needRows, String process, String machine) {
        List<List<String>> need = needRows != null ? needRows : List.of();
        int col = findEquipmentColumn(MasterDispatchSheetEditRules.SheetKind.NEED, need, process, machine);
        if (col < 0) {
            return OptionalInt.empty();
        }
        int baseRow = findNeedBaseRow(need);
        if (baseRow < 0) {
            return OptionalInt.empty();
        }
        String raw = MasterDispatchSheetEditRules.cell(need, baseRow, col);
        if (raw == null || raw.isBlank()) {
            return OptionalInt.empty();
        }
        try {
            int n = Integer.parseInt(stripTrailingDotZero(raw));
            return n >= 1 ? OptionalInt.of(n) : OptionalInt.empty();
        } catch (NumberFormatException e) {
            return OptionalInt.empty();
        }
    }

    public static boolean combinationsComplete(
            List<List<String>> comboRows, String process, String machine, int requiredHeadcount) {
        int k = Math.max(1, requiredHeadcount);
        List<List<String>> rows = comboRows != null ? comboRows : List.of();
        if (rows.size() < 2) {
            return false;
        }
        List<String> header = rows.get(0);
        int procCol = MasterDispatchSheetEditRules.headerIndex(header, "工程名");
        int machCol = MasterDispatchSheetEditRules.headerIndex(header, "機械名");
        int comboCol = MasterDispatchSheetEditRules.headerIndex(header, "工程+機械", "工程＋機械");
        List<Integer> memberCols = new ArrayList<>();
        for (int c = 0; c < header.size(); c++) {
            if (MasterDispatchSheetEditRules.isCombinationMemberColumn(header, c)) {
                memberCols.add(c);
            }
        }
        if (memberCols.size() < k) {
            return false;
        }
        String want =
                MasterTeamCombinationTableReader.normalizedComboKey(
                        process != null ? process : "", machine != null ? machine : "");
        if (want.isEmpty()) {
            return false;
        }
        for (int r = 1; r < rows.size(); r++) {
            if (MasterDispatchSheetEditRules.isColumnTitleSourceRow(
                    MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, r, rows)) {
                continue;
            }
            String proc = procCol >= 0 ? MasterDispatchSheetEditRules.cell(rows, r, procCol) : "";
            String mach = machCol >= 0 ? MasterDispatchSheetEditRules.cell(rows, r, machCol) : "";
            String combo = comboCol >= 0 ? MasterDispatchSheetEditRules.cell(rows, r, comboCol) : "";
            String have = MasterTeamCombinationTableReader.comboKeyFromCells(proc, mach, combo);
            if (!want.equals(have)) {
                continue;
            }
            int filled = 0;
            for (int i = 0; i < k && i < memberCols.size(); i++) {
                String m = MasterDispatchSheetEditRules.cell(rows, r, memberCols.get(i));
                if (m != null && !m.isBlank()) {
                    filled++;
                }
            }
            if (filled >= k) {
                return true;
            }
        }
        return false;
    }

    public static boolean speedComplete(List<List<String>> speedRows, String process, String machine) {
        List<List<String>> speed = speedRows != null ? speedRows : List.of();
        int col = findEquipmentColumn(MasterDispatchSheetEditRules.SheetKind.SPEED, speed, process, machine);
        if (col < 0) {
            return false;
        }
        int baseRow = findSpeedBaseRow(speed);
        if (baseRow < 0) {
            return false;
        }
        String raw = MasterDispatchSheetEditRules.cell(speed, baseRow, col);
        if (raw == null || raw.isBlank()) {
            return false;
        }
        return MasterDispatchSheetEditRules.isSpeedBaseDecimalValid(raw);
    }

    static int findEquipmentColumn(
            MasterDispatchSheetEditRules.SheetKind kind,
            List<List<String>> rows,
            String process,
            String machine) {
        if (!MasterDispatchSheetEditRules.containsEquipmentColumn(kind, rows, process, machine)) {
            return -1;
        }
        List<List<String>> src = rows != null ? rows : List.of();
        int procRow = findProcessHeaderRowPublic(src);
        int machRow = findMachineHeaderRowPublic(src);
        if (procRow < 0 || machRow < 0) {
            return -1;
        }
        int firstEq = kind == MasterDispatchSheetEditRules.SheetKind.NEED
                        || kind == MasterDispatchSheetEditRules.SheetKind.SPEED
                ? 3
                : 1;
        String want =
                MasterTeamCombinationTableReader.normalizedComboKey(
                        process != null ? process : "", machine != null ? machine : "");
        int width = 0;
        for (List<String> row : src) {
            if (row != null) {
                width = Math.max(width, row.size());
            }
        }
        for (int c = firstEq; c < width; c++) {
            String have =
                    MasterTeamCombinationTableReader.normalizedComboKey(
                            MasterDispatchSheetEditRules.cell(src, procRow, c),
                            MasterDispatchSheetEditRules.cell(src, machRow, c));
            if (want.equals(have)) {
                return c;
            }
        }
        return -1;
    }

    private static int findNeedBaseRow(List<List<String>> need) {
        for (int r = 0; r < need.size(); r++) {
            String a = MasterDispatchSheetEditRules.cell(need, r, 0);
            if (a.contains("必須人数") || a.contains("必要人数")) {
                if (a.contains("追加人数") || a.contains("余剰")) {
                    continue;
                }
                return r;
            }
        }
        return -1;
    }

    private static int findSpeedBaseRow(List<List<String>> speed) {
        for (int r = 0; r < speed.size(); r++) {
            String a = MasterDispatchSheetEditRules.cell(speed, r, 0);
            if (a.contains("基本速度")) {
                return r;
            }
        }
        return -1;
    }

    private static int findProcessHeaderRowPublic(List<List<String>> rows) {
        for (int r = 0; r < rows.size(); r++) {
            if ("工程名".equals(MasterDispatchSheetEditRules.cell(rows, r, 0))) {
                return r;
            }
        }
        return -1;
    }

    private static int findMachineHeaderRowPublic(List<List<String>> rows) {
        for (int r = 0; r < rows.size(); r++) {
            if ("機械名".equals(MasterDispatchSheetEditRules.cell(rows, r, 0))) {
                return r;
            }
        }
        return -1;
    }

    private static String stripTrailingDotZero(String s) {
        String t = s != null ? s.strip() : "";
        if (t.endsWith(".0")) {
            return t.substring(0, t.length() - 2);
        }
        return t;
    }

    /** 未使用抑制（将来のステップ集合比較用）。 */
    static Set<Step> allSteps() {
        return EnumSet.allOf(Step.class);
    }
}
