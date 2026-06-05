package jp.co.pm.ai.desktop.dispatch;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo;

import jp.co.pm.ai.desktop.ui.TabularCellHighlight;

/**
 * 配台計画_タスク入力（配台不要でない行）と配台計画手動修正（結果_配台表 JSON）の整合確認。
 */
public final class DispatchPlanInputInteractiveCoverageCheck {

    private static final int MAX_MISSING_LINES_IN_DIALOG = 40;

    private DispatchPlanInputInteractiveCoverageCheck() {}

    /** (依頼NO, 工程名, 機械名) — 手動修正ワイド表と同一視するキー。 */
    public record TaskKey(String requestNo, String process, String machineName) {

        public TaskKey {
            requestNo = requestNo != null ? requestNo.strip() : "";
            process = process != null ? process.strip() : "";
            machineName = machineName != null ? machineName.strip() : "";
        }

        public boolean isComplete() {
            return !requestNo.isEmpty() && !process.isEmpty() && !machineName.isEmpty();
        }

        public String displayLine() {
            return requestNo + " / " + process + " / " + machineName;
        }

        public String identityToken() {
            return requestNo + "\u0001" + process + "\u0001" + machineName;
        }
    }

    /**
     * 配台計画手動修正へ載るべき計画入力行から外すか。
     *
     * <p>Python {@code build_task_queue_from_planning_df} の除外（配台不要・配台計画除外・完了）と
     * {@link TabularCellHighlight#planInputExcludeFromAssignmentIsOn} を Java 側で揃える。
     */
    public static boolean isExcludedFromDispatchCoverage(Map<String, String> row) {
        if (row == null || row.isEmpty()) {
            return true;
        }
        if (TabularCellHighlight.planInputExcludeFromAssignmentIsOn(
                row.getOrDefault("配台不要", ""))) {
            return true;
        }
        String excludeCell = row.getOrDefault("配台不要", "");
        if (excludeCell != null && excludeCell.contains("配台計画除外")) {
            return true;
        }
        String status = row.getOrDefault("ステータス", "");
        if (status != null) {
            String s = status.strip();
            if (s.contains("完了")
                    || s.equalsIgnoreCase("done")
                    || s.equalsIgnoreCase("complete")) {
                return true;
            }
        }
        return false;
    }

    public static boolean isEligiblePlanInputRow(Map<String, String> row) {
        return !isExcludedFromDispatchCoverage(row);
    }

    /**
     * 加工途中・翌日配台ダイアログで 0 m と明示された行を整合確認の期待リストから外す。
     *
     * <p>段階2配台対象から外す意図（ダイアログ文言）に合わせ、手動修正表への必須載せ要件からも除外する。
     */
    public static List<TaskKey> excludeInProgressZeroNextDayFromExpected(
            List<TaskKey> expected, Map<String, String> ui) {
        if (expected == null || expected.isEmpty()) {
            return List.of();
        }
        Set<String> zeroKeys = Stage2InProgressNextDayDispatchIo.zeroNextDayRowKeys(ui);
        if (zeroKeys.isEmpty()) {
            return expected;
        }
        List<TaskKey> filtered = new ArrayList<>(expected.size());
        for (TaskKey key : expected) {
            if (!key.isComplete()) {
                continue;
            }
            String rowKey =
                    Stage2InProgressNextDayDispatchIo.rowKey(
                            key.requestNo(), key.process(), key.machineName());
            if (zeroKeys.contains(rowKey)) {
                continue;
            }
            filtered.add(key);
        }
        return List.copyOf(filtered);
    }

    /** 計画入力の期待キー（出現順・重複除去）のうち、配台表に無いもの。 */
    public static List<TaskKey> findMissingInDispatchTable(
            List<TaskKey> expectedFromPlanInput, List<Map<String, String>> dispatchRows) {
        if (expectedFromPlanInput == null || expectedFromPlanInput.isEmpty()) {
            return List.of();
        }
        Set<String> present = new LinkedHashSet<>();
        if (dispatchRows != null) {
            for (Map<String, String> row : dispatchRows) {
                if (row == null) {
                    continue;
                }
                TaskKey k = taskKeyFromDispatchRow(row);
                if (k.isComplete()) {
                    present.add(k.identityToken());
                }
            }
        }
        List<TaskKey> missing = new ArrayList<>();
        Set<String> seenMissing = new LinkedHashSet<>();
        for (TaskKey key : expectedFromPlanInput) {
            if (!key.isComplete()) {
                continue;
            }
            if (present.contains(key.identityToken())) {
                continue;
            }
            if (seenMissing.add(key.identityToken())) {
                missing.add(key);
            }
        }
        return List.copyOf(missing);
    }

    public static TaskKey taskKeyFromDispatchRow(Map<String, String> row) {
        String id = nz(row.get("依頼NO"));
        if (id.isEmpty()) {
            id = nz(row.get("タスクID"));
        }
        return new TaskKey(id, nz(row.get("工程名")), nz(row.get("機械名")));
    }

    public static String formatMissingTasksDialogMessage(
            List<TaskKey> missing, String resultDispatchJsonPath) {
        StringBuilder sb = new StringBuilder();
        sb.append(
                "配台計画_タスク入力の表で「配台不要」がオフのタスクのうち、"
                        + "次の行が配台計画手動修正表（結果_配台表.json）にありません。\n\n");
        sb.append("手動修正表は段階2の設備タイムライン由来の行だけではなく、"
                + "未配台・加工途中（翌日配台ダイアログ）分も載る想定です。"
                + " 反映漏れのときは段階2ログと JSON 出力先も確認してください。\n\n");
        int show = Math.min(missing.size(), MAX_MISSING_LINES_IN_DIALOG);
        for (int i = 0; i < show; i++) {
            sb.append("・").append(missing.get(i).displayLine()).append('\n');
        }
        if (missing.size() > show) {
            sb.append("・… 他 ").append(missing.size() - show).append(" 件\n");
        }
        if (resultDispatchJsonPath != null && !resultDispatchJsonPath.isBlank()) {
            sb.append("\n結果_配台表.json:\n").append(resultDispatchJsonPath.strip());
        }
        return sb.toString();
    }

    private static String nz(String s) {
        return s == null ? "" : s.strip();
    }
}
