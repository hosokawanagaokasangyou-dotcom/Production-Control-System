package jp.co.pm.ai.desktop.ui;

import java.text.Normalizer;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import javafx.collections.ObservableList;

/**
 * 配台計画_タスク入力／3.0: 行並べ替え後も同一依頼NO内の工程順（§A-1・加工内容のカンマ区切り）を維持する。
 *
 * <p>「配台不要」オン行はブロック集約・工程 rank の対象外（単独行の試行順で並ぶ）。
 */
public final class PlanInputProcessSequenceRowOrder {

    public static final String COL_DISPATCH_TRIAL_ORDER = "配台試行順番";
    public static final String COL_TASK_ID = "依頼NO";
    public static final String COL_PROCESS = "工程名";
    public static final String COL_PROCESS_CONTENT = "加工内容";
    public static final String COL_EXCLUDE_FROM_ASSIGNMENT = "配台不要";

    private static final Pattern WS_COLLAPSE = Pattern.compile("[\\s　]+");
    private static final int RANK_MISSING = 1_000_000_000;
    private static final int DTO_MISSING = 1_000_000_000;

    private PlanInputProcessSequenceRowOrder() {}

    /**
     * ユーザー操作（DnD・↑↓）による行移動。配台不要オフ行をドラッグしたときは、同一依頼NOの配台対象行を
     * まとめて移す（配台不要=yes 行は追従しない）。
     */
    public static void moveRowsForUserReorder(
            List<String> headers,
            ObservableList<ObservableList<String>> rows,
            int from,
            int to) {
        if (headers == null || rows == null || from == to || from < 0 || to < 0) {
            return;
        }
        if (from >= rows.size() || to >= rows.size()) {
            return;
        }
        int colTask = headers.indexOf(COL_TASK_ID);
        int colExclude = headers.indexOf(COL_EXCLUDE_FROM_ASSIGNMENT);
        if (isRowExcludedFromAssignment(rows.get(from), colExclude)) {
            moveSingleRow(rows, from, to);
            return;
        }
        String taskId = colTask >= 0 ? cellAt(rows.get(from), colTask) : "";
        if (taskId.isEmpty()) {
            moveSingleRow(rows, from, to);
            return;
        }
        List<Integer> eligibleIndices = eligibleRowIndices(rows, taskId, colTask, colExclude);
        if (eligibleIndices.size() <= 1 || !eligibleIndices.contains(from)) {
            moveSingleRow(rows, from, to);
            return;
        }
        moveEligibleBlock(rows, eligibleIndices, from, to);
    }

    /**
     * 現在の行順を §A-1 に沿って整え、{@link #COL_DISPATCH_TRIAL_ORDER} を 1..n で振り直す。
     */
    public static void stabilizeAndRenumberDispatchTrialOrder(
            List<String> headers, ObservableList<ObservableList<String>> rows) {
        if (headers == null || rows == null || rows.isEmpty()) {
            return;
        }
        int colDto = headers.indexOf(COL_DISPATCH_TRIAL_ORDER);
        int colTask = headers.indexOf(COL_TASK_ID);
        int colProc = headers.indexOf(COL_PROCESS);
        int colContent = headers.indexOf(COL_PROCESS_CONTENT);
        if (colDto < 0) {
            return;
        }

        int colExclude = headers.indexOf(COL_EXCLUDE_FROM_ASSIGNMENT);

        int n = rows.size();
        seedTrialOrderKeysFromCurrentRowOrder(rows, colDto, n);
        Map<String, List<String>> tokensByTaskId = collectProcessContentTokensByTaskId(rows, colTask, colContent);
        Map<String, Double> eligibleBlockDtoByTaskId = new HashMap<>();
        Map<String, Integer> nextEligibleLineSeq = new HashMap<>();

        List<RowMeta> metas = new ArrayList<>(n);
        for (int i = 0; i < n; i++) {
            ObservableList<String> row = rows.get(i);
            String taskId = colTask >= 0 ? cellAt(row, colTask) : "";
            boolean excluded =
                    colExclude >= 0
                            && TabularCellHighlight.planInputExcludeFromAssignmentIsOn(
                                    cellAt(row, colExclude));
            Double dto = parseTrialOrderSortKey(cellAt(row, colDto));
            List<String> tokens = tokensByTaskId.getOrDefault(taskId, List.of());
            Integer rank = null;
            if (!excluded && colProc >= 0 && !tokens.isEmpty()) {
                rank = processSequenceRank(cellAt(row, colProc), tokens);
            }
            int lineSeq;
            if (!taskId.isEmpty() && !excluded) {
                lineSeq = nextEligibleLineSeq.getOrDefault(taskId, 0);
                nextEligibleLineSeq.put(taskId, lineSeq + 1);
                if (dto != null) {
                    eligibleBlockDtoByTaskId.merge(taskId, dto, Double::min);
                }
            } else {
                lineSeq = i;
            }
            metas.add(new RowMeta(i, taskId, dto, rank, lineSeq, excluded));
        }

        metas.sort(
                Comparator.<RowMeta>comparingInt(m -> m.dto() == null ? 1 : 0)
                        .thenComparingDouble(m -> sortBlockKey(m, eligibleBlockDtoByTaskId))
                        .thenComparingInt(m -> m.excluded() ? 1 : 0)
                        .thenComparingInt(
                                m ->
                                        m.excluded()
                                                ? m.originalIndex()
                                                : (m.rank() != null ? m.rank() : RANK_MISSING))
                        .thenComparingInt(RowMeta::lineSeq)
                        .thenComparingInt(RowMeta::originalIndex));

        List<ObservableList<String>> reordered = new ArrayList<>(n);
        for (RowMeta m : metas) {
            reordered.add(rows.get(m.originalIndex()));
        }
        rows.setAll(reordered);

        for (int i = 0; i < rows.size(); i++) {
            ObservableList<String> row = rows.get(i);
            ensureSize(row, colDto + 1);
            row.set(colDto, Integer.toString(i + 1));
        }
    }

    /**
     * 並べ替え後の行順を試行順キーに反映する。以降のブロック集約はこの位置を基準にする（旧キー最小値へ戻さない）。
     */
    private static boolean isRowExcludedFromAssignment(
            ObservableList<String> row, int colExclude) {
        return colExclude >= 0
                && TabularCellHighlight.planInputExcludeFromAssignmentIsOn(cellAt(row, colExclude));
    }

    private static List<Integer> eligibleRowIndices(
            ObservableList<ObservableList<String>> rows,
            String taskId,
            int colTask,
            int colExclude) {
        List<Integer> out = new ArrayList<>();
        for (int i = 0; i < rows.size(); i++) {
            ObservableList<String> row = rows.get(i);
            if (!taskId.equals(cellAt(row, colTask))) {
                continue;
            }
            if (!isRowExcludedFromAssignment(row, colExclude)) {
                out.add(i);
            }
        }
        return out;
    }

    private static void moveSingleRow(
            ObservableList<ObservableList<String>> rows, int from, int to) {
        ObservableList<String> moved = rows.remove(from);
        rows.add(to, moved);
    }

    private static void moveEligibleBlock(
            ObservableList<ObservableList<String>> rows,
            List<Integer> blockIndices,
            int from,
            int to) {
        List<ObservableList<String>> block = new ArrayList<>(blockIndices.size());
        for (int idx : blockIndices) {
            block.add(rows.get(idx));
        }
        for (int i = blockIndices.size() - 1; i >= 0; i--) {
            rows.remove((int) blockIndices.get(i));
        }
        int insertAt = to;
        if (from < to) {
            insertAt = to - blockIndices.size() + 1;
        }
        rows.addAll(insertAt, block);
    }

    private static void seedTrialOrderKeysFromCurrentRowOrder(
            ObservableList<ObservableList<String>> rows, int colDto, int n) {
        for (int i = 0; i < n; i++) {
            ObservableList<String> row = rows.get(i);
            ensureSize(row, colDto + 1);
            row.set(colDto, Integer.toString(i + 1));
        }
    }

    private static double sortBlockKey(RowMeta m, Map<String, Double> eligibleBlockDtoByTaskId) {
        if (m.dto() == null) {
            return DTO_MISSING;
        }
        if (m.excluded()) {
            return m.dto();
        }
        if (!m.taskId().isEmpty()) {
            Double block = eligibleBlockDtoByTaskId.get(m.taskId());
            if (block != null) {
                return block;
            }
        }
        return m.dto();
    }

    private static Map<String, List<String>> collectProcessContentTokensByTaskId(
            ObservableList<ObservableList<String>> rows, int colTask, int colContent) {
        Map<String, List<String>> out = new LinkedHashMap<>();
        if (colTask < 0 || colContent < 0) {
            return out;
        }
        for (ObservableList<String> row : rows) {
            String taskId = cellAt(row, colTask);
            if (taskId.isEmpty() || out.containsKey(taskId)) {
                continue;
            }
            List<String> tokens = parseProcessContentTokens(cellAt(row, colContent));
            if (!tokens.isEmpty()) {
                out.put(taskId, tokens);
            }
        }
        return out;
    }

    static List<String> parseProcessContentTokens(String raw) {
        if (raw == null || raw.isBlank()) {
            return List.of();
        }
        String normalized = raw.replace('、', ',');
        List<String> out = new ArrayList<>();
        for (String part : normalized.split(",")) {
            String t = part.strip();
            if (!t.isEmpty()) {
                out.add(t);
            }
        }
        return out;
    }

    static Integer processSequenceRank(String processName, List<String> contentTokens) {
        if (contentTokens == null || contentTokens.isEmpty()) {
            return null;
        }
        String proc = normalizeProcessName(processName);
        if (proc.isEmpty()) {
            return null;
        }
        for (int i = 0; i < contentTokens.size(); i++) {
            if (proc.equals(normalizeProcessName(contentTokens.get(i)))) {
                return i;
            }
        }
        return null;
    }

    static String normalizeProcessName(String raw) {
        if (raw == null) {
            return "";
        }
        String t = Normalizer.normalize(raw.strip(), Normalizer.Form.NFKC);
        return WS_COLLAPSE.matcher(t).replaceAll("");
    }

    private static String cellAt(ObservableList<String> row, int colIndex) {
        if (row == null || colIndex < 0 || colIndex >= row.size()) {
            return "";
        }
        String v = row.get(colIndex);
        return v != null ? v.strip() : "";
    }

    private static void ensureSize(ObservableList<String> row, int size) {
        while (row.size() < size) {
            row.add("");
        }
    }

    private static Double parseTrialOrderSortKey(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        try {
            double d = Double.parseDouble(raw.strip());
            if (!Double.isFinite(d)) {
                return null;
            }
            return d;
        } catch (NumberFormatException ex) {
            return null;
        }
    }

    private record RowMeta(
            int originalIndex,
            String taskId,
            Double dto,
            Integer rank,
            int lineSeq,
            boolean excluded) {}
}
