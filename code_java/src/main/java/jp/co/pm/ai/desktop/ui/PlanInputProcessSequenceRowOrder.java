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
 */
public final class PlanInputProcessSequenceRowOrder {

    public static final String COL_DISPATCH_TRIAL_ORDER = "配台試行順番";
    public static final String COL_TASK_ID = "依頼NO";
    public static final String COL_PROCESS = "工程名";
    public static final String COL_PROCESS_CONTENT = "加工内容";

    private static final Pattern WS_COLLAPSE = Pattern.compile("[\\s　]+");
    private static final int RANK_MISSING = 1_000_000_000;
    private static final int DTO_MISSING = 1_000_000_000;

    private PlanInputProcessSequenceRowOrder() {}

    /**
     * 現在の行順を §A-1 に沿って整え、{@link #COL_DISPATCH_TRIAL_ORDER} を 1..n で振り直す。
     *
     * <p>並べ替えキー: 依頼NOブロックの最小試行順 → 加工内容内の工程 rank → 同一依頼内の出現順 → 元行位置。
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

        int n = rows.size();
        Map<String, List<String>> tokensByTaskId = collectProcessContentTokensByTaskId(rows, colTask, colContent);
        Map<String, Double> blockDtoByTaskId = new HashMap<>();
        Map<String, Integer> nextLineSeq = new HashMap<>();

        List<RowMeta> metas = new ArrayList<>(n);
        for (int i = 0; i < n; i++) {
            ObservableList<String> row = rows.get(i);
            String taskId = colTask >= 0 ? cellAt(row, colTask) : "";
            Double dto = parseTrialOrderSortKey(cellAt(row, colDto));
            List<String> tokens = tokensByTaskId.getOrDefault(taskId, List.of());
            Integer rank =
                    colProc >= 0 && !tokens.isEmpty()
                            ? processSequenceRank(cellAt(row, colProc), tokens)
                            : null;
            int lineSeq;
            if (!taskId.isEmpty()) {
                lineSeq = nextLineSeq.getOrDefault(taskId, 0);
                nextLineSeq.put(taskId, lineSeq + 1);
                if (dto != null) {
                    blockDtoByTaskId.merge(taskId, dto, Double::min);
                }
            } else {
                lineSeq = i;
            }
            metas.add(new RowMeta(i, taskId, dto, rank, lineSeq));
        }

        metas.sort(
                Comparator.<RowMeta>comparingInt(m -> m.dto() == null ? 1 : 0)
                        .thenComparingDouble(
                                m -> {
                                    if (m.dto() == null) {
                                        return DTO_MISSING;
                                    }
                                    if (!m.taskId().isEmpty()) {
                                        Double block = blockDtoByTaskId.get(m.taskId());
                                        if (block != null) {
                                            return block;
                                        }
                                    }
                                    return m.dto();
                                })
                        .thenComparingInt(m -> m.rank() != null ? m.rank() : RANK_MISSING)
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
            int originalIndex, String taskId, Double dto, Integer rank, int lineSeq) {}
}
