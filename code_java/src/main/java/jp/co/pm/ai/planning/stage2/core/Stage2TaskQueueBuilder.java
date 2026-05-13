package jp.co.pm.ai.planning.stage2.core;

import java.text.Normalizer;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;

/**
 * 計画入力タブular から依頼NO列を解決し、配台キュー候補のリストを構築する（Python task_queue の足場）。
 */
public final class Stage2TaskQueueBuilder {

    private static final String PRIMARY_REQUEST_HEADER = "依頼NO";

    private Stage2TaskQueueBuilder() {}

    public static List<Stage2QueuedTask> build(Stage2InputSnapshot snap) {
        PlanInputTabularIo.TabularSheet tab = snap.planningTasksSheet();
        List<String> headers = tab.headers();
        List<List<String>> rows = tab.rows();
        int col = findRequestIdColumn(headers);
        List<Stage2QueuedTask> out = new ArrayList<>();
        if (col < 0 || rows == null) {
            return out;
        }
        Set<String> seen = new LinkedHashSet<>();
        int excelRow = 2;
        for (List<String> row : rows) {
            String id = col < row.size() && row.get(col) != null ? row.get(col).strip() : "";
            if (!id.isEmpty() && seen.add(id)) {
                out.add(new Stage2QueuedTask(excelRow, id));
            }
            excelRow++;
        }
        return out;
    }

    /**
     * ゴールデン比較用: 依頼NO の安定した順序（出現順・重複除去後）。
     */
    public static List<String> requestIdsInOrder(List<Stage2QueuedTask> tasks) {
        return tasks.stream().map(Stage2QueuedTask::requestId).toList();
    }

    static int findRequestIdColumn(List<String> headers) {
        if (headers == null) {
            return -1;
        }
        for (int i = 0; i < headers.size(); i++) {
            String h = normalizeHeader(headers.get(i));
            if (PRIMARY_REQUEST_HEADER.equals(h)) {
                return i;
            }
        }
        for (int i = 0; i < headers.size(); i++) {
            String h = normalizeHeader(headers.get(i));
            if (h.contains("依頼") && h.contains("no")) {
                return i;
            }
        }
        return -1;
    }

    private static String normalizeHeader(String raw) {
        if (raw == null) {
            return "";
        }
        String n = Normalizer.normalize(raw.strip(), Normalizer.Form.NFKC);
        return n.toUpperCase(Locale.ROOT);
    }
}
