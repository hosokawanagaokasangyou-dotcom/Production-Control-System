package jp.co.pm.ai.desktop.dispatch;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * 段階2が出力した {@code 結果_配台表.json} と、手動修正タブで表を構築した直後の {@link ResultDispatchDocument} が同一パイプラインで一致するか検証する。
 */
public final class ResultDispatchStage2TableJsonReconciliation {

    /** ダイアログに載せる差分行の上限（超過分はログで案内）。 */
    private static final int MAX_DETAIL_LINES = 40;

    private ResultDispatchStage2TableJsonReconciliation() {}

    public record Result(boolean ok, List<String> detailLines) {}

    /**
     * {@code jsonPath} を再読込し、手動修正タブと同じ列補完・マージ・正規化を適用した結果と {@code displayedAfterPipeline}
     * を比較する。
     */
    public static Result verifyAgainstDiskJson(
            ResultDispatchDocument displayedAfterPipeline, Path jsonPath) {
        List<String> lines = new ArrayList<>();
        if (displayedAfterPipeline == null) {
            lines.add("表示用ドキュメントが null のため比較できません。");
            return new Result(false, lines);
        }
        if (jsonPath == null) {
            lines.add("JSON パスが null のため比較できません。");
            return new Result(false, lines);
        }
        try {
            ResultDispatchDocument fresh = ResultDispatchJsonIo.read(jsonPath);
            ResultDispatchStage2ColumnSupport.ensureStage2RequiredColumns(fresh);
            ResultDispatchInteractiveGridModel.applyWideMergeAndNormalize(fresh);
            return compareDocuments(displayedAfterPipeline, fresh, lines);
        } catch (Exception e) {
            lines.add(
                    "JSON の再読込または比較で例外が発生しました: "
                            + (e.getMessage() != null ? e.getMessage() : e.getClass().getSimpleName()));
            return new Result(false, lines);
        }
    }

    static Result compareDocuments(
            ResultDispatchDocument a, ResultDispatchDocument b, List<String> sink) {
        List<String> colsA = a.columns();
        List<String> colsB = b.columns();
        if (!colsA.equals(colsB)) {
            sink.add(
                    "[列] 列ヘッダの並びが一致しません（A="
                            + colsA.size()
                            + "列 / B="
                            + colsB.size()
                            + "列）。先頭差分の例: "
                            + firstColumnDiffSummary(colsA, colsB));
            return new Result(false, trimDetail(sink));
        }
        List<Map<String, String>> ra = a.rows();
        List<Map<String, String>> rb = b.rows();
        if (ra.size() != rb.size()) {
            sink.add("[行数] 一致しません: 表側=" + ra.size() + " / JSON再処理=" + rb.size());
        }
        int n = Math.min(ra.size(), rb.size());
        for (int i = 0; i < n; i++) {
            Map<String, String> ma = ra.get(i);
            Map<String, String> mb = rb.get(i);
            for (String col : colsA) {
                String va = cell(ma, col);
                String vb = cell(mb, col);
                if (!Objects.equals(va, vb)) {
                    sink.add(
                            "[行"
                                    + (i + 1)
                                    + "]["
                                    + col
                                    + "] 表側="
                                    + abbrev(va)
                                    + " / JSON再処理="
                                    + abbrev(vb));
                    if (sink.size() >= MAX_DETAIL_LINES) {
                        return new Result(false, trimDetail(sink));
                    }
                }
            }
        }
        if (ra.size() != rb.size()) {
            return new Result(false, trimDetail(sink));
        }
        return new Result(sink.isEmpty(), trimDetail(sink));
    }

    private static List<String> trimDetail(List<String> sink) {
        if (sink.size() <= MAX_DETAIL_LINES) {
            return List.copyOf(sink);
        }
        List<String> head = new ArrayList<>(sink.subList(0, MAX_DETAIL_LINES));
        head.add("…（差分が多いためここで打ち切り。ログに全文を出す運用に拡張可能）");
        return head;
    }

    private static String cell(Map<String, String> row, String col) {
        if (row == null) {
            return "";
        }
        String v = row.get(col);
        return v != null ? v : "";
    }

    private static String abbrev(String s) {
        if (s == null) {
            return "（null）";
        }
        if (s.length() > 48) {
            return s.substring(0, 48) + "…";
        }
        return s.isEmpty() ? "（空）" : s;
    }

    private static String firstColumnDiffSummary(List<String> a, List<String> b) {
        int n = Math.min(a.size(), b.size());
        for (int i = 0; i < n; i++) {
            if (!Objects.equals(a.get(i), b.get(i))) {
                return "index " + i + ": 「" + abbrev(a.get(i)) + "」vs「" + abbrev(b.get(i)) + "」";
            }
        }
        if (a.size() != b.size()) {
            return "短い方の末尾まで一致";
        }
        return "（内容は同長）";
    }
}
