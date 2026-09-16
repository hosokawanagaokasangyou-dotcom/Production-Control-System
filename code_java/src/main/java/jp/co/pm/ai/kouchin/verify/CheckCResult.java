package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Path;
import java.util.List;

/**
 * 検証C（湖南・月次処理ファイルの東レ合計照合）の結果。
 *
 * @param file    使用した月次処理ファイル（無ければ null）
 * @param skipped スキップ理由（実施できたときは null）
 * @param rows    4項目の照合行
 */
public record CheckCResult(Path file, String skipped, List<Row> rows) {

    public static final String MATCH = "一致";
    public static final String EXPLAINABLE = "内訳で説明可";
    public static final String MISMATCH = "不一致";
    public static final String NEED_CHECK = "要確認";
    public static final String UNREADABLE = "読取不可";

    public record Row(
            String item, Double monthlyValue, Double ours, Double diff, String judge, String note) {}

    public static CheckCResult skipped(String reason) {
        return new CheckCResult(null, reason, List.of());
    }

    public boolean isSkipped() {
        return skipped != null;
    }

    /** KPI の要確認件数。「内訳で説明可」は入れない。 */
    public int needCheckCount() {
        return (int)
                rows.stream()
                        .filter(r -> MISMATCH.equals(r.judge())
                                || NEED_CHECK.equals(r.judge())
                                || UNREADABLE.equals(r.judge()))
                        .count();
    }
}
