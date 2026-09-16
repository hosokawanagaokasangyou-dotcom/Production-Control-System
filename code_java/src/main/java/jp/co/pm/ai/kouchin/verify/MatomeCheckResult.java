package jp.co.pm.ai.kouchin.verify;

import java.util.List;

/**
 * 検証D（国分②「東レまとめ」の内部整合性）の結果。
 *
 * @param rows       箇所別の明細
 * @param totals     元シート別の合計比較
 * @param mappedRows まとめが元シートを参照している行数
 * @param skipped    スキップ理由（実施できたときは null）
 */
public record MatomeCheckResult(
        List<MatomeRow> rows,
        List<SheetTotal> totals,
        int mappedRows,
        String skipped) {

    /**
     * @param place     場所（{@code 東レまとめ!225 ↔ 東レV.C!83} など）
     * @param irai      依頼NO
     * @param keiyaku   契約NO
     * @param matomeAa  まとめAA
     * @param srcAa     元シートAA
     * @param diff      差額（まとめ-元）
     * @param judge     判定（金額差 / 参照ずれ / 未取込 / 式異常 / 直接入力）
     * @param detail    内容・対処
     */
    public record MatomeRow(
            String place,
            String irai,
            String keiyaku,
            Double matomeAa,
            Double srcAa,
            Double diff,
            String judge,
            String detail) {
    }

    /** 元シート別のAA合計比較。 */
    public record SheetTotal(String sheet, double matomeSum, double srcSum, double diff) {
    }

    public static MatomeCheckResult skipped(String reason) {
        return new MatomeCheckResult(List.of(), List.of(), 0, reason);
    }

    public boolean isSkipped() {
        return skipped != null;
    }

    /** 要修正（金額差・参照ずれ・未取込）の件数。 */
    public int errorCount() {
        return (int) rows.stream().filter(r -> Judge.D_ERRORS.contains(r.judge())).count();
    }

    /** 注意（式異常・直接入力）の件数。 */
    public int noticeCount() {
        return rows.size() - errorCount();
    }

    /** 元シート別合計の差の総和。 */
    public double totalGap() {
        return totals.stream().mapToDouble(SheetTotal::diff).sum();
    }
}
