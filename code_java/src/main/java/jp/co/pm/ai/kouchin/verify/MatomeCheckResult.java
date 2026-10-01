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
     * @param place        場所（{@code 東レまとめ!225 ↔ 東レV.C!83} など）
     * @param irai         依頼NO
     * @param keiyaku      契約NO
     * @param torayAmount  ①東レ金額（当該契約が①に無い検査行は null）
     * @param matomeAa     まとめAA
     * @param srcAa        元シートAA
     * @param diff         ①差額行は ①−まとめ。金額差で①が無い行は まとめ−元
     * @param judge        判定（①差額 / 金額差 / 参照ずれ / 未取込 / 式異常 / 直接入力）
     * @param detail       内容・対処
     */
    public record MatomeRow(
            String place,
            String irai,
            String keiyaku,
            Double torayAmount,
            Double matomeAa,
            Double srcAa,
            Double diff,
            String judge,
            String detail) {
    }

    /** 元シート別のAA合計比較。 */
    public record SheetTotal(String sheet, double matomeSum, double srcSum, double diff) {
    }

    /** 内部検査行。①金額は後から契約NOで埋める。 */
    public static MatomeRow internal(
            String place, String irai, String keiyaku,
            Double matomeAa, Double srcAa, Double diff, String judge, String detail) {
        return new MatomeRow(place, irai, keiyaku, null, matomeAa, srcAa, diff, judge, detail);
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

    /** ①東レ金額とまとめAAが違う契約の件数。要修正には含めない。 */
    public int torayDiffCount() {
        return (int) rows.stream().filter(r -> Judge.TORAY_DIFF.equals(r.judge())).count();
    }

    /** 注意（式異常・直接入力）の件数。 */
    public int noticeCount() {
        return (int) rows.stream()
                .filter(r -> Judge.BAD_FORMULA.equals(r.judge()) || Judge.HARDCODED.equals(r.judge()))
                .count();
    }

    public int countOf(String judge) {
        return (int) rows.stream().filter(r -> judge != null && judge.equals(r.judge())).count();
    }

    /** 元シート別合計の差の総和。 */
    public double totalGap() {
        return totals.stream().mapToDouble(SheetTotal::diff).sum();
    }
}
