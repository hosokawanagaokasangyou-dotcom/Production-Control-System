package jp.co.pm.ai.kouchin.verify;

/**
 * 検証B（② vs ③ 依頼NO突合）の1行。
 *
 * @param irai    依頼NO
 * @param amount2 ②金額
 * @param amount3 ③アラジン金額
 * @param diff    差額(②-③)
 * @param judge   判定
 * @param note    備考
 */
public record RecordB(
        String irai,
        Double amount2,
        Double amount3,
        Double diff,
        String judge,
        String note) {

    /** 並び替え用のスケール（金額の絶対値の最大）。 */
    public double scale() {
        double s = 0.0;
        if (amount2 != null) {
            s = Math.max(s, Math.abs(amount2));
        }
        if (amount3 != null) {
            s = Math.max(s, Math.abs(amount3));
        }
        return s;
    }
}
