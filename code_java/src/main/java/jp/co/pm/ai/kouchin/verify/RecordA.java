package jp.co.pm.ai.kouchin.verify;

/**
 * 検証A（① vs ② 契約NO突合）の1行。
 *
 * @param keiyaku       契約NO
 * @param iraiNo        依頼NO（前月調整・翌月記載は月ラベル付き）
 * @param amount1       ①東レ金額
 * @param amount2       ②金額
 * @param diff          差額(①-②)
 * @param reportAmount  報告計上額（列合計＝報告する過不足）
 * @param judge         判定
 * @param note          備考
 */
public record RecordA(
        String keiyaku,
        String iraiNo,
        Double amount1,
        Double amount2,
        Double diff,
        Double reportAmount,
        String judge,
        String note) {

    /** 並び替え用のスケール（金額の絶対値の最大）。 */
    public double scale() {
        double s = 0.0;
        if (amount1 != null) {
            s = Math.max(s, Math.abs(amount1));
        }
        if (amount2 != null) {
            s = Math.max(s, Math.abs(amount2));
        }
        return s;
    }
}
