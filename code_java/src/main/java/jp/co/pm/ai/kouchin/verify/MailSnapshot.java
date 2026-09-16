package jp.co.pm.ai.kouchin.verify;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

/**
 * 東レ宛 月次報告メールに反映する工場別の数字。
 * 両工場を1通にまとめるため JSON で保存し、他工場の直近結果として読み戻す。
 *
 * @param factoryCode    工場コード（kokubu / konan）
 * @param factoryLabel   工場名
 * @param ymLabel        対象月ラベル（2026年7月度）
 * @param nextYmLabel    翌月ラベル（2026年8月度）
 * @param total1         ① 東レお支払データ総額
 * @param uriage2        ② 当月実売上（契約NO合計）
 * @param adjustPrev     前月調整分
 * @param tougetsu       報告する過不足（当月差異＋翌月記載＋①のみ−②のみ）
 * @param tougetsuCount  報告する過不足の件数
 * @param mismatchCount  金額不一致の件数
 * @param mismatchAmount 金額不一致の差額合計
 * @param nextCount      翌月記載の件数
 * @param nextAmount     翌月記載の金額
 * @param only1Count     ①のみの件数
 * @param only1Amount    ①のみの金額
 * @param only2Count     ②のみの件数
 * @param only2Amount    ②のみの金額
 * @param excludedBasho  対象外入庫場所の表示文字列
 * @param generatedAt    生成日時
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record MailSnapshot(
        String factoryCode,
        String factoryLabel,
        String ymLabel,
        String nextYmLabel,
        long total1,
        long uriage2,
        long adjustPrev,
        long tougetsu,
        int tougetsuCount,
        int mismatchCount,
        long mismatchAmount,
        int nextCount,
        long nextAmount,
        int only1Count,
        long only1Amount,
        int only2Count,
        long only2Amount,
        String excludedBasho,
        String generatedAt) {

    /** メール本文1行分（{@code 国分工場　1,234円（実売上：…）-（前月調整分：…）-（当月：…）}）。 */
    public String amountLine() {
        return factoryLabel + "　" + Fmt.n0(total1) + "円（実売上：" + Fmt.n0(uriage2)
                + "）-（前月調整分：" + Fmt.n0(-adjustPrev) + "）-（当月：" + Fmt.n0(-tougetsu) + "）";
    }

    /** 差異件数の行（{@code 国分工場　3件（-12,345円）の差異（過不足）がありました。}）。 */
    public String diffLine() {
        return factoryLabel + "　" + tougetsuCount + "件（" + Fmt.n0(tougetsu) + "円）の差異（過不足）がありました。";
    }
}
