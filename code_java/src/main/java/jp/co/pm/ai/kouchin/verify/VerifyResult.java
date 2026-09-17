package jp.co.pm.ai.kouchin.verify;

import java.util.List;
import java.util.Map;

/**
 * 1工場分の検証結果。
 *
 * @param profile  工場プロファイル
 * @param targetYm 対象年月（①のファイル名基準・不明なら null）
 * @param recordsA 検証A（① vs ②）の明細（要確認度順）
 * @param recordsB 検証B（② vs ③）の明細（要確認度順）
 * @param info     サマリ用のKPI・総額・使用ファイル等
 * @param warnings データ品質の警告
 * @param mail     メール下書き用の数字
 * @param checkD   検証D（国分のみ・湖南は null）
 */
public record VerifyResult(
        FactoryProfile profile,
        YearMonthKey targetYm,
        List<RecordA> recordsA,
        List<RecordB> recordsB,
        Map<String, Object> info,
        List<String> warnings,
        MailSnapshot mail,
        MatomeCheckResult checkD,
        CheckCResult checkC) {

    public FactoryId factory() {
        return profile.id();
    }

    /** サマリ用の整数値（未設定は 0）。 */
    public long num(String key) {
        Object v = info.get(key);
        return v instanceof Number n ? n.longValue() : 0L;
    }

    /** サマリ用の実数値（未設定は 0）。 */
    public double dbl(String key) {
        Object v = info.get(key);
        return v instanceof Number n ? n.doubleValue() : 0.0;
    }

    /** サマリ用の文字列（未設定は空）。 */
    public String str(String key) {
        Object v = info.get(key);
        return v == null ? "" : String.valueOf(v);
    }

    /** 検証A 要確認件数（不一致＋翌月記載＋前月過不足＋片側のみ。手動判定は含めない）。 */
    public int requiredCheckA() {
        return (int) (num("A不一致") + num("A翌月記載件数") + num("A前月過不足件数") + num("A①のみ") + num("A②のみ"));
    }

    /** 検証B 要確認件数（不一致＋片側のみ）。 */
    public int requiredCheckB() {
        return (int) (num("B不一致") + num("B②のみ") + num("B③のみ"));
    }

    /** ③に対象月データが無く検証Bを実施していない。 */
    public boolean skippedB() {
        Object v = info == null ? null : info.get("検証Bスキップ");
        if (v instanceof Boolean b) {
            return b;
        }
        return "true".equalsIgnoreCase(String.valueOf(v));
    }

    /** 判明済み件数（前月調整・翌月記載・前月過不足・手動判定・月ずれ解消・枝番統合）。 */
    public int knownCount() {
        return (int) (num("A前月調整件数") + num("A翌月記載件数") + num("A前月過不足件数")
                + num("A手動判定件数") + num("B月ずれ解消") + num("B枝番統合"));
    }
}
