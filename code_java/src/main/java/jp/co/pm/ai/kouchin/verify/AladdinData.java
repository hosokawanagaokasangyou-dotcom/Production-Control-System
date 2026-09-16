package jp.co.pm.ai.kouchin.verify;

import java.util.Map;

/**
 * ③ アラジン「依頼NO別問合せ」の読み取り結果。
 *
 * @param byIrai     依頼NO→加工金額（項目「加工金額」行の「--合計--」列）
 * @param taishoText シート内の「対象年月 : ...」の文字列
 * @param taishoYm   対象年月（読めなければ null）
 */
public record AladdinData(Map<String, Double> byIrai, String taishoText, YearMonthKey taishoYm) {

    public double total() {
        return byIrai.values().stream().mapToDouble(Double::doubleValue).sum();
    }

    public int count() {
        return byIrai.size();
    }
}
