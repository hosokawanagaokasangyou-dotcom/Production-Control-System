package jp.co.pm.ai.kouchin.verify;

import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * ① 東レ送付CSV（RVSHEETyyyymm.csv）の読み取り結果。
 *
 * @param byKeiyaku      対象入庫場所の 契約NO→金額合計（H列「-」は符号反転済み）
 * @param byBasho        入庫場所→金額合計（対象外入庫場所の表示に使う）
 * @param nyukoDates     契約NO→入庫月日(yymmdd)の集合
 * @param minusRows      契約NO→取消行（H列「-」）
 * @param subtotalErrors ブロック小計(T行)とデータ行合計の不一致
 * @param warnings       データ品質の警告
 * @param rawRows        原本コピー用の全行
 */
public record TorayCsvData(
        Map<String, Double> byKeiyaku,
        Map<String, Double> byBasho,
        Map<String, Set<String>> nyukoDates,
        Map<String, List<MinusRow>> minusRows,
        List<SubtotalError> subtotalErrors,
        List<String> warnings,
        List<List<String>> rawRows) {

    /** 取消行（H列が「-」の行）。金額は符号反転後。 */
    public record MinusRow(int rowNo, String nyukoDate, double amount) {
    }

    /** 小計差と同額の契約NO候補。 */
    public record Candidate(String keiyaku, String basho, double amount) {
    }

    /** ブロック小計(T行)の検算不一致。 */
    public record SubtotalError(
            int blockStart,
            int subtotalRow,
            String bashos,
            double subtotal,
            double dataSum,
            double diff,
            List<Candidate> candidates) {
    }

    /** 対象入庫場所の総額。 */
    public double total() {
        return byKeiyaku.values().stream().mapToDouble(Double::doubleValue).sum();
    }

    /** 対象入庫場所の契約NO件数。 */
    public int keiyakuCount() {
        return byKeiyaku.size();
    }

    /** 取消行の総件数。 */
    public int minusRowCount() {
        return minusRows.values().stream().mapToInt(List::size).sum();
    }

    /** 取消行の金額合計（反転後）。 */
    public double minusTotal() {
        return minusRows.values().stream()
                .flatMap(List::stream)
                .mapToDouble(MinusRow::amount)
                .sum();
    }
}
