package jp.co.pm.ai.kouchin.trend;

/**
 * 加工賃・加工量の組。読取直後は円 / m。Excel 表示前に千円 / km へ変換する。
 */
public record TrendMetric(double wage, double qty) {

    public TrendMetric plus(TrendMetric other) {
        if (other == null) {
            return this;
        }
        return new TrendMetric(wage + other.wage, qty + other.qty);
    }

    public double value(String metric) {
        return "qty".equals(metric) ? qty : wage;
    }
}
