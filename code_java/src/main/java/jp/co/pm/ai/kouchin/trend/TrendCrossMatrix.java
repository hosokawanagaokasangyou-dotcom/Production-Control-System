package jp.co.pm.ai.kouchin.trend;

import java.util.List;
import java.util.Map;

import jp.co.pm.ai.kouchin.verify.YearMonthKey;

/**
 * 工場×工程（またはグループ）×月の横断表。
 * {@code values[工場][工程][ym]} は欠損月が {@code null}。
 */
public record TrendCrossMatrix(
        List<String> processes,
        List<YearMonthKey> months,
        String metric,
        Map<String, Map<String, Map<YearMonthKey, Double>>> values,
        boolean grouped) {

    public TrendCrossMatrix {
        processes = processes == null ? List.of() : List.copyOf(processes);
        months = months == null ? List.of() : List.copyOf(months);
        metric = metric == null ? "" : metric;
        values = values == null ? Map.of() : Map.copyOf(values);
    }
}
