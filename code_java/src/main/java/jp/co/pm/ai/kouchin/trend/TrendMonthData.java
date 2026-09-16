package jp.co.pm.ai.kouchin.trend;

import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.kouchin.verify.YearMonthKey;

/**
 * 1工場・1か月分の種類別／シート区分別集計。
 */
public record TrendMonthData(
        YearMonthKey ym,
        Path path,
        Map<String, TrendMetric> kinds,
        Map<String, TrendMetric> sheets,
        List<String> warnings) {

    public TrendMonthData {
        kinds = kinds == null ? Map.of() : Map.copyOf(kinds);
        sheets = sheets == null ? Map.of() : Map.copyOf(sheets);
        warnings = warnings == null ? List.of() : List.copyOf(warnings);
    }
}
