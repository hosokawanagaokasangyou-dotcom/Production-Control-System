package jp.co.pm.ai.kouchin.trend;

import jp.co.pm.ai.kouchin.verify.YearMonthKey;

/**
 * 直近2か月の前月差による振替・シフト疑い行。判定は仮説であり断定しない。
 */
public record TrendShiftRow(
        String process,
        YearMonthKey prev,
        YearMonthKey latest,
        double kokubuPrev,
        double kokubuLatest,
        double konanPrev,
        double konanLatest,
        double dKokubu,
        double dKonan,
        double dCombined,
        double offset,
        String badge,
        String metric) {}
