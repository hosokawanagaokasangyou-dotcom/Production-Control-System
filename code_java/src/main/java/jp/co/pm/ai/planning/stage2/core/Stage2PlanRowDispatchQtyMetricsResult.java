package jp.co.pm.ai.planning.stage2.core;

/** {@link Stage2PlanRowDispatchQtyMetrics#compute} の戻り値（ネスト record だと javafx:run で .class 欠落が起きることがあるためトップレベル化）。 */
public record Stage2PlanRowDispatchQtyMetricsResult(
        double remainingM, double doneM, double qtyTotalForDispatchM) {}
