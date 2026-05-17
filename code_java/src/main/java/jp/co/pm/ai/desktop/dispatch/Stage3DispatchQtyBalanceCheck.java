package jp.co.pm.ai.desktop.dispatch;

/**
 * 段階3配台数（実配台合計）が {@code 換算数量 − 実加工数} と一致するかの照合。
 */
public final class Stage3DispatchQtyBalanceCheck {

    public static final String COL_TITLE = "段階3配台照合";

    private static final double EPS = 1e-3;

    private Stage3DispatchQtyBalanceCheck() {}

    /**
     * @param hasStage3ActualColumn 実配台数量列がある（段階3試行済み）
     * @return 空（照合不可）／{@code OK}／{@code NG (期待…／配台…)}
     */
    public static String formatCheck(
            double qtyConverted,
            double actualProcessed,
            double stage3DispatchTotal,
            boolean hasStage3ActualColumn) {
        if (!hasStage3ActualColumn || stage3DispatchTotal <= EPS) {
            return "";
        }
        double expected = qtyConverted - actualProcessed;
        if (Math.abs(stage3DispatchTotal - expected) <= EPS) {
            return "OK";
        }
        return "NG (期待"
                + ResultDispatchNormalizer.formatQty(expected)
                + "／配台"
                + ResultDispatchNormalizer.formatQty(stage3DispatchTotal)
                + ")";
    }

    public static boolean isNgResult(String checkText) {
        return checkText != null && checkText.startsWith("NG");
    }
}
