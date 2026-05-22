package jp.co.pm.ai.desktop.dispatch;

/**
 * 段階3配台数（日別配台 m の合計。手動移動後は当日配台数量、試行直後は実配台と一致）が
 * 配台ロール単位に揃えた残量と一致するかの照合。
 *
 * <p>残量 {@code 換算数量 − 実加工数} がロール単位より小さい場合、配台表・タイムラインには
 * 1 ロール分（ロール単位 m）が載るため、照合もロール切上げ後の m を期待値とする。
 */
public final class Stage3DispatchQtyBalanceCheck {

    public static final String COL_TITLE = "段階3配台照合";

    private static final double EPS = 1e-3;

    private Stage3DispatchQtyBalanceCheck() {}

    /**
     * @param hasStage3ActualColumn 実配台数量列がある（段階3試行済み）
     * @return 空（照合不可）／{@code OK}／{@code 20 (91m)}（残量＋ロール配台）／{@code NG (期待…／配台…)}
     */
    public static String formatCheck(
            double qtyConverted,
            double actualProcessed,
            double stage3DispatchTotal,
            boolean hasStage3ActualColumn) {
        return formatCheck(
                qtyConverted, actualProcessed, stage3DispatchTotal, hasStage3ActualColumn, 0.0);
    }

    /**
     * @param rollUnitM 配台ロール単位 (m)。0 以下のときはロール切上げなし（生の残量で照合）。
     */
    public static String formatCheck(
            double qtyConverted,
            double actualProcessed,
            double stage3DispatchTotal,
            boolean hasStage3ActualColumn,
            double rollUnitM) {
        if (!hasStage3ActualColumn || stage3DispatchTotal <= EPS) {
            return "";
        }
        double rawRemaining = Math.max(0.0, qtyConverted - actualProcessed);
        double expected = rollAlignedDispatchM(rawRemaining, rollUnitM);
        if (Math.abs(stage3DispatchTotal - expected) <= EPS) {
            if (rollUnitM > EPS && expected > rawRemaining + EPS) {
                return formatRollAlignedLabel(rawRemaining, expected);
            }
            return "OK";
        }
        return "NG (期待"
                + ResultDispatchNormalizer.formatQty(expected)
                + "／配台"
                + ResultDispatchNormalizer.formatQty(stage3DispatchTotal)
                + ")";
    }

    /** 段階2 {@code computeDispatchRemainingFromFormula} と同じロール切上げ。 */
    public static double rollAlignedDispatchM(double rawRemainingM, double rollUnitM) {
        if (rawRemainingM <= EPS) {
            return 0.0;
        }
        if (rollUnitM <= EPS) {
            return rawRemainingM;
        }
        int nRolls = (int) Math.ceil(rawRemainingM / rollUnitM - 1e-12);
        return rollUnitM * nRolls;
    }

    static String formatRollAlignedLabel(double rawRemainingM, double rollAlignedM) {
        return ResultDispatchNormalizer.formatQty(rawRemainingM)
                + " ("
                + ResultDispatchNormalizer.formatQty(rollAlignedM)
                + "m)";
    }

    public static boolean isNgResult(String checkText) {
        return checkText != null && checkText.startsWith("NG");
    }
}
