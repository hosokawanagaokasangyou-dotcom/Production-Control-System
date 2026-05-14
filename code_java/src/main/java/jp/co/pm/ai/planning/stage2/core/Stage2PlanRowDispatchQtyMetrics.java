package jp.co.pm.ai.planning.stage2.core;

import java.util.Map;
import java.util.Optional;

/**
 * Python {@code planning_core._core._plan_row_dispatch_qty_metrics} に相当する純粋関数（配台前の行メトリクス）。
 *
 * <p>「未加工」列が無い・空・非数値の行は {@link Optional#empty()}（PassThrough は残量列を埋めない）。
 * ワークスペース解釈（換算数量&gt;0 かつ実加工数=0 かつ未加工セル=0 → 全数未加工＝換算数量ぶん残り）を適用する。
 */
public final class Stage2PlanRowDispatchQtyMetrics {

    private static final double EPS = 1e-12;
    private static final double CEIL_STEP_M = 100.0;

    private Stage2PlanRowDispatchQtyMetrics() {}

    public record Metrics(double remainingM, double doneM, double qtyTotalForDispatchM) {}

    public static Optional<Metrics> compute(Map<String, String> row, Stage2RollUnitLengthTables tables) {
        if (row == null || !row.containsKey("未加工")) {
            return Optional.empty();
        }
        Optional<Double> unpCell = Stage2RollUnitLengthTables.optionalUnprocessedCell(row.get("未加工"));
        if (unpCell.isEmpty()) {
            return Optional.empty();
        }
        double qtyConvRaw = Stage2RollUnitLengthTables.parseFloatSafe(row.get("換算数量"), 0.0);
        double qtyTotalCeiled = ceilRollUnitLengthMToNextStep(qtyConvRaw, CEIL_STEP_M);
        double actualDone = Stage2RollUnitLengthTables.parseFloatSafe(row.get("実加工数"), 0.0);

        double unp = unpCell.get();
        if (qtyConvRaw > EPS && Math.abs(unp) <= EPS && actualDone <= EPS) {
            unp = qtyConvRaw;
        }

        if (unp > EPS) {
            double remainingM = Math.max(0.0, unp);
            double doneM = Math.max(0.0, qtyConvRaw - unp);
            return Optional.of(new Metrics(remainingM, doneM, qtyTotalCeiled));
        }
        double fallbackM =
                Math.max(1e-9, Stage2RollUnitLengthTables.parseFloatSafe(String.valueOf(qtyTotalCeiled), 0.0));
        if (fallbackM <= 1e-9) {
            fallbackM = Math.max(1e-9, qtyConvRaw);
        }
        if (fallbackM <= 1e-9) {
            fallbackM = 1.0;
        }
        double rollM = rollUnitMFromPlanRow(row, fallbackM, tables);
        double baseM = Math.max(0.0, qtyTotalCeiled);
        if (rollM > EPS && qtyConvRaw > EPS) {
            double nRollsRaw = qtyConvRaw / rollM;
            if (Math.abs(nRollsRaw - Math.rint(nRollsRaw)) <= 1e-9) {
                baseM = Math.max(0.0, qtyConvRaw);
            }
        }
        double remainingM = rollM > 0 ? Math.max(baseM, rollM) : baseM;
        double qtyTotalForDispatchM = remainingM;
        return Optional.of(new Metrics(remainingM, 0.0, qtyTotalForDispatchM));
    }

    static double ceilRollUnitLengthMToNextStep(double rollM, double stepM) {
        if (!(rollM > 0)) {
            return rollM;
        }
        double step = stepM > 0 ? stepM : 100.0;
        return Math.ceil(rollM / step) * step;
    }

    static double rollUnitMFromPlanRow(
            Map<String, String> row, double fallbackM, Stage2RollUnitLengthTables tables) {
        String product = nz(row.get("製品名"));
        String usedRaw = nz(row.get("使用原反"));
        double unit = Stage2RollUnitLengthTables.parseFloatSafe(row.get("ロール単位長さ"), 0.0);
        double fb = Math.max(1e-9, fallbackM);
        if (unit <= 0) {
            unit =
                    tables.lookupByUsedRaw(usedRaw)
                            .or(() -> tables.lookupByProductName(product))
                            .orElse(0.0);
        }
        if (unit <= 0) {
            unit = Stage2RollUnitLengthTables.inferFromProductDimensions(product, fb);
        }
        if (unit <= 0) {
            unit = fb;
        }
        return unit;
    }

    private static String nz(String s) {
        return s != null ? s : "";
    }

    /** 残加工量・累計加工量・完了率(実行時点) を結果シート用の文字列で返す。 */
    public static Optional<ResultTaskQtyStrings> toResultSheetStrings(Metrics m) {
        if (m == null) {
            return Optional.empty();
        }
        String rem = Stage2RollUnitLengthTables.formatMetersPlain(m.remainingM());
        String cum = Stage2RollUnitLengthTables.formatMetersPlain(m.doneM());
        double denom = m.qtyTotalForDispatchM();
        String pct =
                denom > EPS
                        ? Stage2RollUnitLengthTables.formatPercentPlain(m.doneM() / denom)
                        : "";
        return Optional.of(new ResultTaskQtyStrings(rem, cum, pct));
    }

    public record ResultTaskQtyStrings(String remainingM, String cumulativeDoneM, String completionPct) {}
}
