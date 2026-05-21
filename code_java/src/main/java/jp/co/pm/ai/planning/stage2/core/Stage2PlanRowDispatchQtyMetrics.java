package jp.co.pm.ai.planning.stage2.core;

import java.util.Map;
import java.util.Optional;

/**
 * Python {@code planning_core._core._plan_row_dispatch_qty_metrics} に相当（結果シート・配台メトリクス）。
 *
 * <p>正: 段階1の列「配台使用残数量」「配台ロール数」（欠損時は段階1式で補完）。
 * 済相当 = max(0, 換算数量 − 配台使用残数量)。総量 = 残り + 済相当（換算数量 raw。100m 切上げなし）。
 *
 * <p>「未加工」列が無い・空・非数値の行は {@link Optional#empty()}（結果シートは残量列を埋めない）。
 * 未加工は行の有効性検証のみ（数値の意味はメトリクスに使わない）。
 */
public final class Stage2PlanRowDispatchQtyMetrics {

    private static final double EPS = 1e-12;

    private static final String COL_DISPATCH_REMAINING = "配台使用残数量";
    private static final String COL_DISPATCH_ROLLS = "配台ロール数";
    private static final String COL_RAW_ROLL = "(原反)ロール単位長さ";
    private static final String COL_RAW_ROLL_ALT = "（原反）ロール単位長さ";

    private Stage2PlanRowDispatchQtyMetrics() {}

    public record Metrics(double remainingM, double doneM, double qtyTotalForDispatchM) {}

    public static Optional<Metrics> compute(Map<String, String> row, Stage2RollUnitLengthTables tables) {
        if (row == null || !row.containsKey("未加工")) {
            return Optional.empty();
        }
        if (Stage2RollUnitLengthTables.optionalUnprocessedCell(row.get("未加工")).isEmpty()) {
            return Optional.empty();
        }
        double qtyConvRaw = Stage2RollUnitLengthTables.parseFloatSafe(row.get("換算数量"), 0.0);
        double remainingM = planDispatchRemainingM(row, tables);
        double doneM = Math.max(0.0, qtyConvRaw - remainingM);
        double qtyTotalForDispatchM = remainingM + doneM;
        return Optional.of(new Metrics(remainingM, doneM, qtyTotalForDispatchM));
    }

    /** 段階1/2: 列「配台使用残数量」を正とする残量(m)。欠損時は段階1式で補完。 */
    static double planDispatchRemainingM(Map<String, String> row, Stage2RollUnitLengthTables tables) {
        double fromCol = Stage2RollUnitLengthTables.parseFloatSafe(row.get(COL_DISPATCH_REMAINING), -1.0);
        if (fromCol >= 0) {
            return fromCol;
        }
        return computeDispatchRemainingFromFormula(row, tables);
    }

    /** 段階1/2: 列「配台ロール数」を正とする本数。欠損時は残量÷原反ロール長で補完。 */
    static double planDispatchRollCount(
            Map<String, String> row, double remainingM, Stage2RollUnitLengthTables tables) {
        double fromCol = Stage2RollUnitLengthTables.parseFloatSafe(row.get(COL_DISPATCH_ROLLS), -1.0);
        if (fromCol >= 0) {
            return fromCol;
        }
        return rollCountFromRemaining(row, remainingM, tables);
    }

    /**
     * 段階2 {@code build_task_queue_from_planning_df} が {@code unit_m = 配台使用残数量 / 配台ロール数} を使う経路か。
     * このときは {@code _effective_roll_unit_m_for_dispatch_task_simulator} による実効化・(原反)ロール列の黄着色は行わない。
     */
    public static boolean stage2SimulatorUsesDispatchRollCountColumns(
            Map<String, String> row, Stage2RollUnitLengthTables tables) {
        if (row == null) {
            return false;
        }
        double rem = planDispatchRemainingM(row, tables);
        double rolls = planDispatchRollCount(row, rem, tables);
        return rem > EPS && rolls > EPS;
    }

    private static double computeDispatchRemainingFromFormula(
            Map<String, String> row, Stage2RollUnitLengthTables tables) {
        double qtyConv = Stage2RollUnitLengthTables.parseFloatSafe(row.get("換算数量"), 0.0);
        double actualDone = Stage2RollUnitLengthTables.parseFloatSafe(row.get("実加工数"), 0.0);
        double b = Math.max(0.0, qtyConv - actualDone);
        double rawRoll = rawRollUnitMFromPlanRow(row, 100.0, tables);
        if (rawRoll <= EPS) {
            return b;
        }
        if (b <= EPS) {
            return 0.0;
        }
        int nRolls = (int) Math.ceil(b / rawRoll);
        return rawRoll * nRolls;
    }

    private static double rollCountFromRemaining(
            Map<String, String> row, double remainingM, Stage2RollUnitLengthTables tables) {
        double rawRoll = rawRollUnitMFromPlanRow(row, 100.0, tables);
        if (rawRoll <= EPS || remainingM <= EPS) {
            return 0.0;
        }
        double n = remainingM / rawRoll;
        if (Math.abs(n - Math.rint(n)) <= 1e-9) {
            return Math.rint(n);
        }
        return n;
    }

    /** Python {@code _dispatch_simulator_unit_m_from_plan_row} に相当（原反ロール長のみ）。 */
    static double rawRollUnitMFromPlanRow(
            Map<String, String> row, double fallbackM, Stage2RollUnitLengthTables tables) {
        String usedRaw = nz(row.get("使用原反"));
        double unit = Stage2RollUnitLengthTables.parseFloatSafe(row.get(COL_RAW_ROLL), 0.0);
        if (unit <= 0) {
            unit = Stage2RollUnitLengthTables.parseFloatSafe(row.get(COL_RAW_ROLL_ALT), 0.0);
        }
        double fb = Math.max(1e-9, fallbackM);
        if (unit <= 0 && tables != null) {
            unit = tables.lookupByUsedRaw(usedRaw).orElse(0.0);
        }
        if (unit <= 0) {
            unit = Stage2RollUnitLengthTables.inferFromProductDimensions(usedRaw, fb);
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

    /**
     * 段階2 {@code build_task_queue_from_planning_df} の {@code unit_m}（配台ロール単位 m）に相当。
     *
     * <p>{@code dispatch_rolls > 0} かつ {@code dispatch_m > 0} のときは {@code dispatch_m / dispatch_rolls}
     *（実効化しない）。それ以外は原反ロール長＋実効ロール化。
     */
    public record DispatchSimulatorUnitM(
            double unitM, double dispatchQtyM, double dispatchRollCount, boolean fromDispatchRollColumns) {}

    public static DispatchSimulatorUnitM dispatchSimulatorUnitMFromPlanRow(
            Map<String, String> row, Stage2RollUnitLengthTables tables) {
        if (row == null || row.isEmpty()) {
            return new DispatchSimulatorUnitM(0.0, 0.0, 0.0, false);
        }
        Stage2RollUnitLengthTables t = tables != null ? tables : Stage2RollUnitLengthTables.empty();
        double rem = planDispatchRemainingM(row, t);
        double rolls = planDispatchRollCount(row, rem, t);
        double qty = Math.max(0.0, rem);
        if (rolls > EPS && qty > EPS) {
            return new DispatchSimulatorUnitM(qty / rolls, qty, rolls, true);
        }
        double qtyConv = Stage2RollUnitLengthTables.parseFloatSafe(row.get("換算数量"), 0.0);
        double actualDone = Stage2RollUnitLengthTables.parseFloatSafe(row.get("実加工数"), 0.0);
        double doneM = Math.max(0.0, qtyConv - actualDone);
        double qtyTotal = qty + doneM;
        double fb = qtyTotal > EPS ? qtyTotal : (qty > EPS ? qty : 100.0);
        double unit = rawRollUnitMFromPlanRow(row, fb, t);
        if (qty > EPS && unit > EPS) {
            unit = effectiveRollUnitMForDispatchTaskSimulator(qty, unit);
        }
        return new DispatchSimulatorUnitM(unit, qty, rolls, false);
    }

    /** Python {@code _effective_roll_unit_m_for_dispatch_task_simulator} に相当。 */
    static double effectiveRollUnitMForDispatchTaskSimulator(double qtyM, double sheetRollUnitM) {
        double q = qtyM;
        double u = sheetRollUnitM;
        if (q <= EPS || u <= EPS) {
            return u > EPS ? u : 0.0;
        }
        double nRaw = q / u;
        if (nRaw <= EPS) {
            return u;
        }
        if (Math.abs(nRaw - Math.rint(nRaw)) <= 1e-9) {
            return u;
        }
        int nWork = (int) Math.floor(nRaw);
        if (nWork < 1) {
            nWork = 1;
        }
        return q / (double) nWork;
    }

    /** 数量が配台ロール単位の整数倍か（0 は可）。 */
    public static boolean isQtyAlignedToRollUnit(double qtyM, double unitM) {
        if (qtyM <= EPS) {
            return true;
        }
        if (unitM <= EPS) {
            return false;
        }
        double n = qtyM / unitM;
        return Math.abs(n - Math.rint(n)) <= 1e-6 * Math.max(1.0, Math.abs(n));
    }

    /** {@code maxM} を超えない最大のロール整数倍数量 (m)。 */
    public static double largestRollMultipleNotExceeding(double maxM, double unitM) {
        if (maxM <= EPS || unitM <= EPS) {
            return 0.0;
        }
        int n = (int) Math.floor((maxM + 1e-9) / unitM);
        if (n < 1) {
            return 0.0;
        }
        return unitM * n;
    }
}
