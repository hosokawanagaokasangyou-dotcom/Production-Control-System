package jp.co.pm.ai.desktop.dispatch;

import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;

/**
 * 配台計画手動修正タブ: 段階3前の日別 {@code 当日配台数量} をアラジン計画に沿うようロール単位で再配分する。
 *
 * <p>換算数量 &lt; 原反ロール長のときはアラジン表示が換算数量（例: 20 m）でも、配台数量はロール単位（例: 300 m）
 * で移動する。
 */
public final class DispatchAladdinPlanAligner {

    private static final double EPS = 1e-9;

    private DispatchAladdinPlanAligner() {}

    public record RowInput(
            double[] currentByDayIndex,
            double[] aladdinByDayIndex,
            double unitM,
            boolean usesConvertedQtyForAladdin) {}

    public record RowResult(double[] newByDayIndex, boolean changed, int rollMoves) {}

    /**
     * 1 タスク行の日別数量を、アラジン計画の優先度に従いロール単位で再配分する。合計 m は維持する（ロール整数倍の範囲内）。
     */
    public static RowResult alignRow(RowInput input) {
        if (input == null
                || input.currentByDayIndex() == null
                || input.aladdinByDayIndex() == null) {
            return unchanged(null);
        }
        int n = input.currentByDayIndex().length;
        if (n == 0 || n != input.aladdinByDayIndex().length) {
            return unchanged(input.currentByDayIndex());
        }
        double unitM = input.unitM();
        if (unitM <= EPS) {
            return unchanged(input.currentByDayIndex());
        }

        double[] current = input.currentByDayIndex();
        double total = 0.0;
        for (double v : current) {
            total += Math.max(0.0, v);
        }
        if (total <= EPS) {
            return unchanged(current);
        }
        if (!Stage2PlanRowDispatchQtyMetrics.isQtyAlignedToRollUnit(total, unitM)) {
            return unchanged(current);
        }

        int totalRolls = (int) Math.round(total / unitM);
        if (totalRolls < 1) {
            return unchanged(current);
        }

        double[] weights = buildPriorityWeights(input.aladdinByDayIndex(), input.usesConvertedQtyForAladdin());
        if (sum(weights) <= EPS) {
            return unchanged(current);
        }

        int[] rolls = allocateRollsByWeight(totalRolls, weights);
        double[] target = new double[n];
        for (int i = 0; i < n; i++) {
            target[i] = rolls[i] * unitM;
        }

        if (arraysNearEqual(current, target)) {
            return new RowResult(current.clone(), false, 0);
        }

        int rollMoves = countRollMoves(current, target, unitM);
        return new RowResult(target, true, rollMoves);
    }

    private static double[] buildPriorityWeights(double[] aladdin, boolean usesConvertedQtyForAladdin) {
        double[] weights = new double[aladdin.length];
        for (int i = 0; i < aladdin.length; i++) {
            double a = Math.max(0.0, aladdin[i]);
            if (usesConvertedQtyForAladdin) {
                weights[i] = a > EPS ? 1.0 : 0.0;
            } else {
                weights[i] = a;
            }
        }
        return weights;
    }

    /** 最大剰余法でロール本数を配分する。重み 0 の日には配分しない。 */
    static int[] allocateRollsByWeight(int totalRolls, double[] weights) {
        int n = weights.length;
        int[] rolls = new int[n];
        if (totalRolls < 1 || n == 0) {
            return rolls;
        }
        double weightSum = sum(weights);
        if (weightSum <= EPS) {
            return rolls;
        }

        double[] remainders = new double[n];
        int assigned = 0;
        for (int i = 0; i < n; i++) {
            if (weights[i] <= EPS) {
                remainders[i] = -1.0;
                continue;
            }
            double exact = totalRolls * weights[i] / weightSum;
            rolls[i] = (int) Math.floor(exact + EPS);
            remainders[i] = exact - rolls[i];
            assigned += rolls[i];
        }

        while (assigned < totalRolls) {
            int best = -1;
            double bestRemainder = -1.0;
            for (int i = 0; i < n; i++) {
                if (weights[i] <= EPS || remainders[i] < 0.0) {
                    continue;
                }
                if (remainders[i] > bestRemainder + EPS) {
                    bestRemainder = remainders[i];
                    best = i;
                }
            }
            if (best < 0) {
                break;
            }
            rolls[best]++;
            remainders[best] = -1.0;
            assigned++;
        }
        return rolls;
    }

    private static int countRollMoves(double[] from, double[] to, double unitM) {
        if (from == null || to == null || from.length != to.length || unitM <= EPS) {
            return 0;
        }
        double surplus = 0.0;
        double deficit = 0.0;
        for (int i = 0; i < from.length; i++) {
            double delta = to[i] - from[i];
            if (delta > EPS) {
                deficit += delta;
            } else if (delta < -EPS) {
                surplus += -delta;
            }
        }
        return (int) Math.round(Math.min(surplus, deficit) / unitM);
    }

    private static boolean arraysNearEqual(double[] a, double[] b) {
        if (a == null || b == null || a.length != b.length) {
            return false;
        }
        for (int i = 0; i < a.length; i++) {
            if (Math.abs(a[i] - b[i]) > EPS) {
                return false;
            }
        }
        return true;
    }

    private static double sum(double[] values) {
        double s = 0.0;
        if (values == null) {
            return s;
        }
        for (double v : values) {
            s += v;
        }
        return s;
    }

    private static RowResult unchanged(double[] current) {
        if (current == null) {
            return new RowResult(new double[0], false, 0);
        }
        return new RowResult(current.clone(), false, 0);
    }
}
