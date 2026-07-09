package jp.co.pm.ai.desktop.dispatch;

import java.util.Optional;

import jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;

/**
 * 加工途中・翌日配台ダイアログ: ロール本数入力と m 換算（{@link DispatchInteractiveRollUnitSupport} と同趣旨）。
 */
public final class Stage2InProgressNextDayRollInput {

    private Stage2InProgressNextDayRollInput() {}

    /** 0 以上の整数ロール本数。空欄は 0。 */
    public static Optional<Integer> parseNonNegativeRollCount(String raw) {
        if (raw == null || raw.isBlank()) {
            return Optional.of(0);
        }
        try {
            String s = raw.strip().replace(",", "");
            if (s.contains(".")) {
                double d = Double.parseDouble(s);
                if (Math.abs(d - Math.rint(d)) <= 1e-9 && d >= 0) {
                    return Optional.of((int) Math.rint(d));
                }
                return Optional.empty();
            }
            int n = Integer.parseInt(s);
            return n >= 0 ? Optional.of(n) : Optional.empty();
        } catch (NumberFormatException e) {
            return Optional.empty();
        }
    }

    public static int maxRolls(double remainingM, double unitM) {
        return DispatchInteractiveRollUnitSupport.maxMoveRollCount(remainingM, unitM);
    }

    public static int defaultRollCount(double remainingM, double unitM) {
        return DispatchInteractiveRollUnitSupport.defaultMoveRollCount(remainingM, unitM);
    }

    /**
     * アラジン当日計画が当日完了したとみなしたときの翌日配台対象 m。
     *
     * <p>アラジン当日量が無いときは残量そのまま。あるときは {@code 残量 − max(0, アラジン当日 − 実加工)}。
     */
    public static double nextDayTargetMetersAssumingAladdinTodayComplete(
            double remainingM, double actualDoneM, double aladdinTodayM) {
        double rem = Math.max(0.0, remainingM);
        if (aladdinTodayM <= 1e-12) {
            return rem;
        }
        double todayShortfall = Math.max(0.0, aladdinTodayM - Math.max(0.0, actualDoneM));
        return Math.max(0.0, rem - todayShortfall);
    }

    /** 上記対象 m をロール換算した初期本数（残量上限内）。 */
    public static int defaultRollCountAssumingAladdinTodayComplete(
            double remainingM, double actualDoneM, double aladdinTodayM, double unitM) {
        double target =
                nextDayTargetMetersAssumingAladdinTodayComplete(
                        remainingM, actualDoneM, aladdinTodayM);
        double cap = Math.min(target, Math.max(0.0, remainingM));
        return defaultRollCount(cap, unitM);
    }

    /** 上限 m（例: アラジン当日量）と残量の小さい方で最大ロール本数を返す。 */
    public static int maxRollsForCap(double capM, double remainingM, double unitM) {
        double effectiveCap = Math.min(Math.max(0.0, capM), Math.max(0.0, remainingM));
        return maxRolls(effectiveCap, unitM);
    }

    public static int defaultRollCountForCap(double capM, double remainingM, double unitM) {
        return maxRollsForCap(capM, remainingM, unitM);
    }

    /**
     * ロール本数から翌日配台 m を解決する。
     *
     * @return empty は入力不正・ロール単位不明・残量超過
     */
    public static Optional<Double> resolveNextDayMeters(int rolls, double remainingM, double unitM) {
        if (rolls < 0) {
            return Optional.empty();
        }
        if (rolls == 0) {
            return Optional.of(0.0);
        }
        if (unitM <= 1e-9) {
            return Optional.empty();
        }
        int maxRolls = maxRolls(remainingM, unitM);
        if (rolls > maxRolls) {
            return Optional.empty();
        }
        double m = DispatchInteractiveRollUnitSupport.metersForRollCount(rolls, unitM);
        if (m > remainingM + 1e-6) {
            return Optional.empty();
        }
        return Optional.of(Stage2InProgressNextDayDispatchIo.sanitizeMeters(m));
    }

    public static String formatConvertedMetersPreview(int rolls, double unitM) {
        if (rolls <= 0) {
            return "0 m";
        }
        if (unitM <= 1e-9) {
            return "—";
        }
        return ResultDispatchNormalizer.formatQty(
                        DispatchInteractiveRollUnitSupport.metersForRollCount(rolls, unitM))
                + " m";
    }

    /** 入力検証。問題なしは empty。 */
    public static Optional<String> validateRollInput(
            String rollCountRaw,
            double remainingM,
            Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo) {
        Optional<Integer> rollsOpt = parseNonNegativeRollCount(rollCountRaw);
        if (rollsOpt.isEmpty()) {
            return Optional.of("翌日配台 (ロール) に 0 以上の整数を入力してください。");
        }
        int rolls = rollsOpt.get();
        if (rolls == 0) {
            return Optional.empty();
        }
        double unitM = unitInfo != null ? unitInfo.unitM() : 0.0;
        if (unitM <= 1e-9) {
            return Optional.of(
                    "配台ロール単位 (m) を決定できません。配台計画_タスク入力の行または使用原反テーブルを確認してください。");
        }
        int maxRolls = maxRolls(remainingM, unitM);
        if (rolls > maxRolls) {
            return Optional.of(
                    String.format(
                            java.util.Locale.ROOT,
                            "最大 %d ロールまでです（残量 %s m、1 ロール = %s m）。",
                            maxRolls,
                            formatPlainM(remainingM),
                            ResultDispatchNormalizer.formatQty(unitM)));
        }
        if (resolveNextDayMeters(rolls, remainingM, unitM).isEmpty()) {
            return Optional.of("翌日配台が残量を超えています。");
        }
        return Optional.empty();
    }

    /** アラジン翌日除外ダイアログ用。上限は残量のみ（アラジン当日量には上限を設けない）。 */
    public static Optional<String> validateExcludeRollInput(
            String rollCountRaw,
            double remainingM,
            Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo) {
        Optional<Integer> rollsOpt = parseNonNegativeRollCount(rollCountRaw);
        if (rollsOpt.isEmpty()) {
            return Optional.of("翌日除外 (ロール) に 0 以上の整数を入力してください。");
        }
        int rolls = rollsOpt.get();
        if (rolls == 0) {
            return Optional.empty();
        }
        double unitM = unitInfo != null ? unitInfo.unitM() : 0.0;
        if (unitM <= 1e-9) {
            return Optional.of(
                    "配台ロール単位 (m) を決定できません。配台計画_タスク入力の行または使用原反テーブルを確認してください。");
        }
        int maxRolls = maxRolls(remainingM, unitM);
        if (rolls > maxRolls) {
            return Optional.of(
                    String.format(
                            java.util.Locale.ROOT,
                            "最大 %d ロールまでです（残量 %s m、1 ロール = %s m）。",
                            maxRolls,
                            formatPlainM(remainingM),
                            ResultDispatchNormalizer.formatQty(unitM)));
        }
        if (resolveNextDayMeters(rolls, remainingM, unitM).isEmpty()) {
            return Optional.of("翌日除外量が残量を超えています。");
        }
        return Optional.empty();
    }

    private static String formatPlainM(double v) {
        if (Math.abs(v - Math.rint(v)) <= 1e-9) {
            return String.valueOf((long) Math.rint(v));
        }
        return String.format(java.util.Locale.ROOT, "%.3f", v).replaceAll("0+$", "").replaceAll("\\.$", "");
    }
}
