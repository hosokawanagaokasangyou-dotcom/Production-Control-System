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

    private static String formatPlainM(double v) {
        if (Math.abs(v - Math.rint(v)) <= 1e-9) {
            return String.valueOf((long) Math.rint(v));
        }
        return String.format(java.util.Locale.ROOT, "%.3f", v).replaceAll("0+$", "").replaceAll("\\.$", "");
    }
}
