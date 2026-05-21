package jp.co.pm.ai.desktop.dispatch;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;

import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.TextInputDialog;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.PlanInputTabController;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;
import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;

/**
 * 配台計画手動修正タブ: 数量の移動・編集を配台ロール単位（段階2 {@code unit_m}）に揃える。
 */
public final class DispatchInteractiveRollUnitSupport {

    private DispatchInteractiveRollUnitSupport() {}

    /**
     * ワイド行プロファイルと（あれば）配台計画タスク入力行から {@code unit_m} を解決する。
     */
    public static Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM resolveUnitM(
            Map<String, String> wideProfile,
            PlanInputTabController planInputOrNull,
            Map<String, String> uiEnv,
            Stage2RollUnitLengthTables tablesOrNull) {
        Map<String, String> row = planRowForWideProfile(wideProfile, planInputOrNull);
        Stage2RollUnitLengthTables tables = tablesOrNull;
        if (tables == null) {
            try {
                tables = Stage2RollUnitLengthTables.load(AppPaths.resolveRepoRoot(uiEnv));
            } catch (Exception ignored) {
                tables = Stage2RollUnitLengthTables.empty();
            }
        }
        return Stage2PlanRowDispatchQtyMetrics.dispatchSimulatorUnitMFromPlanRow(row, tables);
    }

    static Map<String, String> planRowForWideProfile(
            Map<String, String> wideProfile, PlanInputTabController planInputOrNull) {
        if (wideProfile == null) {
            return Map.of();
        }
        String taskId = nz(wideProfile.get("依頼NO"));
        String process = nz(wideProfile.get(ResultDispatchSchema.COL_PROCESS));
        String machine = nz(wideProfile.get(ResultDispatchSchema.COL_MACHINE));
        if (planInputOrNull != null) {
            Optional<Map<String, String>> fromPlan =
                    planInputOrNull.findPlanRowMapByKeys(taskId, process, machine);
            if (fromPlan.isPresent()) {
                return fromPlan.get();
            }
        }
        return new LinkedHashMap<>(wideProfile);
    }

    /**
     * ロール単位での移動量を入力させる。キャンセル時は empty。
     */
    public static Optional<Double> pickRollAlignedMoveQuantity(
            Window owner,
            double maxM,
            Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo,
            String profileHint) {
        double unitM = unitInfo.unitM();
        if (unitM <= 1e-9) {
            Alert a = new Alert(AlertType.WARNING);
            if (owner != null) {
                a.initOwner(owner);
            }
            a.setTitle("配台ロール単位");
            a.setHeaderText("配台ロール単位 (m) を決定できません");
            a.setContentText(
                    (profileHint != null && !profileHint.isBlank() ? profileHint + "\n" : "")
                            + "配台計画_タスク入力の行、または使用原反テーブルを確認してください。");
            a.showAndWait();
            return Optional.empty();
        }
        double defaultAmt = largestMoveDefault(maxM, unitM);
        if (defaultAmt <= 1e-9) {
            Alert a = new Alert(AlertType.WARNING);
            if (owner != null) {
                a.initOwner(owner);
            }
            a.setTitle("配台ロール単位");
            a.setHeaderText("移動できるロール単位の数量がありません");
            a.setContentText(
                    rollUnitDialogHeader(maxM, unitInfo, profileHint)
                            + "\nセル数量が配台ロール単位より小さい可能性があります。");
            a.showAndWait();
            return Optional.empty();
        }

        String lastInput = ResultDispatchNormalizer.formatQty(defaultAmt);
        while (true) {
            TextInputDialog dialog = new TextInputDialog(lastInput);
            if (owner != null) {
                dialog.initOwner(owner);
            }
            dialog.setTitle("移動数量");
            dialog.setHeaderText(rollUnitDialogHeader(maxM, unitInfo, profileHint));
            dialog.setContentText("移動する数量 (m) — 配台ロール単位の整数倍のみ:");
            Optional<String> ov = dialog.showAndWait();
            if (ov.isEmpty() || ov.get().isBlank()) {
                return Optional.empty();
            }
            double v = ResultDispatchNormalizer.parseDouble(ov.get());
            if (v <= 1e-9) {
                warnInvalidRollQty(owner, unitM, "0 より大きい数量を入力してください。");
                lastInput = ov.get();
                continue;
            }
            if (v > maxM + 1e-9) {
                warnInvalidRollQty(
                        owner,
                        unitM,
                        "最大 "
                                + ResultDispatchNormalizer.formatQty(maxM)
                                + " m を超えています。");
                lastInput = ov.get();
                continue;
            }
            if (!Stage2PlanRowDispatchQtyMetrics.isQtyAlignedToRollUnit(v, unitM)) {
                warnInvalidRollQty(
                        owner,
                        unitM,
                        ResultDispatchNormalizer.formatQty(unitM)
                                + " m の整数倍で入力してください（例: "
                                + ResultDispatchNormalizer.formatQty(unitM)
                                + ", "
                                + ResultDispatchNormalizer.formatQty(2 * unitM)
                                + "）。");
                lastInput = ov.get();
                continue;
            }
            return Optional.of(v);
        }
    }

    /** 日付セル直接編集: ロール整数倍に揃えた数量。不正時は empty。 */
    public static Optional<Double> parseRollAlignedTotalQuantity(
            Window owner,
            String input,
            Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo,
            String profileHint) {
        double unitM = unitInfo.unitM();
        if (unitM <= 1e-9) {
            Alert a = new Alert(AlertType.WARNING);
            if (owner != null) {
                a.initOwner(owner);
            }
            a.setTitle("配台ロール単位");
            a.setHeaderText("配台ロール単位 (m) を決定できません");
            a.setContentText(profileHint != null ? profileHint : "");
            a.showAndWait();
            return Optional.empty();
        }
        double v = ResultDispatchNormalizer.parseDouble(input);
        if (!Stage2PlanRowDispatchQtyMetrics.isQtyAlignedToRollUnit(v, unitM)) {
            Alert a = new Alert(AlertType.WARNING);
            if (owner != null) {
                a.initOwner(owner);
            }
            a.setTitle("配台ロール単位");
            a.setHeaderText("数量は配台ロール単位の整数倍のみ設定できます");
            a.setContentText(
                    (profileHint != null && !profileHint.isBlank() ? profileHint + "\n" : "")
                            + "配台ロール単位: "
                            + ResultDispatchNormalizer.formatQty(unitM)
                            + " m");
            a.showAndWait();
            return Optional.empty();
        }
        return Optional.of(v);
    }

    public static String rollUnitDialogHeader(
            double maxM,
            Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo,
            String profileHint) {
        StringBuilder sb = new StringBuilder();
        if (profileHint != null && !profileHint.isBlank()) {
            sb.append(profileHint).append('\n');
        }
        sb.append("配台ロール単位: ")
                .append(ResultDispatchNormalizer.formatQty(unitInfo.unitM()))
                .append(" m");
        if (unitInfo.fromDispatchRollColumns()) {
            sb.append("（配台使用残数量 ÷ 配台ロール数）");
        } else {
            sb.append("（原反ロール長・実効化）");
        }
        if (unitInfo.dispatchRollCount() > 1e-9) {
            sb.append('\n')
                    .append("配台ロール数: ")
                    .append(formatRollCount(unitInfo.dispatchRollCount()))
                    .append(" 本");
        }
        int maxRolls = (int) Math.floor((maxM + 1e-9) / unitInfo.unitM());
        sb.append('\n')
                .append("移動可能: 最大 ")
                .append(ResultDispatchNormalizer.formatQty(maxM))
                .append(" m（")
                .append(maxRolls)
                .append(" ロールまで）");
        return sb.toString();
    }

    private static double largestMoveDefault(double maxM, double unitM) {
        return Stage2PlanRowDispatchQtyMetrics.largestRollMultipleNotExceeding(maxM, unitM);
    }

    private static void warnInvalidRollQty(Window owner, double unitM, String detail) {
        Alert a = new Alert(AlertType.WARNING);
        if (owner != null) {
            a.initOwner(owner);
        }
        a.setTitle("配台ロール単位");
        a.setHeaderText("入力できません");
        a.setContentText(
                "配台ロール単位: "
                        + ResultDispatchNormalizer.formatQty(unitM)
                        + " m\n"
                        + detail);
        a.showAndWait();
    }

    private static String formatRollCount(double rolls) {
        if (Math.abs(rolls - Math.rint(rolls)) <= 1e-6) {
            return String.valueOf((long) Math.rint(rolls));
        }
        return ResultDispatchNormalizer.formatQty(rolls);
    }

    private static String nz(String s) {
        return s != null ? s.strip() : "";
    }
}
