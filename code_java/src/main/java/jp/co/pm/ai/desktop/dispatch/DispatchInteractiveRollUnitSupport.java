package jp.co.pm.ai.desktop.dispatch;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TextInputDialog;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
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
                tables = Stage2RollUnitLengthTables.load(uiEnv);
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

    /** セル数量以下で移動できる最大ロール本数。 */
    public static int maxMoveRollCount(double maxM, double unitM) {
        if (maxM <= 1e-9 || unitM <= 1e-9) {
            return 0;
        }
        return (int) Math.floor((maxM + 1e-9) / unitM);
    }

    /** 移動ダイアログの既定ロール本数（可能ならセル分をすべて）。 */
    public static int defaultMoveRollCount(double maxM, double unitM) {
        return maxMoveRollCount(maxM, unitM);
    }

    public static double metersForRollCount(int rollCount, double unitM) {
        if (rollCount < 1 || unitM <= 1e-9) {
            return 0.0;
        }
        return unitM * rollCount;
    }

    public static String formatMoveMetersPreview(int rollCount, double unitM) {
        return "移動数量: "
                + ResultDispatchNormalizer.formatQty(metersForRollCount(rollCount, unitM))
                + " m（"
                + rollCount
                + " ロール × "
                + ResultDispatchNormalizer.formatQty(unitM)
                + " m）";
    }

    /**
     * ロール本数で移動量を入力させる（スピン＋手入力）。キャンセル時は empty。戻り値は移動 m。
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
        int maxRolls = maxMoveRollCount(maxM, unitM);
        if (maxRolls < 1) {
            Alert a = new Alert(AlertType.WARNING);
            if (owner != null) {
                a.initOwner(owner);
            }
            a.setTitle("配台ロール単位");
            a.setHeaderText("移動できるロールがありません");
            a.setContentText(
                    rollUnitDialogHeader(maxM, unitInfo, profileHint)
                            + "\nセル数量が配台ロール単位より小さい可能性があります。");
            a.showAndWait();
            return Optional.empty();
        }

        while (true) {
            Optional<MoveRollDialogOutcome> outcome =
                    showMoveRollCountDialog(
                            owner, maxM, unitM, maxRolls, unitInfo, profileHint);
            if (outcome.isEmpty()) {
                return Optional.empty();
            }
            MoveRollDialogOutcome o = outcome.get();
            if (o.cancelled()) {
                return Optional.empty();
            }
            if (o.validationMessage() != null) {
                warnInvalidRollQty(owner, unitM, o.validationMessage());
                continue;
            }
            return Optional.of(o.moveMeters());
        }
    }

    private record MoveRollDialogOutcome(boolean cancelled, double moveMeters, String validationMessage) {
        static MoveRollDialogOutcome ok(double moveMeters) {
            return new MoveRollDialogOutcome(false, moveMeters, null);
        }

        static MoveRollDialogOutcome cancel() {
            return new MoveRollDialogOutcome(true, 0.0, null);
        }

        static MoveRollDialogOutcome invalid(String message) {
            return new MoveRollDialogOutcome(false, 0.0, message);
        }
    }

    private static Optional<MoveRollDialogOutcome> showMoveRollCountDialog(
            Window owner,
            double maxM,
            double unitM,
            int maxRolls,
            Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo,
            String profileHint) {
        int defaultRolls = defaultMoveRollCount(maxM, unitM);
        Dialog<MoveRollDialogOutcome> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.setTitle("移動数量");
        dialog.setHeaderText(rollUnitDialogHeader(maxM, unitInfo, profileHint));
        dialog.getDialogPane().getButtonTypes().addAll(ButtonType.OK, ButtonType.CANCEL);

        Label metersPreview = new Label(formatMoveMetersPreview(defaultRolls, unitM));
        metersPreview.setStyle("-fx-font-weight: bold;");

        Spinner<Integer> rollSpinner =
                new Spinner<>(
                        new SpinnerValueFactory.IntegerSpinnerValueFactory(
                                1, maxRolls, defaultRolls, 1));
        rollSpinner.setEditable(true);
        rollSpinner.setPrefWidth(100);

        Runnable syncPreview =
                () -> {
                    int rolls = rollCountFromSpinnerEditor(rollSpinner).orElse(0);
                    if (rolls < 1) {
                        metersPreview.setText("移動数量: — m");
                    } else {
                        metersPreview.setText(formatMoveMetersPreview(rolls, unitM));
                    }
                };
        rollSpinner.valueProperty().addListener((obs, oldV, newV) -> syncPreview.run());
        rollSpinner.getEditor().textProperty().addListener((obs, oldT, newT) -> syncPreview.run());

        Label prompt = new Label("移動するロール数（本）— スピンまたは直接入力:");
        Label unitHint =
                new Label(
                        "1 ロール = "
                                + ResultDispatchNormalizer.formatQty(unitM)
                                + " m（最大 "
                                + maxRolls
                                + " ロール）");
        unitHint.setStyle("-fx-text-fill: #555555;");
        HBox inputRow = new HBox(8, rollSpinner, new Label("本"));
        inputRow.setAlignment(Pos.CENTER_LEFT);
        VBox content = new VBox(10, prompt, inputRow, metersPreview, unitHint);
        content.setPadding(new Insets(4, 0, 0, 0));
        dialog.getDialogPane().setContent(content);
        dialog.getDialogPane().setPrefWidth(420);

        dialog.setResultConverter(
                button -> {
                    if (button != ButtonType.OK) {
                        return MoveRollDialogOutcome.cancel();
                    }
                    Optional<Integer> rollsOpt = parseRollCountFromSpinner(rollSpinner);
                    if (rollsOpt.isEmpty()) {
                        return MoveRollDialogOutcome.invalid(
                                "ロール数は 1 以上の整数で入力してください。");
                    }
                    int rolls = rollsOpt.get();
                    if (rolls < 1) {
                        return MoveRollDialogOutcome.invalid("ロール数は 1 以上で入力してください。");
                    }
                    if (rolls > maxRolls) {
                        return MoveRollDialogOutcome.invalid(
                                "最大 "
                                        + maxRolls
                                        + " ロールまでです（"
                                        + ResultDispatchNormalizer.formatQty(maxM)
                                        + " m）。");
                    }
                    double moveM = metersForRollCount(rolls, unitM);
                    if (moveM > maxM + 1e-9) {
                        return MoveRollDialogOutcome.invalid(
                                "移動 "
                                        + ResultDispatchNormalizer.formatQty(moveM)
                                        + " m はセル数量 "
                                        + ResultDispatchNormalizer.formatQty(maxM)
                                        + " m を超えます。");
                    }
                    return MoveRollDialogOutcome.ok(moveM);
                });

        return dialog.showAndWait();
    }

    /** スピンエディタの文字列を優先して本数を読む（手入力中のプレビュー用。不正時は empty）。 */
    static Optional<Integer> rollCountFromSpinnerEditor(Spinner<Integer> spinner) {
        if (spinner == null || spinner.getEditor() == null) {
            return Optional.empty();
        }
        String text = spinner.getEditor().getText();
        if (text == null || text.isBlank()) {
            Integer v = spinner.getValue();
            return v != null && v >= 1 ? Optional.of(v) : Optional.empty();
        }
        try {
            String s = text.strip().replace(",", "");
            if (s.contains(".")) {
                double d = Double.parseDouble(s);
                if (Math.abs(d - Math.rint(d)) <= 1e-9 && d >= 1) {
                    return Optional.of((int) Math.rint(d));
                }
                return Optional.empty();
            }
            int n = Integer.parseInt(s);
            return n >= 1 ? Optional.of(n) : Optional.empty();
        } catch (NumberFormatException e) {
            return Optional.empty();
        }
    }

    static Optional<Integer> parseRollCountFromSpinner(Spinner<Integer> spinner) {
        return rollCountFromSpinnerEditor(spinner);
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
