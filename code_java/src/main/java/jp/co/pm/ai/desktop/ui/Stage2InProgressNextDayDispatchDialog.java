package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.Optional;

import jp.co.pm.ai.desktop.dispatch.Stage2InProgressNextDayRollInput;
import jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;

/**
 * 段階2直前: 加工途中タスクの翌日配台量をロール本数で一括入力する。
 */
public final class Stage2InProgressNextDayDispatchDialog {

    private static final Stage2NextDayRollDispatchDialogSupport.Theme THEME =
            new Stage2NextDayRollDispatchDialogSupport.Theme(
                    "加工途中タスク — 翌日の配台量",
                    "実加工数が入っている行について、翌日に配台するロール数を指定してください。"
                            + " 0 の行は段階2の配台対象から外します。"
                            + " アラジン当日計画がある行は、当日分が完了した前提で初期ロール数を自動計算します。",
                    "配台計画手動修正タブと同様、配台ロール単位 (m) の整数倍で配台します。"
                            + " 初期値はアラジン当日計画が当日完了したとみなした翌日配台対象量（残量以内の最大ロール本数）です。"
                            + " アラジン当日量が無い行は残量ベースです。"
                            + " 翌日配台は 0 以上・残量に収まるロール整数倍のみ。OK で未確定の入力も反映します。",
                    "実加工",
                    "翌日(ロール)",
                    "",
                    "-fx-font-size: 11px; -fx-text-fill: derive(-fx-text-inner-color, 22%);",
                    true,
                    true);

    private Stage2InProgressNextDayDispatchDialog() {}

    public static final class Row implements Stage2NextDayRollDispatchDialogSupport.RowModel {
        private final String taskId;
        private final String process;
        private final String machineName;
        private final double actualDoneM;
        private final double convertedQtyM;
        private final double dispatchQtyM;
        private final double aladdinTodayM;
        private final double remainingM;
        private final Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo;
        private final javafx.beans.property.SimpleStringProperty nextDayRollCount =
                new javafx.beans.property.SimpleStringProperty();

        public Row(
                String taskId,
                String process,
                String machineName,
                double actualDoneM,
                double convertedQtyM,
                double dispatchQtyM,
                double aladdinTodayM,
                double remainingM,
                Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo) {
            this.taskId = taskId != null ? taskId : "";
            this.process = process != null ? process : "";
            this.machineName = machineName != null ? machineName : "";
            this.actualDoneM = actualDoneM;
            this.convertedQtyM = convertedQtyM;
            this.dispatchQtyM = dispatchQtyM;
            this.aladdinTodayM = aladdinTodayM;
            this.remainingM = remainingM;
            this.unitInfo =
                    unitInfo != null
                            ? unitInfo
                            : new Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM(
                                    0.0, 0.0, 0.0, false);
            int defaultRolls =
                    Stage2InProgressNextDayRollInput.defaultRollCountAssumingAladdinTodayComplete(
                            remainingM, actualDoneM, aladdinTodayM, this.unitInfo.unitM());
            this.nextDayRollCount.set(String.valueOf(defaultRolls));
        }

        @Override
        public String taskId() {
            return taskId;
        }

        @Override
        public String process() {
            return process;
        }

        @Override
        public String machineName() {
            return machineName;
        }

        public double actualDoneM() {
            return actualDoneM;
        }

        public double aladdinTodayM() {
            return aladdinTodayM;
        }

        @Override
        public double aladdinTodayPlanM() {
            return aladdinTodayM;
        }

        @Override
        public double convertedQtyM() {
            return convertedQtyM;
        }

        @Override
        public double dispatchQtyM() {
            return dispatchQtyM;
        }

        @Override
        public double referenceM() {
            return actualDoneM;
        }

        @Override
        public double remainingM() {
            return remainingM;
        }

        @Override
        public Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo() {
            return unitInfo;
        }

        @Override
        public double unitM() {
            return unitInfo.unitM();
        }

        @Override
        public int maxRolls() {
            return Stage2InProgressNextDayRollInput.maxRolls(remainingM, unitM());
        }

        @Override
        public javafx.beans.property.StringProperty rollCountProperty() {
            return nextDayRollCount;
        }

        @Override
        public String targetReason() {
            return "加工途中";
        }

        Stage2InProgressNextDayDispatchIo.Entry toEntryFromNextDayInput() {
            int rolls =
                    Stage2InProgressNextDayRollInput.parseNonNegativeRollCount(
                                    nextDayRollCount.get())
                            .orElse(0);
            double nextDayM =
                    Stage2InProgressNextDayRollInput.resolveNextDayMeters(
                                    rolls, remainingM(), unitM())
                            .orElse(0.0);
            double shortfallM =
                    Stage2InProgressNextDayRollInput.aladdinTodayShortfallMeters(
                            remainingM, actualDoneM, aladdinTodayM);
            return new Stage2InProgressNextDayDispatchIo.Entry(
                    taskId, process, machineName, Math.max(0.0, nextDayM), shortfallM);
        }
    }

    /** @return 確定時は各行の入力値。キャンセル時は empty。 */
    public static Optional<List<Stage2InProgressNextDayDispatchIo.Entry>> prompt(
            javafx.stage.Window owner, List<Row> rows) {
        return Stage2NextDayRollDispatchDialogSupport.prompt(
                owner,
                rows,
                THEME,
                r -> ((Row) r).toEntryFromNextDayInput(),
                r ->
                        Stage2InProgressNextDayRollInput.validateRollInput(
                                r.rollCountProperty().get(), r.remainingM(), r.unitInfo()));
    }
}
