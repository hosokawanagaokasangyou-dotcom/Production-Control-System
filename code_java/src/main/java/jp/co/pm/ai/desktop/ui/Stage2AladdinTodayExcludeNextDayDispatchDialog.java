package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.Optional;

import jp.co.pm.ai.desktop.dispatch.Stage2InProgressNextDayRollInput;
import jp.co.pm.ai.planning.stage2.Stage2AladdinTodayExcludeNextDayDispatchIo;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;

/** 段階2直前: アラジン当日配台がある行の翌日配台ロール本数を入力する。 */
public final class Stage2AladdinTodayExcludeNextDayDispatchDialog {

    private static final Stage2NextDayRollDispatchDialogSupport.Theme THEME =
            new Stage2NextDayRollDispatchDialogSupport.Theme(
                    "アラジン当日配台 — 翌日の配台量",
                    "アラジン加工計画で当日に配台がある行について、翌日に配台するロール数を指定してください。"
                            + " 0 の行は翌日に配台しません。",
                    "配台計画手動修正タブと同様、配台ロール単位 (m) の整数倍で指定します。"
                            + " 初期値は残量からアラジン当日計画分を差し引いたロール本数です。"
                            + " 入力値は翌日の配台上限であり、設備能力などにより実際の配台量は少なくなる場合があります。"
                            + " OK で未確定の入力も反映します。",
                    "実加工",
                    "翌日配台(ロール)",
                    "-fx-background-color: #E3F2FD;",
                    "-fx-font-size: 11px; -fx-text-fill: #1565C0;",
                    true);

    private Stage2AladdinTodayExcludeNextDayDispatchDialog() {}

    public static final class Row implements Stage2NextDayRollDispatchDialogSupport.RowModel {
        private final String taskId;
        private final String process;
        private final String machineName;
        private final double aladdinTodayM;
        private final double remainingM;
        private final Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo;
        private final javafx.beans.property.SimpleStringProperty nextDayRollCount =
                new javafx.beans.property.SimpleStringProperty();

        public Row(
                String taskId,
                String process,
                String machineName,
                double aladdinTodayM,
                double remainingM,
                Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo) {
            this.taskId = taskId != null ? taskId : "";
            this.process = process != null ? process : "";
            this.machineName = machineName != null ? machineName : "";
            this.aladdinTodayM = aladdinTodayM;
            this.remainingM = remainingM;
            this.unitInfo =
                    unitInfo != null
                            ? unitInfo
                            : new Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM(
                                    0.0, 0.0, 0.0, false);
            int defaultRolls =
                    Stage2InProgressNextDayRollInput.defaultNextDayRollCountAfterAladdinToday(
                            aladdinTodayM, remainingM, this.unitInfo.unitM());
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

        public double aladdinTodayM() {
            return aladdinTodayM;
        }

        @Override
        public double referenceM() {
            return 0.0;
        }

        @Override
        public double aladdinTodayPlanM() {
            return aladdinTodayM;
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
            return "アラジン当日";
        }

        Stage2AladdinTodayExcludeNextDayDispatchIo.Entry toEntryFromNextDayInput() {
            int nextDayRolls =
                    Stage2InProgressNextDayRollInput.parseNonNegativeRollCount(
                                    nextDayRollCount.get())
                            .orElse(0);
            double excludeM =
                    Stage2InProgressNextDayRollInput
                            .resolveExcludedMetersFromNextDayRollCount(
                                    nextDayRolls, remainingM, unitM())
                            .orElse(0.0);
            return new Stage2AladdinTodayExcludeNextDayDispatchIo.Entry(
                    taskId, process, machineName, excludeM);
        }
    }

    /** @return 確定時は各行の入力値。キャンセル時は empty。 */
    public static Optional<List<Stage2AladdinTodayExcludeNextDayDispatchIo.Entry>> prompt(
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
