package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.Optional;

import jp.co.pm.ai.desktop.dispatch.Stage2InProgressNextDayRollInput;
import jp.co.pm.ai.planning.stage2.Stage2AladdinTodayExcludeNextDayDispatchIo;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;

/**
 * 段階2直前: アラジン当日配台がある行について、翌日配台から除外するロール本数を入力する。
 */
public final class Stage2AladdinTodayExcludeNextDayDispatchDialog {

    private static final Stage2NextDayRollDispatchDialogSupport.Theme THEME =
            new Stage2NextDayRollDispatchDialogSupport.Theme(
                    "アラジン当日配台 — 翌日からの除外量",
                    "アラジン加工計画で当日に配台がある行について、翌日配台から除外するロール数を指定してください。"
                            + " 実加工数>0 の加工途中行は、ラジオで「①」または「③」を選ぶと別ダイアログで設定します。"
                            + " 0 の行は除外しません。",
                    "配台計画手動修正タブと同様、配台ロール単位 (m) の整数倍で指定します。"
                            + " 初期値はアラジン当日量以内の最大ロール本数です。"
                            + " 除外量は 0 以上・残量およびアラジン当日量に収まるロール整数倍のみ。OK で未確定の入力も反映します。",
                    "アラジン当日",
                    "翌日除外(ロール)",
                    "-fx-background-color: #E3F2FD;",
                    "-fx-font-size: 11px; -fx-text-fill: #1565C0;");

    private Stage2AladdinTodayExcludeNextDayDispatchDialog() {}

    public static final class Row implements Stage2NextDayRollDispatchDialogSupport.RowModel {
        private final String taskId;
        private final String process;
        private final String machineName;
        private final double aladdinTodayM;
        private final double remainingM;
        private final Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo;
        private final javafx.beans.property.SimpleStringProperty excludeRollCount =
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
                    Stage2InProgressNextDayRollInput.defaultRollCountForCap(
                            aladdinTodayM, remainingM, this.unitInfo.unitM());
            this.excludeRollCount.set(String.valueOf(defaultRolls));
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
            return Stage2InProgressNextDayRollInput.maxRollsForCap(
                    aladdinTodayM, remainingM, unitM());
        }

        @Override
        public double effectiveCapM() {
            return Math.min(Math.max(0.0, aladdinTodayM), Math.max(0.0, remainingM));
        }

        @Override
        public javafx.beans.property.StringProperty rollCountProperty() {
            return excludeRollCount;
        }

        Stage2AladdinTodayExcludeNextDayDispatchIo.Entry toEntry(double excludeM) {
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
                r -> {
                    Row row = (Row) r;
                    int rolls =
                            Stage2InProgressNextDayRollInput.parseNonNegativeRollCount(
                                            row.excludeRollCount.get())
                                    .orElse(0);
                    double exclude =
                            Stage2InProgressNextDayRollInput.resolveMetersForCap(
                                            rolls, row.aladdinTodayM(), row.remainingM(), row.unitM())
                                    .orElse(0.0);
                    return row.toEntry(Math.max(0.0, exclude));
                },
                r -> {
                    Row row = (Row) r;
                    return Stage2InProgressNextDayRollInput.validateExcludeRollInput(
                            r.rollCountProperty().get(),
                            row.aladdinTodayM(),
                            r.remainingM(),
                            r.unitInfo());
                });
    }
}
