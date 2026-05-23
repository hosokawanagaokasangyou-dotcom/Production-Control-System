package jp.co.pm.ai.desktop.io.actuals;

import java.util.List;
import java.util.Optional;

/** ダッシュボード1機械分の表示モデル。 */
public record EquipmentMachineStatus(
        String machineName,
        Status status,
        Optional<ActualTaskRow> actualTask,
        List<PlanLine> aladdinPlans,
        List<PlanLine> dispatchPlans) {

    public enum Status {
        /** 選択した実績表示日に加工実績が0件。 */
        STOPPED,
        /** 当日: 実績が0より大きく、かつ実績が当日アラジン計画未満。それ以外の日: 完了率100%未満。 */
        RUNNING,
        /** 当日: 実績が0より大きく、かつ実績が当日アラジン計画以上（または完了率100%）。 */
        COMPLETED
    }

    /** 実績表示日の最新加工開始行。 */
    public record ActualTaskRow(
            String requestNo,
            String processName,
            double qtyConvM,
            /** 0〜100 */
            double completionPct,
            String memberRaw,
            String startDateTime,
            String endDateTime) {}

    /** アラジン／配台予定の1行。 */
    public record PlanLine(String requestNo, String processName, String qtyM) {}
}
