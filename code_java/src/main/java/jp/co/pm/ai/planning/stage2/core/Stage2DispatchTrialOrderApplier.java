package jp.co.pm.ai.planning.stage2.core;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Optional;

import jp.co.pm.ai.planning.stage2.Stage2RunContext;

/**
 * Python {@code _apply_dispatch_trial_order_for_generate_plan} の JVM 側部分移植。
 *
 * <p><b>実装済み</b>: キュー全要素がシート「配台試行順番」に数値を持つとき、シート値＋計画行順でソートし effective を確定する（Python
 * 21695–21708 相当）。
 *
 * <p><b>未移植</b>（ログで明示）: {@code _generate_plan_task_queue_sort_key} によるマスタ・納期・need 加重ソート、
 * {@code _reorder_task_queue_b2_ec_inspection_consecutive} 等の §B 隣接調整、連番付与の完全互換。欠損がある場合は暫定として
 * {@code excelRowIndex} 昇順に並べ 1..n を付与する。
 */
public final class Stage2DispatchTrialOrderApplier {

    private Stage2DispatchTrialOrderApplier() {}

    public static List<Stage2QueuedTask> apply(List<Stage2QueuedTask> in, Stage2RunContext ctx) {
        if (in == null || in.isEmpty()) {
            return List.of();
        }
        ArrayList<Stage2QueuedTask> work = new ArrayList<>(in);
        boolean allFromSheet =
                work.stream().allMatch(t -> t.dispatchTrialOrderFromSheet().isPresent());
        if (allFromSheet) {
            work.sort(
                    Comparator.comparingInt((Stage2QueuedTask t) -> t.dispatchTrialOrderFromSheet().orElse(0))
                            .thenComparingInt(Stage2QueuedTask::excelRowIndex));
            ArrayList<Stage2QueuedTask> out = new ArrayList<>(work.size());
            for (Stage2QueuedTask t : work) {
                int v = t.dispatchTrialOrderFromSheet().orElse(0);
                out.add(
                        new Stage2QueuedTask(
                                t.excelRowIndex(), t.requestId(), t.dispatchTrialOrderFromSheet(), v));
            }
            ctx.log(
                    "[stage2-java] 配台試行順番: 「配台試行順番」列の値のまま使用しました（全 "
                            + out.size()
                            + " 件）。正本: "
                            + Stage2PlanningCorePyBands.APPLY_DISPATCH_TRIAL_ORDER_FOR_GENERATE_PLAN);
            return List.copyOf(out);
        }
        work.sort(Comparator.comparingInt(Stage2QueuedTask::excelRowIndex));
        ArrayList<Stage2QueuedTask> out = new ArrayList<>(work.size());
        int seq = 1;
        for (Stage2QueuedTask t : work) {
            out.add(
                    new Stage2QueuedTask(
                            t.excelRowIndex(), t.requestId(), t.dispatchTrialOrderFromSheet(), seq++));
        }
        ctx.log(
                "[stage2-java] 配台試行順番: 自動計算（暫定: 計画入力行順で 1.."
                        + out.size()
                        + "。マスタ・need 加重・§B-2/3 隣接調整は未移植）。正本: "
                        + Stage2PlanningCorePyBands.APPLY_DISPATCH_TRIAL_ORDER_FOR_GENERATE_PLAN);
        return List.copyOf(out);
    }

    /** 計画入力のセルから試行順（シート列）を解釈する。 */
    static Optional<Integer> parseDispatchTrialOrderFromSheet(String raw) {
        if (raw == null) {
            return Optional.empty();
        }
        String s = raw.strip();
        if (s.isEmpty()) {
            return Optional.empty();
        }
        try {
            double d = Double.parseDouble(s.replace(',', '.'));
            if (Double.isNaN(d) || Double.isInfinite(d)) {
                return Optional.empty();
            }
            return Optional.of((int) Math.round(d));
        } catch (NumberFormatException e) {
            return Optional.empty();
        }
    }
}
