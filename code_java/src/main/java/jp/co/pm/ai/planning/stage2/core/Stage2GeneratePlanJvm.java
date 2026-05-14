package jp.co.pm.ai.planning.stage2.core;

import java.util.List;

import jp.co.pm.ai.planning.stage2.Stage2RunContext;
import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;

/**
 * Python {@code _generate_plan_impl} の JVM 移植用オーケストレータ（段階実装）。
 *
 * <p>現段階では {@link Stage2DispatchLoop} 相当の直後に置ける処理のみを集約する（日内割付・タイムライン・成果物出力は未移植）。
 */
public final class Stage2GeneratePlanJvm {

    private Stage2GeneratePlanJvm() {}

    /**
     * {@code _generate_plan_impl} 内の task_queue 構築直後に相当する後段のうち、JVM で再現済みのフェーズのみ実行する。
     */
    public static List<Stage2QueuedTask> runAfterTaskQueueBuilt(
            Stage2InputSnapshot snap, List<Stage2QueuedTask> taskQueue, Stage2RunContext ctx) {
        ctx.log("[stage2-java] generate_plan(JVM): plan_input=" + snap.planInputPath());
        ctx.log("[stage2-java] generate_plan(JVM): 正本メイン処理 " + Stage2PlanningCorePyBands.GENERATE_PLAN_IMPL);
        ctx.log(
                "[stage2-java] generate_plan(JVM): 試行順確定 "
                        + Stage2PlanningCorePyBands.APPLY_DISPATCH_TRIAL_ORDER_FOR_GENERATE_PLAN);
        return Stage2DispatchTrialOrderApplier.apply(taskQueue, ctx);
    }
}
