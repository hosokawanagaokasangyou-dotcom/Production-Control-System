package jp.co.pm.ai.planning.stage2.core;

import java.util.List;

import jp.co.pm.ai.planning.stage2.Stage2RunContext;
import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;

/**
 * 日内割付・試行順パス等のコアループ（Python {@code _trial_order_first_schedule_pass} / 日次ループの足場）。
 *
 * <p>現段階では {@link Stage2GeneratePlanJvm} に委譲し、試行順確定までを実行する。日内割付本体は未移植。
 */
public interface Stage2DispatchLoop {

    /**
     * @param taskQueue ミュータブルなリストで渡してもよいが、戻り値を正として後段に渡すこと。
     */
    List<Stage2QueuedTask> run(Stage2InputSnapshot snap, List<Stage2QueuedTask> taskQueue, Stage2RunContext ctx);

    enum Default implements Stage2DispatchLoop {
        INSTANCE;

        @Override
        public List<Stage2QueuedTask> run(
                Stage2InputSnapshot snap, List<Stage2QueuedTask> taskQueue, Stage2RunContext ctx) {
            ctx.log(
                    "[stage2-java] dispatch_loop: "
                            + Stage2PlanningCorePyBands.TRIAL_ORDER_FIRST_SCHEDULE_PASS
                            + " 前段 — 試行順確定のみ。日次割付・タイムラインは未移植。");
            return Stage2GeneratePlanJvm.runAfterTaskQueueBuilt(snap, taskQueue, ctx);
        }
    }
}
