package jp.co.pm.ai.planning.stage2.core;

import java.util.List;

import jp.co.pm.ai.planning.stage2.Stage2RunContext;
import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;

/**
 * 日内割付・試行順パス等のコアループ（Python {@code _trial_order_first_schedule_pass} / 日内割付の足場）。
 * 現状は no-op。配台アルゴリズム移植時に実装を差し替える。
 */
public interface Stage2DispatchLoop {

    void run(Stage2InputSnapshot snap, List<Stage2QueuedTask> taskQueue, Stage2RunContext ctx);

    /** ログのみ。Python ログと行番号対応させる際のフック。 */
    enum Default implements Stage2DispatchLoop {
        INSTANCE;

        @Override
        public void run(Stage2InputSnapshot snap, List<Stage2QueuedTask> taskQueue, Stage2RunContext ctx) {
            ctx.log(
                    "[stage2-java] dispatch_loop: placeholder（試行順・日内割付は未移植） task_queue_size="
                            + (taskQueue != null ? taskQueue.size() : 0)
                            + " — 正本: "
                            + Stage2PlanningCorePyBands.TRIAL_ORDER_FIRST_SCHEDULE_PASS);
        }
    }
}
