package jp.co.pm.ai.planning.stage2.dispatch;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

import jp.co.pm.ai.planning.stage2.Stage2RunContext;
import jp.co.pm.ai.planning.stage2.core.Stage2ConstraintDigest;
import jp.co.pm.ai.planning.stage2.core.Stage2DispatchDiagnostics;
import jp.co.pm.ai.planning.stage2.core.Stage2DispatchLoop;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanningCorePyBands;
import jp.co.pm.ai.planning.stage2.core.Stage2QueuedTask;
import jp.co.pm.ai.planning.stage2.core.Stage2TaskQueueBuilder;
import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;

/**
 * JVM 内段階2の統合エントリ: task_queue 構築・制約ダイジェスト・dispatch ループ（足場）の後、現状は成果物を
 * {@link Stage2PassThroughPlanner} に委ねる。配台コア移植が進んだら本クラス内でシート／JSON を埋め、PassThrough を縮退させる。
 */
public final class Stage2JavaDispatchEngine {

    private Stage2JavaDispatchEngine() {}

    public static int run(Stage2RunContext ctx, Stage2InputSnapshot snap, Path outputDir) throws IOException {
        ctx.log("[stage2-java] Stage2JavaDispatchEngine: task_queue / constraint_digest / dispatch_loop（足場）→ 成果物は現状 PassThrough");
        ctx.log("[stage2-java] 正本参照: " + Stage2PlanningCorePyBands.BUILD_TASK_QUEUE_FROM_PLANNING_DF);
        List<Stage2QueuedTask> queue = Stage2TaskQueueBuilder.build(snap);
        ctx.log("[stage2-java] task_queue: size=" + queue.size());
        Stage2DispatchDiagnostics.logMasterProbe(ctx, snap);
        Stage2DispatchDiagnostics.logTaskQueuePreview(ctx, queue, 12);
        Stage2ConstraintDigest digest = Stage2ConstraintDigest.fromSnapshot(snap);
        ctx.log(
                "[stage2-java] constraint_digest: members="
                        + digest.memberCount()
                        + " factory="
                        + digest.factoryStart()
                        + "/"
                        + digest.factoryEnd()
                        + " exclude_rules="
                        + digest.excludeRuleCount()
                        + " plan_rows="
                        + digest.planDataRows()
                        + " sheet="
                        + digest.planSheetResolved());
        Stage2DispatchLoop.Default.INSTANCE.run(snap, queue, ctx);
        return Stage2PassThroughPlanner.run(ctx, snap, outputDir);
    }
}
