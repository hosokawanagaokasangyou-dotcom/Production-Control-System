package jp.co.pm.ai.planning.stage2.dispatch;

import jp.co.pm.ai.planning.stage2.Stage2EnvParsing;
import jp.co.pm.ai.planning.stage2.Stage2RunContext;
import jp.co.pm.ai.planning.stage2.bridge.Stage2DelegatePythonDispatch;
import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;

/** 配台本体のオーケストレーション（段階的にサブモジュールへ分割）。 */
public final class Stage2DispatchOrchestrator {

    private Stage2DispatchOrchestrator() {}

    public static int run(Stage2RunContext ctx, Stage2InputSnapshot snap, java.nio.file.Path outputDir)
            throws Exception {
        if (Stage2EnvParsing.javaDelegatesPythonDispatch(ctx.uiEnv())) {
            ctx.log(
                    "[stage2-java] dispatch: Python _generate_plan_impl（子プロセス正本）— "
                            + "PM_AI_STAGE2_JAVA_DELEGATE_PYTHON_DISPATCH 有効");
            return Stage2DelegatePythonDispatch.runBlocking(ctx);
        }
        ctx.log(
                "[stage2-java] dispatch: Stage2JavaDispatchEngine（task_queue 足場→現状は PassThrough で成果物）");
        return Stage2JavaDispatchEngine.run(ctx, snap, outputDir);
    }
}
