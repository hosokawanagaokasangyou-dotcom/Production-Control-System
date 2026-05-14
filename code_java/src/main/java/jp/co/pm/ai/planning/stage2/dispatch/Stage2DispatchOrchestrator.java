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
        if (Stage2EnvParsing.dispatchCoreExplicitJava(ctx.uiEnv())) {
            if (Stage2EnvParsing.javaDelegatesPythonDispatch(ctx.uiEnv())) {
                ctx.log(
                        "[stage2-java] dispatch: PM_AI_DISPATCH_ENGINE=java のため "
                                + "PM_AI_STAGE2_JAVA_DELEGATE_PYTHON_DISPATCH は無視し、JVM 内配台コア経路へ進みます。");
            }
            ctx.log(
                    "[stage2-java] dispatch: PM_AI_DISPATCH_ENGINE=java — JVM 内配台コア（移植進行中。Python 正本との切替は本変数と python）");
            return Stage2JavaDispatchEngine.run(ctx, snap, outputDir);
        }
        if (Stage2EnvParsing.dispatchCoreExplicitPython(ctx.uiEnv())
                || Stage2EnvParsing.javaDelegatesPythonDispatch(ctx.uiEnv())) {
            ctx.log(
                    "[stage2-java] dispatch: Python _generate_plan_impl（子プロセス正本）— "
                            + (Stage2EnvParsing.dispatchCoreExplicitPython(ctx.uiEnv())
                                    ? "PM_AI_DISPATCH_ENGINE=python"
                                    : "PM_AI_STAGE2_JAVA_DELEGATE_PYTHON_DISPATCH 有効"));
            return Stage2DelegatePythonDispatch.runBlocking(ctx);
        }
        ctx.log(
                "[stage2-java] dispatch: Stage2JavaDispatchEngine（task_queue 足場→現状は PassThrough で成果物）。"
                        + " Python 配台へ切り替えるときは PM_AI_DISPATCH_ENGINE=python または委譲フラグを有効化。");
        return Stage2JavaDispatchEngine.run(ctx, snap, outputDir);
    }
}
