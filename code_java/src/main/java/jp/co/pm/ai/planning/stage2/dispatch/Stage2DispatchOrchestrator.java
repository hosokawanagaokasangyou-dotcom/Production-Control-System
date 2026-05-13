package jp.co.pm.ai.planning.stage2.dispatch;

import jp.co.pm.ai.planning.stage2.Stage2RunContext;
import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;

/** 配台本体のオーケストレーション（段階的にサブモジュールへ分割）。 */
public final class Stage2DispatchOrchestrator {

    private Stage2DispatchOrchestrator() {}

    public static int run(Stage2RunContext ctx, Stage2InputSnapshot snap, java.nio.file.Path outputDir)
            throws Exception {
        ctx.log(
                "[stage2-java] dispatch: Stage2PassThroughPlanner（Python _generate_plan_impl 相当の完全置換は未完了。入力行のプレースホルダ出力）");
        return Stage2PassThroughPlanner.run(ctx, snap, outputDir);
    }
}
