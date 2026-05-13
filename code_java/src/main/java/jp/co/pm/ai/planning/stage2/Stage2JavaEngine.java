package jp.co.pm.ai.planning.stage2;

import java.io.IOException;
import java.nio.file.NoSuchFileException;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.planning.stage2.dispatch.Stage2DispatchOrchestrator;
import jp.co.pm.ai.planning.stage2.input.Stage2InputLoader;

/** JVM 内の段階2エンジン（Python 子プロセスを使わない）。 */
public final class Stage2JavaEngine {

    private Stage2JavaEngine() {}

    public static int run(Stage2RunContext ctx) {
        try {
            ctx.log("[stage2-java] Stage2JavaEngine 開始");
            Stage2InputSnapshot snap = Stage2InputLoader.load(ctx);
            Stage2InputLoader.logSummary(ctx, snap);
            if (snap.memberDisplayNames().isEmpty()) {
                ctx.log(
                        "[stage2-java] メンバー0件のため中断（Python の _generate_plan_impl 早期 return に相当）。成果物は出しません。");
                return Stage2ExitCodes.OK;
            }
            var outDir = AppPaths.defaultPlanningOutputDir(ctx.uiEnv());
            return Stage2DispatchOrchestrator.run(ctx, snap, outDir);
        } catch (NoSuchFileException e) {
            ctx.log("[stage2-java] 入力ファイルがありません: " + e.getMessage());
            return Stage2ExitCodes.FILE_NOT_FOUND;
        } catch (IOException e) {
            String m = e.getMessage() != null ? e.getMessage() : e.toString();
            if (m.contains("見つかりません") || m.contains("not found") || m.contains("not a file")) {
                ctx.log("[stage2-java] " + m);
                return Stage2ExitCodes.FILE_NOT_FOUND;
            }
            ctx.log("[stage2-java] I/O エラー: " + m);
            return Stage2ExitCodes.GENERAL_FAILURE;
        } catch (Exception e) {
            ctx.log("[stage2-java] 未処理例外: " + (e.getMessage() != null ? e.getMessage() : e));
            return Stage2ExitCodes.GENERAL_FAILURE;
        }
    }
}
