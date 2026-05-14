package jp.co.pm.ai.planning.stage2.bridge;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;
import java.util.Map;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.planning.stage2.Stage2ExitCodes;
import jp.co.pm.ai.planning.stage2.Stage2RunContext;

/**
 * {@code PM_AI_STAGE2_ENGINE=java} のとき、配台本体を Python 子で実行する経路。
 *
 * <p>{@link jp.co.pm.ai.planning.stage2.dispatch.Stage2DispatchOrchestrator} から呼ばれる条件は次のいずれか。
 *
 * <ul>
 *   <li>{@link AppPaths#KEY_PM_AI_DISPATCH_ENGINE} が {@code python}（明示）
 *   <li>{@link AppPaths#KEY_PM_AI_STAGE2_JAVA_DELEGATE_PYTHON_DISPATCH} が真（従来）
 * </ul>
 *
 * <p>{@code PM_AI_DISPATCH_ENGINE=java} のときは本クラスは使われず、JVM 内配台コア経路へ進む。
 */
public final class Stage2DelegatePythonDispatch {

    private Stage2DelegatePythonDispatch() {}

    public static int runBlocking(Stage2RunContext ctx) {
        try {
            Map<String, String> ui = ctx.uiEnv();
            String dir = nz(ui.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR));
            if (dir.isEmpty()) {
                ctx.log("[stage2-java] PM_AI_CODE_PYTHON_DIR が空のため Python 子を起動できません。");
                return Stage2ExitCodes.GENERAL_FAILURE;
            }
            Path scriptDir = Path.of(dir).toAbsolutePath().normalize();
            if (!Files.isDirectory(scriptDir)) {
                ctx.log("[stage2-java] PM_AI_CODE_PYTHON_DIR がディレクトリでありません: " + scriptDir);
                return Stage2ExitCodes.GENERAL_FAILURE;
            }
            Path pyExe = resolvePythonExecutable(ui);
            RunRequest req = new RunRequest(pyExe, scriptDir, "plan_simulation_stage2.py", "", ui);
            ctx.log("[stage2-java] Python 子: exe=" + pyExe + " cwd=" + scriptDir);
            return PythonProcessRunner.runBlockingSameThread(req, ctx::log);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            ctx.log("[stage2-java] Python 子プロセス中断: " + e);
            return Stage2ExitCodes.GENERAL_FAILURE;
        } catch (IOException e) {
            String m = e.getMessage() != null ? e.getMessage() : e.toString();
            ctx.log("[stage2-java] Python 子プロセス I/O 失敗: " + m);
            return Stage2ExitCodes.GENERAL_FAILURE;
        }
    }

    private static Path resolvePythonExecutable(Map<String, String> ui) {
        String n = AppPaths.normalizePmAiPythonExecutable(ui != null ? ui.get(AppPaths.KEY_PM_AI_PYTHON) : null);
        if (n != null && !n.isBlank()) {
            return Path.of(n);
        }
        boolean win = System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("win");
        return Path.of(win ? "python" : "python3");
    }

    private static String nz(String s) {
        return s != null ? s.strip() : "";
    }
}
