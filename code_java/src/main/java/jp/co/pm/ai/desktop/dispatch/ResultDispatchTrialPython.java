package jp.co.pm.ai.desktop.dispatch;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.charset.StandardCharsets;
import java.util.Map;
import java.util.concurrent.TimeUnit;
import java.util.function.Consumer;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;

/**
 * Runs {@code dispatch_interactive_trial.py} (writes shortages JSON, updates {@code 結果_配台表.json}
 * and exports xlsx).
 */
public final class ResultDispatchTrialPython {

    private ResultDispatchTrialPython() {}

    /**
     * @return path to shortages JSON (stdout line) or empty
     */
    public static String runTrial(Path jsonPath, Path pythonExe, Path pythonScriptDir) throws Exception {
        return runTrial(jsonPath, pythonExe, pythonScriptDir, null, null);
    }

    /**
     * Runs the trial script, invoking {@code logLine} for each stdout/stderr line as it arrives (when
     * non-null). Merges {@code extraUiEnv} like {@link PythonProcessRunner} (UTF-8, {@code PM_AI_*} rules,
     * {@code PM_AI_PLAN_INPUT_PATH} from the desktop UI).
     *
     * @return path to shortages JSON (stdout line) or empty
     */
    public static String runTrial(
            Path jsonPath,
            Path pythonExe,
            Path pythonScriptDir,
            Map<String, String> extraUiEnv,
            Consumer<String> logLine)
            throws Exception {
        Path script = pythonScriptDir.resolve("dispatch_interactive_trial.py");
        if (!Files.isRegularFile(script)) {
            throw new IllegalStateException("missing dispatch_interactive_trial.py in " + pythonScriptDir);
        }
        ProcessBuilder pb =
                new ProcessBuilder(
                        pythonExe.toString(),
                        script.toAbsolutePath().toString(),
                        jsonPath.toAbsolutePath().toString());
        pb.directory(pythonScriptDir.toFile());
        pb.redirectErrorStream(true);
        PythonProcessRunner.mergeUiEnvIntoProcess(pb, extraUiEnv, pythonScriptDir);
        Process p = pb.start();
        String out;
        try (BufferedReader br =
                new BufferedReader(new InputStreamReader(p.getInputStream(), StandardCharsets.UTF_8))) {
            String line;
            String last = "";
            while ((line = br.readLine()) != null) {
                if (logLine != null) {
                    logLine.accept(line);
                }
                last = line;
            }
            // 先頭の案内行は捨て、最終行（不足情報 JSON パス）だけを戻り値にする
            out = last != null ? last.trim() : "";
        }
        boolean finished = p.waitFor(600, TimeUnit.SECONDS);
        if (!finished) {
            p.destroyForcibly();
            throw new IllegalStateException("dispatch_interactive_trial.py timeout");
        }
        if (p.exitValue() != 0) {
            throw new IllegalStateException(
                    "dispatch_interactive_trial.py exit " + p.exitValue() + ": " + out);
        }
        return out;
    }
}
