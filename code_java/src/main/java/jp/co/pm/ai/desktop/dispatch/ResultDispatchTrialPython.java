package jp.co.pm.ai.desktop.dispatch;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;
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
        java.util.ArrayList<String> cmd = new java.util.ArrayList<>();
        cmd.add(pythonExe.toString());
        cmd.add(script.toAbsolutePath().toString());
        cmd.add(jsonPath.toAbsolutePath().toString());
        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.directory(pythonScriptDir.toFile());
        pb.redirectErrorStream(true);
        PythonProcessRunner.mergeUiEnvIntoProcess(pb, extraUiEnv, pythonScriptDir);
        Process p = pb.start();
        String mergedOut = "";
        String lastLine = "";
        try (BufferedReader br =
                new BufferedReader(new InputStreamReader(p.getInputStream(), StandardCharsets.UTF_8))) {
            String line;
            String last = "";
            StringBuilder merged = new StringBuilder();
            while ((line = br.readLine()) != null) {
                if (logLine != null) {
                    logLine.accept(line);
                }
                if (merged.length() > 0) {
                    merged.append('\n');
                }
                merged.append(line);
                last = line;
            }
            mergedOut = merged.toString();
            lastLine = last != null ? last.trim() : "";
        }
        boolean finished = p.waitFor(600, TimeUnit.SECONDS);
        if (!finished) {
            p.destroyForcibly();
            throw new IllegalStateException("dispatch_interactive_trial.py timeout");
        }
        if (p.exitValue() != 0) {
            String hint = trialFailureHint(mergedOut, pythonExe);
            throw new IllegalStateException(
                    "dispatch_interactive_trial.py exit "
                            + p.exitValue()
                            + ": "
                            + mergedOut
                            + hint);
        }
        return lastLine;
    }

    /**
     * Appends short remediation hints for common embed-site-packages issues (stderr is merged into {@code
     * out}).
     */
    private static String trialFailureHint(String mergedOut, Path pythonExe) {
        if (mergedOut == null) {
            return "";
        }
        String lower = mergedOut.toLowerCase(Locale.ROOT);
        if (lower.contains("modulenotfounderror") || lower.contains("no module named")) {
            return "\n\n[hint] Python 環境に依存パッケージが不足している可能性があります。"
                    + " 環境タブの PM_AI_PYTHON_EXE と planning_core が import できる site-packages を確認してください。";
        }
        if (lower.contains("permission denied") || lower.contains("access is denied")) {
            return "\n\n[hint] JSON/xlsx の書き込み権限またはファイルロックを確認してください。";
        }
        if (pythonExe != null && mergedOut.contains(pythonExe.toString())) {
            return "\n\n[hint] 指定 Python: " + pythonExe;
        }
        return "";
    }
}
