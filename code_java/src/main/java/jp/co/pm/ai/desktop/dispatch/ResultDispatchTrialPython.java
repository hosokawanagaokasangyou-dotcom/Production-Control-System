package jp.co.pm.ai.desktop.dispatch;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
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
        if (mergedOut == null || mergedOut.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        if (mergedOut.contains("No module named 'pydantic'")) {
            sb.append("\n対処: ログ先頭の「Python 実行ファイル」と同じパスに pip install すること（別の python に入れると解消しません）。例: \"")
                    .append(pythonExe)
                    .append("\" -m pip install pydantic\n")
                    .append("またはリポジトリの code/python/requirements.txt を ")
                    .append("\"")
                    .append(pythonExe)
                    .append(
                            "\" -m pip install -r （リポジトリルート）\\\\code\\\\python\\\\requirements.txt で一括。\n")
                    .append("Windows では scripts/pm_ai_embed_pip_install.ps1 も利用可。-PythonExe に上記パスを指定。\n")
                    .append("環境変数タブの PM_AI_PYTHON が別インタープリターを指していないかも確認してください。\n");
        }
        if (mergedOut.contains("pywin32_bootstrap") || mergedOut.contains("pywin32.pth")) {
            sb.append(
                    "pywin32.pth 警告: pywin32 が不完全です。例: \"")
                    .append(pythonExe)
                    .append("\" -m pip install --upgrade --force-reinstall pywin32\n")
                    .append("配台試行のみで COM が不要なら、Lib\\\\site-packages\\\\pywin32.pth をリネームして無効化してもよい場合があります。\n");
        }
        return sb.toString();
    }
}
