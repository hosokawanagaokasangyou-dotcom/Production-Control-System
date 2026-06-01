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
 * Runs {@code build_stage3_input.py} synchronously: decomposes the manually corrected
 * {@code 結果_配台表.json} into stage-3 branch tasks and writes the input3 sheet
 * ({@code 配台計画_タスク入力3.0}) into the plan-input workbook.
 */
public final class Stage3InputBuilderPython {

    private Stage3InputBuilderPython() {}

    /**
     * @return last stdout line (JSON result of the builder).
     */
    public static String buildInput3Sheet(
            Path resultDispatchJsonPath,
            Path planInputWorkbookPath,
            Path pythonExe,
            Path pythonScriptDir,
            Map<String, String> extraUiEnv,
            Consumer<String> logLine)
            throws Exception {
        Path script = pythonScriptDir.resolve("build_stage3_input.py");
        if (!Files.isRegularFile(script)) {
            throw new IllegalStateException("missing build_stage3_input.py in " + pythonScriptDir);
        }
        java.util.ArrayList<String> cmd = new java.util.ArrayList<>();
        cmd.add(pythonExe.toString());
        cmd.add(script.toAbsolutePath().toString());
        cmd.add(resultDispatchJsonPath.toAbsolutePath().toString());
        cmd.add(planInputWorkbookPath.toAbsolutePath().toString());
        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.directory(pythonScriptDir.toFile());
        pb.redirectErrorStream(true);
        Map<String, String> env = new java.util.HashMap<>();
        if (extraUiEnv != null) {
            env.putAll(extraUiEnv);
        }
        PythonProcessRunner.mergeUiEnvIntoProcess(pb, env, pythonScriptDir);
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
        boolean finished = p.waitFor(300, TimeUnit.SECONDS);
        if (!finished) {
            p.destroyForcibly();
            throw new IllegalStateException("build_stage3_input.py timeout");
        }
        if (p.exitValue() != 0) {
            throw new IllegalStateException(
                    "build_stage3_input.py exit " + p.exitValue() + ": " + mergedOut);
        }
        return lastLine;
    }
}
