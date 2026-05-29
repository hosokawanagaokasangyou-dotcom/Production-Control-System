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
import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * Runs {@code plan_simulation_stage2_5_ai.py} (アラジン整列・sidecar・xlsx。学習アーカイブは背景)。
 */
public final class ResultDispatchStage25Python {

    private ResultDispatchStage25Python() {}

    public static String runStage25(
            Path dispatchJsonPath,
            Path stage2RawJsonPath,
            String jobId,
            Path pythonExe,
            Path pythonScriptDir,
            Map<String, String> extraUiEnv,
            Consumer<String> logLine)
            throws Exception {
        Path script = pythonScriptDir.resolve("plan_simulation_stage2_5_ai.py");
        if (!Files.isRegularFile(script)) {
            throw new IllegalStateException("missing plan_simulation_stage2_5_ai.py in " + pythonScriptDir);
        }
        java.util.ArrayList<String> cmd = new java.util.ArrayList<>();
        cmd.add(pythonExe.toString());
        cmd.add(script.toAbsolutePath().toString());
        cmd.add(dispatchJsonPath.toAbsolutePath().toString());
        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.directory(pythonScriptDir.toFile());
        pb.redirectErrorStream(true);
        Map<String, String> env = new java.util.HashMap<>();
        if (extraUiEnv != null) {
            env.putAll(extraUiEnv);
        }
        if (jobId != null && !jobId.isBlank()) {
            env.put(AppPaths.KEY_PM_AI_STAGE2_5_JOB_ID, jobId.trim());
        }
        if (stage2RawJsonPath != null) {
            env.put(
                    AppPaths.KEY_PM_AI_STAGE2_5_STAGE2_RAW_JSON,
                    stage2RawJsonPath.toAbsolutePath().normalize().toString());
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
        boolean finished = p.waitFor(600, TimeUnit.SECONDS);
        if (!finished) {
            p.destroyForcibly();
            throw new IllegalStateException("plan_simulation_stage2_5_ai.py timeout");
        }
        if (p.exitValue() != 0) {
            throw new IllegalStateException(
                    "plan_simulation_stage2_5_ai.py exit " + p.exitValue() + ": " + mergedOut);
        }
        return lastLine;
    }
}
