package jp.co.pm.ai.desktop.bridge;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.Executor;
import java.util.concurrent.Executors;
import java.util.function.Consumer;

import javafx.application.Platform;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * Runs {@code task_extract_stage1.py} / {@code plan_simulation_stage2.py} with
 * {@link ProcessBuilder} (plan: UTF-8 child env from UI).
 *
 * <p>Environment merge: JVM inherits the OS process environment for PATH etc.; application-defined keys
 * come from the UI {@code extraEnv}. Inherited {@link AppPaths#KEY_PM_AI_PLAN_WORKBOOK_JSON} /
 * {@link AppPaths#KEY_PM_AI_MEMBER_SCHEDULE_JSON} are removed unless present in {@code extraEnv}, so OS-level
 * {@code 0} does not suppress stage-2 same-name JSON mirrors when the user did not set those keys in the env tab.
 * Then {@code PYTHONUTF8} and {@code PYTHONIOENCODING}. Use the env tab for {@code PM_AI_PLAN_INPUT_PATH} and other
 * keys; legacy workbook env keys are stripped by the shell controller.
 *
 * <p>Stages 1/2 run as plain Python child processes (openpyxl / pandas, etc.). xlwings is used only on
 * Excel-driven entry points (add-in or legacy macro workflows); it is not required for this JavaFX app.
 */
public final class PythonProcessRunner {

    private static final Executor READER_POOL = Executors.newCachedThreadPool(r -> {
        Thread t = new Thread(r, "python-process-reader");
        t.setDaemon(true);
        return t;
    });

    private PythonProcessRunner() {}

    /**
     * Stage-2 JSON mirrors default to on in Python when unset. Os.environ may still carry {@code 0} from the user or
     * system profile; strip unless the JavaFX env tab explicitly passes the key.
     */
    private static void stripInheritedStage2JsonMirrorEnvUnlessInUi(
            Map<String, String> env, Map<String, String> extraEnv) {
        Map<String, String> ui = extraEnv != null ? extraEnv : Map.of();
        if (!ui.containsKey(AppPaths.KEY_PM_AI_PLAN_WORKBOOK_JSON)) {
            env.remove(AppPaths.KEY_PM_AI_PLAN_WORKBOOK_JSON);
        }
        if (!ui.containsKey(AppPaths.KEY_PM_AI_MEMBER_SCHEDULE_JSON)) {
            env.remove(AppPaths.KEY_PM_AI_MEMBER_SCHEDULE_JSON);
        }
    }

    /**
     * @param taskInputWorkbook Reserved; not copied into the child environment.
     */
    public record RunRequest(
            Path pythonExecutable,
            Path scriptDirectory,
            String scriptFileName,
            String taskInputWorkbook,
            Map<String, String> extraEnv) {}

    /**
     * Same environment and working directory as {@link #runAsync}, but returns the full merged stdout
     * (no per-line {@code [child] } prefix, no {@link Platform#runLater} on each line). For short JSON
     * probe scripts (e.g. {@code pm_ai_actuals_status.py}).
     */
    public static CompletableFuture<CapturedResult> runCaptureAsync(RunRequest req) {
        Objects.requireNonNull(req, "req");
        if (!Files.isDirectory(req.scriptDirectory)) {
            return CompletableFuture.failedFuture(
                    new IOException("script directory not found: " + req.scriptDirectory));
        }
        Path script = req.scriptDirectory.resolve(req.scriptFileName);
        if (!Files.isRegularFile(script)) {
            return CompletableFuture.failedFuture(
                    new IOException("script not found: " + script));
        }
        return CompletableFuture.supplyAsync(
                () -> {
                    try {
                        List<String> cmd = new ArrayList<>();
                        cmd.add(req.pythonExecutable.toString());
                        cmd.add(req.scriptFileName);
                        ProcessBuilder pb = new ProcessBuilder(cmd);
                        pb.directory(req.scriptDirectory.toFile());
                        pb.redirectErrorStream(true);
                        Map<String, String> env = new HashMap<>(pb.environment());
                        stripInheritedStage2JsonMirrorEnvUnlessInUi(env, req.extraEnv);
                        if (req.extraEnv != null) {
                            for (var e : req.extraEnv.entrySet()) {
                                String k = e.getKey();
                                if (k == null || k.isBlank()) {
                                    continue;
                                }
                                env.put(k.trim(), e.getValue() != null ? e.getValue() : "");
                            }
                        }
                        env.put("PYTHONUTF8", "1");
                        env.put("PYTHONIOENCODING", "utf-8");
                        pb.environment().putAll(env);
                        Process p = pb.start();
                        String out = readStreamFully(p.getInputStream());
                        int code = p.waitFor();
                        return new CapturedResult(code, out);
                    } catch (Exception e) {
                        throw new RuntimeException(e);
                    }
                },
                READER_POOL);
    }

    public record CapturedResult(int exitCode, String stdout) {}

    private static String readStreamFully(InputStream in) throws IOException {
        try (BufferedReader r = new BufferedReader(
                new InputStreamReader(in, StandardCharsets.UTF_8))) {
            StringBuilder sb = new StringBuilder();
            String line;
            while ((line = r.readLine()) != null) {
                sb.append(line).append('\n');
            }
            return sb.toString();
        }
    }

    public static CompletableFuture<Integer> runAsync(
            RunRequest req, Consumer<String> lineConsumer, Consumer<Throwable> onError) {
        Objects.requireNonNull(req, "req");
        if (!Files.isDirectory(req.scriptDirectory)) {
            return CompletableFuture.failedFuture(
                    new IOException("script directory not found: " + req.scriptDirectory));
        }
        Path script = req.scriptDirectory.resolve(req.scriptFileName);
        if (!Files.isRegularFile(script)) {
            return CompletableFuture.failedFuture(
                    new IOException("script not found: " + script));
        }
        List<String> cmd = new ArrayList<>();
        cmd.add(req.pythonExecutable.toString());
        cmd.add(req.scriptFileName);

        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.directory(req.scriptDirectory.toFile());
        pb.redirectErrorStream(true);
        Map<String, String> env = new HashMap<>(pb.environment());
        stripInheritedStage2JsonMirrorEnvUnlessInUi(env, req.extraEnv);
        if (req.extraEnv != null) {
            for (var e : req.extraEnv.entrySet()) {
                String k = e.getKey();
                if (k == null || k.isBlank()) {
                    continue;
                }
                env.put(k.trim(), e.getValue() != null ? e.getValue() : "");
            }
        }
        env.put("PYTHONUTF8", "1");
        env.put("PYTHONIOENCODING", "utf-8");
        pb.environment().putAll(env);

        Consumer<String> safeOut = lineConsumer != null ? lineConsumer : s -> {};
        return CompletableFuture.supplyAsync(
                () -> {
                    try {
                        Process p = pb.start();
                        readStreamBlocking(p.getInputStream(), line -> safeOut.accept("[child] " + line));
                        return p.waitFor();
                    } catch (Exception e) {
                        if (onError != null) {
                            Platform.runLater(() -> onError.accept(e));
                        }
                        throw new RuntimeException(e);
                    }
                },
                READER_POOL);
    }

    /** Blocks caller thread until EOF; forwards each line to JavaFX via {@link Platform#runLater}. */
    private static void readStreamBlocking(InputStream in, Consumer<String> perLine) {
        try (BufferedReader r = new BufferedReader(
                new InputStreamReader(in, StandardCharsets.UTF_8))) {
            String line;
            while ((line = r.readLine()) != null) {
                final String l = line;
                Platform.runLater(() -> perLine.accept(l));
            }
        } catch (IOException e) {
            Platform.runLater(() -> perLine.accept("[io-error] " + e.getMessage()));
        }
    }
}
