package jp.co.pm.ai.desktop.bridge;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.regex.Pattern;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.Executor;
import java.util.concurrent.Executors;
import java.util.function.Consumer;

import javafx.application.Platform;

/**
 * Runs {@code task_extract_stage1.py} / {@code plan_simulation_stage2.py} with
 * {@link ProcessBuilder} (plan: UTF-8 child env from UI).
 *
 * <p>Environment merge: JVM inherits the OS process environment for PATH etc. Keys whose names start with
 * {@code PM_AI_} are <strong>not</strong> taken from the OS: any inherited {@code PM_AI_*} is removed unless the
 * same key is present in {@code extraEnv} (JavaFX environment-variable tab + {@code childEnvForPython} overlays).
 * Python therefore does not see OS-level {@code PM_AI_*} unless the user explicitly set them in the app. Then
 * {@code PYTHONUTF8} and {@code PYTHONIOENCODING}. Legacy workbook env keys are stripped by the shell controller
 * ({@code PM_AI_SKIP_WORKBOOK_ENV_SHEET}).
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

    private static boolean isPmAiEnvKey(String k) {
        return k != null && k.length() >= 6 && k.regionMatches(true, 0, "PM_AI_", 0, 6);
    }

    /**
     * Do not let OS-inherited {@code PM_AI_*} affect Python: remove unless {@code extraEnv} explicitly sets that key
     * (case-insensitive key match on Windows).
     */
    private static void stripInheritedPmAiEnvUnlessInUi(Map<String, String> env, Map<String, String> extraEnv) {
        Set<String> allowedLower = new HashSet<>();
        if (extraEnv != null) {
            for (String k : extraEnv.keySet()) {
                if (k != null && !k.isBlank()) {
                    allowedLower.add(k.trim().toLowerCase(Locale.ROOT));
                }
            }
        }
        for (String k : new ArrayList<>(env.keySet())) {
            if (!isPmAiEnvKey(k)) {
                continue;
            }
            String kl = k.toLowerCase(Locale.ROOT);
            if (!allowedLower.contains(kl)) {
                env.remove(k);
            }
        }
    }

    /**
     * Same rules as stage1/2: strip OS-inherited {@code PM_AI_*} unless set in {@code extraEnv}, merge UI keys,
     * then force UTF-8 stdio and unbuffered Python output (Windows-safe Japanese logging).
     */
    public static void mergeUiEnvIntoProcess(ProcessBuilder pb, Map<String, String> extraEnv) {
        mergeUiEnvIntoProcess(pb, extraEnv, null);
    }

    /**
     * @param codePythonDir {@code PM_AI_CODE_PYTHON_DIR} 相当（{@code planning_core} の親ディレクトリ）。指定時は
     *     {@code PYTHONPATH} の先頭へ追加し、{@code No module named 'planning_core'} を防ぐ。
     */
    public static void mergeUiEnvIntoProcess(
            ProcessBuilder pb, Map<String, String> extraEnv, Path codePythonDir) {
        Map<String, String> env = new HashMap<>(pb.environment());
        stripInheritedPmAiEnvUnlessInUi(env, extraEnv);
        if (extraEnv != null) {
            for (var e : extraEnv.entrySet()) {
                String k = e.getKey();
                if (k == null || k.isBlank()) {
                    continue;
                }
                env.put(k.trim(), e.getValue() != null ? e.getValue() : "");
            }
        }
        env.put("PYTHONUTF8", "1");
        env.put("PYTHONIOENCODING", "utf-8");
        env.put("PYTHONUNBUFFERED", "1");
        prependCodePythonDirOnPythonPath(env, codePythonDir);
        pb.environment().putAll(env);
    }

    /**
     * {@code planning_core} を import できるよう、子プロセスの {@code PYTHONPATH} 先頭に {@code code/python} を載せる。
     */
    static void prependCodePythonDirOnPythonPath(Map<String, String> env, Path codePythonDir) {
        if (env == null || codePythonDir == null) {
            return;
        }
        Path norm;
        try {
            norm = codePythonDir.toAbsolutePath().normalize();
        } catch (RuntimeException e) {
            return;
        }
        if (!Files.isDirectory(norm)) {
            return;
        }
        String root = norm.toString();
        String sep = File.pathSeparator;
        String cur = env.get("PYTHONPATH");
        if (cur == null || cur.isBlank()) {
            env.put("PYTHONPATH", root);
            return;
        }
        if (cur.equals(root)) {
            return;
        }
        if (cur.startsWith(root + sep)) {
            return;
        }
        for (String seg : cur.split(Pattern.quote(sep))) {
            String s = seg != null ? seg.strip() : "";
            if (s.isEmpty()) {
                continue;
            }
            try {
                Path p = Path.of(s).toAbsolutePath().normalize();
                if (Files.isDirectory(p) && Files.isSameFile(p, norm)) {
                    return;
                }
            } catch (Exception ignored) {
                /* fall through */
            }
        }
        env.put("PYTHONPATH", root + sep + cur);
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
        return runCaptureAsyncWithLineTap(req, null);
    }

    /**
     * Like {@link #runCaptureAsync}, but invokes {@code lineTap} for each stdout line as it is read
     * (worker thread — use {@link Platform#runLater} in the consumer for UI updates). Does not alter the
     * captured stdout text (tap is observation-only).
     */
    public static CompletableFuture<CapturedResult> runCaptureAsyncWithLineTap(
            RunRequest req, Consumer<String> lineTap) {
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
                        mergeUiEnvIntoProcess(pb, req.extraEnv, req.scriptDirectory);
                        Process p = pb.start();
                        String out = readStreamFullyWithTap(p.getInputStream(), lineTap);
                        int code = p.waitFor();
                        return new CapturedResult(code, out);
                    } catch (Exception e) {
                        throw new RuntimeException(e);
                    }
                },
                READER_POOL);
    }

    public record CapturedResult(int exitCode, String stdout) {}

    private static String readStreamFullyWithTap(InputStream in, Consumer<String> lineTap)
            throws IOException {
        try (BufferedReader r = new BufferedReader(
                new InputStreamReader(in, StandardCharsets.UTF_8))) {
            StringBuilder sb = new StringBuilder();
            String line;
            while ((line = r.readLine()) != null) {
                if (lineTap != null) {
                    lineTap.accept(line);
                }
                sb.append(line).append('\n');
            }
            return sb.toString();
        }
    }

    public static CompletableFuture<Integer> runAsync(
            RunRequest req, Consumer<String> lineConsumer, Consumer<Throwable> onError) {
        return runAsync(req, lineConsumer, onError, null);
    }

    /**
     * @param onProcessStarted invoked on the worker thread immediately after {@link ProcessBuilder#start()}
     *     (e.g. to retain the handle for {@link Process#destroyForcibly()} user cancellation).
     */
    public static CompletableFuture<Integer> runAsync(
            RunRequest req,
            Consumer<String> lineConsumer,
            Consumer<Throwable> onError,
            Consumer<Process> onProcessStarted) {
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
        mergeUiEnvIntoProcess(pb, req.extraEnv, req.scriptDirectory);

        Consumer<String> safeOut = lineConsumer != null ? lineConsumer : s -> {};
        return CompletableFuture.supplyAsync(
                () -> {
                    try {
                        Process p = pb.start();
                        if (onProcessStarted != null) {
                            onProcessStarted.accept(p);
                        }
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
