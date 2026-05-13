package jp.co.pm.ai.desktop;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.management.ManagementFactory;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import jp.co.pm.ai.desktop.debug.AgentDebugLog;

/**
 * 現在の JVM と同じクラスパスで {@link JavaFxGpuProbeApp} を子プロセス起動し、GPU Prism が Canvas で動くか判定する。
 */
final class GpuProbeLauncher {

    private static final int PROBE_TIMEOUT_SEC = 45;

    /** Cursor デバッグセッション用 NDJSON（{@code .cursor/debug-afb084.log}）。 */
    private static final String AGENT_DEBUG_SESSION = "afb084";

    private static final int STDERR_CAPTURE_MAX = 96_000;

    private GpuProbeLauncher() {}

    /**
     * @return GPU パイプラインで Canvas プローブが成功したら true
     */
    static boolean runGpuCanvasProbe() {
        List<String> cmd;
        try {
            cmd = buildCommand();
        } catch (RuntimeException e) {
            PrismGpuBootstrapStatus.recordSoftwareAfterProbe("プローブ起動準備失敗: " + e.getMessage());
            return false;
        }
        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.redirectOutput(ProcessBuilder.Redirect.PIPE);
        pb.redirectError(ProcessBuilder.Redirect.PIPE);
        Process process = null;
        ByteArrayOutputStream outBuf = new ByteArrayOutputStream();
        ByteArrayOutputStream errBuf = new ByteArrayOutputStream();
        try {
            process = pb.start();
            Thread tout = drainStreamToBuffer(process.getInputStream(), outBuf, STDERR_CAPTURE_MAX);
            Thread terr = drainStreamToBuffer(process.getErrorStream(), errBuf, STDERR_CAPTURE_MAX);
            boolean done = process.waitFor(PROBE_TIMEOUT_SEC, TimeUnit.SECONDS);
            joinQuietly(tout, 8_000);
            joinQuietly(terr, 8_000);
            if (!done) {
                process.destroyForcibly();
                PrismGpuBootstrapStatus.recordSoftwareAfterProbe("GPU テストタイムアウト");
                // #region agent log
                logGpuProbeNdjson(
                        "H-timeout",
                        "GpuProbeLauncher.runGpuCanvasProbe",
                        "GPU probe timed out",
                        probeLogData(cmd, -1, outBuf, errBuf));
                // #endregion
                return false;
            }
            int code = process.exitValue();
            if (code != 0) {
                PrismGpuBootstrapStatus.recordSoftwareAfterProbe("GPU テスト終了コード=" + code);
                // #region agent log
                logGpuProbeNdjson(
                        "H-child-exit-nonzero",
                        "GpuProbeLauncher.runGpuCanvasProbe",
                        "GPU probe child exited non-zero",
                        probeLogData(cmd, code, outBuf, errBuf));
                // #endregion
                return false;
            }
            // #region agent log
            logGpuProbeNdjson(
                    "H-probe-ok",
                    "GpuProbeLauncher.runGpuCanvasProbe",
                    "GPU probe child exit 0",
                    probeLogData(cmd, 0, outBuf, errBuf));
            // #endregion
            return true;
        } catch (Exception e) {
            PrismGpuBootstrapStatus.recordSoftwareAfterProbe(
                    "GPU テスト例外: " + e.getClass().getSimpleName() + ": " + e.getMessage());
            // #region agent log
            Map<String, Object> d = probeLogData(cmd, null, outBuf, errBuf);
            d.put("parentException", e.getClass().getName());
            d.put("parentExceptionMessage", String.valueOf(e.getMessage()));
            logGpuProbeNdjson(
                    "H-parent-io",
                    "GpuProbeLauncher.runGpuCanvasProbe",
                    "GPU probe parent IOException/Interrupted",
                    d);
            // #endregion
            return false;
        } finally {
            if (process != null && process.isAlive()) {
                process.destroyForcibly();
            }
        }
    }

    private static Thread drainStreamToBuffer(InputStream in, ByteArrayOutputStream dest, int maxBytes) {
        Thread t =
                new Thread(
                        () -> {
                            try (in) {
                                byte[] buf = new byte[8192];
                                int n;
                                while ((n = in.read(buf)) >= 0) {
                                    synchronized (dest) {
                                        int room = maxBytes - dest.size();
                                        if (room <= 0) {
                                            continue;
                                        }
                                        int take = Math.min(n, room);
                                        dest.write(buf, 0, take);
                                    }
                                }
                            } catch (IOException ignored) {
                                // ignore
                            }
                        },
                        "gpu-probe-drain");
        t.setDaemon(true);
        t.start();
        return t;
    }

    private static void joinQuietly(Thread t, long millis) {
        if (t == null) {
            return;
        }
        try {
            t.join(millis);
        } catch (InterruptedException ie) {
            Thread.currentThread().interrupt();
        }
    }

    private static Map<String, Object> probeLogData(
            List<String> cmd, Integer exitCode, ByteArrayOutputStream outBuf, ByteArrayOutputStream errBuf) {
        Map<String, Object> d = new LinkedHashMap<>();
        d.put("exitCode", exitCode);
        d.put("javaHome", System.getProperty("java.home", ""));
        d.put("osName", System.getProperty("os.name", ""));
        d.put("prismOrderInProbeCmd", prismGpuOrderForProbe());
        d.put("childStdoutUtf8", utf8Bounded(outBuf, 24_000));
        d.put("childStderrUtf8", utf8Bounded(errBuf, 48_000));
        d.put("probeJavaExe", cmd.isEmpty() ? "" : cmd.get(0));
        return d;
    }

    private static String utf8Bounded(ByteArrayOutputStream buf, int maxChars) {
        if (buf == null) {
            return "";
        }
        byte[] raw;
        synchronized (buf) {
            raw = buf.toByteArray();
        }
        String s = new String(raw, StandardCharsets.UTF_8);
        if (s.length() <= maxChars) {
            return s;
        }
        return s.substring(0, maxChars) + "\n…(truncated)";
    }

    private static void logGpuProbeNdjson(
            String hypothesisId, String location, String message, Map<String, ?> data) {
        AgentDebugLog.appendStructured(Map.of(), AGENT_DEBUG_SESSION, hypothesisId, location, message, data);
    }

    private static List<String> buildCommand() {
        boolean win =
                System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("windows");
        Path javaExe = Path.of(System.getProperty("java.home"), "bin", win ? "java.exe" : "java");
        if (!java.nio.file.Files.isExecutable(javaExe)) {
            throw new IllegalStateException("java 実行ファイルが見つかりません: " + javaExe);
        }

        String cp = System.getProperty("java.class.path", "");
        if (cp.isBlank()) {
            throw new IllegalStateException("java.class.path が空です");
        }

        List<String> cmd = new ArrayList<>();
        cmd.add(javaExe.toString());

        appendInheritedJvmOptions(cmd, ManagementFactory.getRuntimeMXBean().getInputArguments());

        cmd.add("-Dfile.encoding=UTF-8");
        cmd.add("-Dprism.order=" + prismGpuOrderForProbe());
        cmd.add("-classpath");
        cmd.add(cp);
        cmd.add(JavaFxGpuProbeApp.class.getName());

        return cmd;
    }

    /**
     * exec:exec@pm-ai-desktop や従来の JavaFX 起動が付ける JavaFX モジュール解決用オプションを子へ継承する。
     *
     * <p>継承しないと子プロセスだけ {@code javafx.application.Application} が解決できず GPU プローブが誤って不合格になる。
     */
    private static void appendInheritedJvmOptions(List<String> cmd, List<String> inputArgs) {
        for (int i = 0; i < inputArgs.size(); i++) {
            String arg = inputArgs.get(i);
            if (arg.startsWith("--module-path=")
                    || arg.startsWith("--upgrade-module-path=")
                    || arg.startsWith("--add-modules=")
                    || arg.startsWith("--limit-modules=")) {
                cmd.add(arg);
                continue;
            }
            if ("--module-path".equals(arg)
                    || "--upgrade-module-path".equals(arg)
                    || "--add-modules".equals(arg)
                    || "--limit-modules".equals(arg)) {
                cmd.add(arg);
                if (i + 1 < inputArgs.size()) {
                    String next = inputArgs.get(i + 1);
                    if (!next.startsWith("-")) {
                        cmd.add(next);
                        i++;
                    }
                }
                continue;
            }
            if (arg.startsWith("--add-opens")
                    || arg.startsWith("--add-exports")
                    || arg.startsWith("--enable-native-access=")
                    || arg.startsWith("--patch-module")) {
                cmd.add(arg);
            }
        }
    }

    private static String prismGpuOrderForProbe() {
        String os = System.getProperty("os.name", "").toLowerCase(Locale.ROOT);
        if (os.contains("windows")) {
            return "d3d,es2,sw";
        }
        if (os.contains("mac")) {
            return "metal,es2,sw";
        }
        return "es2,sw";
    }
}
