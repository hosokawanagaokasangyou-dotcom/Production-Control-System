package jp.co.pm.ai.desktop;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.management.ManagementFactory;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
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
    /** 子プロセス stderr を親側ログに載せる上限（パイプ詰まり防止とログ肥大防止） */
    private static final int STDERR_CAPTURE_MAX_BYTES = 24576;

    private GpuProbeLauncher() {}

    /**
     * @return GPU パイプラインで Canvas プローブが成功したら true
     */
    static boolean runGpuCanvasProbe() {
        // #region agent log
        AgentDebugLog.appendStructured(
                Map.of(),
                "d1d903",
                "H5",
                "GpuProbeLauncher.runGpuCanvasProbe:start",
                "子プロセスGPUプローブ",
                Map.of(
                        "probePrismOrder",
                        prismGpuOrderForProbe(),
                        "osName",
                        System.getProperty("os.name", "")));
        // #endregion
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
        ByteArrayOutputStream errCapture = new ByteArrayOutputStream();
        Thread errDrain = null;
        Thread outDrain = null;
        try {
            process = pb.start();
            outDrain = startDrainToDiscard(process.getInputStream());
            errDrain = startDrainToBoundedBuffer(process.getErrorStream(), errCapture, STDERR_CAPTURE_MAX_BYTES);
            boolean done = process.waitFor(PROBE_TIMEOUT_SEC, TimeUnit.SECONDS);
            if (!done) {
                process.destroyForcibly();
                joinDrainQuiet(outDrain, 5_000);
                joinDrainQuiet(errDrain, 5_000);
                String errTimeout = clipStderrForLog(errCapture);
                // #region agent log
                AgentDebugLog.appendStructured(
                        Map.of(),
                        "d1d903",
                        "H2",
                        "GpuProbeLauncher.runGpuCanvasProbe",
                        "プローブwaitForがタイムアウト",
                        Map.of("timeoutSec", PROBE_TIMEOUT_SEC, "stderrCapture", errTimeout));
                // #endregion
                PrismGpuBootstrapStatus.recordSoftwareAfterProbe("GPU テストタイムアウト");
                return false;
            }
            joinDrainQuiet(outDrain, 5_000);
            joinDrainQuiet(errDrain, 5_000);
            int code = process.exitValue();
            String errFull = clipStderrForLog(errCapture);
            // #region agent log
            AgentDebugLog.appendStructured(
                    Map.of(),
                    "d1d903",
                    "H1",
                    "GpuProbeLauncher.runGpuCanvasProbe",
                    "子プロセス終了",
                    Map.of("exitCode", code, "stderrCapture", errFull));
            // #endregion
            if (code != 0) {
                PrismGpuBootstrapStatus.recordSoftwareAfterProbe("GPU テスト終了コード=" + code);
                return false;
            }
            return true;
        } catch (Exception e) {
            PrismGpuBootstrapStatus.recordSoftwareAfterProbe(
                    "GPU テスト例外: " + e.getClass().getSimpleName() + ": " + e.getMessage());
            return false;
        } finally {
            if (process != null && process.isAlive()) {
                process.destroyForcibly();
            }
        }
    }

    /** NDJSON 用に stderr を上限文字数で切り詰め（改行は維持） */
    private static String clipStderrForLog(ByteArrayOutputStream errCapture) {
        if (errCapture == null) {
            return "";
        }
        String s = errCapture.toString(StandardCharsets.UTF_8).trim();
        int max = 8000;
        if (s.length() <= max) {
            return s;
        }
        return s.substring(0, max) + "…";
    }

    private static void joinDrainQuiet(Thread t, long millis) {
        if (t == null) {
            return;
        }
        try {
            t.join(millis);
        } catch (InterruptedException ie) {
            Thread.currentThread().interrupt();
        }
    }

    private static Thread startDrainToDiscard(InputStream in) {
        Thread thr =
                new Thread(
                        () -> {
                            try (in) {
                                in.transferTo(OutputStream.nullOutputStream());
                            } catch (IOException ignored) {
                                // ignore
                            }
                        },
                        "gpu-probe-drain-out");
        thr.setDaemon(true);
        thr.start();
        return thr;
    }

    private static Thread startDrainToBoundedBuffer(
            InputStream in, ByteArrayOutputStream buf, int maxBytes) {
        Thread thr =
                new Thread(
                        () -> {
                            try (in) {
                                byte[] chunk = new byte[8192];
                                int n;
                                while ((n = in.read(chunk)) != -1 && buf.size() < maxBytes) {
                                    int w = Math.min(n, maxBytes - buf.size());
                                    buf.write(chunk, 0, w);
                                }
                            } catch (IOException ignored) {
                                // ignore
                            }
                        },
                        "gpu-probe-drain-err");
        thr.setDaemon(true);
        thr.start();
        return thr;
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

        for (String arg : ManagementFactory.getRuntimeMXBean().getInputArguments()) {
            if (arg.startsWith("--add-opens")
                    || arg.startsWith("--add-exports")
                    || arg.startsWith("--enable-native-access=")
                    || arg.startsWith("--patch-module")) {
                cmd.add(arg);
            }
        }

        cmd.add("-Dfile.encoding=UTF-8");
        cmd.add("-Dprism.order=" + prismGpuOrderForProbe());
        cmd.add("-classpath");
        cmd.add(cp);
        cmd.add(JavaFxGpuProbeApp.class.getName());

        return cmd;
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
