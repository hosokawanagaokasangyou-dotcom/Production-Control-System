package jp.co.pm.ai.desktop;

import java.lang.management.ManagementFactory;
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
        try {
            process = pb.start();
            drainStream(process.getInputStream());
            drainStream(process.getErrorStream());
            boolean done = process.waitFor(PROBE_TIMEOUT_SEC, TimeUnit.SECONDS);
            if (!done) {
                // #region agent log
                AgentDebugLog.appendStructured(
                        Map.of(),
                        "d1d903",
                        "H2",
                        "GpuProbeLauncher.runGpuCanvasProbe",
                        "プローブwaitForがタイムアウト",
                        Map.of("timeoutSec", PROBE_TIMEOUT_SEC));
                // #endregion
                process.destroyForcibly();
                PrismGpuBootstrapStatus.recordSoftwareAfterProbe("GPU テストタイムアウト");
                return false;
            }
            int code = process.exitValue();
            // #region agent log
            AgentDebugLog.appendStructured(
                    Map.of(),
                    "d1d903",
                    "H1",
                    "GpuProbeLauncher.runGpuCanvasProbe",
                    "子プロセス終了",
                    Map.of("exitCode", code));
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

    private static void drainStream(java.io.InputStream in) {
        Thread t =
                new Thread(
                        () -> {
                            try (in) {
                                in.transferTo(java.io.OutputStream.nullOutputStream());
                            } catch (java.io.IOException ignored) {
                                // ignore
                            }
                        },
                        "gpu-probe-drain");
        t.setDaemon(true);
        t.start();
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
