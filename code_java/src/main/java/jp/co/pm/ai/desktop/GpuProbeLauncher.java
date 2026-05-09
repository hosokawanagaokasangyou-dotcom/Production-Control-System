package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.management.ManagementFactory;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.TimeUnit;

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
                process.destroyForcibly();
                PrismGpuBootstrapStatus.recordSoftwareAfterProbe("GPU テストタイムアウト");
                return false;
            }
            int code = process.exitValue();
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

    private static void drainStream(InputStream in) {
        Thread t =
                new Thread(
                        () -> {
                            try (in) {
                                in.transferTo(OutputStream.nullOutputStream());
                            } catch (IOException ignored) {
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
