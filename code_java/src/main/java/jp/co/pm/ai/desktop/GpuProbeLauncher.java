package jp.co.pm.ai.desktop;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.management.ManagementFactory;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.TimeUnit;
import java.util.regex.Pattern;

/**
 * 現在の JVM と同等の依存解決で {@link JavaFxGpuProbeApp} を子プロセス起動し、GPU Prism が Canvas で動くか判定する。
 *
 * <p>親が {@code jdk.module.path} なしで OpenJFX を {@code java.class.path}（無名モジュール）から読んでいるときは、
 * 子も同一の単一 {@code -classpath} でプローブする。子だけ {@code --module-path} に切り出すと GPU 試験だけ通り、本体で
 * Canvas／RTTexture 系の例外になる偽陽性が出るため。
 */
final class GpuProbeLauncher {

    private static final int PROBE_TIMEOUT_SEC = 45;

    private static final int STDERR_CAPTURE_MAX = 96_000;

    private GpuProbeLauncher() {}

    /**
     * @param mirrorClasspathLikeParent 親が無名 CLASSPATH で JavaFX を読んでいるとき {@code true}。子のコマンドラインを親と
     *     同型にし、GPU 試験の偽陽性を防ぐ。
     * @return GPU パイプラインで Canvas プローブが成功したら true
     */
    static boolean runGpuCanvasProbe(boolean mirrorClasspathLikeParent) {
        List<String> cmd;
        try {
            boolean forceSplitOpenJfx =
                    Boolean.getBoolean("pm.ai.javafx.prism.probeSplitOpenJfx");
            cmd = buildCommand(mirrorClasspathLikeParent && !forceSplitOpenJfx);
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

    private static int jdkFeatureVersion() {
        try {
            return Runtime.version().feature();
        } catch (Throwable ignored) {
            return 21;
        }
    }

    private static List<String> buildCommand(boolean preferClasspathMirror) {
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

        /*
         * JDK 22+ でクラスパス上の JavaFX（無名モジュール）がネイティブを System::load する際、
         * 子プローブだけフラグが無いと「ランタイム不足」扱いで終了することがある（親は Maven が
         * --enable-native-access を付けるが javafx.graphics は無名構成では Unknown module になり得る）。
         */
        if (jdkFeatureVersion() >= 22) {
            cmd.add("--enable-native-access=ALL-UNNAMED");
        }

        appendInheritedJvmOptions(cmd, ManagementFactory.getRuntimeMXBean().getInputArguments());

        cmd.add("-Dfile.encoding=UTF-8");
        cmd.add("-Dprism.order=" + prismGpuOrderForProbe());

        /*
         * 親が module-path 上で JavaFX を名前付きモジュールとして読んでいる場合: 親と同じ java.class.path を子に渡す。
         */
        if (!System.getProperty("jdk.module.path", "").isBlank()) {
            cmd.add("-classpath");
            cmd.add(cp);
            cmd.add(JavaFxGpuProbeApp.class.getName());
            return cmd;
        }

        /*
         * 親が無名 CLASSPATH のときは子も単一 classpath に揃える（OpenJFX を子だけ module-path に載せると
         * 本体との Prism 経路がずれ、GPU 試験が偽陽性になり得る）。
         */
        if (preferClasspathMirror) {
            cmd.add("-classpath");
            cmd.add(cp);
            cmd.add(JavaFxGpuProbeApp.class.getName());
            return cmd;
        }

        /*
         * 親が jdk.module.path 空だが JavaFX が名前付きモジュールとして解決されている等の稀な構成向け:
         * OpenJFX を module-path に載せ、プローブ本体は -classpath 側の無名モジュール（従来の exec:exec 型）。
         */
        LinkedHashSet<String> openJfx = new LinkedHashSet<>();
        LinkedHashSet<String> rest = new LinkedHashSet<>();
        for (String seg : splitClasspathSegments(cp)) {
            String lower = seg.toLowerCase(Locale.ROOT);
            if (lower.contains("openjfx") || lower.contains("javafx-")) {
                openJfx.add(seg);
            } else {
                rest.add(seg);
            }
        }
        if (!openJfx.isEmpty()) {
            cmd.add("--module-path");
            cmd.add(String.join(File.pathSeparator, openJfx));
            cmd.add("--add-modules");
            cmd.add("javafx.controls,javafx.fxml,javafx.graphics,javafx.base,javafx.swing");
            cmd.add("--add-opens=javafx.base/com.sun.javafx.event=ALL-UNNAMED");
            cmd.add("--add-opens=javafx.controls/javafx.scene.control.skin=ALL-UNNAMED");
            cmd.add("--add-exports=javafx.controls/com.sun.javafx.scene.control.behavior=ALL-UNNAMED");
            cmd.add("--enable-native-access=javafx.graphics");
            cmd.add("-classpath");
            cmd.add(String.join(File.pathSeparator, rest));
            cmd.add(JavaFxGpuProbeApp.class.getName());
            return cmd;
        }

        cmd.add("-classpath");
        cmd.add(cp);
        cmd.add(JavaFxGpuProbeApp.class.getName());

        return cmd;
    }

    private static List<String> splitClasspathSegments(String cp) {
        List<String> out = new ArrayList<>();
        for (String s : cp.split(Pattern.quote(File.pathSeparator))) {
            if (s != null && !s.isBlank()) {
                out.add(s);
            }
        }
        return out;
    }

    /**
     * exec:exec@pm-ai-desktop や従来の JavaFX 起動が付ける JavaFX モジュール解決用オプションを子へ継承する。
     *
     * <p>継承しないと子プロセスだけ {@code javafx.application.Application} が解決できず GPU プローブが誤って不合格になる。
     *
     * <p>{@code --add-opens}/{@code --add-exports}/{@code --enable-native-access}/{@code --patch-module} のうち
     * {@code javafx.*} モジュールを対象にするものは継承しない。親がクラスパス上の JavaFX で動いていると子だけ
     * 「Unknown module: javafx.*」となりランタイム不足扱いで終了するため。
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
            if ("--add-opens".equals(arg) || "--add-exports".equals(arg)) {
                if (i + 1 < inputArgs.size()) {
                    String next = inputArgs.get(i + 1);
                    if (!next.startsWith("-")) {
                        if (!referencesJavaFxModuleSpec(next)) {
                            cmd.add(arg);
                            cmd.add(next);
                        }
                        i++;
                    }
                }
                continue;
            }
            if (arg.startsWith("--add-opens=") || arg.startsWith("--add-exports=")) {
                if (!referencesJavaFxModuleSpec(arg)) {
                    cmd.add(arg);
                }
                continue;
            }
            if (arg.startsWith("--enable-native-access=")) {
                if (!referencesJavaFxModuleSpec(arg)) {
                    cmd.add(arg);
                }
                continue;
            }
            if (arg.startsWith("--patch-module")) {
                if (!referencesJavaFxModuleSpec(arg)) {
                    cmd.add(arg);
                }
            }
        }
    }

    /** {@code javafx.} モジュール名を対象にした JVM 引数か（プローブ子では未ロードのため除外する）。 */
    private static boolean referencesJavaFxModuleSpec(String arg) {
        return arg != null && arg.contains("javafx.");
    }

    private static String prismGpuOrderForProbe() {
        String os = System.getProperty("os.name", "").toLowerCase(Locale.ROOT);
        if (os.contains("windows")) {
            // 本体 {@link PmAiFxApp#applyPrismGpuPipelineOrder} と同順（無名モジュール時の Canvas／D3D 不整合回避）
            return "es2,d3d,sw";
        }
        if (os.contains("mac")) {
            return "metal,es2,sw";
        }
        return "es2,sw";
    }
}
