package jp.co.pm.ai.desktop.runtime;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;
import java.util.Optional;

/**
 * Windows で jpackage の {@code PmAiDesktop.exe} から起動したとき、{@code user.dir} が exe と同じフォルダにならない環境がある。
 * {@link jp.co.pm.ai.desktop.config.AppPaths} や {@code pm-ai-data} 解決が崩れ、ネイティブ／クラスパス前提が破綻して無言終了することがあるため、
 * プロセスのコマンドラインが {@code PmAiDesktop.exe} のときだけ {@code user.dir} をその親ディレクトリに合わせる。
 */
public final class WindowsLauncherUserDir {

    private static final String LAUNCHER_EXE = "PmAiDesktop.exe";

    private WindowsLauncherUserDir() {}

    /**
     * Windows かつランチャ exe のときだけ {@code user.dir} を上書きする。IDE / {@code java -jar} / {@code java.exe} では何もしない。
     */
    public static void alignWithPackagedLauncherIfWindows() {
        String os = System.getProperty("os.name", "").toLowerCase(Locale.ROOT);
        if (!os.contains("windows")) {
            return;
        }
        try {
            Optional<String> cmd = ProcessHandle.current().info().command();
            if (cmd.isEmpty()) {
                return;
            }
            Path exe = Path.of(cmd.get());
            if (!Files.isRegularFile(exe)) {
                return;
            }
            Path base = exe.getFileName();
            if (base == null || !LAUNCHER_EXE.equalsIgnoreCase(base.toString())) {
                return;
            }
            Path dir = exe.getParent();
            if (dir != null && Files.isDirectory(dir)) {
                Path abs = dir.toAbsolutePath().normalize();
                System.setProperty("user.dir", abs.toString());
            }
        } catch (Throwable ignored) {
            /* 失敗時は既定の user.dir のまま */
        }
    }
}
