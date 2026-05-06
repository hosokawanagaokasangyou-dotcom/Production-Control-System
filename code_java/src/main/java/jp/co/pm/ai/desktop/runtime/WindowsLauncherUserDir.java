package jp.co.pm.ai.desktop.runtime;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;
import java.util.Optional;

/**
 * When started from jpackage {@code PmAiDesktop.exe} on Windows, {@code user.dir} may not match the
 * install folder. Portable paths (e.g. pm-ai-data) then fail. If the process command line ends with
 * {@code PmAiDesktop.exe}, set {@code user.dir} to that exe's parent directory.
 *
 * <p>Javadoc is ASCII-only so javac never fails on broken multi-byte source encoding on Windows.
 *
 * @see jp.co.pm.ai.desktop.config.AppPaths
 */
public final class WindowsLauncherUserDir {

    private static final String LAUNCHER_EXE = "PmAiDesktop.exe";

    private WindowsLauncherUserDir() {}

    /**
     * Overwrites {@code user.dir} only on Windows when the launcher exe name matches. No-op for IDE,
     * {@code java -jar}, or plain {@code java.exe}.
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
            // keep default user.dir
        }
    }
}
