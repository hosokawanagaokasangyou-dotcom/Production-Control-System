package jp.co.pm.ai.desktop.runtime;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;
import java.util.Optional;

/**
 * When started from jpackage {@code *.exe} on Windows, {@code user.dir} may not match the install
 * folder (e.g. shortcut with another「作業フォルダ」). Portable paths then fail. Align {@code user.dir}
 * to the directory that contains {@code app\} and {@code runtime\}.
 *
 * <p>Javadoc is ASCII-only so javac never fails on broken multi-byte source encoding on Windows.
 *
 * @see jp.co.pm.ai.desktop.config.AppPaths
 */
public final class WindowsLauncherUserDir {

    private static final String PMD_LAUNCHER_EXE = "PMD.exe";
    private static final String RDP_DESKTOP_LAUNCHER_EXE = "PmAiRpaLuncher.exe";

    private WindowsLauncherUserDir() {}

    /**
     * Overwrites {@code user.dir} on Windows when a jpackage portable layout is detected. No-op for
     * IDE, {@code java -jar}, or plain {@code java.exe}.
     */
    public static void alignWithPackagedLauncherIfWindows() {
        String os = System.getProperty("os.name", "").toLowerCase(Locale.ROOT);
        if (!os.contains("windows")) {
            return;
        }
        try {
            if (alignFromJpackageAppPathProperty()) {
                return;
            }
            alignFromCurrentProcessLauncherExe();
        } catch (Throwable ignored) {
            // keep default user.dir
        }
    }

    private static boolean alignFromJpackageAppPathProperty() {
        String raw = System.getProperty("jpackage.app-path");
        if (raw == null || raw.isBlank()) {
            return false;
        }
        Path appPath = Path.of(raw.strip()).toAbsolutePath().normalize();
        Path leaf = appPath.getFileName();
        Path installRoot = null;
        if (leaf != null && "app".equalsIgnoreCase(leaf.toString())) {
            installRoot = appPath.getParent();
        } else if (isPortableBundleInstallRoot(appPath)) {
            installRoot = appPath;
        }
        if (installRoot != null && isPortableBundleInstallRoot(installRoot)) {
            System.setProperty("user.dir", installRoot.toAbsolutePath().normalize().toString());
            return true;
        }
        return false;
    }

    private static void alignFromCurrentProcessLauncherExe() {
        Optional<String> cmd = ProcessHandle.current().info().command();
        if (cmd.isEmpty()) {
            return;
        }
        Path exe = Path.of(cmd.get());
        if (!Files.isRegularFile(exe)) {
            return;
        }
        Path base = exe.getFileName();
        if (base == null || !isKnownPackagedLauncherExe(base.toString())) {
            return;
        }
        Path dir = exe.getParent();
        if (dir != null && Files.isDirectory(dir) && isPortableBundleInstallRoot(dir)) {
            System.setProperty("user.dir", dir.toAbsolutePath().normalize().toString());
        }
    }

    private static boolean isKnownPackagedLauncherExe(String fileName) {
        return PMD_LAUNCHER_EXE.equalsIgnoreCase(fileName)
                || RDP_DESKTOP_LAUNCHER_EXE.equalsIgnoreCase(fileName);
    }

    static boolean isPortableBundleInstallRoot(Path dir) {
        if (dir == null) {
            return false;
        }
        Path abs = dir.toAbsolutePath().normalize();
        return Files.isDirectory(abs.resolve("app")) && Files.isDirectory(abs.resolve("runtime"));
    }
}
