package jp.co.pm.ai.desktop.io;

import java.awt.Desktop;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.TimeUnit;

/**
 * Opens a file with the OS default application ({@link Desktop#open}) with {@code xdg-open} fallback on
 * Linux when desktop integration fails (e.g. some WSL setups).
 */
public final class DesktopFileOpener {

    private DesktopFileOpener() {}

    public static void openFile(Path path) throws IOException {
        if (path == null || !Files.isRegularFile(path)) {
            throw new IOException("not a regular file: " + path);
        }
        Path abs = path.toAbsolutePath().normalize();
        if (!Desktop.isDesktopSupported()) {
            if (isLinux()) {
                tryXdgOpen(abs);
            } else {
                throw new IOException("Desktop API not supported on this platform");
            }
            return;
        }
        Desktop d = Desktop.getDesktop();
        if (!d.isSupported(Desktop.Action.OPEN)) {
            tryXdgOpen(abs);
            return;
        }
        try {
            d.open(abs.toFile());
        } catch (IOException e) {
            if (isLinux()) {
                tryXdgOpen(abs);
            } else {
                throw e;
            }
        }
    }

    private static boolean isLinux() {
        return System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("linux");
    }

    private static void tryXdgOpen(Path abs) throws IOException {
        if (!isLinux()) {
            throw new IOException("xdg-open fallback is Linux-only");
        }
        List<String> cmd = new ArrayList<>();
        cmd.add("xdg-open");
        cmd.add(abs.toString());
        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.redirectErrorStream(true);
        try {
            Process p = pb.start();
            boolean finished = p.waitFor(30, TimeUnit.SECONDS);
            if (!finished) {
                p.destroyForcibly();
                throw new IOException("xdg-open timed out");
            }
            if (p.exitValue() != 0) {
                throw new IOException("xdg-open exit=" + p.exitValue());
            }
        } catch (IOException ex) {
            throw ex;
        } catch (InterruptedException ie) {
            Thread.currentThread().interrupt();
            throw new IOException(ie);
        }
    }
}
