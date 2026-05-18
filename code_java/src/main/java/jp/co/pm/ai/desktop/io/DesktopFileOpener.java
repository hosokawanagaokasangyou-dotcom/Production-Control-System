package jp.co.pm.ai.desktop.io;

import java.awt.Desktop;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
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
        openFile(path, false);
    }

    /**
     * Excel 等を読み取り専用で開く（サマリ配台ブックの UI「開く」向け）。Windows / WSL では {@code EXCEL.EXE /r}
     * を優先する。
     */
    public static void openFileReadOnly(Path path) throws IOException {
        openFile(path, true);
    }

    private static void openFile(Path path, boolean readOnly) throws IOException {
        Path abs = validateRegularFile(path);
        if (readOnly) {
            if (tryOpenReadOnly(abs)) {
                return;
            }
        }
        openWithDesktopOrXdg(abs);
    }

    private static Path validateRegularFile(Path path) throws IOException {
        if (path == null || !Files.isRegularFile(path)) {
            throw new IOException("not a regular file: " + path);
        }
        return path.toAbsolutePath().normalize();
    }

    private static void openWithDesktopOrXdg(Path abs) throws IOException {
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

    /** @return {@code true} if a read-only launch was started */
    private static boolean tryOpenReadOnly(Path abs) throws IOException {
        if (isWindows()) {
            return tryOpenWindowsExcelReadOnly(abs);
        }
        if (isLinux() && isWsl()) {
            return tryOpenWslExcelReadOnly(abs);
        }
        if (isMac()) {
            return tryOpenMacExcelReadOnly(abs);
        }
        if (isLinux()) {
            return tryOpenLinuxOfficeReadOnly(abs);
        }
        return false;
    }

    private static boolean tryOpenWindowsExcelReadOnly(Path abs) throws IOException {
        Path excel = resolveWindowsExcelExe();
        if (excel == null) {
            return tryOpenWindowsExcelReadOnlyViaPowerShell(abs);
        }
        startDetached(List.of(excel.toString(), "/r", abs.toString()));
        return true;
    }

    private static boolean tryOpenWslExcelReadOnly(Path abs) throws IOException {
        String winPath = wslPathToWindows(abs);
        Path excel = resolveWindowsExcelExe();
        if (excel != null) {
            startDetached(
                    List.of(
                            "cmd.exe",
                            "/c",
                            "start",
                            "",
                            excel.toString(),
                            "/r",
                            winPath));
            return true;
        }
        return tryOpenWindowsExcelReadOnlyViaPowerShellWsl(abs, winPath);
    }

    private static boolean tryOpenWindowsExcelReadOnlyViaPowerShell(Path abs) throws IOException {
        String ps =
                "$p='"
                        + escapePowerShellSingleQuoted(abs.toString())
                        + "';"
                        + "$x=New-Object -ComObject Excel.Application;"
                        + "$x.Visible=$true;"
                        + "$null=$x.Workbooks.Open($p,0,$true)";
        startDetached(List.of("powershell.exe", "-NoProfile", "-Command", ps));
        return true;
    }

    private static boolean tryOpenWindowsExcelReadOnlyViaPowerShellWsl(Path abs, String winPath)
            throws IOException {
        String ps =
                "$p='"
                        + escapePowerShellSingleQuoted(winPath)
                        + "';"
                        + "$x=New-Object -ComObject Excel.Application;"
                        + "$x.Visible=$true;"
                        + "$null=$x.Workbooks.Open($p,0,$true)";
        startDetached(List.of("cmd.exe", "/c", "powershell.exe", "-NoProfile", "-Command", ps));
        return true;
    }

    private static boolean tryOpenMacExcelReadOnly(Path abs) throws IOException {
        String script =
                "tell application \"Microsoft Excel\" to open workbook "
                        + "(POSIX file \""
                        + escapeAppleScript(abs.toString())
                        + "\") read only true";
        startDetached(List.of("osascript", "-e", script));
        return true;
    }

    private static boolean tryOpenLinuxOfficeReadOnly(Path abs) throws IOException {
        for (String bin : List.of("libreoffice", "soffice")) {
            if (commandExists(bin)) {
                startDetached(List.of(bin, "--read-only", abs.toString()));
                return true;
            }
        }
        return false;
    }

    static String wslPathToWindows(Path path) {
        String raw = path.toString().replace('\\', '/');
        if (looksLikeWindowsAbsolute(raw)) {
            return raw.replace('/', '\\');
        }
        String s = path.toAbsolutePath().normalize().toString().replace('\\', '/');
        if (!s.startsWith("/mnt/") || s.length() < 6) {
            return path.toString();
        }
        int nextSlash = s.indexOf('/', 5);
        if (nextSlash < 0) {
            char drive = s.charAt(5);
            return Character.toUpperCase(drive) + ":\\";
        }
        char drive = s.charAt(5);
        String rest = s.substring(nextSlash + 1).replace('/', '\\');
        return Character.toUpperCase(drive) + ":\\" + rest;
    }

    private static boolean looksLikeWindowsAbsolute(String normalizedSlashPath) {
        return normalizedSlashPath.length() >= 3
                && Character.isLetter(normalizedSlashPath.charAt(0))
                && normalizedSlashPath.charAt(1) == ':'
                && normalizedSlashPath.charAt(2) == '/';
    }

    private static Path resolveWindowsExcelExe() {
        String pf = System.getenv("ProgramFiles");
        String pfx86 = System.getenv("ProgramFiles(x86)");
        List<Path> candidates = new ArrayList<>();
        if (pf != null) {
            candidates.add(Path.of(pf, "Microsoft Office", "root", "Office16", "EXCEL.EXE"));
            candidates.add(Path.of(pf, "Microsoft Office", "Office16", "EXCEL.EXE"));
        }
        if (pfx86 != null) {
            candidates.add(Path.of(pfx86, "Microsoft Office", "root", "Office16", "EXCEL.EXE"));
            candidates.add(Path.of(pfx86, "Microsoft Office", "Office16", "EXCEL.EXE"));
        }
        candidates.add(Path.of("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"));
        candidates.add(Path.of("C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE"));
        for (Path p : candidates) {
            if (Files.isRegularFile(p)) {
                return p.toAbsolutePath().normalize();
            }
        }
        return null;
    }

    private static void startDetached(List<String> command) throws IOException {
        ProcessBuilder pb = new ProcessBuilder(command);
        pb.redirectErrorStream(true);
        pb.start();
    }

    private static boolean commandExists(String name) {
        try {
            ProcessBuilder pb = new ProcessBuilder("which", name);
            pb.redirectErrorStream(true);
            Process p = pb.start();
            boolean finished = p.waitFor(5, TimeUnit.SECONDS);
            return finished && p.exitValue() == 0;
        } catch (IOException | InterruptedException e) {
            if (e instanceof InterruptedException) {
                Thread.currentThread().interrupt();
            }
            return false;
        }
    }

    private static boolean isWsl() {
        try {
            String version =
                    Files.readString(Path.of("/proc/version"), StandardCharsets.UTF_8)
                            .toLowerCase(Locale.ROOT);
            return version.contains("microsoft") || version.contains("wsl");
        } catch (IOException e) {
            return false;
        }
    }

    private static boolean isWindows() {
        return System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("windows");
    }

    private static boolean isMac() {
        String os = System.getProperty("os.name", "").toLowerCase(Locale.ROOT);
        return os.contains("mac") || os.contains("darwin");
    }

    private static boolean isLinux() {
        return System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("linux");
    }

    private static String escapePowerShellSingleQuoted(String s) {
        return s.replace("'", "''");
    }

    private static String escapeAppleScript(String s) {
        return s.replace("\\", "\\\\").replace("\"", "\\\"");
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
