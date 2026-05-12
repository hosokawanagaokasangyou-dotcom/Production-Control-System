package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.time.Instant;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/**
 * Append startup diagnostics when stdout/stderr are invisible (e.g. GUI-only exe).
 *
 * <p>Writes the same line to:
 *
 * <ul>
 *   <li>{@code ~/.pm-ai-desktop/startup.log}
 *   <li>{@code java.io.tmpdir/pm-ai-desktop-startup.log}
 *   <li>{@code user.dir/pm-ai-desktop-startup.log}
 * </ul>
 *
 * <p>Javadoc and comments are ASCII-only so javac never fails on broken source encoding on Windows.
 */
public final class StartupCrashLog {

    private StartupCrashLog() {}

    /** Install default uncaught-exception handler that appends to log files. */
    public static void installUncaughtExceptionLogging() {
        Thread.setDefaultUncaughtExceptionHandler(
                (thread, ex) -> {
                    appendThrowable("uncaught thread=" + thread.getName(), ex);
                    ex.printStackTrace(System.err);
                });
    }

    /** Primary path shown to users; {@link #append(String)} may also mirror elsewhere. */
    public static Path logPathForUserHint() {
        return Paths.get(System.getProperty("user.home", "."), ".pm-ai-desktop", "startup.log");
    }

    /** Append one line to every writable target path. */
    public static void append(String message) {
        String line =
                "[" + Instant.now() + "] " + message + System.lineSeparator();
        for (Path logFile : distinctLogTargets()) {
            try {
                Path parent = logFile.getParent();
                if (parent != null) {
                    Files.createDirectories(parent);
                }
                Files.writeString(
                        logFile,
                        line,
                        StandardCharsets.UTF_8,
                        StandardOpenOption.CREATE,
                        StandardOpenOption.APPEND);
            } catch (IOException | RuntimeException ignored) {
                // try next path
            }
        }
    }

    private static List<Path> distinctLogTargets() {
        Set<Path> set = new LinkedHashSet<>();
        String home = System.getProperty("user.home");
        if (home != null && !home.isBlank()) {
            set.add(Paths.get(home, ".pm-ai-desktop", "startup.log"));
        }
        String tmp = System.getProperty("java.io.tmpdir");
        if (tmp != null && !tmp.isBlank()) {
            set.add(Paths.get(tmp, "pm-ai-desktop-startup.log"));
        }
        try {
            set.add(
                    Paths.get(System.getProperty("user.dir", "."))
                            .toAbsolutePath()
                            .normalize()
                            .resolve("pm-ai-desktop-startup.log"));
        } catch (RuntimeException ignored) {
            // ignore
        }
        return new ArrayList<>(set);
    }

    public static void appendThrowable(String phase, Throwable t) {
        try {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            t.printStackTrace(pw);
            pw.flush();
            append(phase + ": " + sw.toString().trim());
        } catch (Exception ignored) {
            append(phase + ": " + t.getClass().getName() + ": " + t.getMessage());
        }
    }
}
