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

/**
 * コンソール無し exe 起動時に標準エラーが見えないため、起動診断を {@code ~/.pm-ai-desktop/startup.log} に追記する。
 */
public final class StartupCrashLog {

    private static final Path LOG =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "startup.log");

    private StartupCrashLog() {}

    /** 未捕捉例外をログファイルへ残す（JavaFX 以外のスレッド向け）。 */
    public static void installUncaughtExceptionLogging() {
        Thread.setDefaultUncaughtExceptionHandler(
                (thread, ex) -> {
                    appendThrowable("uncaught thread=" + thread.getName(), ex);
                    ex.printStackTrace(System.err);
                });
    }

    public static void append(String message) {
        try {
            Files.createDirectories(LOG.getParent());
            String line =
                    "[" + Instant.now() + "] " + message + System.lineSeparator();
            Files.writeString(
                    LOG,
                    line,
                    StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE,
                    StandardOpenOption.APPEND);
        } catch (IOException ignored) {
            /* ログに書けないときは諦める */
        }
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

    public static Path logPathForUserHint() {
        return LOG;
    }
}
