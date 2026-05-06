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
 * コンソール無し exe 起動時に標準エラーが見えないため、起動診断を複数パスへミラー追記する。
 *
 * <ul>
 *   <li>優先: {@code ~/.pm-ai-desktop/startup.log}
 *   <li>ミラー: {@code java.io.tmpdir/pm-ai-desktop-startup.log}（権限・ホーム異常時の退避）
 *   <li>ミラー: {@code user.dir/pm-ai-desktop-startup.log}（配布フォルダ直下。{@code user.dir} 合わせ後に有用）
 * </ul>
 */
public final class StartupCrashLog {

    private StartupCrashLog() {}

    /** 未捕捉例外をログファイルへ残す（JavaFX 以外のスレッド向け）。 */
    public static void installUncaughtExceptionLogging() {
        Thread.setDefaultUncaughtExceptionHandler(
                (thread, ex) -> {
                    appendThrowable("uncaught thread=" + thread.getName(), ex);
                    ex.printStackTrace(System.err);
                });
    }

    /**
     * 既定の案内用パス（ホーム配下）。実際のログは {@link #append(String)} が複製先にも書く。
     */
    public static Path logPathForUserHint() {
        return Paths.get(System.getProperty("user.home", "."), ".pm-ai-desktop", "startup.log");
    }

    /** ユーザ／TMP／カレントへ同一行を書く。どれか一つでも成功すればよいが、可能ならすべてに追記する。 */
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
                /* 次の候補へ */
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
            /* ignore */
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
