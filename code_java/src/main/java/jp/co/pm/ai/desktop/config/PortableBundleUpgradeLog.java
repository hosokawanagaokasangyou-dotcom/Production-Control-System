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
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.stream.Stream;

/**
 * ポータブル配布のバージョンアップ処理ログを UTF-8 テキストファイルへ追記する。主な出力先は {@code
 * pm-ai-data/code/log/portable-version-upgrade-yyyyMMdd-HHmmss.log}（同期除外の {@code code/log} 配下）。
 */
public final class PortableBundleUpgradeLog {

    /** {@code pm-ai-data} 直下からのログディレクトリ（相対）。 */
    public static final String LOG_DIR_UNDER_PM_AI_DATA = "code/log";

    private static final String FILE_PREFIX = "portable-version-upgrade-";

    private static final String FILE_SUFFIX = ".log";

    private static final DateTimeFormatter FILE_STAMP =
            DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss").withZone(ZoneId.systemDefault());

    private static final int MAX_RETAINED_LOG_FILES = 40;

    private final Path logFile;
    private final Object writeLock = new Object();
    private boolean closed;

    private PortableBundleUpgradeLog(Path logFile) {
        this.logFile = logFile;
    }

    /**
     * 新規ログファイルを作成し、ヘッダを書き込む。失敗時は {@link IOException}。
     *
     * @param installRoot {@code PMD.exe} の {@code user.dir}（絶対パス推奨）
     * @param pmAiDataRoot ローカル {@code pm-ai-data}
     */
    public static PortableBundleUpgradeLog open(Path installRoot, Path pmAiDataRoot) throws IOException {
        Objects.requireNonNull(installRoot, "installRoot");
        Objects.requireNonNull(pmAiDataRoot, "pmAiDataRoot");
        Path logDir = resolveLogDirectory(installRoot, pmAiDataRoot);
        Files.createDirectories(logDir);
        String stamp = FILE_STAMP.format(Instant.now());
        Path file = logDir.resolve(FILE_PREFIX + stamp + FILE_SUFFIX);
        PortableBundleUpgradeLog session = new PortableBundleUpgradeLog(file);
        session.appendLine("=== ポータルバージョンアップ ログ開始 ===");
        session.appendLine("installRoot=" + installRoot.toAbsolutePath().normalize());
        session.appendLine("pmAiDataRoot=" + pmAiDataRoot.toAbsolutePath().normalize());
        session.appendLine("logFile=" + file.toAbsolutePath().normalize());
        pruneOldLogs(logDir);
        return session;
    }

    /** このセッションのログファイル（絶対パス）。 */
    public Path logFile() {
        return logFile.toAbsolutePath().normalize();
    }

    /** 1 行追記（先頭に ISO-8601 タイムスタンプ）。スレッドセーフ。 */
    public void appendLine(String message) {
        if (message == null || message.isBlank() || closed) {
            return;
        }
        String line =
                "[" + Instant.now() + "] " + message.strip() + System.lineSeparator();
        synchronized (writeLock) {
            if (closed) {
                return;
            }
            try {
                Files.writeString(
                        logFile,
                        line,
                        StandardCharsets.UTF_8,
                        StandardOpenOption.CREATE,
                        StandardOpenOption.APPEND);
            } catch (IOException ignored) {
                /* UI ログは継続；ファイルのみ失敗 */
            }
        }
    }

    /** 例外のスタックトレースを追記。 */
    public void appendThrowable(String phase, Throwable t) {
        if (t == null) {
            return;
        }
        try {
            StringWriter sw = new StringWriter();
            t.printStackTrace(new PrintWriter(sw));
            appendLine(phase + ": " + sw.toString().trim());
        } catch (RuntimeException ex) {
            appendLine(phase + ": " + t.getClass().getName() + ": " + t.getMessage());
        }
    }

    /**
     * 終了フッタを書き、以降の {@link #appendLine} を無効化する。
     *
     * @param success 正常終了なら {@code true}
     * @param summary 1 行サマリ
     */
    public void close(boolean success, String summary) {
        synchronized (writeLock) {
            if (closed) {
                return;
            }
            closed = true;
            String footer =
                    "=== ポータルバージョンアップ "
                            + (success ? "正常終了" : "失敗")
                            + (summary != null && !summary.isBlank() ? " — " + summary.strip() : "")
                            + " ===";
            try {
                Files.writeString(
                        logFile,
                        "[" + Instant.now() + "] " + footer + System.lineSeparator(),
                        StandardCharsets.UTF_8,
                        StandardOpenOption.CREATE,
                        StandardOpenOption.APPEND);
            } catch (IOException ignored) {
                /* best-effort */
            }
        }
    }

    /**
     * ログ出力ディレクトリ。{@code pm-ai-data/code/log} を優先し、作成できないときは {@code
     * ~/.pm-ai-desktop/portable-upgrade-logs}。
     */
    static Path resolveLogDirectory(Path installRoot, Path pmAiDataRoot) {
        Path primary =
                pmAiDataRoot.toAbsolutePath().normalize().resolve(LOG_DIR_UNDER_PM_AI_DATA);
        try {
            Files.createDirectories(primary);
            if (Files.isWritable(primary)) {
                return primary;
            }
        } catch (IOException | SecurityException ignored) {
            /* fallback */
        }
        String home = System.getProperty("user.home", ".");
        return Paths.get(home, ".pm-ai-desktop", "portable-upgrade-logs");
    }

    static void pruneOldLogs(Path logDir) {
        try (Stream<Path> stream = Files.list(logDir)) {
            List<Path> logs = new ArrayList<>();
            stream.filter(Files::isRegularFile)
                    .filter(p -> isUpgradeLogFileName(p.getFileName().toString()))
                    .forEach(logs::add);
            if (logs.size() <= MAX_RETAINED_LOG_FILES) {
                return;
            }
            logs.sort(Comparator.comparingLong(PortableBundleUpgradeLog::lastModifiedSafe).reversed());
            for (int i = MAX_RETAINED_LOG_FILES; i < logs.size(); i++) {
                Files.deleteIfExists(logs.get(i));
            }
        } catch (IOException | SecurityException ignored) {
            /* best-effort */
        }
    }

    private static boolean isUpgradeLogFileName(String name) {
        if (name == null) {
            return false;
        }
        String lower = name.toLowerCase(Locale.ROOT);
        return lower.startsWith(FILE_PREFIX) && lower.endsWith(FILE_SUFFIX);
    }

    private static long lastModifiedSafe(Path p) {
        try {
            return Files.getLastModifiedTime(p).toMillis();
        } catch (IOException e) {
            return 0L;
        }
    }
}
