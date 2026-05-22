package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.Instant;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;

/**
 * サマリ {@link jp.co.pm.ai.desktop.config.AppPaths#SUMMARY_AI_DISPATCH_XLSX} のバックグラウンド出力を、
 * 同一ブックを参照する複数 PC 間で共有するロックファイル（{@code <ブック名>.export.lock}）で示す。
 */
public final class SummaryAiDispatchExportLock {

    public static final String LOCK_SUFFIX = ".export.lock";

    private static final String KEY_VERSION = "version";
    private static final String KEY_WORKBOOK = "workbook";
    private static final String KEY_HOST = "host";
    private static final String KEY_PID = "pid";
    private static final String KEY_USER = "user";
    private static final String KEY_STARTED_AT = "startedAt";

    private SummaryAiDispatchExportLock() {}

    public record LockInfo(
            String workbook,
            String host,
            long pid,
            String user,
            Instant startedAt) {

        public String displayHost() {
            return host != null && !host.isBlank() ? host : "（不明）";
        }
    }

    /** 取得に成功したロック。{@link #release()} で必ず解放する。 */
    public static final class AcquiredLock implements AutoCloseable {
        private final Path lockPath;
        private volatile boolean released;

        private AcquiredLock(Path lockPath) {
            this.lockPath = lockPath;
        }

        public Path lockPath() {
            return lockPath;
        }

        public void release() {
            if (released) {
                return;
            }
            released = true;
            SummaryAiDispatchExportLock.deleteIfExistsQuiet(lockPath);
        }

        @Override
        public void close() {
            release();
        }
    }

    public static Path lockFilePath(Path summaryWorkbookPath) {
        Path wb = summaryWorkbookPath.toAbsolutePath().normalize();
        String name = wb.getFileName().toString() + LOCK_SUFFIX;
        Path parent = wb.getParent();
        return parent != null ? parent.resolve(name) : Path.of(name);
    }

    /**
     * サマリ更新中かどうかの正本判定。{@code <ブック名>.export.lock} が通常ファイルとして存在するときのみ true。
     */
    public static boolean isLocked(Path summaryWorkbookPath) {
        return Files.isRegularFile(lockFilePath(summaryWorkbookPath));
    }

    public static Optional<LockInfo> readLockInfo(Path summaryWorkbookPath) {
        Path lock = lockFilePath(summaryWorkbookPath);
        if (!Files.isRegularFile(lock)) {
            return Optional.empty();
        }
        try {
            String text = Files.readString(lock, StandardCharsets.UTF_8);
            return Optional.of(parseLockPayload(text, summaryWorkbookPath));
        } catch (Exception ignored) {
            return Optional.of(
                    new LockInfo(
                            summaryWorkbookPath.toString(),
                            "",
                            0L,
                            "",
                            Instant.EPOCH));
        }
    }

    /**
     * ロックを取得する。既に存在する場合は empty（他端末または前回クラッシュの残り）。
     */
    public static Optional<AcquiredLock> tryAcquire(Path summaryWorkbookPath) throws IOException {
        Path lock = lockFilePath(summaryWorkbookPath);
        Path parent = lock.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        String payload = formatLockPayload(summaryWorkbookPath);
        try {
            Files.writeString(
                    lock,
                    payload,
                    StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE_NEW,
                    StandardOpenOption.WRITE);
            return Optional.of(new AcquiredLock(lock));
        } catch (java.nio.file.FileAlreadyExistsException ex) {
            return Optional.empty();
        }
    }

    /** 管理者向け: ロックファイルを問答なしで削除する（出力中の他 PC がある場合は競合の恐れあり）。 */
    public static boolean forceRemove(Path summaryWorkbookPath) {
        return deleteIfExistsQuiet(lockFilePath(summaryWorkbookPath));
    }

    private static boolean deleteIfExistsQuiet(Path lock) {
        try {
            return Files.deleteIfExists(lock);
        } catch (IOException ignored) {
            return false;
        }
    }

    private static String formatLockPayload(Path summaryWorkbookPath) {
        Map<String, String> m = new LinkedHashMap<>();
        m.put(KEY_VERSION, "1");
        m.put(KEY_WORKBOOK, summaryWorkbookPath.toAbsolutePath().normalize().toString());
        m.put(KEY_HOST, hostName());
        m.put(KEY_PID, Long.toString(ProcessHandle.current().pid()));
        m.put(KEY_USER, System.getProperty("user.name", ""));
        m.put(KEY_STARTED_AT, Instant.now().toString());
        StringBuilder sb = new StringBuilder();
        for (Map.Entry<String, String> e : m.entrySet()) {
            sb.append(e.getKey()).append('=').append(escapeValue(e.getValue())).append('\n');
        }
        return sb.toString();
    }

    private static LockInfo parseLockPayload(String text, Path fallbackWorkbook) {
        Map<String, String> m = new LinkedHashMap<>();
        for (String line : text.split("\n")) {
            String t = line.strip();
            if (t.isEmpty() || t.startsWith("#")) {
                continue;
            }
            int eq = t.indexOf('=');
            if (eq <= 0) {
                continue;
            }
            m.put(t.substring(0, eq).strip(), unescapeValue(t.substring(eq + 1)));
        }
        long pid = 0L;
        try {
            pid = Long.parseLong(m.getOrDefault(KEY_PID, "0"));
        } catch (NumberFormatException ignored) {
            // ignore
        }
        Instant started = Instant.EPOCH;
        String startedRaw = m.get(KEY_STARTED_AT);
        if (startedRaw != null && !startedRaw.isBlank()) {
            try {
                started = Instant.parse(startedRaw);
            } catch (Exception ignored) {
                // ignore
            }
        }
        return new LockInfo(
                m.getOrDefault(KEY_WORKBOOK, fallbackWorkbook.toString()),
                m.getOrDefault(KEY_HOST, ""),
                pid,
                m.getOrDefault(KEY_USER, ""),
                started);
    }

    private static String escapeValue(String v) {
        if (v == null) {
            return "";
        }
        return v.replace("\\", "\\\\").replace("\n", "\\n").replace("\r", "\\r");
    }

    private static String unescapeValue(String v) {
        if (v == null || v.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder(v.length());
        for (int i = 0; i < v.length(); i++) {
            char c = v.charAt(i);
            if (c == '\\' && i + 1 < v.length()) {
                char n = v.charAt(i + 1);
                if (n == 'n') {
                    sb.append('\n');
                    i++;
                    continue;
                }
                if (n == 'r') {
                    sb.append('\r');
                    i++;
                    continue;
                }
                if (n == '\\') {
                    sb.append('\\');
                    i++;
                    continue;
                }
            }
            sb.append(c);
        }
        return sb.toString();
    }

    private static String hostName() {
        try {
            return java.net.InetAddress.getLocalHost().getHostName();
        } catch (Exception ex) {
            String env = System.getenv("COMPUTERNAME");
            if (env != null && !env.isBlank()) {
                return env;
            }
            env = System.getenv("HOSTNAME");
            return env != null ? env : "";
        }
    }
}
