package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.net.InetAddress;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.Instant;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * 実行時間履歴 JSON（{@link AppPaths#PIPELINE_EXECUTION_TIMING_HISTORY_JSON}）の保存を、
 * 同一共有フォルダを参照する複数 PC 間で排他するロック（{@code <json名>.save.lock}）。
 */
public final class PipelineExecutionTimingHistoryLock {

    public static final String LOCK_SUFFIX = ".save.lock";

    private static final String KEY_VERSION = "version";
    private static final String KEY_HISTORY = "history";
    private static final String KEY_HOST = "host";
    private static final String KEY_HOST_IP = "hostIp";
    private static final String KEY_PID = "pid";
    private static final String KEY_USER = "user";
    private static final String KEY_STARTED_AT = "startedAt";

    private PipelineExecutionTimingHistoryLock() {}

    public record LockInfo(
            String historyPath,
            String host,
            String hostIp,
            long pid,
            String user,
            Instant startedAt) {

        public String displayHost() {
            return host != null && !host.isBlank() ? host : "（不明）";
        }

        public String displayIp() {
            return hostIp != null && !hostIp.isBlank() ? hostIp : "（不明）";
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
            deleteIfExistsQuiet(lockPath);
        }

        @Override
        public void close() {
            release();
        }
    }

    public static Path lockFilePath(Path historyJsonPath) {
        Path json = historyJsonPath.toAbsolutePath().normalize();
        String name = json.getFileName().toString() + LOCK_SUFFIX;
        Path parent = json.getParent();
        return parent != null ? parent.resolve(name) : Path.of(name);
    }

    public static boolean isLocked(Path historyJsonPath) {
        return Files.isRegularFile(lockFilePath(historyJsonPath));
    }

    public static Optional<LockInfo> readLockInfo(Path historyJsonPath) {
        Path lock = lockFilePath(historyJsonPath);
        if (!Files.isRegularFile(lock)) {
            return Optional.empty();
        }
        try {
            String text = Files.readString(lock, StandardCharsets.UTF_8);
            return Optional.of(parseLockPayload(text, historyJsonPath));
        } catch (Exception ignored) {
            return Optional.of(
                    new LockInfo(
                            historyJsonPath.toString(),
                            "",
                            "",
                            0L,
                            "",
                            Instant.EPOCH));
        }
    }

    public static Optional<AcquiredLock> tryAcquire(Path historyJsonPath) throws IOException {
        Path lock = lockFilePath(historyJsonPath);
        Path parent = lock.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        String payload = formatLockPayload(historyJsonPath);
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

    public static String localHostName() {
        try {
            return InetAddress.getLocalHost().getHostName();
        } catch (Exception ex) {
            String env = System.getenv("COMPUTERNAME");
            if (env != null && !env.isBlank()) {
                return env;
            }
            env = System.getenv("HOSTNAME");
            return env != null ? env : "";
        }
    }

    public static String localHostIp() {
        try {
            return InetAddress.getLocalHost().getHostAddress();
        } catch (Exception ex) {
            return "";
        }
    }

    private static boolean deleteIfExistsQuiet(Path lock) {
        try {
            return Files.deleteIfExists(lock);
        } catch (IOException ignored) {
            return false;
        }
    }

    private static String formatLockPayload(Path historyJsonPath) {
        Map<String, String> m = new LinkedHashMap<>();
        m.put(KEY_VERSION, "1");
        m.put(KEY_HISTORY, historyJsonPath.toAbsolutePath().normalize().toString());
        m.put(KEY_HOST, localHostName());
        m.put(KEY_HOST_IP, localHostIp());
        m.put(KEY_PID, Long.toString(ProcessHandle.current().pid()));
        m.put(KEY_USER, System.getProperty("user.name", ""));
        m.put(KEY_STARTED_AT, Instant.now().toString());
        StringBuilder sb = new StringBuilder();
        for (Map.Entry<String, String> e : m.entrySet()) {
            sb.append(e.getKey()).append('=').append(escapeValue(e.getValue())).append('\n');
        }
        return sb.toString();
    }

    private static LockInfo parseLockPayload(String text, Path fallbackHistory) {
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
                m.getOrDefault(KEY_HISTORY, fallbackHistory.toString()),
                m.getOrDefault(KEY_HOST, ""),
                m.getOrDefault(KEY_HOST_IP, ""),
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
}
