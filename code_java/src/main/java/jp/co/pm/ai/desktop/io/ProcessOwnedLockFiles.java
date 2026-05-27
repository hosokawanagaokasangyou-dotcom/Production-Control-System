package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 当 JVM が取得したロックファイル（{@code pid=} 行）を追跡し、終了時に自プロセス分だけ削除する。
 */
public final class ProcessOwnedLockFiles {

    private static final long CURRENT_PID = ProcessHandle.current().pid();
    private static final Set<Path> ACTIVE = ConcurrentHashMap.newKeySet();

    static {
        Runtime.getRuntime()
                .addShutdownHook(
                        new Thread(
                                ProcessOwnedLockFiles::releaseAllOwnedQuietly,
                                "pm-ai-owned-lock-cleanup"));
    }

    private ProcessOwnedLockFiles() {}

    public static void register(Path lockPath) {
        if (lockPath == null) {
            return;
        }
        ACTIVE.add(normalize(lockPath));
    }

    public static void unregister(Path lockPath) {
        if (lockPath == null) {
            return;
        }
        ACTIVE.remove(normalize(lockPath));
    }

    /** 追跡中のロックのうち、ロックファイル上の pid が当プロセスのものだけ削除する。 */
    public static void releaseAllOwnedQuietly() {
        for (Path lockPath : List.copyOf(ACTIVE)) {
            releaseIfOwnedByCurrentProcess(lockPath);
        }
    }

    /**
     * ロックファイルが存在し、かつ {@code pid=} が当プロセスのときだけ削除する。
     *
     * @return 削除したら {@code true}
     */
    public static boolean releaseIfOwnedByCurrentProcess(Path lockPath) {
        Path norm = normalize(lockPath);
        ACTIVE.remove(norm);
        if (!Files.isRegularFile(norm)) {
            return false;
        }
        if (!isOwnedByCurrentProcess(norm)) {
            return false;
        }
        try {
            return Files.deleteIfExists(norm);
        } catch (IOException ignored) {
            return false;
        }
    }

    static boolean isOwnedByCurrentProcess(Path lockPath) {
        long pid = readPid(lockPath);
        return pid > 0L && pid == CURRENT_PID;
    }

    static long readPid(Path lockPath) {
        try {
            String text = Files.readString(lockPath, StandardCharsets.UTF_8);
            for (String line : text.split("\n")) {
                String t = line.strip();
                if (!t.startsWith("pid=")) {
                    continue;
                }
                String raw = t.substring("pid=".length()).strip();
                if (raw.isEmpty()) {
                    return 0L;
                }
                return Long.parseLong(raw);
            }
        } catch (Exception ignored) {
            // ignore
        }
        return 0L;
    }

    private static Path normalize(Path lockPath) {
        return lockPath.toAbsolutePath().normalize();
    }
}
