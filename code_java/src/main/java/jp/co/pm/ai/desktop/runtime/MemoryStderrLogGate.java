package jp.co.pm.ai.desktop.runtime;

import java.util.concurrent.atomic.AtomicInteger;

/**
 * Limits total stderr lines emitted by PM-AI JVM memory logging ({@link JvmMemoryMonitor}, {@link
 * FxJvmMemoryStatusBar}).
 */
final class MemoryStderrLogGate {

    static final int MAX_LINES = 10_000;

    private static final AtomicInteger COUNT = new AtomicInteger(0);

    private MemoryStderrLogGate() {}

    /**
     * Prints one line to stderr unless the global limit is exceeded.
     *
     * @return {@code false} if this line was not printed (limit already reached); callers may stop schedulers.
     */
    static boolean recordLine(String line) {
        if (COUNT.incrementAndGet() > MAX_LINES) {
            return false;
        }
        System.err.println(line);
        return true;
    }
}
