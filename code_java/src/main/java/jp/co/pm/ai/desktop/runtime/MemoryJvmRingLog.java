package jp.co.pm.ai.desktop.runtime;

import java.util.ArrayDeque;
import java.util.Deque;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicReference;

/**
 * In-memory ring buffer for JVM memory log lines ({@link JvmMemoryMonitor}, {@link FxJvmMemoryStatusBar}). Oldest
 * lines are dropped when the buffer exceeds {@link #setMaxLines(int)}.
 */
public final class MemoryJvmRingLog {

    public static final int DEFAULT_MAX_LINES = 10_000;
    public static final int ABS_MIN = 100;
    public static final int ABS_MAX = 500_000;

    private static final Deque<String> LINES = new ArrayDeque<>();
    private static final Object LOCK = new Object();

    private static volatile int maxLines = DEFAULT_MAX_LINES;
    private static final AtomicReference<Runnable> UI_REFRESH = new AtomicReference<>();

    private MemoryJvmRingLog() {}

    /** Called from JavaFX controller; listener should hop to the FX thread if it touches UI. */
    public static void setUiRefreshListener(Runnable listener) {
        UI_REFRESH.set(listener);
    }

    public static int getMaxLines() {
        synchronized (LOCK) {
            return maxLines;
        }
    }

    public static void setMaxLines(int requested) {
        int n = clampSize(requested);
        synchronized (LOCK) {
            maxLines = n;
            while (LINES.size() > maxLines) {
                LINES.removeFirst();
            }
        }
        notifyUi();
    }

    public static void append(String line) {
        Objects.requireNonNull(line, "line");
        synchronized (LOCK) {
            LINES.addLast(line);
            while (LINES.size() > maxLines) {
                LINES.removeFirst();
            }
        }
        notifyUi();
    }

    public static List<String> snapshotLines() {
        synchronized (LOCK) {
            return List.copyOf(LINES);
        }
    }

    /** Replaces buffer contents (e.g. after loading from disk); trims to current {@link #maxLines}. */
    public static void replaceLines(List<String> incoming) {
        synchronized (LOCK) {
            LINES.clear();
            if (incoming != null) {
                for (String line : incoming) {
                    if (line != null) {
                        LINES.addLast(line);
                        while (LINES.size() > maxLines) {
                            LINES.removeFirst();
                        }
                    }
                }
            }
        }
        notifyUi();
    }

    private static void notifyUi() {
        Runnable r = UI_REFRESH.get();
        if (r != null) {
            r.run();
        }
    }

    private static int clampSize(int requested) {
        return Math.max(ABS_MIN, Math.min(ABS_MAX, requested));
    }
}
