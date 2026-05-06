package jp.co.pm.ai.desktop.runtime;

import java.lang.management.ManagementFactory;
import java.lang.management.MemoryMXBean;
import java.lang.management.MemoryPoolMXBean;
import java.lang.management.MemoryUsage;
import java.util.Locale;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.application.Platform;
import javafx.scene.control.Label;
import javafx.stage.Stage;
import javafx.stage.WindowEvent;

/**
 * Toolbar label: heap and Metaspace usage; same line to {@code System.err} (subject to {@link
 * MemoryStderrLogGate#MAX_LINES} combined with {@link JvmMemoryMonitor}).
 *
 * <p>Interval 10 seconds. UI strings are ASCII-only (English) so labels render correctly on Japanese Windows
 * when editor or runtime encoding does not match UTF-8 (same rationale as {@link JvmMemoryMonitor}).
 */
public final class FxJvmMemoryStatusBar {

    /** Prefix for stderr lines; matches {@link JvmMemoryMonitor} tag style. */
    private static final String LOG_PREFIX = "[PM-AI memory]";

    private static final int INTERVAL_SEC = 10;
    private static final String POOL_METASPACE = "Metaspace";

    private static final AtomicBoolean GLOBAL_STARTED = new AtomicBoolean(false);

    private FxJvmMemoryStatusBar() {}

    /** Starts sampling (once per JVM). @param label toolbar target @param stage used to stop on close */
    public static void start(Label label, Stage stage) {
        if (label == null || stage == null) {
            return;
        }
        if (!GLOBAL_STARTED.compareAndSet(false, true)) {
            return;
        }

        MemoryMXBean memoryBean = ManagementFactory.getMemoryMXBean();
        MemoryPoolMXBean metaspacePool = resolveMetaspacePool();

        ScheduledExecutorService scheduler =
                Executors.newSingleThreadScheduledExecutor(
                        r -> {
                            Thread t = new Thread(r, "pm-ai-fx-jvm-memory-status");
                            t.setDaemon(true);
                            return t;
                        });

        AtomicBoolean stopped = new AtomicBoolean(false);
        Runnable shutdown =
                () -> {
                    if (!stopped.compareAndSet(false, true)) {
                        return;
                    }
                    scheduler.shutdown();
                    try {
                        if (!scheduler.awaitTermination(2, TimeUnit.SECONDS)) {
                            scheduler.shutdownNow();
                        }
                    } catch (InterruptedException e) {
                        scheduler.shutdownNow();
                        Thread.currentThread().interrupt();
                    }
                };

        Runnable task =
                () -> {
                    try {
                        MemorySample sample = readSample(memoryBean, metaspacePool);
                        String uiText = formatUiText(sample);
                        String logLine = formatLogLine(sample);
                        Platform.runLater(() -> label.setText(uiText));
                        if (!MemoryStderrLogGate.recordLine(LOG_PREFIX + " " + logLine)) {
                            shutdown.run();
                        }
                    } catch (Throwable ignored) {
                        // status bar only
                    }
                };

        scheduler.scheduleAtFixedRate(task, 0, INTERVAL_SEC, TimeUnit.SECONDS);

        stage.addEventHandler(WindowEvent.WINDOW_CLOSE_REQUEST, e -> shutdown.run());
        Runtime.getRuntime()
                .addShutdownHook(new Thread(shutdown, "pm-ai-fx-jvm-memory-status-shutdown"));
    }

    private static MemoryPoolMXBean resolveMetaspacePool() {
        for (MemoryPoolMXBean pool : ManagementFactory.getMemoryPoolMXBeans()) {
            if (POOL_METASPACE.equals(pool.getName())) {
                return pool;
            }
        }
        return null;
    }

    private static MemorySample readSample(MemoryMXBean memoryBean, MemoryPoolMXBean metaspacePool) {
        MemoryUsage heap = memoryBean.getHeapMemoryUsage();
        long heapUsed = heap.getUsed();
        long heapMax = heap.getMax();
        double heapPct = heapMax > 0 ? 100.0 * heapUsed / (double) heapMax : Double.NaN;

        long metaUsed = -1;
        long metaMax = -1;
        if (metaspacePool != null) {
            MemoryUsage u = metaspacePool.getUsage();
            metaUsed = u.getUsed();
            metaMax = u.getMax();
        }
        return new MemorySample(heapUsed, heapMax, heapPct, metaUsed, metaMax);
    }

    private static String formatUiText(MemorySample s) {
        String heapPctStr =
                !Double.isNaN(s.heapPct)
                        ? String.format(Locale.ROOT, "%.1f%%", s.heapPct)
                        : "n/a";
        return "Heap "
                + formatMiPair(s.heapUsed, s.heapMax)
                + " ("
                + heapPctStr
                + ") | Metaspace "
                + formatMiPair(s.metaUsed, s.metaMax);
    }

    /** Log line: ASCII key=value */
    private static String formatLogLine(MemorySample s) {
        return "heap_used_MiB="
                + toMiB(s.heapUsed)
                + " heap_max_MiB="
                + formatMaxMiB(s.heapMax)
                + " heap_pct="
                + (Double.isNaN(s.heapPct) ? "n/a" : String.format(Locale.ROOT, "%.1f", s.heapPct))
                + " metaspace_used_MiB="
                + (s.metaUsed >= 0 ? toMiB(s.metaUsed) : "n/a")
                + " metaspace_max_MiB="
                + formatMaxMiB(s.metaMax);
    }

    private static String formatMiPair(long usedBytes, long maxBytes) {
        String usedStr = usedBytes >= 0 ? toMiB(usedBytes) + " MiB" : "n/a";
        String maxStr;
        if (maxBytes < 0) {
            maxStr = "n/a";
        } else {
            maxStr = toMiB(maxBytes) + " MiB";
        }
        return usedStr + " / " + maxStr;
    }

    private static String formatMaxMiB(long maxBytes) {
        return maxBytes >= 0 ? Long.toString(toMiB(maxBytes)) : "n/a";
    }

    private static long toMiB(long bytes) {
        return bytes >= 0 ? bytes / (1024 * 1024) : -1;
    }

    private record MemorySample(
            long heapUsed, long heapMax, double heapPct, long metaUsed, long metaMax) {}
}
