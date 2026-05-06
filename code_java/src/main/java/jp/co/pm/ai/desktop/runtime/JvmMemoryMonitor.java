package jp.co.pm.ai.desktop.runtime;

import java.lang.management.ManagementFactory;
import java.lang.management.MemoryMXBean;
import java.lang.management.MemoryUsage;
import java.lang.management.ThreadMXBean;
import java.util.Locale;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * Periodically logs heap and thread usage (disabled unless interval &gt; 0). At most {@link
 * MemoryStderrLogGate#MAX_LINES} lines are written to stderr across all PM-AI memory loggers; then this monitor
 * stops.
 *
 * <p>Enable via system property {@code pm.ai.jvm.memory.monitor.intervalSec} or env {@code
 * PM_AI_JVM_MEMORY_MONITOR_SEC} (seconds). Optional {@code pm.ai.jvm.memory.monitor.warnRatio} (0..1, default
 * 0.85).
 *
 * <p>This source is ASCII-only so compilation succeeds even when the repo path is edited on Windows with a
 * non-UTF-8 editor encoding (javac uses UTF-8 per pom).
 */
public final class JvmMemoryMonitor {

    private static final String PROP_INTERVAL = "pm.ai.jvm.memory.monitor.intervalSec";
    private static final String ENV_INTERVAL = "PM_AI_JVM_MEMORY_MONITOR_SEC";
    private static final String PROP_WARN_RATIO = "pm.ai.jvm.memory.monitor.warnRatio";

    private static final AtomicBoolean STARTED = new AtomicBoolean(false);

    private JvmMemoryMonitor() {}

    /** Called from {@link jp.co.pm.ai.desktop.PmAiFxApp#main}; effective once. */
    public static void startFromMain() {
        if (!STARTED.compareAndSet(false, true)) {
            return;
        }
        long intervalSec = resolveIntervalSec();
        if (intervalSec <= 0) {
            STARTED.set(false);
            return;
        }
        double warnRatio = resolveWarnRatio();

        ScheduledExecutorService scheduler =
                Executors.newSingleThreadScheduledExecutor(
                        r -> {
                            Thread t = new Thread(r, "pm-ai-jvm-memory-monitor");
                            t.setDaemon(true);
                            return t;
                        });

        MemoryMXBean memoryBean = ManagementFactory.getMemoryMXBean();
        ThreadMXBean threadBean = ManagementFactory.getThreadMXBean();

        Runnable task =
                () -> {
                    try {
                        MemoryUsage heap = memoryBean.getHeapMemoryUsage();
                        MemoryUsage nonHeap = memoryBean.getNonHeapMemoryUsage();
                        long used = heap.getUsed();
                        long max = heap.getMax();
                        double pct =
                                max > 0 ? (100.0 * used / (double) max) : Double.NaN;
                        int threads = threadBean.getThreadCount();

                        String line =
                                String.format(
                                        Locale.ROOT,
                                        "[PM-AI heap] used=%d MiB committed=%d MiB max=%s "
                                                + "nonHeap_used=%d MiB threads=%d heap_usage=%s",
                                        toMiB(used),
                                        toMiB(heap.getCommitted()),
                                        max >= 0 ? toMiB(max) + " MiB" : "n/a",
                                        toMiB(nonHeap.getUsed()),
                                        threads,
                                        max > 0
                                                ? String.format(Locale.ROOT, "%.1f%%", pct)
                                                : "n/a");
                        if (!MemoryStderrLogGate.recordLine(line)) {
                            scheduler.shutdown();
                            return;
                        }

                        if (max > 0 && used >= max * warnRatio) {
                            String warn =
                                    "[PM-AI heap] WARN: heap usage exceeded threshold "
                                            + String.format(
                                                    Locale.ROOT, "%.0f%%", warnRatio * 100)
                                            + ". Consider increasing -Xmx or splitting work.";
                            if (!MemoryStderrLogGate.recordLine(warn)) {
                                scheduler.shutdown();
                            }
                        }
                    } catch (Throwable ignored) {
                        // monitoring only
                    }
                };

        scheduler.scheduleAtFixedRate(task, intervalSec, intervalSec, TimeUnit.SECONDS);
        Runtime.getRuntime()
                .addShutdownHook(
                        new Thread(
                                () -> {
                                    scheduler.shutdown();
                                    try {
                                        if (!scheduler.awaitTermination(2, TimeUnit.SECONDS)) {
                                            scheduler.shutdownNow();
                                        }
                                    } catch (InterruptedException e) {
                                        scheduler.shutdownNow();
                                        Thread.currentThread().interrupt();
                                    }
                                },
                                "pm-ai-jvm-memory-monitor-shutdown"));
    }

    private static long resolveIntervalSec() {
        String p = System.getProperty(PROP_INTERVAL, "").trim();
        long fromProp = parsePositiveLong(p);
        if (fromProp > 0) {
            return fromProp;
        }
        String env = System.getenv(ENV_INTERVAL);
        if (env != null) {
            return parsePositiveLong(env.trim());
        }
        return 0;
    }

    private static long parsePositiveLong(String s) {
        if (s == null || s.isEmpty()) {
            return 0;
        }
        try {
            long v = Long.parseLong(s);
            return v > 0 ? v : 0;
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private static double resolveWarnRatio() {
        String p = System.getProperty(PROP_WARN_RATIO, "0.85").trim();
        try {
            double v = Double.parseDouble(p);
            if (v <= 0 || v > 1) {
                return 0.85;
            }
            return v;
        } catch (NumberFormatException e) {
            return 0.85;
        }
    }

    private static long toMiB(long bytes) {
        return bytes >= 0 ? bytes / (1024 * 1024) : -1;
    }
}
