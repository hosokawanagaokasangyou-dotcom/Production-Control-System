package jp.co.pm.ai.desktop.runtime;

import java.lang.management.ManagementFactory;
import java.lang.management.MemoryMXBean;
import java.lang.management.MemoryUsage;
import java.lang.management.ThreadMXBean;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicBoolean;

import jp.co.pm.ai.desktop.debug.AgentDebugLog;

/**
 * 実行中 JVM のヒープ／スレッド使用量を定期的に記録する（既定オフ）。
 *
 * <p>有効化:
 *
 * <ul>
 *   <li>システムプロパティ {@code pm.ai.jvm.memory.monitor.intervalSec} に正の整数（秒）</li>
 *   <li>または環境変数 {@code PM_AI_JVM_MEMORY_MONITOR_SEC}</li>
 * </ul>
 *
 * <p>任意: {@code pm.ai.jvm.memory.monitor.warnRatio}（0?1、既定 0.85）で警告閾値、
 * {@code pm.ai.jvm.memory.monitor.ndjson=true} で NDJSON 出力（{@link AgentDebugLog}）、
 * {@code pm.ai.jvm.memory.monitor.sessionId} でセッション ID。
 */
public final class JvmMemoryMonitor {

    private static final String PROP_INTERVAL = "pm.ai.jvm.memory.monitor.intervalSec";
    private static final String ENV_INTERVAL = "PM_AI_JVM_MEMORY_MONITOR_SEC";
    private static final String PROP_WARN_RATIO = "pm.ai.jvm.memory.monitor.warnRatio";
    private static final String PROP_NDJSON = "pm.ai.jvm.memory.monitor.ndjson";
    private static final String PROP_SESSION = "pm.ai.jvm.memory.monitor.sessionId";

    private static final AtomicBoolean STARTED = new AtomicBoolean(false);

    private JvmMemoryMonitor() {}

    /** {@link jp.co.pm.ai.desktop.PmAiFxApp#main} から起動。複数回呼んでも一度だけ有効。 */
    public static void startFromMain(Map<String, String> ui) {
        if (!STARTED.compareAndSet(false, true)) {
            return;
        }
        long intervalSec = resolveIntervalSec();
        if (intervalSec <= 0) {
            STARTED.set(false);
            return;
        }
        double warnRatio = resolveWarnRatio();
        boolean ndjson = Boolean.getBoolean(PROP_NDJSON);
        String sessionId = resolveSessionId();

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
                        System.err.println(line);

                        if (max > 0 && used >= max * warnRatio) {
                            System.err.println(
                                    "[PM-AI heap] 警告: ヒープ使用率が閾値 "
                                            + String.format(
                                                    Locale.ROOT, "%.0f%%", warnRatio * 100)
                                            + " を超えました。-Xmx の引き上げや処理の分割を検討してください。");
                        }

                        if (ndjson) {
                            Map<String, Object> data = new LinkedHashMap<>();
                            data.put("heapUsedBytes", used);
                            data.put("heapCommittedBytes", heap.getCommitted());
                            data.put("heapMaxBytes", max);
                            data.put("nonHeapUsedBytes", nonHeap.getUsed());
                            data.put("threadCount", threads);
                            if (max > 0) {
                                data.put("heapUsageRatio", used / (double) max);
                            }
                            AgentDebugLog.appendStructured(
                                    ui != null ? ui : Map.of(),
                                    sessionId,
                                    "heap",
                                    "JvmMemoryMonitor",
                                    "jvm_memory_sample",
                                    data);
                        }
                    } catch (Throwable ignored) {
                        // 監視のみ
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

    private static String resolveSessionId() {
        String p = System.getProperty(PROP_SESSION, "").trim();
        if (!p.isEmpty()) {
            return p;
        }
        return AgentDebugLog.DEFAULT_SESSION_ID;
    }

    private static long toMiB(long bytes) {
        return bytes >= 0 ? bytes / (1024 * 1024) : -1;
    }
}
