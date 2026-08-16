package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.EnumMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PipelineExecutionTimingHistoryLock;
import jp.co.pm.ai.desktop.io.PipelineExecutionTimingHistoryLock.AcquiredLock;
import jp.co.pm.ai.desktop.io.PipelineExecutionTimingHistoryLock.LockInfo;

/**
 * 段階1／2.0〜2.1・配台試行・サマリ Excel・納期管理ビューの実行時間履歴。永続化先は
 * {@link AppPaths#pipelineExecutionTimingHistoryPath}（サマリ Excel と同一フォルダ）。
 * 保存前に {@link PipelineExecutionTimingHistoryLock} を取得し、他 PC 更新時はマージして排他保存する。
 */
public final class PipelineExecutionTimingHistoryStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final Path LEGACY_STORE =
            Paths.get(
                    System.getProperty("user.home"),
                    ".pm-ai-desktop",
                    AppPaths.PIPELINE_EXECUTION_TIMING_HISTORY_JSON);

    private static final int DEFAULT_MAX_SAMPLES_PER_KIND = 300;
    private static final int PERSIST_MAX_ATTEMPTS = 24;
    private static final long PERSIST_RETRY_MS = 500L;

    public record Stats(
            long count,
            double avgSec,
            double medianSec,
            double minSec,
            double maxSec,
            double stdDevSec) {

        static Stats empty() {
            return new Stats(0L, 0d, 0d, 0d, 0d, 0d);
        }

        String summaryJa() {
            if (count <= 0L) {
                return "サンプルなし";
            }
            return String.format(
                    Locale.ROOT,
                    "件数 %d　平均 %.2f秒　中央値 %.2f秒　最小 %.2f秒　最大 %.2f秒　標準偏差 %.2f秒",
                    count, avgSec, medianSec, minSec, maxSec, stdDevSec);
        }
    }

    public record HistogramBin(String label, double rangeStartSec, double rangeEndSec, int count) {}

    private final List<PipelineExecutionTimingSample> samples =
            Collections.synchronizedList(new ArrayList<>());
    private final Map<PipelineExecutionTimingKind, Long> startNanos =
            new EnumMap<>(PipelineExecutionTimingKind.class);
    private final Map<PipelineExecutionTimingKind, Long> lastDurationMs =
            new EnumMap<>(PipelineExecutionTimingKind.class);
    private final CopyOnWriteArrayList<Runnable> changeListeners = new CopyOnWriteArrayList<>();

    private volatile Consumer<String> persistLog;

    private Path storePath = AppPaths.pipelineExecutionTimingHistoryPath(Map.of());
    private int maxSamplesPerKind = DEFAULT_MAX_SAMPLES_PER_KIND;
    private boolean configured;

    public void setPersistLog(Consumer<String> persistLog) {
        this.persistLog = persistLog;
    }

    /** {@link AppPaths#summaryAiDispatchXlsxPath} と同じフォルダへ履歴パスを解決して読み込む。 */
    public synchronized void configureFromUi(Map<String, String> ui) {
        Path next = AppPaths.pipelineExecutionTimingHistoryPath(ui);
        if (configured && next.equals(storePath)) {
            reloadMergedFromDiskIfUnlocked();
            return;
        }
        load(next);
        configured = true;
    }

    public synchronized void load(Path path) {
        storePath =
                path != null
                        ? path
                        : AppPaths.pipelineExecutionTimingHistoryPath(Map.of());
        samples.clear();
        lastDurationMs.clear();
        migrateLegacyStoreIfNeeded(storePath);
        mergeDiskIntoMemory(storePath);
        notifyListeners();
    }

    public synchronized Path storagePath() {
        return storePath;
    }

    public boolean isPersistLocked() {
        return PipelineExecutionTimingHistoryLock.isLocked(storePath);
    }

    public Optional<LockInfo> readPersistLockInfo() {
        return PipelineExecutionTimingHistoryLock.readLockInfo(storePath);
    }

    public void addChangeListener(Runnable listener) {
        if (listener != null) {
            changeListeners.add(listener);
        }
    }

    public void removeChangeListener(Runnable listener) {
        if (listener != null) {
            changeListeners.remove(listener);
        }
    }

    public synchronized void begin(PipelineExecutionTimingKind kind) {
        if (kind == null) {
            return;
        }
        startNanos.put(kind, System.nanoTime());
    }

    public synchronized void end(PipelineExecutionTimingKind kind) {
        if (kind == null) {
            return;
        }
        Long start = startNanos.remove(kind);
        if (start == null) {
            return;
        }
        long durationMs = Math.max(0L, (System.nanoTime() - start) / 1_000_000L);
        lastDurationMs.put(kind, durationMs);
        samples.add(
                new PipelineExecutionTimingSample(
                        kind,
                        Instant.now().toEpochMilli(),
                        durationMs,
                        PipelineExecutionTimingHistoryLock.localHostName(),
                        PipelineExecutionTimingHistoryLock.localHostIp()));
        trimKind(kind);
        persistAsync();
        notifyListeners();
    }

    public synchronized long lastDurationMs(PipelineExecutionTimingKind kind) {
        if (kind == null) {
            return -1L;
        }
        return lastDurationMs.getOrDefault(kind, -1L);
    }

    public synchronized List<PipelineExecutionTimingSample> recentSamples(
            PipelineExecutionTimingKind kind, int limit) {
        int cap = Math.max(1, limit);
        List<PipelineExecutionTimingSample> out = new ArrayList<>();
        synchronized (samples) {
            for (int i = samples.size() - 1; i >= 0 && out.size() < cap; i--) {
                PipelineExecutionTimingSample s = samples.get(i);
                if (kind == null || s.kind() == kind) {
                    out.add(s);
                }
            }
        }
        Collections.reverse(out);
        return List.copyOf(out);
    }

    public synchronized List<PipelineExecutionTimingSample> recentSamplesForKind(
            PipelineExecutionTimingKind kind, int limit) {
        if (kind == null) {
            return List.of();
        }
        return recentSamples(kind, limit);
    }

    public synchronized void clearAll() {
        samples.clear();
        lastDurationMs.clear();
        startNanos.clear();
        persistAsync();
        notifyListeners();
    }

    public synchronized void clearKind(PipelineExecutionTimingKind kind) {
        if (kind == null) {
            return;
        }
        samples.removeIf(s -> s.kind() == kind);
        lastDurationMs.remove(kind);
        startNanos.remove(kind);
        persistAsync();
        notifyListeners();
    }

    static List<PipelineExecutionTimingSample> mergeSamples(
            List<PipelineExecutionTimingSample> disk, List<PipelineExecutionTimingSample> memory) {
        LinkedHashMap<String, PipelineExecutionTimingSample> merged = new LinkedHashMap<>();
        if (disk != null) {
            for (PipelineExecutionTimingSample s : disk) {
                merged.put(sampleKey(s), s);
            }
        }
        if (memory != null) {
            for (PipelineExecutionTimingSample s : memory) {
                merged.put(sampleKey(s), s);
            }
        }
        List<PipelineExecutionTimingSample> out = new ArrayList<>(merged.values());
        out.sort(Comparator.comparingLong(PipelineExecutionTimingSample::finishedAtEpochMs));
        return List.copyOf(out);
    }

    public static Stats computeStats(List<PipelineExecutionTimingSample> list) {
        if (list == null || list.isEmpty()) {
            return Stats.empty();
        }
        List<Double> secs = new ArrayList<>(list.size());
        for (PipelineExecutionTimingSample s : list) {
            secs.add(s.durationMs() / 1000.0);
        }
        Collections.sort(secs);
        double min = secs.getFirst();
        double max = secs.getLast();
        double sum = 0d;
        for (double v : secs) {
            sum += v;
        }
        double avg = sum / secs.size();
        double median;
        int n = secs.size();
        if (n % 2 == 1) {
            median = secs.get(n / 2);
        } else {
            median = (secs.get(n / 2 - 1) + secs.get(n / 2)) / 2.0;
        }
        double var = 0d;
        for (double v : secs) {
            double d = v - avg;
            var += d * d;
        }
        double std = n > 1 ? Math.sqrt(var / (n - 1)) : 0d;
        return new Stats(n, avg, median, min, max, std);
    }

    public static List<HistogramBin> computeHistogram(
            List<PipelineExecutionTimingSample> list, int requestedBinCount) {
        if (list == null || list.isEmpty()) {
            return List.of();
        }
        List<Double> secs = new ArrayList<>(list.size());
        for (PipelineExecutionTimingSample s : list) {
            secs.add(s.durationMs() / 1000.0);
        }
        Collections.sort(secs);
        double min = secs.getFirst();
        double max = secs.getLast();
        if (Math.abs(max - min) < 1e-9) {
            return List.of(new HistogramBin(formatBinLabel(min, max), min, max, secs.size()));
        }
        int bins = Math.max(3, Math.min(24, requestedBinCount));
        double width = (max - min) / bins;
        if (width <= 0d) {
            width = 1d;
        }
        int[] counts = new int[bins];
        for (double v : secs) {
            int idx = (int) Math.floor((v - min) / width);
            if (idx >= bins) {
                idx = bins - 1;
            }
            if (idx < 0) {
                idx = 0;
            }
            counts[idx]++;
        }
        List<HistogramBin> out = new ArrayList<>(bins);
        for (int i = 0; i < bins; i++) {
            double start = min + i * width;
            double end = i == bins - 1 ? max : min + (i + 1) * width;
            out.add(new HistogramBin(formatBinLabel(start, end), start, end, counts[i]));
        }
        return List.copyOf(out);
    }

    private synchronized void reloadMergedFromDiskIfUnlocked() {
        if (PipelineExecutionTimingHistoryLock.isLocked(storePath)) {
            return;
        }
        mergeDiskIntoMemory(storePath);
        notifyListeners();
    }

    private synchronized void mergeDiskIntoMemory(Path target) {
        List<PipelineExecutionTimingSample> disk = readSamplesFromFile(target);
        List<PipelineExecutionTimingSample> merged = mergeSamples(disk, snapshotSamples());
        replaceSamples(merged);
    }

    private synchronized List<PipelineExecutionTimingSample> snapshotSamples() {
        return List.copyOf(samples);
    }

    private synchronized void replaceSamples(List<PipelineExecutionTimingSample> merged) {
        samples.clear();
        lastDurationMs.clear();
        samples.addAll(merged);
        for (PipelineExecutionTimingKind kind : PipelineExecutionTimingKind.values()) {
            trimKind(kind);
        }
        for (PipelineExecutionTimingSample s : samples) {
            lastDurationMs.put(s.kind(), s.durationMs());
        }
    }

    private static String sampleKey(PipelineExecutionTimingSample s) {
        return s.kind().name()
                + '|'
                + s.finishedAtEpochMs()
                + '|'
                + s.durationMs()
                + '|'
                + nz(s.writerHost())
                + '|'
                + nz(s.writerIp());
    }

    private static String nz(String v) {
        return v != null ? v : "";
    }

    private static String formatBinLabel(double startSec, double endSec) {
        if (startSec >= 60d || endSec >= 60d) {
            return String.format(Locale.ROOT, "%.1f–%.1f分", startSec / 60d, endSec / 60d);
        }
        return String.format(Locale.ROOT, "%.1f–%.1f秒", startSec, endSec);
    }

    private void trimKind(PipelineExecutionTimingKind kind) {
        int perKind = 0;
        for (int i = samples.size() - 1; i >= 0; i--) {
            if (samples.get(i).kind() == kind) {
                perKind++;
                if (perKind > maxSamplesPerKind) {
                    samples.remove(i);
                }
            }
        }
    }

    private void persistAsync() {
        Path target = storePath;
        List<PipelineExecutionTimingSample> memorySnap;
        synchronized (this) {
            memorySnap = snapshotSamples();
        }
        Thread t =
                new Thread(
                        () -> {
                            try {
                                persistNow(target, memorySnap);
                            } catch (IOException ex) {
                                logPersist(
                                        "[pipeline-timing] 履歴保存失敗: "
                                                + (ex.getMessage() != null
                                                        ? ex.getMessage()
                                                        : ex.toString()));
                            }
                        },
                        "pipeline-timing-history-save");
        t.setDaemon(true);
        t.start();
    }

    private void persistNow(Path target, List<PipelineExecutionTimingSample> memorySnap)
            throws IOException {
        boolean loggedWait = false;
        for (int attempt = 0; attempt < PERSIST_MAX_ATTEMPTS; attempt++) {
            if (attempt > 0) {
                try {
                    Thread.sleep(PERSIST_RETRY_MS);
                } catch (InterruptedException ie) {
                    Thread.currentThread().interrupt();
                    return;
                }
            }
            Optional<LockInfo> holder = PipelineExecutionTimingHistoryLock.readLockInfo(target);
            if (holder.isPresent() && !loggedWait) {
                LockInfo info = holder.get();
                logPersist(
                        "[pipeline-timing] 履歴保存: 他端末の更新待ち（"
                                + info.displayHost()
                                + " / "
                                + info.displayIp()
                                + "）");
                loggedWait = true;
            }
            Optional<AcquiredLock> lock = PipelineExecutionTimingHistoryLock.tryAcquire(target);
            if (lock.isEmpty()) {
                continue;
            }
            try (AcquiredLock acquired = lock.get()) {
                List<PipelineExecutionTimingSample> disk = readSamplesFromFile(target);
                List<PipelineExecutionTimingSample> merged = mergeSamples(disk, memorySnap);
                synchronized (this) {
                    replaceSamples(merged);
                }
                writeSamplesToFile(target, merged);
                logPersist(
                        "[pipeline-timing] 履歴保存: "
                                + target
                                + " （"
                                + merged.size()
                                + " 件, 記録端末 "
                                + PipelineExecutionTimingHistoryLock.localHostIp()
                                + "）");
                notifyListeners();
                return;
            }
        }
        logPersist("[pipeline-timing] 履歴保存失敗: ロック取得タイムアウト（" + target + "）");
    }

    private static List<PipelineExecutionTimingSample> readSamplesFromFile(Path target) {
        if (!Files.isRegularFile(target)) {
            return List.of();
        }
        try {
            JsonNode root = JSON.readTree(target.toFile());
            JsonNode arr = root.get("samples");
            if (arr == null || !arr.isArray()) {
                return List.of();
            }
            List<PipelineExecutionTimingSample> out = new ArrayList<>();
            for (JsonNode el : arr) {
                PipelineExecutionTimingSample s = parseSample(el);
                if (s != null) {
                    out.add(s);
                }
            }
            return out;
        } catch (IOException ex) {
            return List.of();
        }
    }

    private static PipelineExecutionTimingSample parseSample(JsonNode el) {
        if (el == null || !el.isObject()) {
            return null;
        }
        PipelineExecutionTimingKind kind = parseKind(text(el, "kind"));
        if (kind == null) {
            return null;
        }
        long finishedAt = el.path("finishedAtEpochMs").asLong(0L);
        long durationMs = el.path("durationMs").asLong(-1L);
        if (durationMs < 0L) {
            return null;
        }
        return new PipelineExecutionTimingSample(
                kind,
                finishedAt,
                durationMs,
                text(el, "writerHost"),
                text(el, "writerIp"));
    }

    private static void writeSamplesToFile(
            Path target, List<PipelineExecutionTimingSample> merged) throws IOException {
        ObjectNode root = JSON.createObjectNode();
        ArrayNode arr = root.putArray("samples");
        for (PipelineExecutionTimingSample s : merged) {
            ObjectNode o = arr.addObject();
            o.put("kind", s.kind().name());
            o.put("finishedAtEpochMs", s.finishedAtEpochMs());
            o.put("durationMs", s.durationMs());
            o.put("writerHost", nz(s.writerHost()));
            o.put("writerIp", nz(s.writerIp()));
        }
        if (target.getParent() != null) {
            Files.createDirectories(target.getParent());
        }
        Path tmp = target.resolveSibling(target.getFileName() + ".tmp");
        JSON.writerWithDefaultPrettyPrinter().writeValue(tmp.toFile(), root);
        Files.move(tmp, target, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
    }

    private void logPersist(String line) {
        Consumer<String> log = persistLog;
        if (log != null) {
            try {
                log.accept(line);
            } catch (RuntimeException ignored) {
                // ログ出力失敗で計測本体を止めない
            }
        }
    }

    private void notifyListeners() {
        for (Runnable r : changeListeners) {
            try {
                r.run();
            } catch (RuntimeException ignored) {
                // listener 側の UI 更新失敗で計測本体を止めない
            }
        }
    }

    private static PipelineExecutionTimingKind parseKind(String name) {
        if (name == null || name.isBlank()) {
            return null;
        }
        String trimmed = name.trim();
        if ("STAGE2".equals(trimmed)) {
            return PipelineExecutionTimingKind.STAGE2_0;
        }
        if ("STAGE3".equals(trimmed)
                || "STAGE3_0".equals(trimmed)
                || "STAGE3_1".equals(trimmed)
                || "STAGE3_2".equals(trimmed)) {
            return PipelineExecutionTimingKind.DISPATCH_TRIAL;
        }
        try {
            return PipelineExecutionTimingKind.valueOf(trimmed);
        } catch (IllegalArgumentException ex) {
            return null;
        }
    }

    private static void migrateLegacyStoreIfNeeded(Path target) {
        if (Files.isRegularFile(target) || !Files.isRegularFile(LEGACY_STORE)) {
            return;
        }
        try {
            if (target.getParent() != null) {
                Files.createDirectories(target.getParent());
            }
            Files.copy(LEGACY_STORE, target, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException ignored) {
            // 旧 ~/.pm-ai-desktop からの移行に失敗しても新規作成で続行
        }
    }

    private static String text(JsonNode node, String field) {
        JsonNode v = node.get(field);
        return v != null && !v.isNull() ? v.asText("") : "";
    }
}
