package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Collections;
import java.util.EnumMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.CopyOnWriteArrayList;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * 段階1～3・サマリ Excel・納期管理ビューの実行時間履歴。永続化先は
 * {@code ~/.pm-ai-desktop/pipeline-execution-timing-history.json}。
 */
public final class PipelineExecutionTimingHistoryStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final Path DEFAULT_STORE =
            Paths.get(
                    System.getProperty("user.home"),
                    ".pm-ai-desktop",
                    "pipeline-execution-timing-history.json");

    private static final int DEFAULT_MAX_SAMPLES_PER_KIND = 300;

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

    private Path storePath = DEFAULT_STORE;
    private int maxSamplesPerKind = DEFAULT_MAX_SAMPLES_PER_KIND;

    public void loadFromDefaultPath() {
        load(storePath);
    }

    public synchronized void load(Path path) {
        storePath = path != null ? path : DEFAULT_STORE;
        samples.clear();
        lastDurationMs.clear();
        if (!Files.isRegularFile(storePath)) {
            return;
        }
        try {
            JsonNode root = JSON.readTree(storePath.toFile());
            JsonNode arr = root.get("samples");
            if (arr == null || !arr.isArray()) {
                return;
            }
            for (JsonNode el : arr) {
                if (el == null || !el.isObject()) {
                    continue;
                }
                String kindName = text(el, "kind");
                PipelineExecutionTimingKind kind = parseKind(kindName);
                if (kind == null) {
                    continue;
                }
                long finishedAt = el.path("finishedAtEpochMs").asLong(0L);
                long durationMs = el.path("durationMs").asLong(-1L);
                if (durationMs < 0L) {
                    continue;
                }
                samples.add(new PipelineExecutionTimingSample(kind, finishedAt, durationMs));
                lastDurationMs.put(kind, durationMs);
            }
            trimAllKinds();
        } catch (IOException ignored) {
            samples.clear();
            lastDurationMs.clear();
        }
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
                        kind, Instant.now().toEpochMilli(), durationMs));
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

    private void trimAllKinds() {
        for (PipelineExecutionTimingKind kind : PipelineExecutionTimingKind.values()) {
            trimKind(kind);
        }
    }

    private void persistAsync() {
        Path target = storePath;
        Thread t =
                new Thread(
                        () -> {
                            try {
                                persistNow(target);
                            } catch (IOException ignored) {
                                // 履歴保存失敗は UI 動作を止めない
                            }
                        },
                        "pipeline-timing-history-save");
        t.setDaemon(true);
        t.start();
    }

    private synchronized void persistNow(Path target) throws IOException {
        ObjectNode root = JSON.createObjectNode();
        ArrayNode arr = root.putArray("samples");
        synchronized (samples) {
            for (PipelineExecutionTimingSample s : samples) {
                ObjectNode o = arr.addObject();
                o.put("kind", s.kind().name());
                o.put("finishedAtEpochMs", s.finishedAtEpochMs());
                o.put("durationMs", s.durationMs());
            }
        }
        Files.createDirectories(target.getParent());
        Path tmp = target.resolveSibling(target.getFileName() + ".tmp");
        JSON.writerWithDefaultPrettyPrinter().writeValue(tmp.toFile(), root);
        Files.move(tmp, target, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
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
        try {
            return PipelineExecutionTimingKind.valueOf(name.trim());
        } catch (IllegalArgumentException ex) {
            return null;
        }
    }

    private static String text(JsonNode node, String field) {
        JsonNode v = node.get(field);
        return v != null && !v.isNull() ? v.asText("") : "";
    }
}
