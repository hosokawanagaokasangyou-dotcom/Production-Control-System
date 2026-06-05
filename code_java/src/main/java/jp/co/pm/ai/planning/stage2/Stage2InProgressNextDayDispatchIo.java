package jp.co.pm.ai.planning.stage2;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import jp.co.pm.ai.desktop.config.AppPaths;

/** 段階2直前に書く「加工途中・翌日配台量」JSON（Python {@code planning_core._core} と対応）。 */
public final class Stage2InProgressNextDayDispatchIo {

    private static final ObjectMapper JSON =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    private static final int VERSION = 1;

    private Stage2InProgressNextDayDispatchIo() {}

    public record Entry(String taskId, String process, String machineName, double nextDayDispatchM) {
        public Entry {
            nextDayDispatchM = sanitizeMeters(nextDayDispatchM);
        }
    }

    /** Python {@code _sanitize_dispatch_qty_m} と同趣旨（整数に近い m は整数化）。 */
    public static double sanitizeMeters(double m) {
        if (!Double.isFinite(m) || m < 0) {
            return 0.0;
        }
        long r = Math.round(m);
        if (Math.abs(m - r) <= 1e-6) {
            return (double) r;
        }
        return Math.round(m * 1000.0) / 1000.0;
    }

    public static String rowKey(String taskId, String process, String machineName) {
        String tid = taskId != null ? taskId.strip() : "";
        String proc = process != null ? process.strip() : "";
        String mach = machineName != null ? machineName.strip() : "";
        return tid + "\u001e" + proc + "\u001e" + mach;
    }

    public static Path defaultCachePath(Map<String, String> ui) {
        return AppPaths.resolveRepoRoot(ui)
                .resolve(".pm-ai-cache")
                .resolve("stage2-in-progress-next-day-dispatch.json");
    }

    /** 環境変数 → 既定キャッシュの順で読み込み対象 JSON を解決する。 */
    public static Path resolveReadPath(Map<String, String> ui) {
        if (ui != null) {
            String fromEnv = ui.get(AppPaths.KEY_PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON);
            if (fromEnv != null && !fromEnv.isBlank()) {
                Path p = Path.of(fromEnv.strip()).toAbsolutePath().normalize();
                if (Files.isRegularFile(p)) {
                    return p;
                }
            }
        }
        Path cached = defaultCachePath(ui != null ? ui : Map.of());
        return Files.isRegularFile(cached) ? cached : null;
    }

    public static List<Entry> readEntries(Path path) throws IOException {
        if (path == null || !Files.isRegularFile(path)) {
            return List.of();
        }
        @SuppressWarnings("unchecked")
        Map<String, Object> root = JSON.readValue(path.toFile(), Map.class);
        Object raw = root != null ? root.get("entries") : null;
        if (!(raw instanceof List<?> list)) {
            return List.of();
        }
        List<Entry> out = new ArrayList<>(list.size());
        for (Object item : list) {
            if (!(item instanceof Map<?, ?>)) {
                continue;
            }
            @SuppressWarnings("unchecked")
            Map<String, Object> entMap = (Map<String, Object>) item;
            String taskId = String.valueOf(entMap.getOrDefault("task_id", "")).strip();
            String process = String.valueOf(entMap.getOrDefault("process", "")).strip();
            String machineName = String.valueOf(entMap.getOrDefault("machine_name", "")).strip();
            double meters = 0.0;
            Object mObj = entMap.get("next_day_dispatch_m");
            if (mObj instanceof Number n) {
                meters = sanitizeMeters(n.doubleValue());
            } else if (mObj != null) {
                try {
                    meters = sanitizeMeters(Double.parseDouble(String.valueOf(mObj)));
                } catch (NumberFormatException ignored) {
                    meters = 0.0;
                }
            }
            if (taskId.isEmpty()) {
                continue;
            }
            out.add(new Entry(taskId, process, machineName, meters));
        }
        return List.copyOf(out);
    }

    /** 翌日配台 m が 0 の加工途中行キー（Python {@code rowKey} 形式）。 */
    public static Set<String> zeroNextDayRowKeys(Map<String, String> ui) {
        Path path;
        try {
            path = resolveReadPath(ui);
            if (path == null) {
                return Set.of();
            }
            List<Entry> entries = readEntries(path);
            if (entries.isEmpty()) {
                return Set.of();
            }
            LinkedHashSet<String> out = new LinkedHashSet<>();
            for (Entry e : entries) {
                if (e.nextDayDispatchM() <= 1e-12) {
                    out.add(rowKey(e.taskId(), e.process(), e.machineName()));
                }
            }
            return out.isEmpty() ? Set.of() : Collections.unmodifiableSet(out);
        } catch (IOException ignored) {
            return Set.of();
        }
    }

    public static void write(Path path, List<Entry> entries) throws IOException {
        if (path == null) {
            throw new IOException("path is null");
        }
        Path parent = path.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        List<Map<String, Object>> list = new ArrayList<>();
        for (Entry e : entries) {
            Map<String, Object> o = new LinkedHashMap<>();
            o.put("task_id", e.taskId());
            o.put("process", e.process());
            o.put("machine_name", e.machineName());
            o.put("next_day_dispatch_m", e.nextDayDispatchM());
            list.add(o);
        }
        Map<String, Object> root = new LinkedHashMap<>();
        root.put("version", VERSION);
        root.put("entries", list);
        JSON.writeValue(path.toFile(), root);
    }

    public static void deleteIfExists(Path path) {
        if (path == null) {
            return;
        }
        try {
            Files.deleteIfExists(path);
        } catch (IOException ignored) {
            // best-effort
        }
    }
}
