package jp.co.pm.ai.planning.stage2;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

import jp.co.pm.ai.desktop.config.AppPaths;

/** 段階2直前に書く「アラジン当日・翌日除外量」JSON（Python planning_core と対応）。 */
public final class Stage2AladdinTodayExcludeNextDayDispatchIo {

    private static final ObjectMapper JSON =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    private static final int VERSION = 1;

    private Stage2AladdinTodayExcludeNextDayDispatchIo() {}

    public record Entry(String taskId, String process, String machineName, double excludeNextDayM) {
        public Entry {
            excludeNextDayM = Stage2InProgressNextDayDispatchIo.sanitizeMeters(excludeNextDayM);
        }
    }

    public static String rowKey(String taskId, String process, String machineName) {
        return Stage2InProgressNextDayDispatchIo.rowKey(taskId, process, machineName);
    }

    public static Path defaultCachePath(Map<String, String> ui) {
        return AppPaths.resolveRepoRoot(ui)
                .resolve(".pm-ai-cache")
                .resolve("stage2-aladdin-today-exclude-next-day-dispatch.json");
    }

    public static Path resolveReadPath(Map<String, String> ui) {
        if (ui != null) {
            String fromEnv = ui.get(AppPaths.KEY_PM_AI_STAGE2_ALADDIN_TODAY_EXCLUDE_NEXT_DAY_JSON);
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
            Object mObj = entMap.get("exclude_next_day_m");
            if (mObj instanceof Number n) {
                meters = Stage2InProgressNextDayDispatchIo.sanitizeMeters(n.doubleValue());
            } else if (mObj != null) {
                try {
                    meters =
                            Stage2InProgressNextDayDispatchIo.sanitizeMeters(
                                    Double.parseDouble(String.valueOf(mObj)));
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
            o.put("exclude_next_day_m", e.excludeNextDayM());
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

    public static void writeEmpty(Path path) throws IOException {
        write(path, List.of());
    }
}
