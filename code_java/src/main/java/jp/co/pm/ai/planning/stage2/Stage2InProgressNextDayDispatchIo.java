package jp.co.pm.ai.planning.stage2;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

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
