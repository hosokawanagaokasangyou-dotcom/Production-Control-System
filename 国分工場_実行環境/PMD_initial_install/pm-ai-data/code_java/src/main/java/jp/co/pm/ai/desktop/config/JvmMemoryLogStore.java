package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.runtime.MemoryJvmRingLog;

/**
 * JVM memory ring log (max lines + text lines) storage, separate from {@link DesktopSessionStateStore} and from the
 * main run-tab log ({@code mainRunLogLines}).
 *
 * <p>File: {@code ~/.pm-ai-desktop/jvm-memory-log.json}
 */
public final class JvmMemoryLogStore {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final Path DIR = Paths.get(System.getProperty("user.home"), ".pm-ai-desktop");
    private static final Path FILE = DIR.resolve("jvm-memory-log.json");
    private static final Path SESSION_LEGACY = DIR.resolve("session-state.json");

    private JvmMemoryLogStore() {}

    /**
     * Loads persisted settings and ring lines into {@link MemoryJvmRingLog}. If {@link #FILE} is absent, migrates
     * {@code memoryJvmLogMaxLines} from legacy {@link #SESSION_LEGACY} when present.
     */
    public static void bootstrapRingFromDisk() {
        long max = MemoryJvmRingLog.DEFAULT_MAX_LINES;
        List<String> lines = List.of();
        try {
            if (Files.isRegularFile(FILE)) {
                JsonNode root = JSON.readTree(FILE.toFile());
                max = clampLong(root.path("memoryJvmLogMaxLines").asLong(MemoryJvmRingLog.DEFAULT_MAX_LINES));
                lines = loadLinesArray(root.get("lines"));
            } else {
                max = loadLegacyMaxLinesFromSessionStateOnly();
            }
        } catch (IOException ignored) {
            max = loadLegacyMaxLinesFromSessionStateOnly();
        }
        MemoryJvmRingLog.setMaxLines((int) max);
        MemoryJvmRingLog.replaceLines(lines);
    }

    /** Writes max-lines + ring snapshot (does not touch {@code session-state.json}). */
    public static void persistSnapshot(long maxLinesValue, List<String> ringLines) {
        try {
            Files.createDirectories(DIR);
            ObjectNode root = JSON.createObjectNode();
            root.put("memoryJvmLogMaxLines", clampLong(maxLinesValue));
            ArrayNode arr = JSON.createArrayNode();
            if (ringLines != null) {
                for (String line : ringLines) {
                    arr.add(line != null ? line : "");
                }
            }
            root.set("lines", arr);
            JSON.writerWithDefaultPrettyPrinter().writeValue(FILE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    private static long loadLegacyMaxLinesFromSessionStateOnly() {
        try {
            if (!Files.isRegularFile(SESSION_LEGACY)) {
                return MemoryJvmRingLog.DEFAULT_MAX_LINES;
            }
            JsonNode root = JSON.readTree(SESSION_LEGACY.toFile());
            JsonNode n = root.get("memoryJvmLogMaxLines");
            if (n != null && n.isNumber()) {
                return clampLong(n.asLong());
            }
        } catch (IOException ignored) {
        }
        return MemoryJvmRingLog.DEFAULT_MAX_LINES;
    }

    private static List<String> loadLinesArray(JsonNode linesNode) {
        if (linesNode == null || !linesNode.isArray()) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        for (JsonNode el : linesNode) {
            if (el != null && el.isTextual()) {
                out.add(el.asText(""));
            } else if (el != null && el.isValueNode()) {
                out.add(el.asText(""));
            }
        }
        return List.copyOf(out);
    }

    private static long clampLong(long v) {
        return Math.max(MemoryJvmRingLog.ABS_MIN, Math.min(MemoryJvmRingLog.ABS_MAX, v));
    }
}
