package jp.co.pm.ai.desktop.gemini;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * {@link AppPaths#geminiFreeTierFlashLiteModelsCachePath(Map)} への JSON 永続化。
 */
public final class GeminiFreeTierModelsCache {

    public static final Duration DEFAULT_MAX_AGE = Duration.ofDays(1);

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private GeminiFreeTierModelsCache() {}

    public record Snapshot(
            long refreshedAtEpochMillis,
            List<String> modelIds,
            String lastError,
            String source) {

        public Snapshot {
            modelIds = modelIds != null ? List.copyOf(modelIds) : List.of();
        }

        public boolean hasModels() {
            return !modelIds.isEmpty();
        }
    }

    public static Path resolvePath(Map<String, String> ui) {
        return AppPaths.geminiFreeTierFlashLiteModelsCachePath(ui);
    }

    public static Optional<Snapshot> read(Path path) {
        if (path == null || !Files.isRegularFile(path)) {
            return Optional.empty();
        }
        try {
            JsonNode root = MAPPER.readTree(path.toFile());
            long at = root.path("refreshedAtEpochMillis").asLong(0L);
            List<String> ids = new ArrayList<>();
            JsonNode arr = root.path("modelIds");
            if (arr.isArray()) {
                for (JsonNode n : arr) {
                    String s = n.asText("").strip();
                    if (!s.isEmpty()) {
                        ids.add(s);
                    }
                }
            }
            String err = root.path("lastError").asText(null);
            if (err != null && err.isBlank()) {
                err = null;
            }
            String source = root.path("source").asText("");
            return Optional.of(new Snapshot(at, List.copyOf(ids), err, source));
        } catch (IOException e) {
            return Optional.empty();
        }
    }

    public static void write(Path path, Snapshot snapshot) throws IOException {
        if (path == null) {
            throw new IllegalArgumentException("path");
        }
        Path parent = path.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        ObjectNode root = JsonNodeFactory.instance.objectNode();
        root.put("refreshedAtEpochMillis", snapshot.refreshedAtEpochMillis());
        ArrayNode ids = JsonNodeFactory.instance.arrayNode();
        for (String id : snapshot.modelIds()) {
            ids.add(id);
        }
        root.set("modelIds", ids);
        if (snapshot.lastError() != null) {
            root.put("lastError", snapshot.lastError());
        }
        root.put("source", snapshot.source() != null ? snapshot.source() : "");
        MAPPER.writerWithDefaultPrettyPrinter().writeValue(path.toFile(), root);
    }

    public static boolean isStale(Snapshot snapshot, Duration maxAge) {
        if (snapshot == null) {
            return true;
        }
        Duration age = maxAge != null ? maxAge : DEFAULT_MAX_AGE;
        long maxMs = age.toMillis();
        if (maxMs <= 0) {
            return false;
        }
        long at = snapshot.refreshedAtEpochMillis();
        if (at <= 0) {
            return true;
        }
        return System.currentTimeMillis() - at >= maxMs;
    }
}
