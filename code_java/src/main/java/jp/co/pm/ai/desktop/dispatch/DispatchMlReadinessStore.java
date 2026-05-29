package jp.co.pm.ai.desktop.dispatch;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.config.AppPaths;

/** {@code ml_readiness.json} と {@code process_machine_speed.json} の UI 向けスナップショット。 */
public final class DispatchMlReadinessStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private DispatchMlReadinessStore() {}

    public record MlLayerState(boolean eligible, boolean enabled, List<String> blockers) {}

    public record ReadinessSnapshot(
            String updatedAt,
            String mlModeActive,
            int archiveJobCount,
            int speedKeyCount,
            int speedApplicableKeyCount,
            Map<String, MlLayerState> layers) {

        public static ReadinessSnapshot empty() {
            return new ReadinessSnapshot("", "off", 0, 0, 0, Map.of());
        }
    }

    public record SpeedKeyEntry(
            String key,
            String process,
            String machine,
            int n,
            double p50,
            Double appliedSpeed,
            List<Integer> histogramCounts,
            List<Double> histogramEdges) {}

    public static ReadinessSnapshot loadReadiness(Path archiveRoot) {
        if (archiveRoot == null) {
            return ReadinessSnapshot.empty();
        }
        Path p = archiveRoot.resolve("ml_readiness.json");
        if (!Files.isRegularFile(p)) {
            return ReadinessSnapshot.empty();
        }
        try {
            JsonNode root = JSON.readTree(p.toFile());
            Map<String, MlLayerState> layers = new LinkedHashMap<>();
            JsonNode layersNode = root.path("layers");
            if (layersNode.isObject()) {
                Iterator<Map.Entry<String, JsonNode>> it = layersNode.fields();
                while (it.hasNext()) {
                    Map.Entry<String, JsonNode> e = it.next();
                    JsonNode ln = e.getValue();
                    List<String> blockers = new ArrayList<>();
                    if (ln.path("blockers").isArray()) {
                        for (JsonNode b : ln.path("blockers")) {
                            blockers.add(b.asText(""));
                        }
                    }
                    layers.put(
                            e.getKey(),
                            new MlLayerState(
                                    ln.path("eligible").asBoolean(false),
                                    ln.path("enabled").asBoolean(false),
                                    List.copyOf(blockers)));
                }
            }
            return new ReadinessSnapshot(
                    root.path("updated_at").asText(""),
                    root.path("ml_mode_active").asText("off"),
                    root.path("archive_job_count").asInt(0),
                    root.path("speed_key_count").asInt(0),
                    root.path("speed_applicable_key_count").asInt(0),
                    Collections.unmodifiableMap(layers));
        } catch (Exception ex) {
            return ReadinessSnapshot.empty();
        }
    }

    public static List<SpeedKeyEntry> loadSpeedKeys(Path archiveRoot) {
        if (archiveRoot == null) {
            return List.of();
        }
        Path p = archiveRoot.resolve("speed-distributions").resolve("process_machine_speed.json");
        if (!Files.isRegularFile(p)) {
            return List.of();
        }
        try {
            JsonNode root = JSON.readTree(p.toFile());
            List<SpeedKeyEntry> out = new ArrayList<>();
            Iterator<Map.Entry<String, JsonNode>> it = root.fields();
            while (it.hasNext()) {
                Map.Entry<String, JsonNode> e = it.next();
                JsonNode n = e.getValue();
                List<Integer> counts = new ArrayList<>();
                List<Double> edges = new ArrayList<>();
                JsonNode hist = n.path("histogram");
                if (hist.path("counts").isArray()) {
                    for (JsonNode c : hist.path("counts")) {
                        counts.add(c.asInt(0));
                    }
                }
                if (hist.path("bin_edges").isArray()) {
                    for (JsonNode c : hist.path("bin_edges")) {
                        edges.add(c.asDouble(0));
                    }
                }
                Double applied = null;
                if (!n.path("applied_speed_m_per_min").isNull() && n.has("applied_speed_m_per_min")) {
                    applied = n.path("applied_speed_m_per_min").asDouble();
                }
                out.add(
                        new SpeedKeyEntry(
                                e.getKey(),
                                n.path("process").asText(""),
                                n.path("machine").asText(""),
                                n.path("n").asInt(0),
                                n.path("p50").asDouble(0),
                                applied,
                                List.copyOf(counts),
                                List.copyOf(edges)));
            }
            out.sort((a, b) -> a.key().compareTo(b.key()));
            return out;
        } catch (Exception ex) {
            return List.of();
        }
    }

    public static Path defaultArchiveRoot(Map<String, String> ui) {
        return AppPaths.resolveDispatchLearningArchiveRoot(ui);
    }
}
