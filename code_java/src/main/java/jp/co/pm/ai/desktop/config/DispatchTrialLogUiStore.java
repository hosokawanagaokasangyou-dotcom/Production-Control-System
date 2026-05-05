package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * Persists the dispatch-trial log modal geometry and font under {@code ~/.pm-ai-desktop/dispatch-trial-log-ui.json}.
 */
public final class DispatchTrialLogUiStore {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final Path STORE =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "dispatch-trial-log-ui.json");

    public record DispatchTrialLogUiSnapshot(double width, double height, String fontFamily, double fontSize) {

        public static DispatchTrialLogUiSnapshot defaults() {
            return new DispatchTrialLogUiSnapshot(720, 480, "", 12);
        }
    }

    private DispatchTrialLogUiStore() {}

    public static DispatchTrialLogUiSnapshot load() {
        try {
            if (!Files.isRegularFile(STORE)) {
                return DispatchTrialLogUiSnapshot.defaults();
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return DispatchTrialLogUiSnapshot.defaults();
            }
            double w = optionalDouble(root, "width", 0d);
            double h = optionalDouble(root, "height", 0d);
            String fam = text(root, "fontFamily");
            double fs = optionalDouble(root, "fontSize", 0d);
            return new DispatchTrialLogUiSnapshot(
                    w,
                    h,
                    fam != null ? fam : "",
                    Double.isFinite(fs) ? fs : 0d);
        } catch (IOException e) {
            return DispatchTrialLogUiSnapshot.defaults();
        }
    }

    public static void save(DispatchTrialLogUiSnapshot snapshot) {
        if (snapshot == null) {
            return;
        }
        try {
            Files.createDirectories(STORE.getParent());
            ObjectNode root = JSON.createObjectNode();
            double w = snapshot.width();
            double h = snapshot.height();
            if (Double.isFinite(w) && w > 0) {
                root.put("width", w);
            }
            if (Double.isFinite(h) && h > 0) {
                root.put("height", h);
            }
            String fam = snapshot.fontFamily();
            if (fam != null && !fam.isBlank()) {
                root.put("fontFamily", fam.trim());
            }
            double fs = snapshot.fontSize();
            if (Double.isFinite(fs) && fs > 0) {
                root.put("fontSize", fs);
            }
            JSON.writerWithDefaultPrettyPrinter().writeValue(STORE.toFile(), root);
        } catch (IOException ignored) {
        }
    }

    private static String text(JsonNode root, String key) {
        JsonNode n = root.get(key);
        if (n == null || n.isNull() || !n.isTextual()) {
            return "";
        }
        return n.asText("");
    }

    private static double optionalDouble(JsonNode root, String key, double defaultValue) {
        JsonNode n = root.get(key);
        if (n == null || n.isNull() || !n.isNumber()) {
            return defaultValue;
        }
        return n.asDouble();
    }
}
