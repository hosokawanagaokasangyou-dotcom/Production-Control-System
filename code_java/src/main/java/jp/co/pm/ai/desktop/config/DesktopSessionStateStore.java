package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * Persists last-used paths under {@code ~/.pm-ai-desktop/session-state.json} so tabs reload the same files on
 * the next launch.
 */
public final class DesktopSessionStateStore {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final Path STORE =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "session-state.json");

    private DesktopSessionStateStore() {}

    public static DesktopSessionState load() {
        try {
            if (!Files.isRegularFile(STORE)) {
                return DesktopSessionState.empty();
            }
            JsonNode root = JSON.readTree(STORE.toFile());
            if (root == null || !root.isObject()) {
                return DesktopSessionState.empty();
            }
            return new DesktopSessionState(
                    text(root, "planInputPath"),
                    text(root, "planInputSheet"),
                    text(root, "stage1PreviewPath"),
                    text(root, "stage1PreviewSheet"),
                    text(root, "excludeRulesPath"),
                    text(root, "mainRunWorkbook"),
                    text(root, "mainRunPythonExe"),
                    text(root, "mainRunScriptDir"),
                    optionalDouble(root, "windowWidth", 0d),
                    optionalDouble(root, "windowHeight", 0d),
                    optionalDouble(root, "windowX", Double.NaN),
                    optionalDouble(root, "windowY", Double.NaN),
                    text(root, "uiTheme"));
        } catch (IOException e) {
            return DesktopSessionState.empty();
        }
    }

    public static void save(DesktopSessionState state) {
        try {
            Files.createDirectories(STORE.getParent());
            ObjectNode root = JSON.createObjectNode();
            put(root, "planInputPath", state.planInputPath());
            put(root, "planInputSheet", state.planInputSheet());
            put(root, "stage1PreviewPath", state.stage1PreviewPath());
            put(root, "stage1PreviewSheet", state.stage1PreviewSheet());
            put(root, "excludeRulesPath", state.excludeRulesPath());
            put(root, "mainRunWorkbook", state.mainRunWorkbook());
            put(root, "mainRunPythonExe", state.mainRunPythonExe());
            put(root, "mainRunScriptDir", state.mainRunScriptDir());
            put(root, "uiTheme", state.uiTheme());
            putWindowGeometry(root, state);
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

    private static void put(ObjectNode root, String key, String value) {
        if (value != null && !value.isBlank()) {
            root.put(key, value.trim());
        }
    }

    private static void putWindowGeometry(ObjectNode root, DesktopSessionState state) {
        double w = state.windowWidth();
        double h = state.windowHeight();
        if (Double.isFinite(w) && w > 0) {
            root.put("windowWidth", w);
        }
        if (Double.isFinite(h) && h > 0) {
            root.put("windowHeight", h);
        }
        double x = state.windowX();
        double y = state.windowY();
        if (Double.isFinite(x)) {
            root.put("windowX", x);
        }
        if (Double.isFinite(y)) {
            root.put("windowY", y);
        }
    }
}
