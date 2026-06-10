package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * 起動プロファイルのメタデータ（名称・説明・区分など）を JSON で読み書きする。
 *
 * <p>保存先は {@link AppPaths#resolveRdpLaunchProfilesFile}（通常は {@code RAP設定.ini} と同階層の
 * {@link AppPaths#RDP_LAUNCH_PROFILES_BASENAME}）。ファイルが無いときは同梱既定を返す。
 */
public final class RdpLaunchProfileCatalog {

    public static final int SCHEMA_VERSION = 1;

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final String DEFAULT_RESOURCE =
            "/jp/co/pm/ai/desktop/rdp-launcher/rdp-launch-profiles-default.json";

    private RdpLaunchProfileCatalog() {}

    public static List<RdpLaunchProfile> load(Path path) throws IOException {
        if (path != null && Files.isRegularFile(path)) {
            JsonNode root = JSON.readTree(Files.readString(path, StandardCharsets.UTF_8));
            return parseProfiles(root);
        }
        return loadBundledDefaults();
    }

    /** ファイルが無い／読込失敗時は同梱既定。 */
    public static List<RdpLaunchProfile> loadOrDefaults(Path path) {
        try {
            if (path != null && Files.isRegularFile(path)) {
                return load(path);
            }
        } catch (IOException ignored) {
            // fall through
        }
        return loadBundledDefaults();
    }

    public static List<RdpLaunchProfile> loadBundledDefaults() {
        try (InputStream in = RdpLaunchProfileCatalog.class.getResourceAsStream(DEFAULT_RESOURCE)) {
            if (in == null) {
                return seedEmptyProfiles(3);
            }
            JsonNode root = JSON.readTree(in);
            return parseProfiles(root);
        } catch (IOException ex) {
            return seedEmptyProfiles(3);
        }
    }

    public static void save(Path path, List<RdpLaunchProfile> profiles) throws IOException {
        Objects.requireNonNull(path, "path");
        Objects.requireNonNull(profiles, "profiles");
        Path parent = path.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", SCHEMA_VERSION);
        ArrayNode array = root.putArray("profiles");
        for (RdpLaunchProfile profile : normalizeList(profiles)) {
            array.add(toJson(profile));
        }
        JSON.writerWithDefaultPrettyPrinter().writeValue(path.toFile(), root);
    }

    /** 番号キーでマージ（右側優先）。 */
    public static List<RdpLaunchProfile> mergeByNumber(
            List<RdpLaunchProfile> base, List<RdpLaunchProfile> overlay) {
        Map<Integer, RdpLaunchProfile> merged = new LinkedHashMap<>();
        for (RdpLaunchProfile profile : normalizeList(base)) {
            merged.put(profile.number(), profile);
        }
        for (RdpLaunchProfile profile : normalizeList(overlay)) {
            merged.put(profile.number(), profile);
        }
        return new ArrayList<>(merged.values());
    }

    /** ini のスロット数に合わせてプロファイル行数を確保する。 */
    public static List<RdpLaunchProfile> ensureCount(List<RdpLaunchProfile> profiles, int count) {
        Map<Integer, RdpLaunchProfile> byNumber = new LinkedHashMap<>();
        if (profiles != null) {
            for (RdpLaunchProfile profile : profiles) {
                byNumber.put(profile.number(), profile);
            }
        }
        int target = Math.min(RdpRemoteLauncherIni.MAX_SLOTS, Math.max(1, count));
        List<RdpLaunchProfile> out = new ArrayList<>();
        for (int n = 1; n <= target; n++) {
            out.add(byNumber.getOrDefault(n, RdpLaunchProfile.empty(n)));
        }
        return out;
    }

    private static List<RdpLaunchProfile> parseProfiles(JsonNode root) {
        if (root == null || !root.isObject()) {
            return seedEmptyProfiles(3);
        }
        JsonNode profilesNode = root.get("profiles");
        if (profilesNode == null || !profilesNode.isArray()) {
            return seedEmptyProfiles(3);
        }
        Map<Integer, RdpLaunchProfile> byNumber = new LinkedHashMap<>();
        for (JsonNode node : profilesNode) {
            RdpLaunchProfile profile = fromJson(node);
            if (profile != null) {
                byNumber.put(profile.number(), profile);
            }
        }
        if (byNumber.isEmpty()) {
            return seedEmptyProfiles(3);
        }
        int max = byNumber.keySet().stream().mapToInt(Integer::intValue).max().orElse(1);
        return ensureCount(new ArrayList<>(byNumber.values()), max);
    }

    private static RdpLaunchProfile fromJson(JsonNode node) {
        if (node == null || !node.isObject()) {
            return null;
        }
        JsonNode numberNode = node.get("number");
        if (numberNode == null || !numberNode.canConvertToInt()) {
            return null;
        }
        int number = numberNode.intValue();
        if (number < 1 || number > RdpRemoteLauncherIni.MAX_SLOTS) {
            return null;
        }
        RdpSessionEndAction sessionEndAction = null;
        JsonNode sessionEndActionNode = node.get("sessionEndAction");
        if (sessionEndActionNode != null && sessionEndActionNode.isTextual()) {
            sessionEndAction = RdpSessionEndAction.fromProfileJson(sessionEndActionNode.asText());
        }
        return new RdpLaunchProfile(
                number,
                textOrEmpty(node.get("name")),
                textOrEmpty(node.get("description")),
                textOrEmpty(node.get("category")),
                nullableBoolean(node.get("disconnectOnChildExit")),
                sessionEndAction,
                nullableBoolean(node.get("fullScreen")),
                nullableInteger(node.get("desktopWidth")),
                nullableInteger(node.get("desktopHeight")),
                nullableBoolean(node.get("rpaEternal")));
    }

    private static ObjectNode toJson(RdpLaunchProfile profile) {
        ObjectNode node = JSON.createObjectNode();
        node.put("number", profile.number());
        node.put("name", profile.name());
        node.put("description", profile.description());
        node.put("category", profile.category());
        if (profile.sessionEndAction() != null) {
            node.put("sessionEndAction", profile.sessionEndAction().iniValue());
        } else if (profile.disconnectOnChildExit() != null) {
            node.put("disconnectOnChildExit", profile.disconnectOnChildExit());
        }
        if (profile.fullScreen() != null) {
            node.put("fullScreen", profile.fullScreen());
        }
        if (profile.desktopWidth() != null) {
            node.put("desktopWidth", profile.desktopWidth());
        }
        if (profile.desktopHeight() != null) {
            node.put("desktopHeight", profile.desktopHeight());
        }
        if (profile.rpaEternal() != null) {
            node.put("rpaEternal", profile.rpaEternal());
        }
        return node;
    }

    private static List<RdpLaunchProfile> normalizeList(List<RdpLaunchProfile> profiles) {
        if (profiles == null || profiles.isEmpty()) {
            return seedEmptyProfiles(1);
        }
        int maxNumber = profiles.stream().mapToInt(RdpLaunchProfile::number).max().orElse(1);
        return ensureCount(profiles, maxNumber);
    }

    private static List<RdpLaunchProfile> seedEmptyProfiles(int count) {
        List<RdpLaunchProfile> out = new ArrayList<>();
        for (int i = 1; i <= count; i++) {
            out.add(RdpLaunchProfile.empty(i));
        }
        return out;
    }

    private static String textOrEmpty(JsonNode node) {
        if (node == null || node.isNull()) {
            return "";
        }
        return node.asText("").trim();
    }

    private static Boolean nullableBoolean(JsonNode node) {
        if (node == null || node.isNull()) {
            return null;
        }
        if (node.isBoolean()) {
            return node.booleanValue();
        }
        String raw = node.asText("").trim().toLowerCase(java.util.Locale.ROOT);
        return switch (raw) {
            case "1", "true", "on", "yes" -> Boolean.TRUE;
            case "0", "false", "off", "no" -> Boolean.FALSE;
            default -> null;
        };
    }

    private static Integer nullableInteger(JsonNode node) {
        if (node == null || node.isNull()) {
            return null;
        }
        if (node.canConvertToInt()) {
            return node.intValue();
        }
        try {
            return Integer.parseInt(node.asText("").trim());
        } catch (NumberFormatException ex) {
            return null;
        }
    }
}
