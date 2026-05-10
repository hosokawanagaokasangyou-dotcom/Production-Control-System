package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.stream.Stream;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * ユーザープロファイル（{@code ~/.pm-ai-desktop/user-profiles/*.json}）。アップデート同期の対象外で、配布の上書きを受けない。
 */
public final class UserProfileStore {

    public static final int MAX_PROFILES = 100;

    private static final int SCHEMA_VERSION = 1;

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final Path DIRECTORY =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "user-profiles");

    /** 名前省略時の表示・{@code displayName} 保存に使う日時（秒まで）。 */
    private static final DateTimeFormatter STAMP =
            DateTimeFormatter.ofPattern("yyyyMMddHHmmss").withLocale(Locale.JAPAN);

    private UserProfileStore() {}

    public static Path directory() {
        return DIRECTORY;
    }

    /** 一覧表示用（逆順＝新しいファイル名が先頭になりやすい）。 */
    public record ListedProfile(Path path, String displayLabel, String savedAtIso) {}

    public static List<ListedProfile> listProfiles() throws IOException {
        if (!Files.isDirectory(DIRECTORY)) {
            return List.of();
        }
        List<ListedProfile> out = new ArrayList<>();
        try (Stream<Path> stream = Files.list(DIRECTORY)) {
            stream.filter(p -> Files.isRegularFile(p) && p.getFileName().toString().endsWith(".json"))
                    .sorted(Comparator.comparing(p -> p.getFileName().toString(), Comparator.reverseOrder()))
                    .forEach(
                            p -> {
                                try {
                                    JsonNode root = JSON.readTree(p.toFile());
                                    if (root == null || !root.isObject()) {
                                        return;
                                    }
                                    String dn = root.path("displayName").asText("").strip();
                                    String sa = root.path("savedAt").asText("").strip();
                                    String label =
                                            dn.isEmpty()
                                                    ? (sa.isEmpty() ? p.getFileName().toString() : sa)
                                                    : dn;
                                    out.add(new ListedProfile(p, label, sa));
                                } catch (IOException ignored) {
                                }
                            });
        }
        return List.copyOf(out);
    }

    public static int countProfiles() throws IOException {
        if (!Files.isDirectory(DIRECTORY)) {
            return 0;
        }
        try (Stream<Path> stream = Files.list(DIRECTORY)) {
            return (int)
                    stream.filter(p -> Files.isRegularFile(p) && p.getFileName().toString().endsWith(".json"))
                            .count();
        }
    }

    /**
     * @param optionalDisplayName 空なら {@link #STAMP} に相当する文字列を名前として保存
     * @return 書き込んだファイルパス
     */
    public static Path saveProfile(String optionalDisplayName, ObjectNode sessionJson, JsonNode tableColumnRoot)
            throws IOException {
        if (countProfiles() >= MAX_PROFILES) {
            throw new IllegalStateException("プロファイルは最大 " + MAX_PROFILES + " 個までです。");
        }
        Files.createDirectories(DIRECTORY);
        String stamp = STAMP.format(LocalDateTime.now());
        String displayStored =
                optionalDisplayName != null && !optionalDisplayName.isBlank()
                        ? optionalDisplayName.strip()
                        : stamp;
        String base =
                optionalDisplayName != null && !optionalDisplayName.isBlank()
                        ? stamp + "_" + sanitizeFileSegment(optionalDisplayName.strip())
                        : stamp;
        Path dest = uniqueDestination(base);
        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", SCHEMA_VERSION);
        root.put("savedAt", LocalDateTime.now().toString());
        root.put("displayName", displayStored);
        root.set("session", sessionJson != null ? sessionJson.deepCopy() : JSON.createObjectNode());
        root.set(
                "tableColumnOrder",
                tableColumnRoot != null && tableColumnRoot.isObject()
                        ? tableColumnRoot.deepCopy()
                        : JSON.createObjectNode());
        JSON.writerWithDefaultPrettyPrinter().writeValue(dest.toFile(), root);
        return dest;
    }

    private static Path uniqueDestination(String base) {
        Path p = DIRECTORY.resolve(base + ".json");
        if (!Files.isRegularFile(p)) {
            return p;
        }
        for (int i = 2; i < 10_000; i++) {
            Path q = DIRECTORY.resolve(base + "_" + i + ".json");
            if (!Files.isRegularFile(q)) {
                return q;
            }
        }
        throw new IllegalStateException("ユニークなプロファイル名を生成できませんでした。");
    }

    private static String sanitizeFileSegment(String s) {
        String t = s.replaceAll("[^a-zA-Z0-9._\\-\\u3040-\\u30ff\\u3400-\\u9fff]", "_");
        if (t.length() > 80) {
            t = t.substring(0, 80);
        }
        return t.isEmpty() ? "profile" : t;
    }

    public record LoadedProfile(String displayName, String savedAt, DesktopSessionState session, JsonNode tableColumnOrder) {}

    public static LoadedProfile loadProfile(Path file) throws IOException {
        JsonNode root = JSON.readTree(file.toFile());
        if (root == null || !root.isObject()) {
            throw new IOException("プロファイルが読めません: " + file);
        }
        int ver = root.path("schemaVersion").asInt(0);
        if (ver != SCHEMA_VERSION) {
            throw new IOException("未対応のプロファイル形式です (schemaVersion=" + ver + ")");
        }
        String displayName = root.path("displayName").asText("");
        String savedAt = root.path("savedAt").asText("");
        JsonNode sess = root.get("session");
        DesktopSessionState state = DesktopSessionStateStore.desktopSessionFromStoredJson(sess);
        JsonNode tables = root.get("tableColumnOrder");
        if (tables == null || !tables.isObject()) {
            tables = JSON.createObjectNode();
        }
        return new LoadedProfile(displayName, savedAt, state, tables);
    }

    public static void deleteProfile(Path file) throws IOException {
        Files.deleteIfExists(file);
    }
}
