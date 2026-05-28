package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * 工場別の配台システム操作者名（起動時選択・作成者表示用）。
 *
 * <p>永続化: {@code ~/.pm-ai-desktop/factory-operator-users.json}
 */
public final class FactoryOperatorUserStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final Path DEFAULT_STORE =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "factory-operator-users.json");

    public static final int SCHEMA_VERSION = 1;
    public static final int MAX_NAMES_PER_FACTORY = 50;
    public static final int MAX_NAME_LENGTH = 40;

    public static final List<String> DEFAULT_NAMES =
            List.of("砂田", "古家", "図司", "細川");

    private static volatile String sessionOperatorName = "";

    public record FactoryOperatorUsers(List<String> names, String lastSelected) {

        public FactoryOperatorUsers {
            names = names != null ? List.copyOf(names) : List.of();
            lastSelected = lastSelected != null ? lastSelected.strip() : "";
        }
    }

    private FactoryOperatorUserStore() {}

    public static Path storePath() {
        String test = System.getProperty("pm.ai.test.factoryOperatorUserStore");
        if (test != null && !test.isBlank()) {
            return Path.of(test).toAbsolutePath().normalize();
        }
        return DEFAULT_STORE;
    }

    /** 現在セッションで選択中の操作者名（起動時に設定）。 */
    public static String sessionOperatorName() {
        return sessionOperatorName != null ? sessionOperatorName.strip() : "";
    }

    public static void clearSessionOperatorName() {
        sessionOperatorName = "";
    }

    public static List<String> namesForFactory(FactorySite site) throws IOException {
        return new ArrayList<>(loadFactory(site).names());
    }

    public static String lastSelectedForFactory(FactorySite site) throws IOException {
        return loadFactory(site).lastSelected();
    }

    /**
     * セッション操作者を設定し、工場別 {@code lastSelected} を永続化する。
     *
     * @throws IllegalArgumentException 名前が一覧に無い／空
     */
    public static void selectSessionOperator(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("操作者名が空です。");
        }
        Document doc = loadDocument();
        FactoryOperatorUsers current = doc.factories().get(factory);
        if (current == null || !current.names().contains(normalized)) {
            throw new IllegalArgumentException("操作者名が一覧にありません: " + normalized);
        }
        sessionOperatorName = normalized;
        doc.factories().put(factory, new FactoryOperatorUsers(current.names(), normalized));
        saveDocument(doc);
    }

    public static void addName(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("名前が空です。");
        }
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        if (current.names().contains(normalized)) {
            throw new IllegalArgumentException("同じ名前が既にあります: " + normalized);
        }
        if (current.names().size() >= MAX_NAMES_PER_FACTORY) {
            throw new IllegalArgumentException("名前は工場あたり最大 " + MAX_NAMES_PER_FACTORY + " 件です。");
        }
        List<String> next = new ArrayList<>(current.names());
        next.add(normalized);
        doc.factories().put(factory, new FactoryOperatorUsers(next, current.lastSelected()));
        saveDocument(doc);
    }

    public static void removeName(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        if (current.names().size() <= 1) {
            throw new IllegalStateException("最後の1件は削除できません。");
        }
        if (!current.names().contains(normalized)) {
            return;
        }
        List<String> next = new ArrayList<>(current.names());
        next.remove(normalized);
        String last =
                normalized.equals(current.lastSelected()) ? "" : current.lastSelected();
        doc.factories().put(factory, new FactoryOperatorUsers(next, last));
        if (normalized.equals(sessionOperatorName) && factory == GlobalInitSettingTarget.load()) {
            sessionOperatorName = "";
        }
        saveDocument(doc);
    }

    public static void resetNamesToDefaults(FactorySite site) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        String last =
                DEFAULT_NAMES.contains(current.lastSelected()) ? current.lastSelected() : "";
        doc.factories().put(factory, new FactoryOperatorUsers(DEFAULT_NAMES, last));
        if (!DEFAULT_NAMES.contains(sessionOperatorName) && factory == GlobalInitSettingTarget.load()) {
            sessionOperatorName = "";
        }
        saveDocument(doc);
    }

    private static FactoryOperatorUsers loadFactory(FactorySite site) throws IOException {
        return ensureFactory(loadDocument(), site != null ? site : FactorySite.KONAN);
    }

    private static Document loadDocument() throws IOException {
        Path path = storePath();
        if (!Files.isRegularFile(path)) {
            return defaultDocument();
        }
        JsonNode root = JSON.readTree(path.toFile());
        if (root == null || !root.isObject()) {
            return defaultDocument();
        }
        int ver = root.path("schemaVersion").asInt(0);
        if (ver <= 0) {
            ver = 1;
        }
        if (ver > SCHEMA_VERSION) {
            throw new IOException(
                    "操作者名設定はより新しいアプリ向けです (schemaVersion=" + ver + ")");
        }
        Map<FactorySite, FactoryOperatorUsers> factories = new LinkedHashMap<>();
        JsonNode factoriesNode = root.get("factories");
        if (factoriesNode != null && factoriesNode.isObject()) {
            for (FactorySite site : FactorySite.values()) {
                JsonNode n = factoriesNode.get(site.name());
                if (n != null && n.isObject()) {
                    factories.put(site, parseFactory(n));
                }
            }
        }
        Document doc = new Document(ver, factories);
        for (FactorySite site : FactorySite.values()) {
            ensureFactory(doc, site);
        }
        return doc;
    }

    private static FactoryOperatorUsers parseFactory(JsonNode node) {
        List<String> names = new ArrayList<>();
        JsonNode arr = node.get("names");
        if (arr != null && arr.isArray()) {
            LinkedHashSet<String> dedupe = new LinkedHashSet<>();
            for (JsonNode n : arr) {
                if (n == null || n.isNull()) {
                    continue;
                }
                String v = normalizeName(n.asText(""));
                if (!v.isEmpty()) {
                    dedupe.add(v);
                }
            }
            names.addAll(dedupe);
        }
        if (names.isEmpty()) {
            names.addAll(DEFAULT_NAMES);
        }
        String last = normalizeName(node.path("lastSelected").asText(""));
        if (!last.isEmpty() && !names.contains(last)) {
            last = "";
        }
        return new FactoryOperatorUsers(names, last);
    }

    private static void saveDocument(Document doc) throws IOException {
        Path path = storePath();
        if (path.getParent() != null) {
            Files.createDirectories(path.getParent());
        }
        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", SCHEMA_VERSION);
        ObjectNode factories = root.putObject("factories");
        for (Map.Entry<FactorySite, FactoryOperatorUsers> e : doc.factories().entrySet()) {
            ObjectNode fo = factories.putObject(e.getKey().name());
            ArrayNode arr = fo.putArray("names");
            for (String name : e.getValue().names()) {
                arr.add(name);
            }
            fo.put("lastSelected", e.getValue().lastSelected());
        }
        JSON.writerWithDefaultPrettyPrinter().writeValue(path.toFile(), root);
    }

    private static Document defaultDocument() {
        Map<FactorySite, FactoryOperatorUsers> factories = new LinkedHashMap<>();
        for (FactorySite site : FactorySite.values()) {
            factories.put(site, new FactoryOperatorUsers(DEFAULT_NAMES, ""));
        }
        return new Document(SCHEMA_VERSION, factories);
    }

    private static FactoryOperatorUsers ensureFactory(Document doc, FactorySite site) {
        FactoryOperatorUsers current = doc.factories().get(site);
        if (current != null && !current.names().isEmpty()) {
            return current;
        }
        FactoryOperatorUsers created = new FactoryOperatorUsers(DEFAULT_NAMES, "");
        doc.factories().put(site, created);
        return created;
    }

    private static String normalizeName(String raw) {
        if (raw == null) {
            return "";
        }
        String t = raw.strip();
        if (t.length() > MAX_NAME_LENGTH) {
            t = t.substring(0, MAX_NAME_LENGTH);
        }
        return t;
    }

    private record Document(int schemaVersion, Map<FactorySite, FactoryOperatorUsers> factories) {

        Document {
            factories = factories != null ? new LinkedHashMap<>(factories) : new LinkedHashMap<>();
        }
    }

    /** テスト用: ストアを既定状態へ戻す。 */
    public static void resetStoreForTests() throws IOException {
        sessionOperatorName = "";
        Path path = storePath();
        Files.deleteIfExists(path);
    }

    /** テスト用: ファイルを直接書き込む。 */
    public static void writeRawJsonForTests(String json) throws IOException {
        Path path = storePath();
        if (path.getParent() != null) {
            Files.createDirectories(path.getParent());
        }
        Files.writeString(path, Objects.requireNonNull(json), StandardCharsets.UTF_8);
    }
}
