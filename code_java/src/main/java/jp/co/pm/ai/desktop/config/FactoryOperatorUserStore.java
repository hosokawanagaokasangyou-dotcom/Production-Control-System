package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.security.SecureRandom;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HexFormat;
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
 * 工場別の配台システム操作者名（起動時選択・作成者表示用）と PIN（4～10 桁数字）。
 *
 * <p>永続化: {@link AppPaths#factoryOperatorUsersStorePath}（サマリ Excel と同一フォルダの
 * {@link AppPaths#FACTORY_OPERATOR_USERS_BIN}、バイナリ形式）。旧 {@code ~/.pm-ai-desktop/factory-operator-users.json}
 * から初回読込時に移行する。
 */
public final class FactoryOperatorUserStore {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final SecureRandom SECURE_RANDOM = new SecureRandom();
    private static final HexFormat HEX = HexFormat.of();

    private static final byte[] BINARY_MAGIC = {'P', 'M', 'O', 'U'};
    private static final int BINARY_FORMAT_VERSION = 1;

    private static final Path LEGACY_JSON_STORE =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "factory-operator-users.json");

    private static volatile Path configuredStorePath;
    private static volatile boolean storeConfigured;

    public static final int SCHEMA_VERSION = 3;
    public static final int MAX_NAMES_PER_FACTORY = 50;
    public static final int MAX_NAME_LENGTH = 40;
    public static final int MIN_PIN_LENGTH = 4;
    public static final int MAX_PIN_LENGTH = 10;

    /** @deprecated {@link #MIN_PIN_LENGTH} を使用 */
    @Deprecated
    public static final int PIN_LENGTH = MIN_PIN_LENGTH;
    public static final int MAX_CONSECUTIVE_PIN_FAILURES = 20;

    /** ユーザー管理者タブを開くための管理者パスワード（平文）。 */
    public static final String ADMIN_TAB_PASSWORD = "nagaoka123";

    public static final List<String> DEFAULT_NAMES =
            List.of("砂田", "古家", "図司", "細川");

    private static volatile String sessionOperatorName = "";

    public enum PinVerificationResult {
        SUCCESS,
        NO_PIN_REQUIRED,
        WRONG_PIN,
        LOCKED,
        INVALID_PIN
    }

    public record FactoryOperatorUsers(
            List<String> names,
            String lastSelected,
            Map<String, String> pinHashes,
            Map<String, Integer> pinFailedAttempts) {

        public FactoryOperatorUsers {
            names = names != null ? List.copyOf(names) : List.of();
            lastSelected = lastSelected != null ? lastSelected.strip() : "";
            pinHashes = pinHashes != null ? Map.copyOf(pinHashes) : Map.of();
            pinFailedAttempts = pinFailedAttempts != null ? Map.copyOf(pinFailedAttempts) : Map.of();
        }

        public FactoryOperatorUsers(List<String> names, String lastSelected) {
            this(names, lastSelected, Map.of(), Map.of());
        }

        public FactoryOperatorUsers(List<String> names, String lastSelected, Map<String, String> pinHashes) {
            this(names, lastSelected, pinHashes, Map.of());
        }
    }

    private FactoryOperatorUserStore() {}

    /** {@link AppPaths#summaryAiDispatchXlsxPath} と同じフォルダへストアパスを解決する。 */
    public static synchronized void configureFromUi(Map<String, String> ui) {
        Path next = AppPaths.factoryOperatorUsersStorePath(ui != null ? ui : Map.of());
        if (storeConfigured && next.equals(configuredStorePath)) {
            return;
        }
        configuredStorePath = next;
        storeConfigured = true;
    }

    public static Path storePath() {
        String test = System.getProperty("pm.ai.test.factoryOperatorUserStore");
        if (test != null && !test.isBlank()) {
            return Path.of(test).toAbsolutePath().normalize();
        }
        if (storeConfigured && configuredStorePath != null) {
            return configuredStorePath;
        }
        return AppPaths.factoryOperatorUsersStorePath(Map.of());
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

    public static boolean hasPin(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            return false;
        }
        String hash = loadFactory(factory).pinHashes().get(normalized);
        return hash != null && !hash.isBlank();
    }

    public static boolean isPinLocked(FactorySite site, String name) throws IOException {
        return pinFailedAttemptCount(site, name) >= MAX_CONSECUTIVE_PIN_FAILURES;
    }

    public static int pinFailedAttemptCount(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            return 0;
        }
        Integer count = loadFactory(factory).pinFailedAttempts().get(normalized);
        return count != null ? Math.max(0, count) : 0;
    }

    public static int remainingPinAttempts(FactorySite site, String name) throws IOException {
        int failed = pinFailedAttemptCount(site, name);
        if (failed >= MAX_CONSECUTIVE_PIN_FAILURES) {
            return 0;
        }
        return MAX_CONSECUTIVE_PIN_FAILURES - failed;
    }

    /** 副作用なしの PIN 照合（失敗回数は増やさない）。 */
    public static boolean verifyPin(FactorySite site, String name, String pin) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        String pinNorm = normalizePin(pin);
        if (normalized.isEmpty() || pinNorm == null) {
            return false;
        }
        String expected = loadFactory(factory).pinHashes().get(normalized);
        if (expected == null || expected.isBlank()) {
            return true;
        }
        return MessageDigest.isEqual(
                expected.getBytes(StandardCharsets.UTF_8),
                hashPin(factory, normalized, pinNorm).getBytes(StandardCharsets.UTF_8));
    }

    /**
     * 起動時 PIN 入力向け。連続失敗を記録し、{@link #MAX_CONSECUTIVE_PIN_FAILURES} 回でロックする。
     */
    public static PinVerificationResult verifyPinAttempt(FactorySite site, String name, String pin)
            throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            return PinVerificationResult.INVALID_PIN;
        }
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        if (!current.names().contains(normalized)) {
            return PinVerificationResult.INVALID_PIN;
        }
        int failures = current.pinFailedAttempts().getOrDefault(normalized, 0);
        if (failures >= MAX_CONSECUTIVE_PIN_FAILURES) {
            return PinVerificationResult.LOCKED;
        }
        String expected = current.pinHashes().get(normalized);
        if (expected == null || expected.isBlank()) {
            return PinVerificationResult.NO_PIN_REQUIRED;
        }
        String pinNorm = normalizePin(pin);
        if (pinNorm == null) {
            return PinVerificationResult.INVALID_PIN;
        }
        if (MessageDigest.isEqual(
                expected.getBytes(StandardCharsets.UTF_8),
                hashPin(factory, normalized, pinNorm).getBytes(StandardCharsets.UTF_8))) {
            clearPinFailures(doc, factory, normalized);
            saveDocument(doc);
            return PinVerificationResult.SUCCESS;
        }
        int nextFailures = failures + 1;
        Map<String, Integer> attempts = new LinkedHashMap<>(current.pinFailedAttempts());
        attempts.put(normalized, nextFailures);
        doc.factories()
                .put(
                        factory,
                        new FactoryOperatorUsers(
                                current.names(),
                                current.lastSelected(),
                                current.pinHashes(),
                                attempts));
        saveDocument(doc);
        return nextFailures >= MAX_CONSECUTIVE_PIN_FAILURES
                ? PinVerificationResult.LOCKED
                : PinVerificationResult.WRONG_PIN;
    }

    /** 管理者が PIN ロックを解除する（PIN 自体は変更しない）。 */
    public static void unlockPin(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("名前が空です。");
        }
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        if (!current.names().contains(normalized)) {
            throw new IllegalArgumentException("操作者名が一覧にありません: " + normalized);
        }
        clearPinFailures(doc, factory, normalized);
        saveDocument(doc);
    }

    public static String pinLengthRangeDescriptionJa() {
        return MIN_PIN_LENGTH + "～" + MAX_PIN_LENGTH + "桁の数字";
    }

    /**
     * ログイン中の操作者が自分の PIN を変更する（未設定のときは新規設定）。
     *
     * @throws IllegalStateException 自分以外／ロック中
     * @throws IllegalArgumentException PIN 形式不正・現在 PIN 不一致
     */
    public static void changePinByUser(
            FactorySite site, String name, String currentPin, String newPin) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("名前が空です。");
        }
        if (!normalized.equals(sessionOperatorName())) {
            throw new IllegalStateException("自分の PIN のみ変更できます。");
        }
        if (isPinLocked(factory, normalized)) {
            throw new IllegalStateException(
                    "PIN がロックされています。ユーザー管理者タブでロック解除してください。");
        }
        String newPinNorm = normalizePin(newPin);
        if (newPinNorm == null) {
            throw new IllegalArgumentException("新しい PIN は " + pinLengthRangeDescriptionJa() + " です。");
        }
        if (hasPin(factory, normalized)) {
            String currentNorm = normalizePin(currentPin);
            if (currentNorm == null || !verifyPin(factory, normalized, currentNorm)) {
                throw new IllegalArgumentException("現在の PIN が正しくありません。");
            }
        }
        assignPin(factory, normalized, newPinNorm);
    }

    /**
     * 管理者向け: ランダム 4 桁 PIN を新規発行する（既存 PIN があれば上書き）。
     *
     * @return 発行した PIN（このタイミングのみ平文表示可能）
     */
    public static String issuePin(FactorySite site, String name) throws IOException {
        return assignPin(site, name, generatePin());
    }

    /**
     * 管理者向け: {@link #issuePin} と同じだが意図を明示する別名。
     */
    public static String reissuePin(FactorySite site, String name) throws IOException {
        return issuePin(site, name);
    }

    private static String assignPin(FactorySite site, String name, String pin) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        String pinNorm = normalizePin(pin);
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("名前が空です。");
        }
        if (pinNorm == null) {
            throw new IllegalArgumentException("PIN は " + pinLengthRangeDescriptionJa() + " です。");
        }
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        if (!current.names().contains(normalized)) {
            throw new IllegalArgumentException("操作者名が一覧にありません: " + normalized);
        }
        Map<String, String> pins = new LinkedHashMap<>(current.pinHashes());
        pins.put(normalized, hashPin(factory, normalized, pinNorm));
        Map<String, Integer> attempts = new LinkedHashMap<>(current.pinFailedAttempts());
        attempts.remove(normalized);
        doc.factories()
                .put(
                        factory,
                        new FactoryOperatorUsers(
                                current.names(), current.lastSelected(), pins, attempts));
        saveDocument(doc);
        return pinNorm;
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
        doc.factories()
                .put(
                        factory,
                        new FactoryOperatorUsers(
                                current.names(),
                                normalized,
                                current.pinHashes(),
                                current.pinFailedAttempts()));
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
        doc.factories()
                .put(
                        factory,
                        new FactoryOperatorUsers(
                                next,
                                current.lastSelected(),
                                current.pinHashes(),
                                current.pinFailedAttempts()));
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
        Map<String, String> pins = new LinkedHashMap<>(current.pinHashes());
        pins.remove(normalized);
        Map<String, Integer> attempts = new LinkedHashMap<>(current.pinFailedAttempts());
        attempts.remove(normalized);
        doc.factories().put(factory, new FactoryOperatorUsers(next, last, pins, attempts));
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
        Map<String, String> pins = new LinkedHashMap<>();
        Map<String, Integer> attempts = new LinkedHashMap<>();
        for (String n : DEFAULT_NAMES) {
            String h = current.pinHashes().get(n);
            if (h != null && !h.isBlank()) {
                pins.put(n, h);
            }
            Integer failed = current.pinFailedAttempts().get(n);
            if (failed != null && failed > 0) {
                attempts.put(n, failed);
            }
        }
        doc.factories().put(factory, new FactoryOperatorUsers(DEFAULT_NAMES, last, pins, attempts));
        if (!DEFAULT_NAMES.contains(sessionOperatorName) && factory == GlobalInitSettingTarget.load()) {
            sessionOperatorName = "";
        }
        saveDocument(doc);
    }

    /** 一覧表示用: PIN / ロック状態。 */
    public static String pinStatusLabel(FactorySite site, String name) throws IOException {
        if (isPinLocked(site, name)) {
            return "ロック";
        }
        return hasPin(site, name) ? "設定済" : "未設定";
    }

    private static FactoryOperatorUsers loadFactory(FactorySite site) throws IOException {
        return ensureFactory(loadDocument(), site != null ? site : FactorySite.KONAN);
    }

    private static Document loadDocument() throws IOException {
        Path path = storePath();
        migrateLegacyStoreIfNeeded(path);
        if (!Files.isRegularFile(path)) {
            return defaultDocument();
        }
        JsonNode root = readStoreRoot(path);
        if (root == null || !root.isObject()) {
            return defaultDocument();
        }
        return parseDocumentRoot(root);
    }

    private static JsonNode readStoreRoot(Path path) throws IOException {
        byte[] bytes = Files.readAllBytes(path);
        if (isBinaryStore(bytes)) {
            return decodeBinaryPayload(bytes);
        }
        return JSON.readTree(bytes);
    }

    private static Document parseDocumentRoot(JsonNode root) throws IOException {
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

    private static void migrateLegacyStoreIfNeeded(Path targetBin) throws IOException {
        if (Files.isRegularFile(targetBin)) {
            return;
        }
        Path legacySibling = targetBin.resolveSibling("factory-operator-users.json");
        for (Path legacy : List.of(resolveLegacyJsonStorePath(), legacySibling)) {
            if (!Files.isRegularFile(legacy)) {
                continue;
            }
            JsonNode root = JSON.readTree(legacy.toFile());
            Document doc = parseDocumentRoot(root);
            saveDocumentToPath(targetBin, doc);
            return;
        }
    }

    private static boolean isBinaryStore(byte[] bytes) {
        return bytes.length >= BINARY_MAGIC.length
                && Arrays.compare(
                                Arrays.copyOf(bytes, BINARY_MAGIC.length),
                                BINARY_MAGIC)
                        == 0;
    }

    private static JsonNode decodeBinaryPayload(byte[] fileBytes) throws IOException {
        ByteBuffer buf = ByteBuffer.wrap(fileBytes);
        byte[] magic = new byte[BINARY_MAGIC.length];
        buf.get(magic);
        if (!Arrays.equals(magic, BINARY_MAGIC)) {
            throw new IOException("操作者名設定ファイルの形式が不正です。");
        }
        int formatVersion = buf.getShort() & 0xffff;
        if (formatVersion != BINARY_FORMAT_VERSION) {
            throw new IOException(
                    "操作者名設定のバイナリ形式が未対応です (formatVersion=" + formatVersion + ")");
        }
        int payloadLen = buf.getInt();
        if (payloadLen < 0 || buf.remaining() < payloadLen) {
            throw new IOException("操作者名設定ファイルが壊れています。");
        }
        byte[] payload = new byte[payloadLen];
        buf.get(payload);
        return JSON.readTree(payload);
    }

    private static byte[] encodeBinaryDocument(ObjectNode root) throws IOException {
        byte[] payload = JSON.writeValueAsBytes(root);
        ByteBuffer buf = ByteBuffer.allocate(BINARY_MAGIC.length + 2 + 4 + payload.length);
        buf.put(BINARY_MAGIC);
        buf.putShort((short) BINARY_FORMAT_VERSION);
        buf.putInt(payload.length);
        buf.put(payload);
        return buf.array();
    }

    private static Path resolveLegacyJsonStorePath() {
        String test = System.getProperty("pm.ai.test.factoryOperatorUserLegacyStore");
        if (test != null && !test.isBlank()) {
            return Path.of(test).toAbsolutePath().normalize();
        }
        return LEGACY_JSON_STORE;
    }

    private static ObjectNode documentToObjectNode(Document doc) {
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
            ObjectNode pins = fo.putObject("pinHashes");
            for (Map.Entry<String, String> pe : e.getValue().pinHashes().entrySet()) {
                if (e.getValue().names().contains(pe.getKey())) {
                    pins.put(pe.getKey(), pe.getValue());
                }
            }
            ObjectNode attempts = fo.putObject("pinFailedAttempts");
            for (Map.Entry<String, Integer> ae : e.getValue().pinFailedAttempts().entrySet()) {
                if (e.getValue().names().contains(ae.getKey()) && ae.getValue() != null && ae.getValue() > 0) {
                    attempts.put(ae.getKey(), ae.getValue());
                }
            }
        }
        return root;
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
        Map<String, String> pinHashes = new LinkedHashMap<>();
        JsonNode pinsNode = node.get("pinHashes");
        if (pinsNode != null && pinsNode.isObject()) {
            pinsNode
                    .fields()
                    .forEachRemaining(
                            e -> {
                                String key = normalizeName(e.getKey());
                                if (key.isEmpty() || !names.contains(key)) {
                                    return;
                                }
                                JsonNode v = e.getValue();
                                if (v != null && v.isTextual()) {
                                    String hash = v.asText("").strip();
                                    if (!hash.isEmpty()) {
                                        pinHashes.put(key, hash);
                                    }
                                }
                            });
        }
        Map<String, Integer> pinFailedAttempts = new LinkedHashMap<>();
        JsonNode attemptsNode = node.get("pinFailedAttempts");
        if (attemptsNode != null && attemptsNode.isObject()) {
            attemptsNode
                    .fields()
                    .forEachRemaining(
                            e -> {
                                String key = normalizeName(e.getKey());
                                if (key.isEmpty() || !names.contains(key)) {
                                    return;
                                }
                                JsonNode v = e.getValue();
                                if (v != null && v.isNumber()) {
                                    int count = v.asInt(0);
                                    if (count > 0) {
                                        pinFailedAttempts.put(key, count);
                                    }
                                }
                            });
        }
        return new FactoryOperatorUsers(names, last, pinHashes, pinFailedAttempts);
    }

    private static void saveDocument(Document doc) throws IOException {
        saveDocumentToPath(storePath(), doc);
    }

    private static void saveDocumentToPath(Path path, Document doc) throws IOException {
        if (path.getParent() != null) {
            Files.createDirectories(path.getParent());
        }
        byte[] encoded = encodeBinaryDocument(documentToObjectNode(doc));
        Files.write(
                path,
                encoded,
                StandardOpenOption.CREATE,
                StandardOpenOption.TRUNCATE_EXISTING,
                StandardOpenOption.WRITE);
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

    public static String normalizePin(String raw) {
        if (raw == null) {
            return null;
        }
        String t = raw.strip();
        if (t.length() < MIN_PIN_LENGTH || t.length() > MAX_PIN_LENGTH || !t.chars().allMatch(Character::isDigit)) {
            return null;
        }
        return t;
    }

    private static String generatePin() {
        int n = SECURE_RANDOM.nextInt(10_000);
        return String.format("%04d", n);
    }

    private static void clearPinFailures(Document doc, FactorySite factory, String normalized) {
        FactoryOperatorUsers current = doc.factories().get(factory);
        if (current == null || !current.pinFailedAttempts().containsKey(normalized)) {
            return;
        }
        Map<String, Integer> attempts = new LinkedHashMap<>(current.pinFailedAttempts());
        attempts.remove(normalized);
        doc.factories()
                .put(
                        factory,
                        new FactoryOperatorUsers(
                                current.names(),
                                current.lastSelected(),
                                current.pinHashes(),
                                attempts));
    }

    private static String hashPin(FactorySite factory, String name, String pin) {
        String payload =
                factory.name() + "|" + normalizeName(name) + "|" + Objects.requireNonNull(pin);
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            byte[] digest = md.digest(payload.getBytes(StandardCharsets.UTF_8));
            return HEX.formatHex(digest);
        } catch (NoSuchAlgorithmException ex) {
            throw new IllegalStateException(ex);
        }
    }

    private record Document(int schemaVersion, Map<FactorySite, FactoryOperatorUsers> factories) {

        Document {
            factories = factories != null ? new LinkedHashMap<>(factories) : new LinkedHashMap<>();
        }
    }

    /** テスト用: ストアを既定状態へ戻す。 */
    public static void resetStoreForTests() throws IOException {
        sessionOperatorName = "";
        configuredStorePath = null;
        storeConfigured = false;
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
