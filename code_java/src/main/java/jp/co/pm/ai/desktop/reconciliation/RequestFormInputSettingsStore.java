package jp.co.pm.ai.desktop.reconciliation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.DesktopSessionStateStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Map;
import java.util.Optional;
import java.util.Properties;

/**
 * 依頼書入力タブの個人設定（ComboBox 候補・原本フォルダ・受注ファイルパス）を、サマリ AI 配台 Excel と
 * 同一フォルダの {@link AppPaths#REQUEST_FORM_INPUT_SETTINGS_JSON_FILENAME} に保存する。
 */
public final class RequestFormInputSettingsStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String KEY_TARGET_FOLDER = "targetFolder";
    private static final String KEY_JUCHU_FILE_PATH = "juchuFilePath";

    /** 移行元（旧バージョンのユーザーホーム配置）。 */
    private static final Path LEGACY_PROPERTIES_FILE =
            Path.of(System.getProperty("user.home"), ".pm-ai-desktop", "request-form-reconciliation.properties");

    private RequestFormInputSettingsStore() {}

    /**
     * 受注ファイルの実効パス。環境変数 {@link AppPaths#KEY_PM_AI_REQUEST_FORM_JUCHU_FILE} が非空ならそれを正とし、
     * 空のときだけ設定 JSON を使う（利用工場と矛盾するパスは捨てて工場既定へ）。
     */
    public static String resolveEffectiveJuchuFilePath(Map<String, String> ui) {
        Map<String, String> env = ui != null ? ui : Map.of();
        FactorySite site = GlobalInitSettingTarget.loadEffective(env);
        String fromEnv = textOrEmpty(env.get(AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE));
        if (!fromEnv.isEmpty()) {
            return Path.of(fromEnv).toAbsolutePath().normalize().toString();
        }
        Optional<Settings> settings = load(env);
        if (settings.isPresent()) {
            String saved = settings.get().paths().juchuFilePath();
            if (saved != null
                    && !saved.isBlank()
                    && !AppPaths.factoryPathHintConflictsWithSite(saved, site)) {
                return Path.of(saved.strip()).toAbsolutePath().normalize().toString();
            }
        }
        return Path.of(AppPaths.defaultRequestFormJuchuFileForFactory(site))
                .toAbsolutePath()
                .normalize()
                .toString();
    }

    private static String textOrEmpty(String value) {
        return value != null ? value.strip() : "";
    }

    /** 原本フォルダ（保存値があれば UI 環境より優先して復元）。受注ファイルパスは {@link #resolveEffectiveJuchuFilePath}。 */
    public record ReconciliationPaths(String targetFolder, String juchuFilePath) {}

    public record Settings(RequestFormComboChoices comboChoices, ReconciliationPaths paths) {

        public static Settings empty() {
            return new Settings(RequestFormComboChoices.empty(), new ReconciliationPaths("", ""));
        }
    }

    public static Path resolveStorePath(Map<String, String> ui) {
        return AppPaths.requestFormInputSettingsJsonPath(ui);
    }

    public static Optional<Settings> load(Map<String, String> ui) {
        Optional<Settings> fromPrimary = loadFromFile(resolveStorePath(ui));
        if (fromPrimary.isPresent()) {
            return fromPrimary;
        }
        Path legacyPath = AppPaths.requestFormInputSettingsJsonPathLegacy(ui);
        Path primaryPath = resolveStorePath(ui);
        if (!legacyPath.equals(primaryPath)) {
            fromPrimary = loadFromFile(legacyPath);
            if (fromPrimary.isPresent()) {
                return fromPrimary;
            }
        }
        return loadLegacyMigration(ui);
    }

    private static Optional<Settings> loadFromFile(Path storePath) {
        if (!Files.isRegularFile(storePath)) {
            return Optional.empty();
        }
        try {
            JsonNode root = JSON.readTree(storePath.toFile());
            return Optional.of(parseRoot(root));
        } catch (IOException ex) {
            System.err.println(
                    "Could not load request form input settings: " + storePath + " — " + ex.getMessage());
            return Optional.empty();
        }
    }

    public static RequestFormComboChoices loadComboChoices(Map<String, String> ui, FactorySite site) {
        Optional<Settings> settings = load(ui);
        if (settings.isPresent() && !settings.get().comboChoices().isEmpty()) {
            return settings.get().comboChoices().mergedWithDefaults();
        }
        return DesktopSessionStateStore.loadFactoryRequestFormComboChoices(ui, site).mergedWithDefaults();
    }

    public static void save(Map<String, String> ui, Settings settings) throws IOException {
        if (settings == null) {
            return;
        }
        Path storePath = resolveStorePath(ui);
        Path parent = storePath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        ObjectNode root = readExistingObjectOrEmpty(storePath);
        if (settings.comboChoices() != null && !settings.comboChoices().isEmpty()) {
            settings.comboChoices().mergeIntoSettingsRoot(root);
        }
        ReconciliationPaths paths = settings.paths();
        if (paths != null) {
            if (paths.targetFolder() != null && !paths.targetFolder().isBlank()) {
                root.put(KEY_TARGET_FOLDER, paths.targetFolder().strip());
            }
            if (paths.juchuFilePath() != null) {
                root.put(KEY_JUCHU_FILE_PATH, paths.juchuFilePath().strip());
            }
        }
        writeJsonAtomically(storePath, root);
    }

    public static void save(
            Map<String, String> ui,
            RequestFormComboChoices comboChoices,
            String targetFolder,
            String juchuFilePath)
            throws IOException {
        save(
                ui,
                new Settings(
                        comboChoices != null ? comboChoices : RequestFormComboChoices.empty(),
                        new ReconciliationPaths(
                                targetFolder != null ? targetFolder : "",
                                juchuFilePath != null ? juchuFilePath : "")));
    }

    /**
     * 設定タブの JSON 直接編集用。ファイルがあれば内容を返す。構文が壊れていても生テキストを返し、
     * 正常なら pretty-print する。未作成なら空オブジェクト。
     */
    public static String readTextForEditor(Map<String, String> ui) throws IOException {
        Path storePath = resolveStorePath(ui);
        if (!Files.isRegularFile(storePath)) {
            Path legacyPath = AppPaths.requestFormInputSettingsJsonPathLegacy(ui);
            if (legacyPath != null && !legacyPath.equals(storePath) && Files.isRegularFile(legacyPath)) {
                storePath = legacyPath;
            } else {
                return JSON.writerWithDefaultPrettyPrinter().writeValueAsString(JSON.createObjectNode());
            }
        }
        String raw = Files.readString(storePath, StandardCharsets.UTF_8);
        try {
            JsonNode root = JSON.readTree(raw);
            return JSON.writerWithDefaultPrettyPrinter().writeValueAsString(root);
        } catch (IOException ex) {
            return raw;
        }
    }

    /**
     * 設定タブの JSON 直接編集の保存。ルートは JSON オブジェクトのみ。pretty-print して上書きする。
     *
     * @throws IOException 構文不正・ルートがオブジェクトでない・書込失敗
     */
    public static Settings savePrettyJson(Map<String, String> ui, String rawText) throws IOException {
        if (rawText == null || rawText.isBlank()) {
            throw new IOException("JSON が空です。");
        }
        JsonNode root;
        try {
            root = JSON.readTree(rawText);
        } catch (IOException ex) {
            throw new IOException("JSON の構文が不正です: " + ex.getMessage(), ex);
        }
        if (root == null || !root.isObject()) {
            throw new IOException("JSON のルートはオブジェクトである必要があります。");
        }
        Path storePath = resolveStorePath(ui);
        Path parent = storePath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        writeJsonAtomically(storePath, (ObjectNode) root);
        return parseRoot(root);
    }

    private static ObjectNode readExistingObjectOrEmpty(Path storePath) {
        if (!Files.isRegularFile(storePath)) {
            return JSON.createObjectNode();
        }
        try {
            JsonNode existing = JSON.readTree(storePath.toFile());
            if (existing != null && existing.isObject()) {
                return (ObjectNode) existing;
            }
        } catch (IOException ignored) {
            return JSON.createObjectNode();
        }
        return JSON.createObjectNode();
    }

    private static void writeJsonAtomically(Path storePath, ObjectNode root) throws IOException {
        Path parent = storePath.getParent();
        Path staging =
                parent != null
                        ? Files.createTempFile(parent, "rfis-", ".tmp")
                        : Files.createTempFile("rfis-", ".tmp");
        try {
            JSON.writerWithDefaultPrettyPrinter().writeValue(staging.toFile(), root);
            try {
                Files.move(
                        staging,
                        storePath,
                        StandardCopyOption.REPLACE_EXISTING,
                        StandardCopyOption.ATOMIC_MOVE);
            } catch (AtomicMoveNotSupportedException ex) {
                Files.move(staging, storePath, StandardCopyOption.REPLACE_EXISTING);
            }
        } catch (IOException ex) {
            Files.deleteIfExists(staging);
            throw ex;
        }
    }

    private static Settings parseRoot(JsonNode root) {
        if (root == null || !root.isObject()) {
            return Settings.empty();
        }
        RequestFormComboChoices combo = RequestFormComboChoices.fromSettingsFileRoot(root);
        String targetFolder = textOrEmpty(root.get(KEY_TARGET_FOLDER));
        String juchuPath = textOrEmpty(root.get(KEY_JUCHU_FILE_PATH));
        return new Settings(combo, new ReconciliationPaths(targetFolder, juchuPath));
    }

    private static String textOrEmpty(JsonNode node) {
        if (node == null || !node.isTextual()) {
            return "";
        }
        return node.asText("").strip();
    }

    /** 旧 {@code ~/.pm-ai-desktop/request-form-reconciliation.properties} から初回のみ移行する。 */
    private static Optional<Settings> loadLegacyMigration(Map<String, String> ui) {
        if (!Files.isRegularFile(LEGACY_PROPERTIES_FILE)) {
            return Optional.empty();
        }
        Properties props = new Properties();
        try (InputStream in = Files.newInputStream(LEGACY_PROPERTIES_FILE)) {
            props.load(in);
        } catch (IOException ex) {
            return Optional.empty();
        }
        String folder = props.getProperty(KEY_TARGET_FOLDER, "").strip();
        String juchu = props.getProperty(KEY_JUCHU_FILE_PATH, "").strip();
        RequestFormComboChoices combo =
                DesktopSessionStateStore.loadFactoryRequestFormComboChoices(
                        ui, GlobalInitSettingTarget.load());
        Settings migrated =
                new Settings(
                        combo != null ? combo : RequestFormComboChoices.empty(),
                        new ReconciliationPaths(folder, juchu));
        if (!folder.isBlank() || !juchu.isBlank() || !migrated.comboChoices().isEmpty()) {
            try {
                save(ui, migrated);
            } catch (IOException ex) {
                System.err.println(
                        "Could not migrate request form input settings: "
                                + resolveStorePath(ui)
                                + " — "
                                + ex.getMessage());
            }
        }
        return Optional.of(migrated);
    }
}
