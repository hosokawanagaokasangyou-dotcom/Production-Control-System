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
import java.nio.file.Files;
import java.nio.file.Path;
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

    /** 原本フォルダと受注ファイルパス（保存値があれば UI 環境より優先して復元）。 */
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

    public static void save(Map<String, String> ui, Settings settings) {
        if (settings == null) {
            return;
        }
        Path storePath = resolveStorePath(ui);
        try {
            Files.createDirectories(storePath.getParent());
            ObjectNode root = JSON.createObjectNode();
            if (settings.comboChoices() != null && !settings.comboChoices().isEmpty()) {
                settings.comboChoices().writeToObjectNode(root);
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
            JSON.writerWithDefaultPrettyPrinter().writeValue(storePath.toFile(), root);
        } catch (IOException ex) {
            System.err.println(
                    "Could not save request form input settings: " + storePath + " — " + ex.getMessage());
        }
    }

    public static void save(
            Map<String, String> ui,
            RequestFormComboChoices comboChoices,
            String targetFolder,
            String juchuFilePath) {
        save(
                ui,
                new Settings(
                        comboChoices != null ? comboChoices : RequestFormComboChoices.empty(),
                        new ReconciliationPaths(
                                targetFolder != null ? targetFolder : "",
                                juchuFilePath != null ? juchuFilePath : "")));
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
            save(ui, migrated);
        }
        return Optional.of(migrated);
    }
}
