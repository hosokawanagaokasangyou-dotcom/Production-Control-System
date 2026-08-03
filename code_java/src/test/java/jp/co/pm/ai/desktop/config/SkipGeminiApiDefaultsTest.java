package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import org.junit.jupiter.api.Test;

/**
 * 段階1・段階2 の AI API 呼び出しは既定で有効（スキップは OFF）。文言に「開発用」を残さない。
 */
class SkipGeminiApiDefaultsTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Test
    void emptySessionEnablesApiCallsForBothStages() {
        DesktopSessionState s = DesktopSessionState.empty();
        assertFalse(s.mainRunSkipGeminiApi());
        assertFalse(s.planInputStage2SkipGeminiApi());
    }

    @Test
    void runTabCheckBoxStartsUncheckedWithoutDevelopmentWording() throws IOException {
        String fxml = fxml("MainRunTab.fxml");
        String checkBox = elementOf(fxml, "skipGeminiApiCheckBox");
        assertTrue(checkBox.contains("selected=\"false\""), checkBox);
        assertTrue(checkBox.contains("text=\"AI API 呼び出しをスキップ\""), checkBox);
    }

    @Test
    void planInputCheckBoxStartsUncheckedWithoutDevelopmentWording() throws IOException {
        String fxml = fxml("PlanInputTab.fxml");
        String checkBox = elementOf(fxml, "stage2SkipGeminiApiCheckBox");
        assertTrue(checkBox.contains("selected=\"false\""), checkBox);
        assertTrue(checkBox.contains("text=\"AI API 呼び出しをスキップ\""), checkBox);
    }

    @Test
    void skipGeminiApiWordingDropsDevelopmentLabel() throws IOException {
        assertFalse(elementOf(fxml("MainRunTab.fxml"), "skipGeminiApiCheckBox").contains("開発用"));
        assertFalse(
                elementOf(fxml("PlanInputTab.fxml"), "stage2SkipGeminiApiCheckBox").contains("開発用"));
        assertFalse(EnvVarDocs.logicOnly(AppPaths.KEY_PM_AI_SKIP_GEMINI_API).contains("開発用"));
    }

    @Test
    void factorySessionDefaultsEnableApiCallsForBothStages() throws IOException {
        for (String name :
                new String[] {"session_defaults_konan.json", "session_defaults_kokubu.json"}) {
            JsonNode root = MAPPER.readTree(initSetting(name).toFile());
            assertFalse(root.path("mainRunSkipGeminiApi").asBoolean(true), name);
            assertFalse(root.path("planInputStage2SkipGeminiApi").asBoolean(true), name);
        }
    }

    @Test
    void storedSessionFromBeforeTheDefaultChangeLosesItsSkipFlagOnce() {
        ObjectNode root = MAPPER.createObjectNode();
        root.put("mainRunSkipGeminiApi", true);
        root.put("planInputStage2SkipGeminiApi", true);

        DesktopSessionState s = DesktopSessionStateStore.desktopSessionFromStoredJson(root);
        assertFalse(s.mainRunSkipGeminiApi());
        assertFalse(s.planInputStage2SkipGeminiApi());
    }

    @Test
    void skipFlagIsHonouredAgainOnceTheMigrationMarkerIsStored() {
        ObjectNode root = MAPPER.createObjectNode();
        root.put("mainRunSkipGeminiApi", true);
        root.put("planInputStage2SkipGeminiApi", true);
        root.put(DesktopSessionStateStore.SKIP_GEMINI_API_DEFAULT_OFF_MIGRATED_KEY, true);

        DesktopSessionState s = DesktopSessionStateStore.desktopSessionFromStoredJson(root);
        assertTrue(s.mainRunSkipGeminiApi());
        assertTrue(s.planInputStage2SkipGeminiApi());
    }

    @Test
    void savedSessionCarriesTheMigrationMarker() {
        ObjectNode root = DesktopSessionStateStore.toJsonObject(DesktopSessionState.empty());
        assertTrue(root.path(DesktopSessionStateStore.SKIP_GEMINI_API_DEFAULT_OFF_MIGRATED_KEY).asBoolean(false));
    }

    @Test
    void bundledEnvDefaultKeepsSkipDisabled() throws IOException {
        try (InputStream in =
                SkipGeminiApiDefaultsTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/ui_ref_env_defaults.json")) {
            assertNotNull(in);
            JsonNode rows = MAPPER.readTree(in.readAllBytes()).path("entries");
            JsonNode hit = null;
            for (JsonNode row : rows) {
                if (AppPaths.KEY_PM_AI_SKIP_GEMINI_API.equals(row.path("key").asText())) {
                    hit = row;
                    break;
                }
            }
            assertNotNull(hit);
            assertTrue("0".equals(hit.path("value").asText()));
            assertFalse(hit.path("description").asText().contains("開発用"));
        }
    }

    /** {@code <CheckBox fx:id="…" …>} 開始タグだけを切り出す（属性の並びに依存しないため）。 */
    private static String elementOf(String fxml, String fxId) {
        int idAt = fxml.indexOf("fx:id=\"" + fxId + "\"");
        assertTrue(idAt >= 0, fxId + " が FXML に見つかりません。");
        int start = fxml.lastIndexOf('<', idAt);
        int end = fxml.indexOf('>', idAt);
        assertTrue(start >= 0 && end > start);
        return fxml.substring(start, end + 1);
    }

    private static String fxml(String name) throws IOException {
        try (InputStream in =
                SkipGeminiApiDefaultsTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/" + name)) {
            assertNotNull(in, name);
            return new String(in.readAllBytes(), StandardCharsets.UTF_8);
        }
    }

    private static Path initSetting(String name) {
        Path base = Path.of(System.getProperty("user.dir"));
        Path repo = base.getParent() != null ? base.getParent() : base;
        Path p = repo.resolve("init_setting").resolve(name);
        assertTrue(Files.isRegularFile(p), p.toString());
        return p;
    }
}
