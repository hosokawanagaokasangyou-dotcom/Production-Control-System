package jp.co.pm.ai.desktop.benchmark;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

/**
 * 思考トークンは応答待ち時間に直結するため、抽出系の REST 呼び出しでは既定で無効にする。
 */
class GeminiGenerateContentRestClientRequestJsonTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @BeforeEach
    void resetThinkingConfigMemo() {
        GeminiGenerateContentRestClient.forgetThinkingConfigRejections();
    }

    @Test
    void requestDisablesThinkingByDefault() throws Exception {
        JsonNode gen = generationConfig(GeminiGenerateContentRestClient.buildRequestJson("こんにちは", 4096));
        assertEquals(0, gen.path("thinkingConfig").path("thinkingBudget").asInt(-1));
    }

    @Test
    void requestKeepsMaxOutputTokensAndTemperature() throws Exception {
        JsonNode gen = generationConfig(GeminiGenerateContentRestClient.buildRequestJson("x", 64));
        assertEquals(64, gen.path("maxOutputTokens").asInt());
        assertEquals(0.0, gen.path("temperature").asDouble());
    }

    @Test
    void requestClampsMaxOutputTokensIntoSupportedRange() throws Exception {
        assertEquals(
                1,
                generationConfig(GeminiGenerateContentRestClient.buildRequestJson("x", 0))
                        .path("maxOutputTokens")
                        .asInt());
        assertEquals(
                8192,
                generationConfig(GeminiGenerateContentRestClient.buildRequestJson("x", 99999))
                        .path("maxOutputTokens")
                        .asInt());
    }

    @Test
    void requestCarriesPromptText() throws Exception {
        JsonNode root = MAPPER.readTree(GeminiGenerateContentRestClient.buildRequestJson("配台の質問", 32));
        JsonNode parts = root.path("contents").get(0).path("parts");
        assertTrue(parts.isArray());
        assertEquals("配台の質問", parts.get(0).path("text").asText());
    }

    @Test
    void requestOmitsThinkingConfigWhenNotRequested() throws Exception {
        JsonNode gen = generationConfig(GeminiGenerateContentRestClient.buildRequestJson("x", 64, false));
        assertTrue(gen.path("thinkingConfig").isMissingNode());
        assertEquals(64, gen.path("maxOutputTokens").asInt());
    }

    /** gemini-3.5-flash-lite は理由を書かない 400 で thinkingBudget を拒む。 */
    @Test
    void modelThatRejectedThinkingConfigIsRemembered() {
        String model = "gemini-3.5-flash-lite";
        assertFalse(GeminiGenerateContentRestClient.modelRejectsThinkingConfig(model));
        GeminiGenerateContentRestClient.rememberThinkingConfigRejection(model);
        assertTrue(GeminiGenerateContentRestClient.modelRejectsThinkingConfig(model));
        assertFalse(GeminiGenerateContentRestClient.modelRejectsThinkingConfig("gemini-3.5-flash"));
    }

    @Test
    void rejectionMemoIsKeyedOnNormalizedModelId() {
        GeminiGenerateContentRestClient.rememberThinkingConfigRejection("models/gemini-3.5-flash-lite");
        assertTrue(GeminiGenerateContentRestClient.modelRejectsThinkingConfig("gemini-3.5-flash-lite"));
    }

    private static JsonNode generationConfig(String json) throws Exception {
        return MAPPER.readTree(json).path("generationConfig");
    }
}
