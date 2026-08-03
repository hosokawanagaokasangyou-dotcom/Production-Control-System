package jp.co.pm.ai.desktop.benchmark;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.junit.jupiter.api.Test;

/**
 * 思考トークンは応答待ち時間に直結するため、抽出系の REST 呼び出しでは既定で無効にする。
 */
class GeminiGenerateContentRestClientRequestJsonTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

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

    private static JsonNode generationConfig(String json) throws Exception {
        return MAPPER.readTree(json).path("generationConfig");
    }
}
