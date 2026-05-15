package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.databind.JsonNode;

class ExcludeRuleLogicGeminiServiceTest {

    @Test
    void parseModelJsonObject_acceptsMarkdownFence() throws Exception {
        String raw = "```json\n{\"version\": 1, \"mode\": \"always_exclude\"}\n```";
        JsonNode n = ExcludeRuleLogicGeminiService.parseModelJsonObject(raw);
        JsonNode v = ExcludeRuleLogicGeminiService.validateRuleJson(n);
        assertNotNull(v);
        assertEquals(1, v.path("version").asInt());
        assertEquals("always_exclude", v.path("mode").asText());
    }
}
