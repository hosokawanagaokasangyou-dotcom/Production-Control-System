package jp.co.pm.ai.desktop.dispatch.rules.simulation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

class DispatchRuleSimulationServiceTest {

    private static final ObjectMapper JSON = new ObjectMapper();

    @Test
    void extractLastJsonObjectLine_ignoresBootstrapLogAfterJson() throws Exception {
        String merged =
                """
                {"final_blocked": true, "summary_ja": "候補から除外", "steps": []}
                2026-05-27 08:00:00,000 - INFO - planning_core bootstrap noise
                """;
        String payload = DispatchRuleSimulationService.extractLastJsonObjectLine(merged);
        JsonNode root = JSON.readTree(payload);
        assertTrue(root.path("final_blocked").asBoolean(false));
        assertEquals("候補から除外", root.path("summary_ja").asText());
    }
}
