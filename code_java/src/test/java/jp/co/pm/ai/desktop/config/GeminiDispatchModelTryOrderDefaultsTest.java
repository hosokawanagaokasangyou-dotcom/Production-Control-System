package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class GeminiDispatchModelTryOrderDefaultsTest {

    @Test
    void resolveEffectiveModelTryOrder_usesTryOrderCsvWhenUnpinned() {
        Map<String, String> env =
                Map.of(
                        GeminiDispatchModelTryOrderDefaults.ENV_GEMINI_MODEL_TRY_ORDER,
                        "gemini-3.5-flash, gemini-3.1-flash-lite");
        List<String> models = GeminiDispatchModelTryOrderDefaults.resolveEffectiveModelTryOrder(env);
        assertEquals(List.of("gemini-3.5-flash", "gemini-3.1-flash-lite"), models);
    }

    @Test
    void resolveEffectiveModelTryOrder_prefersPinnedModel() {
        Map<String, String> env =
                Map.of(
                        GeminiDispatchModelTryOrderDefaults.ENV_GEMINI_MODEL,
                        "gemini-2.5-flash-lite",
                        GeminiDispatchModelTryOrderDefaults.ENV_GEMINI_MODEL_TRY_ORDER,
                        "gemini-3.5-flash");
        assertEquals(
                List.of("gemini-2.5-flash-lite"),
                GeminiDispatchModelTryOrderDefaults.resolveEffectiveModelTryOrder(env));
    }

    @Test
    void resolveEffectiveModelTryOrder_fallsBackToCodeDefaults() {
        List<String> models = GeminiDispatchModelTryOrderDefaults.resolveEffectiveModelTryOrder(Map.of());
        assertTrue(models.contains("gemini-3.1-flash-lite"));
        assertEquals(
                GeminiDispatchModelTryOrderDefaults.PLANNING_CORE_FALLBACK_TRY_ORDER.size(),
                models.size());
    }
}
