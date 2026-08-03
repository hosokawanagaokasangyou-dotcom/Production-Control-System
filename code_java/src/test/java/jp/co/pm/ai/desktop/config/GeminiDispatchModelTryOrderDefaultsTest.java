package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
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
        assertEquals(
                GeminiDispatchModelTryOrderDefaults.PLANNING_CORE_TOP_PRIORITY_MODEL,
                models.getFirst());
        assertTrue(models.contains("gemini-3.1-flash-lite"));
        assertEquals(
                GeminiDispatchModelTryOrderDefaults.PLANNING_CORE_FALLBACK_TRY_ORDER.size(),
                models.size());
    }

    @Test
    void codeDefaultsDropGenerationsWithoutFreeTierAllocation() {
        for (String id : GeminiDispatchModelTryOrderDefaults.PLANNING_CORE_FALLBACK_TRY_ORDER) {
            assertTrue(
                    GeminiDispatchModelTryOrderDefaults.hasFreeTierAllocation(id),
                    "無料枠の割り当てが無い世代を既定列に残さない: " + id);
        }
    }

    @Test
    void hasFreeTierAllocation_rejectsExhaustedGenerationsAndPro() {
        assertTrue(GeminiDispatchModelTryOrderDefaults.hasFreeTierAllocation("gemini-3.5-flash"));
        assertTrue(
                GeminiDispatchModelTryOrderDefaults.hasFreeTierAllocation(
                        "models/gemini-3.1-flash-lite"));
        assertTrue(GeminiDispatchModelTryOrderDefaults.hasFreeTierAllocation("gemini-flash-latest"));

        assertFalse(GeminiDispatchModelTryOrderDefaults.hasFreeTierAllocation("gemini-2.5-flash-lite"));
        assertFalse(GeminiDispatchModelTryOrderDefaults.hasFreeTierAllocation("gemini-2.0-flash-lite"));
        assertFalse(GeminiDispatchModelTryOrderDefaults.hasFreeTierAllocation("models/gemini-1.5-flash"));
        assertFalse(GeminiDispatchModelTryOrderDefaults.hasFreeTierAllocation("gemini-3.5-pro"));
        assertFalse(GeminiDispatchModelTryOrderDefaults.hasFreeTierAllocation("  "));
        assertFalse(GeminiDispatchModelTryOrderDefaults.hasFreeTierAllocation(null));
    }

    @Test
    void withPlanningCorePriorityFirst_prependsTopPriorityBeforeFlashLiteCandidates() {
        List<String> merged =
                GeminiDispatchModelTryOrderDefaults.withPlanningCorePriorityFirst(
                        List.of("gemini-3.5-flash-lite", "gemini-3.1-flash-lite"));
        assertEquals(
                GeminiDispatchModelTryOrderDefaults.PLANNING_CORE_TOP_PRIORITY_MODEL,
                merged.getFirst());
        assertEquals(
                List.of(
                        "gemini-3.5-flash",
                        "gemini-3.5-flash-lite",
                        "gemini-3.1-flash-lite",
                        "gemini-3.1-flash-lite-preview",
                        "gemini-flash-latest"),
                merged);
    }

    @Test
    void withPlanningCorePriorityFirst_dropsCandidatesWithoutFreeTierAllocation() {
        List<String> merged =
                GeminiDispatchModelTryOrderDefaults.withPlanningCorePriorityFirst(
                        List.of("gemini-2.5-flash-lite", "gemini-2.0-flash-lite"));
        assertFalse(merged.contains("gemini-2.5-flash-lite"));
        assertFalse(merged.contains("gemini-2.0-flash-lite"));
        assertEquals(GeminiDispatchModelTryOrderDefaults.PLANNING_CORE_FALLBACK_TRY_ORDER, merged);
    }
}
