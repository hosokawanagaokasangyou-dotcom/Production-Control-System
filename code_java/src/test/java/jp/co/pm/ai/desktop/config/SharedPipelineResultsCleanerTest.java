package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class SharedPipelineResultsCleanerTest {

    @Test
    void shouldDeleteStageArtifactFileName_keepsAladdinAndSummary() {
        assertFalse(
                SharedPipelineResultsCleaner.shouldDeleteStageArtifactFileName(
                        "アラジン入力用_配台計画_20260715-113257.xlsx"));
        assertFalse(
                SharedPipelineResultsCleaner.shouldDeleteStageArtifactFileName("サマリ_AI配台.xlsx"));
    }

    @Test
    void shouldDeleteStageArtifactFileName_targetsStage1And2Cores() {
        assertTrue(
                SharedPipelineResultsCleaner.shouldDeleteStageArtifactFileName(
                        AppPaths.STAGE1_PLAN_TASKS_FILENAME));
        assertTrue(
                SharedPipelineResultsCleaner.shouldDeleteStageArtifactFileName(
                        AppPaths.STAGE1_EXCLUDE_RULES_JSON_FILENAME));
        assertTrue(
                SharedPipelineResultsCleaner.shouldDeleteStageArtifactFileName(
                        AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME));
        assertTrue(
                SharedPipelineResultsCleaner.shouldDeleteStageArtifactFileName(
                        "計画_20260715-120000.xlsx"));
        assertTrue(
                SharedPipelineResultsCleaner.shouldDeleteStageArtifactFileName(
                        "shaped_aladdin_plan.json"));
    }
}
