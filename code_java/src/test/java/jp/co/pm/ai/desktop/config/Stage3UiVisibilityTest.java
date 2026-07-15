package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;

import org.junit.jupiter.api.Test;

class Stage3UiVisibilityTest {

    @Test
    void defaultsToHiddenAndAcceptsTruthyValues() {
        assertFalse(Stage3UiVisibility.isVisible(Map.of()));
        assertFalse(
                Stage3UiVisibility.isVisible(
                        Map.of(AppPaths.KEY_PM_AI_STAGE3_UI_VISIBLE, "0")));
        assertTrue(
                Stage3UiVisibility.isVisible(
                        Map.of(AppPaths.KEY_PM_AI_STAGE3_UI_VISIBLE, "1")));
    }

    @Test
    void bundledEnvironmentDefaultIsHidden() {
        var entry =
                UiRefEnvDefaults.loadOrEmpty().stream()
                        .filter(e -> AppPaths.KEY_PM_AI_STAGE3_UI_VISIBLE.equals(e.key()))
                        .findFirst()
                        .orElseThrow();
        assertTrue("0".equals(entry.value()));
        assertFalse(EnvVarDocs.logicOnly(AppPaths.KEY_PM_AI_STAGE3_UI_VISIBLE).isBlank());
    }
}
