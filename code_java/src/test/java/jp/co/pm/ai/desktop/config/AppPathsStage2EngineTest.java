package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.HashMap;
import java.util.Map;

import org.junit.jupiter.api.Test;

class AppPathsStage2EngineTest {

    @Test
    void stage2EngineUsesJava_onlyForJavaToken() {
        Map<String, String> ui = new HashMap<>();
        assertFalse(AppPaths.stage2EngineUsesJava(ui));
        ui.put(AppPaths.KEY_PM_AI_STAGE2_ENGINE, "");
        assertFalse(AppPaths.stage2EngineUsesJava(ui));
        ui.put(AppPaths.KEY_PM_AI_STAGE2_ENGINE, "python");
        assertFalse(AppPaths.stage2EngineUsesJava(ui));
        ui.put(AppPaths.KEY_PM_AI_STAGE2_ENGINE, "JAVA");
        assertTrue(AppPaths.stage2EngineUsesJava(ui));
    }
}
