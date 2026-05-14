package jp.co.pm.ai.planning.stage2;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.AppPaths;

class Stage2EnvParsingTest {

    @Test
    void stage2WriteExcel_defaultOn() {
        assertTrue(Stage2EnvParsing.stage2WriteExcel(Map.of()));
        assertTrue(Stage2EnvParsing.stage2WriteExcel(null));
    }

    @Test
    void stage2WriteExcel_offTokens() {
        assertFalse(Stage2EnvParsing.stage2WriteExcel(Map.of(AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL, "0")));
        assertFalse(Stage2EnvParsing.stage2WriteExcel(Map.of(AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL, "false")));
    }

    @Test
    void envEnabled_truthyAndOff() {
        assertTrue(Stage2EnvParsing.envEnabled("K", Map.of("K", "1"), false));
        assertFalse(Stage2EnvParsing.envEnabled("K", Map.of("K", "0"), true));
    }
}
