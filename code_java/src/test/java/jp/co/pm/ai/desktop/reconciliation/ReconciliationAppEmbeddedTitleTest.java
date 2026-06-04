package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;

class ReconciliationAppEmbeddedTitleTest {

    @Test
    void embeddedWindowTitle_hasFactorySuffix() {
        String title = ReconciliationApp.embeddedWindowTitle(Map.of());
        assertTrue(title.endsWith("統合管理データベース (JavaFX版)"));
        assertTrue(title.contains("工場 "));
    }

    @Test
    void embeddedWindowTitle_reflectsKonanFromUiEnv() {
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        FactorySite.KONAN.taskInputSourceDir());
        assertEquals("湖南工場 統合管理データベース (JavaFX版)", ReconciliationApp.embeddedWindowTitle(ui));
    }

    @Test
    void embeddedWindowTitle_reflectsKokubuFromUiEnv() {
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        FactorySite.KOKUBU.taskInputSourceDir());
        String title = ReconciliationApp.embeddedWindowTitle(ui);
        assertEquals("国分工場 統合管理データベース (JavaFX版)", title);
    }
}
