package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class FactorySiteTest {

    @Test
    void konanMatchesAppPathsDefaults() {
        assertEquals(AppPaths.DEFAULT_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR, FactorySite.KONAN.portableBundleSourceDir());
        assertEquals(AppPaths.DEFAULT_PM_AI_TASK_INPUT_SOURCE_DIR, FactorySite.KONAN.taskInputSourceDir());
        assertEquals(AppPaths.DEFAULT_PM_AI_ACTUAL_DETAIL_SOURCE_DIR, FactorySite.KONAN.actualDetailSourceDir());
        assertEquals("", FactorySite.KONAN.masterWorkbookFileBasename());
    }

    @Test
    void kokubuUsesKokubuNetworkLayout() {
        assertTrue(FactorySite.KOKUBU.portableBundleSourceDir().contains("国分工場"));
        assertTrue(FactorySite.KOKUBU.portableBundleSourceDir().endsWith("PMD_version_upgrade.zip"));
        assertTrue(FactorySite.KOKUBU.taskInputSourceDir().contains("DATA\\計画"));
        assertTrue(FactorySite.KOKUBU.actualDetailSourceDir().contains("DATA\\実績"));
        assertEquals("国分master.xlsm", FactorySite.KOKUBU.masterWorkbookFileBasename());
    }

    @Test
    void toStringUsesJapaneseLabel() {
        assertEquals("湖南工場", FactorySite.KONAN.toString());
        assertEquals("国分工場", FactorySite.KOKUBU.toString());
    }
}
