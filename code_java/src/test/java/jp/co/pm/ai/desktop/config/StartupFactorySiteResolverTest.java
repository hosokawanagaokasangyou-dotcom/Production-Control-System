package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class StartupFactorySiteResolverTest {

    private String priorHome;
    private String priorUserHome;

    @BeforeEach
    void setUp(@TempDir Path tmp) {
        priorHome = AppPaths.desktopAppHomeDirName();
        priorUserHome = System.getProperty("user.home");
        System.setProperty("user.home", tmp.toString());
        AppPaths.setDesktopAppHomeDirName(".pm-ai-desktop-test");
        LastLaunchedFactorySiteStore.resetForTests();
    }

    @AfterEach
    void tearDown() {
        LastLaunchedFactorySiteStore.resetForTests();
        AppPaths.setDesktopAppHomeDirName(priorHome);
        System.setProperty("user.home", priorUserHome);
    }

    @Test
    void resolveForSplash_whenJsonMissing_returnsKonan() {
        GlobalInitSettingTarget.save(FactorySite.KOKUBU);
        DesktopSessionStateStore.patchUiEnvMapAndTheme(
                Map.of(AppPaths.KEY_PM_AI_FACTORY_SITE, FactorySite.KOKUBU.name()), "");

        assertEquals(FactorySite.KONAN, StartupFactorySiteResolver.resolve());
        assertEquals(FactorySite.KONAN, StartupFactorySiteResolver.resolveForSplash());
    }

    @Test
    void resolveForSplash_usesLastLaunchedJson() {
        LastLaunchedFactorySiteStore.save(FactorySite.KOKUBU);
        GlobalInitSettingTarget.save(FactorySite.KONAN);

        assertEquals(FactorySite.KOKUBU, StartupFactorySiteResolver.resolve());
        assertEquals(FactorySite.KOKUBU, StartupFactorySiteResolver.resolveForSplash());
        assertEquals(FactorySite.KONAN, GlobalInitSettingTarget.load());
    }

    @Test
    void resolveForPortableUpgrade_prefersLastLaunchedOverKonanEnvUnc() {
        LastLaunchedFactorySiteStore.save(FactorySite.KOKUBU);
        GlobalInitSettingTarget.save(FactorySite.KOKUBU);
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR,
                        FactorySite.KONAN.portableBundleSourceDir(),
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        FactorySite.KONAN.taskInputSourceDir(),
                        AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR,
                        FactorySite.KONAN.actualDetailSourceDir());

        assertEquals(
                FactorySite.KOKUBU,
                StartupFactorySiteResolver.resolveForPortableUpgrade(
                        java.util.Optional.empty(), ui, java.util.Optional.empty()));
    }

    @Test
    void requiresStartupSwitch_onlyWhenPersistedDiffersFromAdopted() {
        assertTrue(
                StartupFactorySiteResolver.requiresStartupSwitch(
                        FactorySite.KONAN, FactorySite.KOKUBU));
        assertFalse(
                StartupFactorySiteResolver.requiresStartupSwitch(
                        FactorySite.KOKUBU, FactorySite.KOKUBU));
        assertFalse(
                StartupFactorySiteResolver.requiresStartupSwitch(
                        FactorySite.KONAN, FactorySite.RDP_LAUNCHER));
        assertTrue(StartupFactorySiteResolver.requiresStartupSwitch(null, FactorySite.KONAN));
        assertFalse(StartupFactorySiteResolver.requiresStartupSwitch(FactorySite.KONAN, null));
    }
}
