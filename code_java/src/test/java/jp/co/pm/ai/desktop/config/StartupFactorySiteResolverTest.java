package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;

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

        assertEquals(FactorySite.KONAN, StartupFactorySiteResolver.resolveForSplash());
    }

    @Test
    void resolveForSplash_usesLastLaunchedJson() {
        LastLaunchedFactorySiteStore.save(FactorySite.KOKUBU);
        GlobalInitSettingTarget.save(FactorySite.KONAN);

        assertEquals(FactorySite.KOKUBU, StartupFactorySiteResolver.resolveForSplash());
        assertEquals(FactorySite.KONAN, GlobalInitSettingTarget.load());
    }
}
