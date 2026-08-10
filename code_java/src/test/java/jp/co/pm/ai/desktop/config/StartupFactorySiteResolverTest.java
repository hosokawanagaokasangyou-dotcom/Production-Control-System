package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class StartupFactorySiteResolverTest {

    private String priorHome;
    private String priorUserHome;
    private String priorOperatorStore;
    private String priorLastSelectedDir;

    @BeforeEach
    void setUp(@TempDir Path tmp) throws Exception {
        priorHome = AppPaths.desktopAppHomeDirName();
        priorUserHome = System.getProperty("user.home");
        priorOperatorStore = System.getProperty("pm.ai.test.factoryOperatorUserStore");
        priorLastSelectedDir = System.getProperty("pm.ai.test.factoryOperatorLastSelectedDir");

        System.setProperty("user.home", tmp.toString());
        AppPaths.setDesktopAppHomeDirName(".pm-ai-desktop-test");
        System.setProperty(
                "pm.ai.test.factoryOperatorUserStore", tmp.resolve("operators.bin").toString());
        System.setProperty(
                "pm.ai.test.factoryOperatorLastSelectedDir",
                tmp.resolve("last-selected").toString());
        FactoryOperatorUserStore.resetStoreForTests();
        FactorySiteWorkspaceStore.resetForTests();
    }

    @AfterEach
    void tearDown() throws Exception {
        FactorySiteWorkspaceStore.resetForTests();
        FactoryOperatorUserStore.resetStoreForTests();
        AppPaths.setDesktopAppHomeDirName(priorHome);
        System.setProperty("user.home", priorUserHome);
        restoreProperty("pm.ai.test.factoryOperatorUserStore", priorOperatorStore);
        restoreProperty("pm.ai.test.factoryOperatorLastSelectedDir", priorLastSelectedDir);
    }

    private static void restoreProperty(String key, String value) {
        if (value != null) {
            System.setProperty(key, value);
        } else {
            System.clearProperty(key);
        }
    }

    @Test
    void resolveForSplash_usesFactorySiteKeyFromSession() {
        GlobalInitSettingTarget.save(FactorySite.KONAN);
        DesktopSessionStateStore.patchUiEnvMapAndTheme(
                Map.of(AppPaths.KEY_PM_AI_FACTORY_SITE, FactorySite.KOKUBU.name()), "");

        assertEquals(FactorySite.KOKUBU, StartupFactorySiteResolver.resolveForSplash());
        assertEquals(FactorySite.KONAN, GlobalInitSettingTarget.load());
    }

    @Test
    void resolveForSplash_ignoresUiRefFactorySiteWhenNotInSession() {
        GlobalInitSettingTarget.save(FactorySite.KOKUBU);
        DesktopSessionStateStore.patchUiEnvMapAndTheme(
                Map.of(
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        FactorySite.KOKUBU.taskInputSourceDir()),
                "");

        assertEquals(FactorySite.KOKUBU, StartupFactorySiteResolver.resolveForSplash());
        assertEquals(FactorySite.KOKUBU, GlobalInitSettingTarget.load());
    }

    @Test
    void resolveOperatorWorkspaceLastFactory_returnsLastWhenReachable(@TempDir Path sharedData)
            throws Exception {
        Files.createDirectories(sharedData);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, sharedData.toString());

        GlobalInitSettingTarget.save(FactorySite.KONAN);
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
        FactoryOperatorUserStore.clearSessionOperatorName();
        FactorySiteWorkspaceStore.saveLastFactorySite("砂田", FactorySite.KOKUBU);

        Optional<FactorySite> resolved =
                StartupFactorySiteResolver.resolveOperatorWorkspaceLastFactory(
                        ui, FactorySite.KONAN);
        assertEquals(FactorySite.KOKUBU, resolved.orElseThrow());
    }

    @Test
    void resolveOperatorWorkspaceLastFactory_skipsWhenOperatorNotInTargetList(@TempDir Path sharedData)
            throws Exception {
        Files.createDirectories(sharedData);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, sharedData.toString());

        FactoryOperatorUserStore.addName(FactorySite.KONAN, "湖南専用");
        FactoryOperatorUserStore.removeName(FactorySite.KOKUBU, "湖南専用");
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "湖南専用");
        FactoryOperatorUserStore.clearSessionOperatorName();
        FactorySiteWorkspaceStore.saveLastFactorySite("湖南専用", FactorySite.KOKUBU);

        assertTrue(
                StartupFactorySiteResolver.resolveOperatorWorkspaceLastFactory(
                                ui, FactorySite.KONAN)
                        .isEmpty());
    }

    @Test
    void resolveOperatorWorkspaceLastFactory_skipsWhenPinLocked(@TempDir Path sharedData)
            throws Exception {
        Files.createDirectories(sharedData);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, sharedData.toString());

        FactoryOperatorUserStore.assignPinByAdmin(FactorySite.KONAN, "砂田", "1234");
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
        FactoryOperatorUserStore.clearSessionOperatorName();
        for (int i = 0; i < FactoryOperatorUserStore.MAX_CONSECUTIVE_PIN_FAILURES; i++) {
            FactoryOperatorUserStore.verifyPinAttempt(FactorySite.KONAN, "砂田", "0000");
        }
        FactorySiteWorkspaceStore.saveLastFactorySite("砂田", FactorySite.KOKUBU);

        assertTrue(
                StartupFactorySiteResolver.resolveOperatorWorkspaceLastFactory(
                                ui, FactorySite.KONAN)
                        .isEmpty());
    }
}
