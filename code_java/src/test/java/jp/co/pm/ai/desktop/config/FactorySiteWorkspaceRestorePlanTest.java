package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import java.util.Optional;

import org.junit.jupiter.api.Test;

class FactorySiteWorkspaceRestorePlanTest {

    @Test
    void savedUiEnvRows_loadsWorkspaceEnvWithoutBundledReset() {
        FactorySiteWorkspaceSnapshot snapshot =
                new FactorySiteWorkspaceSnapshot(
                        List.of(new UiEnvRowSnapshot(AppPaths.KEY_PM_AI_FACTORY_SITE, "KOKUBU", "")),
                        DesktopSessionState.empty());

        FactorySiteWorkspaceRestorePlan plan =
                FactorySiteWorkspaceRestorePlan.of(Optional.of(snapshot));

        assertTrue(plan.applyInitSettingBeforeEnv());
        assertTrue(plan.restoreSavedUiEnvRows());
        assertFalse(plan.bundledEnvReset());
        assertFalse(plan.overlayFactoryNetworkDefaults());
        assertTrue(plan.applySessionFragment());
        assertFalse(plan.preserveEnvInitializationInSessionFragment());
    }

    @Test
    void missingWorkspace_initializesBundledEnvWithFactoryOverlay() {
        FactorySiteWorkspaceRestorePlan plan = FactorySiteWorkspaceRestorePlan.of(Optional.empty());

        assertTrue(plan.applyInitSettingBeforeEnv());
        assertFalse(plan.restoreSavedUiEnvRows());
        assertTrue(plan.bundledEnvReset());
        assertTrue(plan.overlayFactoryNetworkDefaults());
        assertFalse(plan.applySessionFragment());
        assertTrue(plan.preserveEnvInitializationInSessionFragment());
    }

    @Test
    void workspaceWithoutUiEnvRows_initializesEnvThenAppliesSessionFragment() {
        FactorySiteWorkspaceSnapshot snapshot =
                new FactorySiteWorkspaceSnapshot(List.of(), DesktopSessionState.empty());

        FactorySiteWorkspaceRestorePlan plan =
                FactorySiteWorkspaceRestorePlan.of(Optional.of(snapshot));

        assertTrue(plan.applyInitSettingBeforeEnv());
        assertFalse(plan.restoreSavedUiEnvRows());
        assertTrue(plan.bundledEnvReset());
        assertTrue(plan.overlayFactoryNetworkDefaults());
        assertTrue(plan.applySessionFragment());
        assertTrue(plan.preserveEnvInitializationInSessionFragment());
    }
}
