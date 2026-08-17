package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class PortableBundleUpgradeFollowUpTest {

    @Test
    void isPendingFor_matchesInstallRoot(@TempDir Path tmp) throws Exception {
        Path home = tmp.resolve("home");
        Path install = tmp.resolve("install");
        Files.createDirectories(install);
        String prev = System.getProperty("user.home");
        try {
            System.setProperty("user.home", home.toString());
            PortableBundleUpgradeFollowUp.writePending(install, "7.17", FactorySite.KONAN);
            assertTrue(PortableBundleUpgradeFollowUp.isPendingFor(install));
            assertFalse(PortableBundleUpgradeFollowUp.isPendingFor(tmp.resolve("other")));
            PortableBundleUpgradeFollowUp.clear();
            assertFalse(PortableBundleUpgradeFollowUp.isPendingFor(install));
        } finally {
            if (prev != null) {
                System.setProperty("user.home", prev);
            } else {
                System.clearProperty("user.home");
            }
        }
    }

    @Test
    void writePending_invalidatesEnvVarsInitializationRecord(@TempDir Path tmp) throws Exception {
        Path home = tmp.resolve("home");
        Path install = tmp.resolve("install");
        Files.createDirectories(install);
        String prev = System.getProperty("user.home");
        try {
            System.setProperty("user.home", home.toString());
            EnvVarsInitializedAtStore.recordNow();
            EnvVarsInitializedAtStore.recordEnvFingerprint(
                    java.util.Map.of("PM_AI_REPO_ROOT", "C:\\repo"), k -> true);
            assertTrue(EnvVarsInitializedAtStore.isRecorded());

            PortableBundleUpgradeFollowUp.writePending(install, "7.17", FactorySite.KONAN);

            assertTrue(PortableBundleUpgradeFollowUp.isPendingFor(install));
            assertFalse(EnvVarsInitializedAtStore.isRecorded());
            assertFalse(EnvVarsInitializedAtStore.loadEnvFingerprint().isPresent());
            PortableBundleUpgradeFollowUp.clear();
        } finally {
            if (prev != null) {
                System.setProperty("user.home", prev);
            } else {
                System.clearProperty("user.home");
            }
        }
    }
}
