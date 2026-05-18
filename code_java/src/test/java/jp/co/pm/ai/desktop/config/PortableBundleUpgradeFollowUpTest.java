package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
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
            PortableBundleUpgradeFollowUp.writePending(install, "7.17");
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
}
