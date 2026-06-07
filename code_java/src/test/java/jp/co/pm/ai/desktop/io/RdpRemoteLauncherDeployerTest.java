package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class RdpRemoteLauncherDeployerTest {

    @Test
    void isAutoDeployEnabled_defaultsTrue() {
        assertTrue(RdpRemoteLauncherDeployer.isAutoDeployEnabled(Map.of()));
        assertFalse(
                RdpRemoteLauncherDeployer.isAutoDeployEnabled(
                        Map.of(AppPaths.KEY_PM_AI_RDP_LAUNCHER_AUTO_DEPLOY, "0")));
    }

    @Test
    void ensureDeployed_copiesWhenMissing(@TempDir Path tmp) throws Exception {
        Path deployDir = tmp.resolve("deploy");
        Files.createDirectories(deployDir);
        Path exe = deployDir.resolve(AppPaths.RDP_LAUNCHER_EXE_BASENAME);
        Path version = deployDir.resolve(AppPaths.RDP_LAUNCHER_VERSION_BASENAME);

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_RDP_LAUNCHER_EXE,
                        exe.toString(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        deployDir.resolve("summary.xlsx").toString());

        RdpRemoteLauncherDeployer.DeployOutcome outcome =
                RdpRemoteLauncherDeployer.ensureDeployed(ui);
        if (RdpRemoteLauncherDeployer.readBundledVersion().isEmpty()) {
            assertFalse(outcome.copied());
            return;
        }
        assertTrue(outcome.copied() || outcome.upToDate());
        if (outcome.copied()) {
            assertTrue(Files.isRegularFile(exe));
            assertTrue(Files.isRegularFile(version));
        }
    }

    @Test
    void ensureDeployed_skipsWhenUpToDate(@TempDir Path tmp) throws Exception {
        var bundledVer = RdpRemoteLauncherDeployer.readBundledVersion();
        if (bundledVer.isEmpty()) {
            return;
        }
        Path deployDir = tmp.resolve("deploy");
        Files.createDirectories(deployDir);
        Path exe = deployDir.resolve(AppPaths.RDP_LAUNCHER_EXE_BASENAME);
        Path version = deployDir.resolve(AppPaths.RDP_LAUNCHER_VERSION_BASENAME);

        try (InputStream in =
                RdpRemoteLauncherDeployer.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/rdp-launcher/" + AppPaths.RDP_LAUNCHER_EXE_BASENAME)) {
            if (in != null) {
                Files.copy(in, exe, StandardCopyOption.REPLACE_EXISTING);
            }
        }
        Files.writeString(version, bundledVer.get().toPlainString() + "\n", StandardCharsets.UTF_8);

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_RDP_LAUNCHER_EXE,
                        exe.toString(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        deployDir.resolve("summary.xlsx").toString());

        RdpRemoteLauncherDeployer.DeployOutcome outcome =
                RdpRemoteLauncherDeployer.ensureDeployed(ui);
        assertTrue(outcome.upToDate());
        assertFalse(outcome.copied());
    }

    @Test
    void parseVersionFile_readsDecimal(@TempDir Path tmp) throws Exception {
        Path f = tmp.resolve("v.txt");
        Files.writeString(f, "10.78\n", StandardCharsets.UTF_8);
        assertTrue(
                RdpRemoteLauncherDeployer.parseVersionFile(f)
                        .map(v -> v.equals(new BigDecimal("10.78")))
                        .orElse(false));
    }
}
