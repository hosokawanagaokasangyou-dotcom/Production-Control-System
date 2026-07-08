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
import java.util.Optional;

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
    void ensureDeployed_skipsWhenExeHashMatchesDespiteOlderSharedVersion(@TempDir Path tmp)
            throws Exception {
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
            if (in == null) {
                return;
            }
            Files.copy(in, exe, StandardCopyOption.REPLACE_EXISTING);
        }
        BigDecimal olderShared = bundledVer.get().subtract(new BigDecimal("0.01"));
        Files.writeString(version, olderShared.toPlainString() + "\n", StandardCharsets.UTF_8);

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_RDP_LAUNCHER_EXE,
                        exe.toString(),
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        deployDir.resolve("summary.xlsx").toString());

        assertFalse(RdpRemoteLauncherDeployer.needsExeDeploy(ui));

        RdpRemoteLauncherDeployer.DeployOutcome outcome =
                RdpRemoteLauncherDeployer.ensureDeployed(ui);
        assertTrue(outcome.upToDate());
        assertFalse(outcome.copied());
        assertTrue(
                outcome.message().orElse("").contains("最新")
                        || RdpRemoteLauncherDeployer.parseVersionFile(version)
                                .map(v -> v.compareTo(bundledVer.get()) >= 0)
                                .orElse(false));
    }

    @Test
    void needsExeDeploy_falseWhenAutoDeployDisabled() {
        assertFalse(
                RdpRemoteLauncherDeployer.needsExeDeploy(
                        Map.of(AppPaths.KEY_PM_AI_RDP_LAUNCHER_AUTO_DEPLOY, "0")));
    }

    @Test
    void forceDeploy_overwritesUpToDateVersion(@TempDir Path tmp) throws Exception {
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

        RdpRemoteLauncherDeployer.DeployOutcome ensure =
                RdpRemoteLauncherDeployer.ensureDeployed(ui);
        assertFalse(ensure.copied());

        RdpRemoteLauncherDeployer.DeployOutcome forced =
                RdpRemoteLauncherDeployer.forceDeploy(ui);
        assertTrue(forced.copied() || forced.message().orElse("").contains("強制転送"));
    }

    @Test
    void deployOutcome_succeeded() {
        assertTrue(
                new RdpRemoteLauncherDeployer.DeployOutcome(true, false, Optional.empty())
                        .succeeded());
        assertTrue(
                new RdpRemoteLauncherDeployer.DeployOutcome(false, true, Optional.empty())
                        .succeeded());
        assertFalse(
                new RdpRemoteLauncherDeployer.DeployOutcome(false, false, Optional.empty())
                        .succeeded());
    }

    @Test
    void looksLikeFileInUse_detectsSharingViolation() {
        assertTrue(
                RdpRemoteLauncherDeployer.looksLikeFileInUse(
                        new java.io.IOException(
                                "The process cannot access the file because it is being used by"
                                        + " another process")));
        assertTrue(
                RdpRemoteLauncherDeployer.looksLikeFileInUse(
                        new java.io.IOException("別のプロセスによって使用されているため、プロセスはファイルにアクセスできません")));
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
