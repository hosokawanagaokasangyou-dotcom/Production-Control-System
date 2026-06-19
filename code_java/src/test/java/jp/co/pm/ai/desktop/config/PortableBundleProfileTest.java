package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class PortableBundleProfileTest {

    @Test
    void rdpLayout_detectsExeAndApp(@TempDir Path tmp) throws Exception {
        Files.createDirectories(tmp.resolve("app"));
        Files.writeString(tmp.resolve(AppPaths.RDP_DESKTOP_LAUNCHER_EXE_BASENAME), "stub");
        assertTrue(PortableBundleProfile.RDP_LAUNCHER.isPortableBundleLayout(tmp));
    }

    @Test
    void rdpUpgradeZip_resolvedUnderReleaseFolder(@TempDir Path release) throws Exception {
        Path zip = release.resolve(PortableBundleProfile.RDP_LAUNCHER.upgradeZipName());
        Files.writeString(zip, "zip");
        assertEquals(
                zip,
                PortableBundleSelfUpdater.resolveEffectiveUpgradeZip(
                        PortableBundleProfile.RDP_LAUNCHER, release)
                        .orElseThrow());
    }

    @Test
    void rdpPortableBundleDefaultCanonicalMatchesBoardShare() {
        assertEquals(
                "\\\\192.168.0.101\\共有フォルダ\\掲示板\\rpa_luncher\\RDPランチャ",
                AppPaths.DEFAULT_PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR);
    }

    @Test
    void defaultRdpSharedDataDir_underRpaLauncherShare() {
        assertEquals(
                "\\\\192.168.0.101\\共有フォルダ\\掲示板\\rpa_luncher\\"
                        + AppPaths.RDP_LAUNCHER_SHARED_DATA_DIR_LEAF,
                AppPaths.DEFAULT_PM_AI_RPA_LAUNCHER_OPERATOR_USERS_STORE_DIR);
        assertEquals(
                AppPaths.DEFAULT_PM_AI_RPA_LAUNCHER_OPERATOR_USERS_STORE_DIR,
                AppPaths.DEFAULT_PM_AI_RDP_SHARED_DATA_DIR);
        Path store = AppPaths.rdpLauncherOperatorUsersStorePath();
        assertTrue(store.toString().contains(AppPaths.RDP_LAUNCHER_SHARED_DATA_DIR_LEAF));
        assertEquals(AppPaths.RDP_LAUNCHER_OPERATOR_USERS_BIN, store.getFileName().toString());
    }

    @Test
    void pmdRdpLauncherDefaults_useKonanSharedDataDrive() {
        assertEquals(
                AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR,
                AppPaths.DEFAULT_PM_AI_RDP_LAUNCHER_DEPLOY_DIR);
        assertEquals(
                AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR,
                AppPaths.DEFAULT_PM_AI_RDP_OPERATOR_USERS_STORE_DIR);
    }
}
